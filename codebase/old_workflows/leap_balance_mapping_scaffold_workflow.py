#%%
from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any

import openpyxl
import pandas as pd

from codebase.utilities.master_config import read_config_table

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.utilities.energy_balance_template_extractor import TemplateBalanceExtractor  # noqa: E402
from codebase.utilities.leap_balance_export_resolver import resolve_balance_export_workbook  # noqa: E402
from codebase.utilities.workflow_common import archive_config_dir_once_per_day  # noqa: E402
from codebase.utilities.workflow_outputs import build_workflow_output_layout, write_output_manifest  # noqa: E402
from codebase.utilities.output_paths import MAPPINGS_ROOT  # noqa: E402


#%%
def _resolve(path: Path | str) -> Path:
    raw = str(path).replace("\\", "/")
    drive_match = re.match(r"^([a-zA-Z]):/(.*)$", raw)
    if drive_match:
        drive = drive_match.group(1).lower()
        rest = drive_match.group(2)
        return Path(f"/mnt/{drive}/{rest}")
    candidate = Path(raw)
    return candidate if candidate.is_absolute() else (REPO_ROOT / candidate)


def _clean_token(value: object) -> str:
    text = str(value or "").strip()
    return "" if text.lower() in {"", "nan", "none"} else text


def _list_balance_sheets(workbook_path: Path) -> list[str]:
    wb = openpyxl.load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        names = [str(name).strip() for name in wb.sheetnames]
    finally:
        wb.close()

    selected: list[str] = []
    for name in names:
        key = name.lower()
        if key.startswith("ebal|"):
            selected.append(name)
            continue
        if key.startswith("energy balance") or key.startswith("targt energy balance"):
            selected.append(name)
            continue
    return selected


def _coalesce_unique(series: pd.Series) -> str:
    cleaned = [str(v).strip() for v in series.fillna("").astype(str).tolist() if str(v).strip()]
    unique = sorted(set(cleaned))
    return unique[0] if unique else ""


def _is_total_like(value: object) -> bool:
    text = _clean_token(value).lower()
    if not text:
        return False
    if text == "total":
        return True
    if text.endswith(" total") or text.startswith("total "):
        return True
    if "_total" in text or text == "19_total":
        return True
    if re.match(r"^\d{2}(?:\.\d{2})*\s+total\b", text):
        return True
    return False


def _drop_total_rows(mapped: pd.DataFrame) -> pd.DataFrame:
    out = mapped.copy()
    check_cols = [
        "leap_sector_name_raw",
        "leap_sector_name",
        "leap_fuel_name",
        "leap_sector",
        "leap_fuel",
        "esto_flow",
        "esto_product",
    ]
    mask = pd.Series(False, index=out.index)
    for col in check_cols:
        if col not in out.columns:
            continue
        mask = mask | out[col].map(_is_total_like)
    return out.loc[~mask].copy()


def _append_mapping_cardinality(
    *,
    out: pd.DataFrame,
    relations: pd.DataFrame,
    left_keys: list[str],
    right_keys: list[str],
) -> pd.DataFrame:
    result = out.copy()
    rel = relations.copy()
    for col in left_keys + right_keys:
        if col not in rel.columns:
            rel[col] = ""
        rel[col] = rel[col].fillna("").astype(str).str.strip()
    rel = rel[
        rel[left_keys].apply(lambda r: all(str(v).strip() for v in r), axis=1)
        & rel[right_keys].apply(lambda r: all(str(v).strip() for v in r), axis=1)
    ].drop_duplicates(left_keys + right_keys)

    if rel.empty:
        result["mapping_scope"] = "unmapped"
        result["mapping_cardinality"] = "unmapped"
        result["left_to_right_count"] = 0
        result["right_to_left_count"] = 0
        return result

    left_counts = rel.groupby(left_keys, as_index=False).size().rename(columns={"size": "left_to_right_count"})
    right_counts = rel.groupby(right_keys, as_index=False).size().rename(columns={"size": "right_to_left_count"})
    result = result.merge(left_counts, on=left_keys, how="left")
    result = result.merge(right_counts, on=right_keys, how="left")
    result["left_to_right_count"] = pd.to_numeric(result["left_to_right_count"], errors="coerce").fillna(0).astype(int)
    result["right_to_left_count"] = pd.to_numeric(result["right_to_left_count"], errors="coerce").fillna(0).astype(int)

    right_non_empty = result[right_keys].apply(lambda r: all(str(v).strip() for v in r), axis=1)
    product_only = result[right_keys[0]].fillna("").astype(str).str.strip().eq("") & result[right_keys[1]].fillna("").astype(str).str.strip().ne("")
    result["mapping_scope"] = "unmapped"
    result.loc[right_non_empty, "mapping_scope"] = "pair"
    result.loc[product_only, "mapping_scope"] = "product_only"

    result["mapping_cardinality"] = "unmapped"
    one_to_one = right_non_empty & result["left_to_right_count"].eq(1) & result["right_to_left_count"].eq(1)
    one_to_many = right_non_empty & result["left_to_right_count"].gt(1) & result["right_to_left_count"].eq(1)
    many_to_one = right_non_empty & result["left_to_right_count"].eq(1) & result["right_to_left_count"].gt(1)
    many_to_many = right_non_empty & result["left_to_right_count"].gt(1) & result["right_to_left_count"].gt(1)
    result.loc[one_to_one, "mapping_cardinality"] = "one_to_one"
    result.loc[one_to_many, "mapping_cardinality"] = "one_to_many"
    result.loc[many_to_one, "mapping_cardinality"] = "many_to_one"
    result.loc[many_to_many, "mapping_cardinality"] = "many_to_many"

    # For product-only rows, classify against product only (left->product is one; inverse may be many).
    product_rel = rel.groupby([*left_keys, right_keys[1]], as_index=False).size()
    product_left = product_rel.groupby(left_keys, as_index=False).size().rename(columns={"size": "left_to_product_count"})
    product_right = product_rel.groupby([right_keys[1]], as_index=False).size().rename(columns={"size": "product_to_left_count"})
    result = result.merge(product_left, on=left_keys, how="left")
    result = result.merge(product_right, on=[right_keys[1]], how="left")
    result["left_to_product_count"] = pd.to_numeric(result["left_to_product_count"], errors="coerce").fillna(0).astype(int)
    result["product_to_left_count"] = pd.to_numeric(result["product_to_left_count"], errors="coerce").fillna(0).astype(int)
    product_one_to_one = product_only & result["left_to_product_count"].eq(1) & result["product_to_left_count"].eq(1)
    product_many_to_one = product_only & result["left_to_product_count"].eq(1) & result["product_to_left_count"].gt(1)
    result.loc[product_one_to_one, "mapping_cardinality"] = "one_to_one"
    result.loc[product_many_to_one, "mapping_cardinality"] = "many_to_one"
    return result


def _extract_code_prefix(label: object) -> str:
    text = _clean_token(label)
    if not text:
        return ""
    m = re.match(r"^(\d{2}(?:\.\d{2})*)\b", text)
    return m.group(1) if m else ""


def _build_code_parent_lookup(codes: set[str], sep: str) -> dict[str, bool]:
    out: dict[str, bool] = {}
    cleaned = sorted({str(c).strip() for c in codes if str(c).strip()})
    for code in cleaned:
        prefix = f"{code}{sep}"
        out[code] = any(other != code and other.startswith(prefix) for other in cleaned)
    return out


def _derive_leap_parent_flags(mapped: pd.DataFrame) -> pd.DataFrame:
    df = mapped.copy()
    for col in [
        "leap_sector_name",
        "leap_sector_name_raw",
        "sector_name_reassigned",
        "leap_sector_row_has_fuel_children",
        "leap_sector_row_has_sector_children",
    ]:
        if col not in df.columns:
            df[col] = ""
    df["leap_sector_name"] = df["leap_sector_name"].fillna("").astype(str).str.strip()
    df["leap_sector_name_raw"] = df["leap_sector_name_raw"].fillna("").astype(str).str.strip()
    reassigned = df["sector_name_reassigned"].fillna(False).astype(bool)
    structure_parent = df["leap_sector_row_has_fuel_children"].fillna(False).astype(bool)
    grouped = (
        df.groupby("leap_sector_name", as_index=False)
        .agg(
            leap_parent_reassigned_rows=("leap_sector_name", lambda s: 0),
            leap_direct_rows=("leap_sector_name", lambda s: 0),
            leap_parent_structure_rows=("leap_sector_name", lambda s: 0),
        )
        .reset_index(drop=True)
    )
    # Fill counts using masks to avoid lambda closure ambiguity.
    reassigned_counts = (
        df[reassigned]
        .groupby("leap_sector_name", as_index=False)
        .size()
        .rename(columns={"size": "leap_parent_reassigned_rows"})
    )
    direct_counts = (
        df[~reassigned]
        .groupby("leap_sector_name", as_index=False)
        .size()
        .rename(columns={"size": "leap_direct_rows"})
    )
    structure_counts = (
        df[structure_parent]
        .groupby("leap_sector_name", as_index=False)
        .size()
        .rename(columns={"size": "leap_parent_structure_rows"})
    )
    grouped = grouped.drop(
        columns=["leap_parent_reassigned_rows", "leap_direct_rows", "leap_parent_structure_rows"],
        errors="ignore",
    )
    grouped = grouped.merge(reassigned_counts, on="leap_sector_name", how="left").merge(
        direct_counts, on="leap_sector_name", how="left"
    ).merge(
        structure_counts, on="leap_sector_name", how="left"
    )
    grouped["leap_parent_reassigned_rows"] = pd.to_numeric(
        grouped["leap_parent_reassigned_rows"], errors="coerce"
    ).fillna(0).astype(int)
    grouped["leap_direct_rows"] = pd.to_numeric(grouped["leap_direct_rows"], errors="coerce").fillna(0).astype(int)
    grouped["leap_parent_structure_rows"] = pd.to_numeric(
        grouped["leap_parent_structure_rows"], errors="coerce"
    ).fillna(0).astype(int)
    grouped["leap_is_parent"] = grouped["leap_parent_reassigned_rows"].gt(0) | grouped["leap_parent_structure_rows"].gt(0)
    grouped["leap_is_leaf"] = ~grouped["leap_is_parent"]
    return grouped


def _derive_esto_parent_lookups(esto_table_path: Path) -> tuple[dict[str, bool], dict[str, bool]]:
    esto = pd.read_csv(esto_table_path)
    flow_codes = {_extract_code_prefix(v) for v in esto.get("flows", pd.Series(dtype=object)).fillna("").astype(str)}
    product_codes = {
        _extract_code_prefix(v) for v in esto.get("products", pd.Series(dtype=object)).fillna("").astype(str)
    }
    flow_lookup = _build_code_parent_lookup({c for c in flow_codes if c}, ".")
    product_lookup = _build_code_parent_lookup({c for c in product_codes if c}, ".")
    return flow_lookup, product_lookup


def _derive_ninth_parent_lookups(codebook_path: Path) -> tuple[dict[str, bool], dict[str, bool]]:
    codebook = read_config_table(codebook_path, sheet_name="code_to_name", dtype=str).fillna("")
    sector_codes = set(
        codebook.loc[
            codebook["9th_column"].astype(str).str.strip().isin(
                ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]
            ),
            "9th_label",
        ]
        .astype(str)
        .str.strip()
        .tolist()
    )
    fuel_codes = set(
        codebook.loc[
            codebook["9th_column"].astype(str).str.strip().isin(["fuels", "subfuels"]),
            "9th_label",
        ]
        .astype(str)
        .str.strip()
        .tolist()
    )
    sector_lookup = _build_code_parent_lookup({c for c in sector_codes if c}, "_")
    fuel_lookup = _build_code_parent_lookup({c for c in fuel_codes if c}, "_")
    return sector_lookup, fuel_lookup


def _apply_parent_flags_to_esto_mapping(
    frame: pd.DataFrame,
    *,
    leap_parent_flags: pd.DataFrame,
    esto_flow_parent_lookup: dict[str, bool],
    esto_product_parent_lookup: dict[str, bool],
) -> pd.DataFrame:
    out = frame.copy()
    out = out.merge(leap_parent_flags, on="leap_sector_name", how="left")
    out["leap_is_parent"] = out["leap_is_parent"].fillna(False).astype(bool)
    out["leap_is_leaf"] = out["leap_is_leaf"].fillna(True).astype(bool)
    out["leap_parent_reassigned_rows"] = pd.to_numeric(out["leap_parent_reassigned_rows"], errors="coerce").fillna(0).astype(int)
    out["leap_direct_rows"] = pd.to_numeric(out["leap_direct_rows"], errors="coerce").fillna(0).astype(int)
    out["leap_parent_structure_rows"] = pd.to_numeric(
        out.get("leap_parent_structure_rows", 0), errors="coerce"
    ).fillna(0).astype(int)

    out["esto_flow_code"] = out["esto_flow"].map(_extract_code_prefix)
    out["esto_product_code"] = out["esto_product"].map(_extract_code_prefix)
    out["esto_flow_is_parent"] = out["esto_flow_code"].map(esto_flow_parent_lookup).fillna(False).astype(bool)
    out["esto_flow_is_leaf"] = out["esto_flow"].fillna("").astype(str).str.strip().ne("") & (~out["esto_flow_is_parent"])
    out["esto_product_is_parent"] = out["esto_product_code"].map(esto_product_parent_lookup).fillna(False).astype(bool)
    out["esto_product_is_leaf"] = (
        out["esto_product"].fillna("").astype(str).str.strip().ne("") & (~out["esto_product_is_parent"])
    )
    return out


def _apply_parent_flags_to_ninth_mapping(
    frame: pd.DataFrame,
    *,
    leap_parent_flags: pd.DataFrame,
    ninth_sector_parent_lookup: dict[str, bool],
    ninth_fuel_parent_lookup: dict[str, bool],
) -> pd.DataFrame:
    out = frame.copy()
    out = out.merge(leap_parent_flags, on="leap_sector_name", how="left")
    out["leap_is_parent"] = out["leap_is_parent"].fillna(False).astype(bool)
    out["leap_is_leaf"] = out["leap_is_leaf"].fillna(True).astype(bool)
    out["leap_parent_reassigned_rows"] = pd.to_numeric(out["leap_parent_reassigned_rows"], errors="coerce").fillna(0).astype(int)
    out["leap_direct_rows"] = pd.to_numeric(out["leap_direct_rows"], errors="coerce").fillna(0).astype(int)
    out["leap_parent_structure_rows"] = pd.to_numeric(
        out.get("leap_parent_structure_rows", 0), errors="coerce"
    ).fillna(0).astype(int)

    out["ninth_sector_is_parent"] = out["ninth_sector"].map(ninth_sector_parent_lookup).fillna(False).astype(bool)
    out["ninth_sector_is_leaf"] = out["ninth_sector"].fillna("").astype(str).str.strip().ne("") & (~out["ninth_sector_is_parent"])
    out["ninth_fuel_is_parent"] = out["ninth_fuel"].map(ninth_fuel_parent_lookup).fillna(False).astype(bool)
    out["ninth_fuel_is_leaf"] = out["ninth_fuel"].fillna("").astype(str).str.strip().ne("") & (~out["ninth_fuel_is_parent"])
    return out


def _extract_combined_balance_rows(
    *,
    ref_workbook_path: Path,
    tgt_workbook_path: Path,
    mapping_workbook_path: Path,
    codebook_path: Path,
    template_sheet: str,
) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    for workbook in [ref_workbook_path, tgt_workbook_path]:
        extractor = TemplateBalanceExtractor(
            template_sheet=template_sheet,
            mapping_pairs_path=mapping_workbook_path,
            codebook_path=codebook_path,
            reinterpret_fuel_rows_as_parent_sector=True,
        )
        extractor.load_mappings()
        selected = _list_balance_sheets(workbook)
        _, mapped, _, _, _ = extractor.extract_workbook(
            workbook,
            include_zero_values=False,
            sheet_name_filter=selected,
            convert_units_to_petajoule=True,
        )
        frames.append(mapped)
    return pd.concat(frames, ignore_index=True, sort=False)


def _build_leap_name_mapping_scaffold(mapped: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    working = mapped.copy()
    for col in [
        "leap_sector_name",
        "leap_fuel_name",
        "leap_sector",
        "leap_fuel",
        "esto_flow",
        "esto_product",
        "mapping_status",
    ]:
        if col not in working.columns:
            working[col] = ""
        working[col] = working[col].fillna("").astype(str).str.strip()
    working["value_petajoule"] = pd.to_numeric(
        working.get("value_petajoule", working.get("value", pd.NA)),
        errors="coerce",
    )

    pair_keys = ["leap_sector_name", "leap_fuel_name"]
    grouped = (
        working.groupby(pair_keys, as_index=False)
        .agg(
            row_count=("leap_sector_name", "size"),
            value_pj_sum=("value_petajoule", "sum"),
            value_pj_abs_sum=("value_petajoule", lambda s: s.abs().sum()),
            leap_sector_codes=("leap_sector", lambda s: "|".join(sorted(set([v for v in s if v])))),
            leap_fuel_codes=("leap_fuel", lambda s: "|".join(sorted(set([v for v in s if v])))),
        )
        .reset_index(drop=True)
    )

    pairs = (
        working[["leap_sector_name", "leap_fuel_name", "esto_flow", "esto_product"]]
        .drop_duplicates()
        .copy()
    )
    non_empty = pairs[pairs["esto_flow"].ne("") & pairs["esto_product"].ne("")]
    pair_counts = (
        non_empty.groupby(pair_keys, as_index=False)
        .size()
        .rename(columns={"size": "esto_pair_count"})
    )
    single_pair = non_empty.groupby(pair_keys, as_index=False).agg(
        esto_flow=("esto_flow", _coalesce_unique),
        esto_product=("esto_product", _coalesce_unique),
    )

    out = grouped.merge(pair_counts, on=pair_keys, how="left").merge(single_pair, on=pair_keys, how="left")
    out["esto_pair_count"] = pd.to_numeric(out["esto_pair_count"], errors="coerce").fillna(0).astype(int)
    out["esto_flow"] = out["esto_flow"].fillna("").astype(str)
    out["esto_product"] = out["esto_product"].fillna("").astype(str)

    out["mapping_status"] = "unmapped"
    out.loc[out["esto_pair_count"].eq(1), "mapping_status"] = "mapped"
    out.loc[out["esto_pair_count"].gt(1), "mapping_status"] = "ambiguous_multiple_pairs"
    out.loc[out["esto_pair_count"].ne(1), ["esto_flow", "esto_product"]] = ""
    out["mapping_method"] = "balance_existing_observed_pairs"

    # Product-only backfill (independent of flow): infer by leap_fuel_name across all rows.
    product_by_fuel = (
        non_empty.groupby("leap_fuel_name", as_index=False)["esto_product"]
        .agg(lambda s: sorted(set([str(v).strip() for v in s if str(v).strip()])))
        .rename(columns={"esto_product": "candidate_products"})
    )
    product_by_fuel["unique_product"] = product_by_fuel["candidate_products"].apply(
        lambda vals: vals[0] if len(vals) == 1 else ""
    )
    product_lookup = dict(
        zip(
            product_by_fuel["leap_fuel_name"].astype(str),
            product_by_fuel["unique_product"].astype(str),
        )
    )
    missing_product = out["esto_product"].fillna("").astype(str).str.strip().eq("")
    inferred_product = out["leap_fuel_name"].map(product_lookup).fillna("").astype(str)
    fill_mask = missing_product & inferred_product.str.strip().ne("")
    out.loc[fill_mask, "esto_product"] = inferred_product.loc[fill_mask]
    out.loc[fill_mask & out["mapping_status"].eq("unmapped"), "mapping_status"] = "product_only_mapped"
    out.loc[fill_mask, "mapping_method"] = out.loc[fill_mask, "mapping_method"].astype(str) + "|fuel_name_product_only"

    out = _append_mapping_cardinality(
        out=out,
        relations=non_empty,
        left_keys=pair_keys,
        right_keys=["esto_flow", "esto_product"],
    )
    out = out.sort_values(["leap_sector_name", "leap_fuel_name"], kind="mergesort").reset_index(drop=True)

    ambiguous = out[out["mapping_status"].eq("ambiguous_multiple_pairs")][pair_keys]
    ambiguous_pairs = non_empty.merge(ambiguous, on=pair_keys, how="inner").sort_values(
        ["leap_sector_name", "leap_fuel_name", "esto_flow", "esto_product"],
        kind="mergesort",
    )
    return out, ambiguous_pairs


def _build_leap_code_mapping_scaffold(mapped: pd.DataFrame) -> pd.DataFrame:
    working = mapped.copy()
    for col in [
        "leap_sector",
        "leap_fuel",
        "leap_sector_name",
        "leap_fuel_name",
        "esto_flow",
        "esto_product",
    ]:
        if col not in working.columns:
            working[col] = ""
        working[col] = working[col].fillna("").astype(str).str.strip()
    working["value_petajoule"] = pd.to_numeric(
        working.get("value_petajoule", working.get("value", pd.NA)),
        errors="coerce",
    )
    working = working[working["leap_sector"].ne("") & working["leap_fuel"].ne("")].copy()

    keys = ["leap_sector", "leap_fuel"]
    grouped = (
        working.groupby(keys, as_index=False)
        .agg(
            leap_sector_name=("leap_sector_name", _coalesce_unique),
            leap_fuel_name=("leap_fuel_name", _coalesce_unique),
            row_count=("leap_sector", "size"),
            value_pj_sum=("value_petajoule", "sum"),
            value_pj_abs_sum=("value_petajoule", lambda s: s.abs().sum()),
        )
        .reset_index(drop=True)
    )

    pairs = working[keys + ["esto_flow", "esto_product"]].drop_duplicates()
    non_empty = pairs[pairs["esto_flow"].ne("") & pairs["esto_product"].ne("")]
    pair_counts = non_empty.groupby(keys, as_index=False).size().rename(columns={"size": "esto_pair_count"})
    single_pair = non_empty.groupby(keys, as_index=False).agg(
        esto_flow=("esto_flow", _coalesce_unique),
        esto_product=("esto_product", _coalesce_unique),
    )
    out = grouped.merge(pair_counts, on=keys, how="left").merge(single_pair, on=keys, how="left")
    out["esto_pair_count"] = pd.to_numeric(out["esto_pair_count"], errors="coerce").fillna(0).astype(int)
    out["esto_flow"] = out["esto_flow"].fillna("").astype(str)
    out["esto_product"] = out["esto_product"].fillna("").astype(str)
    out["mapping_status"] = "unmapped"
    out.loc[out["esto_pair_count"].eq(1), "mapping_status"] = "mapped"
    out.loc[out["esto_pair_count"].gt(1), "mapping_status"] = "ambiguous_multiple_pairs"
    out.loc[out["esto_pair_count"].ne(1), ["esto_flow", "esto_product"]] = ""
    out["mapping_method"] = "balance_existing_observed_pairs"

    # Product-only backfill (independent of flow): infer by leap_fuel code across all rows.
    product_by_fuel_code = (
        non_empty.groupby("leap_fuel", as_index=False)["esto_product"]
        .agg(lambda s: sorted(set([str(v).strip() for v in s if str(v).strip()])))
        .rename(columns={"esto_product": "candidate_products"})
    )
    product_by_fuel_code["unique_product"] = product_by_fuel_code["candidate_products"].apply(
        lambda vals: vals[0] if len(vals) == 1 else ""
    )
    product_lookup = dict(
        zip(
            product_by_fuel_code["leap_fuel"].astype(str),
            product_by_fuel_code["unique_product"].astype(str),
        )
    )
    missing_product = out["esto_product"].fillna("").astype(str).str.strip().eq("")
    inferred_product = out["leap_fuel"].map(product_lookup).fillna("").astype(str)
    fill_mask = missing_product & inferred_product.str.strip().ne("")
    out.loc[fill_mask, "esto_product"] = inferred_product.loc[fill_mask]
    out.loc[fill_mask & out["mapping_status"].eq("unmapped"), "mapping_status"] = "product_only_mapped"
    out.loc[fill_mask, "mapping_method"] = out.loc[fill_mask, "mapping_method"].astype(str) + "|fuel_code_product_only"

    out = _append_mapping_cardinality(
        out=out,
        relations=non_empty,
        left_keys=keys,
        right_keys=["esto_flow", "esto_product"],
    )
    return out.sort_values(["leap_sector", "leap_fuel"], kind="mergesort").reset_index(drop=True)


def _build_leap_name_to_ninth_scaffold(mapped: pd.DataFrame) -> pd.DataFrame:
    working = mapped.copy()
    for col in ["leap_sector_name", "leap_fuel_name", "leap_sector", "leap_fuel"]:
        if col not in working.columns:
            working[col] = ""
        working[col] = working[col].fillna("").astype(str).str.strip()
    working["value_petajoule"] = pd.to_numeric(working.get("value_petajoule", working.get("value", pd.NA)), errors="coerce")

    keys = ["leap_sector_name", "leap_fuel_name"]
    grouped = (
        working.groupby(keys, as_index=False)
        .agg(
            row_count=("leap_sector_name", "size"),
            value_pj_sum=("value_petajoule", "sum"),
            value_pj_abs_sum=("value_petajoule", lambda s: s.abs().sum()),
            leap_sector_codes=("leap_sector", lambda s: "|".join(sorted(set([v for v in s if v])))),
            leap_fuel_codes=("leap_fuel", lambda s: "|".join(sorted(set([v for v in s if v])))),
        )
        .reset_index(drop=True)
    )
    pairs = working[keys + ["leap_sector", "leap_fuel"]].drop_duplicates()
    non_empty = pairs[pairs["leap_sector"].ne("") & pairs["leap_fuel"].ne("")]
    pair_counts = non_empty.groupby(keys, as_index=False).size().rename(columns={"size": "ninth_pair_count"})
    single_pair = non_empty.groupby(keys, as_index=False).agg(
        ninth_sector=("leap_sector", _coalesce_unique),
        ninth_fuel=("leap_fuel", _coalesce_unique),
    )
    out = grouped.merge(pair_counts, on=keys, how="left").merge(single_pair, on=keys, how="left")
    out["ninth_pair_count"] = pd.to_numeric(out["ninth_pair_count"], errors="coerce").fillna(0).astype(int)
    out["ninth_sector"] = out["ninth_sector"].fillna("").astype(str)
    out["ninth_fuel"] = out["ninth_fuel"].fillna("").astype(str)
    out["mapping_status"] = "unmapped"
    out.loc[out["ninth_pair_count"].eq(1), "mapping_status"] = "mapped"
    out.loc[out["ninth_pair_count"].gt(1), "mapping_status"] = "ambiguous_multiple_pairs"
    out.loc[out["ninth_pair_count"].ne(1), ["ninth_sector", "ninth_fuel"]] = ""
    out["mapping_method"] = "observed_leap_code_pair"
    out = _append_mapping_cardinality(
        out=out,
        relations=non_empty.rename(columns={"leap_sector": "ninth_sector", "leap_fuel": "ninth_fuel"}),
        left_keys=keys,
        right_keys=["ninth_sector", "ninth_fuel"],
    )
    return out.sort_values(keys, kind="mergesort").reset_index(drop=True)


def _build_leap_code_to_ninth_scaffold(mapped: pd.DataFrame) -> pd.DataFrame:
    working = mapped.copy()
    for col in ["leap_sector", "leap_fuel", "leap_sector_name", "leap_fuel_name"]:
        if col not in working.columns:
            working[col] = ""
        working[col] = working[col].fillna("").astype(str).str.strip()
    working["value_petajoule"] = pd.to_numeric(working.get("value_petajoule", working.get("value", pd.NA)), errors="coerce")
    working = working[working["leap_sector"].ne("") & working["leap_fuel"].ne("")].copy()
    out = (
        working.groupby(["leap_sector", "leap_fuel"], as_index=False)
        .agg(
            leap_sector_name=("leap_sector_name", _coalesce_unique),
            leap_fuel_name=("leap_fuel_name", _coalesce_unique),
            row_count=("leap_sector", "size"),
            value_pj_sum=("value_petajoule", "sum"),
            value_pj_abs_sum=("value_petajoule", lambda s: s.abs().sum()),
        )
        .reset_index(drop=True)
    )
    out["ninth_sector"] = out["leap_sector"]
    out["ninth_fuel"] = out["leap_fuel"]
    out["mapping_status"] = "mapped"
    out["mapping_method"] = "identity_code_pair"
    out["mapping_scope"] = "pair"
    out["mapping_cardinality"] = "one_to_one"
    out["left_to_right_count"] = 1
    out["right_to_left_count"] = 1
    out["left_to_product_count"] = 1
    out["product_to_left_count"] = 1
    return out.sort_values(["leap_sector", "leap_fuel"], kind="mergesort").reset_index(drop=True)


def _build_esto_coverage_checks(
    *,
    leap_name_map: pd.DataFrame,
    esto_table_path: Path,
    economy: str,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    esto = pd.read_csv(esto_table_path)
    for col in ["economy", "flows", "products"]:
        if col not in esto.columns:
            raise KeyError(f"ESTO table missing required column: {col}")
    year_cols = [c for c in esto.columns if str(c).isdigit() and len(str(c)) == 4]
    if not year_cols:
        raise ValueError("ESTO table has no 4-digit year columns.")

    econ_norm = economy.replace("_", "").upper()
    working = esto[esto["economy"].astype(str).str.replace("_", "", regex=False).str.upper().eq(econ_norm)].copy()
    working["value_abs_sum"] = working[year_cols].apply(pd.to_numeric, errors="coerce").fillna(0).abs().sum(axis=1)
    with_data = working[working["value_abs_sum"].gt(0)].copy()

    esto_pairs_with_data = (
        with_data[["flows", "products", "value_abs_sum"]]
        .groupby(["flows", "products"], as_index=False)
        .agg(value_abs_sum=("value_abs_sum", "sum"))
        .rename(columns={"flows": "esto_flow", "products": "esto_product"})
        .sort_values(["esto_flow", "esto_product"], kind="mergesort")
        .reset_index(drop=True)
    )

    mapped_pairs = (
        leap_name_map[["esto_flow", "esto_product"]]
        .copy()
        .fillna("")
    )
    mapped_pairs = mapped_pairs[mapped_pairs["esto_flow"].astype(str).str.strip().ne("") & mapped_pairs["esto_product"].astype(str).str.strip().ne("")].drop_duplicates()

    missing_esto = esto_pairs_with_data.merge(mapped_pairs, on=["esto_flow", "esto_product"], how="left", indicator=True)
    missing_esto = missing_esto[missing_esto["_merge"].eq("left_only")].drop(columns=["_merge"]).reset_index(drop=True)

    mapped_not_in_esto = mapped_pairs.merge(
        esto_pairs_with_data[["esto_flow", "esto_product"]],
        on=["esto_flow", "esto_product"],
        how="left",
        indicator=True,
    )
    mapped_not_in_esto = mapped_not_in_esto[mapped_not_in_esto["_merge"].eq("left_only")].drop(columns=["_merge"]).reset_index(drop=True)
    return esto_pairs_with_data, missing_esto, mapped_not_in_esto


#%%
BALANCE_EXPORT_ECONOMY = "20_USA"
REF_BALANCE_EXPORT_DATE_ID: str | None = None
TGT_BALANCE_EXPORT_DATE_ID: str | None = None
REF_WORKBOOK_PATH = resolve_balance_export_workbook(
    economy=BALANCE_EXPORT_ECONOMY,
    scenario="REF",
    date_id=REF_BALANCE_EXPORT_DATE_ID,
)
TGT_WORKBOOK_PATH = resolve_balance_export_workbook(
    economy=BALANCE_EXPORT_ECONOMY,
    scenario="TGT",
    date_id=TGT_BALANCE_EXPORT_DATE_ID,
)
CODEBOOK_PATH = _resolve("config/sector_fuel_codes_to_names.xlsx")
MAPPING_WORKBOOK_PATH = _resolve("config/leap_to_esto_balance_full_mapping_slim.xlsx")
ESTO_TABLE_PATH = _resolve("data/00APEC_2025_low_with_subtotals.csv")
OUTPUT_DIR = MAPPINGS_ROOT / "leap_balance_mapping_scaffold"
OUTPUT_WORKBOOK_PATH = _resolve("config/leap_to_esto_balance_full_mapping_slim.xlsx")

ECONOMY = "20USA"
TEMPLATE_SHEET = "EBal|2060"


def run_workflow() -> dict[str, Any]:
    archive_config_dir_once_per_day()
    layout = build_workflow_output_layout(OUTPUT_DIR)

    mapped = _extract_combined_balance_rows(
        ref_workbook_path=REF_WORKBOOK_PATH,
        tgt_workbook_path=TGT_WORKBOOK_PATH,
        mapping_workbook_path=MAPPING_WORKBOOK_PATH,
        codebook_path=CODEBOOK_PATH,
        template_sheet=TEMPLATE_SHEET,
    )
    mapped = _drop_total_rows(mapped)
    leap_parent_flags = _derive_leap_parent_flags(mapped)
    esto_flow_parent_lookup, esto_product_parent_lookup = _derive_esto_parent_lookups(ESTO_TABLE_PATH)
    ninth_sector_parent_lookup, ninth_fuel_parent_lookup = _derive_ninth_parent_lookups(CODEBOOK_PATH)

    leap_name_map, ambiguous_name_pairs = _build_leap_name_mapping_scaffold(mapped)
    leap_code_map = _build_leap_code_mapping_scaffold(mapped)
    leap_name_to_ninth = _build_leap_name_to_ninth_scaffold(mapped)
    leap_code_to_ninth = _build_leap_code_to_ninth_scaffold(mapped)

    leap_name_map = _apply_parent_flags_to_esto_mapping(
        leap_name_map,
        leap_parent_flags=leap_parent_flags,
        esto_flow_parent_lookup=esto_flow_parent_lookup,
        esto_product_parent_lookup=esto_product_parent_lookup,
    )
    leap_code_map = _apply_parent_flags_to_esto_mapping(
        leap_code_map,
        leap_parent_flags=leap_parent_flags,
        esto_flow_parent_lookup=esto_flow_parent_lookup,
        esto_product_parent_lookup=esto_product_parent_lookup,
    )
    leap_name_to_ninth = _apply_parent_flags_to_ninth_mapping(
        leap_name_to_ninth,
        leap_parent_flags=leap_parent_flags,
        ninth_sector_parent_lookup=ninth_sector_parent_lookup,
        ninth_fuel_parent_lookup=ninth_fuel_parent_lookup,
    )
    leap_code_to_ninth = _apply_parent_flags_to_ninth_mapping(
        leap_code_to_ninth,
        leap_parent_flags=leap_parent_flags,
        ninth_sector_parent_lookup=ninth_sector_parent_lookup,
        ninth_fuel_parent_lookup=ninth_fuel_parent_lookup,
    )

    leap_axis_sectors = (
        mapped[["leap_sector_name"]]
        .drop_duplicates()
        .rename(columns={"leap_sector_name": "leap_sector_name"})
        .sort_values(["leap_sector_name"], kind="mergesort")
        .reset_index(drop=True)
    )
    leap_axis_fuels = (
        mapped[["leap_fuel_name"]]
        .drop_duplicates()
        .rename(columns={"leap_fuel_name": "leap_fuel_name"})
        .sort_values(["leap_fuel_name"], kind="mergesort")
        .reset_index(drop=True)
    )

    esto_pairs_with_data, missing_esto_pairs_with_data, mapped_not_in_esto_data = _build_esto_coverage_checks(
        leap_name_map=leap_name_map,
        esto_table_path=ESTO_TABLE_PATH,
        economy=ECONOMY,
    )
    unmapped_leap_pairs = leap_name_map[leap_name_map["esto_flow"].eq("") | leap_name_map["esto_product"].eq("")].copy()
    unmapped_leap_to_ninth = leap_name_to_ninth[
        leap_name_to_ninth["ninth_sector"].eq("") | leap_name_to_ninth["ninth_fuel"].eq("")
    ].copy()

    summary = {
        "rows_mapped_input": int(len(mapped)),
        "unique_leap_name_pairs": int(len(leap_name_map)),
        "unique_leap_code_pairs": int(len(leap_code_map)),
        "mapped_name_pairs": int((leap_name_map["mapping_status"] == "mapped").sum()),
        "ambiguous_name_pairs": int((leap_name_map["mapping_status"] == "ambiguous_multiple_pairs").sum()),
        "unmapped_name_pairs": int((leap_name_map["mapping_status"] == "unmapped").sum()),
        "product_only_name_pairs": int((leap_name_map["mapping_status"] == "product_only_mapped").sum()),
        "one_to_one_name_pairs": int((leap_name_map["mapping_cardinality"] == "one_to_one").sum()),
        "one_to_many_name_pairs": int((leap_name_map["mapping_cardinality"] == "one_to_many").sum()),
        "many_to_one_name_pairs": int((leap_name_map["mapping_cardinality"] == "many_to_one").sum()),
        "many_to_many_name_pairs": int((leap_name_map["mapping_cardinality"] == "many_to_many").sum()),
        "leap_parent_name_pairs": int(leap_name_map["leap_is_parent"].sum()),
        "leap_leaf_name_pairs": int(leap_name_map["leap_is_leaf"].sum()),
        "unique_leap_name_to_ninth_pairs": int(len(leap_name_to_ninth)),
        "unmapped_name_to_ninth_pairs": int(len(unmapped_leap_to_ninth)),
        "esto_pairs_with_data": int(len(esto_pairs_with_data)),
        "missing_esto_pairs_with_data": int(len(missing_esto_pairs_with_data)),
        "mapped_pairs_not_in_esto_data": int(len(mapped_not_in_esto_data)),
    }

    with pd.ExcelWriter(OUTPUT_WORKBOOK_PATH, engine="openpyxl") as writer:
        pd.DataFrame([summary]).to_excel(writer, sheet_name="summary", index=False)
        leap_code_map.to_excel(writer, sheet_name="leap_code_to_esto_pair", index=False)
        leap_name_map.to_excel(writer, sheet_name="leap_name_to_esto_pair", index=False)
        leap_code_to_ninth.to_excel(writer, sheet_name="leap_code_to_ninth_pair", index=False)
        leap_name_to_ninth.to_excel(writer, sheet_name="leap_name_to_ninth_pair", index=False)
        unmapped_leap_pairs.to_excel(writer, sheet_name="unmapped_leap_pairs", index=False)
        unmapped_leap_to_ninth.to_excel(writer, sheet_name="unmapped_leap_to_ninth_pairs", index=False)
        ambiguous_name_pairs.to_excel(writer, sheet_name="ambiguous_name_pairs", index=False)
        leap_axis_sectors.to_excel(writer, sheet_name="leap_axis_sectors", index=False)
        leap_axis_fuels.to_excel(writer, sheet_name="leap_axis_fuels", index=False)
        missing_esto_pairs_with_data.to_excel(writer, sheet_name="missing_esto_with_data", index=False)
        mapped_not_in_esto_data.to_excel(writer, sheet_name="mapped_not_in_esto_data", index=False)
        esto_pairs_with_data.to_excel(writer, sheet_name="esto_pairs_with_data", index=False)

    summary_json_path = layout.runtime / "leap_balance_mapping_scaffold_summary.json"
    summary_json_path.write_text(json.dumps(summary, ensure_ascii=True, indent=2), encoding="utf-8")
    missing_esto_csv = layout.checks / "missing_esto_pairs_with_data.csv"
    missing_esto_pairs_with_data.to_csv(missing_esto_csv, index=False)
    unmapped_leap_csv = layout.checks / "unmapped_leap_pairs.csv"
    unmapped_leap_pairs.to_csv(unmapped_leap_csv, index=False)
    manifest_path = write_output_manifest(
        out_dir=layout.root,
        primary_outputs={"output_workbook": str(OUTPUT_WORKBOOK_PATH)},
        supporting_outputs={
            "summary_json": str(summary_json_path),
            "missing_esto_pairs_with_data_csv": str(missing_esto_csv),
            "unmapped_leap_pairs_csv": str(unmapped_leap_csv),
        },
        primary_output_descriptions={
            "output_workbook": "Primary balance-mapping scaffold workbook with authored scaffold sheets.",
        },
        supporting_output_descriptions={
            "summary_json": "Run summary for the balance mapping scaffold workflow.",
            "missing_esto_pairs_with_data_csv": "ESTO pairs with data that remain unmapped in the scaffold.",
            "unmapped_leap_pairs_csv": "LEAP pairs that still lack ESTO mappings in the scaffold.",
        },
        notes=[
            "The scaffold workbook stays at the workflow root.",
            "Audit CSVs and summary JSON live under supporting_files/.",
        ],
    )

    return {
        "output_workbook": str(OUTPUT_WORKBOOK_PATH),
        "summary_json": str(summary_json_path),
        "missing_esto_pairs_with_data_csv": str(missing_esto_csv),
        "unmapped_leap_pairs_csv": str(unmapped_leap_csv),
        "summary": summary,
        "output_manifest_json": str(manifest_path),
    }


#%%
RUN_WORKFLOW = False
WORKFLOW_RESULT: dict[str, Any] | None = None
if RUN_WORKFLOW:
    WORKFLOW_RESULT = run_workflow()
    print("[OK] LEAP balance mapping scaffold workflow complete.")
    for key, value in WORKFLOW_RESULT.items():
        print(f"- {key}: {value}")
