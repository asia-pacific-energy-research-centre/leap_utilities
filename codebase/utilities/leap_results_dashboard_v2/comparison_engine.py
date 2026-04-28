from __future__ import annotations

from functools import lru_cache
import pandas as pd

from codebase.utilities.master_config import config_table_exists, read_config_table

from codebase.utilities.leap_results_dashboard_utils import (
    _collapse_base_family_rows_for_display,
    build_comparisons,
)
from .pathing import resolve_path


def _with_measure(keys: list[str], frame: pd.DataFrame) -> list[str]:
    return keys + (["measure"] if "measure" in frame.columns else [])


def _chart_scope_keys(keys: list[str], frame: pd.DataFrame) -> list[str]:
    """Use explicit chart_group_key for chart-level grouping, falling back to legacy sheet."""
    has_chart_group = (
        "chart_group_key" in frame.columns
        and frame["chart_group_key"].fillna("").astype(str).str.strip().ne("").any()
    )
    scoped = ["chart_group_key" if has_chart_group and key == "sheet" else key for key in keys]
    return _with_measure(scoped, frame)


def _hierarchy_descriptor_cols(frame: pd.DataFrame) -> list[str]:
    return [
        col
        for col in ["sheet", "page_key", "page_label", "chart_group_label"]
        if col in frame.columns
    ]

from .mapping_engine import annotate_mapping_status
from .models import DashboardV2Settings

DEFAULT_X_HIERARCHY_OVERRIDES = "config/leap_results_x_hierarchy_overrides.csv"


def _to_bool_series(values: object, default: bool = True) -> pd.Series:
    if values is None:
        return pd.Series(dtype="bool")
    if isinstance(values, pd.Series):
        raw = values.copy()
    else:
        raw = pd.Series(values)
    if raw.dtype == bool:
        return raw.fillna(default).astype(bool)
    text = raw.fillna(default).astype(str).str.strip().str.lower()
    true_tokens = {"1", "true", "yes", "y", "t"}
    false_tokens = {"0", "false", "no", "n", "f"}
    out = pd.Series(default, index=raw.index, dtype="bool")
    out.loc[text.isin(true_tokens)] = True
    out.loc[text.isin(false_tokens)] = False
    return out


def _clean_key(value: object) -> str:
    text = str(value or "").strip()
    return "" if text.lower() == "nan" else text


def _hierarchy_prefix_tokens(code: object) -> list[str]:
    text = _clean_key(code)
    if not text:
        return []
    out: list[str] = []
    for token in text.split("_"):
        tok = token.strip()
        if not tok:
            break
        if tok.isdigit() or tok.lower() == "x":
            out.append(tok.lower())
            continue
        break
    return out


def _numeric_prefix_before_x(code: object) -> list[str]:
    tokens = _hierarchy_prefix_tokens(code)
    out: list[str] = []
    for token in tokens:
        if token == "x":
            break
        out.append(token)
    return out


@lru_cache(maxsize=1)
def _load_x_hierarchy_overrides() -> dict[tuple[str, str], dict[str, object]]:
    path = resolve_path(DEFAULT_X_HIERARCHY_OVERRIDES)
    if not config_table_exists(path):
        return {}
    df = read_config_table(path)
    df.columns = [str(col).strip().lower() for col in df.columns]
    for col in ["domain", "code", "resolved_level", "hierarchy_behavior", "parent_code", "notes"]:
        if col not in df.columns:
            df[col] = ""
    df["domain"] = df["domain"].fillna("").astype(str).str.strip().str.lower()
    df["code"] = df["code"].fillna("").astype(str).str.strip()
    df["resolved_level"] = pd.to_numeric(df["resolved_level"], errors="coerce").fillna(0).astype(int)
    df["hierarchy_behavior"] = df["hierarchy_behavior"].fillna("").astype(str).str.strip().str.lower()
    df["parent_code"] = df["parent_code"].fillna("").astype(str).str.strip()
    df["notes"] = df["notes"].fillna("").astype(str).str.strip()
    df = df[df["domain"].isin({"sector", "fuel"}) & df["code"].ne("")].copy()
    return {
        (str(row.domain), str(row.code)): {
            "resolved_level": int(row.resolved_level),
            "hierarchy_behavior": str(row.hierarchy_behavior),
            "parent_code": str(row.parent_code),
            "notes": str(row.notes),
        }
        for row in df.itertuples(index=False)
    }


def _x_override(domain: str, code: object) -> dict[str, object]:
    return _load_x_hierarchy_overrides().get((str(domain).strip().lower(), _clean_key(code)), {})


def _code_depth(code: object) -> int:
    return sum(1 for token in _hierarchy_prefix_tokens(code) if token != "x")


def _effective_code_depth(code: object, domain: str) -> int:
    override = _x_override(domain, code)
    level = int(override.get("resolved_level", 0) or 0)
    if level > 0:
        return level
    return _code_depth(code)


def _is_inferred_parent_code(parent_code: object, child_code: object, *, domain: str) -> bool:
    parent_override = _x_override(domain, parent_code)
    child_override = _x_override(domain, child_code)
    explicit_child_parent = _clean_key(child_override.get("parent_code", ""))
    if explicit_child_parent and explicit_child_parent == _clean_key(parent_code):
        return True

    parent_tokens = _hierarchy_prefix_tokens(parent_code)
    child_tokens = _hierarchy_prefix_tokens(child_code)
    if not parent_tokens or not child_tokens:
        return False

    parent_behavior = str(parent_override.get("hierarchy_behavior", "") or "").strip().lower()
    child_behavior = str(child_override.get("hierarchy_behavior", "") or "").strip().lower()
    if parent_behavior in {"leaf_bucket", "catch_all"}:
        return False
    if child_behavior == "catch_all":
        return False

    if any(token == "x" for token in parent_tokens):
        if parent_behavior != "parent_aggregate":
            return False
        prefix = _numeric_prefix_before_x(parent_code)
        if not prefix:
            return False
        child_depth = _effective_code_depth(child_code, domain)
        parent_depth = _effective_code_depth(parent_code, domain)
        return child_tokens[: len(prefix)] == prefix and child_depth > parent_depth

    if len(parent_tokens) >= len(child_tokens):
        return False
    return child_tokens[: len(parent_tokens)] == parent_tokens


def _comparator_meta_rows(frame: pd.DataFrame, mapping_status: pd.DataFrame | None) -> pd.DataFrame:
    comp = frame.copy()
    if "measure" not in comp.columns:
        comp["measure"] = ""
    comp["measure"] = comp["measure"].fillna("").astype(str)
    comp["fuel_label"] = comp["fuel_label"].fillna("").astype(str)
    comp["source"] = comp["source"].fillna("").astype(str)
    if "chart_group_key" not in comp.columns:
        comp["chart_group_key"] = ""
    comp["chart_group_key"] = comp["chart_group_key"].fillna("").astype(str).str.strip()

    ms = mapping_status.copy() if mapping_status is not None else pd.DataFrame()
    for col in [
        "sheet",
        "chart_group_key",
        "measure",
        "fuel_label",
        "sector_code_9th",
        "ninth_fuel_code",
        "projection_parent_sector_code",
        "mapping_note",
    ]:
        if col not in ms.columns:
            ms[col] = ""
        ms[col] = ms[col].fillna("").astype(str)
    mapping_meta_full = ms[
        [
            "sheet",
            "chart_group_key",
            "measure",
            "fuel_label",
            "sector_code_9th",
            "ninth_fuel_code",
            "projection_parent_sector_code",
            "mapping_note",
        ]
    ]
    mapping_meta = mapping_meta_full.drop_duplicates(subset=["chart_group_key", "measure", "fuel_label"], keep="first")
    if comp["chart_group_key"].ne("").any() and mapping_meta["chart_group_key"].ne("").any():
        comp = comp.merge(
            mapping_meta.drop(columns=["sheet"], errors="ignore"),
            on=["chart_group_key", "measure", "fuel_label"],
            how="left",
            suffixes=("", "_map"),
        )
    else:
        mapping_meta = mapping_meta_full.drop_duplicates(subset=["sheet", "measure", "fuel_label"], keep="first")
        comp = comp.merge(mapping_meta, on=["sheet", "measure", "fuel_label"], how="left", suffixes=("", "_map"))

    for col in ["sector_code_9th", "ninth_fuel_code", "projection_parent_sector_code", "mapping_note"]:
        comp[col] = comp.get(col, "").fillna("").astype(str).str.strip()
    comp["is_residual_parent_row"] = (
        comp["mapping_note"].str.lower().str.contains(
            "base parent product reduced by mapped child products",
            regex=False,
        )
        & comp["sheet"].str.lower().str.contains("refining", regex=False)
    )

    is_comparator = comp["source"].isin(
        ["base", "base_estimated", "base_mixed", "projection", "projection_estimated", "projection_mixed"]
    )
    comp["comparator_sector_key"] = ""
    comp["comparator_fuel_key"] = ""
    comp.loc[is_comparator, "comparator_sector_key"] = comp.loc[is_comparator, "sector_code_9th"]
    comp.loc[is_comparator, "comparator_fuel_key"] = comp.loc[is_comparator, "ninth_fuel_code"]
    comp["comparator_key_kind"] = "display_fuel"
    comp.loc[is_comparator, "comparator_key_kind"] = "comparator_exact_key"
    comp["exact_comparator_key"] = "__fuel__:" + comp["fuel_label"].astype(str)
    comp.loc[is_comparator, "exact_comparator_key"] = (
        comp.loc[is_comparator, "comparator_sector_key"].fillna("").astype(str).str.strip()
        + "|"
        + comp.loc[is_comparator, "comparator_fuel_key"].fillna("").astype(str).str.strip()
    )
    missing_exact = is_comparator & comp["exact_comparator_key"].eq("|")
    comp.loc[missing_exact, "exact_comparator_key"] = "__unmapped__:" + comp.loc[missing_exact, "fuel_label"].astype(str)

    comp["comparator_sector_depth"] = comp["comparator_sector_key"].map(lambda value: _effective_code_depth(value, "sector"))
    comp["comparator_fuel_depth"] = comp["comparator_fuel_key"].map(lambda value: _effective_code_depth(value, "fuel"))
    comp["comparator_parent_sector_key"] = comp["projection_parent_sector_code"].fillna("").astype(str).str.strip()
    comp["is_comparator_leaf_level"] = is_comparator & comp["comparator_parent_sector_key"].eq("")
    comp["comparator_sector_behavior"] = comp["comparator_sector_key"].map(
        lambda value: str(_x_override("sector", value).get("hierarchy_behavior", ""))
    )
    comp["comparator_fuel_behavior"] = comp["comparator_fuel_key"].map(
        lambda value: str(_x_override("fuel", value).get("hierarchy_behavior", ""))
    )
    return comp


def _drop_parent_child_comparator_overlaps(
    chart_rows: pd.DataFrame,
    mapping_status: pd.DataFrame | None,
) -> pd.DataFrame:
    """
    Remove parent-level comparator rows when child-level rows for the same
    rendered group are already present.

    This is an upstream guard for chart rendering: totals should be built from
    one resolved hierarchy level, not a mix of parent and child comparator rows.
    The relation is currently determined from explicit parent-fallback lineage
    (`projection_parent_sector_code`) carried in mapping_status.
    """
    if chart_rows.empty:
        return chart_rows.copy()

    comp = _comparator_meta_rows(chart_rows, mapping_status)
    if "used_in_total" in comp.columns:
        used = comp["used_in_total"].copy()
        if used.dtype != bool:
            used = used.fillna(True).astype(str).str.strip().str.lower().isin({"1", "true", "yes", "y", "t"})
        comp["used_in_total"] = used.astype(bool)
    else:
        comp["used_in_total"] = True
    non_total = comp[comp["fuel_label"].astype(str).ne("Total")].copy()
    if non_total.empty:
        return chart_rows.copy()

    comparator_sources = {"base", "base_estimated", "base_mixed", "projection", "projection_estimated", "projection_mixed"}
    candidate = non_total[
        non_total["source"].isin(comparator_sources)
        & non_total["comparator_sector_key"].ne("")
        & non_total["comparator_fuel_key"].ne("")
        & non_total["used_in_total"]
    ].copy()
    if candidate.empty:
        return chart_rows.copy()

    scope_keys = _with_measure(["economy", "sheet", "scenario", "source", "year"], candidate)
    drop_keys: list[dict[str, object]] = []
    for values, group in candidate.groupby(scope_keys, dropna=False):
        key_map = dict(zip(scope_keys, values if isinstance(values, tuple) else (values,), strict=False))
        rows = list(group.itertuples(index=False))
        for parent_row in rows:
            parent_sector = _clean_key(getattr(parent_row, "comparator_sector_key", ""))
            parent_fuel = _clean_key(getattr(parent_row, "comparator_fuel_key", ""))
            parent_is_residual = bool(getattr(parent_row, "is_residual_parent_row", False))
            if not parent_sector or not parent_fuel:
                continue
            if parent_is_residual:
                continue
            for child_row in rows:
                if parent_row == child_row:
                    continue
                child_sector = _clean_key(getattr(child_row, "comparator_sector_key", ""))
                child_fuel = _clean_key(getattr(child_row, "comparator_fuel_key", ""))
                child_parent_sector = _clean_key(getattr(child_row, "comparator_parent_sector_key", ""))
                if not child_sector or not child_fuel:
                    continue

                sector_parent_match = False
                if parent_fuel == child_fuel:
                    sector_parent_match = (
                        parent_sector == child_parent_sector
                        or _is_inferred_parent_code(parent_sector, child_sector, domain="sector")
                    )

                fuel_parent_match = False
                if parent_sector == child_sector:
                    fuel_parent_match = _is_inferred_parent_code(parent_fuel, child_fuel, domain="fuel")

                if sector_parent_match or fuel_parent_match:
                    drop_keys.append(
                        {
                            **{k: key_map[k] for k in ["economy", "sheet", "scenario", "source", "year"] if k in key_map},
                            "measure": key_map.get("measure", ""),
                            "fuel_label": getattr(parent_row, "fuel_label"),
                            "comparator_sector_key": parent_sector,
                            "comparator_fuel_key": parent_fuel,
                        }
                    )
                    break

    if not drop_keys:
        return chart_rows.copy()

    drop_df = pd.DataFrame(drop_keys).drop_duplicates()
    key_cols = [
        "economy",
        "sheet",
        "measure",
        "scenario",
        "source",
        "year",
        "fuel_label",
        "comparator_sector_key",
        "comparator_fuel_key",
    ]
    comp = comp.merge(drop_df.assign(_drop_parent_overlap=True), on=key_cols, how="left")
    drop_mask = comp["_drop_parent_overlap"].fillna(False).astype(bool)
    filtered = comp.loc[~drop_mask].drop(columns=["_drop_parent_overlap"], errors="ignore")

    original_cols = list(chart_rows.columns)
    return filtered[[col for col in original_cols if col in filtered.columns]].copy()


def _common_level_only_filter(comparison_long: pd.DataFrame, mapping_status: pd.DataFrame) -> pd.DataFrame:
    """Keep comparator rows only where both base and projection are available for the sheet/fuel/scenario group."""
    if comparison_long.empty:
        return comparison_long

    comp = comparison_long.copy()
    comp["value"] = pd.to_numeric(comp["value"], errors="coerce")

    availability_key = _with_measure(["sheet", "fuel_label", "scenario"], comp)
    availability = (
        comp[comp["source"].isin(["base", "base_estimated", "base_mixed", "projection", "projection_estimated", "projection_mixed"])]
        .assign(
            has_base=lambda d: d["source"].isin(["base", "base_estimated", "base_mixed"]) & d["value"].notna(),
            has_projection=lambda d: d["source"].isin(["projection", "projection_estimated", "projection_mixed"]) & d["value"].notna(),
        )
        .groupby(availability_key, as_index=False)[["has_base", "has_projection"]]
        .max()
    )
    availability["common_level_ok"] = availability["has_base"] & availability["has_projection"]

    comp = comp.merge(availability[availability_key + ["common_level_ok"]], on=availability_key, how="left")
    keep = comp["source"].eq("leap") | comp["common_level_ok"].fillna(False)
    filtered = comp.loc[keep].drop(columns=["common_level_ok"])

    if not mapping_status.empty:
        idx = mapping_status[_with_measure(["sheet", "fuel_label"], mapping_status)].drop_duplicates()
        idx = idx.merge(
            availability.groupby(_with_measure(["sheet", "fuel_label"], availability), as_index=False)["common_level_ok"].max(),
            on=_with_measure(["sheet", "fuel_label"], idx),
            how="left",
        )
        filtered_keys = idx[idx["common_level_ok"].fillna(False)][_with_measure(["sheet", "fuel_label"], idx)]
        if not filtered_keys.empty:
            # no-op marker use downstream via status annotation
            pass

    return filtered


def filter_full_comparator_chart_rows(comparison_long: pd.DataFrame) -> pd.DataFrame:
    """
    Keep chart rows only for sheet/fuel/scenario groups that have both
    base-family and projection-family non-null comparator values.
    """
    if comparison_long.empty:
        return comparison_long.copy()

    comp = comparison_long.copy()
    comp["value"] = pd.to_numeric(comp["value"], errors="coerce")
    key = _with_measure(["sheet", "fuel_label", "scenario"], comp)
    fam = (
        comp.groupby(key)
        .apply(
            lambda g: pd.Series(
                {
                    "has_base": g.loc[
                        g["source"].isin(["base", "base_estimated", "base_mixed"]),
                        "value",
                    ].notna().any(),
                    "has_projection": g.loc[
                        g["source"].isin(["projection", "projection_estimated", "projection_mixed"]),
                        "value",
                    ].notna().any(),
                }
            ),
            include_groups=False,
        )
        .reset_index()
    )
    fam["full_compare"] = fam["has_base"] & fam["has_projection"]
    keep = fam[fam["full_compare"]][key]
    if keep.empty:
        return comp.iloc[0:0].copy()
    return comp.merge(keep, on=key, how="inner")


def build_total_rows_for_charts(
    chart_rows: pd.DataFrame,
    mapping_status: pd.DataFrame,
) -> pd.DataFrame:
    """
    Build `fuel_label == Total` rows from strict chart rows.

    Comparator totals are summed directly from child rows.
    """
    if chart_rows.empty:
        return chart_rows.iloc[0:0].copy()

    df = chart_rows.copy()
    if "measure" not in df.columns:
        df["measure"] = ""
    df["measure"] = df["measure"].fillna("").astype(str)
    df["fuel_label"] = df["fuel_label"].astype(str)
    df["source"] = df["source"].astype(str)
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    if "used_in_total" not in df.columns:
        df["used_in_total"] = True
    df["used_in_total"] = _to_bool_series(df["used_in_total"], default=True)
    fuels = df[df["fuel_label"].ne("Total")].copy()
    if fuels.empty:
        return fuels
    fuels = _collapse_base_family_rows_for_display(fuels)
    if "used_in_total" not in fuels.columns:
        fuels["used_in_total"] = True
    fuels["used_in_total"] = _to_bool_series(fuels["used_in_total"], default=True)

    keys = _with_measure(["economy", "sheet", "scenario", "source", "year"], df)
    scope_keys = _with_measure(["economy", "sheet"], df)

    leap_totals = (
        fuels[fuels["source"].eq("leap")]
        .groupby(keys, as_index=False)["value"]
        .sum(min_count=1)
        .assign(fuel_label="Total")
    )

    projection_sources = ["projection", "projection_estimated", "projection_mixed"]
    projection_comp = fuels[fuels["source"].isin(projection_sources) & fuels["used_in_total"]].copy()
    if projection_comp.empty:
        projection_totals = projection_comp.iloc[0:0].copy()
    else:
        projection_totals = (
            projection_comp.groupby(keys, as_index=False)["value"]
            .sum(min_count=1)
            .assign(fuel_label="Total")
        )

    base_sources = ["base", "base_estimated", "base_mixed"]
    base_comp = fuels[fuels["source"].isin(base_sources) & fuels["used_in_total"]].copy()
    if base_comp.empty:
        base_totals = base_comp.iloc[0:0].copy()
    else:
        base_totals = (
            base_comp.groupby(_with_measure(["economy", "sheet", "source", "year"], df), as_index=False)["value"]
            .sum(min_count=1)
        )
        base_totals["scenario"] = ""
        base_totals["fuel_label"] = "Total"

    out = pd.concat([leap_totals, projection_totals, base_totals], ignore_index=True, sort=False)
    return out


def build_chart_line_mapping_ledger(
    chart_rows: pd.DataFrame,
    mapping_status: pd.DataFrame,
) -> pd.DataFrame:
    """
    Build a per-point ledger for rendered chart rows with mapping metadata.

    For normal fuel lines, metadata is taken directly from ``mapping_status``.
    For ``fuel_label == Total`` lines, metadata is derived from the displayed
    child rows and counts exact comparator identities, not value-based dedupe
    buckets.
    """
    if chart_rows.empty:
        return chart_rows.copy()

    ledger = chart_rows.copy()
    ledger["sheet"] = ledger["sheet"].astype(str)
    if "measure" not in ledger.columns:
        ledger["measure"] = ""
    ledger["measure"] = ledger["measure"].fillna("").astype(str)
    ledger["fuel_label"] = ledger["fuel_label"].astype(str)
    ledger["source"] = ledger["source"].astype(str)
    ledger["scenario"] = ledger["scenario"].astype(str)
    ledger["value"] = pd.to_numeric(ledger["value"], errors="coerce")
    ledger["is_total_line"] = ledger["fuel_label"].eq("Total")
    ledger["mapping_record_type"] = ledger["is_total_line"].map({True: "derived_total", False: "direct"})
    ledger["aggregate_group_key"] = pd.NA
    ledger["aggregate_group_member_count"] = pd.NA
    ledger["first_of_aggregate"] = False
    ledger["first_of_aggregate_or_non_aggregate"] = False
    ledger["exact_comparator_key"] = ""
    ledger["duplicate_exact_comparator_key_count"] = pd.NA
    ledger["comparator_sector_key"] = ""
    ledger["comparator_fuel_key"] = ""
    ledger["comparator_sector_depth"] = 0
    ledger["comparator_fuel_depth"] = 0
    ledger["comparator_parent_sector_key"] = ""
    ledger["is_comparator_leaf_level"] = False

    meta_cols = [
        "sector_code_9th",
        "ninth_fuel_code",
        "esto_flow",
        "esto_product",
        "mapped",
        "mapping_source",
        "flow_source",
        "fuel_source",
        "sector_match_method",
        "mapping_note",
        "projection_parent_fallback",
        "projection_parent_sector_code",
        "comparator_scope",
    ]
    ms = mapping_status.copy() if mapping_status is not None else pd.DataFrame()
    if ms.empty:
        for col in ["sheet", "fuel_label"] + meta_cols:
            if col not in ledger.columns:
                ledger[col] = ""
        return ledger

    if "chart_group_key" not in ledger.columns:
        ledger["chart_group_key"] = ""
    ledger["chart_group_key"] = ledger["chart_group_key"].fillna("").astype(str).str.strip()
    for col in ["sheet", "chart_group_key", "fuel_label"] + meta_cols:
        if col not in ms.columns:
            ms[col] = ""
    if "measure" not in ms.columns:
        ms["measure"] = ""
    ms["measure"] = ms["measure"].fillna("").astype(str)
    ms["chart_group_key"] = ms["chart_group_key"].fillna("").astype(str).str.strip()
    if ledger["chart_group_key"].ne("").any() and ms["chart_group_key"].ne("").any():
        mapping_meta = ms[["chart_group_key", "measure", "fuel_label"] + meta_cols].drop_duplicates(
            subset=["chart_group_key", "measure", "fuel_label"],
            keep="first",
        )
        ledger = ledger.merge(mapping_meta, on=["chart_group_key", "measure", "fuel_label"], how="left")
    else:
        mapping_meta = ms[["sheet", "measure", "fuel_label"] + meta_cols].drop_duplicates(
            subset=["sheet", "measure", "fuel_label"],
            keep="first",
        )
        ledger = ledger.merge(mapping_meta, on=["sheet", "measure", "fuel_label"], how="left")

    keys = _chart_scope_keys(["economy", "sheet", "scenario", "source", "year"], ledger)
    fuels = _comparator_meta_rows(ledger[~ledger["is_total_line"]].copy(), ms)
    if fuels.empty:
        return ledger

    helper_cols = [
        "sheet",
        "chart_group_key",
        "measure",
        "fuel_label",
        "source",
        "exact_comparator_key",
        "comparator_sector_key",
        "comparator_fuel_key",
        "comparator_sector_depth",
        "comparator_fuel_depth",
        "comparator_parent_sector_key",
        "is_comparator_leaf_level",
    ]
    helper_join_cols = ["chart_group_key", "measure", "fuel_label", "source"] if ledger["chart_group_key"].ne("").any() else ["sheet", "measure", "fuel_label", "source"]
    ledger = ledger.merge(
        fuels[helper_cols].drop_duplicates(
            subset=helper_join_cols,
            keep="first",
        ),
        on=helper_join_cols,
        how="left",
        suffixes=("", "_meta"),
    )
    for col in [
        "exact_comparator_key",
        "comparator_sector_key",
        "comparator_fuel_key",
        "comparator_parent_sector_key",
    ]:
        meta_col = f"{col}_meta"
        if meta_col in ledger.columns:
            ledger[col] = ledger[meta_col].combine_first(ledger[col])
            ledger = ledger.drop(columns=[meta_col], errors="ignore")
        ledger[col] = ledger[col].fillna("").astype(str)
    for col in ["comparator_sector_depth", "comparator_fuel_depth"]:
        meta_col = f"{col}_meta"
        if meta_col in ledger.columns:
            ledger[col] = ledger[meta_col].combine_first(ledger[col])
            ledger = ledger.drop(columns=[meta_col], errors="ignore")
        ledger[col] = pd.to_numeric(ledger[col], errors="coerce").fillna(0).astype(int)
    if "is_comparator_leaf_level_meta" in ledger.columns:
        ledger["is_comparator_leaf_level"] = ledger["is_comparator_leaf_level_meta"].combine_first(
            ledger["is_comparator_leaf_level"]
        )
        ledger = ledger.drop(columns=["is_comparator_leaf_level_meta"], errors="ignore")
    ledger["is_comparator_leaf_level"] = (
        ledger["is_comparator_leaf_level"].astype("boolean").fillna(False).astype(bool)
    )

    comp_sources = {"base", "base_estimated", "base_mixed", "projection", "projection_estimated", "projection_mixed"}
    leap_sources = {"leap"}

    leap_comp = (
        fuels[fuels["source"].isin(leap_sources)]
        .groupby(keys, as_index=False)["fuel_label"]
        .nunique()
        .rename(columns={"fuel_label": "total_component_bucket_count"})
    )

    comp = fuels[fuels["source"].isin(comp_sources)].copy()
    if comp.empty:
        comp_comp = pd.DataFrame(columns=keys + ["total_component_bucket_count"])
    else:
        comp_comp = (
            comp.groupby(keys, as_index=False)["exact_comparator_key"]
            .nunique()
            .rename(columns={"exact_comparator_key": "total_component_bucket_count"})
        )

    total_meta = pd.concat([leap_comp, comp_comp], ignore_index=True, sort=False)
    if not total_meta.empty:
        total_meta = (
            total_meta.groupby(keys, as_index=False)["total_component_bucket_count"]
            .max()
        )
        ledger = ledger.merge(total_meta, on=keys, how="left")
    else:
        ledger["total_component_bucket_count"] = pd.NA

    # Add explicit duplicate helpers keyed to exact comparator identity.
    non_total = ledger[~ledger["is_total_line"]].copy()
    if not non_total.empty:
        non_total = non_total.reset_index().rename(columns={"index": "__row_id"})
        non_total["aggregate_group_key"] = non_total["exact_comparator_key"].fillna("").astype(str)
        group_key_cols = keys + ["aggregate_group_key"]
        non_total["aggregate_group_member_count"] = non_total.groupby(group_key_cols)["fuel_label"].transform("size")
        non_total["duplicate_exact_comparator_key_count"] = non_total["aggregate_group_member_count"]
        non_total = non_total.sort_values(group_key_cols + ["fuel_label"], kind="mergesort")
        non_total["__group_order"] = non_total.groupby(group_key_cols).cumcount()
        non_total["first_of_aggregate"] = non_total["__group_order"].eq(0)
        non_total["first_of_aggregate_or_non_aggregate"] = (
            non_total["aggregate_group_member_count"].le(1) | non_total["first_of_aggregate"]
        )

        helper = non_total[
            [
                "__row_id",
                "aggregate_group_key",
                "aggregate_group_member_count",
                "duplicate_exact_comparator_key_count",
                "first_of_aggregate",
                "first_of_aggregate_or_non_aggregate",
            ]
        ].copy()
        ledger = ledger.reset_index().rename(columns={"index": "__row_id"})
        ledger = ledger.merge(helper, on="__row_id", how="left", suffixes=("", "_new"))
        ledger["aggregate_group_key"] = ledger["aggregate_group_key_new"].combine_first(ledger["aggregate_group_key"])
        ledger["aggregate_group_member_count"] = ledger["aggregate_group_member_count_new"].combine_first(
            ledger["aggregate_group_member_count"]
        )
        ledger["duplicate_exact_comparator_key_count"] = ledger["duplicate_exact_comparator_key_count_new"].combine_first(
            ledger["duplicate_exact_comparator_key_count"]
        )
        ledger["first_of_aggregate"] = (
            ledger["first_of_aggregate_new"]
            .combine_first(ledger["first_of_aggregate"])
            .astype("boolean")
            .fillna(False)
            .astype(bool)
        )
        ledger["first_of_aggregate_or_non_aggregate"] = (
            ledger["first_of_aggregate_or_non_aggregate_new"]
            .combine_first(ledger["first_of_aggregate_or_non_aggregate"])
            .astype("boolean")
            .fillna(False)
            .astype(bool)
        )
        ledger = ledger.drop(
            columns=[
                "__row_id",
                "aggregate_group_key_new",
                "aggregate_group_member_count_new",
                "duplicate_exact_comparator_key_count_new",
                "first_of_aggregate_new",
                "first_of_aggregate_or_non_aggregate_new",
            ],
            errors="ignore",
        )
    ledger["duplicate_exact_comparator_key_count"] = pd.to_numeric(
        ledger["duplicate_exact_comparator_key_count"], errors="coerce"
    ).fillna(0).astype(int)

    return ledger


def build_total_component_ledger(
    chart_rows: pd.DataFrame,
    mapping_status: pd.DataFrame,
) -> pd.DataFrame:
    """
    Build detailed component rows that feed each ``fuel_label == Total`` point.

    Comparator rows follow the rendered chart rule directly: sum displayed child
    rows, and flag repeated exact comparator identities instead of suppressing
    them with a value-based heuristic.
    """
    if chart_rows.empty:
        return pd.DataFrame(
            columns=[
                "economy",
                "sheet",
                "scenario",
                "source",
                "year",
                "dedupe_bucket",
                "exact_comparator_key",
                "ninth_fuel_code",
                "sector_code_9th",
                "projection_parent_sector_code",
                "sector_depth",
                "fuel_depth",
                "is_leaf_level",
                "member_fuel_label",
                "member_value",
                "component_selected_value",
                "is_selected_max",
                "component_included_in_total",
                "duplicate_exact_comparator_key_count",
            ]
        )

    df = chart_rows.copy()
    df["fuel_label"] = df["fuel_label"].astype(str)
    df["source"] = df["source"].astype(str)
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    fuels = df[df["fuel_label"].ne("Total")].copy()
    if fuels.empty:
        return pd.DataFrame(
            columns=[
                "economy",
                "sheet",
                "scenario",
                "source",
                "year",
                "dedupe_bucket",
                "exact_comparator_key",
                "ninth_fuel_code",
                "sector_code_9th",
                "projection_parent_sector_code",
                "sector_depth",
                "fuel_depth",
                "is_leaf_level",
                "member_fuel_label",
                "member_value",
                "component_selected_value",
                "is_selected_max",
                "component_included_in_total",
                "duplicate_exact_comparator_key_count",
            ]
        )

    keys = _chart_scope_keys(["economy", "sheet", "scenario", "source", "year"], df)
    descriptor_cols = [col for col in _hierarchy_descriptor_cols(df) if col not in keys]
    fuels = _comparator_meta_rows(fuels, mapping_status)

    leap = fuels[fuels["source"].eq("leap")].copy()
    if not leap.empty:
        leap["dedupe_bucket"] = "__fuel__:" + leap["fuel_label"].astype(str)
        leap["exact_comparator_key"] = leap["dedupe_bucket"]
        leap["ninth_fuel_code"] = ""
        leap["sector_code_9th"] = ""
        leap["projection_parent_sector_code"] = ""
        leap["sector_depth"] = 0
        leap["fuel_depth"] = 0
        leap["is_leaf_level"] = True
        leap["member_fuel_label"] = leap["fuel_label"].astype(str)
        leap["member_value"] = leap["value"]
        leap["component_selected_value"] = leap["value"]
        leap["is_selected_max"] = True
        leap["component_included_in_total"] = _to_bool_series(leap.get("used_in_total", True), default=True)
        leap["duplicate_exact_comparator_key_count"] = 1
        leap_rows = leap[
            keys
            + descriptor_cols
            + [
                "dedupe_bucket",
                "exact_comparator_key",
                "ninth_fuel_code",
                "sector_code_9th",
                "projection_parent_sector_code",
                "sector_depth",
                "fuel_depth",
                "is_leaf_level",
                "member_fuel_label",
                "member_value",
                "component_selected_value",
                "is_selected_max",
                "component_included_in_total",
                "duplicate_exact_comparator_key_count",
            ]
        ].copy()
    else:
        leap_rows = pd.DataFrame()

    comp_sources = {"base", "base_estimated", "base_mixed", "projection", "projection_estimated", "projection_mixed"}
    comp = fuels[fuels["source"].isin(comp_sources)].copy()
    if not comp.empty:
        comp["dedupe_bucket"] = comp["exact_comparator_key"]
        duplicate_counts = (
            comp.groupby(keys + ["exact_comparator_key"], as_index=False)["fuel_label"]
            .size()
            .rename(columns={"size": "duplicate_exact_comparator_key_count"})
        )
        comp = comp.merge(duplicate_counts, on=keys + ["exact_comparator_key"], how="left")
        comp["member_fuel_label"] = comp["fuel_label"].astype(str)
        comp["member_value"] = comp["value"]
        comp["component_selected_value"] = comp["value"]
        comp["is_selected_max"] = True
        comp["component_included_in_total"] = _to_bool_series(comp.get("used_in_total", True), default=True)
        comp["sector_code_9th"] = comp["comparator_sector_key"]
        comp["projection_parent_sector_code"] = comp["comparator_parent_sector_key"]
        comp["sector_depth"] = comp["comparator_sector_depth"]
        comp["fuel_depth"] = comp["comparator_fuel_depth"]
        comp["is_leaf_level"] = comp["is_comparator_leaf_level"]
        comp_rows = comp[
            keys
            + descriptor_cols
            + [
                "dedupe_bucket",
                "exact_comparator_key",
                "ninth_fuel_code",
                "sector_code_9th",
                "projection_parent_sector_code",
                "sector_depth",
                "fuel_depth",
                "is_leaf_level",
                "member_fuel_label",
                "member_value",
                "component_selected_value",
                "is_selected_max",
                "component_included_in_total",
                "duplicate_exact_comparator_key_count",
            ]
        ].copy()
    else:
        comp_rows = pd.DataFrame()

    out = pd.concat([leap_rows, comp_rows], ignore_index=True, sort=False)
    sort_cols = [col for col in ["page_key", "chart_group_key", "sheet", "scenario", "source", "year", "dedupe_bucket", "member_fuel_label"] if col in out.columns]
    out = out.sort_values(sort_cols).reset_index(drop=True)
    return out


def collapse_sheets_by_sector_name_for_charts(
    comparison_long: pd.DataFrame,
    sheet_map: pd.DataFrame,
) -> pd.DataFrame:
    """
    Collapse sheets that share the same final chart category into one chart namespace.

    This is intended for chart/dashboard rendering only. Rows from grouped sheets
    are rewritten so ``sheet`` becomes the shared final category label.
    """
    if comparison_long.empty or sheet_map.empty:
        return comparison_long.copy()

    sm = sheet_map.copy()
    for col in ["sheet_name", "sector_name", "final_category_name", "final category name", "notes"]:
        if col not in sm.columns:
            sm[col] = ""
        sm[col] = sm[col].fillna("").astype(str).str.strip()
    final_category = sm["final_category_name"].where(
        sm["final_category_name"].ne(""),
        sm["final category name"],
    )
    sm["grouped_category_name"] = final_category.where(
        final_category.ne(""),
        sm["sector_name"],
    )
    sm = sm[sm["sheet_name"].ne("") & sm["grouped_category_name"].ne("")].copy()
    # Keep derived loss/own-use sheets in their own namespace so dashboard
    # routing can place them under the dedicated Losses & own use section.
    # If these are collapsed to sector names (e.g., "Oil refineries"), they
    # are re-routed into transformation pages instead.
    sm = sm[~sm["sheet_name"].str.lower().str.endswith("_loss_own_use_total")].copy()
    if sm.empty:
        return comparison_long.copy()

    # Drop sector-name alias sheets when a final-category sheet exists for the
    # same sector code. This prevents duplicate charts like "Chemicals" and
    # "Chemical (incl. petrochemical)" when the latter is just a sector label.
    drop_alias_sheets: set[str] = set()
    if "sector_code_9th" in sm.columns:
        alias_cols = ["sheet_name", "sector_name", "final_category_name", "sector_code_9th"]
        for col in alias_cols:
            if col not in sm.columns:
                sm[col] = ""
            sm[col] = sm[col].fillna("").astype(str).str.strip()
        for _, row in sm.iterrows():
            sheet_name = str(row["sheet_name"]).strip()
            sector_name = str(row["sector_name"]).strip()
            final_name = str(row["final_category_name"]).strip()
            sector_code = str(row["sector_code_9th"]).strip()
            if not sheet_name or not sector_name or not final_name or not sector_code:
                continue
            if sheet_name != sector_name:
                continue
            if final_name == sheet_name:
                continue
            has_final_sheet = bool(
                sm[
                    (sm["sector_code_9th"].eq(sector_code))
                    & (sm["sheet_name"].eq(final_name))
                ].shape[0]
            )
            if has_final_sheet:
                drop_alias_sheets.add(sheet_name)
    if drop_alias_sheets:
        comp_drop = comparison_long.copy()
        comp_drop["sheet"] = comp_drop["sheet"].astype(str)
        comp_drop = comp_drop[~comp_drop["sheet"].isin(drop_alias_sheets)]
        comparison_long = comp_drop

    role_keywords = (
        "input",
        "inputs",
        "output",
        "outputs",
        "feedstock",
        "product",
    )
    protected_category_names: set[str] = set()
    for category_name, group in sm.groupby("grouped_category_name", dropna=False):
        if group["sheet_name"].nunique() <= 1:
            continue
        notes = group["notes"].astype(str).str.strip().str.lower()
        distinct_notes = {note for note in notes if note}
        has_role_specific_note = any(any(keyword in note for keyword in role_keywords) for note in distinct_notes)
        if has_role_specific_note and len(distinct_notes) > 1:
            protected_category_names.add(str(category_name))

    comp = comparison_long.copy()
    comp["sheet"] = comp["sheet"].astype(str)
    if "measure" not in comp.columns:
        comp["measure"] = ""
    comp["measure"] = comp["measure"].fillna("").astype(str)
    comp["source"] = comp["source"].astype(str)
    comp["value"] = pd.to_numeric(comp["value"], errors="coerce")
    has_atomic_allocated = "atomic_allocated" in comp.columns
    if has_atomic_allocated:
        comp["atomic_allocated"] = comp["atomic_allocated"].fillna(False).astype(bool)
    has_force_show_chart = "force_show_chart" in comp.columns
    if has_force_show_chart:
        comp["force_show_chart"] = comp["force_show_chart"].fillna(False).astype(bool)
    has_projection_parent_fallback = "projection_parent_fallback" in comp.columns
    if has_projection_parent_fallback:
        comp["projection_parent_fallback"] = comp["projection_parent_fallback"].fillna(False).astype(bool)
    has_used_in_total = "used_in_total" in comp.columns
    if has_used_in_total:
        comp["used_in_total"] = _to_bool_series(comp["used_in_total"], default=True)
    for col in ["sector_code_9th", "ninth_fuel_code", "projection_parent_sector_code", "comparator_scope"]:
        if col in comp.columns:
            comp[col] = comp[col].fillna("").astype(str).str.strip()

    def _effective_comparator_identity(frame: pd.DataFrame) -> pd.Series:
        if frame.empty:
            return pd.Series(dtype="object", index=frame.index)
        sector = frame["sector_code_9th"] if "sector_code_9th" in frame.columns else pd.Series("", index=frame.index)
        fuel = frame["ninth_fuel_code"] if "ninth_fuel_code" in frame.columns else pd.Series("", index=frame.index)
        parent = (
            frame["projection_parent_sector_code"]
            if "projection_parent_sector_code" in frame.columns
            else pd.Series("", index=frame.index)
        )
        parent_fallback = (
            frame["projection_parent_fallback"]
            if "projection_parent_fallback" in frame.columns
            else pd.Series(False, index=frame.index)
        )
        effective_sector = sector.where(~parent_fallback.astype(bool), parent)
        effective_sector = effective_sector.fillna("").astype(str).str.strip().str.lower()
        fuel = fuel.fillna("").astype(str).str.strip().str.lower()
        return effective_sector.where(
            effective_sector.ne("") | fuel.ne(""),
            "",
        ) + "|" + fuel.where(effective_sector.ne("") | fuel.ne(""), "")

    comp["__collapse_comparator_identity"] = _effective_comparator_identity(comp)

    key_cols = _with_measure(["economy", "sheet", "fuel_label", "scenario", "source", "year"], comp)

    grouped = (
        sm.groupby("grouped_category_name", as_index=False)["sheet_name"]
        .nunique()
        .rename(columns={"sheet_name": "sheet_count"})
    )
    out = comp.copy()
    # Reconcile:
    # 1) true grouped sectors (many child sheets -> one sector label), and
    # 2) one-to-one aliases only when both labels are already present in data.
    category_names: set[str] = {
        str(name)
        for name in grouped[grouped["sheet_count"] > 1]["grouped_category_name"].astype(str).tolist()
        if str(name) not in protected_category_names
    }
    labels_present = set(comp["sheet"].astype(str).tolist())
    single_rows = grouped[grouped["sheet_count"] == 1].copy()
    if not single_rows.empty:
        single_map = sm.merge(single_rows[["grouped_category_name"]], on="grouped_category_name", how="inner")
        both_present = single_map[
            single_map["sheet_name"].astype(str).isin(labels_present)
            & single_map["grouped_category_name"].astype(str).isin(labels_present)
        ]
        if not both_present.empty:
            category_names.update(
                str(name)
                for name in both_present["grouped_category_name"].astype(str).tolist()
                if str(name) not in protected_category_names
            )

    grouped_sheet_counts = {
        str(row["grouped_category_name"]): int(row["sheet_count"])
        for _, row in grouped.iterrows()
    }

    if category_names:
        map_rows = sm[sm["grouped_category_name"].isin(category_names)][["sheet_name", "grouped_category_name"]].drop_duplicates()

        # Build per-sector reconciliation to preserve existing parent rows while
        # adding child-sheet information.
        reconciled_parts: list[pd.DataFrame] = []
        drop_sheet_names: set[str] = set()
        drop_sheet_names.update(map_rows["sheet_name"].astype(str).tolist())
        drop_sheet_names.update(category_names)

        for category_name in sorted(category_names):
            child_sheets = set(
                map_rows.loc[map_rows["grouped_category_name"].eq(category_name), "sheet_name"].astype(str).tolist()
            )
            prefer_child_rows = grouped_sheet_counts.get(str(category_name), 0) > 1
            group_rows = comp[comp["sheet"].isin(child_sheets | {category_name})].copy()
            if group_rows.empty:
                continue
            group_rows["sheet"] = category_name
            group_rows["is_parent_row"] = comp.loc[group_rows.index, "sheet"].astype(str).eq(category_name)

            records: list[dict[str, object]] = []
            for values, g in group_rows.groupby(key_cols, dropna=False):
                key_map = dict(zip(key_cols, values if isinstance(values, tuple) else (values,), strict=False))
                economy = key_map.get("economy")
                sheet_name = key_map.get("sheet")
                fuel_label = key_map.get("fuel_label")
                scenario = key_map.get("scenario")
                source = key_map.get("source")
                year = key_map.get("year")
                src = str(source or "").strip().lower()
                parent_vals = pd.to_numeric(g.loc[g["is_parent_row"], "value"], errors="coerce").dropna()
                child_vals = pd.to_numeric(g.loc[~g["is_parent_row"], "value"], errors="coerce").dropna()
                all_vals = pd.to_numeric(g["value"], errors="coerce").dropna()
                allocated_rows = bool(g["atomic_allocated"].fillna(False).any()) if has_atomic_allocated else False

                if prefer_child_rows and child_vals.size:
                    if src == "leap":
                        # For LEAP grouped demand sectors, preserve additive totals:
                        # always sum displayed child sheets at the resolved level.
                        out_val = float(child_vals.sum())
                    else:
                        if allocated_rows:
                            out_val = float(child_vals.sum())
                        else:
                            uniq = child_vals.unique()
                            out_val = float(uniq[0]) if len(uniq) == 1 else float(child_vals.sum())
                elif parent_vals.size:
                    out_val = float(parent_vals.iloc[0])
                elif not all_vals.size:
                    out_val = float("nan")
                elif src == "leap":
                    # Preserve additive LEAP totals: sum all displayed components.
                    out_val = float(all_vals.sum())
                else:
                    # Comparator series often repeat across children; collapse duplicates.
                    if allocated_rows:
                        # Atomic-allocated rows already partition shared units across
                        # grouped sheets; sum to recover the parent comparator total.
                        out_val = float(all_vals.sum())
                    else:
                        comparator_ids = (
                            g["__collapse_comparator_identity"].fillna("").astype(str).str.strip()
                            if "__collapse_comparator_identity" in g.columns
                            else pd.Series("", index=g.index, dtype="object")
                        )
                        identified = g[comparator_ids.ne("")].copy()
                        unidentified = g[comparator_ids.eq("")].copy()
                        dedup_total = 0.0
                        saw_component = False
                        if not identified.empty:
                            identified["__collapse_comparator_identity"] = comparator_ids.loc[identified.index]
                            for _, ident_g in identified.groupby("__collapse_comparator_identity", dropna=False):
                                ident_vals = pd.to_numeric(ident_g["value"], errors="coerce").dropna()
                                if ident_vals.empty:
                                    continue
                                saw_component = True
                                uniq_ident = ident_vals.unique()
                                dedup_total += float(uniq_ident[0]) if len(uniq_ident) == 1 else float(ident_vals.iloc[0])
                        if not unidentified.empty:
                            unidentified_vals = pd.to_numeric(unidentified["value"], errors="coerce").dropna()
                            if not unidentified_vals.empty:
                                saw_component = True
                                uniq = unidentified_vals.unique()
                                dedup_total += float(uniq[0]) if len(uniq) == 1 else float(unidentified_vals.sum())
                        out_val = dedup_total if saw_component else float("nan")

                rec: dict[str, object] = {
                    "economy": economy,
                    "sheet": sheet_name,
                    "measure": key_map.get("measure", ""),
                    "fuel_label": fuel_label,
                    "scenario": scenario,
                    "source": source,
                    "year": year,
                    "value": out_val,
                }
                if has_atomic_allocated:
                    rec["atomic_allocated"] = allocated_rows
                if has_force_show_chart:
                    rec["force_show_chart"] = bool(g["force_show_chart"].fillna(False).any())
                if has_used_in_total:
                    rec["used_in_total"] = bool(_to_bool_series(g["used_in_total"], default=True).any())
                records.append(rec)

            reconciled_parts.append(pd.DataFrame(records))

        if reconciled_parts:
            reconciled = pd.concat(reconciled_parts, ignore_index=True, sort=False)
            base = comp[~comp["sheet"].isin(drop_sheet_names)].copy()
            out = pd.concat([base, reconciled], ignore_index=True, sort=False)
            out = out.drop_duplicates(subset=key_cols, keep="first")

    # Also collapse one-to-one alias pairs (sheet_name <-> final category) when the
    # rendered series are truly identical. This prevents duplicate charts such as
    # "Chemicals" and "Chemical (incl. petrochemical)" while preserving real
    # parent/child differences (which are not value-identical).
    pair_rows = sm[["sheet_name", "grouped_category_name"]].drop_duplicates()
    labels_present = set(out["sheet"].astype(str))
    pair_rows = pair_rows[
        pair_rows["sheet_name"].isin(labels_present)
        & pair_rows["grouped_category_name"].isin(labels_present)
        & pair_rows["sheet_name"].ne(pair_rows["grouped_category_name"])
    ].copy()
    if not pair_rows.empty:
        def _has_explicit_measure_rows(sheet_name: str) -> bool:
            sheet_rows = out[out["sheet"].eq(sheet_name)]
            if sheet_rows.empty or "measure" not in sheet_rows.columns:
                return False
            measures = [
                str(value).strip().lower()
                for value in sheet_rows["measure"].dropna().astype(str).tolist()
                if str(value).strip()
            ]
            return any(
                any(token in measure for token in ("input", "output", "feedstock", "product"))
                for measure in measures
            )

        pair_rows = pair_rows[
            ~pair_rows["sheet_name"].map(_has_explicit_measure_rows)
            & ~pair_rows["grouped_category_name"].map(_has_explicit_measure_rows)
        ].copy()
    if pair_rows.empty:
        return out

    def _series_equivalent(left_sheet: str, right_sheet: str) -> bool:
        left_raw = out[out["sheet"].eq(left_sheet)].copy()
        right_raw = out[out["sheet"].eq(right_sheet)].copy()
        compare_left = left_raw.copy()
        compare_right = right_raw.copy()

        compare_key_cols = ["economy", "fuel_label", "scenario", "source", "year"]
        if "measure" in out.columns:
            left_measures = {
                str(value).strip()
                for value in compare_left["measure"].dropna().astype(str)
                if str(value).strip() and str(value).strip().lower() != "nan"
            }
            right_measures = {
                str(value).strip()
                for value in compare_right["measure"].dropna().astype(str)
                if str(value).strip() and str(value).strip().lower() != "nan"
            }
            combined_measures = left_measures | right_measures
            if len(combined_measures) == 1:
                normalized_measure = next(iter(combined_measures))
                compare_left["measure"] = compare_left["measure"].fillna("").astype(str).str.strip()
                compare_right["measure"] = compare_right["measure"].fillna("").astype(str).str.strip()
                compare_left.loc[compare_left["measure"].eq(""), "measure"] = normalized_measure
                compare_right.loc[compare_right["measure"].eq(""), "measure"] = normalized_measure
            compare_key_cols.append("measure")

        def _aggregate_series(frame: pd.DataFrame, value_name: str) -> pd.DataFrame:
            records: list[dict[str, object]] = []
            for values, g in frame.groupby(compare_key_cols, dropna=False):
                key_map = dict(zip(compare_key_cols, values if isinstance(values, tuple) else (values,), strict=False))
                vals = pd.to_numeric(g["value"], errors="coerce").dropna()
                if vals.empty:
                    out_val = float("nan")
                else:
                    uniq = vals.unique()
                    out_val = float(uniq[0]) if len(uniq) == 1 else float(vals.sum())
                key_map[value_name] = out_val
                records.append(key_map)
            return pd.DataFrame(records)

        left = _aggregate_series(compare_left, "left_value")
        right = _aggregate_series(compare_right, "right_value")
        merged = left.merge(right, on=compare_key_cols, how="outer")
        if merged.empty:
            return False
        diff = (merged["left_value"] - merged["right_value"]).abs()
        same = (
            (merged["left_value"].isna() & merged["right_value"].isna())
            | (merged["left_value"].notna() & merged["right_value"].notna() & diff.le(1e-9))
        )
        return bool(same.all())

    alias_map: dict[str, str] = {}
    for _, row in pair_rows.iterrows():
        sheet_name = str(row["sheet_name"])
        category_name = str(row["grouped_category_name"])
        if sheet_name in alias_map:
            continue
        if _series_equivalent(sheet_name, category_name):
            alias_map[sheet_name] = category_name

    if not alias_map:
        return out

    def _resolve_sheet(label: str) -> str:
        current = str(label)
        seen: set[str] = set()
        while current in alias_map and current not in seen:
            seen.add(current)
            current = alias_map[current]
        return current

    out = out.copy()
    out["sheet"] = out["sheet"].map(_resolve_sheet)
    out = out.drop_duplicates(subset=key_cols, keep="first")
    return out


def _fail_fast_leaf_holes(comparison_long: pd.DataFrame, mapping_status: pd.DataFrame) -> None:
    """Fail when mapped rows have missing comparator values across required sources."""
    if comparison_long.empty or mapping_status.empty:
        return

    ms = mapping_status.copy()
    ms["mapped"] = ms.get("mapped", False)
    if "measure" not in ms.columns:
        ms["measure"] = ""
    mapped_keys = ms[ms["mapped"].fillna(False).astype(bool)][["sheet", "measure", "fuel_label"]].drop_duplicates()
    if mapped_keys.empty:
        return

    if "measure" not in comparison_long.columns:
        comparison_long = comparison_long.assign(measure="")
    comp = comparison_long.merge(mapped_keys.assign(_mapped=True), on=["sheet", "measure", "fuel_label"], how="left")
    comp = comp[comp["_mapped"].fillna(False)]
    comp["value"] = pd.to_numeric(comp["value"], errors="coerce")
    comp = comp[comp["source"].isin(["base", "projection", "base_estimated", "projection_estimated", "base_mixed", "projection_mixed"])]
    if comp.empty:
        return

    holes = comp[comp["value"].isna()][["sheet", "measure", "fuel_label", "scenario", "source", "year"]].drop_duplicates()
    if holes.empty:
        return

    sample = holes.head(20).to_dict("records")
    raise RuntimeError(
        "V2 fail-fast: residual missing comparator values after normalization. "
        f"Missing rows: {len(holes)}. Examples: {sample}"
    )


def _split_sector_codes(raw_value: object) -> list[str]:
    text = str(raw_value or "").strip()
    if not text or text.lower() == "nan":
        return []
    parts = pd.Series([text]).str.split(r"\s*(?:,|;|\||\band\b)\s*", regex=True).iloc[0]
    out: list[str] = []
    seen: set[str] = set()
    for part in parts:
        token = str(part or "").strip()
        if not token:
            continue
        key = token.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(token)
    return out


def aggregate_leap_for_shared_sector_groups(
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
) -> pd.DataFrame:
    """
    For sheets that map to the same single 9th sector, aggregate LEAP values across
    those sheets and apply the aggregate back to each sheet's LEAP rows.
    """
    if comparison_long.empty or mapping_status.empty:
        return comparison_long

    ms = mapping_status.copy()
    ms["sheet"] = ms.get("sheet", "").fillna("").astype(str).str.strip()
    ms["sector_code_9th"] = ms.get("sector_code_9th", "").fillna("").astype(str).str.strip()
    ms = ms[ms["sheet"].ne("")].copy()
    if ms.empty:
        return comparison_long

    # Keep only sheets with exactly one mapped sector code.
    sheet_sector_rows: list[dict[str, str]] = []
    for sheet, g in ms.groupby("sheet", dropna=False):
        codes: set[str] = set()
        for raw in g["sector_code_9th"].tolist():
            codes.update(_split_sector_codes(raw))
        if len(codes) == 1:
            sheet_sector_rows.append({"sheet": str(sheet), "sector_code_9th": sorted(codes)[0]})
    if not sheet_sector_rows:
        return comparison_long

    sheet_sector = pd.DataFrame(sheet_sector_rows)
    shared = (
        sheet_sector.groupby("sector_code_9th", as_index=False)["sheet"]
        .nunique()
        .rename(columns={"sheet": "sheet_count"})
    )
    shared = shared[shared["sheet_count"] > 1].copy()
    if shared.empty:
        return comparison_long
    shared_sector_codes = set(shared["sector_code_9th"].astype(str))
    sheet_sector = sheet_sector[sheet_sector["sector_code_9th"].isin(shared_sector_codes)].copy()

    comp = comparison_long.copy()
    comp["sheet"] = comp["sheet"].astype(str)
    if "measure" not in comp.columns:
        comp["measure"] = ""
    comp["measure"] = comp["measure"].fillna("").astype(str)
    comp["source"] = comp["source"].astype(str)
    comp["value"] = pd.to_numeric(comp["value"], errors="coerce")

    leap = comp[comp["source"].eq("leap")].copy()
    if leap.empty:
        return comp
    leap = leap.merge(sheet_sector, on="sheet", how="inner")
    if leap.empty:
        return comp

    grouped = (
        leap.groupby(["sector_code_9th", "measure", "fuel_label", "scenario", "year"], as_index=False)["value"]
        .sum(min_count=1)
        .rename(columns={"value": "group_leap_value"})
    )
    leap = leap.merge(grouped, on=["sector_code_9th", "measure", "fuel_label", "scenario", "year"], how="left")

    # Replace original leap rows for shared-sector sheets with grouped totals.
    replace_keys = leap[["sheet", "measure", "fuel_label", "scenario", "year", "source"]].drop_duplicates()
    comp = comp.merge(
        replace_keys.assign(_replace=True),
        on=["sheet", "measure", "fuel_label", "scenario", "year", "source"],
        how="left",
    )
    comp = comp[comp["_replace"] != True].drop(columns=["_replace"], errors="ignore")

    leap["value"] = pd.to_numeric(leap["group_leap_value"], errors="coerce")
    keep_cols = [c for c in comparison_long.columns if c in leap.columns]
    comp = pd.concat([comp, leap[keep_cols]], ignore_index=True, sort=False)
    return comp


def add_parent_totals_from_dashboard_overrides(
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    sheet_map: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Add derived parent sheets for dashboard parent overrides when the parent
    sheet does not already exist.

    Parent rows are sums of the child sheet values for each measure/fuel/year/source.
    Mapping status rows are derived from child mapping rows with a derived note.
    """
    if comparison_long.empty or sheet_map.empty:
        return comparison_long, mapping_status

    sm = sheet_map.copy()
    for col in ["sheet_name", "dashboard_page", "measure", "sector_code_9th"]:
        if col not in sm.columns:
            sm[col] = ""
        sm[col] = sm[col].fillna("").astype(str).str.strip()
    child_rows = sm[sm["dashboard_page"].ne("") & sm["sheet_name"].ne("")].copy()
    if child_rows.empty:
        return comparison_long, mapping_status

    def _numeric_prefix_tokens(code: str) -> list[str]:
        parts = [p for p in str(code).strip().split("_") if p]
        numeric: list[str] = []
        for token in parts:
            if token.isdigit():
                numeric.append(token)
            else:
                break
        return numeric

    def _is_sector_descendant(parent_code: str, child_code: str) -> bool:
        parent_parts = _numeric_prefix_tokens(parent_code)
        child_parts = _numeric_prefix_tokens(child_code)
        if not parent_parts or not child_parts:
            return False
        if len(child_parts) <= len(parent_parts):
            return False
        return child_parts[: len(parent_parts)] == parent_parts

    parent_to_children: dict[str, list[tuple[str, str]]] = {}
    for _, row in child_rows.iterrows():
        parent = str(row["dashboard_page"]).strip()
        child = str(row["sheet_name"]).strip()
        child_code = str(row.get("sector_code_9th", "")).strip()
        if not parent or not child:
            continue
        parent_to_children.setdefault(parent, [])
        if not any(existing_child == child for existing_child, _ in parent_to_children[parent]):
            parent_to_children[parent].append((child, child_code))
    if not parent_to_children:
        return comparison_long, mapping_status

    # Avoid parent-total double counting when a parent override includes both an
    # intermediate node (e.g. Manufacturing) and its descendants (e.g. Chemicals).
    # Keep leaf-most contributors and drop ancestor rows when descendants exist.
    resolved_children_by_parent: dict[str, list[str]] = {}
    for parent, child_entries in parent_to_children.items():
        keep: list[str] = []
        for child, child_code in child_entries:
            if not child:
                continue
            is_ancestor_of_other = False
            if child_code:
                for other_child, other_code in child_entries:
                    if other_child == child or not other_code:
                        continue
                    if _is_sector_descendant(child_code, other_code):
                        is_ancestor_of_other = True
                        break
            if not is_ancestor_of_other:
                keep.append(child)
        deduped_keep = []
        seen: set[str] = set()
        for child in keep:
            if child in seen:
                continue
            seen.add(child)
            deduped_keep.append(child)
        resolved_children_by_parent[parent] = deduped_keep

    comp = comparison_long.copy()
    comp["sheet"] = comp["sheet"].astype(str)
    if "measure" not in comp.columns:
        comp["measure"] = ""
    comp["measure"] = comp["measure"].fillna("").astype(str)
    comp["value"] = pd.to_numeric(comp["value"], errors="coerce")

    new_rows: list[pd.DataFrame] = []
    for parent, children in resolved_children_by_parent.items():
        if not children:
            continue
        if comp["sheet"].eq(parent).any():
            continue
        child_vals = comp[comp["sheet"].isin(children)].copy()
        if child_vals.empty:
            continue
        keys = ["economy", "scenario", "source", "year", "fuel_label", "measure"]
        grouped = (
            child_vals.groupby(keys, as_index=False)["value"]
            .sum(min_count=1)
            .assign(sheet=parent)
        )
        for col in ["effective_comparator_sector_code", "sector_codes_list"]:
            if col not in child_vals.columns:
                continue
            unique = [
                str(value).strip()
                for value in child_vals[col].dropna().astype(str).tolist()
                if str(value).strip() and str(value).strip().lower() != "nan"
            ]
            unique_vals = sorted(set(unique))
            fill_value = unique_vals[0] if len(unique_vals) == 1 else ""
            grouped[col] = fill_value
        # Align output columns to comparison_long
        for col in comparison_long.columns:
            if col not in grouped.columns:
                grouped[col] = ""
        grouped = grouped[comparison_long.columns]
        new_rows.append(grouped)

    if not new_rows:
        return comparison_long, mapping_status

    comp = pd.concat([comp, *new_rows], ignore_index=True, sort=False)

    if mapping_status is None or mapping_status.empty:
        return comp, mapping_status

    ms = mapping_status.copy()
    if "measure" not in ms.columns:
        ms["measure"] = ""
    ms["measure"] = ms["measure"].fillna("").astype(str)
    ms["sheet"] = ms["sheet"].fillna("").astype(str)
    ms["fuel_label"] = ms["fuel_label"].fillna("").astype(str)

    bool_any_cols = {
        "has_any_mapping",
        "mapped",
        "partially_mapped",
        "base_mapping_optional",
        "projection_parent_fallback",
        "uses_parent_flow",
        "allow_parent_estimate",
        "aggregated_mapping",
    }
    bool_all_cols = {"base_mapping_complete", "projection_mapping_complete"}
    numeric_cols = {"esto_flow_depth", "min_sector_depth", "derived_parent_flow_levels"}

    new_status_rows: list[dict[str, object]] = []
    for parent, children in resolved_children_by_parent.items():
        if ms["sheet"].eq(parent).any():
            continue
        child_ms = ms[ms["sheet"].isin(children)].copy()
        if child_ms.empty:
            continue
        for (fuel_label, measure), g in child_ms.groupby(["fuel_label", "measure"], dropna=False):
            row: dict[str, object] = {}
            for col in ms.columns:
                if col == "sheet":
                    row[col] = parent
                elif col == "fuel_label":
                    row[col] = fuel_label
                elif col == "measure":
                    row[col] = measure
                elif col in bool_any_cols:
                    row[col] = bool(g[col].fillna(False).astype(bool).any()) if col in g.columns else False
                elif col in bool_all_cols:
                    row[col] = bool(g[col].fillna(False).astype(bool).all()) if col in g.columns else False
                elif col in numeric_cols:
                    vals = pd.to_numeric(g[col], errors="coerce") if col in g.columns else pd.Series([], dtype="float64")
                    row[col] = float(vals.min()) if not vals.dropna().empty else ""
                else:
                    if col in g.columns:
                        values = [
                            str(value).strip()
                            for value in g[col].dropna().astype(str).tolist()
                            if str(value).strip() and str(value).strip().lower() != "nan"
                        ]
                        unique_vals = sorted(set(values))
                        row[col] = unique_vals[0] if len(unique_vals) == 1 else ""
                    else:
                        row[col] = ""
            row["mapping_source"] = "derived_parent_total"
            row["mapping_note"] = f"Derived from children: {', '.join(children)}"
            row["mapping_precedence"] = "derived_parent_total"
            new_status_rows.append(row)

    if new_status_rows:
        ms = pd.concat([ms, pd.DataFrame(new_status_rows)], ignore_index=True, sort=False)
    return comp, ms


def _rebuild_comparison_wide(comparison_long: pd.DataFrame) -> pd.DataFrame:
    if comparison_long.empty:
        return comparison_long.copy()
    wide = (
        comparison_long.pivot_table(
            index=["economy", "scenario", "sheet", "fuel_label", "year"],
            columns="source",
            values="value",
            aggfunc="first",
        )
        .reset_index()
    )
    if hasattr(wide.columns, "name"):
        wide.columns.name = None
    return wide


def build_comparisons_v2(
    *,
    leap_long: pd.DataFrame,
    sheet_map: pd.DataFrame,
    fuel_mapping: dict[str, dict[str, str]],
    sector_flow_mapping: dict[str, str],
    ninth_pairs: pd.DataFrame,
    base_df: pd.DataFrame,
    ninth_df: pd.DataFrame,
    explicit_mappings: pd.DataFrame,
    base_year: int,
    base_economy: str,
    projection_economy: str,
    projection_years: tuple[int, ...],
    scenario_map: dict[str, str],
    use_esto_agg_only: bool,
    sibling_comparator_mode: str,
    include_sibling_parent_totals: bool,
    settings: DashboardV2Settings,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    comparison_long, comparison_wide, mapping_status = build_comparisons(
        leap_long,
        sheet_map=sheet_map,
        fuel_mapping=fuel_mapping,
        sector_flow_mapping=sector_flow_mapping,
        ninth_pairs=ninth_pairs,
        base_df=base_df,
        ninth_df=ninth_df,
        explicit_mappings=explicit_mappings,
        base_year=base_year,
        base_economy=base_economy,
        projection_economy=projection_economy,
        projection_years=projection_years,
        scenario_map=scenario_map,
        use_esto_agg_only=use_esto_agg_only,
        sibling_comparator_mode=sibling_comparator_mode,
        include_sibling_parent_totals=include_sibling_parent_totals,
    )

    mapping_status = annotate_mapping_status(mapping_status)

    # Keep full comparison outputs for diagnostics and downstream analysis.
    # `common_level_only` filtering is applied later for chart/dashboard inputs only.

    if settings.leaf_hole_policy == "fail_fast":
        check_frame = comparison_long
        if settings.mapping_graph_mode == "common_level_only":
            check_frame = _common_level_only_filter(comparison_long, mapping_status)
        _fail_fast_leaf_holes(check_frame, mapping_status)

    return comparison_long, comparison_wide, mapping_status
