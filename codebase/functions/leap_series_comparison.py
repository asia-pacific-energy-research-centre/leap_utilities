from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Any

import pandas as pd

from codebase.utilities.master_config import config_table_exists, read_config_table

from codebase.functions.leap_excel_io import read_export_sheet
from codebase.functions.leap_expressions import expression_to_series
from codebase.functions.leap_labels import clean_fuel_label_for_leap
from codebase.functions.ninth_projection_mapping import (
    build_esto_projection_table,
    normalize_economy_key,
)
from codebase.scrapbook.utilities import (
    apply_matt_subtotal_mapping,
    filter_matt_subtotals,
    load_augmented_reference_tables,
)

FUEL_GROUP_LABELS = ("Output Fuels", "Feedstock Fuels", "Auxiliary Fuels")
REQUIRED_MAPPING_COLUMNS = [
    "series_id",
    "sector_tag",
    "leap_variable",
    "leap_branch_contains",
    "leap_fuel_label",
    "esto_flow",
    "esto_product",
    "ninth_sector_expected",
    "ninth_fuel_expected",
    "active",
    "notes",
]
TRANSPORT_BRANCH_MAPPING_COLUMNS = [
    "branch_path",
    "ninth_sector_code",
    "active",
    "include_in_demand_total",
    "notes",
]
TRANSPORT_FUEL_ALIAS_COLUMNS = [
    "leap_fuel_label",
    "codebook_name",
    "ninth_fuel_override",
    "esto_product_override",
    "active",
    "notes",
]
NINTH_SECTOR_COLUMNS_BY_DEPTH = {
    1: "sectors",
    2: "sub1sectors",
    3: "sub2sectors",
    4: "sub3sectors",
    5: "sub4sectors",
}
NINTH_FUEL_COLUMNS = {"fuels", "subfuels"}


@dataclass
class ComparisonRunConfig:
    leap_file: str | Path
    leap_sheet: str
    mapping_csv: str | Path
    economy: str
    scenario: str
    region: str
    esto_data_path: str | Path
    ninth_data_path: str | Path
    subtotal_mapping_path: str | Path
    ninth_to_esto_mapping_path: str | Path
    base_year: int = 2022
    projection_start_year: int = 2023
    projection_end_year: int = 2060
    output_dir: str | Path = Path("outputs") / "series_comparison"


@dataclass
class ComparisonArtifacts:
    comparison_long_csv: Path
    comparison_wide_csv: Path
    comparison_summary_csv: Path
    mapping_status_csv: Path
    unmatched_leap_rows_csv: Path
    charts_dir: Path


@dataclass
class TransportResultsComparisonConfig:
    leap_results_file: str | Path
    economy: str
    scenario: str
    region: str
    branch_sector_mapping_csv: str | Path = Path()
    fuel_aliases_csv: str | Path = Path()
    code_to_name_path: str | Path = Path("config") / "sector_fuel_codes_to_names.xlsx"
    code_to_name_sheet: str = "code_to_name"
    esto_data_path: str | Path = Path("data") / "00APEC_2024_low.csv"
    ninth_data_path: str | Path = (
        Path("data") / "merged_file_energy_ALL_20250814_pre_trump.csv"
    )
    subtotal_mapping_path: str | Path = Path("config") / "ESTO_subtotal_mapping.xlsx"
    ninth_to_esto_mapping_path: str | Path = (
        Path("config") / "ninth_pairs_to_esto_pairs.xlsx"
    )
    base_year: int = 2022
    projection_start_year: int = 2023
    projection_end_year: int = 2060
    share_year_offset: int = 1
    ninth_scenario: str = "reference"
    output_dir: str | Path = (
        Path("outputs") / "transport_results_series_comparison"
    )


def _coerce_year_from_header(value: object) -> int | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, int):
        return value if 1000 <= value <= 3000 else None
    if isinstance(value, float) and value.is_integer():
        year = int(value)
        return year if 1000 <= year <= 3000 else None
    text = str(value).strip()
    if re.fullmatch(r"\d{4}", text):
        return int(text)
    if re.fullmatch(r"\d{4}\.0+", text):
        return int(float(text))
    return None


def _normalize_year_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, list[int]]:
    rename_map: dict[Any, Any] = {}
    year_cols: list[int] = []
    for col in df.columns:
        year_int = _coerce_year_from_header(col)
        if year_int is not None:
            rename_map[col] = year_int
            year_cols.append(year_int)
    working = df.rename(columns=rename_map)
    return working, sorted(set(year_cols))


def _normalize_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _normalize_key_text(value: object) -> str:
    return _normalize_text(value).lower()


def _parse_active_flag(value: object) -> bool:
    text = _normalize_key_text(value)
    if text in {"", "0", "false", "no", "n"}:
        return False
    if text in {"1", "true", "yes", "y"}:
        return True
    return bool(text)


def _safe_filename_token(value: object) -> str:
    text = _normalize_text(value)
    if not text:
        return "series"
    safe = "".join(ch if ch.isalnum() or ch in {"_", "-"} else "_" for ch in text)
    return safe.strip("_") or "series"


def _extract_fuel_label_from_branch_path(branch_path: str) -> str:
    parts = [part.strip() for part in _normalize_text(branch_path).split("\\") if part.strip()]
    if not parts:
        return ""
    for idx, token in enumerate(parts[:-1]):
        if token in FUEL_GROUP_LABELS:
            candidate_idx = idx + 1
            if candidate_idx < len(parts):
                return clean_fuel_label_for_leap(parts[candidate_idx]).strip()
    return clean_fuel_label_for_leap(parts[-1]).strip()


def _read_leap_table(path: Path, sheet_name: str) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return read_config_table(path)

    try:
        _, data, _ = read_export_sheet(path, sheet_name=sheet_name)
        return data
    except Exception:
        pass

    try:
        return read_config_table(path, sheet_name=sheet_name, header=2)
    except Exception:
        return read_config_table(path, sheet_name=sheet_name)


def _load_and_normalize_leap_rows(
    leap_file: Path,
    leap_sheet: str,
    scenario: str,
    region: str,
) -> tuple[pd.DataFrame, list[int]]:
    working = _read_leap_table(leap_file, leap_sheet).copy()
    if working.empty:
        return pd.DataFrame(), []

    working.columns = [str(col).strip() if not isinstance(col, int) else col for col in working.columns]
    working, explicit_year_cols = _normalize_year_columns(working)

    required_cols = {"Branch Path", "Variable"}
    missing = required_cols - set(working.columns)
    if missing:
        raise ValueError(f"LEAP input is missing required columns: {sorted(missing)}")

    if "Scenario" in working.columns:
        scenario_norm = working["Scenario"].astype(str).str.strip().str.lower()
        working = working[scenario_norm == _normalize_key_text(scenario)].copy()
    if "Region" in working.columns:
        region_norm = working["Region"].astype(str).str.strip().str.lower()
        working = working[region_norm == _normalize_key_text(region)].copy()

    working["__row_id"] = range(len(working))
    working["__fuel_label_raw"] = working["Branch Path"].apply(_extract_fuel_label_from_branch_path)
    working["__fuel_label_norm"] = working["__fuel_label_raw"].apply(
        lambda value: _normalize_key_text(clean_fuel_label_for_leap(value))
    )
    working["__branch_path_norm"] = working["Branch Path"].astype(str).str.strip().str.lower()
    working["__variable_norm"] = working["Variable"].astype(str).str.strip().str.lower()

    long_rows: list[dict[str, object]] = []
    if explicit_year_cols:
        for _, row in working.iterrows():
            for year in explicit_year_cols:
                value = pd.to_numeric(row.get(year), errors="coerce")
                if pd.isna(value):
                    continue
                long_rows.append(
                    {
                        "__row_id": int(row["__row_id"]),
                        "Branch Path": row["Branch Path"],
                        "Variable": row["Variable"],
                        "Scenario": row.get("Scenario"),
                        "Region": row.get("Region"),
                        "__fuel_label_raw": row["__fuel_label_raw"],
                        "__fuel_label_norm": row["__fuel_label_norm"],
                        "__branch_path_norm": row["__branch_path_norm"],
                        "__variable_norm": row["__variable_norm"],
                        "year": int(year),
                        "value": float(value),
                    }
                )
    elif "Expression" in working.columns:
        for _, row in working.iterrows():
            series = expression_to_series(row.get("Expression"), years=None, base_year=None)
            if series is None:
                continue
            for year, value in sorted(series.items()):
                if value is None or pd.isna(value):
                    continue
                long_rows.append(
                    {
                        "__row_id": int(row["__row_id"]),
                        "Branch Path": row["Branch Path"],
                        "Variable": row["Variable"],
                        "Scenario": row.get("Scenario"),
                        "Region": row.get("Region"),
                        "__fuel_label_raw": row["__fuel_label_raw"],
                        "__fuel_label_norm": row["__fuel_label_norm"],
                        "__branch_path_norm": row["__branch_path_norm"],
                        "__variable_norm": row["__variable_norm"],
                        "year": int(year),
                        "value": float(value),
                    }
                )
    else:
        raise ValueError(
            "LEAP input does not include explicit year columns or an Expression column."
        )

    if not long_rows:
        return pd.DataFrame(), []
    long_df = pd.DataFrame(long_rows)
    long_df = long_df.sort_values(["__row_id", "year"]).reset_index(drop=True)
    years = sorted(long_df["year"].dropna().astype(int).unique().tolist())
    return long_df, years


def _load_mapping(mapping_csv: Path) -> pd.DataFrame:
    mapping = read_config_table(mapping_csv).fillna("")
    missing = [col for col in REQUIRED_MAPPING_COLUMNS if col not in mapping.columns]
    if missing:
        raise ValueError(
            f"Mapping CSV '{mapping_csv}' is missing required columns: {missing}"
        )
    mapping = mapping.copy()
    mapping["active"] = mapping["active"].apply(_parse_active_flag)
    mapping = mapping[mapping["active"]].copy()
    mapping = mapping.reset_index(drop=True)
    mapping["__mapping_index"] = mapping.index
    return mapping


def _load_esto_data(
    esto_data_path: Path,
    subtotal_mapping_path: Path,
) -> pd.DataFrame:
    df, _ = load_augmented_reference_tables(
        esto_path=esto_data_path,
        ninth_path=Path("data/merged_file_energy_ALL_20251106.csv"),
        subtotal_mapping_path=subtotal_mapping_path,
        synthetic_rules_path=Path("config/synthetic_reference_rows.csv"),
        cache_dir=Path("data/.cache/leap_series_comparison_reference_tables"),
        apply_esto_subtotal_map=True,
        filter_esto_subtotals_flag=True,
        filter_ninth_subtotals_flag=False,
    )
    df, _ = _normalize_year_columns(df)
    df["flows"] = df["flows"].astype(str).str.strip()
    df["products"] = df["products"].astype(str).str.strip()
    df["economy_key"] = df["economy"].apply(normalize_economy_key)
    return df


def _load_ninth_data(ninth_data_path: Path) -> pd.DataFrame:
    _, df = load_augmented_reference_tables(
        esto_path=Path("data/00APEC_2025_low_with_subtotals.csv"),
        ninth_path=ninth_data_path,
        synthetic_rules_path=Path("config/synthetic_reference_rows.csv"),
        cache_dir=Path("data/.cache/leap_series_comparison_reference_tables"),
        apply_esto_subtotal_map=False,
        filter_esto_subtotals_flag=False,
        filter_ninth_subtotals_flag=False,
    )
    df, _ = _normalize_year_columns(df)
    return df


def _load_mapping_pairs(mapping_path: Path) -> pd.DataFrame:
    if mapping_path.suffix.lower() in {".xlsx", ".xls"}:
        df = read_config_table(mapping_path, dtype=str).fillna("")
    else:
        df = read_config_table(mapping_path, dtype=str).fillna("")
    for col in ["9th_sector", "9th_fuel", "esto_flow", "esto_product"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    return df


def _build_reference_long(
    economy_key: str,
    esto_df: pd.DataFrame,
    ninth_df: pd.DataFrame,
    ninth_to_esto_mapping_path: Path,
    base_year: int,
    projection_years: list[int],
) -> pd.DataFrame:
    year_col = base_year if base_year in esto_df.columns else str(base_year)
    if year_col not in esto_df.columns:
        raise ValueError(f"ESTO data does not include base year column '{base_year}'.")

    esto_base = esto_df[esto_df["economy_key"] == economy_key].copy()
    if esto_base.empty:
        raise ValueError(f"No ESTO rows found for economy key '{economy_key}'.")
    esto_base[year_col] = pd.to_numeric(esto_base[year_col], errors="coerce")

    base_ref = (
        esto_base.groupby(["economy_key", "flows", "products"], dropna=False)[year_col]
        .sum()
        .reset_index()
        .rename(
            columns={
                "flows": "esto_flow",
                "products": "esto_product",
                year_col: "reference_value",
            }
        )
    )
    base_ref["year"] = int(base_year)
    base_ref["reference_source"] = "esto_base_year"

    proj_df, _diagnostics = build_esto_projection_table(
        ninth_data=ninth_df,
        esto_data=esto_df,
        mapping_path=ninth_to_esto_mapping_path,
        base_year=base_year,
        projection_years=projection_years,
        scenario="reference",
    )
    if proj_df is None or proj_df.empty:
        projection_ref = pd.DataFrame(
            columns=[
                "economy_key",
                "esto_flow",
                "esto_product",
                "reference_value",
                "year",
                "reference_source",
            ]
        )
    else:
        proj_working = proj_df[proj_df["economy_key"] == economy_key].copy()
        projection_rows: list[dict[str, object]] = []
        for _, row in proj_working.iterrows():
            for year in projection_years:
                if year not in proj_working.columns:
                    continue
                value = pd.to_numeric(row.get(year), errors="coerce")
                if pd.isna(value):
                    continue
                projection_rows.append(
                    {
                        "economy_key": row["economy_key"],
                        "esto_flow": row["esto_flow"],
                        "esto_product": row["esto_product"],
                        "reference_value": float(value),
                        "year": int(year),
                        "reference_source": "ninth_projection_allocated",
                    }
                )
        projection_ref = pd.DataFrame(projection_rows)

    reference_long = pd.concat(
        [
            base_ref[
                ["economy_key", "esto_flow", "esto_product", "reference_value", "year", "reference_source"]
            ],
            projection_ref[
                ["economy_key", "esto_flow", "esto_product", "reference_value", "year", "reference_source"]
            ]
            if not projection_ref.empty
            else projection_ref,
        ],
        ignore_index=True,
    )
    return reference_long


def _match_leap_rows_for_mapping(
    leap_long_df: pd.DataFrame,
    mapping_row: pd.Series,
) -> pd.DataFrame:
    matched = leap_long_df.copy()

    variable_filter = _normalize_key_text(mapping_row.get("leap_variable"))
    if variable_filter:
        matched = matched[matched["__variable_norm"] == variable_filter]

    branch_contains = _normalize_key_text(mapping_row.get("leap_branch_contains"))
    if branch_contains:
        matched = matched[
            matched["__branch_path_norm"].str.contains(
                branch_contains,
                na=False,
                regex=False,
            )
        ]

    fuel_filter = _normalize_key_text(mapping_row.get("leap_fuel_label"))
    if fuel_filter:
        matched = matched[matched["__fuel_label_norm"] == fuel_filter]

    return matched


def _validate_expected_ninth_pair(
    mapping_pairs: pd.DataFrame,
    mapping_row: pd.Series,
) -> tuple[bool, str]:
    ninth_sector_expected = _normalize_text(mapping_row.get("ninth_sector_expected"))
    ninth_fuel_expected = _normalize_text(mapping_row.get("ninth_fuel_expected"))
    esto_flow = _normalize_text(mapping_row.get("esto_flow"))
    esto_product = _normalize_text(mapping_row.get("esto_product"))

    if not ninth_sector_expected and not ninth_fuel_expected:
        return False, ""
    if mapping_pairs.empty:
        return False, "missing_mapping_pairs_table"

    subset = mapping_pairs[
        (mapping_pairs["esto_flow"] == esto_flow)
        & (mapping_pairs["esto_product"] == esto_product)
    ].copy()
    if subset.empty:
        return False, "esto_pair_not_in_mapping_table"

    pair_subset = subset[
        (subset["9th_sector"] == ninth_sector_expected)
        & (subset["9th_fuel"] == ninth_fuel_expected)
    ]
    if pair_subset.empty:
        return False, "expected_9th_pair_not_found_for_esto_pair"
    return True, ""


def _build_comparison_outputs(
    config: ComparisonRunConfig,
    leap_long_df: pd.DataFrame,
    mapping_df: pd.DataFrame,
    reference_long_df: pd.DataFrame,
    mapping_pairs_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    projection_years = list(range(config.projection_start_year, config.projection_end_year + 1))
    target_years = [config.base_year] + projection_years

    comparison_rows: list[dict[str, object]] = []
    mapping_status_rows: list[dict[str, object]] = []
    matched_row_ids: set[int] = set()

    for _, mapping_row in mapping_df.iterrows():
        matched = _match_leap_rows_for_mapping(leap_long_df, mapping_row)
        if not matched.empty:
            matched_row_ids.update(matched["__row_id"].astype(int).tolist())

        leap_series = (
            matched.groupby("year", dropna=False)["value"].sum().to_dict()
            if not matched.empty
            else {}
        )
        esto_flow = _normalize_text(mapping_row.get("esto_flow"))
        esto_product = _normalize_text(mapping_row.get("esto_product"))
        ref_subset = reference_long_df[
            (reference_long_df["esto_flow"] == esto_flow)
            & (reference_long_df["esto_product"] == esto_product)
        ].copy()
        ref_by_year = {
            int(row["year"]): (float(row["reference_value"]), _normalize_text(row["reference_source"]))
            for _, row in ref_subset.iterrows()
        }

        for year in target_years:
            leap_value = leap_series.get(year)
            ref_pair = ref_by_year.get(year)
            reference_value = ref_pair[0] if ref_pair else pd.NA
            reference_source = ref_pair[1] if ref_pair else "missing_reference"
            if leap_value is None:
                leap_value = pd.NA
            delta = pd.NA
            abs_delta = pd.NA
            pct_delta = pd.NA
            if pd.notna(leap_value) and pd.notna(reference_value):
                delta = float(leap_value) - float(reference_value)
                abs_delta = abs(delta)
                if abs(float(reference_value)) > 1e-12:
                    pct_delta = delta / float(reference_value)
            comparison_rows.append(
                {
                    "series_id": _normalize_text(mapping_row.get("series_id")),
                    "sector_tag": _normalize_text(mapping_row.get("sector_tag")),
                    "economy": config.economy,
                    "scenario": config.scenario,
                    "region": config.region,
                    "leap_variable": _normalize_text(mapping_row.get("leap_variable")),
                    "leap_branch_contains": _normalize_text(mapping_row.get("leap_branch_contains")),
                    "leap_fuel_label": _normalize_text(mapping_row.get("leap_fuel_label")),
                    "esto_flow": esto_flow,
                    "esto_product": esto_product,
                    "year": int(year),
                    "leap_value": leap_value,
                    "reference_value": reference_value,
                    "delta": delta,
                    "abs_delta": abs_delta,
                    "pct_delta": pct_delta,
                    "reference_source": reference_source,
                    "mapping_index": int(mapping_row["__mapping_index"]),
                }
            )

        expected_ok, expected_issue = _validate_expected_ninth_pair(mapping_pairs_df, mapping_row)
        missing_reference_years = [
            year for year in target_years if year not in ref_by_year
        ]
        mapping_status_rows.append(
            {
                "series_id": _normalize_text(mapping_row.get("series_id")),
                "sector_tag": _normalize_text(mapping_row.get("sector_tag")),
                "mapping_index": int(mapping_row["__mapping_index"]),
                "active": True,
                "leap_rows_matched": int(matched["__row_id"].nunique()) if not matched.empty else 0,
                "leap_points_matched": int(len(matched)),
                "has_leap_match": bool(not matched.empty),
                "has_base_year_reference": bool(config.base_year in ref_by_year),
                "has_projection_reference": bool(
                    any(year in ref_by_year for year in projection_years)
                ),
                "missing_reference_years_count": int(len(missing_reference_years)),
                "missing_reference_years": ";".join(str(year) for year in missing_reference_years),
                "expected_pair_provided": bool(
                    _normalize_text(mapping_row.get("ninth_sector_expected"))
                    or _normalize_text(mapping_row.get("ninth_fuel_expected"))
                ),
                "expected_pair_matches_mapping_table": expected_ok,
                "expected_pair_issue": expected_issue,
                "esto_flow": _normalize_text(mapping_row.get("esto_flow")),
                "esto_product": _normalize_text(mapping_row.get("esto_product")),
                "ninth_sector_expected": _normalize_text(mapping_row.get("ninth_sector_expected")),
                "ninth_fuel_expected": _normalize_text(mapping_row.get("ninth_fuel_expected")),
                "notes": _normalize_text(mapping_row.get("notes")),
            }
        )

    comparison_long = pd.DataFrame(comparison_rows)
    if not comparison_long.empty:
        comparison_long = comparison_long.sort_values(["series_id", "year"]).reset_index(drop=True)

    mapping_status = pd.DataFrame(mapping_status_rows)
    if not mapping_status.empty:
        mapping_status = mapping_status.sort_values(["series_id", "mapping_index"]).reset_index(drop=True)

    unmatched = leap_long_df[~leap_long_df["__row_id"].isin(matched_row_ids)].copy()
    if not unmatched.empty:
        unmatched = unmatched[
            [
                "__row_id",
                "Branch Path",
                "Variable",
                "Scenario",
                "Region",
                "__fuel_label_raw",
                "year",
                "value",
            ]
        ].rename(columns={"__fuel_label_raw": "fuel_label"})
        unmatched = unmatched.sort_values(["__row_id", "year"]).reset_index(drop=True)
    else:
        unmatched = pd.DataFrame(
            columns=[
                "__row_id",
                "Branch Path",
                "Variable",
                "Scenario",
                "Region",
                "fuel_label",
                "year",
                "value",
            ]
        )

    comparison_wide = _build_wide_comparison_table(comparison_long)
    comparison_summary = _build_summary_table(comparison_long, config.base_year)
    return comparison_long, comparison_wide, comparison_summary, mapping_status, unmatched


def _build_wide_comparison_table(comparison_long: pd.DataFrame) -> pd.DataFrame:
    if comparison_long.empty:
        return pd.DataFrame()

    key_cols = [
        "series_id",
        "sector_tag",
        "economy",
        "scenario",
        "region",
        "leap_variable",
        "leap_branch_contains",
        "leap_fuel_label",
        "esto_flow",
        "esto_product",
        "mapping_index",
    ]
    working = comparison_long.copy()
    working["year"] = working["year"].astype(int)

    pivot_parts: list[pd.DataFrame] = []
    for value_col, prefix in [
        ("leap_value", "leap"),
        ("reference_value", "reference"),
        ("delta", "delta"),
    ]:
        part = (
            working.pivot_table(
                index=key_cols,
                columns="year",
                values=value_col,
                aggfunc="first",
            )
            .reset_index()
        )
        part.columns = [
            col if not isinstance(col, int) else f"{prefix}_{col}"
            for col in part.columns
        ]
        pivot_parts.append(part)

    merged = pivot_parts[0]
    for part in pivot_parts[1:]:
        merged = merged.merge(part, on=key_cols, how="outer")
    return merged.sort_values(["series_id", "mapping_index"]).reset_index(drop=True)


def _build_summary_table(comparison_long: pd.DataFrame, base_year: int) -> pd.DataFrame:
    if comparison_long.empty:
        return pd.DataFrame()

    rows: list[dict[str, object]] = []
    for series_id, group in comparison_long.groupby("series_id", dropna=False):
        base_row = group[group["year"] == int(base_year)]
        base_year_delta = pd.NA
        if not base_row.empty:
            value = pd.to_numeric(base_row["delta"], errors="coerce").dropna()
            if not value.empty:
                base_year_delta = float(value.iloc[0])

        projection_group = group[group["year"] > int(base_year)].copy()
        projection_abs_delta_sum = float(
            pd.to_numeric(projection_group["abs_delta"], errors="coerce").fillna(0.0).sum()
        )
        max_abs_delta = float(
            pd.to_numeric(group["abs_delta"], errors="coerce").fillna(0.0).max()
        )
        pct_vals = pd.to_numeric(group["pct_delta"], errors="coerce").dropna()
        mean_abs_pct_delta = float(pct_vals.abs().mean()) if not pct_vals.empty else pd.NA

        leap_vals = pd.to_numeric(group["leap_value"], errors="coerce")
        ref_vals = pd.to_numeric(group["reference_value"], errors="coerce")
        valid_sign = leap_vals.notna() & ref_vals.notna() & leap_vals.ne(0.0) & ref_vals.ne(0.0)
        sign_mismatch_count = int(((leap_vals * ref_vals) < 0).where(valid_sign, False).sum())

        rows.append(
            {
                "series_id": _normalize_text(series_id),
                "sector_tag": _normalize_text(group["sector_tag"].iloc[0]),
                "economy": _normalize_text(group["economy"].iloc[0]),
                "scenario": _normalize_text(group["scenario"].iloc[0]),
                "region": _normalize_text(group["region"].iloc[0]),
                "esto_flow": _normalize_text(group["esto_flow"].iloc[0]),
                "esto_product": _normalize_text(group["esto_product"].iloc[0]),
                "base_year_delta": base_year_delta,
                "projection_abs_delta_sum": projection_abs_delta_sum,
                "max_abs_delta": max_abs_delta,
                "mean_abs_pct_delta": mean_abs_pct_delta,
                "sign_mismatch_count": sign_mismatch_count,
                "year_count": int(group["year"].nunique()),
            }
        )

    return pd.DataFrame(rows).sort_values("series_id").reset_index(drop=True)


def _write_comparison_charts(
    comparison_long: pd.DataFrame,
    charts_dir: Path,
) -> None:
    if comparison_long.empty:
        return

    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except Exception as exc:
        print(
            "[WARN] Skipping chart generation because matplotlib is unavailable: "
            f"{exc}"
        )
        return

    charts_dir.mkdir(parents=True, exist_ok=True)
    grouped = comparison_long.groupby("series_id", dropna=False)
    for series_id, group in grouped:
        working = group.sort_values("year").copy()
        years = working["year"].astype(int).tolist()
        leap_vals = pd.to_numeric(working["leap_value"], errors="coerce")
        ref_vals = pd.to_numeric(working["reference_value"], errors="coerce")
        delta_vals = pd.to_numeric(working["delta"], errors="coerce")

        fig, (ax_top, ax_bottom) = plt.subplots(
            nrows=2,
            ncols=1,
            figsize=(10, 6),
            sharex=True,
            gridspec_kw={"height_ratios": [3, 1]},
        )

        ax_top.plot(years, leap_vals, label="LEAP", marker="o")
        ax_top.plot(years, ref_vals, label="Reference", marker="o")
        ax_top.set_ylabel("Value")
        ax_top.grid(True, alpha=0.3)
        ax_top.legend(loc="best")
        title = (
            f"{_normalize_text(series_id)} | "
            f"{_normalize_text(working['esto_flow'].iloc[0])} | "
            f"{_normalize_text(working['esto_product'].iloc[0])}"
        )
        ax_top.set_title(title)

        ax_bottom.axhline(0.0, color="black", linewidth=0.8)
        ax_bottom.bar(years, delta_vals.fillna(0.0), color="#4C78A8")
        ax_bottom.set_xlabel("Year")
        ax_bottom.set_ylabel("Delta")
        ax_bottom.grid(True, axis="y", alpha=0.3)

        fig.tight_layout()
        out_name = f"{_safe_filename_token(series_id)}.png"
        fig.savefig(charts_dir / out_name, dpi=160)
        plt.close(fig)


def run_leap_series_comparison(config: ComparisonRunConfig) -> ComparisonArtifacts:
    leap_file = Path(config.leap_file)
    mapping_csv = Path(config.mapping_csv)
    esto_data_path = Path(config.esto_data_path)
    ninth_data_path = Path(config.ninth_data_path)
    subtotal_mapping_path = Path(config.subtotal_mapping_path)
    ninth_to_esto_mapping_path = Path(config.ninth_to_esto_mapping_path)
    output_dir = Path(config.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    charts_dir = output_dir / "charts"
    charts_dir.mkdir(parents=True, exist_ok=True)

    if config.projection_end_year < config.projection_start_year:
        raise ValueError("projection_end_year must be >= projection_start_year.")

    mapping_df = _load_mapping(mapping_csv)
    leap_long_df, leap_years = _load_and_normalize_leap_rows(
        leap_file=leap_file,
        leap_sheet=config.leap_sheet,
        scenario=config.scenario,
        region=config.region,
    )
    if leap_long_df.empty:
        raise ValueError("No LEAP rows remain after filtering by scenario/region and year parsing.")

    projection_years = [
        year for year in range(config.projection_start_year, config.projection_end_year + 1)
    ]
    if leap_years:
        max_leap_year = max(leap_years)
        projection_years = [year for year in projection_years if year <= max_leap_year]

    economy_key = normalize_economy_key(config.economy)
    esto_df = _load_esto_data(esto_data_path, subtotal_mapping_path)
    ninth_df = _load_ninth_data(ninth_data_path)
    mapping_pairs_df = _load_mapping_pairs(ninth_to_esto_mapping_path)
    reference_long_df = _build_reference_long(
        economy_key=economy_key,
        esto_df=esto_df,
        ninth_df=ninth_df,
        ninth_to_esto_mapping_path=ninth_to_esto_mapping_path,
        base_year=config.base_year,
        projection_years=projection_years,
    )

    comparison_long, comparison_wide, comparison_summary, mapping_status, unmatched = (
        _build_comparison_outputs(
            config=config,
            leap_long_df=leap_long_df,
            mapping_df=mapping_df,
            reference_long_df=reference_long_df,
            mapping_pairs_df=mapping_pairs_df,
        )
    )

    _write_comparison_charts(comparison_long, charts_dir)

    comparison_long_csv = output_dir / "comparison_long.csv"
    comparison_wide_csv = output_dir / "comparison_wide.csv"
    comparison_summary_csv = output_dir / "comparison_summary.csv"
    mapping_status_csv = output_dir / "mapping_status.csv"
    unmatched_leap_rows_csv = output_dir / "unmatched_leap_rows.csv"

    comparison_long.to_csv(comparison_long_csv, index=False)
    comparison_wide.to_csv(comparison_wide_csv, index=False)
    comparison_summary.to_csv(comparison_summary_csv, index=False)
    mapping_status.to_csv(mapping_status_csv, index=False)
    unmatched.to_csv(unmatched_leap_rows_csv, index=False)

    return ComparisonArtifacts(
        comparison_long_csv=comparison_long_csv,
        comparison_wide_csv=comparison_wide_csv,
        comparison_summary_csv=comparison_summary_csv,
        mapping_status_csv=mapping_status_csv,
        unmatched_leap_rows_csv=unmatched_leap_rows_csv,
        charts_dir=charts_dir,
    )


def _normalize_lookup_label(value: object) -> str:
    text = _normalize_text(value).lower()
    if not text:
        return ""
    text = text.replace("&", " and ")
    text = text.replace("/", " ")
    text = text.replace("-", " ")
    text = re.sub(r"[^a-z0-9\s]+", " ", text)
    return " ".join(text.split())


def _parse_scenario_region_from_results_metadata(value: object) -> tuple[str, str]:
    text = _normalize_text(value)
    if not text:
        return "", ""
    scenario = ""
    region = ""
    if "Scenario:" in text:
        payload = text.split("Scenario:", 1)[1].strip()
    else:
        payload = text
    if ", Region:" in payload:
        scenario_part, region_part = payload.split(", Region:", 1)
        scenario = scenario_part.strip().rstrip(",")
        region = region_part.strip()
    else:
        scenario = payload.strip()
    region = re.sub(r",\s*All Fuels\s*$", "", region, flags=re.IGNORECASE).strip()
    return scenario, region


def _unit_factor_from_units_text(units_text: object) -> float:
    text = _normalize_key_text(units_text)
    if "thousand petajoule" in text:
        return 1000.0
    if "petajoule" in text:
        return 1.0
    return 1.0


def _find_results_table_header_row(raw: pd.DataFrame) -> int | None:
    if raw.empty:
        return None
    for idx in range(len(raw)):
        first_cell = _normalize_key_text(raw.iloc[idx, 0] if raw.shape[1] > 0 else "")
        if first_cell not in {"fuel", "branch"}:
            continue
        row = raw.iloc[idx].tolist()
        has_year = any(_coerce_year_from_header(cell) is not None for cell in row[1:])
        if has_year:
            return int(idx)
    return None


def _load_transport_results_tables(
    leap_results_file: Path,
    scenario: str,
    region: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    long_rows: list[dict[str, object]] = []
    inventory_rows: list[dict[str, object]] = []
    row_id_counter = 0

    scenario_key = _normalize_key_text(scenario)
    region_key = _normalize_key_text(region)

    with pd.ExcelFile(leap_results_file) as workbook:
        for sheet_name in workbook.sheet_names:
            raw = workbook.parse(sheet_name=sheet_name, header=None)
            a1 = raw.iloc[0, 0] if raw.shape[0] > 0 and raw.shape[1] > 0 else ""
            a2 = raw.iloc[1, 0] if raw.shape[0] > 1 and raw.shape[1] > 0 else ""
            a3 = raw.iloc[2, 0] if raw.shape[0] > 2 and raw.shape[1] > 0 else ""
            a4 = raw.iloc[3, 0] if raw.shape[0] > 3 and raw.shape[1] > 0 else ""
            parsed_scenario, parsed_region = _parse_scenario_region_from_results_metadata(a2)
            branch_path = ""
            if _normalize_key_text(a3).startswith("branch:"):
                branch_path = _normalize_text(a3).split(":", 1)[1].strip()

            include = True
            status = "accepted"
            reason = ""
            if _normalize_key_text(a1) != "final energy demand":
                include = False
                status = "skipped"
                reason = "a1_not_final_energy_demand"
            elif not branch_path:
                include = False
                status = "skipped"
                reason = "a3_missing_branch_prefix"
            elif parsed_scenario and _normalize_key_text(parsed_scenario) != scenario_key:
                include = False
                status = "skipped"
                reason = "scenario_mismatch"
            elif parsed_region and _normalize_key_text(parsed_region) != region_key:
                include = False
                status = "skipped"
                reason = "region_mismatch"

            header_row = _find_results_table_header_row(raw) if include else None
            year_cols: list[int] = []
            first_col_name = ""
            parsed_rows_count = 0
            if include and header_row is None:
                status = "error"
                reason = "header_row_not_found"
                include = False

            if include and header_row is not None:
                table = raw.iloc[header_row + 1 :].copy()
                table.columns = raw.iloc[header_row].tolist()
                table = table.reset_index(drop=True)
                table, year_cols = _normalize_year_columns(table)
                if not len(table.columns):
                    status = "error"
                    reason = "empty_table_after_header"
                    include = False
                else:
                    first_col_name = _normalize_text(table.columns[0])
                if include and not year_cols:
                    status = "error"
                    reason = "no_year_columns_found"
                    include = False

            unit_factor = _unit_factor_from_units_text(a4)
            if include and header_row is not None:
                first_col = table.columns[0]
                for _, row in table.iterrows():
                    label = _normalize_text(row.get(first_col))
                    if not label or _normalize_key_text(label) in {"nan", "none"}:
                        continue
                    row_id = row_id_counter
                    row_id_counter += 1
                    point_count = 0
                    for year in year_cols:
                        value = pd.to_numeric(row.get(year), errors="coerce")
                        if pd.isna(value):
                            continue
                        point_count += 1
                        long_rows.append(
                            {
                                "__row_id": row_id,
                                "sheet_name": sheet_name,
                                "branch_path": branch_path,
                                "row_label": label,
                                "year": int(year),
                                "value": float(value) * unit_factor,
                                "units_raw": _normalize_text(a4),
                                "unit_factor": unit_factor,
                                "scenario": parsed_scenario,
                                "region": parsed_region,
                            }
                        )
                    parsed_rows_count += int(point_count > 0)

            inventory_rows.append(
                {
                    "sheet_name": sheet_name,
                    "a1": _normalize_text(a1),
                    "a2": _normalize_text(a2),
                    "a3": _normalize_text(a3),
                    "a4": _normalize_text(a4),
                    "parsed_scenario": parsed_scenario,
                    "parsed_region": parsed_region,
                    "branch_path": branch_path,
                    "status": status,
                    "reason": reason,
                    "header_row": header_row,
                    "first_column_name": first_col_name,
                    "year_start": min(year_cols) if year_cols else pd.NA,
                    "year_end": max(year_cols) if year_cols else pd.NA,
                    "parsed_row_count": parsed_rows_count,
                    "unit_factor": unit_factor,
                }
            )

    long_df = pd.DataFrame(long_rows)
    inventory_df = pd.DataFrame(inventory_rows)
    if not inventory_df.empty:
        inventory_df = inventory_df.sort_values("sheet_name").reset_index(drop=True)
    if not long_df.empty:
        long_df = long_df.sort_values(
            ["sheet_name", "branch_path", "row_label", "year"]
        ).reset_index(drop=True)
    return long_df, inventory_df


def _load_transport_branch_mapping(path: Path) -> pd.DataFrame:
    mapping = read_config_table(path).fillna("")
    missing = [
        col
        for col in TRANSPORT_BRANCH_MAPPING_COLUMNS
        if col not in mapping.columns
    ]
    if missing:
        raise ValueError(
            f"Transport branch mapping CSV '{path}' is missing required columns: {missing}"
        )
    mapping = mapping.copy()
    mapping["active"] = mapping["active"].apply(_parse_active_flag)
    mapping["include_in_demand_total"] = mapping["include_in_demand_total"].apply(
        _parse_active_flag
    )
    mapping = mapping[mapping["active"]].copy()
    mapping["branch_path"] = mapping["branch_path"].astype(str).str.strip()
    mapping["ninth_sector_code"] = mapping["ninth_sector_code"].astype(str).str.strip()
    mapping = mapping[
        (mapping["branch_path"] != "") & (mapping["ninth_sector_code"] != "")
    ].copy()
    return mapping.reset_index(drop=True)


def _load_transport_fuel_aliases(path: Path) -> pd.DataFrame:
    aliases = read_config_table(path).fillna("")
    missing = [col for col in TRANSPORT_FUEL_ALIAS_COLUMNS if col not in aliases.columns]
    if missing:
        raise ValueError(
            f"Transport fuel aliases CSV '{path}' is missing required columns: {missing}"
        )
    aliases = aliases.copy()
    aliases["active"] = aliases["active"].apply(_parse_active_flag)
    aliases = aliases[aliases["active"]].copy()
    aliases["leap_fuel_label"] = aliases["leap_fuel_label"].astype(str).str.strip()
    aliases["codebook_name"] = aliases["codebook_name"].astype(str).str.strip()
    aliases["ninth_fuel_override"] = aliases["ninth_fuel_override"].astype(str).str.strip()
    aliases["esto_product_override"] = aliases["esto_product_override"].astype(str).str.strip()
    aliases["__leap_norm"] = aliases["leap_fuel_label"].map(_normalize_lookup_label)
    aliases = aliases[aliases["__leap_norm"] != ""].copy()
    return aliases.reset_index(drop=True)


def _load_code_to_name(path: Path, sheet_name: str) -> pd.DataFrame:
    codebook = read_config_table(path, sheet_name=sheet_name, dtype=str).fillna("")
    required = {"9th_label", "9th_column", "esto_label", "esto_column", "name"}
    missing = required - set(codebook.columns)
    if missing:
        raise ValueError(
            f"Code-to-name sheet '{path}:{sheet_name}' is missing columns: {sorted(missing)}"
        )
    codebook = codebook.copy()
    for col in ["9th_label", "9th_column", "esto_label", "esto_column", "name"]:
        codebook[col] = codebook[col].astype(str).str.strip()
    codebook["__name_norm"] = codebook["name"].map(_normalize_lookup_label)
    return codebook


def _numeric_prefix_tokens(code: object) -> list[str]:
    tokens: list[str] = []
    for part in _normalize_text(code).split("_"):
        if part.isdigit():
            tokens.append(part)
        else:
            break
    return tokens


def _sector_column_from_code(sector_code: object) -> str:
    depth = len(_numeric_prefix_tokens(sector_code))
    if depth not in NINTH_SECTOR_COLUMNS_BY_DEPTH:
        raise ValueError(
            f"Could not infer 9th sector level from code '{_normalize_text(sector_code)}'."
        )
    return NINTH_SECTOR_COLUMNS_BY_DEPTH[depth]


def _build_sector_parent_children_maps(
    codebook_df: pd.DataFrame,
) -> tuple[dict[str, str], dict[str, list[str]]]:
    sector_rows = codebook_df[
        codebook_df["9th_column"].isin(NINTH_SECTOR_COLUMNS_BY_DEPTH.values())
        & codebook_df["9th_label"].ne("")
    ].copy()
    if sector_rows.empty:
        return {}, {}
    sector_rows["__depth"] = sector_rows["9th_label"].map(
        lambda value: len(_numeric_prefix_tokens(value))
    )
    sector_rows = sector_rows[
        sector_rows["__depth"].isin(NINTH_SECTOR_COLUMNS_BY_DEPTH.keys())
    ].copy()

    parent_map: dict[str, str] = {}
    children_map: dict[str, list[str]] = {}
    by_depth: dict[int, pd.DataFrame] = {
        depth: sector_rows[sector_rows["__depth"] == depth].copy()
        for depth in sorted(sector_rows["__depth"].unique())
    }
    for _, row in sector_rows.iterrows():
        child = row["9th_label"]
        depth = int(row["__depth"])
        if depth <= 1:
            continue
        prefix_tokens = _numeric_prefix_tokens(child)
        parent_depth = depth - 1
        parent_prefix = "_".join(prefix_tokens[:parent_depth]) + "_"
        candidates = by_depth.get(parent_depth, pd.DataFrame())
        if candidates.empty:
            continue
        subset = candidates[candidates["9th_label"].str.startswith(parent_prefix)]
        if subset.empty:
            continue
        parent = sorted(subset["9th_label"].unique().tolist())[0]
        parent_map[child] = parent
        children = children_map.get(parent, [])
        if child not in children:
            children.append(child)
        children_map[parent] = sorted(children)
    return parent_map, children_map


def _truthy(value: object) -> bool:
    if isinstance(value, bool):
        return value
    text = _normalize_key_text(value)
    return text in {"1", "true", "yes", "y", "t"}


def _is_apec_aggregate_economy(value: object) -> bool:
    return normalize_economy_key(_normalize_text(value)) == "00APEC"


def _prepare_ninth_rows(
    ninth_df: pd.DataFrame,
    scenario: str,
    economy_key: str,
    aggregate_all_economies: bool = False,
) -> pd.DataFrame:
    working = ninth_df.copy()
    if "scenarios" in working.columns and scenario:
        scenario_key = _normalize_key_text(scenario)
        working = working[
            working["scenarios"].astype(str).str.strip().str.lower() == scenario_key
        ]

    working["economy_key"] = working["economy"].map(normalize_economy_key)
    if not aggregate_all_economies:
        working = working[working["economy_key"] == economy_key].copy()
    for col in ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors", "fuels", "subfuels"]:
        if col in working.columns:
            working[col] = working[col].astype(str).str.strip()
    return working.reset_index(drop=True)


def _build_fuel_column_lookup(codebook_df: pd.DataFrame) -> dict[str, str]:
    rows = codebook_df[
        codebook_df["9th_column"].isin(NINTH_FUEL_COLUMNS)
        & codebook_df["9th_label"].ne("")
    ][["9th_label", "9th_column"]].drop_duplicates()
    return {row["9th_label"]: row["9th_column"] for _, row in rows.iterrows()}


def _aggregate_ninth_series(
    ninth_economy_df: pd.DataFrame,
    sector_code: str,
    fuel_code: str,
    fuel_column: str,
    years: list[int],
    base_year: int | None = None,
) -> dict[int, float]:
    if ninth_economy_df.empty:
        return {}
    try:
        sector_col = _sector_column_from_code(sector_code)
    except ValueError:
        return {}
    if sector_col not in ninth_economy_df.columns:
        return {}
    subset = ninth_economy_df[ninth_economy_df[sector_col] == sector_code].copy()
    if (
        base_year is not None
        and years
        and min(int(year) for year in years) > int(base_year)
        and "subtotal_results" in subset.columns
    ):
        subtotal_flag = subset["subtotal_results"].fillna(False).map(_truthy)
        subset = subset[~subtotal_flag].copy()
    if fuel_code:
        if not fuel_column or fuel_column not in subset.columns:
            return {}
        subset = subset[subset[fuel_column] == fuel_code].copy()
    if subset.empty:
        return {}
    out: dict[int, float] = {}
    for year in years:
        if year not in subset.columns:
            continue
        out[int(year)] = float(pd.to_numeric(subset[year], errors="coerce").fillna(0.0).sum())
    return out


def _resolve_fuel_mapping(
    leap_fuel_label: str,
    aliases_df: pd.DataFrame,
    codebook_df: pd.DataFrame,
    fuel_column_lookup: dict[str, str],
) -> dict[str, object]:
    leap_norm = _normalize_lookup_label(leap_fuel_label)
    alias_row = aliases_df[aliases_df["__leap_norm"] == leap_norm].head(1)
    alias_used = not alias_row.empty
    codebook_name = (
        _normalize_text(alias_row["codebook_name"].iloc[0])
        if alias_used
        else _normalize_text(leap_fuel_label)
    )
    codebook_norm = _normalize_lookup_label(codebook_name)
    name_matches = codebook_df[codebook_df["__name_norm"] == codebook_norm].copy()

    ninth_fuel_override = (
        _normalize_text(alias_row["ninth_fuel_override"].iloc[0]) if alias_used else ""
    )
    esto_product_override = (
        _normalize_text(alias_row["esto_product_override"].iloc[0]) if alias_used else ""
    )

    ninth_fuel_code = ninth_fuel_override
    if not ninth_fuel_code:
        fuel_rows = name_matches[
            name_matches["9th_column"].isin(NINTH_FUEL_COLUMNS)
            & name_matches["9th_label"].ne("")
        ]
        if not fuel_rows.empty:
            ninth_fuel_code = _normalize_text(fuel_rows["9th_label"].iloc[0])

    ninth_fuel_column = (
        fuel_column_lookup.get(ninth_fuel_code, "") if ninth_fuel_code else ""
    )
    if not ninth_fuel_column and ninth_fuel_code:
        fallback = codebook_df[codebook_df["9th_label"] == ninth_fuel_code]
        if not fallback.empty and fallback["9th_column"].iloc[0] in NINTH_FUEL_COLUMNS:
            ninth_fuel_column = _normalize_text(fallback["9th_column"].iloc[0])

    esto_product = esto_product_override
    if not esto_product:
        product_rows = name_matches[
            (name_matches["esto_column"] == "products") & name_matches["esto_label"].ne("")
        ]
        if not product_rows.empty:
            esto_product = _normalize_text(product_rows["esto_label"].iloc[0])

    unresolved_reason = ""
    if not ninth_fuel_code and not esto_product:
        unresolved_reason = "missing_ninth_and_esto_mapping"
    elif not ninth_fuel_code:
        unresolved_reason = "missing_ninth_fuel_mapping"
    elif not esto_product:
        unresolved_reason = "missing_esto_product_mapping"

    return {
        "leap_fuel_label": leap_fuel_label,
        "leap_fuel_norm": leap_norm,
        "alias_used": alias_used,
        "codebook_name": codebook_name,
        "ninth_fuel_code": ninth_fuel_code,
        "ninth_fuel_column": ninth_fuel_column,
        "esto_product": esto_product,
        "resolved": unresolved_reason == "",
        "unresolved_reason": unresolved_reason,
    }


def _build_transport_wide_table(comparison_long: pd.DataFrame) -> pd.DataFrame:
    if comparison_long.empty:
        return pd.DataFrame()
    key_cols = [
        "series_id",
        "branch_path",
        "fuel_label",
        "economy",
        "scenario",
        "region",
        "ninth_sector_codes",
        "ninth_fuel_code",
        "esto_product",
    ]
    working = comparison_long.copy()
    working["year"] = working["year"].astype(int)
    parts: list[pd.DataFrame] = []
    for value_col, prefix in [
        ("leap_value", "leap"),
        ("reference_value", "reference"),
        ("delta", "delta"),
    ]:
        part = (
            working.pivot_table(
                index=key_cols,
                columns="year",
                values=value_col,
                aggfunc="first",
            )
            .reset_index()
        )
        part.columns = [
            col if not isinstance(col, int) else f"{prefix}_{col}" for col in part.columns
        ]
        parts.append(part)
    merged = parts[0]
    for part in parts[1:]:
        merged = merged.merge(part, on=key_cols, how="outer")
    return merged.sort_values(["branch_path", "fuel_label"]).reset_index(drop=True)


def _build_transport_summary_table(
    comparison_long: pd.DataFrame,
    base_year: int,
) -> pd.DataFrame:
    if comparison_long.empty:
        return pd.DataFrame()
    rows: list[dict[str, object]] = []
    for series_id, group in comparison_long.groupby("series_id", dropna=False):
        group = group.sort_values("year")
        base_row = group[group["year"] == int(base_year)]
        base_year_delta = pd.NA
        if not base_row.empty:
            value = pd.to_numeric(base_row["delta"], errors="coerce").dropna()
            if not value.empty:
                base_year_delta = float(value.iloc[0])

        projection_group = group[group["year"] > int(base_year)].copy()
        projection_abs_delta_sum = float(
            pd.to_numeric(projection_group["abs_delta"], errors="coerce").fillna(0.0).sum()
        )
        max_abs_delta = float(
            pd.to_numeric(group["abs_delta"], errors="coerce").fillna(0.0).max()
        )
        pct_vals = pd.to_numeric(group["pct_delta"], errors="coerce").dropna()
        mean_abs_pct_delta = float(pct_vals.abs().mean()) if not pct_vals.empty else pd.NA
        leap_vals = pd.to_numeric(group["leap_value"], errors="coerce")
        ref_vals = pd.to_numeric(group["reference_value"], errors="coerce")
        valid_sign = leap_vals.notna() & ref_vals.notna() & leap_vals.ne(0.0) & ref_vals.ne(0.0)
        sign_mismatch_count = int(((leap_vals * ref_vals) < 0).where(valid_sign, False).sum())

        rows.append(
            {
                "series_id": _normalize_text(series_id),
                "branch_path": _normalize_text(group["branch_path"].iloc[0]),
                "fuel_label": _normalize_text(group["fuel_label"].iloc[0]),
                "economy": _normalize_text(group["economy"].iloc[0]),
                "scenario": _normalize_text(group["scenario"].iloc[0]),
                "region": _normalize_text(group["region"].iloc[0]),
                "ninth_sector_codes": _normalize_text(group["ninth_sector_codes"].iloc[0]),
                "ninth_fuel_code": _normalize_text(group["ninth_fuel_code"].iloc[0]),
                "esto_product": _normalize_text(group["esto_product"].iloc[0]),
                "base_year_delta": base_year_delta,
                "projection_abs_delta_sum": projection_abs_delta_sum,
                "max_abs_delta": max_abs_delta,
                "mean_abs_pct_delta": mean_abs_pct_delta,
                "sign_mismatch_count": sign_mismatch_count,
                "year_count": int(group["year"].nunique()),
            }
        )
    return pd.DataFrame(rows).sort_values(["branch_path", "fuel_label"]).reset_index(drop=True)


def _write_transport_comparison_charts(
    comparison_long: pd.DataFrame,
    charts_dir: Path,
    reference_label: str = "Reference",
) -> None:
    if comparison_long.empty:
        return
    try:
        import matplotlib

        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except Exception as exc:
        print(
            "[WARN] Skipping chart generation because matplotlib is unavailable: "
            f"{exc}"
        )
        return

    charts_dir.mkdir(parents=True, exist_ok=True)
    grouped = comparison_long.groupby("series_id", dropna=False)
    for _, group in grouped:
        working = group.sort_values("year").copy()
        years = working["year"].astype(int).tolist()
        leap_vals = pd.to_numeric(working["leap_value"], errors="coerce")
        ref_vals = pd.to_numeric(working["reference_value"], errors="coerce")
        delta_vals = pd.to_numeric(working["delta"], errors="coerce")

        fig, (ax_top, ax_bottom) = plt.subplots(
            nrows=2,
            ncols=1,
            figsize=(12, 4.8),
            sharex=True,
            gridspec_kw={"height_ratios": [3.4, 1]},
            constrained_layout=True,
        )
        ax_top.plot(years, leap_vals, label="LEAP", marker="o", markersize=12.0, linewidth=7.2)
        ax_top.plot(
            years,
            ref_vals,
            label=reference_label,
            marker="o",
            markersize=12.0,
            linewidth=7.2,
        )
        ax_top.set_ylabel("PJ", fontsize=28)
        ax_top.grid(True, alpha=0.3)
        ax_top.legend(loc="best", fontsize=26)
        ax_top.tick_params(axis="both", labelsize=26)

        ax_bottom.axhline(0.0, color="black", linewidth=0.8)
        ax_bottom.bar(years, delta_vals.fillna(0.0), color="#4C78A8")
        ax_bottom.set_xlabel("Year", fontsize=28)
        ax_bottom.set_ylabel("Delta", fontsize=28)
        ax_bottom.grid(True, axis="y", alpha=0.3)
        ax_bottom.tick_params(axis="both", labelsize=26)

        branch_slug = _safe_filename_token(
            _normalize_text(working["branch_path"].iloc[0]).replace("\\", "_")
        )
        fuel_slug = _safe_filename_token(_normalize_text(working["fuel_label"].iloc[0]))
        out_name = f"{branch_slug}__{fuel_slug}.png"
        fig.savefig(charts_dir / out_name, dpi=220, bbox_inches="tight", pad_inches=0.03)
        plt.close(fig)


def run_transport_results_table_comparison(
    config: TransportResultsComparisonConfig,
) -> ComparisonArtifacts:
    raise RuntimeError(
        "The transport results-table comparison workflow has been removed. "
        "Use codebase/leap_results_dashboard_workflow.py instead."
    )

    leap_results_file = Path(config.leap_results_file)
    branch_sector_mapping_csv = Path(config.branch_sector_mapping_csv)
    fuel_aliases_csv = Path(config.fuel_aliases_csv)
    code_to_name_path = Path(config.code_to_name_path)
    esto_data_path = Path(config.esto_data_path)
    ninth_data_path = Path(config.ninth_data_path)
    subtotal_mapping_path = Path(config.subtotal_mapping_path)
    ninth_to_esto_mapping_path = Path(config.ninth_to_esto_mapping_path)
    output_dir = Path(config.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    charts_dir = output_dir / "charts"
    charts_dir.mkdir(parents=True, exist_ok=True)

    if config.projection_end_year < config.projection_start_year:
        raise ValueError("projection_end_year must be >= projection_start_year.")

    leap_rows, sheet_inventory = _load_transport_results_tables(
        leap_results_file=leap_results_file,
        scenario=config.scenario,
        region=config.region,
    )
    if leap_rows.empty:
        raise ValueError(
            "No LEAP results-table rows matched the requested scenario/region scope."
        )
    sheet_inventory.to_csv(output_dir / "sheet_inventory.csv", index=False)

    branch_mapping = _load_transport_branch_mapping(branch_sector_mapping_csv)
    aliases_df = _load_transport_fuel_aliases(fuel_aliases_csv)
    codebook_df = _load_code_to_name(code_to_name_path, config.code_to_name_sheet)
    fuel_column_lookup = _build_fuel_column_lookup(codebook_df)
    parent_map, children_map = _build_sector_parent_children_maps(codebook_df)

    branch_to_sectors = (
        branch_mapping.groupby("branch_path", dropna=False)["ninth_sector_code"]
        .apply(lambda series: sorted(set(series.tolist())))
        .to_dict()
    )
    branch_to_include_demand = (
        branch_mapping.groupby("branch_path", dropna=False)["include_in_demand_total"]
        .any()
        .to_dict()
    )

    mapping_pairs = _load_mapping_pairs(ninth_to_esto_mapping_path)
    for col in ["9th_sector", "9th_fuel", "esto_flow", "esto_product"]:
        if col in mapping_pairs.columns:
            mapping_pairs[col] = mapping_pairs[col].astype(str).str.strip()
    mapping_pairs = mapping_pairs[
        mapping_pairs["9th_sector"].ne("")
        & mapping_pairs["9th_fuel"].ne("")
        & mapping_pairs["esto_flow"].ne("")
        & mapping_pairs["esto_product"].ne("")
    ][["9th_sector", "9th_fuel", "esto_flow", "esto_product"]].drop_duplicates()

    ninth_df = _load_ninth_data(ninth_data_path)
    economy_key = normalize_economy_key(config.economy)
    aggregate_all_economies = _is_apec_aggregate_economy(config.economy)
    ninth_rows = _prepare_ninth_rows(
        ninth_df=ninth_df,
        scenario=config.ninth_scenario,
        economy_key=economy_key,
        aggregate_all_economies=aggregate_all_economies,
    )

    esto_df = _load_esto_data(esto_data_path, subtotal_mapping_path)
    if aggregate_all_economies:
        esto_econ = esto_df.copy()
    else:
        esto_econ = esto_df[esto_df["economy_key"] == economy_key].copy()
    if esto_econ.empty:
        raise ValueError(f"No ESTO rows found for economy '{config.economy}'.")
    year_col = config.base_year if config.base_year in esto_econ.columns else str(config.base_year)
    if year_col not in esto_econ.columns:
        raise ValueError(
            f"ESTO data does not include base year column '{config.base_year}'."
        )
    esto_econ[year_col] = pd.to_numeric(esto_econ[year_col], errors="coerce")
    base_grouped = (
        esto_econ.groupby(["flows", "products"], dropna=False)[year_col]
        .sum()
        .reset_index()
    )
    esto_base_lookup = {
        (_normalize_text(row["flows"]), _normalize_text(row["products"])): float(row[year_col])
        for _, row in base_grouped.iterrows()
    }

    target_years_from_leap = sorted(set(leap_rows["year"].astype(int).tolist()))
    target_years = [
        year
        for year in target_years_from_leap
        if year == config.base_year
        or (config.projection_start_year <= year <= config.projection_end_year)
    ]
    if config.base_year not in target_years:
        target_years = [config.base_year] + target_years
    target_years = sorted(set(target_years))
    projection_years = [year for year in target_years if year >= config.projection_start_year]

    scoped_rows = leap_rows[leap_rows["branch_path"].str.lower() != "demand"].copy()
    scoped_rows["fuel_label"] = scoped_rows["row_label"].where(
        scoped_rows["row_label"].str.strip().str.lower() != "total",
        "__total__",
    )
    leap_series = (
        scoped_rows.groupby(["branch_path", "fuel_label", "year"], dropna=False)["value"]
        .sum()
        .reset_index()
        .rename(columns={"value": "leap_value"})
    )

    fuel_labels = sorted(
        set(
            scoped_rows[
                scoped_rows["fuel_label"].ne("__total__")
            ]["fuel_label"].tolist()
        )
    )
    fuel_resolution_cache: dict[str, dict[str, object]] = {}
    fuel_status_rows: list[dict[str, object]] = []
    for fuel_label in fuel_labels:
        resolved = _resolve_fuel_mapping(
            leap_fuel_label=fuel_label,
            aliases_df=aliases_df,
            codebook_df=codebook_df,
            fuel_column_lookup=fuel_column_lookup,
        )
        fuel_resolution_cache[fuel_label] = resolved
        fuel_status_rows.append(
            {
                "fuel_label": fuel_label,
                "alias_used": bool(resolved["alias_used"]),
                "codebook_name": _normalize_text(resolved["codebook_name"]),
                "ninth_fuel_code": _normalize_text(resolved["ninth_fuel_code"]),
                "ninth_fuel_column": _normalize_text(resolved["ninth_fuel_column"]),
                "esto_product": _normalize_text(resolved["esto_product"]),
                "resolved": bool(resolved["resolved"]),
                "unresolved_reason": _normalize_text(resolved["unresolved_reason"]),
            }
        )
    fuel_mapping_status_df = pd.DataFrame(fuel_status_rows).sort_values("fuel_label")
    fuel_mapping_status_df.to_csv(output_dir / "fuel_mapping_status.csv", index=False)

    projection_cache: dict[tuple[str, str, str], dict[int, float]] = {}
    base_cache: dict[tuple[str, str, str], tuple[object, str, dict[str, object]]] = {}

    comparison_rows: list[dict[str, object]] = []
    mapping_status_rows: list[dict[str, object]] = []
    matched_row_ids: set[int] = set()

    branch_fuel_reference: dict[tuple[str, int], float] = {}
    branch_fuel_reference_counts: dict[tuple[str, int], int] = {}

    for (branch_path, fuel_label), group in leap_series.groupby(
        ["branch_path", "fuel_label"], dropna=False
    ):
        leap_by_year = {
            int(row["year"]): float(row["leap_value"])
            for _, row in group.iterrows()
        }
        row_ids = set(
            scoped_rows[
                (scoped_rows["branch_path"] == branch_path)
                & (scoped_rows["fuel_label"] == fuel_label)
            ]["__row_id"].astype(int).tolist()
        )

        is_total = fuel_label == "__total__"
        sectors = branch_to_sectors.get(branch_path, [])
        has_branch_mapping = bool(sectors)

        ninth_fuel_code = ""
        ninth_fuel_column = ""
        esto_product = ""
        has_fuel_mapping = True
        fuel_unresolved_reason = ""
        if not is_total:
            fuel_info = fuel_resolution_cache.get(fuel_label, {})
            ninth_fuel_code = _normalize_text(fuel_info.get("ninth_fuel_code"))
            ninth_fuel_column = _normalize_text(fuel_info.get("ninth_fuel_column"))
            esto_product = _normalize_text(fuel_info.get("esto_product"))
            has_fuel_mapping = bool(ninth_fuel_code or esto_product)
            fuel_unresolved_reason = _normalize_text(fuel_info.get("unresolved_reason"))

        ref_by_year: dict[int, object] = {}
        ref_source_by_year: dict[int, str] = {}
        base_sources: set[str] = set()

        if has_branch_mapping and (is_total or has_fuel_mapping):
            if is_total:
                # Filled after fuel series are processed for this branch.
                pass
            else:
                for sector_code in sectors:
                    cache_key = (sector_code, ninth_fuel_code, ninth_fuel_column)
                    if cache_key not in projection_cache:
                        projection_cache[cache_key] = _aggregate_ninth_series(
                            ninth_economy_df=ninth_rows,
                            sector_code=sector_code,
                            fuel_code=ninth_fuel_code,
                            fuel_column=ninth_fuel_column,
                            years=projection_years,
                            base_year=config.base_year,
                        )
                    series = projection_cache[cache_key]
                    for year, value in series.items():
                        ref_by_year[year] = float(ref_by_year.get(year, 0.0)) + float(value)
                        ref_source_by_year[year] = "ninth_projection"

                    base_key = (sector_code, ninth_fuel_code, ninth_fuel_column)
                    if base_key not in base_cache:
                        direct_pairs = mapping_pairs[
                            (mapping_pairs["9th_sector"] == sector_code)
                            & (mapping_pairs["9th_fuel"] == ninth_fuel_code)
                        ]
                        if not direct_pairs.empty:
                            value_sum = 0.0
                            found = False
                            for _, pair in direct_pairs.iterrows():
                                lookup_key = (
                                    _normalize_text(pair["esto_flow"]),
                                    _normalize_text(pair["esto_product"]),
                                )
                                if lookup_key in esto_base_lookup:
                                    value_sum += float(esto_base_lookup[lookup_key])
                                    found = True
                            if found:
                                base_cache[base_key] = (
                                    float(value_sum),
                                    "esto_base_year_direct",
                                    {},
                                )
                            else:
                                base_cache[base_key] = (pd.NA, "missing_base_reference", {})
                        else:
                            parent_sector = parent_map.get(sector_code, "")
                            parent_pairs = mapping_pairs[
                                (mapping_pairs["9th_sector"] == parent_sector)
                                & (mapping_pairs["9th_fuel"] == ninth_fuel_code)
                            ]
                            if not parent_sector or parent_pairs.empty:
                                base_cache[base_key] = (
                                    pd.NA,
                                    "missing_base_reference",
                                    {"parent_sector": parent_sector},
                                )
                            else:
                                parent_value = 0.0
                                parent_found = False
                                for _, pair in parent_pairs.iterrows():
                                    lookup_key = (
                                        _normalize_text(pair["esto_flow"]),
                                        _normalize_text(pair["esto_product"]),
                                    )
                                    if lookup_key in esto_base_lookup:
                                        parent_value += float(esto_base_lookup[lookup_key])
                                        parent_found = True
                                if not parent_found:
                                    base_cache[base_key] = (
                                        pd.NA,
                                        "missing_base_reference",
                                        {"parent_sector": parent_sector},
                                    )
                                else:
                                    siblings = children_map.get(parent_sector, [])
                                    if not siblings:
                                        siblings = [sector_code]
                                    if sector_code not in siblings:
                                        siblings = sorted(set(siblings + [sector_code]))

                                    share_year = int(config.base_year + config.share_year_offset)
                                    child_share_series = _aggregate_ninth_series(
                                        ninth_economy_df=ninth_rows,
                                        sector_code=sector_code,
                                        fuel_code=ninth_fuel_code,
                                        fuel_column=ninth_fuel_column,
                                        years=[share_year],
                                        base_year=config.base_year,
                                    )
                                    numerator = float(child_share_series.get(share_year, 0.0))
                                    denominator = 0.0
                                    for sibling in siblings:
                                        sibling_series = _aggregate_ninth_series(
                                            ninth_economy_df=ninth_rows,
                                            sector_code=sibling,
                                            fuel_code=ninth_fuel_code,
                                            fuel_column=ninth_fuel_column,
                                            years=[share_year],
                                            base_year=config.base_year,
                                        )
                                        denominator += float(
                                            sibling_series.get(share_year, 0.0)
                                        )
                                    share_year_used: int | None = share_year
                                    if abs(denominator) <= 1e-12:
                                        fallback_series = _aggregate_ninth_series(
                                            ninth_economy_df=ninth_rows,
                                            sector_code=sector_code,
                                            fuel_code=ninth_fuel_code,
                                            fuel_column=ninth_fuel_column,
                                            years=[config.base_year],
                                            base_year=config.base_year,
                                        )
                                        numerator = float(
                                            fallback_series.get(config.base_year, 0.0)
                                        )
                                        denominator = 0.0
                                        for sibling in siblings:
                                            sibling_series = _aggregate_ninth_series(
                                                ninth_economy_df=ninth_rows,
                                                sector_code=sibling,
                                                fuel_code=ninth_fuel_code,
                                                fuel_column=ninth_fuel_column,
                                                years=[config.base_year],
                                                base_year=config.base_year,
                                            )
                                            denominator += float(
                                                sibling_series.get(config.base_year, 0.0)
                                            )
                                        share_year_used = int(config.base_year)
                                    if abs(denominator) <= 1e-12:
                                        share = 1.0 / float(len(siblings))
                                        share_year_used = None
                                    else:
                                        share = numerator / denominator
                                    base_cache[base_key] = (
                                        float(parent_value) * float(share),
                                        "esto_base_year_allocated_from_parent",
                                        {
                                            "parent_sector": parent_sector,
                                            "share_year_used": share_year_used,
                                            "share": share,
                                        },
                                    )

                    base_value, base_source, _details = base_cache[base_key]
                    if pd.notna(base_value):
                        ref_by_year[config.base_year] = float(
                            ref_by_year.get(config.base_year, 0.0)
                        ) + float(base_value)
                        base_sources.add(base_source)

                is_bunker = any(
                    _normalize_text(sector).startswith("04_")
                    or _normalize_text(sector).startswith("05_")
                    for sector in sectors
                )
                if is_bunker:
                    for year, value in list(ref_by_year.items()):
                        if pd.notna(value):
                            ref_by_year[year] = abs(float(value))

                if config.base_year in ref_by_year:
                    if base_sources == {"esto_base_year_direct"}:
                        ref_source_by_year[config.base_year] = "esto_base_year_direct"
                    elif "esto_base_year_allocated_from_parent" in base_sources:
                        ref_source_by_year[config.base_year] = (
                            "esto_base_year_allocated_from_parent"
                        )
                    else:
                        ref_source_by_year[config.base_year] = "esto_base_year_mixed"

                matched_row_ids.update(row_ids)

                for year in projection_years:
                    if year in ref_by_year:
                        key = (branch_path, int(year))
                        branch_fuel_reference[key] = float(
                            branch_fuel_reference.get(key, 0.0)
                        ) + float(ref_by_year[year])
                        branch_fuel_reference_counts[key] = (
                            int(branch_fuel_reference_counts.get(key, 0)) + 1
                        )
                if config.base_year in ref_by_year:
                    key = (branch_path, int(config.base_year))
                    branch_fuel_reference[key] = float(
                        branch_fuel_reference.get(key, 0.0)
                    ) + float(ref_by_year[config.base_year])
                    branch_fuel_reference_counts[key] = (
                        int(branch_fuel_reference_counts.get(key, 0)) + 1
                    )

        series_id = f"{branch_path}|{fuel_label}"
        missing_reference_years: list[int] = []
        for year in target_years:
            leap_value = leap_by_year.get(year, pd.NA)
            reference_value = ref_by_year.get(year, pd.NA)
            reference_source = ref_source_by_year.get(year, "missing_reference")
            if pd.isna(reference_value):
                missing_reference_years.append(year)
            delta = pd.NA
            abs_delta = pd.NA
            pct_delta = pd.NA
            if pd.notna(leap_value) and pd.notna(reference_value):
                delta = float(leap_value) - float(reference_value)
                abs_delta = abs(delta)
                if abs(float(reference_value)) > 1e-12:
                    pct_delta = delta / float(reference_value)

            comparison_rows.append(
                {
                    "series_id": series_id,
                    "branch_path": branch_path,
                    "fuel_label": fuel_label,
                    "economy": config.economy,
                    "scenario": config.scenario,
                    "region": config.region,
                    "ninth_sector_codes": ";".join(sectors),
                    "ninth_fuel_code": ninth_fuel_code,
                    "esto_product": esto_product,
                    "year": int(year),
                    "leap_value": leap_value,
                    "reference_value": reference_value,
                    "delta": delta,
                    "abs_delta": abs_delta,
                    "pct_delta": pct_delta,
                    "reference_source": reference_source,
                }
            )

        mapping_status_rows.append(
            {
                "series_id": series_id,
                "branch_path": branch_path,
                "fuel_label": fuel_label,
                "is_total_series": bool(is_total),
                "has_branch_mapping": bool(has_branch_mapping),
                "has_fuel_mapping": bool(is_total or has_fuel_mapping),
                "fuel_unresolved_reason": fuel_unresolved_reason,
                "ninth_sector_codes": ";".join(sectors),
                "ninth_fuel_code": ninth_fuel_code,
                "ninth_fuel_column": ninth_fuel_column,
                "esto_product": esto_product,
                "leap_points": int(len(group)),
                "has_base_year_reference": bool(config.base_year in ref_by_year),
                "has_projection_reference": bool(
                    any(year in ref_by_year for year in projection_years)
                ),
                "missing_reference_years_count": int(len(missing_reference_years)),
                "missing_reference_years": ";".join(str(year) for year in missing_reference_years),
            }
        )

    comparison_long = pd.DataFrame(comparison_rows)
    if comparison_long.empty:
        raise ValueError("No comparison rows were produced for transport results tables.")

    # Build branch total references from summed fuel references.
    total_series_rows: list[dict[str, object]] = []
    total_status_rows: list[dict[str, object]] = []
    branch_totals = leap_series[leap_series["fuel_label"] == "__total__"].copy()
    for branch_path, group in branch_totals.groupby("branch_path", dropna=False):
        leap_by_year = {
            int(row["year"]): float(row["leap_value"])
            for _, row in group.iterrows()
        }
        ref_by_year = {
            year: branch_fuel_reference.get((branch_path, year), pd.NA)
            for year in target_years
        }
        ref_source_by_year = {
            year: (
                "sum_branch_fuels"
                if (branch_path, year) in branch_fuel_reference_counts
                else "missing_reference"
            )
            for year in target_years
        }
        series_id = f"{branch_path}|__total__"
        missing_reference_years: list[int] = []
        for year in target_years:
            leap_value = leap_by_year.get(year, pd.NA)
            reference_value = ref_by_year.get(year, pd.NA)
            if pd.isna(reference_value):
                missing_reference_years.append(year)
            delta = pd.NA
            abs_delta = pd.NA
            pct_delta = pd.NA
            if pd.notna(leap_value) and pd.notna(reference_value):
                delta = float(leap_value) - float(reference_value)
                abs_delta = abs(delta)
                if abs(float(reference_value)) > 1e-12:
                    pct_delta = delta / float(reference_value)
            total_series_rows.append(
                {
                    "series_id": series_id,
                    "branch_path": branch_path,
                    "fuel_label": "__total__",
                    "economy": config.economy,
                    "scenario": config.scenario,
                    "region": config.region,
                    "ninth_sector_codes": ";".join(branch_to_sectors.get(branch_path, [])),
                    "ninth_fuel_code": "",
                    "esto_product": "",
                    "year": int(year),
                    "leap_value": leap_value,
                    "reference_value": reference_value,
                    "delta": delta,
                    "abs_delta": abs_delta,
                    "pct_delta": pct_delta,
                    "reference_source": ref_source_by_year.get(year, "missing_reference"),
                }
            )
        total_status_rows.append(
            {
                "series_id": series_id,
                "branch_path": branch_path,
                "fuel_label": "__total__",
                "is_total_series": True,
                "has_branch_mapping": bool(branch_to_sectors.get(branch_path)),
                "has_fuel_mapping": True,
                "fuel_unresolved_reason": "",
                "ninth_sector_codes": ";".join(branch_to_sectors.get(branch_path, [])),
                "ninth_fuel_code": "",
                "ninth_fuel_column": "",
                "esto_product": "",
                "leap_points": int(len(group)),
                "has_base_year_reference": bool(
                    pd.notna(ref_by_year.get(config.base_year, pd.NA))
                ),
                "has_projection_reference": bool(
                    any(pd.notna(ref_by_year.get(year, pd.NA)) for year in projection_years)
                ),
                "missing_reference_years_count": int(len(missing_reference_years)),
                "missing_reference_years": ";".join(str(year) for year in missing_reference_years),
            }
        )

    comparison_long = pd.concat(
        [comparison_long, pd.DataFrame(total_series_rows)], ignore_index=True
    )
    mapping_status = pd.DataFrame(mapping_status_rows + total_status_rows)

    # Demand total from mapped component branch totals.
    demand_branches = sorted(
        [
            branch
            for branch, include in branch_to_include_demand.items()
            if include and branch in set(branch_totals["branch_path"].unique().tolist())
        ]
    )
    if demand_branches:
        leap_demand_by_year: dict[int, float] = {}
        ref_demand_by_year: dict[int, float] = {}
        for branch in demand_branches:
            branch_leap_rows = branch_totals[branch_totals["branch_path"] == branch]
            for _, row in branch_leap_rows.iterrows():
                year = int(row["year"])
                leap_demand_by_year[year] = float(
                    leap_demand_by_year.get(year, 0.0) + float(row["leap_value"])
                )
            for year in target_years:
                ref_value = branch_fuel_reference.get((branch, year), pd.NA)
                if pd.notna(ref_value):
                    ref_demand_by_year[year] = float(
                        ref_demand_by_year.get(year, 0.0) + float(ref_value)
                    )

        demand_rows: list[dict[str, object]] = []
        missing_reference_years: list[int] = []
        for year in target_years:
            leap_value = leap_demand_by_year.get(year, pd.NA)
            reference_value = ref_demand_by_year.get(year, pd.NA)
            if pd.isna(reference_value):
                missing_reference_years.append(year)
            delta = pd.NA
            abs_delta = pd.NA
            pct_delta = pd.NA
            if pd.notna(leap_value) and pd.notna(reference_value):
                delta = float(leap_value) - float(reference_value)
                abs_delta = abs(delta)
                if abs(float(reference_value)) > 1e-12:
                    pct_delta = delta / float(reference_value)
            demand_rows.append(
                {
                    "series_id": "Demand|__total__",
                    "branch_path": "Demand",
                    "fuel_label": "__total__",
                    "economy": config.economy,
                    "scenario": config.scenario,
                    "region": config.region,
                    "ninth_sector_codes": ";".join(
                        sorted(
                            set(
                                sector
                                for branch in demand_branches
                                for sector in branch_to_sectors.get(branch, [])
                            )
                        )
                    ),
                    "ninth_fuel_code": "",
                    "esto_product": "",
                    "year": int(year),
                    "leap_value": leap_value,
                    "reference_value": reference_value,
                    "delta": delta,
                    "abs_delta": abs_delta,
                    "pct_delta": pct_delta,
                    "reference_source": (
                        "sum_component_branch_totals"
                        if pd.notna(reference_value)
                        else "missing_reference"
                    ),
                }
            )
        comparison_long = pd.concat(
            [comparison_long, pd.DataFrame(demand_rows)], ignore_index=True
        )
        mapping_status = pd.concat(
            [
                mapping_status,
                pd.DataFrame(
                    [
                        {
                            "series_id": "Demand|__total__",
                            "branch_path": "Demand",
                            "fuel_label": "__total__",
                            "is_total_series": True,
                            "has_branch_mapping": True,
                            "has_fuel_mapping": True,
                            "fuel_unresolved_reason": "",
                            "ninth_sector_codes": ";".join(
                                sorted(
                                    set(
                                        sector
                                        for branch in demand_branches
                                        for sector in branch_to_sectors.get(branch, [])
                                    )
                                )
                            ),
                            "ninth_fuel_code": "",
                            "ninth_fuel_column": "",
                            "esto_product": "",
                            "leap_points": int(len(target_years)),
                            "has_base_year_reference": bool(
                                pd.notna(ref_demand_by_year.get(config.base_year, pd.NA))
                            ),
                            "has_projection_reference": bool(
                                any(
                                    pd.notna(ref_demand_by_year.get(year, pd.NA))
                                    for year in projection_years
                                )
                            ),
                            "missing_reference_years_count": int(len(missing_reference_years)),
                            "missing_reference_years": ";".join(
                                str(year) for year in missing_reference_years
                            ),
                        }
                    ]
                ),
            ],
            ignore_index=True,
        )

    comparison_long = comparison_long.sort_values(["branch_path", "fuel_label", "year"]).reset_index(drop=True)
    mapping_status = mapping_status.sort_values(["branch_path", "fuel_label"]).reset_index(drop=True)

    unmatched = scoped_rows[
        (~scoped_rows["__row_id"].isin(matched_row_ids))
        & (scoped_rows["fuel_label"] != "__total__")
    ].copy()
    if not unmatched.empty:
        branch_mapped = unmatched["branch_path"].map(lambda value: value in branch_to_sectors)
        fuel_mapped = unmatched["fuel_label"].map(
            lambda value: bool(fuel_resolution_cache.get(value, {}).get("resolved"))
        )
        unmatched["unmatched_reason"] = pd.NA
        unmatched.loc[~branch_mapped, "unmatched_reason"] = "branch_not_mapped"
        unmatched.loc[branch_mapped & ~fuel_mapped, "unmatched_reason"] = "fuel_not_mapped"
        unmatched = unmatched[
            [
                "__row_id",
                "sheet_name",
                "branch_path",
                "row_label",
                "year",
                "value",
                "unmatched_reason",
            ]
        ].sort_values(["sheet_name", "branch_path", "row_label", "year"]).reset_index(drop=True)
    else:
        unmatched = pd.DataFrame(
            columns=[
                "__row_id",
                "sheet_name",
                "branch_path",
                "row_label",
                "year",
                "value",
                "unmatched_reason",
            ]
        )

    comparison_wide = _build_transport_wide_table(comparison_long)
    comparison_summary = _build_transport_summary_table(comparison_long, config.base_year)
    ninth_scenario_label = _normalize_text(config.ninth_scenario)
    if ninth_scenario_label:
        ninth_scenario_label = (
            ninth_scenario_label[:1].upper() + ninth_scenario_label[1:]
        )
    reference_label = f"ESTO {config.base_year}, 9th {ninth_scenario_label or 'Reference'}"
    _write_transport_comparison_charts(
        comparison_long,
        charts_dir,
        reference_label=reference_label,
    )

    comparison_long_csv = output_dir / "comparison_long.csv"
    comparison_wide_csv = output_dir / "comparison_wide.csv"
    comparison_summary_csv = output_dir / "comparison_summary.csv"
    mapping_status_csv = output_dir / "mapping_status.csv"
    unmatched_leap_rows_csv = output_dir / "unmatched_leap_rows.csv"

    comparison_long.to_csv(comparison_long_csv, index=False)
    comparison_wide.to_csv(comparison_wide_csv, index=False)
    comparison_summary.to_csv(comparison_summary_csv, index=False)
    mapping_status.to_csv(mapping_status_csv, index=False)
    unmatched.to_csv(unmatched_leap_rows_csv, index=False)

    return ComparisonArtifacts(
        comparison_long_csv=comparison_long_csv,
        comparison_wide_csv=comparison_wide_csv,
        comparison_summary_csv=comparison_summary_csv,
        mapping_status_csv=mapping_status_csv,
        unmatched_leap_rows_csv=unmatched_leap_rows_csv,
        charts_dir=charts_dir,
    )


__all__ = [
    "ComparisonRunConfig",
    "ComparisonArtifacts",
    "run_leap_series_comparison",
]
