from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal

import pandas as pd

from codebase.utilities.master_config import read_config_table
from codebase.functions.leap_excel_io import read_export_sheet, save_export_files
from codebase.functions.leap_series_adapter import (
    SeriesFormat,
    collect_available_years,
    detect_series_format,
    extract_row_series,
    inject_row_series,
    scale_row_series,
    sum_rows_series,
)
from codebase.functions.leap_labels import clean_fuel_label_for_leap
from codebase.functions.ninth_projection_mapping import (
    build_esto_base_year_values,
    compute_esto_base_year_shares,
    filter_ninth_projection_rows,
    normalize_economy_key,
)
from codebase.scrapbook.utilities import (
    apply_matt_subtotal_mapping,
    filter_matt_subtotals,
    load_augmented_reference_tables,
)
from codebase.utilities.workflow_common import archive_config_dir_once_per_day


DEFAULT_ESTO_INDUSTRY_FLOW = "14 Industry sector"
DEFAULT_HYDROGEN_SUBFUELS = ["16_x_ammonia", "16_x_efuel", "16_x_hydrogen"]
CURRENT_ACCOUNT_LABELS = {"current accounts", "current account"}


@dataclass
class MappingRow:
    industry_fuel: str
    canonical_industry_fuel: str
    mapping_mode: str
    target_fuels: list[str]
    notes: str


def _clean_level_value(value) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none"}:
        return ""
    return text


def _normalize_year_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {col: int(col) for col in df.columns if str(col).isdigit()}
    return df.rename(columns=rename_map)


def _load_esto_data(path: Path, subtotal_mapping_path: Path) -> pd.DataFrame:
    archive_config_dir_once_per_day()
    df, _ = load_augmented_reference_tables(
        esto_path=path,
        ninth_path=Path("data/merged_file_energy_ALL_20251106.csv"),
        subtotal_mapping_path=subtotal_mapping_path,
        synthetic_rules_path=Path("config/synthetic_reference_rows.csv"),
        cache_dir=Path("data/.cache/industry_reference_tables"),
        apply_esto_subtotal_map=True,
        filter_esto_subtotals_flag=True,
        filter_ninth_subtotals_flag=False,
    )
    df = _normalize_year_columns(df)
    df["economy_key"] = df["economy"].apply(normalize_economy_key)
    df["flows"] = df["flows"].astype(str).str.strip()
    df["products"] = df["products"].astype(str).str.strip()
    return df


def _compute_hydrogen_shares(
    ninth_df: pd.DataFrame,
    economy_key: str,
    scenario: str,
    sector: str,
    subfuels: list[str],
) -> dict[str, dict[int, float]]:
    if ninth_df.empty:
        return {subfuel: {} for subfuel in subfuels}
    filtered = filter_ninth_projection_rows(ninth_df, scenario=scenario)
    filtered = _normalize_year_columns(filtered)
    filtered["economy_key"] = filtered["economy"].apply(normalize_economy_key)
    subset = filtered[
        (filtered["economy_key"] == economy_key)
        & (filtered["sectors"] == sector)
        & (filtered["subfuels"].isin(subfuels))
    ].copy()
    year_cols = [col for col in subset.columns if isinstance(col, int)]
    if not year_cols:
        return {subfuel: {} for subfuel in subfuels}
    grouped = subset.groupby("subfuels", dropna=False)[year_cols].sum()
    grouped = grouped.reindex(subfuels).fillna(0.0)
    totals = grouped.sum(axis=0)
    shares_by_subfuel: dict[str, dict[int, float]] = {subfuel: {} for subfuel in subfuels}
    for year in year_cols:
        total = float(totals.loc[year])
        if total > 0:
            for subfuel in subfuels:
                shares_by_subfuel[subfuel][int(year)] = float(grouped.loc[subfuel, year]) / total
        else:
            for subfuel in subfuels:
                shares_by_subfuel[subfuel][int(year)] = 0.0

    # Fill zero-total years with previous non-zero shares or equal split.
    years_sorted = sorted(int(year) for year in year_cols)
    prev_shares = None
    for year in years_sorted:
        year_shares = {subfuel: shares_by_subfuel[subfuel].get(year, 0.0) for subfuel in subfuels}
        if sum(year_shares.values()) > 0:
            prev_shares = year_shares
            continue
        if prev_shares is not None:
            for subfuel in subfuels:
                shares_by_subfuel[subfuel][year] = prev_shares[subfuel]
        else:
            equal = 1.0 / len(subfuels)
            for subfuel in subfuels:
                shares_by_subfuel[subfuel][year] = equal
    return shares_by_subfuel


def _load_mapping(mapping_path: Path) -> dict[str, MappingRow]:
    df = read_config_table(mapping_path).fillna("")
    rows = {}
    for _, row in df.iterrows():
        industry_fuel = str(row.get("industry_fuel", "")).strip()
        if not industry_fuel:
            continue
        canonical = str(row.get("canonical_industry_fuel", "")).strip() or industry_fuel
        mapping_mode = str(row.get("mapping_mode", "")).strip().lower()
        if not mapping_mode:
            mapping_type = str(row.get("mapping_type", "")).strip().lower()
            if mapping_type == "direct":
                mapping_mode = "direct"
            elif mapping_type == "aggregate_split":
                mapping_mode = "split_base_year"
            else:
                mapping_mode = "direct"
        target_fuels_raw = str(row.get("target_fuels", "")).strip()
        if not target_fuels_raw:
            target_fuels_raw = str(row.get("esto_products", "")).strip()
        target_fuels = [part.strip() for part in target_fuels_raw.split(";") if part.strip()]
        notes = str(row.get("notes", "")).strip()
        rows[industry_fuel] = MappingRow(
            industry_fuel=industry_fuel,
            canonical_industry_fuel=canonical,
            mapping_mode=mapping_mode,
            target_fuels=target_fuels,
            notes=notes,
        )
    return rows


def _load_subtotal_products_for_flow(
    subtotal_mapping_path: Path,
    flow_name: str,
) -> set[str]:
    """Return subtotal products flagged for a specific ESTO flow."""
    try:
        mapping = read_config_table(subtotal_mapping_path, dtype=str)
    except Exception:
        return set()

    mapping = mapping.rename(columns={col: str(col).strip().lower() for col in mapping.columns})
    if {"flows", "products", "is_subtotal"}.issubset(mapping.columns):
        mapping = mapping.rename(columns={"flows": "flow", "products": "product"})
    if not {"flow", "product", "is_subtotal"}.issubset(mapping.columns):
        return set()

    mapping["flow"] = mapping["flow"].astype(str).str.strip()
    mapping["product"] = mapping["product"].astype(str).str.strip()
    mapping["is_subtotal"] = (
        mapping["is_subtotal"]
        .astype(str)
        .str.strip()
        .str.lower()
        .isin(["true", "1", "yes"])
    )
    subset = mapping[
        (mapping["flow"] == str(flow_name).strip()) & mapping["is_subtotal"]
    ]
    return set(subset["product"].dropna().astype(str).str.strip().tolist())


def _resolve_mapping(mapping_by_fuel: dict[str, MappingRow], fuel: str) -> MappingRow | None:
    current = mapping_by_fuel.get(fuel)
    visited = set()
    while current and current.mapping_mode == "alias":
        if current.industry_fuel in visited:
            return None
        visited.add(current.industry_fuel)
        next_fuel = current.canonical_industry_fuel or current.industry_fuel
        current = mapping_by_fuel.get(next_fuel)
    return current


def _is_fuel_row(row: pd.Series, level_cols: list[str]) -> tuple[bool, str, int, str]:
    levels = [_clean_level_value(row.get(col)) for col in level_cols]
    if not levels or levels[0] != "Demand" or (len(levels) > 1 and levels[1] != "Industry"):
        return False, "", -1, ""
    sector = levels[2] if len(levels) > 2 else ""
    if not sector:
        return False, "", -1, ""
    if sector == "Manufacturing":
        fuel_idx = 4
    else:
        fuel_idx = 3
    if fuel_idx >= len(levels):
        return False, "", -1, ""
    fuel = levels[fuel_idx]
    if not fuel:
        return False, "", -1, ""
    last_non_empty = max((idx for idx, val in enumerate(levels) if val), default=-1)
    if last_non_empty != fuel_idx:
        return False, "", -1, ""
    return True, fuel, fuel_idx, sector


def _update_branch_path(row: pd.Series, level_cols: list[str], fuel_idx: int, new_fuel: str) -> pd.Series:
    levels = [_clean_level_value(row.get(col)) for col in level_cols]
    if fuel_idx >= len(levels):
        return row
    levels[fuel_idx] = clean_fuel_label_for_leap(new_fuel)
    branch_path = "\\".join([val for val in levels if val])
    updated = row.copy()
    updated["Branch Path"] = branch_path
    for col, value in zip(level_cols, levels):
        updated[col] = value if value else pd.NA
    return updated


def _rebuild_level_columns(df: pd.DataFrame, level_cols: list[str]) -> pd.DataFrame:
    if not level_cols:
        return df
    max_levels = len(level_cols)
    working = df.copy()
    parts = working["Branch Path"].fillna("").astype(str).str.split("\\")
    for idx in range(max_levels):
        col = level_cols[idx]
        working[col] = parts.str.get(idx).fillna("")
        working[col] = working[col].replace({"": pd.NA})
    return working


def _collect_fuel_series_years(
    df: pd.DataFrame,
    level_cols: list[str],
    series_format: SeriesFormat,
    year_cols: list[int],
) -> set[int]:
    years: set[int] = set()
    for _, row in df.iterrows():
        is_fuel, _fuel, _fuel_idx, _sector = _is_fuel_row(row, level_cols)
        if not is_fuel:
            continue
        series = extract_row_series(
            row,
            series_format=series_format,
            year_cols=year_cols if year_cols else None,
        )
        if isinstance(series, dict):
            years.update(int(year) for year in series.keys())
    return years


def _normalize_key_part(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _extract_series_value_for_year(
    row: pd.Series,
    year: int,
    series_format: SeriesFormat,
    year_cols: list[int] | None = None,
) -> float | None:
    series = extract_row_series(
        row,
        series_format=series_format,
        year_cols=year_cols if year_cols else None,
    )
    if isinstance(series, float):
        return float(series)
    if isinstance(series, dict) and int(year) in series:
        return float(series[int(year)])
    return None


def _anchor_projected_series_to_base_year(
    df: pd.DataFrame,
    base_year: int,
    series_format: SeriesFormat,
    year_cols: list[int],
) -> tuple[pd.DataFrame, dict[str, object]]:
    required_cols = {"Scenario", "Branch Path", "Variable", "Region"}
    if not required_cols.issubset(set(df.columns)):
        return df, {
            "anchored_rows": 0,
            "rows_missing_current_accounts_value": 0,
            "skipped_missing_columns": sorted(required_cols - set(df.columns)),
            "sample_missing_current_accounts_rows": [],
        }

    # Split mappings intentionally duplicate source rows, which can leave duplicate
    # index labels. Use a clean RangeIndex so scalar `.at[...]` updates only touch
    # one row.
    working = df.reset_index(drop=True).copy()
    scenario_norm = working["Scenario"].astype("string").fillna("").str.strip().str.lower()
    is_current_accounts = scenario_norm.isin(CURRENT_ACCOUNT_LABELS)
    key_cols = ["Branch Path", "Variable", "Region"]

    ca_values: dict[tuple[str, str, str], float] = {}
    for _, row in working.loc[is_current_accounts].iterrows():
        key = tuple(_normalize_key_part(row.get(col)) for col in key_cols)
        base_value = _extract_series_value_for_year(
            row,
            base_year,
            series_format=series_format,
            year_cols=year_cols,
        )
        if base_value is None:
            continue
        if key not in ca_values:
            ca_values[key] = float(base_value)

    anchored_rows = 0
    missing_rows = 0
    missing_samples: list[dict[str, str]] = []
    for idx, row in working.loc[~is_current_accounts].iterrows():
        existing_base = _extract_series_value_for_year(
            row,
            base_year,
            series_format=series_format,
            year_cols=year_cols,
        )
        if existing_base is not None:
            continue
        key = tuple(_normalize_key_part(row.get(col)) for col in key_cols)
        base_value = ca_values.get(key)
        if base_value is None:
            missing_rows += 1
            if len(missing_samples) < 5:
                missing_samples.append(
                    {
                        "scenario": _normalize_key_part(row.get("Scenario")),
                        "branch_path": _normalize_key_part(row.get("Branch Path")),
                        "variable": _normalize_key_part(row.get("Variable")),
                    }
                )
            continue
        series = extract_row_series(
            row,
            series_format=series_format,
            year_cols=year_cols if year_cols else None,
        )
        if isinstance(series, float):
            continue
        series_payload = dict(series or {})
        series_payload[int(base_year)] = float(base_value)
        updated_row = inject_row_series(
            row,
            series=series_payload,
            series_format=series_format,
            year_cols=year_cols if year_cols else None,
        )
        working.loc[idx] = updated_row
        anchored_rows += 1

    return working, {
        "anchored_rows": anchored_rows,
        "rows_missing_current_accounts_value": missing_rows,
        "skipped_missing_columns": [],
        "sample_missing_current_accounts_rows": missing_samples,
    }


def _validate_projected_base_year_coverage(
    df: pd.DataFrame,
    base_year: int,
    series_format: SeriesFormat,
    year_cols: list[int],
) -> dict[str, object]:
    required_cols = {"Scenario", "Branch Path", "Variable"}
    if not required_cols.issubset(set(df.columns)):
        return {
            "checked_series_rows": 0,
            "missing_base_year_rows": 0,
            "sample_missing_rows": [],
            "skipped_missing_columns": sorted(required_cols - set(df.columns)),
        }

    working = df.copy()
    scenario_norm = working["Scenario"].astype("string").fillna("").str.strip().str.lower()
    projected = working.loc[~scenario_norm.isin(CURRENT_ACCOUNT_LABELS)]
    missing_rows: list[dict[str, str]] = []
    checked_series_rows = 0
    missing_count = 0
    for _, row in projected.iterrows():
        if series_format == "expression":
            series = extract_row_series(
                row,
                series_format=series_format,
                year_cols=year_cols if year_cols else None,
            )
            if not isinstance(series, dict):
                continue
        checked_series_rows += 1
        if _extract_series_value_for_year(
            row,
            base_year,
            series_format=series_format,
            year_cols=year_cols,
        ) is not None:
            continue
        missing_count += 1
        if len(missing_rows) < 5:
            missing_rows.append(
                {
                    "scenario": _normalize_key_part(row.get("Scenario")),
                    "branch_path": _normalize_key_part(row.get("Branch Path")),
                    "variable": _normalize_key_part(row.get("Variable")),
                }
            )
    return {
        "checked_series_rows": checked_series_rows,
        "missing_base_year_rows": missing_count,
        "sample_missing_rows": missing_rows,
        "skipped_missing_columns": [],
    }


def _extract_model_name(header_rows: pd.DataFrame) -> str:
    if "Variable" in header_rows.columns:
        values = header_rows["Variable"].dropna().astype(str).str.strip()
        values = values[values != ""]
        if not values.empty:
            return values.iloc[0]
    return "Model"


def _build_for_viewing_df_from_series(
    export_df: pd.DataFrame,
    base_year: int,
    series_format: SeriesFormat,
    year_cols: list[int],
) -> tuple[pd.DataFrame, list[int]]:
    if series_format == "year_columns":
        years = year_cols or collect_available_years(export_df, series_format)
        return export_df.copy(), years

    years = year_cols or collect_available_years(export_df, series_format)
    if not years:
        years = [int(base_year)]
    rows_for_years: list[dict[int, object]] = []
    for _, row in export_df.iterrows():
        series = extract_row_series(row, series_format="expression", year_cols=years)
        if not isinstance(series, dict):
            rows_for_years.append({year: pd.NA for year in years})
            continue
        rows_for_years.append({year: series.get(year, pd.NA) for year in years})

    year_values_df = pd.DataFrame(rows_for_years, index=export_df.index)
    viewing_df = pd.concat([export_df.drop(columns=["Expression"]), year_values_df], axis=1)
    level_cols = [col for col in viewing_df.columns if str(col).startswith("Level ")]
    non_level_cols = [col for col in viewing_df.columns if col not in level_cols]
    non_year_cols = [col for col in non_level_cols if not isinstance(col, int)]
    viewing_df = viewing_df[non_year_cols + years + level_cols]
    return viewing_df, years


def _summarize_fuel_values(
    df: pd.DataFrame,
    level_cols: list[str],
    years: list[int],
    base_year: int,
    series_format: SeriesFormat,
) -> tuple[dict[tuple[str, int], float], dict[tuple[str, str, int], float], list[str]]:
    overall: dict[tuple[str, int], float] = {}
    by_sector: dict[tuple[str, str, int], float] = {}
    unknown: list[str] = []
    for _, row in df.iterrows():
        is_fuel, _fuel, _fuel_idx, sector = _is_fuel_row(row, level_cols)
        if not is_fuel:
            continue
        series = extract_row_series(
            row,
            series_format=series_format,
            year_cols=years,
        )
        if not isinstance(series, dict):
            unknown.append(str(row.get("Branch Path", "")))
            continue
        variable = str(row.get("Variable", "")).strip()
        for year, value in series.items():
            overall_key = (variable, int(year))
            overall[overall_key] = overall.get(overall_key, 0.0) + float(value)
            sector_key = (variable, sector, int(year))
            by_sector[sector_key] = by_sector.get(sector_key, 0.0) + float(value)
    return overall, by_sector, unknown


def _build_validation_report(
    before_df: pd.DataFrame,
    after_df: pd.DataFrame,
    level_cols: list[str],
    base_year: int,
    tolerance: float,
    series_format: SeriesFormat,
    year_cols: list[int],
) -> tuple[pd.DataFrame, dict[str, float | int]]:
    years = _collect_fuel_series_years(before_df, level_cols, series_format, year_cols)
    years |= _collect_fuel_series_years(after_df, level_cols, series_format, year_cols)
    if not years:
        years = {int(base_year)}
    year_list = sorted(int(year) for year in years)

    overall_before, sector_before, unknown_before = _summarize_fuel_values(
        before_df, level_cols, year_list, base_year, series_format
    )
    overall_after, sector_after, unknown_after = _summarize_fuel_values(
        after_df, level_cols, year_list, base_year, series_format
    )

    rows = []
    for (variable, year) in sorted(set(overall_before) | set(overall_after)):
        before_val = overall_before.get((variable, year), 0.0)
        after_val = overall_after.get((variable, year), 0.0)
        diff = after_val - before_val
        rows.append(
            {
                "group_type": "overall",
                "variable": variable,
                "sector": "",
                "year": int(year),
                "before_value": before_val,
                "after_value": after_val,
                "diff": diff,
                "abs_diff": abs(diff),
                "within_tolerance": abs(diff) <= tolerance,
            }
        )

    for (variable, sector, year) in sorted(set(sector_before) | set(sector_after)):
        before_val = sector_before.get((variable, sector, year), 0.0)
        after_val = sector_after.get((variable, sector, year), 0.0)
        diff = after_val - before_val
        rows.append(
            {
                "group_type": "sector",
                "variable": variable,
                "sector": sector,
                "year": int(year),
                "before_value": before_val,
                "after_value": after_val,
                "diff": diff,
                "abs_diff": abs(diff),
                "within_tolerance": abs(diff) <= tolerance,
            }
        )

    report = pd.DataFrame(rows)
    summary = {
        "max_abs_diff_overall": float(
            report.loc[report["group_type"] == "overall", "abs_diff"].max()
            if not report.empty
            else 0.0
        ),
        "max_abs_diff_sector": float(
            report.loc[report["group_type"] == "sector", "abs_diff"].max()
            if not report.empty
            else 0.0
        ),
        "unknown_expressions_before": int(len(unknown_before)),
        "unknown_expressions_after": int(len(unknown_after)),
    }
    return report, summary


def _convert_series_format(
    df: pd.DataFrame,
    source_format: SeriesFormat,
    target_format: SeriesFormat,
    base_year: int,
) -> tuple[pd.DataFrame, list[int]]:
    if source_format == target_format:
        return df.copy(), collect_available_years(df, source_format)

    working = df.copy()
    source_years = collect_available_years(working, source_format)
    if not source_years:
        source_years = [int(base_year)]

    if source_format == "expression" and target_format == "year_columns":
        year_values = []
        for _, row in working.iterrows():
            series = extract_row_series(row, "expression", year_cols=source_years)
            if not isinstance(series, dict):
                year_values.append({year: pd.NA for year in source_years})
                continue
            year_values.append({year: series.get(year, pd.NA) for year in source_years})
        year_df = pd.DataFrame(year_values, index=working.index)
        if "Expression" in working.columns:
            working = working.drop(columns=["Expression"])
        working = pd.concat([working, year_df], axis=1)
        return working, source_years

    if source_format == "year_columns" and target_format == "expression":
        expressions = []
        for _, row in working.iterrows():
            series = extract_row_series(row, "year_columns", year_cols=source_years)
            if not isinstance(series, dict):
                expressions.append("")
                continue
            if not series:
                expressions.append("")
                continue
            expression = inject_row_series(
                pd.Series({"Expression": ""}),
                series=series,
                series_format="expression",
                year_cols=source_years,
            ).get("Expression", "")
            expressions.append(expression)
        working["Expression"] = expressions
        year_like_cols = [
            col
            for col in working.columns
            if isinstance(col, int) or (isinstance(col, str) and col.isdigit())
        ]
        if year_like_cols:
            working = working.drop(columns=year_like_cols)
        return working, source_years

    return working, source_years


def remap_industry_export_fuels(
    input_path: str | Path,
    output_path: str | Path,
    mapping_csv_path: str | Path,
    esto_data_path: str | Path,
    ninth_data_path: str | Path,
    subtotal_mapping_path: str | Path,
    economy: str,
    base_year: int,
    scenario: str = "reference",
    sheet_name: str = "Export",
    esto_industry_flow: str = DEFAULT_ESTO_INDUSTRY_FLOW,
    hydrogen_subfuels: list[str] | None = None,
    include_extra_others: bool = False,
    extra_others_products: list[str] | None = None,
    report_path: str | Path | None = None,
    validation_path: str | Path | None = None,
    validation_tolerance: float = 1e-6,
    validate: bool = True,
    ensure_base_year_from_current_accounts: bool = True,
    enforce_base_year_presence: bool = False,
    output_series_format: Literal["preserve", "expression", "year_columns"] = "preserve",
) -> dict[str, object]:
    input_path = Path(input_path)
    output_path = Path(output_path)
    mapping_csv_path = Path(mapping_csv_path)
    esto_data_path = Path(esto_data_path)
    ninth_data_path = Path(ninth_data_path)
    subtotal_mapping_path = Path(subtotal_mapping_path)

    header_rows, df, columns = read_export_sheet(input_path, sheet_name)
    input_series_format = detect_series_format(df)
    input_year_cols = collect_available_years(df, input_series_format)
    mapping_by_fuel = _load_mapping(mapping_csv_path)
    issues: dict[str, object] = {
        "unmapped_fuels": [],
        "extra_others_products": [],
        "subtotal_targets_removed": [],
    }

    level_cols = [col for col in columns if col.startswith("Level ")]
    id_cols = [col for col in ["BranchID", "VariableID", "ScenarioID", "RegionID"] if col in columns]

    economy_key = normalize_economy_key(economy)
    esto_df = _load_esto_data(esto_data_path, subtotal_mapping_path)
    base_values = build_esto_base_year_values(esto_df, base_year)
    ninth_df = pd.read_csv(ninth_data_path)

    if hydrogen_subfuels is None:
        hydrogen_subfuels = DEFAULT_HYDROGEN_SUBFUELS

    subtotal_products = _load_subtotal_products_for_flow(
        subtotal_mapping_path=subtotal_mapping_path,
        flow_name=esto_industry_flow,
    )
    if subtotal_products:
        filtered_mapping: dict[str, MappingRow] = {}
        removed_rows: list[dict[str, str]] = []
        for fuel_key, mapping_row in mapping_by_fuel.items():
            original_targets = list(mapping_row.target_fuels)
            is_others_bucket = str(mapping_row.industry_fuel).strip().lower() == "others"
            if is_others_bucket and mapping_row.mapping_mode in {"direct", "split_base_year"}:
                filtered_targets = [t for t in original_targets if t not in subtotal_products]
            else:
                filtered_targets = original_targets
            removed_targets = [t for t in original_targets if t not in filtered_targets]
            for removed in removed_targets:
                removed_rows.append(
                    {
                        "industry_fuel": fuel_key,
                        "removed_target": removed,
                        "reason": "subtotal_product",
                    }
                )
            filtered_mapping[fuel_key] = MappingRow(
                industry_fuel=mapping_row.industry_fuel,
                canonical_industry_fuel=mapping_row.canonical_industry_fuel,
                mapping_mode=mapping_row.mapping_mode,
                target_fuels=filtered_targets,
                notes=mapping_row.notes,
            )
        mapping_by_fuel = filtered_mapping
        issues["subtotal_targets_removed"] = removed_rows

    hydrogen_shares = _compute_hydrogen_shares(
        ninth_df,
        economy_key=economy_key,
        scenario=scenario,
        sector="14_industry_sector",
        subfuels=hydrogen_subfuels,
    )

    # Build mapping targets for base-year splits.
    base_year_share_cache: dict[tuple[str, str], dict[str, float]] = {}

    def _get_base_year_shares(mapping_row: MappingRow) -> dict[str, float]:
        key = (mapping_row.industry_fuel, mapping_row.mapping_mode)
        if key in base_year_share_cache:
            return base_year_share_cache[key]
        shares = compute_esto_base_year_shares(
            base_values,
            economy_key=economy_key,
            esto_flow=esto_industry_flow,
            esto_products=mapping_row.target_fuels,
        )
        base_year_share_cache[key] = shares
        return shares

    # Detect extra products for Others.
    mapped_products = set()
    for mapping in mapping_by_fuel.values():
        if mapping.mapping_mode in {"direct", "split_base_year"}:
            mapped_products.update(mapping.target_fuels)

    year_col = base_year if base_year in esto_df.columns else str(base_year)
    industry_products = set()
    if year_col in esto_df.columns:
        subset = esto_df[
            (esto_df["economy_key"] == economy_key) & (esto_df["flows"] == esto_industry_flow)
        ].copy()
        subset[year_col] = pd.to_numeric(subset[year_col], errors="coerce").fillna(0.0)
        subset = subset[subset[year_col] != 0]
        industry_products = set(subset["products"].tolist())

    others_row = mapping_by_fuel.get("Others")
    extra_products = []
    if industry_products and others_row:
        extra_products = sorted(industry_products - mapped_products - set(others_row.target_fuels))
        if extra_products:
            issues["extra_others_products"] = extra_products
            if include_extra_others:
                others_row.target_fuels.extend(extra_products)
            elif extra_others_products:
                chosen = [product for product in extra_others_products if product in extra_products]
                others_row.target_fuels.extend(chosen)

    new_rows = []
    for _, row in df.iterrows():
        is_fuel, fuel, fuel_idx, _sector = _is_fuel_row(row, level_cols)
        if not is_fuel:
            new_rows.append(row)
            continue
        mapping_row = _resolve_mapping(mapping_by_fuel, fuel)
        if not mapping_row:
            issues["unmapped_fuels"].append(fuel)
            new_rows.append(row)
            continue

        if mapping_row.mapping_mode == "direct":
            targets = list(mapping_row.target_fuels)
            if not targets:
                issues["unmapped_fuels"].append(f"{fuel} (no non-subtotal targets)")
                new_rows.append(row)
                continue
            shares = {target: 1.0 for target in targets}
            for target in targets:
                updated = _update_branch_path(row, level_cols, fuel_idx, target)
                updated = scale_row_series(
                    updated,
                    scale_by=shares[target],
                    base_year=base_year,
                    series_format=input_series_format,
                    year_cols=input_year_cols if input_year_cols else None,
                )
                for col in id_cols:
                    updated[col] = pd.NA
                new_rows.append(updated)
            continue

        if mapping_row.mapping_mode == "split_base_year":
            if not mapping_row.target_fuels:
                issues["unmapped_fuels"].append(f"{fuel} (no non-subtotal targets)")
                new_rows.append(row)
                continue
            shares = _get_base_year_shares(mapping_row)
            fallback_share = None
            if shares:
                fallback_share = sum(shares.values()) / max(len(shares), 1)
            for target in mapping_row.target_fuels:
                share = shares.get(target, 0.0)
                updated = _update_branch_path(row, level_cols, fuel_idx, target)
                updated = scale_row_series(
                    updated,
                    scale_by=share,
                    base_year=base_year,
                    series_format=input_series_format,
                    year_cols=input_year_cols if input_year_cols else None,
                    fallback_share=fallback_share,
                )
                for col in id_cols:
                    updated[col] = pd.NA
                new_rows.append(updated)
            continue

        if mapping_row.mapping_mode == "split_ninth_hydrogen":
            for target in mapping_row.target_fuels:
                share_by_year = hydrogen_shares.get(target, {})
                fallback = None
                if share_by_year:
                    fallback = sum(share_by_year.values()) / max(len(share_by_year), 1)
                updated = _update_branch_path(row, level_cols, fuel_idx, target)
                updated = scale_row_series(
                    updated,
                    scale_by=share_by_year,
                    base_year=base_year,
                    series_format=input_series_format,
                    year_cols=input_year_cols if input_year_cols else None,
                    fallback_share=fallback,
                )
                for col in id_cols:
                    updated[col] = pd.NA
                new_rows.append(updated)
            continue

        # Fallback: keep as-is
        new_rows.append(row)

    mapped_df = pd.DataFrame(new_rows, columns=columns)

    # Aggregate any duplicate Branch Path + Variable + Scenario + Region rows.
    key_cols = [col for col in ["Branch Path", "Variable", "Scenario", "Region"] if col in mapped_df.columns]
    if key_cols:
        grouped = []
        for _, group in mapped_df.groupby(key_cols, dropna=False):
            if len(group) == 1:
                grouped.append(group.iloc[0])
                continue
            row = group.iloc[0].copy()
            combined = sum_rows_series(
                group,
                series_format=input_series_format,
                year_cols=input_year_cols if input_year_cols else None,
            )
            row = inject_row_series(
                row,
                series=combined,
                series_format=input_series_format,
                year_cols=input_year_cols if input_year_cols else None,
            )
            for col in id_cols:
                row[col] = pd.NA
            grouped.append(row)
        mapped_df = pd.DataFrame(grouped, columns=mapped_df.columns)

    mapped_df = _rebuild_level_columns(mapped_df, level_cols)

    if ensure_base_year_from_current_accounts:
        mapped_df, anchor_summary = _anchor_projected_series_to_base_year(
            mapped_df,
            base_year=base_year,
            series_format=input_series_format,
            year_cols=input_year_cols,
        )
        issues["base_year_anchor_summary"] = anchor_summary
        if anchor_summary.get("anchored_rows", 0):
            print(
                f"[INFO] Added base-year {base_year} points from Current Accounts to "
                f"{anchor_summary['anchored_rows']} projected rows."
            )
        missing_from_ca = int(anchor_summary.get("rows_missing_current_accounts_value", 0))
        if missing_from_ca:
            print(
                f"[WARN] Could not anchor base-year {base_year} for {missing_from_ca} projected rows "
                "because no matching Current Accounts base-year value was found."
            )
            for sample in anchor_summary.get("sample_missing_current_accounts_rows", []):
                print(
                    "[WARN]   "
                    f"{sample.get('scenario', '')} | "
                    f"{sample.get('branch_path', '')} | "
                    f"{sample.get('variable', '')}"
                )

    base_year_coverage = _validate_projected_base_year_coverage(
        mapped_df,
        base_year,
        series_format=input_series_format,
        year_cols=input_year_cols,
    )
    issues["base_year_coverage"] = base_year_coverage
    missing_base_year_rows = int(base_year_coverage.get("missing_base_year_rows", 0))
    if missing_base_year_rows:
        message = (
            f"{missing_base_year_rows} projected series rows are missing base-year {base_year} points."
        )
        print(f"[WARN] {message}")
        for sample in base_year_coverage.get("sample_missing_rows", []):
            print(
                "[WARN]   "
                f"{sample.get('scenario', '')} | "
                f"{sample.get('branch_path', '')} | "
                f"{sample.get('variable', '')}"
            )
        if enforce_base_year_presence:
            raise ValueError(message)

    if output_series_format not in {"preserve", "expression", "year_columns"}:
        raise ValueError(
            "output_series_format must be one of: preserve, expression, year_columns."
        )
    target_series_format: SeriesFormat = (
        input_series_format
        if output_series_format == "preserve"
        else "expression"
        if output_series_format == "expression"
        else "year_columns"
    )
    mapped_df, mapped_years = _convert_series_format(
        mapped_df,
        source_format=input_series_format,
        target_format=target_series_format,
        base_year=base_year,
    )

    for_viewing_df, viewing_years = _build_for_viewing_df_from_series(
        mapped_df,
        base_year=base_year,
        series_format=target_series_format,
        year_cols=mapped_years,
    )
    save_export_files(
        leap_export_df=mapped_df,
        export_df_for_viewing=for_viewing_df,
        leap_export_filename=output_path,
        base_year=base_year,
        final_year=max(viewing_years) if viewing_years else base_year,
        model_name=_extract_model_name(header_rows),
    )

    if report_path:
        report_path = Path(report_path)
        report_rows = []
        for fuel in sorted(set(issues.get("unmapped_fuels", []))):
            report_rows.append(
                {"issue_type": "unmapped_fuel", "detail": fuel, "suggestion": "Add to mapping CSV."}
            )
        for product in issues.get("extra_others_products", []):
            report_rows.append(
                {
                    "issue_type": "extra_others_product",
                    "detail": product,
                    "suggestion": "Consider adding to Others target_fuels.",
                }
            )
        for item in issues.get("subtotal_targets_removed", []):
            report_rows.append(
                {
                    "issue_type": "subtotal_target_removed",
                    "detail": f"{item.get('industry_fuel', '')} -> {item.get('removed_target', '')}",
                    "suggestion": "Removed automatically using ESTO subtotal mapping.",
                }
            )
        if report_rows:
            report_path.parent.mkdir(parents=True, exist_ok=True)
            pd.DataFrame(report_rows).to_csv(report_path, index=False)

    if validate:
        validation_source_df = df
        validation_years = mapped_years
        if target_series_format != input_series_format:
            validation_source_df, validation_years = _convert_series_format(
                df,
                source_format=input_series_format,
                target_format=target_series_format,
                base_year=base_year,
            )
        validation_report, validation_summary = _build_validation_report(
            validation_source_df,
            mapped_df,
            level_cols,
            base_year,
            validation_tolerance,
            series_format=target_series_format,
            year_cols=validation_years,
        )
        issues["validation_summary"] = validation_summary
        if validation_path:
            validation_path = Path(validation_path)
            validation_path.parent.mkdir(parents=True, exist_ok=True)
            validation_report.to_csv(validation_path, index=False)
        if validation_summary:
            print(
                "[INFO] Validation max abs diff overall="
                f"{validation_summary['max_abs_diff_overall']:.6g}, "
                "sector="
                f"{validation_summary['max_abs_diff_sector']:.6g}."
            )

    return issues


__all__ = ["remap_industry_export_fuels"]
