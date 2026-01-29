from __future__ import annotations

from pathlib import Path
from typing import Iterable, Sequence

import pandas as pd

DEFAULT_SCENARIO = "reference"
NINTH_SECTOR_COLS = [
    "sub4sectors",
    "sub3sectors",
    "sub2sectors",
    "sub1sectors",
    "sectors",
]
NINTH_FUEL_COLS = ["subfuels", "fuels"]


def normalize_economy_key(value: str | None) -> str:
    """Return a canonical economy key for cross-dataset joins."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none"}:
        return ""
    return text.replace("_", "").upper()


def _clean_label_series(series: pd.Series) -> pd.Series:
    cleaned = series.fillna("").astype(str).str.strip()
    return cleaned.mask(cleaned.str.lower() == "x", pd.NA)


def add_ninth_pair_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Add most-specific sector/fuel columns for 9th data."""
    working = df.copy()
    sector_cols = [col for col in NINTH_SECTOR_COLS if col in working.columns]
    fuel_cols = [col for col in NINTH_FUEL_COLS if col in working.columns]
    if sector_cols:
        sector_values = pd.DataFrame(
            {col: _clean_label_series(working[col]) for col in sector_cols}
        )
        working["9th_sector"] = sector_values.bfill(axis=1).iloc[:, 0].fillna("")
    else:
        working["9th_sector"] = ""
    if fuel_cols:
        fuel_values = pd.DataFrame(
            {col: _clean_label_series(working[col]) for col in fuel_cols}
        )
        working["9th_fuel"] = fuel_values.bfill(axis=1).iloc[:, 0].fillna("")
    else:
        working["9th_fuel"] = ""
    return working


def filter_ninth_projection_rows(
    df: pd.DataFrame, scenario: str = DEFAULT_SCENARIO
) -> pd.DataFrame:
    """Filter 9th data to the reference scenario and non-subtotal rows."""
    working = df.copy()
    if scenario and "scenarios" in working.columns:
        scenario_key = str(scenario).strip().lower()
        working = working[
            working["scenarios"].astype(str).str.strip().str.lower() == scenario_key
        ]
    if "subtotal_results" in working.columns:
        working = working[working["subtotal_results"] == False]
    return working


def build_ninth_projection_series(
    ninth_df: pd.DataFrame, projection_years: Sequence[int]
) -> pd.DataFrame:
    """Aggregate projected-year values by economy + 9th pair."""
    if not projection_years or ninth_df.empty:
        return pd.DataFrame()
    year_cols = [year for year in projection_years if year in ninth_df.columns]
    if not year_cols:
        return pd.DataFrame()
    working = ninth_df.copy()
    working = working[(working["9th_sector"] != "") & (working["9th_fuel"] != "")]
    if working.empty:
        return pd.DataFrame()
    for year in year_cols:
        working[year] = pd.to_numeric(working[year], errors="coerce").fillna(0.0)
    grouped = (
        working.groupby(["economy_key", "9th_sector", "9th_fuel"], dropna=False)[year_cols]
        .sum()
        .reset_index()
    )
    return grouped


def build_esto_base_year_values(
    esto_df: pd.DataFrame, base_year: int
) -> pd.DataFrame:
    """Return base-year values per economy/flow/product."""
    if esto_df.empty:
        return pd.DataFrame()
    year_col = base_year if base_year in esto_df.columns else str(base_year)
    if year_col not in esto_df.columns:
        return pd.DataFrame()
    working = esto_df.copy()
    working["economy_key"] = working["economy"].apply(normalize_economy_key)
    working["esto_flow"] = working["flows"].astype(str).str.strip()
    working["esto_product"] = working["products"].astype(str).str.strip()
    working[year_col] = pd.to_numeric(working[year_col], errors="coerce").fillna(0.0)
    grouped = (
        working.groupby(["economy_key", "esto_flow", "esto_product"], dropna=False)[
            year_col
        ]
        .sum()
        .reset_index()
        .rename(columns={year_col: "base_value"})
    )
    grouped["base_value_abs"] = grouped["base_value"].abs()
    return grouped


def allocate_ninth_projection_to_esto(
    mapping_df: pd.DataFrame,
    ninth_series: pd.DataFrame,
    base_values: pd.DataFrame,
    projection_years: Sequence[int],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Allocate 9th projections to ESTO pairs using base-year shares."""
    if mapping_df.empty or ninth_series.empty or not projection_years:
        return pd.DataFrame(), pd.DataFrame()
    mapping = mapping_df.copy()
    mapping["9th_sector"] = mapping["9th_sector"].fillna("").astype(str).str.strip()
    mapping["9th_fuel"] = mapping["9th_fuel"].fillna("").astype(str).str.strip()
    mapping["esto_flow"] = mapping["esto_flow"].fillna("").astype(str).str.strip()
    mapping["esto_product"] = mapping["esto_product"].fillna("").astype(str).str.strip()
    mapping = mapping[(mapping["9th_sector"] != "") & (mapping["9th_fuel"] != "")]
    if mapping.empty:
        return pd.DataFrame(), pd.DataFrame()

    base_values = base_values.copy()
    if not base_values.empty:
        base_values["esto_flow"] = base_values["esto_flow"].astype(str).str.strip()
        base_values["esto_product"] = base_values["esto_product"].astype(str).str.strip()
        base_values["economy_key"] = base_values["economy_key"].astype(str).str.strip()

    apec_base = (
        base_values.groupby(["esto_flow", "esto_product"], dropna=False)["base_value_abs"]
        .sum()
        .reset_index()
    )
    mapping_apec = mapping.merge(apec_base, on=["esto_flow", "esto_product"], how="left")
    mapping_apec["base_value_abs"] = mapping_apec["base_value_abs"].fillna(0.0)
    mapping_apec["apec_group_total"] = mapping_apec.groupby(
        ["9th_sector", "9th_fuel"], dropna=False
    )["base_value_abs"].transform("sum")
    mapping_apec["apec_share"] = 0.0
    apec_mask = mapping_apec["apec_group_total"] > 0
    mapping_apec.loc[apec_mask, "apec_share"] = (
        mapping_apec.loc[apec_mask, "base_value_abs"]
        / mapping_apec.loc[apec_mask, "apec_group_total"]
    )

    merged = mapping.merge(
        ninth_series, on=["9th_sector", "9th_fuel"], how="inner"
    )
    merged = merged.merge(
        base_values[["economy_key", "esto_flow", "esto_product", "base_value_abs"]],
        on=["economy_key", "esto_flow", "esto_product"],
        how="left",
    )
    merged["base_value_abs"] = merged["base_value_abs"].fillna(0.0)
    merged = merged.merge(
        mapping_apec[
            [
                "9th_sector",
                "9th_fuel",
                "esto_flow",
                "esto_product",
                "apec_group_total",
                "apec_share",
            ]
        ],
        on=["9th_sector", "9th_fuel", "esto_flow", "esto_product"],
        how="left",
    )
    merged["apec_group_total"] = merged["apec_group_total"].fillna(0.0)
    merged["apec_share"] = merged["apec_share"].fillna(0.0)
    merged["group_total"] = merged.groupby(
        ["economy_key", "9th_sector", "9th_fuel"], dropna=False
    )["base_value_abs"].transform("sum")
    merged["group_count"] = merged.groupby(
        ["9th_sector", "9th_fuel"], dropna=False
    )["esto_flow"].transform("count").astype(float)
    merged["share"] = 0.0
    merged["share_source"] = "economy"
    economy_mask = merged["group_total"] > 0
    merged.loc[economy_mask, "share"] = (
        merged.loc[economy_mask, "base_value_abs"]
        / merged.loc[economy_mask, "group_total"]
    )
    fallback_mask = ~economy_mask
    apec_mask = fallback_mask & (merged["apec_group_total"] > 0)
    merged.loc[apec_mask, "share"] = merged.loc[apec_mask, "apec_share"]
    merged.loc[apec_mask, "share_source"] = "apec"
    equal_mask = fallback_mask & ~apec_mask
    merged.loc[equal_mask, "share"] = (
        1.0
        / merged.loc[equal_mask, "group_count"].replace(0, pd.NA)
    )
    merged.loc[equal_mask, "share_source"] = "equal"
    merged["share"] = merged["share"].fillna(0.0)

    year_cols = [year for year in projection_years if year in merged.columns]
    for year in year_cols:
        merged[year] = pd.to_numeric(merged[year], errors="coerce").fillna(0.0)
    merged[year_cols] = merged[year_cols].multiply(merged["share"], axis=0)

    projection_df = (
        merged.groupby(["economy_key", "esto_flow", "esto_product"], dropna=False)[
            year_cols
        ]
        .sum()
        .reset_index()
    )
    diagnostics = merged.loc[
        merged["share_source"] != "economy",
        [
            "economy_key",
            "9th_sector",
            "9th_fuel",
            "esto_flow",
            "esto_product",
            "share_source",
            "group_total",
            "apec_group_total",
            "base_value_abs",
            "share",
        ],
    ].copy()
    return projection_df, diagnostics


def build_esto_projection_table(
    ninth_data: pd.DataFrame,
    esto_data: pd.DataFrame,
    mapping_path: str | Path,
    base_year: int,
    projection_years: Sequence[int],
    scenario: str = DEFAULT_SCENARIO,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return projected values for ESTO pairs plus allocation diagnostics."""
    mapping_path = Path(mapping_path)
    if not mapping_path.exists():
        return pd.DataFrame(), pd.DataFrame()
    if mapping_path.suffix.lower() in {".xlsx", ".xls"}:
        mapping_df = pd.read_excel(mapping_path, dtype=str).fillna("")
    else:
        mapping_df = pd.read_csv(mapping_path, dtype=str).fillna("")
    if mapping_df.empty:
        return pd.DataFrame(), pd.DataFrame()
    ninth_filtered = filter_ninth_projection_rows(ninth_data, scenario=scenario)
    ninth_pairs = add_ninth_pair_columns(ninth_filtered)
    ninth_pairs["economy_key"] = ninth_pairs["economy"].apply(normalize_economy_key)
    ninth_series = build_ninth_projection_series(ninth_pairs, projection_years)
    base_values = build_esto_base_year_values(esto_data, base_year)
    return allocate_ninth_projection_to_esto(
        mapping_df,
        ninth_series,
        base_values,
        projection_years,
    )


def merge_projection_into_esto(
    esto_df: pd.DataFrame,
    projection_df: pd.DataFrame,
    projection_years: Sequence[int],
) -> pd.DataFrame:
    """Return an ESTO dataframe with projection years appended."""
    if projection_df is None or projection_df.empty or not projection_years:
        return esto_df
    working = esto_df.copy()
    working["economy_key"] = working["economy"].apply(normalize_economy_key)
    working["flows"] = working["flows"].astype(str).str.strip()
    working["products"] = working["products"].astype(str).str.strip()

    proj = projection_df.copy()
    proj["esto_flow"] = proj["esto_flow"].astype(str).str.strip()
    proj["esto_product"] = proj["esto_product"].astype(str).str.strip()
    proj_cols = [year for year in projection_years if year in proj.columns]
    if not proj_cols:
        return esto_df
    proj = proj.rename(columns={year: f"{year}_proj" for year in proj_cols})

    merged = working.merge(
        proj,
        left_on=["economy_key", "flows", "products"],
        right_on=["economy_key", "esto_flow", "esto_product"],
        how="left",
    )
    for year in proj_cols:
        proj_col = f"{year}_proj"
        merged[year] = merged[proj_col].fillna(0.0)
    drop_cols = [
        "economy_key",
        "esto_flow",
        "esto_product",
    ] + [f"{year}_proj" for year in proj_cols]
    merged = merged.drop(columns=[col for col in drop_cols if col in merged.columns])

    base_cols = [col for col in esto_df.columns if col not in proj_cols]
    existing_years = [col for col in base_cols if str(col).isdigit()]
    non_year_cols = [col for col in base_cols if col not in existing_years]
    ordered_years = sorted(set(existing_years + proj_cols))
    ordered_cols = non_year_cols + ordered_years
    merged = merged[ordered_cols]
    return merged


def build_projection_lookup(projection_df: pd.DataFrame) -> pd.DataFrame | None:
    """Return a MultiIndex lookup for projection values."""
    if projection_df is None or projection_df.empty:
        return None
    grouped = (
        projection_df.groupby(
            ["economy_key", "esto_flow", "esto_product"], dropna=False
        )
        .sum(numeric_only=True)
    )
    return grouped
