#%%
# Build flattened 9th rows with most-specific sector/fuel columns and
# summarize subtotal economy lists by pair for base year and base year + 1.
# NOTE: We observed no projected-year (> base year) subtotal pattern differences
# across economies, so downstream mapping can focus on > base year data.
# Outputs are written to outputs/.
#%%
from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[3]
DATA_PATH = REPO_ROOT / "data" / "merged_file_energy_ALL_20250814.csv"
OUTPUT_DIR = REPO_ROOT / "outputs"

BASE_YEAR = 2022
PROJECTED_YEAR = BASE_YEAR + 1
SCENARIO_FILTER = "reference"

SECTOR_COLS = ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]
FUEL_COLS = ["fuels", "subfuels"]
META_COLS = ["economy", "scenarios", "subtotal_layout", "subtotal_results"]


def _find_year_columns(columns: Iterable[str]) -> list[str]:
    return [col for col in columns if str(col).isdigit()]


def _coerce_string_columns(df: pd.DataFrame, columns: Iterable[str]) -> None:
    for col in columns:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str)


def _most_specific_label(row: pd.Series, columns: list[str]) -> str:
    for col in columns:
        value = row.get(col, "")
        if value and value != "x":
            return value
    return ""


def _build_flattened_rows(df: pd.DataFrame) -> pd.DataFrame:
    sector_order = list(reversed(SECTOR_COLS))
    fuel_order = list(reversed(FUEL_COLS))
    df = df.copy()
    df["sector"] = df.apply(lambda row: _most_specific_label(row, sector_order), axis=1)
    df["fuel"] = df.apply(lambda row: _most_specific_label(row, fuel_order), axis=1)
    return df


def _flag_nonzero(series: pd.Series) -> pd.Series:
    values = pd.to_numeric(series, errors="coerce").fillna(0)
    return values != 0


def _summarize_subtotals_by_pair(
    df: pd.DataFrame,
    year_col: str,
    subtotal_col: str,
) -> pd.DataFrame:
    if year_col not in df.columns:
        raise ValueError(f"Missing year column: {year_col}")
    if subtotal_col not in df.columns:
        raise ValueError(f"Missing subtotal column: {subtotal_col}")

    working = df.copy()
    working = working[_flag_nonzero(working[year_col])]

    required = ["economy", "sector", "fuel", subtotal_col]
    missing = [col for col in required if col not in working.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    working[subtotal_col] = working[subtotal_col].fillna(False).astype(bool)

    grouped = (
        working.groupby(["sector", "fuel", "economy"])[subtotal_col]
        .agg(lambda values: bool(values.max()))
        .reset_index()
    )

    def _collect(series: pd.Series, flag_value: bool) -> list[str]:
        return sorted(series[grouped[subtotal_col] == flag_value].unique().tolist())

    results = []
    for (sector, fuel), group in grouped.groupby(["sector", "fuel"]):
        subtotal_econs = sorted(group.loc[group[subtotal_col], "economy"].unique().tolist())
        nonsub_econs = sorted(group.loc[~group[subtotal_col], "economy"].unique().tolist())
        results.append(
            {
                "sector": sector,
                "fuel": fuel,
                "subtotal_econs": ", ".join(subtotal_econs),
                "nonsubtotal_econs": ", ".join(nonsub_econs),
            }
        )

    return pd.DataFrame(results)


def _coerce_subtotal_flag(series: pd.Series) -> pd.Series:
    return series.fillna(False).astype(bool)


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    raw_df = pd.read_csv(DATA_PATH, low_memory=False)
    year_cols = _find_year_columns(raw_df.columns)

    keep_cols = [col for col in META_COLS + SECTOR_COLS + FUEL_COLS if col in raw_df.columns]
    keep_cols += year_cols
    df = raw_df[keep_cols].copy()

    _coerce_string_columns(df, SECTOR_COLS + FUEL_COLS + ["economy", "scenarios"])
    if "subtotal_results" in df.columns:
        df["subtotal_results"] = _coerce_subtotal_flag(df["subtotal_results"])
        df = df[~df["subtotal_results"]]
    if SCENARIO_FILTER:
        df = df[df["scenarios"].str.lower() == SCENARIO_FILTER.lower()]

    flattened = _build_flattened_rows(df)
    flattened_path = OUTPUT_DIR / "ninth_flattened_rows.csv"
    flattened.to_csv(flattened_path, index=False)
    print(f"Wrote flattened rows to {flattened_path}")

    base_year_col = str(BASE_YEAR)
    proj_year_col = str(PROJECTED_YEAR)

    base_summary = _summarize_subtotals_by_pair(
        flattened, base_year_col, "subtotal_layout"
    ).rename(
        columns={
            "subtotal_econs": "subtotal_econs_base",
            "nonsubtotal_econs": "nonsubtotal_econs_base",
        }
    )
    base_summary["subtotal_count_base"] = base_summary["subtotal_econs_base"].apply(
        lambda value: 0 if pd.isna(value) or not str(value).strip() else len(str(value).split(", "))
    )
    base_summary["nonsubtotal_count_base"] = base_summary["nonsubtotal_econs_base"].apply(
        lambda value: 0 if pd.isna(value) or not str(value).strip() else len(str(value).split(", "))
    )
    proj_summary = _summarize_subtotals_by_pair(
        flattened, proj_year_col, "subtotal_results"
    ).rename(
        columns={
            "subtotal_econs": "subtotal_econs_proj",
            "nonsubtotal_econs": "nonsubtotal_econs_proj",
        }
    )
    proj_summary["subtotal_count_proj"] = proj_summary["subtotal_econs_proj"].apply(
        lambda value: 0 if pd.isna(value) or not str(value).strip() else len(str(value).split(", "))
    )
    proj_summary["nonsubtotal_count_proj"] = proj_summary["nonsubtotal_econs_proj"].apply(
        lambda value: 0 if pd.isna(value) or not str(value).strip() else len(str(value).split(", "))
    )

    combined = pd.merge(base_summary, proj_summary, on=["sector", "fuel"], how="outer")
    combined_path = OUTPUT_DIR / "ninth_pair_subtotal_econ_report.csv"
    combined.to_csv(combined_path, index=False)
    print(f"Wrote subtotal pair report to {combined_path}")

    non_subtotals = flattened.copy()
    if "subtotal_results" in non_subtotals.columns:
        non_subtotals["subtotal_results"] = _coerce_subtotal_flag(non_subtotals["subtotal_results"])
        non_subtotals = non_subtotals[~non_subtotals["subtotal_results"]]

    projected_years = [
        col for col in year_cols if col.isdigit() and int(col) >= PROJECTED_YEAR
    ]
    if not projected_years:
        raise ValueError("No projected year columns found for 2023+ filtering.")

    projected_mask = (
        non_subtotals[projected_years]
        .apply(pd.to_numeric, errors="coerce")
        .fillna(0)
        .abs()
        .sum(axis=1)
        != 0
    )
    base_nonzero_pairs = (
        non_subtotals[projected_mask]
        .groupby(["sector", "fuel"])
        .size()
        .reset_index(name="row_count")
        .sort_values(["sector", "fuel"])
    )
    base_pair_path = OUTPUT_DIR / "ninth_pairs_base_year_nonzero.csv"
    base_nonzero_pairs.to_csv(base_pair_path, index=False)
    print(f"Wrote base-year nonzero pair list to {base_pair_path}")


if __name__ == "__main__":
    main()
