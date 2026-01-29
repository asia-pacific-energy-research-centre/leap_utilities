#%%
# Summary: Compute LEAP import parameters for transformation flows (LNG, gas works,
# blending, coal subtypes, charcoal, and nonspecified) using ESTO/9th datasets.
# How it works:
# - Loads ESTO/9th data, normalizes year columns, and cleans subtotals.
# - Uses explicit transformation flow codes per sector to select rows.
# - For each flow and economy, identifies primary input/output fuels and totals.
# - Computes efficiency as output / (feedstock + losses) using loss flow codes.
# - Treats own-use/loss fuels as auxiliary fuels unless they match feedstock.
# - Prints a LEAP-style structure block for manual import into LEAP.
import os
import re
import sys
from pathlib import Path

import pandas as pd

# Allow the repository root to be importable regardless of the working directory.
REPO_ROOT = Path(__file__).resolve().parents[1]
CURRENT_DIR = Path.cwd()
if CURRENT_DIR != REPO_ROOT:
    os.chdir(REPO_ROOT)
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from leap_utils.all_products_and_flows import ESTO_SECTORS
from leap_utils.config import (
    BRANCH_DEMAND_CATEGORY,
    BRANCH_DEMAND_TECHNOLOGY,
    BASE_YEAR,
)
from leap_utils.leap_core import (
    connect_to_leap,
    create_branches_from_export_file,
    fill_branches_from_export_file,
    sanitize_leap_branch_path,
)
from leap_utils.scrapbook.utilities import (
    apply_matt_subtotal_mapping,
    filter_matt_subtotals,
    save_subtotal_labeled_data,
)
from leap_utils.leap_excel_io import finalise_export_df, save_export_files
from leap_utils.ninth_projection_mapping import (
    build_esto_projection_table,
    merge_projection_into_esto,
)
#%%

#%%
######### CONSTANTS (UNLIKELY TO CHANGE) #########
ESTO_DATA_PATH = "data/merged_file_energy_ALL_20250814.csv"
MATT_DATA_PATH = "data/00APEC_2024_low.csv"
CONFIG_DIR = REPO_ROOT / "config"
SUBTOTAL_MAPPING_PATH = CONFIG_DIR / "ESTO_subtotal_mapping.xlsx"
NINTH_TO_ESTO_MAPPING_PATH = CONFIG_DIR / "ninth_pairs_to_esto_pairs.xlsx"
CODE_TO_NAME_PATHS = [
    CONFIG_DIR / "sector_fuel_codes_to_names.updated.xlsx",
    CONFIG_DIR / "sector_fuel_codes_to_names.xlsx",
]

YEAR_START_FOR_ANALYSIS = BASE_YEAR
PROJECTION_START_YEAR = 2023
PROJECTION_END_YEAR = 2061
PROJECTION_YEAR_RANGE = list(range(PROJECTION_START_YEAR, PROJECTION_END_YEAR + 1))
LOSS_SECTOR_CODE_9TH = "10_losses_and_own_use"
# DEFAULT_SCENARIO = "Target"
DEFAULT_REGION = "United States of America"
DEFAULT_OUTPUT_UNITS = "Petajoule"
DEFAULT_EFFICIENCY_UNITS = "Percent"
DEFAULT_FEEDSTOCK_UNITS = ""
DEFAULT_AUXILIARY_UNITS = "Gigajoule"
DEFAULT_AUXILIARY_PER = "Gigajoule"
ENABLE_DEBUG_BREAKPOINTS = True
PRINT_SECTOR_ROWS = True
PRINT_TOP_FUEL_ROWS = 12
PRINT_ONLY_NONZERO_ROWS = True
USE_CODE_TO_NAME_MAPPING = True
AUXILIARY_THRESHOLD_RATIO = 0.1
INCLUDE_ALL_AUXILIARY = False
PRINT_GAS_PROCESSING_SUMMARY = False

MAJOR_SECTOR_CONFIG = {
    "lng": {
        "dataset_key": "ninth",
        "title": "NG Liquefaction",
        "transformation_sub1": "09_06_gas_processing_plants",
        "transformation_sub2": ["09_06_02_liquefaction_regasification_plants"],
        "loss_sub2": ["10_01_03_liquefaction_regasification_plants"],
    },
    "gas_works": {
        "dataset_key": "esto",
        "title": "Gas works plants",
        "transformation_flow_codes": ["09.06.01 Gas works plants"],
        "loss_flow_codes": ["10.01.02 Gas works plants"],
    },
    "gas_blending": {
        "dataset_key": "esto",
        "title": "Natural gas blending plants",
        "transformation_flow_codes": ["09.06.03 Natural gas blending plants"],
        "loss_flow_codes": [],
    },
    "coal_coke_ovens": {
        "dataset_key": "esto",
        "title": "Coke ovens",
        "transformation_flow_codes": ["09.08.01 Coke ovens"],
        "loss_flow_codes": ["10.01.05 Coke ovens"],
    },
    "coal_blast_furnaces": {
        "dataset_key": "esto",
        "title": "Blast furnaces",
        "transformation_flow_codes": ["09.08.02 Blast furnaces"],
        "loss_flow_codes": ["10.01.07 Blast furnaces"],
    },
    "coal_patent_fuel_plants": {
        "dataset_key": "esto",
        "title": "Patent fuel plants",
        "transformation_flow_codes": ["09.08.03 Patent fuel plants"],
        "loss_flow_codes": [],
    },
    "coal_bkb_pb_plants": {
        "dataset_key": "esto",
        "title": "BKB/PB plants",
        "transformation_flow_codes": ["09.08.04 BKB/PB plants"],
        "loss_flow_codes": [],
    },
    "coal_liquefaction": {
        "dataset_key": "esto",
        "title": "Liquefaction (coal to oil)",
        "transformation_flow_codes": ["09.08.05 Liquefaction (coal to oil)"],
        "loss_flow_codes": [],
    },
    "coal_mines": {
        "dataset_key": "esto",
        "title": "Coal mines",
        "transformation_flow_codes": ["09.08.06 Coal mines"],
        "loss_flow_codes": ["10.01.06 Coal mines"],
    },
    "charcoal_processing": {
        "dataset_key": "esto",
        "title": "Charcoal processing",
        "transformation_flow_codes": ["09.11 Charcoal processing"],
        "loss_flow_codes": [],
    },
    "nonspecified_transformation": {
        "dataset_key": "esto",
        "title": "Non-specified transformation",
        "transformation_flow_codes": ["09.12 Non-specified transformation"],
        "loss_flow_codes": ["10.01.17 Non-specified own uses"],
    },
} 
# Additional constants for import/export lookups and output control
ESTO_IMPORT_EXPORT_REFERENCE_DATA = None
ESTO_IMPORT_EXPORT_YEAR_COLS = []
ESTO_IMPORT_SECTOR_LABEL = next(
    (sector for sector in ESTO_SECTORS if sector.startswith("02 ")),
    "02 Imports",
)
ESTO_EXPORT_SECTOR_LABEL = next(
    (sector for sector in ESTO_SECTORS if sector.startswith("03 ")),
    "03 Exports",
)
TRANSFORMATION_OUTPUT_VARIABLES = {
    "output": True,
    "output_import_target": True,
    "output_export_target": True,
    "feedstock_share": True,
    "process_efficiency": True,
    "auxiliary_ratio": True,
    "loss_value": True,
}
#%%
# Available sectors (with non zero data) in the ESTO dataset:
# 10.01 Own Use
# 10.01.01 Electricity, CHP and heat plants
# 10.01.02 Gas works plants
# 10.01.03 Liquefaction/regasification plants
# 10.01.05 Coke ovens
# 10.01.06 Coal mines
# 10.01.07 Blast furnaces
# 10.01.11 Oil refineries
# 10.01.12 Oil and gas extraction
# 10.01.13 Pump storage plants
# 10.01.17 Non-specified own uses
# 10.02 Transmission and distribution losses
#transformation sectors:
# 09.01 Main activity producer
# 09.01.01 Electricity plants
# 09.01.02 CHP plants
# 09.01.03 Heat plants
# 09.02 Autoproducers
# 09.02.01 Electricity plants
# 09.02.02 CHP plants
# 09.02.03 Heat plants
# 09.04 Electric boilers
# 09.05 Chemical heat for electricity production
# 09.06 Gas processing plants
# 09.06.01 Gas works plants
# 09.06.02 Liquefaction/regasification plants
# 09.06.03 Natural gas blending plants
# 09.06.04 Gas-to-liquids plants
# 09.07 Oil refineries
# 09.08 Coal transformation
# 09.08.01 Coke ovens
# 09.08.02 Blast furnaces
# 09.08.03 Patent fuel plants
# 09.08.04 BKB/PB plants
# 09.08.05 Liquefaction (coal to oil)
# 09.09 Petrochemical industry
# 09.11 Charcoal processing
# 09.12 Non-specified transformation
# Unused sectors (from provided 09/10 lists; keep for reference)
# UNUSED_SECTORS = [
#     "10.01 Own Use",
#     "10.01.01 Electricity, CHP and heat plants",
#     "10.01.11 Oil refineries",
#     "10.01.12 Oil and gas extraction",
#     "10.01.13 Pump storage plants",
#     "10.02 Transmission and distribution losses",
#     "09.01 Main activity producer",
#     "09.01.01 Electricity plants",
#     "09.01.02 CHP plants",
#     "09.01.03 Heat plants",
#     "09.02 Autoproducers",
#     "09.02.01 Electricity plants",
#     "09.02.02 CHP plants",
#     "09.02.03 Heat plants",
#     "09.04 Electric boilers",
#     "09.05 Chemical heat for electricity production",
#     "09.07 Oil refineries",
#     "09.09 Petrochemical industry",
# ]
# Build mapping helpers from MAJOR_SECTOR_CONFIG when needed.
#%%

#%%
######### FUNCTIONS #########
def ensure_repo_root():
    """Move to repo root if running from the scrapbook folder."""
    try:
        if os.getcwd().endswith("scrapbook"):
            os.chdir("../../")
    except Exception as exc:
        print(f"Failed to set repo root: {exc}")
        try_debug_breakpoint()
        raise


def try_debug_breakpoint():
    """Trigger a debug breakpoint when enabled (safe to call anywhere)."""
    if not ENABLE_DEBUG_BREAKPOINTS:
        return
    try:
        breakpoint()
    except Exception as breakpoint_exc:
        print(f"Debug breakpoint failed: {breakpoint_exc}")


def load_csv_data(path, label):
    """Load a CSV file and return a pandas DataFrame."""
    try:
        df = pd.read_csv(path)
        print(f"Loaded {label}: {df.shape[0]} rows, {df.shape[1]} columns")
        return df
    except Exception as exc:
        print(f"Failed to load {label} from {path}: {exc}")
        try_debug_breakpoint()
        raise


def filter_reference_scenario(df, label):
    """Filter to the Reference scenario when a scenarios column is present."""
    try:
        if "scenarios" not in df.columns:
            return df
        scenarios = df["scenarios"].astype(str).str.strip().str.lower()
        filtered = df[scenarios == "reference"].copy()
        unique_vals = sorted(set(scenarios.unique()))
        print(
            f"{label}: filtering to scenarios=reference "
            f"(available={unique_vals}, rows={filtered.shape[0]})"
        )
        return filtered
    except Exception as exc:
        print(f"Failed to filter scenarios for {label}: {exc}")
        try_debug_breakpoint()
        raise


def normalize_year_columns(df):
    """Convert year-like columns to int and return (df, year_cols)."""
    try:
        year_cols = [int(col) for col in df.columns if str(col).isdigit()]
        df.columns = [int(col) if str(col).isdigit() else col for col in df.columns]
        return df, year_cols
    except Exception as exc:
        print(f"Failed to normalize year columns: {exc}")
        try_debug_breakpoint()
        raise


def get_years_from(year_cols, start_year):
    """Return a list with the base year column when available."""
    try:
        return [year for year in year_cols if year >= start_year]
    except Exception as exc:
        print(f"Failed to filter year columns from {start_year}: {exc}")
        try_debug_breakpoint()
        raise


def _extract_numeric_segments(value):
    """Return a list of numeric code segments from a code or label."""
    try:
        if value is None:
            return []
        text = str(value)
        if text == "x":
            return []
        segments = []
        for chunk in text.replace(".", "_").split("_"):
            match = re.match(r"^(\d+)", chunk)
            if not match:
                break
            segments.append(match.group(1).zfill(2))
        return segments
    except Exception as exc:
        print(f"Failed to extract numeric segments from {value}: {exc}")
        try_debug_breakpoint()
        raise


def _match_code_prefix(label_value, code_value):
    """Check if label_value shares the numeric prefix of code_value."""
    try:
        code_segments = _extract_numeric_segments(code_value)
        if not code_segments:
            return False
        label_segments = _extract_numeric_segments(label_value)
        if len(label_segments) < len(code_segments):
            return False
        return label_segments[: len(code_segments)] == code_segments
    except Exception as exc:
        print(f"Failed to match code prefix for {code_value}: {exc}")
        try_debug_breakpoint()
        raise


def _normalize_economy_value(value):
    """Normalize economy codes to a common underscore-free form."""
    try:
        if value is None:
            return ""
        return str(value).replace("_", "").strip()
    except Exception as exc:
        print(f"Failed to normalize economy value {value}: {exc}")
        try_debug_breakpoint()
        raise


def select_rows(df, filters):
    """Return filtered rows based on a dict of column -> value."""
    try:
        mask = pd.Series(True, index=df.index)
        for column, value in filters.items():
            if column in df.columns:
                if column == "economy":
                    target = _normalize_economy_value(value)
                    mask &= df[column].apply(_normalize_economy_value).eq(target)
                else:
                    mask &= df[column].eq(value)
                continue

            if column in ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]:
                if "flows" in df.columns:
                    mask &= df["flows"].apply(lambda flow: _match_code_prefix(flow, value))
                    continue
            if column in ["fuels", "subfuels"] and "products" in df.columns:
                mask &= df["products"].apply(lambda product: _match_code_prefix(product, value))
                continue

            mask &= False
        return df.loc[mask]
    except Exception as exc:
        print(f"Failed to filter rows with {filters}: {exc}")
        try_debug_breakpoint()
        raise


def sum_years(df, year_cols):
    """Sum values over year columns, returning a float."""
    try:
        if df.empty:
            return 0.0
        return df[year_cols].sum().sum()
    except Exception as exc:
        print(f"Failed to sum years for frame: {exc}")
        try_debug_breakpoint()
        raise


def clean_esto_subtotals(df, year_cols):
    """Remove subtotal rows for pre/post 2022 and return a cleaned dataset."""
    try:
        required_cols = ["subtotal_2022_and_before", "subtotal_2023_and_after"]
        if not all(col in df.columns for col in required_cols):
            print("Subtotal flags missing; skipping ESTO subtotal cleanup.")
            return df
        pre_years = [col for col in year_cols if col <= 2022]
        post_years = [col for col in year_cols if col >= 2023]

        pre_non_subtotal = df[df["subtotal_2022_and_before"] == False].copy()
        post_non_subtotal = df[df["subtotal_2023_and_after"] == False].copy()

        pre_non_subtotal[post_years] = 0
        post_non_subtotal[pre_years] = 0

        key_cols = [
            col
            for col in df.columns
            if col not in year_cols
            and col not in ["subtotal_2022_and_before", "subtotal_2023_and_after"]
        ]

        combined = (
            pd.concat([pre_non_subtotal, post_non_subtotal], ignore_index=True)
            .groupby(key_cols, dropna=False)[year_cols]
            .sum()
            .reset_index()
        )
        combined["subtotal_2022_and_before"] = False
        combined["subtotal_2023_and_after"] = False
        return combined
    except Exception as exc:
        print(f"Failed to clean ESTO subtotals: {exc}")
        try_debug_breakpoint()
        raise


def normalize_esto_economy_codes(df):
    """Insert underscore in ESTO economy codes (e.g., 01AUS -> 01_AUS)."""
    try:
        if "economy" not in df.columns:
            return df
        updated = df.copy()
        updated["economy"] = (
            updated["economy"]
            .astype(str)
            .str.replace(r"^(\d{2})([A-Z].+)$", r"\1_\2", regex=True)
        )
        return updated
    except Exception as exc:
        print(f"Failed to normalize ESTO economy codes: {exc}")
        try_debug_breakpoint()
        raise


def filter_total_energy_rows(df):
    """Drop total/renewables summary rows from fuels or products."""
    try:
        updated = df.copy()
        total_codes = {"19_total", "20_total_renewables", "21_modern_renewables"}
        total_labels = {
            "19 Total",
            "20 Total Renewables",
            "21 Modern renewables",
        }
        if "fuels" in updated.columns:
            updated = updated[~updated["fuels"].astype(str).isin(total_codes)]
        if "subfuels" in updated.columns:
            updated = updated[~updated["subfuels"].astype(str).isin(total_codes)]
        if "products" in updated.columns:
            updated = updated[~updated["products"].astype(str).isin(total_labels)]
        return updated
    except Exception as exc:
        print(f"Failed to filter total energy rows: {exc}")
        try_debug_breakpoint()
        raise


def add_all_economy_total(df, year_cols, economy_label="ALL"):
    """Append an all-economy total row set to a dataset."""
    try:
        if "economy" not in df.columns or df.empty:
            return df
        if df["economy"].astype(str).eq(economy_label).any():
            return df
        group_cols = [
            col for col in df.columns if col not in year_cols and col != "economy"
        ]
        totals = (
            df.groupby(group_cols, dropna=False)[year_cols]
            .sum()
            .reset_index()
        )
        totals["economy"] = economy_label
        totals = totals[df.columns.tolist()]
        return pd.concat([df, totals], ignore_index=True)
    except Exception as exc:
        print(f"Failed to add all-economy totals: {exc}")
        try_debug_breakpoint()
        raise


def get_economy_list(df, requested_economies=None):
    """Return a list of economies to analyze."""
    try:
        available = sorted(df["economy"].dropna().unique())
        if requested_economies:
            requested = [econ for econ in requested_economies if econ in available]
            missing = [econ for econ in requested_economies if econ not in available]
            if missing:
                print(
                    "Warning: requested economies not found in this dataset: "
                    f"{', '.join(missing)}"
                )
            if requested:
                return requested
            print("Warning: no requested economies found; using all available economies.")
        return available
    except Exception as exc:
        print(f"Failed to build economy list: {exc}")
        try_debug_breakpoint()
        raise


def print_sector_rows(df, sector_label, filters, year_cols, start_year, code_to_name_mapping=None):
    """Print rows for a sector so the user can manually inspect inputs/outputs."""
    try:
        if not PRINT_SECTOR_ROWS:
            return
        sector_rows = select_rows(df, filters)
        if sector_rows.empty:
            print(f"\n{sector_label}: no rows found")
            return
        year_cols_from_start = get_years_from(year_cols, start_year)
        summary = sector_rows.copy()
        summary["total_from_start"] = summary[year_cols_from_start].sum(axis=1)
        if PRINT_ONLY_NONZERO_ROWS:
            summary = summary[summary["total_from_start"] != 0]
        if summary.empty:
            print(f"\n{sector_label}: no nonzero rows after filtering")
            return
        columns_to_show = [
            "scenarios",
            "economy",
            "sectors",
            "sub1sectors",
            "sub2sectors",
            "sub3sectors",
            "sub4sectors",
            "fuels",
            "subfuels",
            "flows",
            "products",
            "total_from_start",
        ]
        columns_to_show = [col for col in columns_to_show if col in summary.columns]
        print(f"\n{sector_label}: rows {summary.shape[0]}")
        summary_to_show = summary[columns_to_show].copy()
        if code_to_name_mapping:
            columns_to_map = [
                "sectors",
                "sub1sectors",
                "sub2sectors",
                "sub3sectors",
                "sub4sectors",
                "fuels",
                "subfuels",
            ]
            for column in columns_to_map:
                if column in summary_to_show.columns:
                    summary_to_show[column] = summary_to_show[column].apply(
                        lambda value: map_code_label(value, code_to_name_mapping)
                    )
        print(summary_to_show.head(PRINT_TOP_FUEL_ROWS).to_string(index=False))
    except Exception as exc:
        print(f"Failed to print sector rows for {sector_label}: {exc}")
        try_debug_breakpoint()
        raise


def print_sector_rows_from_df(
    sector_rows, sector_label, year_cols, start_year, code_to_name_mapping=None
):
    """Print already-filtered rows for a sector."""
    try:
        if not PRINT_SECTOR_ROWS:
            return
        if sector_rows.empty:
            print(f"\n{sector_label}: no rows found")
            return
        year_cols_from_start = get_years_from(year_cols, start_year)
        summary = sector_rows.copy()
        summary["total_from_start"] = summary[year_cols_from_start].sum(axis=1)
        if PRINT_ONLY_NONZERO_ROWS:
            summary = summary[summary["total_from_start"] != 0]
        if summary.empty:
            print(f"\n{sector_label}: no nonzero rows after filtering")
            return
        columns_to_show = [
            "scenarios",
            "economy",
            "sectors",
            "sub1sectors",
            "sub2sectors",
            "sub3sectors",
            "sub4sectors",
            "fuels",
            "subfuels",
            "flows",
            "products",
            "total_from_start",
        ]
        columns_to_show = [col for col in columns_to_show if col in summary.columns]
        print(f"\n{sector_label}: rows {summary.shape[0]}")
        summary_to_show = summary[columns_to_show].copy()
        if code_to_name_mapping:
            columns_to_map = [
                "sectors",
                "sub1sectors",
                "sub2sectors",
                "sub3sectors",
                "sub4sectors",
                "fuels",
                "subfuels",
            ]
            for column in columns_to_map:
                if column in summary_to_show.columns:
                    summary_to_show[column] = summary_to_show[column].apply(
                        lambda value: map_code_label(value, code_to_name_mapping)
                    )
        print(summary_to_show.head(PRINT_TOP_FUEL_ROWS).to_string(index=False))
    except Exception as exc:
        print(f"Failed to print sector rows for {sector_label}: {exc}")
        try_debug_breakpoint()
        raise


def summarize_fuels_by_subfuel(df, year_cols, start_year):
    """Summarize inputs (negative) and outputs (positive) per fuel label."""
    try:
        totals, _ = summarize_fuel_totals(df, year_cols, start_year, allow_all_years_fallback=False)
        negatives = totals[totals < 0].sort_values()
        positives = totals[totals > 0].sort_values(ascending=False)
        return negatives, positives
    except Exception as exc:
        print(f"Failed to summarize fuels by subfuel: {exc}")
        try_debug_breakpoint()
        raise


def get_fuel_labels(df):
    """Return a series of fuel labels for grouping."""
    try:
        if "subfuels" in df.columns and "fuels" in df.columns:
            return df["subfuels"].where(df["subfuels"] != "x", df["fuels"])
        if "products" in df.columns:
            return df["products"]
        return None
    except Exception as exc:
        print(f"Failed to get fuel labels: {exc}")
        try_debug_breakpoint()
        raise


def summarize_fuel_totals(df, year_cols, start_year, allow_all_years_fallback=True):
    """Return totals by fuel label and whether all-years fallback was used."""
    try:
        fuel_labels = get_fuel_labels(df)
        if fuel_labels is None:
            return pd.Series(dtype=float), False
        year_cols_from_start = get_years_from(year_cols, start_year)
        totals = (
            df.assign(fuel_label=fuel_labels)
            .groupby("fuel_label")[year_cols_from_start]
            .sum()
            .sum(axis=1)
        )
        if allow_all_years_fallback and (totals[totals < 0].empty or totals[totals > 0].empty):
            totals = (
                df.assign(fuel_label=fuel_labels)
                .groupby("fuel_label")[year_cols]
                .sum()
                .sum(axis=1)
            )
            return totals.sort_values(), True
        return totals.sort_values(), False
    except Exception as exc:
        print(f"Failed to summarize fuel totals: {exc}")
        try_debug_breakpoint()
        raise


def summarize_fuel_timeseries(df, year_cols, start_year, allow_all_years_fallback=True):
    """Return (timeseries_df, used_all_years) grouped by fuel label and year."""
    try:
        fuel_labels = get_fuel_labels(df)
        if fuel_labels is None:
            return pd.DataFrame(), False
        year_cols_from_start = get_years_from(year_cols, start_year)
        timeseries = (
            df.assign(fuel_label=fuel_labels)
            .groupby("fuel_label")[year_cols_from_start]
            .sum()
        )
        totals = timeseries.sum(axis=1)
        if allow_all_years_fallback and (totals[totals < 0].empty or totals[totals > 0].empty):
            timeseries = (
                df.assign(fuel_label=fuel_labels)
                .groupby("fuel_label")[year_cols]
                .sum()
            )
            return timeseries, True
        return timeseries, False
    except Exception as exc:
        print(f"Failed to summarize fuel timeseries: {exc}")
        try_debug_breakpoint()
        raise


def get_label_timeseries(timeseries_df, label):
    """Return a series for a label, matching on code prefix when needed."""
    try:
        if timeseries_df is None or timeseries_df.empty:
            return pd.Series(dtype=float)
        if label in timeseries_df.index:
            return timeseries_df.loc[label]
        matches = [
            idx for idx in timeseries_df.index if _match_code_prefix(idx, label)
        ]
        if matches:
            return timeseries_df.loc[matches[0]]
        return pd.Series(dtype=float)
    except Exception as exc:
        print(f"Failed to get label timeseries for {label}: {exc}")
        try_debug_breakpoint()
        raise


def sum_years_by_year(df, year_cols, start_year):
    """Return a Series of year -> total for selected years."""
    try:
        year_cols_from_start = get_years_from(year_cols, start_year)
        if not year_cols_from_start:
            return pd.Series(dtype=float)
        return df[year_cols_from_start].sum()
    except Exception as exc:
        print(f"Failed to sum years by year: {exc}")
        try_debug_breakpoint()
        raise


def ensure_full_year_series(series, base_year, final_year):
    """Return a Series indexed by the full year range, filling missing with 0."""
    try:
        full_years = list(range(base_year, final_year + 1))
        if series is None or series.empty:
            return pd.Series({year: 0.0 for year in full_years})
        return series.reindex(full_years, fill_value=0.0)
    except Exception as exc:
        print(f"Failed to ensure full year series: {exc}")
        try_debug_breakpoint()
        raise


def series_to_year_dict(series, base_year, final_year):
    """Return a dict of year -> value for the full year range."""
    try:
        full_series = ensure_full_year_series(series, base_year, final_year)
        return full_series.to_dict()
    except Exception as exc:
        print(f"Failed to convert series to year dict: {exc}")
        try_debug_breakpoint()
        raise


def _match_est_product_label(product_value, code_value):
    """Check whether a product label shares the prefix of a target fuel code."""
    try:
        if product_value is None or (isinstance(product_value, float) and pd.isna(product_value)):
            return False
        return _match_code_prefix(str(product_value), code_value)
    except Exception as exc:
        print(f"Failed to match ESTO product label {product_value} to {code_value}: {exc}")
        try_debug_breakpoint()
        raise


def _filter_est_import_export_rows(economy, fuel_label, sector_label):
    """Return ESTO rows for a flow sector and fuel that match the economy."""
    try:
        if ESTO_IMPORT_EXPORT_REFERENCE_DATA is None or ESTO_IMPORT_EXPORT_REFERENCE_DATA.empty:
            return pd.DataFrame()
        df = ESTO_IMPORT_EXPORT_REFERENCE_DATA
        mask = pd.Series(True, index=df.index)
        if "flows" in df.columns:
            mask &= df["flows"].eq(sector_label)
        else:
            mask &= False
        if "economy" in df.columns:
            target = _normalize_economy_value(economy)
            mask &= df["economy"].apply(_normalize_economy_value).eq(target)
        else:
            mask &= False
        if "products" in df.columns:
            mask &= df["products"].apply(
                lambda value: _match_est_product_label(value, fuel_label)
            )
        else:
            mask &= False
        if not mask.any():
            return df.iloc[0:0]
        return df.loc[mask]
    except Exception as exc:
        print(f"Failed to filter import/export rows for {fuel_label}: {exc}")
        try_debug_breakpoint()
        raise


def build_est_output_target_dict(
    economy,
    fuel_label,
    sector_label,
    start_year,
    base_year,
    final_year,
):
    """Build a per-year dictionary of import/export totals for a fuel."""
    try:
        if not ESTO_IMPORT_EXPORT_YEAR_COLS:
            return {}
        rows = _filter_est_import_export_rows(economy, fuel_label, sector_label)
        if rows.empty:
            return {}
        series = sum_years_by_year(rows, ESTO_IMPORT_EXPORT_YEAR_COLS, start_year)
        series = series.abs()
        if series.sum() == 0:
            return {}
        full_series = ensure_full_year_series(series, base_year, final_year)
        return series_to_year_dict(full_series, base_year, final_year)
    except Exception as exc:
        print(f"Failed to build import/export target dict for {fuel_label}: {exc}")
        try_debug_breakpoint()
        raise


def gather_output_target_dicts(economy, output_labels, base_year, final_year):
    """Return dictionaries for import/export targets across output fuels."""
    try:
        import_targets = {}
        export_targets = {}
        if not output_labels:
            return import_targets, export_targets
        if (
            ESTO_IMPORT_EXPORT_REFERENCE_DATA is None
            or ESTO_IMPORT_EXPORT_REFERENCE_DATA.empty
            or not ESTO_IMPORT_EXPORT_YEAR_COLS
        ):
            return import_targets, export_targets
        for label in output_labels:
            if TRANSFORMATION_OUTPUT_VARIABLES.get("output_import_target"):
                import_dict = build_est_output_target_dict(
                    economy,
                    label,
                    ESTO_IMPORT_SECTOR_LABEL,
                    YEAR_START_FOR_ANALYSIS,
                    base_year,
                    final_year,
                )
                if import_dict:
                    import_targets[label] = import_dict
            if TRANSFORMATION_OUTPUT_VARIABLES.get("output_export_target"):
                export_dict = build_est_output_target_dict(
                    economy,
                    label,
                    ESTO_EXPORT_SECTOR_LABEL,
                    YEAR_START_FOR_ANALYSIS,
                    base_year,
                    final_year,
                )
                if export_dict:
                    export_targets[label] = export_dict
        return import_targets, export_targets
    except Exception as exc:
        print(f"Failed to gather output target dicts for {output_labels}: {exc}")
        try_debug_breakpoint()
        raise


def safe_divide_series(numerator, denominator):
    """Return numerator/denominator with zeros where denominator is 0."""
    try:
        if numerator is None or denominator is None:
            return pd.Series(dtype=float)
        aligned = numerator.align(denominator, fill_value=0.0)
        num, denom = aligned
        result = num.copy()
        result[denom == 0] = 0.0
        result[denom != 0] = num[denom != 0] / denom[denom != 0]
        return result
    except Exception as exc:
        print(f"Failed to divide series safely: {exc}")
        try_debug_breakpoint()
        raise


def build_auxiliary_ratios_by_year(negative_timeseries, auxiliary_fuels, output_series):
    """Return auxiliary fuel ratios by year for each auxiliary fuel."""
    try:
        ratios = {}
        if negative_timeseries is None or output_series is None:
            return ratios
        for label in auxiliary_fuels:
            if label not in negative_timeseries.index:
                continue
            ratio_series = safe_divide_series(
                negative_timeseries.loc[label].abs(),
                output_series,
            )
            ratios[label] = ratio_series.to_dict()
        return ratios
    except Exception as exc:
        print(f"Failed to build auxiliary ratios by year: {exc}")
        try_debug_breakpoint()
        raise


def build_auxiliary_from_losses_by_year(loss_values_by_year, output_series):
    """Return auxiliary fuels and ratios by year derived from losses."""
    try:
        if not loss_values_by_year:
            return [], {}
        fuels = []
        ratios = {}
        for label, series in loss_values_by_year.items():
            fuels.append(label)
            ratio_series = safe_divide_series(pd.Series(series).abs(), output_series)
            ratios[label] = ratio_series.to_dict()
        return fuels, ratios
    except Exception as exc:
        print(f"Failed to build auxiliary fuels from losses by year: {exc}")
        try_debug_breakpoint()
        raise


def merge_loss_into_auxiliary_by_year(
    auxiliary_fuels, auxiliary_ratios, loss_values_by_year, output_series, feedstock_label
):
    """Treat own use/loss fuels as auxiliary by year (unless same as feedstock)."""
    try:
        if not loss_values_by_year or output_series is None or output_series.empty:
            return auxiliary_fuels, auxiliary_ratios
        updated_fuels = list(auxiliary_fuels) if auxiliary_fuels else []
        updated_ratios = dict(auxiliary_ratios) if auxiliary_ratios else {}
        for label, series in loss_values_by_year.items():
            if label == feedstock_label:
                continue
            if label not in updated_fuels:
                updated_fuels.append(label)
            ratio_series = safe_divide_series(pd.Series(series).abs(), output_series)
            updated_ratios[label] = ratio_series.to_dict()
        return updated_fuels, updated_ratios
    except Exception as exc:
        print(f"Failed to merge loss fuels into auxiliary list by year: {exc}")
        try_debug_breakpoint()
        raise


def filter_loss_values_for_feedstock_by_year(loss_values_by_year, feedstock_label):
    """Return loss values by year for the feedstock fuel only."""
    try:
        if not loss_values_by_year or not feedstock_label:
            return {}
        if feedstock_label not in loss_values_by_year:
            return {}
        return {feedstock_label: loss_values_by_year[feedstock_label]}
    except Exception as exc:
        print(f"Failed to filter loss values for feedstock by year: {exc}")
        try_debug_breakpoint()
        raise


def get_loss_total_for_efficiency_by_year(loss_values_by_year, feedstock_label, output_label, years):
    """Return year->loss total using feedstock/output labels only."""
    try:
        if not loss_values_by_year:
            return pd.Series({year: 0.0 for year in years})
        relevant_labels = {feedstock_label, output_label}
        totals = {year: 0.0 for year in years}
        for label in relevant_labels:
            series = loss_values_by_year.get(label)
            if not series:
                continue
            for year, value in series.items():
                totals[int(year)] = totals.get(int(year), 0.0) + abs(value)
        return pd.Series(totals)
    except Exception as exc:
        print(f"Failed to build loss total for efficiency by year: {exc}")
        try_debug_breakpoint()
        raise


def compute_efficiency_by_year(output_series, input_series, loss_series):
    """Return efficiency by year: output / (input + losses)."""
    try:
        denom = input_series.add(loss_series, fill_value=0.0)
        return safe_divide_series(output_series, denom)
    except Exception as exc:
        print(f"Failed to compute efficiency by year: {exc}")
        try_debug_breakpoint()
        raise

def compute_primary_io(negative_series, positive_series):
    """Return primary input/output labels and totals.

    Expects `negative_series` to contain the feedstock or own-use rows (negative balances)
    and `positive_series` to hold the corresponding outputs. The returned input total is
    always reported as a positive volume for LEAP.
    """
    try:
        primary_input = negative_series.idxmin()
        primary_output = positive_series.idxmax()
        input_total = abs(negative_series.loc[primary_input])
        output_total = positive_series.loc[primary_output]
        return primary_input, primary_output, input_total, output_total
    except Exception as exc:
        print(f"Failed to compute primary input/output: {exc}")
        try_debug_breakpoint()
        raise


def calculate_efficiency_with_losses(output_total, input_total, loss_total):
    """Return efficiency including losses (output / (input + losses))."""
    try:
        denominator = input_total + loss_total
        if denominator == 0:
            return 0.0
        return output_total / denominator
    except Exception as exc:
        print(f"Failed to calculate efficiency with losses: {exc}")
        try_debug_breakpoint()
        raise


def build_loss_context(
    loss_data,
    loss_year_cols,
    start_year,
    economy,
    sector_key,
    sub2_code=None,
    flow_code=None,
):
    """Return loss series, total, and value dict for a transformation code."""
    try:
        sector_config = MAJOR_SECTOR_CONFIG.get(sector_key, {})
        loss_sub2_map = sector_config.get("loss_sub2_map", {})
        loss_sub2_list = sector_config.get("loss_sub2", [])
        loss_flow_list = sector_config.get("loss_flow_codes", [])
        loss_sub2_code = None
        loss_flow_code = None
        if sub2_code and sub2_code in loss_sub2_map:
            loss_sub2_code = loss_sub2_map[sub2_code]
        elif loss_sub2_list:
            loss_sub2_code = loss_sub2_list[0]
        if loss_flow_list:
            loss_flow_code = loss_flow_list[0]
        loss_series, loss_total, loss_values_by_year = summarize_own_use_losses_by_year(
            loss_data,
            loss_year_cols,
            start_year,
            economy,
            loss_sub2_code,
            loss_flow_code,
            allow_all_years_fallback=True,
        )
        loss_values = {label: abs(value) for label, value in loss_series.items()}
        return loss_series, loss_total, loss_values, loss_values_by_year
    except Exception as exc:
        print(f"Failed to build loss context: {exc}")
        try_debug_breakpoint()
        raise


def summarize_own_use_losses(
    data,
    year_cols,
    start_year,
    economy,
    loss_sub2_code=None,
    loss_flow_code=None,
    allow_all_years_fallback=True,
):
    """Return (loss_series, loss_total) for own use/losses tied to a code.

    Source rows are negative because they belong to own-use/loss sectors, so the returned
    totals and loss_series entries are always converted to absolute (positive) values before
    reaching LEAP.
    """
    try:
        year_cols_from_start = get_years_from(year_cols, start_year)
        if loss_sub2_code:
            loss_rows = select_rows(
                data,
                {
                    "economy": economy,
                    "sectors": LOSS_SECTOR_CODE_9TH,
                    "sub2sectors": loss_sub2_code,
                },
            )
        elif loss_flow_code and "flows" in data.columns:
            loss_rows = select_rows(
                data,
                {
                    "economy": economy,
                    "flows": loss_flow_code,
                },
            )
        else:
            return pd.Series(dtype=float), 0.0
        if loss_rows.empty:
            return pd.Series(dtype=float), 0.0
        fuel_labels = get_fuel_labels(loss_rows)
        if fuel_labels is None:
            return pd.Series(dtype=float), 0.0
        totals = (
            loss_rows.assign(fuel_label=fuel_labels)
            .groupby("fuel_label")[year_cols_from_start]
            .sum()
            .sum(axis=1)
        )
        loss_series = totals[totals != 0].sort_values()
        if loss_series.empty and allow_all_years_fallback and year_cols:
            totals = (
                loss_rows.assign(fuel_label=fuel_labels)
                .groupby("fuel_label")[year_cols]
                .sum()
                .sum(axis=1)
            )
            loss_series = totals[totals != 0].sort_values()
        loss_total = loss_series.abs().sum()
        return loss_series, loss_total
    except Exception as exc:
        print(f"Failed to summarize own use losses: {exc}")
        try_debug_breakpoint()
        raise


def summarize_own_use_losses_by_year(
    data,
    year_cols,
    start_year,
    economy,
    loss_sub2_code=None,
    loss_flow_code=None,
    allow_all_years_fallback=True,
):
    """Return (loss_series_totals, loss_total, loss_values_by_year) for own use/losses."""
    try:
        year_cols_from_start = get_years_from(year_cols, start_year)
        if loss_sub2_code:
            loss_rows = select_rows(
                data,
                {
                    "economy": economy,
                    "sectors": LOSS_SECTOR_CODE_9TH,
                    "sub2sectors": loss_sub2_code,
                },
            )
        elif loss_flow_code and "flows" in data.columns:
            loss_rows = select_rows(
                data,
                {
                    "economy": economy,
                    "flows": loss_flow_code,
                },
            )
        else:
            return pd.Series(dtype=float), 0.0, {}
        if loss_rows.empty:
            return pd.Series(dtype=float), 0.0, {}
        fuel_labels = get_fuel_labels(loss_rows)
        if fuel_labels is None:
            return pd.Series(dtype=float), 0.0, {}
        timeseries = (
            loss_rows.assign(fuel_label=fuel_labels)
            .groupby("fuel_label")[year_cols_from_start]
            .sum()
        )
        if timeseries.empty and allow_all_years_fallback and year_cols:
            timeseries = (
                loss_rows.assign(fuel_label=fuel_labels)
                .groupby("fuel_label")[year_cols]
                .sum()
            )
        if timeseries.empty:
            return pd.Series(dtype=float), 0.0, {}
        timeseries = timeseries.loc[(timeseries != 0).any(axis=1)]
        loss_series = timeseries.sum(axis=1).sort_values()
        loss_total_by_year = timeseries.abs().sum(axis=0)
        loss_total = loss_total_by_year.sum()
        loss_values_by_year = {
            label: timeseries.loc[label].abs().to_dict()
            for label in timeseries.index
        }
        return loss_series, loss_total, loss_values_by_year
    except Exception as exc:
        print(f"Failed to summarize own use losses by year: {exc}")
        try_debug_breakpoint()
        raise


def get_flow_list(data, flow_codes=None):
    """Return a list of flow codes using explicit list."""
    try:
        if flow_codes:
            return list(flow_codes)
        return []
    except Exception as exc:
        print(f"Failed to build flow list: {exc}")
        try_debug_breakpoint()
        raise


def select_flow_rows(data, economy, flow_code):
    """Select rows for a single flow code."""
    try:
        if not flow_code or "flows" not in data.columns:
            return data.iloc[0:0]
        return select_rows(data, {"economy": economy, "flows": flow_code})
    except Exception as exc:
        print(f"Failed to select flow rows for {flow_code}: {exc}")
        try_debug_breakpoint()
        raise


def summarize_loss_sectors(
    data,
    year_cols,
    start_year,
    economy,
    loss_sub2_codes,
    code_to_name_mapping,
    title_prefix,
):
    """Print summaries for loss/own-use sectors."""
    try:
        print(f"\n==== {title_prefix} ({economy}) ====")
        if not has_required_columns(
            data,
            [["sub2sectors", "subfuels", "fuels"], ["flows", "products"]],
            title_prefix,
        ):
            return
        for loss_sub2 in loss_sub2_codes:
            label = map_code_label(loss_sub2, code_to_name_mapping)
            print_sector_rows(
                data,
                f"{title_prefix} rows ({label})",
                {
                    "economy": economy,
                    "sectors": LOSS_SECTOR_CODE_9TH,
                    "sub2sectors": loss_sub2,
                },
                year_cols,
                start_year,
                code_to_name_mapping,
            )
            loss_rows = select_rows(
                data,
                {
                    "economy": economy,
                    "sectors": LOSS_SECTOR_CODE_9TH,
                    "sub2sectors": loss_sub2,
                },
            )
            if loss_rows.empty:
                continue
            negatives, positives = summarize_fuels_by_subfuel(
                loss_rows, year_cols, start_year
            )
            if not negatives.empty:
                print("Loss inputs by fuel label:")
                print(map_series_index(negatives, code_to_name_mapping).to_string())
            if not positives.empty:
                print("Loss outputs by fuel label:")
                print(map_series_index(positives, code_to_name_mapping).to_string())
    except Exception as exc:
        print(f"Failed to summarize loss sectors: {exc}")
        try_debug_breakpoint()
        raise


def build_dataset_map(esto_data, esto_year_cols, ninth_data, ninth_year_cols, matt_data, matt_year_cols):
    """Return a dataset map keyed by dataset_key."""
    try:
        return {
            "esto": (esto_data, esto_year_cols),
            "ninth": (ninth_data, ninth_year_cols),
            "matt": (matt_data, matt_year_cols),
        }
    except Exception as exc:
        print(f"Failed to build dataset map: {exc}")
        try_debug_breakpoint()
        raise


def resolve_dataset(dataset_map, dataset_key):
    """Return (data, year_cols) for a dataset key."""
    try:
        if dataset_key not in dataset_map:
            raise KeyError(f"Unknown dataset key: {dataset_key}")
        return dataset_map[dataset_key]
    except Exception as exc:
        print(f"Failed to resolve dataset key {dataset_key}: {exc}")
        try_debug_breakpoint()
        raise


def load_code_to_name_mapping(path_candidates):
    """Load the code-to-name mapping from the first available workbook."""
    try:
        for path in path_candidates:
            if not os.path.exists(path):
                continue
            try:
                mapping_df = pd.read_excel(path, sheet_name="code_to_name")
            except Exception as exc:
                print(f"Failed to read code-to-name mapping from {path}: {exc}")
                continue
            required_cols = {"esto_label", "9th_label", "name"}
            if not required_cols.issubset(set(mapping_df.columns)):
                missing = sorted(required_cols - set(mapping_df.columns))
                print(
                    f"Missing {missing} columns in {path}; trying next file."
                )
                continue

            mapping = {}
            for _, row in mapping_df.iterrows():
                name = row.get("name")
                if name is None or (isinstance(name, float) and pd.isna(name)):
                    continue
                name = str(name).strip()
                if not name:
                    continue

                for col in ["esto_label", "9th_label"]:
                    label = row.get(col)
                    if label is None or (isinstance(label, float) and pd.isna(label)):
                        continue
                    label = str(label).strip()
                    if not label:
                        continue
                    if label in mapping and mapping[label] != name:
                        if col == "esto_label":
                            print(
                                f"Warning: overriding label {label} name "
                                f"{mapping[label]} with {name} (esto_label)."
                            )
                            mapping[label] = name
                        else:
                            print(
                                f"Warning: keeping existing name for label {label} "
                                f"({mapping[label]}); skipping {name} from {col}."
                            )
                        continue
                    mapping[label] = name

            if not mapping:
                print(f"No usable mappings found in {path}; trying next file.")
                continue

            print(f"Loaded code-to-name mapping from {path}: {len(mapping)} entries")
            return mapping

        raise ValueError("Code-to-name mapping not found in configured files.")
    except Exception as exc:
        print(f"Failed to load code-to-name mapping: {exc}")
        try_debug_breakpoint()
        raise


def is_code_like_label(label):
    """Return True when the label looks like a coded fuel/sector/flow."""
    try:
        if label is None:
            return False
        text = str(label).strip()
        if not text:
            return False
        if text[0].isdigit():
            return True
        return any(token in text for token in ["_", "."])
    except Exception as exc:
        print(f"Failed to check if label is code-like: {exc}")
        try_debug_breakpoint()
        raise


def resolve_label_name(label, code_to_name_mapping, context_label=""):
    """Return a mapped label name or raise when missing."""
    try:
        if not code_to_name_mapping:
            raise ValueError("Code-to-name mapping is empty.")
        if label is None:
            return label
        if isinstance(label, float) and pd.isna(label):
            return label
        text = str(label).strip()
        if text == "":
            return text
        if not is_code_like_label(text):
            return text
        if text in code_to_name_mapping:
            return code_to_name_mapping[text]
        if text in code_to_name_mapping.values():
            return text
        context_text = f" ({context_label})" if context_label else ""
        raise ValueError(f"Missing code-to-name mapping for label: {text}{context_text}")
    except Exception as exc:
        print(f"Failed to resolve label name for {label}: {exc}")
        try_debug_breakpoint()
        raise


def map_code_label(label, code_to_name_mapping):
    """Return a label mapped to a human-readable name when available."""
    try:
        return resolve_label_name(label, code_to_name_mapping)
    except Exception as exc:
        print(f"Failed to map label {label}: {exc}")
        try_debug_breakpoint()
        raise


def map_label_list(labels, code_to_name_mapping):
    """Map a list of labels to human-readable names."""
    try:
        return [
            resolve_label_name(label, code_to_name_mapping, context_label="label list")
            for label in labels
        ]
    except Exception as exc:
        print(f"Failed to map label list: {exc}")
        try_debug_breakpoint()
        raise


def format_fuel_label(label, code_to_name_mapping):
    """Return a fuel label formatted with numeric code prefix when available."""
    try:
        if label is None:
            return label
        text = str(label)
        if text == "nan":
            return label
        name = resolve_label_name(label, code_to_name_mapping, context_label="format_fuel_label")
        segments = _extract_numeric_segments(label)
        if segments:
            code_prefix = ".".join(segments)
            return f"{code_prefix} {name}"
        return name
    except Exception as exc:
        print(f"Failed to format fuel label {label}: {exc}")
        try_debug_breakpoint()
        raise


def map_series_index(series, code_to_name_mapping):
    """Map a Series index to human-readable names."""
    try:
        new_index = [
            resolve_label_name(idx, code_to_name_mapping, context_label="series index")
            for idx in series.index
        ]
        return series.rename(index=dict(zip(series.index, new_index)))
    except Exception as exc:
        print(f"Failed to map series index: {exc}")
        try_debug_breakpoint()
        raise


def split_auxiliary_fuels(negative_series, primary_input, threshold_ratio, include_all=False):
    """Split negative fuels into primary input and auxiliary candidates."""
    try:
        if negative_series is None or negative_series.empty:
            return []
        primary_label = primary_input
        if primary_label not in negative_series.index:
            matches = [
                label
                for label in negative_series.index
                if _match_code_prefix(label, primary_input)
            ]
            if matches:
                primary_label = matches[0]
            else:
                return []
        primary_value = abs(negative_series.loc[primary_label])
        auxiliary = []
        for fuel_label, value in negative_series.items():
            if fuel_label == primary_label:
                continue
            if include_all or abs(value) <= primary_value * threshold_ratio:
                auxiliary.append(fuel_label)
        return auxiliary
    except Exception as exc:
        print(f"Failed to split auxiliary fuels: {exc}")
        try_debug_breakpoint()
        raise


def get_all_other_negative_fuels(negative_series, primary_input):
    """Return all negative fuel labels except the primary input."""
    try:
        if negative_series is None or negative_series.empty:
            return []
        return [label for label in negative_series.index if label != primary_input]
    except Exception as exc:
        print(f"Failed to build auxiliary fuels from negatives: {exc}")
        try_debug_breakpoint()
        raise


def has_required_columns(df, required_sets, context_label):
    """Check for required column sets and warn/skip when missing."""
    try:
        for required_columns in required_sets:
            if all(col in df.columns for col in required_columns):
                return True
        print(
            f"{context_label}: missing required columns for this dataset, skipping."
        )
        return False
    except Exception as exc:
        print(f"Failed to validate columns for {context_label}: {exc}")
        try_debug_breakpoint()
        raise


def build_auxiliary_ratios(negative_series, auxiliary_fuels, output_total):
    """Return auxiliary fuel ratios (abs(input)/output)."""
    try:
        ratios = {}
        if negative_series is None or output_total == 0:
            return ratios
        for label in auxiliary_fuels:
            if label in negative_series.index:
                ratios[label] = abs(negative_series.get(label)) / output_total
        return ratios
    except Exception as exc:
        print(f"Failed to build auxiliary ratios: {exc}")
        try_debug_breakpoint()
        raise


def build_auxiliary_from_losses(loss_values, output_total):
    """Return auxiliary fuels/ratios derived from loss values."""
    try:
        if not loss_values:
            return [], {}
        fuels = []
        ratios = {}
        for label, value in loss_values.items():
            fuels.append(label)
            ratios[label] = abs(value) / output_total if output_total else 0.0
        return fuels, ratios
    except Exception as exc:
        print(f"Failed to build auxiliary fuels from losses: {exc}")
        try_debug_breakpoint()
        raise


def merge_loss_into_auxiliary(
    auxiliary_fuels, auxiliary_ratios, loss_values, output_total, feedstock_label
):
    """Treat own use/loss fuels as auxiliary (unless same as feedstock)."""
    try:
        if not loss_values or output_total == 0:
            return auxiliary_fuels, auxiliary_ratios
        updated_fuels = list(auxiliary_fuels) if auxiliary_fuels else []
        updated_ratios = dict(auxiliary_ratios) if auxiliary_ratios else {}
        for label, value in loss_values.items():
            if label == feedstock_label:
                continue
            if label not in updated_fuels:
                updated_fuels.append(label)
            updated_ratios[label] = abs(value) / output_total
        return updated_fuels, updated_ratios
    except Exception as exc:
        print(f"Failed to merge loss fuels into auxiliary list: {exc}")
        try_debug_breakpoint()
        raise


def filter_loss_values_for_feedstock(loss_values, feedstock_label):
    """Return loss values for the feedstock fuel only."""
    try:
        if not loss_values or not feedstock_label:
            return {}
        if feedstock_label not in loss_values:
            return {}
        return {feedstock_label: loss_values[feedstock_label]}
    except Exception as exc:
        print(f"Failed to filter loss values for feedstock: {exc}")
        try_debug_breakpoint()
        raise


def get_loss_total_for_efficiency(loss_values, feedstock_label, output_label):
    """Return loss total for efficiency using feedstock/output fuel losses only."""
    try:
        if not loss_values:
            return 0.0
        relevant_labels = {feedstock_label, output_label}
        return sum(
            value for label, value in loss_values.items() if label in relevant_labels
        )
    except Exception as exc:
        print(f"Failed to build loss total for efficiency: {exc}")
        try_debug_breakpoint()
        raise


def print_leap_structure_header(title):
    """Print a section header that mirrors the LEAP branch structure."""
    try:
        print("")
        print(title)
        print("-" * len(title))
    except Exception as exc:
        print(f"Failed to print LEAP structure header: {exc}")
        try_debug_breakpoint()
        raise


def format_value(value):
    """Format numeric values for LEAP structure output."""
    try:
        if isinstance(value, str):
            return value
        if value is None or pd.isna(value):
            return ""
        return f"{float(value):.6f}"
    except Exception as exc:
        print(f"Failed to format value {value}: {exc}")
        try_debug_breakpoint()
        raise


def build_year_rows(branch_path, measure, scenario, value_by_year, units, scale, per_value):
    """Return log-style rows for a LEAP import file.

    Inputs:
        branch_path: LEAP branch path string (e.g., Transformation\\Coke ovens)
        measure: LEAP variable name (e.g., "Process Efficiency")
        scenario: scenario label (e.g., "Current Accounts")
        value_by_year: dict of year -> value
        units: units string for LEAP
        scale: scale string for LEAP
        per_value: per... string for LEAP

    Outputs:
        List[dict] with fields expected by finalise_export_df.

    Side effects:
        None.
    """
    try:
        rows = []
        for year, value in sorted(value_by_year.items()):
            rows.append(
                {
                    "Branch_Path": branch_path,
                    "Scenario": scenario,
                    "Measure": measure,
                    "Units": units,
                    "Scale": scale,
                    "Per...": per_value,
                    "Date": int(year),
                    "Value": float(value),
                }
            )
        return rows
    except Exception as exc:
        print(f"Failed to build year rows for {branch_path}: {exc}")
        try_debug_breakpoint()
        raise


def build_value_by_year(value, base_year, final_year):
    """Return a dict of year -> value for the given range."""
    try:
        return {year: value for year in range(base_year, final_year + 1)}
    except Exception as exc:
        print(f"Failed to build value-by-year map: {exc}")
        try_debug_breakpoint()
        raise


def coerce_value_by_year(value, base_year, final_year):
    """Return a year->value dict from a scalar, series, or dict."""
    try:
        if isinstance(value, dict):
            return {int(year): float(val) for year, val in value.items()}
        if isinstance(value, pd.Series):
            return {int(year): float(val) for year, val in value.items()}
        return build_value_by_year(value, base_year, final_year)
    except Exception as exc:
        print(f"Failed to coerce value to year map: {exc}")
        try_debug_breakpoint()
        raise


def summarize_numeric_value(value, summary="sum"):
    """Summarize a scalar or year->value mapping for tables."""
    try:
        if isinstance(value, dict):
            values = [val for val in value.values() if val is not None]
            if not values:
                return None
            if summary == "mean":
                return sum(values) / len(values)
            return sum(values)
        return value
    except Exception as exc:
        print(f"Failed to summarize numeric value: {exc}")
        try_debug_breakpoint()
        raise


def format_filename_segment(value):
    """Return a file-safe string for economy or scenario labels."""
    try:
        if value is None:
            return ""
        text = str(value).strip()
        if not text:
            return ""
        sanitized = re.sub(r"[^A-Za-z0-9_-]+", "_", text)
        return sanitized.strip("_") or text
    except Exception as exc:
        print(f"Failed to format filename segment for {value}: {exc}")
        try_debug_breakpoint()
        raise


def build_export_filename(template, fallback, economy, scenario):
    """Format an export filename with economy/scenario segments."""
    try:
        if not template:
            return fallback
        economy_segment = format_filename_segment(economy)
        scenario_segment = format_filename_segment(scenario)
        if "{economy}" not in template and "{scenario}" not in template:
            return template
        return template.format(
            economy=economy_segment,
            scenario=scenario_segment,
        )
    except Exception as exc:
        print(f"Failed to build export filename: {exc}")
        try_debug_breakpoint()
        return fallback or template

def build_process_record(
    economy,
    sector_title,
    process_name,
    output_values,
    feedstock_values,
    efficiency,
    auxiliary_ratios,
    loss_values,
    loss_total,
    feedstock_shares=None,
    input_total=None,
    output_import_targets=None,
    output_export_targets=None,
):
    """Return a standardized record for a transformation process."""
    try:
        return {
            "economy": economy,
            "sector_title": sector_title,
            "process_name": process_name,
            "output_values": dict(output_values or {}),
            "feedstock_values": dict(feedstock_values or {}),
            "feedstock_shares": dict(feedstock_shares or {}),
            "efficiency": efficiency,
            "auxiliary_ratios": dict(auxiliary_ratios or {}),
            "loss_values": dict(loss_values or {}),
            "loss_total": loss_total,
            "input_total": input_total,
            "output_import_targets": dict(output_import_targets or {}),
            "output_export_targets": dict(output_export_targets or {}),
        }
    except Exception as exc:
        print(f"Failed to build process record for {process_name}: {exc}")
        try_debug_breakpoint()
        raise


def append_process_record(process_records, record):
    """Append a process record to the list when provided."""
    try:
        if process_records is None:
            return
        process_records.append(record)
    except Exception as exc:
        print(f"Failed to append process record: {exc}")
        try_debug_breakpoint()
        raise


def select_primary_label(value_map):
    """Return the label with the largest absolute value."""
    try:
        if not value_map:
            return ""
        return max(
            value_map,
            key=lambda label: abs(summarize_numeric_value(value_map.get(label, 0), summary="sum") or 0),
        )
    except Exception as exc:
        print(f"Failed to select primary label: {exc}")
        try_debug_breakpoint()
        raise


def build_transformation_process_table(process_records, code_to_name_mapping):
    """Return a process-level summary table for transformations."""
    try:
        rows = []
        for record in process_records:
            output_label = select_primary_label(record.get("output_values"))
            feedstock_label = select_primary_label(record.get("feedstock_values"))
            output_value = summarize_numeric_value(
                record.get("output_values", {}).get(output_label), summary="sum"
            )
            feedstock_value = summarize_numeric_value(
                record.get("feedstock_values", {}).get(feedstock_label), summary="sum"
            )
            efficiency_value = summarize_numeric_value(
                record.get("efficiency"), summary="mean"
            )
            rows.append(
                {
                    "economy": record.get("economy"),
                    "sector_title": record.get("sector_title"),
                    "process_name": record.get("process_name"),
                    "output_label": format_fuel_label(output_label, code_to_name_mapping),
                    "output_value": output_value,
                    "feedstock_label": format_fuel_label(feedstock_label, code_to_name_mapping),
                    "feedstock_value": feedstock_value,
                    "efficiency": efficiency_value,
                    "loss_total": record.get("loss_total"),
                    "auxiliary_count": len(record.get("auxiliary_ratios", {})),
                }
            )
        return pd.DataFrame(rows)
    except Exception as exc:
        print(f"Failed to build transformation process table: {exc}")
        try_debug_breakpoint()
        raise


def build_transformation_detail_table(process_records, code_to_name_mapping):
    """Return a long-form detail table for outputs, feedstocks, and auxiliaries."""
    try:
        rows = []
        for record in process_records:
            economy = record.get("economy")
            sector_title = record.get("sector_title")
            process_name = record.get("process_name")
            if (
                INCLUDE_OUTPUT_SERIES_IN_LEAP_EXPORT
                and TRANSFORMATION_OUTPUT_VARIABLES.get("output")
            ):
                for label, value in record.get("output_values", {}).items():
                    summary_value = summarize_numeric_value(value, summary="sum")
                    rows.append(
                        {
                            "economy": economy,
                            "sector_title": sector_title,
                            "process_name": process_name,
                            "category": "output",
                            "fuel_label": label,
                            "fuel_label_display": format_fuel_label(label, code_to_name_mapping),
                            "value": summary_value,
                            "units": DEFAULT_OUTPUT_UNITS,
                            "per": "",
                        }
                    )
            if TRANSFORMATION_OUTPUT_VARIABLES.get("output_import_target"):
                for label, value in record.get("output_import_targets", {}).items():
                    summary_value = summarize_numeric_value(value, summary="sum")
                    rows.append(
                        {
                            "economy": economy,
                            "sector_title": sector_title,
                            "process_name": process_name,
                            "category": "output_import_target",
                            "fuel_label": label,
                            "fuel_label_display": format_fuel_label(label, code_to_name_mapping),
                            "value": summary_value,
                            "units": DEFAULT_OUTPUT_UNITS,
                            "per": "",
                        }
                    )
            if TRANSFORMATION_OUTPUT_VARIABLES.get("output_export_target"):
                for label, value in record.get("output_export_targets", {}).items():
                    summary_value = summarize_numeric_value(value, summary="sum")
                    rows.append(
                        {
                            "economy": economy,
                            "sector_title": sector_title,
                            "process_name": process_name,
                            "category": "output_export_target",
                            "fuel_label": label,
                            "fuel_label_display": format_fuel_label(label, code_to_name_mapping),
                            "value": summary_value,
                            "units": DEFAULT_OUTPUT_UNITS,
                            "per": "",
                        }
                    )
            if TRANSFORMATION_OUTPUT_VARIABLES.get("feedstock_share"):
                for label, value in record.get("feedstock_shares", {}).items():
                    summary_value = summarize_numeric_value(value, summary="mean")
                    rows.append(
                        {
                            "economy": economy,
                            "sector_title": sector_title,
                            "process_name": process_name,
                            "category": "feedstock_share",
                            "fuel_label": label,
                            "fuel_label_display": format_fuel_label(label, code_to_name_mapping),
                            "value": summary_value,
                            "units": DEFAULT_FEEDSTOCK_UNITS,
                            "per": "",
                        }
                    )
            if TRANSFORMATION_OUTPUT_VARIABLES.get("process_efficiency") and record.get("efficiency") is not None:
                efficiency_value = summarize_numeric_value(
                    record.get("efficiency"), summary="mean"
                )
                rows.append(
                    {
                        "economy": economy,
                        "sector_title": sector_title,
                        "process_name": process_name,
                        "category": "process_efficiency",
                        "fuel_label": "",
                        "fuel_label_display": "",
                        "value": efficiency_value,
                        "units": DEFAULT_EFFICIENCY_UNITS,
                        "per": "",
                    }
                )
            if TRANSFORMATION_OUTPUT_VARIABLES.get("auxiliary_ratio"):
                for label, value in record.get("auxiliary_ratios", {}).items():
                    summary_value = summarize_numeric_value(value, summary="mean")
                    rows.append(
                        {
                            "economy": economy,
                            "sector_title": sector_title,
                            "process_name": process_name,
                            "category": "auxiliary_ratio",
                            "fuel_label": label,
                            "fuel_label_display": format_fuel_label(label, code_to_name_mapping),
                            "value": summary_value,
                            "units": DEFAULT_AUXILIARY_UNITS,
                            "per": DEFAULT_AUXILIARY_PER,
                        }
                    )
            if TRANSFORMATION_OUTPUT_VARIABLES.get("loss_value"):
                for label, value in record.get("loss_values", {}).items():
                    summary_value = summarize_numeric_value(value, summary="sum")
                    rows.append(
                        {
                            "economy": economy,
                            "sector_title": sector_title,
                            "process_name": process_name,
                            "category": "loss_value",
                            "fuel_label": label,
                            "fuel_label_display": format_fuel_label(label, code_to_name_mapping),
                            "value": summary_value,
                            "units": DEFAULT_OUTPUT_UNITS,
                            "per": "",
                        }
                    )
        return pd.DataFrame(rows)
    except Exception as exc:
        print(f"Failed to build transformation detail table: {exc}")
        try_debug_breakpoint()
        raise


def build_branch_path(parts):
    """Return a LEAP branch path from parts."""
    try:
        cleaned_parts = [str(part).strip() for part in parts if part and str(part).strip()]
        return sanitize_leap_branch_path("\\".join(cleaned_parts))
    except Exception as exc:
        print(f"Failed to build branch path from {parts}: {exc}")
        try_debug_breakpoint()
        raise


def build_scenario_specific_rows(
    process_records,
    scenario,
    scenario_config,
    base_year,
    final_year,
):
    """Return scenario-specific LEAP rows (hook for future custom rows)."""
    try:
        if not scenario_config or not scenario_config.get("include_current_account_rows"):
            return []
        # Placeholder for future Current Accounts-only rows. No additional rows today.
        return []
    except Exception as exc:
        print(f"Failed to build scenario-specific rows for {scenario}: {exc}")
        try_debug_breakpoint()
        raise


def build_transformation_log_rows(
    process_records,
    scenario,
    region,
    base_year,
    final_year,
    code_to_name_mapping,
    scenario_config=None,
):
    """Return log-style rows for LEAP import from process records."""
    try:
        rows = []
        for record in process_records:
            sector_title = map_code_label(record.get("sector_title"), code_to_name_mapping)
            process_name = map_code_label(record.get("process_name"), code_to_name_mapping)
            output_values = record.get("output_values", {})
            feedstock_shares = record.get("feedstock_shares", {})
            auxiliary_ratios = record.get("auxiliary_ratios", {})
            efficiency = record.get("efficiency")

            if (
                INCLUDE_OUTPUT_SERIES_IN_LEAP_EXPORT
                and TRANSFORMATION_OUTPUT_VARIABLES.get("output")
            ):
                for label, value in output_values.items():
                    value_by_year = coerce_value_by_year(value, base_year, final_year)
                    branch_path = build_branch_path(
                        [
                            "Transformation",
                            sector_title,
                            "Output Fuels",
                            map_code_label(label, code_to_name_mapping),
                        ]
                    )
                    rows.extend(
                        build_year_rows(
                            branch_path,
                            "Output",
                            scenario,
                            value_by_year,
                            DEFAULT_OUTPUT_UNITS,
                            "",
                            "",
                        )
                    )
            if TRANSFORMATION_OUTPUT_VARIABLES.get("output_import_target"):
                for label, value in record.get("output_import_targets", {}).items():
                    value_by_year = coerce_value_by_year(value, base_year, final_year)
                    branch_path = build_branch_path(
                        [
                            "Transformation",
                            sector_title,
                            "Output Fuels",
                            map_code_label(label, code_to_name_mapping),
                        ]
                    )
                    rows.extend(
                        build_year_rows(
                            branch_path,
                            "Import Target",
                            scenario,
                            value_by_year,
                            DEFAULT_OUTPUT_UNITS,
                            "",
                            "",
                        )
                    )
            if TRANSFORMATION_OUTPUT_VARIABLES.get("output_export_target"):
                for label, value in record.get("output_export_targets", {}).items():
                    value_by_year = coerce_value_by_year(value, base_year, final_year)
                    branch_path = build_branch_path(
                        [
                            "Transformation",
                            sector_title,
                            "Output Fuels",
                            map_code_label(label, code_to_name_mapping),
                        ]
                    )
                    rows.extend(
                        build_year_rows(
                            branch_path,
                            "Export Target",
                            scenario,
                            value_by_year,
                            DEFAULT_OUTPUT_UNITS,
                            "",
                            "",
                        )
                    )

            if TRANSFORMATION_OUTPUT_VARIABLES.get("process_efficiency") and efficiency is not None:
                efficiency_by_year = coerce_value_by_year(efficiency, base_year, final_year)
                branch_path = build_branch_path(
                    ["Transformation", sector_title, "Processes", str(process_name)]
                )
                rows.extend(
                    build_year_rows(
                        branch_path,
                        "Process Efficiency",
                        scenario,
                        efficiency_by_year,
                        DEFAULT_EFFICIENCY_UNITS,
                        "",
                        "",
                    )
                )

            if TRANSFORMATION_OUTPUT_VARIABLES.get("feedstock_share"):
                for label, value in feedstock_shares.items():
                    value_by_year = coerce_value_by_year(value, base_year, final_year)
                    branch_path = build_branch_path(
                        [
                            "Transformation",
                            sector_title,
                            "Processes",
                            str(process_name),
                            "Feedstock Fuels",
                            map_code_label(label, code_to_name_mapping),
                        ]
                    )
                    rows.extend(
                        build_year_rows(
                            branch_path,
                            "Feedstock Fuel Share",
                            scenario,
                            value_by_year,
                            DEFAULT_FEEDSTOCK_UNITS,
                            "",
                            "",
                        )
                    )

            if TRANSFORMATION_OUTPUT_VARIABLES.get("auxiliary_ratio"):
                for label, value in auxiliary_ratios.items():
                    value_by_year = coerce_value_by_year(value, base_year, final_year)
                    branch_path = build_branch_path(
                        [
                            "Transformation",
                            sector_title,
                            "Processes",
                            str(process_name),
                            "Auxiliary Fuels",
                            map_code_label(label, code_to_name_mapping),
                        ]
                    )
                    rows.extend(
                        build_year_rows(
                            branch_path,
                            "Auxiliary Fuel Use",
                            scenario,
                            value_by_year,
                            DEFAULT_AUXILIARY_UNITS,
                            "",
                            DEFAULT_AUXILIARY_PER,
                        )
                    )

        rows.extend(
            build_scenario_specific_rows(
                process_records,
                scenario,
                scenario_config,
                base_year,
                final_year,
            )
        )
        return rows
    except Exception as exc:
        print(f"Failed to build transformation log rows: {exc}")
        try_debug_breakpoint()
        raise


def build_data_expression(row, year_cols):
    """Return a LEAP Data(...) expression from year columns."""
    try:
        parts = []
        for year in year_cols:
            value = row.get(year)
            if value is None or pd.isna(value):
                value = 0.0
            parts.append(f"{int(year)},{float(value)}")
        return f"Data({', '.join(parts)})"
    except Exception as exc:
        print(f"Failed to build Data expression: {exc}")
        try_debug_breakpoint()
        raise


def build_expression_export_df(export_df):
    """Return a LEAP sheet df with Expression and no year columns."""
    try:
        year_cols = sorted([col for col in export_df.columns if str(col).isdigit()])
        expression_df = export_df.copy()
        expression_df["Expression"] = expression_df.apply(
            lambda row: build_data_expression(row, year_cols),
            axis=1,
        )
        expression_df = expression_df.drop(columns=year_cols)
        base_cols = ["Branch Path", "Variable", "Scenario", "Region", "Scale", "Units", "Per...", "Expression"]
        level_cols = [col for col in expression_df.columns if col.startswith("Level ")]
        expression_df = expression_df[base_cols + level_cols]
        return expression_df
    except Exception as exc:
        print(f"Failed to build expression export df: {exc}")
        try_debug_breakpoint()
        raise


def build_export_from_log_rows(log_rows, scenario_label, region, base_year, final_year):
    """Finalize a log row list into LEAP export/log DataFrames."""
    try:
        log_df = pd.DataFrame(log_rows)
        export_df = finalise_export_df(log_df, scenario_label, region, base_year, final_year)
        return export_df, log_df
    except Exception as exc:
        print(f"Failed to build export from log rows: {exc}")
        try_debug_breakpoint()
        raise


def save_transformation_summaries(
    process_records,
    code_to_name_mapping,
    output_dir,
    process_filename,
    detail_filename,
):
    """Save transformation summary tables to CSV."""
    try:
        if not process_records:
            print("No process records available for summary tables.")
            return None, None
        process_summary = build_transformation_process_table(
            process_records,
            code_to_name_mapping,
        )
        detail_summary = build_transformation_detail_table(
            process_records,
            code_to_name_mapping,
        )
        os.makedirs(output_dir, exist_ok=True)
        process_summary_path = os.path.join(output_dir, process_filename)
        detail_summary_path = os.path.join(output_dir, detail_filename)
        process_summary.to_csv(process_summary_path, index=False)
        detail_summary.to_csv(detail_summary_path, index=False)
        print(f"Saved transformation process summary to {process_summary_path}")
        print(f"Saved transformation detail summary to {detail_summary_path}")
        return process_summary, detail_summary
    except Exception as exc:
        print(f"Failed to save transformation summary tables: {exc}")
        try_debug_breakpoint()
        raise


def save_transformation_export(
    process_records,
    region,
    base_year,
    final_year,
    code_to_name_mapping,
    output_dir,
    output_filename,
    model_name,
    scenarios,
):
    """Save a LEAP import file built from process records across scenarios."""
    try:
        if not process_records:
            print("No process records available for LEAP export.")
            return None
        scenario_configs = {
            scenario: get_scenario_export_config(scenario) for scenario in scenarios
        }
        combined_base_year, combined_final_year = compute_combined_year_range(
            base_year, final_year, scenario_configs
        )
        combined_rows = []
        for scenario in scenarios:
            scenario_config = scenario_configs.get(scenario, {})
            combined_rows.extend(
                build_transformation_log_rows(
                    process_records,
                    scenario,
                    region,
                    combined_base_year,
                    combined_final_year,
                    code_to_name_mapping,
                    scenario_config=scenario_config,
                )
            )
        if not combined_rows:
            print("No log rows generated across scenarios; skipping export.")
            return None
        scenario_label = ", ".join(scenarios)
        export_df, log_df = build_export_from_log_rows(
            combined_rows,
            scenario_label,
            region,
            combined_base_year,
            combined_final_year,
        )
        if export_df is None:
            print("No export dataframe created for LEAP export.")
            return None
        leap_expression_df = build_expression_export_df(export_df)
        os.makedirs(output_dir, exist_ok=True)
        export_path = os.path.join(output_dir, output_filename)
        save_export_files(
            leap_expression_df,
            export_df,
            export_path,
            combined_base_year,
            combined_final_year,
            model_name,
        )
        return export_path
    except Exception as exc:
        print(f"Failed to save transformation LEAP export: {exc}")
        try_debug_breakpoint()
        raise


def print_leap_structure_block(
    title,
    output_fuels,
    process_name,
    feedstock_fuels,
    auxiliary_fuels,
    loss_fuels=None,
    code_to_name_mapping=None,
    output_fuel_values=None,
    process_value=None,
    feedstock_fuel_values=None,
    auxiliary_fuel_values=None,
    loss_fuel_values=None,
    other_feedstock_fuels=None,
    other_feedstock_values=None,
    other_feedstock_ratios=None,
):
    """Print a LEAP-structure outline for a transformation process."""
    try:
        output_pairs = [
            (label, format_fuel_label(label, code_to_name_mapping)) for label in output_fuels
        ]
        feedstock_pairs = [
            (label, format_fuel_label(label, code_to_name_mapping)) for label in feedstock_fuels
        ]
        auxiliary_pairs = [
            (label, format_fuel_label(label, code_to_name_mapping)) for label in auxiliary_fuels
        ]
        loss_pairs = [
            (label, format_fuel_label(label, code_to_name_mapping))
            for label in (loss_fuels or [])
        ]
        process_name = map_code_label(process_name, code_to_name_mapping)

        print_leap_structure_header(title)
        print("Output fuels (export target, import target):")
        for raw_label, fuel in output_pairs:
            fuel_value = ""
            if output_fuel_values is not None:
                fuel_value = format_value(output_fuel_values.get(raw_label))
            print(f"  - {fuel}" + (f" {fuel_value}" if fuel_value else ""))
        print("Processes (process efficiency):")
        process_value_text = ""
        if process_value is not None:
            process_value_text = f" {format_value(process_value)}"
        print(f"  - {process_name}:{process_value_text}")
        if feedstock_fuels:
            print("      Feedstock fuels:")
            for raw_label, fuel in feedstock_pairs:
                fuel_value = ""
                if feedstock_fuel_values is not None:
                    fuel_value = format_value(feedstock_fuel_values.get(raw_label))
                print(f"        - {fuel}" + (f" {fuel_value}" if fuel_value else ""))
        if auxiliary_fuels:
            print("      Auxiliary fuels (Aux fuel use pj/pj output):")
            for raw_label, fuel in auxiliary_pairs:
                fuel_value = ""
                if auxiliary_fuel_values is not None:
                    fuel_value = format_value(auxiliary_fuel_values.get(raw_label))
                print(f"        - {fuel}" + (f" {fuel_value}" if fuel_value else ""))
        if other_feedstock_fuels:
            other_feedstock_pairs = [
                (label, format_fuel_label(label, code_to_name_mapping))
                for label in other_feedstock_fuels
            ]
            total_other_feedstock = 0.0
            if other_feedstock_values:
                total_other_feedstock = sum(
                    value for value in other_feedstock_values.values() if value is not None
                )
            total_text = f" (total {format_value(total_other_feedstock)})"
            print(
                "      Other feedstock fuels (set as aux fuel use)"
                + total_text
                + ":"
            )
            for raw_label, fuel in other_feedstock_pairs:
                fuel_value = ""
                fuel_ratio = ""
                if other_feedstock_values is not None:
                    fuel_value = format_value(other_feedstock_values.get(raw_label))
                if other_feedstock_ratios is not None:
                    fuel_ratio = format_value(other_feedstock_ratios.get(raw_label))
                value_text = f" {fuel_value}" if fuel_value else ""
                ratio_text = f" ({fuel_ratio} pj/pj)" if fuel_ratio else ""
                print(f"        - {fuel}" + value_text + ratio_text)
        if loss_pairs:
            print("      Own use and losses (PJ):")
            for raw_label, fuel in loss_pairs:
                fuel_value = ""
                if loss_fuel_values is not None:
                    fuel_value = format_value(loss_fuel_values.get(raw_label))
                print(f"        - {fuel}" + (f" {fuel_value}" if fuel_value else ""))
        print("")
    except Exception as exc:
        print(f"Failed to print LEAP structure block: {exc}")
        try_debug_breakpoint()
        raise


def calculate_efficiency(output_df, input_df, loss_df, year_cols):
    """Compute process efficiency as output / (abs(input) + abs(losses))."""
    try:
        output_total = sum_years(output_df, year_cols)
        input_total = abs(sum_years(input_df, year_cols))
        loss_total = abs(sum_years(loss_df, year_cols)) if loss_df is not None else 0.0
        if (input_total + loss_total) == 0:
            return 0.0
        return output_total / (input_total + loss_total)
    except Exception as exc:
        print(f"Failed to calculate efficiency: {exc}")
        try_debug_breakpoint()
        raise


def calculate_aux_fuel_use(aux_df, output_df, year_cols):
    """Compute auxiliary fuel use as abs(aux input) / output."""
    try:
        output_total = sum_years(output_df, year_cols)
        aux_total = abs(sum_years(aux_df, year_cols))
        if output_total == 0:
            return 0.0
        return aux_total / output_total
    except Exception as exc:
        print(f"Failed to calculate auxiliary fuel use: {exc}")
        try_debug_breakpoint()
        raise


def analyze_lng_liquefaction_regas(
    esto_data,
    year_cols,
    start_year,
    economy,
    code_to_name_mapping,
    loss_data,
    loss_year_cols,
    sector_config=None,
    process_records=None,
):
    """Estimate LNG liquefaction/regasification efficiency and auxiliary fuel use."""
    try:
        lng_config = sector_config or MAJOR_SECTOR_CONFIG["lng"]
        fuel_codes = {
            "natural_gas": "08_01_natural_gas",
            "lng": "08_02_lng",
            "gas_works_gas": "08_03_gas_works_gas",
            "lignite": "01_05_lignite",
            "electricity": "17_electricity",
        }
        lng_sub2 = lng_config["transformation_sub2"][0]
        print(f"\n==== LNG liquefaction/regasification ({economy}) ====")
        if not has_required_columns(
            esto_data,
            [["sub2sectors", "subfuels", "fuels"], ["flows", "products"]],
            "LNG liquefaction/regasification",
        ):
            return
        year_cols_from_start = get_years_from(year_cols, start_year)

        # LNG data uses the ninth data file due to adjustments in merged_file_energy_ALL_20250814.
        regas_output = select_rows(
            esto_data,
            {
                "economy": economy,
                "sub2sectors": lng_sub2,
                "subfuels": fuel_codes["natural_gas"],
            },
        )
        liquefaction_output = select_rows(
            esto_data,
            {"economy": economy, "sub2sectors": lng_sub2, "subfuels": fuel_codes["lng"]},
        )
        print_sector_rows(
            esto_data,
            "LNG liquefaction/regas rows",
            {"economy": economy, "sub2sectors": lng_sub2},
            year_cols,
            start_year,
            code_to_name_mapping,
        )
        loss_series, loss_total, loss_fuel_values, loss_values_by_year = build_loss_context(
            loss_data,
            loss_year_cols,
            start_year,
            economy,
            "lng",
            lng_sub2,
        )
        export_base_year = EXPORT_BASE_YEAR
        export_final_year = EXPORT_FINAL_YEAR
        export_years = list(range(export_base_year, export_final_year + 1))
        is_regas_negative = sum_years(regas_output, year_cols_from_start) < 0
        is_liq_negative = sum_years(liquefaction_output, year_cols_from_start) < 0

        lng_importer = False
        lng_exporter = False
        regas_message = None
        liq_message = None
        if is_regas_negative:
            regas_message = "Regas output negative: likely LNG exporter (no regas)."
            lng_exporter = True
        if is_liq_negative:
            liq_message = "Liquefaction output negative: likely LNG importer (no liquefaction)."
            lng_importer = True

        if lng_importer:
            lng_input = select_rows(
                esto_data,
                {"economy": economy, "sub2sectors": lng_sub2, "subfuels": fuel_codes["lng"]},
            )
            electricity_input = select_rows(
                esto_data,
                {
                    "economy": economy,
                    "sub2sectors": "10_01_03_liquefaction_regasification_plants",
                    "fuels": fuel_codes["electricity"],
                },
            )
            regas_output_series = ensure_full_year_series(
                sum_years_by_year(regas_output, year_cols, start_year),
                export_base_year,
                export_final_year,
            )
            lng_input_series = ensure_full_year_series(
                sum_years_by_year(lng_input, year_cols, start_year).abs(),
                export_base_year,
                export_final_year,
            )
            regas_output_total = regas_output_series.sum()
            lng_input_total = lng_input_series.sum()
            regas_loss_total_by_year = get_loss_total_for_efficiency_by_year(
                loss_values_by_year,
                fuel_codes["lng"],
                fuel_codes["natural_gas"],
                export_years,
            )
            regas_loss_total = loss_total
            efficiency_regas_series = compute_efficiency_by_year(
                regas_output_series,
                lng_input_series,
                regas_loss_total_by_year,
            )
            electricity_series = ensure_full_year_series(
                sum_years_by_year(electricity_input, year_cols, start_year).abs(),
                export_base_year,
                export_final_year,
            )
            aux_fuel_use_regas_series = safe_divide_series(
                electricity_series,
                regas_output_series,
            )
            regas_aux_values = {
                fuel_codes["electricity"]: series_to_year_dict(
                    aux_fuel_use_regas_series, export_base_year, export_final_year
                ),
            }
            regas_aux_fuels, regas_aux_values = merge_loss_into_auxiliary_by_year(
                [fuel_codes["electricity"]],
                regas_aux_values,
                loss_values_by_year,
                regas_output_series,
                fuel_codes["lng"],
            )
            regas_loss_values = filter_loss_values_for_feedstock_by_year(
                loss_values_by_year,
                fuel_codes["lng"],
            )
            regas_loss_values_display = {
                label: summarize_numeric_value(values, summary="sum")
                for label, values in regas_loss_values.items()
            }
            print_leap_structure_block(
                "LNG regasification",
                [fuel_codes["natural_gas"]],
                "Regasification",
                [fuel_codes["lng"]],
                regas_aux_fuels,
                loss_fuels=list(regas_loss_values.keys()),
                code_to_name_mapping=code_to_name_mapping,
                output_fuel_values={fuel_codes["natural_gas"]: regas_output_total},
                process_value=f"{efficiency_regas_series.mean():.4f}",
                feedstock_fuel_values={fuel_codes["lng"]: lng_input_total},
                auxiliary_fuel_values={
                    label: summarize_numeric_value(values, summary="mean")
                    for label, values in regas_aux_values.items()
                },
                loss_fuel_values=regas_loss_values_display,
            )
            output_import_targets, output_export_targets = gather_output_target_dicts(
                economy,
                [fuel_codes["natural_gas"]],
                export_base_year,
                export_final_year,
            )
            record = build_process_record(
                economy,
                lng_config.get("title", "NG Liquefaction"),
                "Regasification",
                {fuel_codes["natural_gas"]: series_to_year_dict(regas_output_series, export_base_year, export_final_year)},
                {fuel_codes["lng"]: series_to_year_dict(lng_input_series, export_base_year, export_final_year)},
                series_to_year_dict(efficiency_regas_series, export_base_year, export_final_year),
                regas_aux_values,
                regas_loss_values,
                regas_loss_total,
                feedstock_shares={fuel_codes["lng"]: 1.0},
                input_total=lng_input_total,
                output_import_targets=output_import_targets,
                output_export_targets=output_export_targets,
            )
            append_process_record(process_records, record)
            print(
                "Estimated regas aux fuel use (electricity per PJ output, mean): "
                f"{aux_fuel_use_regas_series.mean():.6f}"
            )
            if liq_message:
                print(liq_message)
            print(f"Estimated regasification efficiency (mean): {efficiency_regas_series.mean():.4f}")

        if lng_exporter:
            natgas_input = select_rows(
                esto_data,
                {
                    "economy": economy,
                    "sub2sectors": lng_sub2,
                    "subfuels": fuel_codes["natural_gas"],
                },
            )
            electricity_input = select_rows(
                esto_data,
                {
                    "economy": economy,
                    "sub2sectors": "10_01_03_liquefaction_regasification_plants",
                    "fuels": fuel_codes["electricity"],
                },
            )
            liquefaction_output_series = ensure_full_year_series(
                sum_years_by_year(liquefaction_output, year_cols, start_year),
                export_base_year,
                export_final_year,
            )
            natgas_input_series = ensure_full_year_series(
                sum_years_by_year(natgas_input, year_cols, start_year).abs(),
                export_base_year,
                export_final_year,
            )
            liquefaction_output_total = liquefaction_output_series.sum()
            natgas_input_total = natgas_input_series.sum()
            liq_loss_total_by_year = get_loss_total_for_efficiency_by_year(
                loss_values_by_year,
                fuel_codes["natural_gas"],
                fuel_codes["lng"],
                export_years,
            )
            liq_loss_total = loss_total
            efficiency_liquefaction_series = compute_efficiency_by_year(
                liquefaction_output_series,
                natgas_input_series,
                liq_loss_total_by_year,
            )
            electricity_series = ensure_full_year_series(
                sum_years_by_year(electricity_input, year_cols, start_year).abs(),
                export_base_year,
                export_final_year,
            )
            aux_fuel_use_liquefaction_series = safe_divide_series(
                electricity_series,
                liquefaction_output_series,
            )
            liq_aux_values = {
                fuel_codes["electricity"]: series_to_year_dict(
                    aux_fuel_use_liquefaction_series, export_base_year, export_final_year
                ),
            }
            liq_aux_fuels, liq_aux_values = merge_loss_into_auxiliary_by_year(
                [fuel_codes["electricity"]],
                liq_aux_values,
                loss_values_by_year,
                liquefaction_output_series,
                fuel_codes["natural_gas"],
            )
            liq_loss_values = filter_loss_values_for_feedstock_by_year(
                loss_values_by_year,
                fuel_codes["natural_gas"],
            )
            liq_loss_values_display = {
                label: summarize_numeric_value(values, summary="sum")
                for label, values in liq_loss_values.items()
            }
            print_leap_structure_block(
                "LNG liquefaction",
                [fuel_codes["lng"]],
                "Liquefaction",
                [fuel_codes["natural_gas"]],
                liq_aux_fuels,
                loss_fuels=list(liq_loss_values.keys()),
                code_to_name_mapping=code_to_name_mapping,
                output_fuel_values={fuel_codes["lng"]: liquefaction_output_total},
                process_value=f"{efficiency_liquefaction_series.mean():.4f}",
                feedstock_fuel_values={fuel_codes["natural_gas"]: natgas_input_total},
                auxiliary_fuel_values={
                    label: summarize_numeric_value(values, summary="mean")
                    for label, values in liq_aux_values.items()
                },
                loss_fuel_values=liq_loss_values_display,
            )
            output_import_targets, output_export_targets = gather_output_target_dicts(
                economy,
                [fuel_codes["lng"]],
                export_base_year,
                export_final_year,
            )
            record = build_process_record(
                economy,
                lng_config.get("title", "NG Liquefaction"),
                "Liquefaction",
                {fuel_codes["lng"]: series_to_year_dict(liquefaction_output_series, export_base_year, export_final_year)},
                {fuel_codes["natural_gas"]: series_to_year_dict(natgas_input_series, export_base_year, export_final_year)},
                series_to_year_dict(efficiency_liquefaction_series, export_base_year, export_final_year),
                liq_aux_values,
                liq_loss_values,
                liq_loss_total,
                feedstock_shares={fuel_codes["natural_gas"]: 1.0},
                input_total=natgas_input_total,
                output_import_targets=output_import_targets,
                output_export_targets=output_export_targets,
            )
            append_process_record(process_records, record)
            print(
                "Estimated liquefaction aux fuel use (electricity per PJ output): "
                f"{aux_fuel_use_liquefaction_series.mean():.6f}"
            )
            if regas_message:
                print(regas_message)
            print(f"Estimated liquefaction efficiency (mean): {efficiency_liquefaction_series.mean():.4f}")
    except Exception as exc:
        print(f"LNG analysis failed: {exc}")
        try_debug_breakpoint()
        raise


def analyze_gas_processing(
    esto_data,
    year_cols,
    start_year,
    economy,
    code_to_name_mapping,
    loss_data,
    loss_year_cols,
    sector_config=None,
    process_records=None,
):
    """Estimate efficiencies for gas works and natural gas blending plants."""
    try:
        if sector_config is None:
            raise ValueError("Gas processing analysis requires a sector_config")
        gas_config = sector_config
        fuel_codes = {
            "natural_gas": "08_01_natural_gas",
            "lng": "08_02_lng",
            "gas_works_gas": "08_03_gas_works_gas",
            "lignite": "01_05_lignite",
            "electricity": "17_electricity",
        }
        transformation_sub2 = gas_config.get("transformation_sub2") or []
        gas_works_sub2 = transformation_sub2[0] if len(transformation_sub2) > 0 else None
        blending_sub2 = transformation_sub2[1] if len(transformation_sub2) > 1 else None
        gas_works_flow_code = gas_config.get("flow_code_gas_works")
        blending_flow_code = gas_config.get("flow_code_blending")
        product_code_natural_gas = gas_config.get("product_code_natural_gas")
        product_code_gas_works_gas = gas_config.get("product_code_gas_works_gas")
        product_code_lignite = gas_config.get("product_code_lignite")
        print(f"\n==== Gas processing (no imports/exports expected) ({economy}) ====")
        if not has_required_columns(
            esto_data,
            [["sub2sectors", "subfuels", "fuels"], ["flows", "products"]],
            "Gas processing",
        ):
            return
        year_cols_from_start = get_years_from(year_cols, start_year)
        export_base_year = EXPORT_BASE_YEAR
        export_final_year = EXPORT_FINAL_YEAR
        export_years = list(range(export_base_year, export_final_year + 1))

        if "sub2sectors" in esto_data.columns and gas_works_sub2:
            gas_works_output = select_rows(
                esto_data,
                {
                    "economy": economy,
                    "sub2sectors": gas_works_sub2,
                    "subfuels": fuel_codes["gas_works_gas"],
                },
            )
            gas_works_input = select_rows(
                esto_data,
                {"economy": economy, "sub2sectors": gas_works_sub2, "subfuels": fuel_codes["lignite"]},
            )
            gas_works_rows = select_rows(
                esto_data,
                {"economy": economy, "sub2sectors": gas_works_sub2},
            )
        else:
            if gas_works_flow_code:
                gas_works_output = select_rows(
                    esto_data,
                    {
                        "economy": economy,
                        "flows": gas_works_flow_code,
                        "products": product_code_gas_works_gas,
                    },
                )
                gas_works_input = select_rows(
                    esto_data,
                    {
                        "economy": economy,
                        "flows": gas_works_flow_code,
                        "products": product_code_lignite,
                    },
                )
                gas_works_rows = select_rows(
                    esto_data,
                    {"economy": economy, "flows": gas_works_flow_code},
                )
            else:
                gas_works_output = esto_data.iloc[0:0]
                gas_works_input = esto_data.iloc[0:0]
                gas_works_rows = esto_data.iloc[0:0]
        print_sector_rows_from_df(
            gas_works_rows,
            "Gas works rows",
            year_cols,
            start_year,
            code_to_name_mapping,
        )

        gas_works_negatives, _ = summarize_fuels_by_subfuel(
            gas_works_rows, year_cols, start_year
        )
        gas_works_timeseries, _ = summarize_fuel_timeseries(
            gas_works_rows, year_cols, start_year, allow_all_years_fallback=False
        )
        gas_works_primary_label = fuel_codes["lignite"]
        if "products" in esto_data.columns:
            gas_works_primary_label = product_code_lignite
        gas_works_aux = split_auxiliary_fuels(
            gas_works_negatives,
            gas_works_primary_label,
            AUXILIARY_THRESHOLD_RATIO,
            INCLUDE_ALL_AUXILIARY,
        )
        gas_works_output_series = ensure_full_year_series(
            sum_years_by_year(gas_works_output, year_cols, start_year),
            export_base_year,
            export_final_year,
        )
        gas_works_input_series = ensure_full_year_series(
            sum_years_by_year(gas_works_input, year_cols, start_year).abs(),
            export_base_year,
            export_final_year,
        )
        gas_works_output_total = gas_works_output_series.sum()
        gas_works_input_total = gas_works_input_series.sum()
        gas_works_loss_series, gas_works_loss_total, gas_works_loss_values, gas_works_loss_values_by_year = build_loss_context(
            loss_data,
            loss_year_cols,
            start_year,
            economy,
            "gas_processing",
            gas_works_sub2,
        )
        gas_works_output_label_for_eff = fuel_codes["gas_works_gas"]
        if "products" in esto_data.columns:
            gas_works_output_label_for_eff = product_code_gas_works_gas
        gas_works_loss_total_for_eff = get_loss_total_for_efficiency_by_year(
            gas_works_loss_values_by_year,
            gas_works_primary_label,
            gas_works_output_label_for_eff,
            export_years,
        )
        gas_works_eff_series = compute_efficiency_by_year(
            gas_works_output_series,
            gas_works_input_series,
            gas_works_loss_total_for_eff,
        )
        gas_works_aux_ratios = build_auxiliary_ratios_by_year(
            gas_works_timeseries, gas_works_aux, gas_works_output_series
        )
        gas_works_aux, gas_works_aux_ratios = merge_loss_into_auxiliary_by_year(
            gas_works_aux,
            gas_works_aux_ratios,
            gas_works_loss_values_by_year,
            gas_works_output_series,
            gas_works_primary_label,
        )

        if not gas_works_rows.empty:
            output_label = fuel_codes["gas_works_gas"]
            feedstock_label = fuel_codes["lignite"]
            if "products" in esto_data.columns:
                output_label = product_code_gas_works_gas
                feedstock_label = product_code_lignite
            gas_works_loss_values = filter_loss_values_for_feedstock_by_year(
                gas_works_loss_values_by_year,
                feedstock_label,
            )
            gas_works_loss_values_display = {
                label: summarize_numeric_value(values, summary="sum")
                for label, values in gas_works_loss_values.items()
            }
            print_leap_structure_block(
                "Gas works",
                [output_label],
                "Gas works",
                [feedstock_label],
                gas_works_aux,
                loss_fuels=list(gas_works_loss_values.keys()),
                code_to_name_mapping=code_to_name_mapping,
                output_fuel_values={output_label: gas_works_output_total},
                process_value=f"{gas_works_eff_series.mean():.4f}",
                feedstock_fuel_values={feedstock_label: gas_works_input_total},
                auxiliary_fuel_values={
                    label: summarize_numeric_value(values, summary="mean")
                    for label, values in gas_works_aux_ratios.items()
                },
                loss_fuel_values=gas_works_loss_values_display,
            )
            output_import_targets, output_export_targets = gather_output_target_dicts(
                economy,
                [output_label],
                export_base_year,
                export_final_year,
            )
            record = build_process_record(
                economy,
                gas_config.get("title", "Gas works"),
                "Gas works",
                {output_label: series_to_year_dict(gas_works_output_series, export_base_year, export_final_year)},
                {feedstock_label: series_to_year_dict(gas_works_input_series, export_base_year, export_final_year)},
                series_to_year_dict(gas_works_eff_series, export_base_year, export_final_year),
                gas_works_aux_ratios,
                gas_works_loss_values,
                gas_works_loss_total,
                feedstock_shares={feedstock_label: 1.0},
                input_total=gas_works_input_total,
                output_import_targets=output_import_targets,
                output_export_targets=output_export_targets,
            )
            append_process_record(process_records, record)
            print(f"Estimated gas works efficiency (incl losses, mean): {gas_works_eff_series.mean():.4f}")
            print("=" * 45)

        if "sub2sectors" in esto_data.columns and blending_sub2:
            blending_output = select_rows(
                esto_data,
                {
                    "economy": economy,
                    "sub2sectors": blending_sub2,
                    "subfuels": fuel_codes["natural_gas"],
                },
            )
            blending_input = select_rows(
                esto_data,
                {
                    "economy": economy,
                    "sub2sectors": blending_sub2,
                    "subfuels": fuel_codes["gas_works_gas"],
                },
            )
            blending_rows = select_rows(
                esto_data,
                {"economy": economy, "sub2sectors": blending_sub2},
            )
        else:
            if blending_flow_code:
                blending_output = select_rows(
                    esto_data,
                    {
                        "economy": economy,
                        "flows": blending_flow_code,
                        "products": product_code_natural_gas,
                    },
                )
                blending_input = select_rows(
                    esto_data,
                    {
                        "economy": economy,
                        "flows": blending_flow_code,
                        "products": product_code_gas_works_gas,
                    },
                )
                blending_rows = select_rows(
                    esto_data,
                    {"economy": economy, "flows": blending_flow_code},
                )
            else:
                blending_output = esto_data.iloc[0:0]
                blending_input = esto_data.iloc[0:0]
                blending_rows = esto_data.iloc[0:0]
        print_sector_rows_from_df(
            blending_rows,
            "Natural gas blending rows",
            year_cols,
            start_year,
            code_to_name_mapping,
        )
        blending_negatives, _ = summarize_fuels_by_subfuel(
            blending_rows, year_cols, start_year
        )
        blending_timeseries, _ = summarize_fuel_timeseries(
            blending_rows, year_cols, start_year, allow_all_years_fallback=False
        )
        blending_primary_label = fuel_codes["gas_works_gas"]
        if "products" in esto_data.columns:
            blending_primary_label = product_code_gas_works_gas
        blending_aux = split_auxiliary_fuels(
            blending_negatives,
            blending_primary_label,
            AUXILIARY_THRESHOLD_RATIO,
            INCLUDE_ALL_AUXILIARY,
        )
        blending_output_series = ensure_full_year_series(
            sum_years_by_year(blending_output, year_cols, start_year),
            export_base_year,
            export_final_year,
        )
        blending_input_series = ensure_full_year_series(
            sum_years_by_year(blending_input, year_cols, start_year).abs(),
            export_base_year,
            export_final_year,
        )
        blending_output_total = blending_output_series.sum()
        blending_input_total = blending_input_series.sum()
        blending_loss_series, blending_loss_total, blending_loss_values, blending_loss_values_by_year = build_loss_context(
            loss_data,
            loss_year_cols,
            start_year,
            economy,
            "gas_processing",
            blending_sub2,
        )
        blending_output_label_for_eff = fuel_codes["natural_gas"]
        if "products" in esto_data.columns:
            blending_output_label_for_eff = product_code_natural_gas
        blending_loss_total_for_eff = get_loss_total_for_efficiency_by_year(
            blending_loss_values_by_year,
            blending_primary_label,
            blending_output_label_for_eff,
            export_years,
        )
        blending_eff_series = compute_efficiency_by_year(
            blending_output_series,
            blending_input_series,
            blending_loss_total_for_eff,
        )
        blending_aux_ratios = build_auxiliary_ratios_by_year(
            blending_timeseries, blending_aux, blending_output_series
        )
        blending_aux, blending_aux_ratios = merge_loss_into_auxiliary_by_year(
            blending_aux,
            blending_aux_ratios,
            blending_loss_values_by_year,
            blending_output_series,
            blending_primary_label,
        )

        if not blending_rows.empty:
            output_label = fuel_codes["natural_gas"]
            feedstock_label = fuel_codes["gas_works_gas"]
            if "products" in esto_data.columns:
                output_label = product_code_natural_gas
                feedstock_label = product_code_gas_works_gas
            blending_loss_values = filter_loss_values_for_feedstock_by_year(
                blending_loss_values_by_year,
                feedstock_label,
            )
            blending_loss_values_display = {
                label: summarize_numeric_value(values, summary="sum")
                for label, values in blending_loss_values.items()
            }
            print_leap_structure_block(
                "Natural gas blending",
                [output_label],
                "Natural gas blending",
                [feedstock_label],
                blending_aux,
                loss_fuels=list(blending_loss_values.keys()),
                code_to_name_mapping=code_to_name_mapping,
                output_fuel_values={output_label: blending_output_total},
                process_value=f"{blending_eff_series.mean():.6f}",
                feedstock_fuel_values={feedstock_label: blending_input_total},
                auxiliary_fuel_values={
                    label: summarize_numeric_value(values, summary="mean")
                    for label, values in blending_aux_ratios.items()
                },
                loss_fuel_values=blending_loss_values_display,
            )
            output_import_targets, output_export_targets = gather_output_target_dicts(
                economy,
                [output_label],
                export_base_year,
                export_final_year,
            )
            record = build_process_record(
                economy,
                gas_config.get("title", "Natural gas blending"),
                "Natural gas blending",
                {output_label: series_to_year_dict(blending_output_series, export_base_year, export_final_year)},
                {feedstock_label: series_to_year_dict(blending_input_series, export_base_year, export_final_year)},
                series_to_year_dict(blending_eff_series, export_base_year, export_final_year),
                blending_aux_ratios,
                blending_loss_values,
                blending_loss_total,
                feedstock_shares={feedstock_label: 1.0},
                input_total=blending_input_total,
                output_import_targets=output_import_targets,
                output_export_targets=output_export_targets,
            )
            append_process_record(process_records, record)
            print(f"Estimated natural gas blending efficiency (mean): {blending_eff_series.mean():.6f}")
            print("=" * 44)

        if PRINT_GAS_PROCESSING_SUMMARY:
            flow_list = [
                code for code in [gas_works_flow_code, blending_flow_code] if code
            ]
            gas_processing_rows = pd.concat(
                [select_flow_rows(esto_data, economy, code) for code in flow_list],
                ignore_index=True,
            ) if flow_list else esto_data.iloc[0:0]
            negatives, positives = summarize_fuels_by_subfuel(
                gas_processing_rows, year_cols, start_year
            )
            if not negatives.empty:
                print("Gas processing inputs by fuel label:")
                print(map_series_index(negatives, code_to_name_mapping).to_string())
            if not positives.empty:
                print("Gas processing outputs by fuel label:")
                print(map_series_index(positives, code_to_name_mapping).to_string())
    except Exception as exc:
        print(f"Gas processing analysis failed: {exc}")
        try_debug_breakpoint()
        raise


def summarize_transformation_flows(
    data,
    year_cols,
    start_year,
    economy,
    flow_codes,
    title,
    code_to_name_mapping,
    loss_data,
    loss_year_cols,
    sector_key,
    process_records=None,
):
    """Summarize transformation flows with primary input/output fuels."""
    try:
        print(f"\n==== {title} ({economy}) ====")
        if not has_required_columns(
            data,
            [["flows", "products"], ["flows", "subfuels", "fuels"]],
            title,
        ):
            return
        flow_list = get_flow_list(data, flow_codes)
        if not flow_list:
            print(f"{title}: no flows configured or found")
            return

        for flow_code in flow_list:
            flow_rows = select_flow_rows(data, economy, flow_code)
            if flow_rows.empty:
                continue

            totals, used_all_years = summarize_fuel_totals(
                flow_rows, year_cols, start_year, allow_all_years_fallback=True
            )
            timeseries, _ = summarize_fuel_timeseries(
                flow_rows, year_cols, start_year, allow_all_years_fallback=True
            )
            # Negative totals represent feedstocks/own-use inputs, while positives are outputs.
            negative = totals[totals < 0]
            positive = totals[totals > 0]

            if negative.empty or positive.empty:
                print(f"{flow_code}: missing input/output balance for {start_year}+ and all years")
                continue
            if used_all_years:
                print(f"{flow_code}: no {start_year}+ activity, using all years for summary")

            primary_input, primary_output, input_total, output_total = compute_primary_io(
                negative, positive
            )
            # Loss rows come in negative, but build_loss_context returns the absolute values we feed to LEAP.
            loss_series, loss_total, loss_values, loss_values_by_year = build_loss_context(
                loss_data,
                loss_year_cols,
                start_year,
                economy,
                sector_key,
                flow_code=flow_code,
            )
            export_base_year = EXPORT_BASE_YEAR
            export_final_year = EXPORT_FINAL_YEAR
            export_years = list(range(export_base_year, export_final_year + 1))
            # Outputs should already be positive, but we still enforce the year range here.
            output_series = ensure_full_year_series(
                get_label_timeseries(timeseries, primary_output),
                export_base_year,
                export_final_year,
            )
            # Feedstock rows are negative in the balance table, so we take absolutes before building LEAP series.
            input_series = ensure_full_year_series(
                get_label_timeseries(timeseries, primary_input).abs(),
                export_base_year,
                export_final_year,
            )
            loss_total_for_eff = get_loss_total_for_efficiency_by_year(
                loss_values_by_year,
                primary_input,
                primary_output,
                export_years,
            )
            efficiency_series = compute_efficiency_by_year(
                output_series,
                input_series,
                loss_total_for_eff,
            )
            input_name = map_code_label(primary_input, code_to_name_mapping)
            output_name = map_code_label(primary_output, code_to_name_mapping)
            other_feedstock_fuels = []
            if INCLUDE_ALL_FEEDSTOCKS_AS_AUXILIARY:
                other_feedstock_fuels = get_all_other_negative_fuels(
                    negative,
                    primary_input,
                )
                auxiliary_fuels, auxiliary_ratios = build_auxiliary_from_losses_by_year(
                    loss_values_by_year,
                    output_series,
                )
            else:
                auxiliary_fuels = split_auxiliary_fuels(
                    negative,
                    primary_input,
                    AUXILIARY_THRESHOLD_RATIO,
                    INCLUDE_ALL_AUXILIARY,
                )
                auxiliary_ratios = build_auxiliary_ratios_by_year(
                    timeseries, auxiliary_fuels, output_series
                )
                auxiliary_fuels, auxiliary_ratios = merge_loss_into_auxiliary_by_year(
                    auxiliary_fuels,
                    auxiliary_ratios,
                    loss_values_by_year,
                    output_series,
                    primary_input,
                )
            other_feedstock_values = {}
            other_feedstock_ratios = {}
            if INCLUDE_ALL_FEEDSTOCKS_AS_AUXILIARY:
                other_feedstock_values = {
                    label: abs(negative.get(label))
                    for label in other_feedstock_fuels
                    if label in negative.index
                }
                other_feedstock_ratios = build_auxiliary_ratios_by_year(
                    timeseries, other_feedstock_fuels, output_series
                )

            print(
                f"{flow_code}: output {output_name} ({output_total:.2f}), "
                f"input {input_name} ({-input_total:.2f}), "
                f"eff {efficiency_series.mean():.4f}"
            )

            flow_loss_values = filter_loss_values_for_feedstock(
                loss_values,
                primary_input,
            )
            print_leap_structure_block(
                f"{title} - {flow_code}",
                [primary_output],
                flow_code,
                [primary_input],
                auxiliary_fuels,
                loss_fuels=list(flow_loss_values.keys()),
                code_to_name_mapping=code_to_name_mapping,
                output_fuel_values={primary_output: output_total},
                process_value=f"{efficiency_series.mean():.4f}",
                feedstock_fuel_values={primary_input: input_total},
                auxiliary_fuel_values={
                    label: summarize_numeric_value(values, summary="mean")
                    for label, values in auxiliary_ratios.items()
                },
                loss_fuel_values=flow_loss_values,
                other_feedstock_fuels=(
                    other_feedstock_fuels if INCLUDE_ALL_FEEDSTOCKS_AS_AUXILIARY else None
                ),
                other_feedstock_values=(
                    other_feedstock_values if INCLUDE_ALL_FEEDSTOCKS_AS_AUXILIARY else None
                ),
                other_feedstock_ratios=(
                    {
                        label: summarize_numeric_value(values, summary="mean")
                        for label, values in other_feedstock_ratios.items()
                    } if INCLUDE_ALL_FEEDSTOCKS_AS_AUXILIARY else None
                ),
            )

            output_import_targets, output_export_targets = gather_output_target_dicts(
                economy,
                [primary_output],
                export_base_year,
                export_final_year,
            )
            # breakpoint()

            record = build_process_record(
                economy,
                title,
                flow_code,
                {primary_output: series_to_year_dict(output_series, export_base_year, export_final_year)},
                {primary_input: series_to_year_dict(input_series, export_base_year, export_final_year)},
                series_to_year_dict(efficiency_series, export_base_year, export_final_year),
                auxiliary_ratios,
                loss_values_by_year,
                loss_total,
                feedstock_shares={primary_input: 1.0},
                input_total=input_total,
                output_import_targets=output_import_targets,
                output_export_targets=output_export_targets,
            )
            append_process_record(process_records, record)

            print_sector_rows_from_df(
                flow_rows,
                f"{title} rows ({flow_code})",
                year_cols,
                start_year,
                code_to_name_mapping,
            )

            negatives, positives = summarize_fuels_by_subfuel(
                flow_rows, year_cols, start_year
            )
            if not negatives.empty:
                print("Inputs by fuel label:")
                print(map_series_index(negatives, code_to_name_mapping).to_string())
            if not positives.empty:
                print("Outputs by fuel label:")
                print(map_series_index(positives, code_to_name_mapping).to_string())
    except Exception as exc:
        print(f"{title} flow analysis failed: {exc}")
        try_debug_breakpoint()
        raise
#%%

#%%
def run_analysis_for_sector(run_flag, sector_key, analysis_callback, process_records):
    """Resolve dataset + economies and execute an analysis callback."""
    try:
        if not run_flag:
            return
        sector_config = dict(resolve_sector_config(sector_key))
        sector_config["sector_key"] = sector_key
        data, year_cols = resolve_dataset(DATASET_MAP, sector_config["dataset_key"])
        loss_data, loss_year_cols = data, year_cols
        for economy in get_economy_list(data, ECONOMIES_TO_ANALYZE):
            analysis_callback(
                data,
                year_cols,
                economy,
                loss_data,
                loss_year_cols,
                sector_config,
                process_records,
            )
    except Exception as exc:
        print(f"Analysis runner failed for {sector_key}: {exc}")
        try_debug_breakpoint()
        raise
#%%

#%%
def resolve_sector_config(sector_key):
    """Return sector config for a sector key."""
    try:
        return MAJOR_SECTOR_CONFIG[sector_key]
    except Exception as exc:
        print(f"Failed to resolve config for {sector_key}: {exc}")
        try_debug_breakpoint()
        raise


def run_lng_analysis(
    data,
    year_cols,
    economy,
    loss_data,
    loss_year_cols,
    sector_config,
    process_records,
):
    """Run LNG analysis for a single economy."""
    analyze_lng_liquefaction_regas(
        data,
        year_cols,
        YEAR_START_FOR_ANALYSIS,
        economy,
        code_to_name_mapping,
        loss_data,
        loss_year_cols,
        sector_config,
        process_records,
    )


def run_gas_processing_analysis(
    data, year_cols, economy, loss_data, loss_year_cols, sector_config, process_records
):
    """Run gas processing analysis for a single economy."""
    analyze_gas_processing(
        data,
        year_cols,
        YEAR_START_FOR_ANALYSIS,
        economy,
        code_to_name_mapping,
        loss_data,
        loss_year_cols,
        sector_config,
        process_records,
    )


def run_flow_sector_analysis(
    data, year_cols, economy, loss_data, loss_year_cols, sector_config, process_records
):
    """Run a flow-based transformation analysis for a single economy."""
    summarize_transformation_flows(
        data,
        year_cols,
        YEAR_START_FOR_ANALYSIS,
        economy,
        sector_config.get("transformation_flow_codes"),
        sector_config.get("title", "Transformation"),
        code_to_name_mapping,
        loss_data,
        loss_year_cols,
        sector_config.get("sector_key", ""),
        process_records,
    )


def run_coal_transformation_analysis(
    data, year_cols, economy, loss_data, loss_year_cols, sector_config, process_records
):
    """Run coal transformation analysis for a single economy."""
    run_flow_sector_analysis(
        data,
        year_cols,
        economy,
        loss_data,
        loss_year_cols,
        sector_config,
        process_records,
    )


def run_charcoal_processing_analysis(
    data, year_cols, economy, loss_data, loss_year_cols, sector_config, process_records
):
    """Run charcoal processing analysis for a single economy."""
    run_flow_sector_analysis(
        data,
        year_cols,
        economy,
        loss_data,
        loss_year_cols,
        sector_config,
        process_records,
    )


def run_nonspecified_transformation_analysis(
    data, year_cols, economy, loss_data, loss_year_cols, sector_config, process_records
):
    """Run nonspecified transformation analysis for a single economy."""
    run_flow_sector_analysis(
        data,
        year_cols,
        economy,
        loss_data,
        loss_year_cols,
        sector_config,
        process_records,
    )
#%%

#%%
######### CONSTANTS (LIKELY TO CHANGE) #########
RUN_LNG_ANALYSIS = True
RUN_GAS_PROCESSING_ANALYSIS = True
RUN_COAL_TRANSFORMATION_ANALYSIS = True
RUN_CHARCOAL_PROCESSING_ANALYSIS = True
RUN_NONSPECIFIED_TRANSFORMATION_ANALYSIS = True
INCLUDE_ALL_ECONOMIES = True
ALL_ECONOMY_LABEL = "20_USA"
INCLUDE_ALL_FEEDSTOCKS_AS_AUXILIARY = True
ECONOMIES_TO_ANALYZE = [ALL_ECONOMY_LABEL]
SAVE_ESTO_SUBTOTAL_LABELED = True
ESTO_SUBTOTAL_LABELED_OUTPUT_PATH = "data/00APEC_2024_low_with_subtotals.csv"
BUILD_LEAP_EXPORT = True
SAVE_LEAP_EXPORT_FILE = True
SAVE_SUMMARY_TABLES = True
EXPORT_OUTPUT_DIR = os.path.join("outputs", "leap_exports")
EXPORT_FILENAME_FALLBACK = "transformation_leap_imports.xlsx"
EXPORT_FILENAME_TEMPLATE = "transformation_leap_imports_{economy}_{scenario}.xlsx"
EXPORT_MODEL_NAME = "LEAP Transformation Imports"
EXPORT_REGION = "United States of America"
SCENARIOS_TO_EXPORT = ['Reference','Target','Current Accounts']
EXPORT_BASE_YEAR = 2022
EXPORT_FINAL_YEAR = PROJECTION_END_YEAR
SUMMARY_OUTPUT_DIR = EXPORT_OUTPUT_DIR
PROCESS_SUMMARY_FILENAME = "transformation_process_summary.csv"
DETAIL_SUMMARY_FILENAME = "transformation_detail_summary.csv"

INCLUDE_OUTPUT_SERIES_IN_LEAP_EXPORT = False  # Skip emitting output series rows that LEAP does not expect.

SAVE_PROJECTION_DIAGNOSTICS = False
PROJECTION_DIAGNOSTICS_PATH = os.path.join(
    EXPORT_OUTPUT_DIR, "ninth_projection_allocation_fallbacks.csv"
)

SCENARIO_EXPORT_OVERRIDES = {
    "Current Account": {
        "export_base_year": 2022,
        "export_final_year": 2022,
        "include_current_account_rows": True,
    }
}


def get_scenario_export_config(scenario):
    """Return export overrides for the given scenario."""
    overrides = SCENARIO_EXPORT_OVERRIDES.get(scenario, {})
    return {
        "export_base_year": overrides.get("export_base_year", EXPORT_BASE_YEAR),
        "export_final_year": overrides.get("export_final_year", EXPORT_FINAL_YEAR),
        "include_current_account_rows": overrides.get("include_current_account_rows", False),
    }


def compute_combined_year_range(base_year, final_year, scenario_configs):
    """Return the min base year and max final year over all scenarios."""
    base_candidates = [base_year]
    final_candidates = [final_year]
    for cfg in scenario_configs.values():
        base_candidates.append(cfg.get("export_base_year", base_year))
        final_candidates.append(cfg.get("export_final_year", final_year))
    return min(base_candidates), max(final_candidates)
#%%

#%%
######### LOAD DATA #########
ensure_repo_root()
ninth_data_raw = load_csv_data(ESTO_DATA_PATH, "ESTO data (9th)")
esto_data_raw = load_csv_data(MATT_DATA_PATH, "Matt data")
 # Note: matt data lacks sub-sector columns; keep available for dataset_key="matt" only.

ninth_data_raw, ninth_year_cols = normalize_year_columns(ninth_data_raw)
esto_data_raw, esto_year_cols = normalize_year_columns(esto_data_raw)
esto_year_cols_raw = list(esto_year_cols)

ninth_data = clean_esto_subtotals(ninth_data_raw, ninth_year_cols)
ninth_data = filter_reference_scenario(ninth_data, "9th data")
if "subtotal_results" in ninth_data.columns:
    ninth_data = ninth_data[ninth_data["subtotal_results"] == False].copy()
esto_data_raw = normalize_esto_economy_codes(esto_data_raw)
esto_data_raw = filter_total_energy_rows(esto_data_raw)
ninth_data = filter_total_energy_rows(ninth_data)
esto_data_with_subtotals = apply_matt_subtotal_mapping(esto_data_raw, SUBTOTAL_MAPPING_PATH)
if SAVE_ESTO_SUBTOTAL_LABELED:
    save_subtotal_labeled_data(
        esto_data_with_subtotals,
        ESTO_SUBTOTAL_LABELED_OUTPUT_PATH,
        "ESTO (Matt) data",
    )
esto_data = filter_matt_subtotals(esto_data_with_subtotals)
if INCLUDE_ALL_ECONOMIES:
    ninth_data = add_all_economy_total(ninth_data, ninth_year_cols, ALL_ECONOMY_LABEL)
    esto_data = add_all_economy_total(esto_data, esto_year_cols, ALL_ECONOMY_LABEL)
projection_df, projection_diagnostics = build_esto_projection_table(
    ninth_data=ninth_data,
    esto_data=esto_data,
    mapping_path=NINTH_TO_ESTO_MAPPING_PATH,
    base_year=BASE_YEAR,
    projection_years=PROJECTION_YEAR_RANGE,
)
esto_data = merge_projection_into_esto(
    esto_data, projection_df, PROJECTION_YEAR_RANGE
)
esto_year_cols = sorted([col for col in esto_data.columns if str(col).isdigit()])
if SAVE_PROJECTION_DIAGNOSTICS and projection_diagnostics is not None:
    if not projection_diagnostics.empty:
        os.makedirs(EXPORT_OUTPUT_DIR, exist_ok=True)
        projection_diagnostics.to_csv(PROJECTION_DIAGNOSTICS_PATH, index=False)
        print(f"Saved projection fallback report to {PROJECTION_DIAGNOSTICS_PATH}")
code_to_name_mapping = (
    load_code_to_name_mapping(CODE_TO_NAME_PATHS) if USE_CODE_TO_NAME_MAPPING else {}
)
DATASET_MAP = build_dataset_map(
    esto_data,
    esto_year_cols,
    ninth_data,
    ninth_year_cols,
    esto_data_raw,
    esto_year_cols_raw,
)
ESTO_IMPORT_EXPORT_REFERENCE_DATA = esto_data
ESTO_IMPORT_EXPORT_YEAR_COLS = esto_year_cols
#%%

#%%
# Transformation\NG Liquefaction\Processes\Regasification\Auxiliary Fuels\Electricity Auxiliary Fuel Use Current Accounts United States of America Gigajoule Gigajoule 0.00428 Transformation NG Liquefaction Processes Regasification Auxiliary Fuels Electricity
# Transformation\NG Liquefaction\Processes\Regasification\Auxiliary Fuels\Electricity Auxiliary Fuel Use Reference United States of America Gigajoule Gigajoule 0.00428 Transformation NG Liquefaction Processes Regasification Auxiliary Fuels Electricity
# Transformation\NG Liquefaction\Processes\Regasification\Auxiliary Fuels\Electricity Auxiliary Fuel Use Target United States of America Gigajoule Gigajoule 0.00428 Transformation NG Liquefaction Processes Regasification Auxiliary Fuels Electricity
#%%

#%%
# LNG done basically. Next up is the other gas processing sectors:
# Using ESTO data here (Matt file does not include sector detail columns).
#%%
