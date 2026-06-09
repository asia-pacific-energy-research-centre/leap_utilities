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

from codebase.utilities.master_config import config_table_exists, read_config_table

# Allow the repository root to be importable regardless of the working directory.
REPO_ROOT = Path(__file__).resolve().parents[2]
CURRENT_DIR = Path.cwd()
if CURRENT_DIR != REPO_ROOT:
    os.chdir(REPO_ROOT)
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.configuration.all_products_and_flows import ESTO_SECTORS
from codebase.configuration.config import (
    BRANCH_DEMAND_CATEGORY,
    BRANCH_DEMAND_TECHNOLOGY,
    BASE_YEAR,
)
from codebase.functions.leap_core import (
    connect_to_leap,
    create_branches_from_export_file,
    fill_branches_from_export_file,
    sanitize_leap_branch_path,
)
from codebase.scrapbook.utilities import (
    apply_matt_subtotal_mapping,
    filter_matt_subtotals,
    load_augmented_reference_tables,
    save_subtotal_labeled_data,
)
from codebase.functions.leap_excel_io import finalise_export_df, save_export_files
from codebase.functions.ninth_projection_mapping import (
    build_esto_projection_table,
    merge_projection_into_esto,
)
from codebase.configuration import workflow_config as workflow_cfg
from codebase.utilities import workflow_common
#%%

#%%
######### CONSTANTS (UNLIKELY TO CHANGE) #########
ESTO_DATA_PATH = REPO_ROOT / "data" / "merged_file_energy_ALL_20250814_pre_trump.csv"
# Use merged_file_energy_ALL_20251106.csv and merged_file_energy_00_APEC_20251106 for exact 9th edition projection matching.
MATT_DATA_PATH = REPO_ROOT / "data" / "00APEC_2024_low.csv"
CONFIG_DIR = REPO_ROOT / "config"
SUBTOTAL_MAPPING_PATH = CONFIG_DIR / "ESTO_subtotal_mapping.xlsx"
NINTH_TO_ESTO_MAPPING_PATH = CONFIG_DIR / "ninth_pairs_to_esto_pairs.xlsx"
CODE_TO_NAME_PATHS = [
    CONFIG_DIR / "sector_fuel_codes_to_names.updated.xlsx",
    CONFIG_DIR / "sector_fuel_codes_to_names.xlsx",
]

YEAR_START_FOR_ANALYSIS = BASE_YEAR
PROJECTION_START_YEAR = 2023
PROJECTION_END_YEAR = 2060
PROJECTION_YEAR_RANGE = list(range(PROJECTION_START_YEAR, PROJECTION_END_YEAR + 1))
REFERENCE_CACHE_DIR = REPO_ROOT / "data" / ".cache" / "transformation_reference_tables"
# Projection split policy for transformation workflows:
# - PROJECTION_SIGN_STABLE_MODE controls allocation behavior:
#   - "all": apply sign-stable routing to every mapped ESTO flow.
#   - "selected": apply only to SIGN_STABLE_PROJECTION_FLOWS.
#   - "off": disable sign-stable routing (legacy abs-share behavior).
# - PROJECTION_STRICT_CONSERVATION raises on any source-vs-allocated mismatch.
PROJECTION_SIGN_STABLE_MODE = "all"
SIGN_STABLE_PROJECTION_FLOWS = [
    "09.08.01 Coke ovens",
    "09.08.02 Blast furnaces",
    "09.08.03 Patent fuel plants",
    "09.08.04 BKB/PB plants",
    "09.08.05 Liquefaction (coal to oil)",
    "09.08.06 Coal mines",
]
PROJECTION_STRICT_CONSERVATION = True


def resolve_projection_sign_stable_flows(mode, selected_flows):
    """Return sign_stable_flows argument for build_esto_projection_table."""
    mode_key = str(mode or "").strip().lower()
    if mode_key in {"off", "none", ""}:
        return []
    if mode_key == "selected":
        return list(selected_flows or [])
    if mode_key == "all":
        return "all"
    raise ValueError(
        f"Invalid PROJECTION_SIGN_STABLE_MODE={mode!r}. "
        "Use one of: 'all', 'selected', 'off'."
    )
LOSS_SECTOR_CODE_9TH = "10_losses_and_own_use"
# DEFAULT_SCENARIO = "Target"
DEFAULT_REGION = "United States"
DEFAULT_OUTPUT_UNITS = "Petajoule"
DEFAULT_EFFICIENCY_UNITS = "Percent"
DEFAULT_FEEDSTOCK_UNITS = ""
DEFAULT_AUXILIARY_UNITS = "Petajoule"
DEFAULT_AUXILIARY_PER = "Petajoule"
ENABLE_DEBUG_BREAKPOINTS = True
PRINT_SECTOR_ROWS = True
PRINT_TOP_FUEL_ROWS = 12
PRINT_ONLY_NONZERO_ROWS = True
USE_CODE_TO_NAME_MAPPING = True
AUXILIARY_THRESHOLD_RATIO = 0.1
INCLUDE_ALL_AUXILIARY = False
PRINT_GAS_PROCESSING_SUMMARY = False

HYDROGEN_OUTPUT_SUBFUELS = [
    "16_x_hydrogen",
    "16_x_ammonia",
    "16_x_efuel",
]
HYDROGEN_PROCESS_CONFIG = [
    {
        "process_code": "09_13_01_electrolysers",
        "source_sub2sectors": "09_13_01_electrolysers",
        "input_fuel_codes": ["17_x_green_electricity"],
        "output_subfuels": list(HYDROGEN_OUTPUT_SUBFUELS),
        "enabled": True,
    },
    {
        "process_code": "09_13_03_smr_w_ccs",
        "source_sub2sectors": "09_13_03_smr_w_ccs",
        "input_fuel_codes": ["08_01_natural_gas"],
        "output_subfuels": list(HYDROGEN_OUTPUT_SUBFUELS),
        "enabled": True,
    },
    {
        "process_code": "09_13_02_smr_wo_ccs",
        "source_sub2sectors": "09_13_02_smr_wo_ccs",
        "input_fuel_codes": ["08_01_natural_gas"],
        "output_subfuels": list(HYDROGEN_OUTPUT_SUBFUELS),
        "enabled": True,
    },
    {
        "process_code": "electrolysers_non_green",
        "source_sub2sectors": "09_13_01_electrolysers",
        "input_fuel_codes": ["17_electricity"],
        "output_subfuels": list(HYDROGEN_OUTPUT_SUBFUELS),
        "enabled": False,
    },
]

MAJOR_SECTOR_CONFIG = {
    "lng": {
        "dataset_key": "ninth",
        "title": "NG Liquefaction",
        "liquefaction_title": "NG Liquefaction",
        "regasification_title": "LNG regasification",
        "transformation_sub1": "09_06_gas_processing_plants",
        "transformation_sub2": ["09_06_02_liquefaction_regasification_plants"],
        "loss_sub2": ["10_01_03_liquefaction_regasification_plants"],
    },
    "gas_works": {
        "dataset_key": "esto",
        "title": "Gas works plants",
        "transformation_sub2": [
            "09_06_01_gas_works_plants",
            "09_06_03_natural_gas_blending_plants",
        ],
        "flow_code_gas_works": "09.06.01 Gas works plants",
        "flow_code_blending": "09.06.03 Natural gas blending plants",
        "product_code_natural_gas": "08.01 Natural gas",
        "product_code_gas_works_gas": "08.03 Gas works gas",
        "product_code_lignite": "01.05 Lignite",
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
        "title": "BKB and PB plants",
        "transformation_flow_codes": ["09.08.04 BKB/PB plants"],
        "loss_flow_codes": [],
    },
    "coal_liquefaction": {
        "dataset_key": "esto",
        "title": "Liquefaction (coal to oil)",
        "transformation_flow_codes": ["09.08.05 Liquefaction (coal to oil)"],
        "loss_flow_codes": [],
    },
    "electric_boilers": {
        "dataset_key": "esto",
        "title": "Electric boilers",
        "transformation_flow_codes": ["09.04 Electric boilers"],
        "loss_flow_codes": [],
    },
    "chemical_heat_for_electricity_production": {
        "dataset_key": "esto",
        "title": "Chemical heat for electricity production",
        "transformation_flow_codes": ["09.05 Chemical heat for electricity production"],
        "loss_flow_codes": [],
    },
    "petrochemical_industry": {
        "dataset_key": "esto",
        "title": "Petrochemical industry",
        "transformation_flow_codes": ["09.09 Petrochemical industry"],
        "loss_flow_codes": [],
    },
    "gas_to_liquids_plants": {
        "dataset_key": "esto",
        "title": "Gas to liquids plants",
        "transformation_flow_codes": ["09.06.04 Gas-to-liquids plants"],
        "loss_flow_codes": [],
    },
    "biofuels_processing": {
        "dataset_key": "esto",
        "title": "Biofuels processing",
        "transformation_flow_codes": ["09.10 Biofuels processing"],
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
    # Oil refineries: multi-output sector (gasoline, diesel, jet fuel, etc.).
    # run_flow_sector_analysis handles the primary output only — a full multi-output
    # analysis (similar to hydrogen) is needed to capture all product streams.
    "oil_refineries": {
        "dataset_key": "esto",
        "title": "Oil Refining",
        "transformation_flow_codes": ["09.07 Oil refineries"],
        "loss_flow_codes": ["10.01.11 Oil refineries"],
    },
    "hydrogen_transformation": {
        "dataset_key": "ninth",
        "title": "Hydrogen transformation",
        "transformation_sub1": "09_13_hydrogen_transformation",
        "output_fuels": ["16_others"],
        "output_subfuels": list(HYDROGEN_OUTPUT_SUBFUELS),
        "process_config": HYDROGEN_PROCESS_CONFIG,
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
FEEDSTOCK_METHOD_SINGLE_AUX = "single_feedstock_aux_others"
FEEDSTOCK_METHOD_SPLIT = "split_processes_per_feedstock"
FEEDSTOCK_METHOD_MULTI = "multi_feedstock_single_process"
FEEDSTOCK_METHOD = FEEDSTOCK_METHOD_MULTI
# When False, loss/own-use fuels always go to aux even if they match a feedstock fuel.
# When True, loss fuels that match a feedstock label are excluded from aux.
LOSS_AUX_EXCLUDE_FEEDSTOCKS = False
# When True, input fuels detected as negative in the source data are always written to the
# export even if their values sum to zero over the export year range (e.g. data only exists
# outside the export window, such as pre-2022 historical lignite for AUS BKB/PB).
# Setting False drops these zero-sum fuels — if all inputs are dropped the process is
# skipped entirely, which is correct for sectors with no export-window activity.
# Root cause of phantom rows: the all-years fallback in summarize_fuel_totals can detect
# historical fuels as "inputs" even when the sector has zero activity in 2022+.
INCLUDE_ALL_DETECTED_INPUT_FUELS = False

_dropped_input_fuel_log: list[dict] = []
# Tracks LEAP-mapped sector titles for every sector run_analysis_for_sector visits,
# even those where all economies returned no data.  Used by zero-fill to clear catalog
# branches for sectors this workflow owns but had no ESTO data for.
_analyzed_sector_titles: set[str] = set()
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
    """Move to repo root so relative workflow paths resolve consistently."""
    try:
        if Path.cwd() != REPO_ROOT:
            os.chdir(REPO_ROOT)
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
        df = read_config_table(path)
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
    """Drop total/aggregate summary rows from fuels or products.

    Removes:
    - Named total rows (19 Total, 20 Total Renewables, 21 Modern renewables)
    - ESTO aggregate product codes that have no decimal point (e.g. '08 Gas',
      '01 Coal', '07 Petroleum products'). These are subtotals of their
      sub-product groups and should never appear as LEAP branch fuels.
    """
    try:
        import re as _re
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
            # Drop named totals
            mask_named = updated["products"].astype(str).isin(total_labels)
            # Drop aggregate codes: two-digit number + space + name, no decimal
            # e.g. "08 Gas", "01 Coal", "07 Petroleum products"
            mask_aggregate = updated["products"].astype(str).str.match(
                r"^\d{2} ", na=False
            ) & ~updated["products"].astype(str).str.contains(r"\.", na=False)
            updated = updated[~(mask_named | mask_aggregate)]
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
    output_series=None,
    cap_to_output=False,
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
        if output_series is not None and not output_series.empty:
            output_series = ensure_full_year_series(output_series.abs(), base_year, final_year)
            base_output = output_series.get(base_year, 0.0)
            base_export = full_series.get(base_year, 0.0)
            if base_output > 0 and base_export > 0:
                share = base_export / base_output
                max_export_year = max(series.index) if len(series.index) else base_year
                for year in range(base_year, final_year + 1):
                    if year <= max_export_year:
                        continue
                    if full_series.get(year, 0.0) == 0 and output_series.get(year, 0.0) != 0:
                        full_series[year] = output_series.get(year, 0.0) * share
            if cap_to_output:
                for year in range(base_year, final_year + 1):
                    export_value = float(full_series.get(year, 0.0))
                    output_value = max(float(output_series.get(year, 0.0)), 0.0)
                    if export_value > output_value:
                        full_series[year] = output_value
        return series_to_year_dict(full_series, base_year, final_year)
    except Exception as exc:
        print(f"Failed to build import/export target dict for {fuel_label}: {exc}")
        try_debug_breakpoint()
        raise


def gather_output_target_dicts(economy, output_labels, base_year, final_year, output_series_by_fuel=None):
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
            output_series = None
            if output_series_by_fuel:
                output_series = output_series_by_fuel.get(label)
            if TRANSFORMATION_OUTPUT_VARIABLES.get("output_import_target"):
                import_dict = build_est_output_target_dict(
                    economy,
                    label,
                    ESTO_IMPORT_SECTOR_LABEL,
                    YEAR_START_FOR_ANALYSIS,
                    base_year,
                    final_year,
                    output_series=output_series,
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
                    output_series=output_series,
                    cap_to_output=True,
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


def to_input_only_series(series):
    """Return the input-only component of a signed series (negative values as positives)."""
    try:
        if series is None:
            return pd.Series(dtype=float)
        return series.clip(upper=0).abs()
    except Exception as exc:
        print(f"Failed to build input-only series: {exc}")
        try_debug_breakpoint()
        raise


def to_output_only_series(series):
    """Return the output-only component of a signed series (positive values only)."""
    try:
        if series is None:
            return pd.Series(dtype=float)
        return series.clip(lower=0)
    except Exception as exc:
        print(f"Failed to build output-only series: {exc}")
        try_debug_breakpoint()
        raise


def build_auxiliary_ratios_by_year(negative_timeseries, auxiliary_fuels, output_series):
    """Return auxiliary fuel ratios by year for each auxiliary fuel."""
    try:
        ratios = {}
        if negative_timeseries is None or output_series is None:
            return ratios
        output_series_pos = to_output_only_series(output_series)
        for label in auxiliary_fuels:
            if label not in negative_timeseries.index:
                continue
            aux_input_only = to_input_only_series(negative_timeseries.loc[label])
            ratio_series = safe_divide_series(
                aux_input_only,
                output_series_pos,
            )
            ratios[label] = ratio_series.to_dict()
        return ratios
    except Exception as exc:
        print(f"Failed to build auxiliary ratios by year: {exc}")
        try_debug_breakpoint()
        raise


def build_auxiliary_from_losses_by_year(loss_values_by_year, output_series, feedstock_labels=None):
    """Return auxiliary fuels and ratios by year derived from losses, excluding feedstock fuels."""
    try:
        if not loss_values_by_year:
            return [], {}
        exclude = {str(label) for label in (feedstock_labels or [])}
        output_series_pos = to_output_only_series(output_series)
        fuels = []
        ratios = {}
        for label, series in loss_values_by_year.items():
            if str(label) in exclude:
                continue
            fuels.append(label)
            ratio_series = safe_divide_series(pd.Series(series).abs(), output_series_pos)
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
        output_series_pos = to_output_only_series(output_series)
        updated_fuels = list(auxiliary_fuels) if auxiliary_fuels else []
        updated_ratios = dict(auxiliary_ratios) if auxiliary_ratios else {}
        for label, series in loss_values_by_year.items():
            if label == feedstock_label:
                continue
            if label not in updated_fuels:
                updated_fuels.append(label)
            ratio_series = safe_divide_series(pd.Series(series).abs(), output_series_pos)
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
        output_series_pos = to_output_only_series(output_series)
        denom = input_series.add(loss_series, fill_value=0.0)
        denom = denom.clip(lower=0.0)
        return safe_divide_series(output_series_pos, denom)
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


def resolve_feedstock_method(method=None):
    """Return a normalized feedstock handling method."""
    try:
        raw = FEEDSTOCK_METHOD if method is None else method
        text = str(raw or "").strip().lower()
        mapping = {
            FEEDSTOCK_METHOD_SINGLE_AUX: FEEDSTOCK_METHOD_SINGLE_AUX,
            "single": FEEDSTOCK_METHOD_SINGLE_AUX,
            "single_feedstock": FEEDSTOCK_METHOD_SINGLE_AUX,
            "single_feedstock_aux": FEEDSTOCK_METHOD_SINGLE_AUX,
            "single_feedstock_aux_others": FEEDSTOCK_METHOD_SINGLE_AUX,
            FEEDSTOCK_METHOD_SPLIT: FEEDSTOCK_METHOD_SPLIT,
            "split": FEEDSTOCK_METHOD_SPLIT,
            "split_processes": FEEDSTOCK_METHOD_SPLIT,
            "split_processes_per_feedstock": FEEDSTOCK_METHOD_SPLIT,
            "per_feedstock": FEEDSTOCK_METHOD_SPLIT,
            FEEDSTOCK_METHOD_MULTI: FEEDSTOCK_METHOD_MULTI,
            "multi": FEEDSTOCK_METHOD_MULTI,
            "multi_feedstock": FEEDSTOCK_METHOD_MULTI,
            "multi_feedstock_single_process": FEEDSTOCK_METHOD_MULTI,
        }
        if text in mapping:
            return mapping[text]
        raise ValueError(
            f"Invalid feedstock method '{raw}'. "
            f"Use one of: {FEEDSTOCK_METHOD_SINGLE_AUX}, {FEEDSTOCK_METHOD_SPLIT}, {FEEDSTOCK_METHOD_MULTI}."
        )
    except Exception as exc:
        print(f"Failed to resolve feedstock method: {exc}")
        try_debug_breakpoint()
        raise


def build_input_series_map(timeseries, negative_labels, base_year, final_year):
    """Return (series_map, zero_sum_labels) for negative fuels.

    zero_sum_labels lists fuels that had no input values in the export year range.
    When INCLUDE_ALL_DETECTED_INPUT_FUELS is True those fuels are still included in
    series_map (with zero values); otherwise they are omitted.
    """
    try:
        series_map = {}
        zero_sum_labels = []
        if timeseries is None or negative_labels is None:
            return series_map, zero_sum_labels
        for label in negative_labels:
            if label not in timeseries.index:
                continue
            series = ensure_full_year_series(
                to_input_only_series(timeseries.loc[label]),
                base_year,
                final_year,
            )
            if series.sum() == 0:
                zero_sum_labels.append(label)
                if not INCLUDE_ALL_DETECTED_INPUT_FUELS:
                    continue
            series_map[label] = series
        return series_map, zero_sum_labels
    except Exception as exc:
        print(f"Failed to build input series map: {exc}")
        try_debug_breakpoint()
        raise


def log_dropped_input_fuels(economy, flow_code, zero_sum_labels, export_base_year, export_final_year):
    """Append zero-sum input fuel events to the module-level log."""
    for label in zero_sum_labels:
        included = INCLUDE_ALL_DETECTED_INPUT_FUELS
        _dropped_input_fuel_log.append(
            {
                "economy": economy,
                "flow_code": flow_code,
                "fuel_label": label,
                "export_base_year": export_base_year,
                "export_final_year": export_final_year,
                "outcome": "force_included_with_zeros" if included else "dropped",
            }
        )
        action = "force-included with zero values (INCLUDE_ALL_DETECTED_INPUT_FUELS=True)" if included else "dropped"
        print(
            f"[WARN] {flow_code} ({economy}): input fuel '{label}' has no usable values in "
            f"{export_base_year}–{export_final_year}; {action}."
        )


def get_dropped_fuel_log() -> list[dict]:
    """Return the current dropped-input-fuel log."""
    return list(_dropped_input_fuel_log)


def reset_dropped_fuel_log() -> None:
    """Clear the dropped-input-fuel log."""
    _dropped_input_fuel_log.clear()


def get_analyzed_sector_titles() -> set[str]:
    """Return the set of LEAP-mapped sector titles visited during the current run."""
    return set(_analyzed_sector_titles)


def reset_analyzed_sector_titles() -> None:
    """Clear the analyzed-sector-titles tracking set."""
    _analyzed_sector_titles.clear()


def print_dropped_fuel_summary() -> None:
    """Print a consolidated end-of-run summary of all dropped/force-included input fuels."""
    if not _dropped_input_fuel_log:
        print("[INFO] No input fuels were dropped or force-included during this run.")
        return
    dropped = [r for r in _dropped_input_fuel_log if r["outcome"] == "dropped"]
    forced = [r for r in _dropped_input_fuel_log if r["outcome"] == "force_included_with_zeros"]
    print(
        f"\n{'='*70}\n"
        f"INPUT FUEL SUMMARY ({len(_dropped_input_fuel_log)} event(s) total)\n"
        f"  Dropped (no export-range data):          {len(dropped)}\n"
        f"  Force-included with zero values:         {len(forced)}\n"
        f"{'='*70}"
    )
    if dropped:
        print("  DROPPED fuels:")
        for r in dropped:
            print(f"    [{r['economy']}] {r['flow_code']} -> {r['fuel_label']}")
    if forced:
        print("  FORCE-INCLUDED (zero values) fuels:")
        for r in forced:
            print(f"    [{r['economy']}] {r['flow_code']} -> {r['fuel_label']}")
    print(f"{'='*70}\n")


def save_dropped_fuel_report(path) -> None:
    """Save the dropped/force-included fuel log to a CSV and print summary."""
    print_dropped_fuel_summary()
    if not _dropped_input_fuel_log:
        return
    try:
        import csv as _csv
        path = str(path)
        os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
        fieldnames = ["economy", "flow_code", "fuel_label", "export_base_year", "export_final_year", "outcome"]
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = _csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(_dropped_input_fuel_log)
        print(f"[INFO] Dropped/force-included fuel report saved to: {path}")
    except Exception as exc:
        print(f"[WARN] Failed to save dropped fuel report: {exc}")


def build_total_input_series(input_series_map, years):
    """Return total input series from per-label input series."""
    try:
        total = pd.Series({year: 0.0 for year in years}, dtype=float)
        if not input_series_map:
            return total
        for series in input_series_map.values():
            if series is None:
                continue
            aligned = series.reindex(years, fill_value=0.0)
            total = total.add(aligned, fill_value=0.0)
        return total
    except Exception as exc:
        print(f"Failed to build total input series: {exc}")
        try_debug_breakpoint()
        raise


def build_input_share_series(input_series, total_input_series, fallback_to_one=False):
    """Return input share series with optional fallback when totals are zero."""
    try:
        share = safe_divide_series(input_series, total_input_series)
        if total_input_series is None or total_input_series.empty:
            return share
        zero_mask = total_input_series.eq(0.0)
        if zero_mask.any():
            if fallback_to_one:
                share = share.where(~zero_mask, 1.0)
            else:
                share = share.where(~zero_mask, 0.0)
        return share
    except Exception as exc:
        print(f"Failed to build input share series: {exc}")
        try_debug_breakpoint()
        raise


def allocate_outputs_by_share(output_series_by_label, share_series):
    """Return output series scaled by a share series."""
    try:
        if not output_series_by_label:
            return {}
        if share_series is None or share_series.empty:
            return {label: series.copy() for label, series in output_series_by_label.items()}
        allocated = {}
        for label, series in output_series_by_label.items():
            if series is None:
                continue
            allocated[label] = series.mul(share_series, fill_value=0.0)
        return allocated
    except Exception as exc:
        print(f"Failed to allocate outputs by share: {exc}")
        try_debug_breakpoint()
        raise


def allocate_loss_by_share(loss_values_by_year, share_series):
    """Return loss values scaled by a share series."""
    try:
        if not loss_values_by_year:
            return {}
        if share_series is None or share_series.empty:
            return dict(loss_values_by_year)
        share_map = {int(year): float(val) for year, val in share_series.items()}
        allocated = {}
        for label, values in loss_values_by_year.items():
            if not values:
                continue
            scaled = {}
            for year, share in share_map.items():
                value = values.get(year, values.get(str(year), 0.0))
                if value is None or pd.isna(value):
                    continue
                scaled[int(year)] = float(value) * float(share)
            if scaled:
                allocated[label] = scaled
        return allocated
    except Exception as exc:
        print(f"Failed to allocate losses by share: {exc}")
        try_debug_breakpoint()
        raise


def sum_loss_values_by_year(loss_values_by_year, years):
    """Return a year-indexed series with total losses across labels."""
    try:
        totals = {int(year): 0.0 for year in years}
        if not loss_values_by_year:
            return pd.Series(totals, dtype=float)
        for values in loss_values_by_year.values():
            if not values:
                continue
            for year in years:
                value = values.get(year, values.get(str(year), 0.0))
                if value is None or pd.isna(value):
                    continue
                totals[int(year)] = totals.get(int(year), 0.0) + abs(float(value))
        return pd.Series(totals, dtype=float)
    except Exception as exc:
        print(f"Failed to sum loss values by year: {exc}")
        try_debug_breakpoint()
        raise


def scale_year_dict_by_share(year_dict, share_series):
    """Scale a year->value dict by a share series."""
    try:
        if not year_dict:
            return {}
        if share_series is None or share_series.empty:
            return dict(year_dict)
        share_map = {int(year): float(val) for year, val in share_series.items()}
        scaled = {}
        for year, value in year_dict.items():
            if value is None or pd.isna(value):
                continue
            year_int = int(year)
            share = share_map.get(year_int, share_map.get(str(year_int), 0.0))
            scaled[year_int] = float(value) * float(share)
        return scaled
    except Exception as exc:
        print(f"Failed to scale year dict by share: {exc}")
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
            if not config_table_exists(path, sheet_name="code_to_name"):
                continue
            try:
                mapping_df = read_config_table(path, sheet_name="code_to_name")
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


def normalize_feedstock_share_to_percent(value_by_year):
    """Convert 0-1 feedstock shares to 0-100 percentages for LEAP."""
    try:
        if not value_by_year:
            return value_by_year
        values = [abs(float(v)) for v in value_by_year.values() if v is not None and not pd.isna(v)]
        if not values:
            return value_by_year
        if max(values) <= 1.000001:
            return {int(year): float(val) * 100.0 for year, val in value_by_year.items()}
        return {int(year): float(val) for year, val in value_by_year.items()}
    except Exception as exc:
        print(f"Failed to normalize feedstock shares to percent: {exc}")
        try_debug_breakpoint()
        raise


def resolve_scenario_year_range(base_year, final_year, scenario_config=None):
    """Return the year window to export for the active scenario."""
    try:
        scenario_base_year = int(base_year)
        scenario_final_year = int(final_year)
        if scenario_config:
            scenario_base_year = int(
                scenario_config.get("export_base_year", scenario_base_year)
            )
            scenario_final_year = int(
                scenario_config.get("export_final_year", scenario_final_year)
            )
        if scenario_final_year < scenario_base_year:
            scenario_final_year = scenario_base_year
        return scenario_base_year, scenario_final_year
    except Exception as exc:
        print(f"Failed to resolve scenario year range: {exc}")
        try_debug_breakpoint()
        raise


def clip_value_by_year_range(value_by_year, start_year, end_year):
    """Keep only year entries within [start_year, end_year]."""
    try:
        if not value_by_year:
            return {}
        start = int(start_year)
        end = int(end_year)
        return {
            int(year): float(value)
            for year, value in value_by_year.items()
            if start <= int(year) <= end and value is not None and not pd.isna(value)
        }
    except Exception as exc:
        print(f"Failed to clip year map to range {start_year}-{end_year}: {exc}")
        try_debug_breakpoint()
        raise


def normalize_feedstock_shares_for_export(feedstock_shares, base_year, final_year):
    """Normalize per-fuel shares so each process-year sums to exactly 100%."""
    try:
        if not feedstock_shares:
            return {}
        labels = list(feedstock_shares.keys())
        if not labels:
            return {}

        normalized = {}
        for label, raw_share in feedstock_shares.items():
            value_by_year = coerce_value_by_year(raw_share, base_year, final_year)
            value_by_year = clip_value_by_year_range(
                value_by_year,
                base_year,
                final_year,
            )
            value_by_year = normalize_feedstock_share_to_percent(value_by_year)
            normalized[label] = {
                int(year): max(float(value), 0.0)
                for year, value in value_by_year.items()
                if value is not None and not pd.isna(value)
            }

        # Build a stable fallback distribution for years with zero/undefined totals.
        fallback_weights = {
            label: sum(max(float(value), 0.0) for value in normalized[label].values())
            for label in labels
        }
        fallback_total = sum(fallback_weights.values())
        if fallback_total > 0.0:
            fallback_distribution = {
                label: fallback_weights[label] * 100.0 / fallback_total
                for label in labels
            }
            anchor_label = max(fallback_distribution, key=fallback_distribution.get)
            fallback_distribution[anchor_label] += 100.0 - sum(
                fallback_distribution.values()
            )
        else:
            anchor_label = labels[0]
            fallback_distribution = {label: 0.0 for label in labels}
            fallback_distribution[anchor_label] = 100.0

        # Build explicit normalized profiles for years that have a valid total.
        positive_profiles = {}
        zero_years = []
        tolerance = 1e-12
        for year in range(int(base_year), int(final_year) + 1):
            year_total = sum(normalized[label].get(year, 0.0) for label in labels)
            if year_total <= tolerance:
                zero_years.append(year)
                continue

            scaled_by_label = {
                label: normalized[label].get(year, 0.0) * 100.0 / year_total
                for label in labels
            }
            anchor_label = max(scaled_by_label, key=scaled_by_label.get)
            residual = 100.0 - sum(scaled_by_label.values())
            scaled_by_label[anchor_label] += residual
            positive_profiles[year] = scaled_by_label

        # For years with zero/undefined totals, copy nearest nonzero profile.
        # Preference: nearest future year, then nearest past year.
        valid_years = sorted(positive_profiles.keys())
        for year in zero_years:
            chosen_profile = None
            if valid_years:
                chosen_year = min(
                    valid_years,
                    key=lambda candidate: (
                        abs(candidate - year),
                        0 if candidate >= year else 1,
                    ),
                )
                chosen_profile = positive_profiles.get(chosen_year)
            if chosen_profile is None:
                chosen_profile = fallback_distribution
            positive_profiles[year] = dict(chosen_profile)

        for year in range(int(base_year), int(final_year) + 1):
            scaled_by_label = positive_profiles.get(year, fallback_distribution)
            for label, value in scaled_by_label.items():
                normalized[label][year] = value
        return normalized
    except Exception as exc:
        print(f"Failed to normalize feedstock shares for export: {exc}")
        try_debug_breakpoint()
        raise


def _build_feedstock_label_universe(process_records):
    """Return per-process feedstock label sets discovered in records."""
    try:
        label_lookup = {}
        for record in process_records or []:
            key = (
                str(record.get("sector_title") or "").strip(),
                str(record.get("process_name") or "").strip(),
            )
            if not key[0] or not key[1]:
                continue
            labels = []
            for source in (record.get("feedstock_values"), record.get("feedstock_shares")):
                if isinstance(source, dict):
                    labels.extend(str(label).strip() for label in source.keys() if str(label).strip())
            if not labels:
                continue
            bucket = label_lookup.setdefault(key, [])
            for label in labels:
                if label not in bucket:
                    bucket.append(label)
        return label_lookup
    except Exception as exc:
        print(f"Failed to build feedstock label universe: {exc}")
        try_debug_breakpoint()
        raise


def prepare_feedstock_shares_for_export(
    feedstock_shares,
    feedstock_values,
    process_feedstock_labels,
    base_year,
    final_year,
):
    """Return normalized feedstock shares with zero-use safeguards for LEAP export."""
    try:
        labels = []
        for source in (process_feedstock_labels, feedstock_values, feedstock_shares):
            if isinstance(source, dict):
                values = source.keys()
            elif isinstance(source, list):
                values = source
            else:
                values = []
            for raw in values:
                label = str(raw).strip()
                if label and label not in labels:
                    labels.append(label)
        if not labels:
            return {}

        start = int(base_year)
        end = int(final_year)
        years = list(range(start, end + 1))
        prepared_shares = {}
        usage_totals = {}

        for label in labels:
            share_map = {}
            if isinstance(feedstock_shares, dict) and label in feedstock_shares:
                share_map = coerce_value_by_year(feedstock_shares.get(label), start, end)
                share_map = clip_value_by_year_range(share_map, start, end)
                share_map = normalize_feedstock_share_to_percent(share_map)
            prepared_shares[label] = {
                int(year): float(share_map.get(year, 0.0))
                for year in years
            }

            usage_map = {}
            if isinstance(feedstock_values, dict) and label in feedstock_values:
                usage_map = coerce_value_by_year(feedstock_values.get(label), start, end)
                usage_map = clip_value_by_year_range(usage_map, start, end)
            usage_totals[label] = sum(abs(float(usage_map.get(year, 0.0))) for year in years)

        # Fuels not used in the current economy/process stay at zero share in all years.
        usage_tolerance = 1e-12
        for label, usage_total in usage_totals.items():
            if usage_total <= usage_tolerance:
                prepared_shares[label] = {year: 0.0 for year in years}

        process_used = any(usage_total > usage_tolerance for usage_total in usage_totals.values())
        if not process_used:
            # Process branch exists but is unused: spread 100% equally to avoid LEAP share errors.
            equal_share = 100.0 / float(len(labels))
            for label in labels:
                prepared_shares[label] = {year: equal_share for year in years}
            anchor_label = labels[0]
            for year in years:
                residual = 100.0 - sum(prepared_shares[label][year] for label in labels)
                prepared_shares[anchor_label][year] += residual

        return normalize_feedstock_shares_for_export(prepared_shares, start, end)
    except Exception as exc:
        print(f"Failed to prepare feedstock shares for export: {exc}")
        try_debug_breakpoint()
        raise


def normalize_process_efficiency_to_percent(value_by_year, input_scale="ratio"):
    """Normalize process efficiency to percent using explicit input-scale rules."""
    try:
        if not value_by_year:
            return value_by_year
        cleaned = {
            int(year): float(val)
            for year, val in value_by_year.items()
            if val is not None and not pd.isna(val)
        }
        if not cleaned:
            return value_by_year
        scale_key = str(input_scale or "ratio").strip().lower()
        if scale_key == "ratio":
            cleaned = {year: value * 100.0 for year, value in cleaned.items()}
        elif scale_key == "percent":
            pass
        else:
            raise ValueError(
                f"Unknown process efficiency input_scale={input_scale!r}. "
                "Use 'ratio' or 'percent'."
            )
        # LEAP rejects non-positive process efficiencies. Prefer carrying forward
        # the previous valid year to avoid artificial tiny values in exports.
        years_sorted = sorted(cleaned.keys())
        first_positive = next((cleaned[year] for year in years_sorted if cleaned[year] > 0.0), None)
        fallback = float(first_positive) if first_positive is not None else 1.0
        last_valid = None
        normalized: dict[int, float] = {}
        for year in years_sorted:
            value = cleaned[year]
            if value > 0.0:
                normalized[year] = value
                last_valid = value
            elif last_valid is not None:
                normalized[year] = last_valid
            else:
                normalized[year] = fallback
        cleaned = normalized
        return cleaned
    except Exception as exc:
        print(f"Failed to normalize process efficiency to percent: {exc}")
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

def compute_own_use_ratios_for_record(loss_values_by_year, input_series_map, years):
    """Return {fuel_label: scalar_ratio} for fuels that appear in BOTH losses and feedstock inputs.

    ratio = mean_own_use / (mean_own_use + mean_feedstock_input) over years.

    Only fuels where 0 < ratio < 1 are returned — i.e. fuels that are ambiguously used as
    both feedstock and own-use.  Pure-feedstock fuels (ratio=0) and pure-own-use fuels
    (ratio=1, not in input_series_map) are omitted.

    These ratios are stored on the process record so that on subsequent LEAP results
    readbacks, the reported feedstock 'Inputs' can be split into true feedstock vs
    estimated own-use, allowing auxiliary_ratios to be recalibrated.
    """
    try:
        ratios = {}
        year_index = pd.Index(years)
        for fuel, loss_by_year in (loss_values_by_year or {}).items():
            input_series = (input_series_map or {}).get(fuel)
            if input_series is None:
                continue
            loss_mean = float(
                pd.Series(loss_by_year, dtype=float).reindex(year_index, fill_value=0.0).mean()
            )
            input_mean = float(input_series.reindex(year_index, fill_value=0.0).mean())
            total_mean = loss_mean + input_mean
            if total_mean > 0 and loss_mean > 0:
                ratios[fuel] = loss_mean / total_mean
        return ratios
    except Exception as exc:
        print(f"Failed to compute own-use ratios: {exc}")
        try_debug_breakpoint()
        raise


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
    loss_values_for_efficiency=None,
    feedstock_shares=None,
    input_total=None,
    output_import_targets=None,
    output_export_targets=None,
    efficiency_scale="ratio",
    own_use_ratios=None,
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
            "efficiency_scale": str(efficiency_scale or "ratio"),
            "auxiliary_ratios": dict(auxiliary_ratios or {}),
            "loss_values": dict(loss_values or {}),
            "loss_total": loss_total,
            "loss_values_for_efficiency": dict(loss_values_for_efficiency or {}),
            "input_total": input_total,
            "output_import_targets": dict(output_import_targets or {}),
            "output_export_targets": dict(output_export_targets or {}),
            "own_use_ratios": dict(own_use_ratios or {}),
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
                efficiency_data = record.get("efficiency")
                if isinstance(efficiency_data, dict):
                    efficiency_data = normalize_process_efficiency_to_percent(
                        efficiency_data,
                        input_scale=record.get("efficiency_scale", "ratio"),
                    )
                efficiency_value = summarize_numeric_value(
                    efficiency_data, summary="mean"
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


def _sum_series_map_by_year(value_map, years):
    """Return year totals from a fuel->(year->value) mapping."""
    try:
        totals = {int(year): 0.0 for year in years}
        if not isinstance(value_map, dict):
            return totals
        for raw_series in value_map.values():
            if isinstance(raw_series, dict):
                for year in years:
                    raw_value = raw_series.get(year, raw_series.get(str(year), 0.0))
                    if raw_value is None or pd.isna(raw_value):
                        continue
                    totals[int(year)] += abs(float(raw_value))
                continue
            if raw_series is None or pd.isna(raw_series):
                continue
            scalar_value = abs(float(raw_series))
            for year in years:
                totals[int(year)] += scalar_value
        return totals
    except Exception as exc:
        print(f"Failed to sum series map by year: {exc}")
        try_debug_breakpoint()
        raise


def _build_process_share_lookup(
    process_records,
    code_to_name_mapping,
    base_year,
    final_year,
):
    """Return per-process share percentages by economy/sector/year."""
    try:
        years = list(range(int(base_year), int(final_year) + 1))
        process_sets = {}
        process_activity = {}
        for record in process_records or []:
            economy = str(record.get("economy") or "").strip()
            sector_title = map_code_label(record.get("sector_title"), code_to_name_mapping)
            process_name = map_code_label(record.get("process_name"), code_to_name_mapping)
            if not sector_title or not process_name:
                continue
            key = (economy, str(sector_title))
            process_sets.setdefault(key, set()).add(str(process_name))
            process_key = (economy, str(sector_title), str(process_name))
            output_totals = _sum_series_map_by_year(record.get("output_values"), years)
            feedstock_totals = _sum_series_map_by_year(record.get("feedstock_values"), years)
            loss_totals = _sum_series_map_by_year(record.get("loss_values"), years)
            totals = {}
            for year in years:
                # Use output-first activity as the main basis for Process Share.
                output_value = float(output_totals.get(year, 0.0))
                # Fallback to wider activity if output is unavailable in that year.
                total_activity = (
                    output_value
                    + float(feedstock_totals.get(year, 0.0))
                    + float(loss_totals.get(year, 0.0))
                )
                totals[int(year)] = (
                    output_value
                    if output_value > 0.0
                    else total_activity
                )
            process_activity[process_key] = totals

        share_lookup = {}
        usage_tolerance = 1e-12
        for (economy, sector_title), processes in process_sets.items():
            process_list = sorted(processes)
            if not process_list:
                continue
            for process_name in process_list:
                share_lookup[(economy, sector_title, process_name)] = {
                    int(year): 0.0 for year in years
                }
            for year in years:
                active_processes = []
                denominator = 0.0
                for process_name in process_list:
                    activity = process_activity.get(
                        (economy, sector_title, process_name),
                        {},
                    ).get(year, 0.0)
                    if float(activity) > usage_tolerance:
                        activity_value = float(activity)
                        active_processes.append((process_name, activity_value))
                        denominator += activity_value
                if not active_processes:
                    continue
                if denominator <= usage_tolerance:
                    continue
                for process_name, activity_value in active_processes:
                    share_value = (activity_value / denominator) * 100.0
                    share_lookup[(economy, sector_title, process_name)][int(year)] = share_value
        return share_lookup
    except Exception as exc:
        print(f"Failed to build process share lookup: {exc}")
        try_debug_breakpoint()
        raise


def _normalize_share_percentages(raw_share_by_label, rounding_decimals=6):
    """Return rounded shares that sum to exactly 100.0."""
    try:
        if not raw_share_by_label:
            return {}
        ordered = sorted(raw_share_by_label.items(), key=lambda item: str(item[0]))
        rounded = {
            label: round(max(float(value), 0.0), int(rounding_decimals))
            for label, value in ordered
        }
        rounded_sum = round(sum(rounded.values()), int(rounding_decimals))
        residual = round(100.0 - rounded_sum, int(rounding_decimals))
        if abs(residual) > 0.0:
            target_label = sorted(
                raw_share_by_label.items(),
                key=lambda item: (-float(item[1]), str(item[0])),
            )[0][0]
            rounded[target_label] = round(
                float(rounded.get(target_label, 0.0)) + residual,
                int(rounding_decimals),
            )
        return rounded
    except Exception as exc:
        print(f"Failed to normalize share percentages: {exc}")
        try_debug_breakpoint()
        raise


def _build_output_share_lookup(
    process_records,
    code_to_name_mapping,
    base_year,
    final_year,
):
    """Return output fuel share percentages by (economy, sector, fuel, year)."""
    try:
        years = list(range(int(base_year), int(final_year) + 1))
        sector_fuels = {}
        output_totals = {}
        value_tolerance = 1e-12
        for record in process_records or []:
            economy = str(record.get("economy") or "").strip()
            sector_title = map_code_label(record.get("sector_title"), code_to_name_mapping)
            if not economy or not sector_title:
                continue
            sector_key = (economy, str(sector_title))
            output_values = record.get("output_values") or {}
            for label, value in output_values.items():
                fuel_label = map_code_label(label, code_to_name_mapping)
                if not fuel_label:
                    continue
                sector_fuels.setdefault(sector_key, set()).add(str(fuel_label))
                value_by_year = coerce_value_by_year(value, base_year, final_year)
                for year in years:
                    year_value = float(value_by_year.get(year, value_by_year.get(str(year), 0.0)) or 0.0)
                    year_value = max(year_value, 0.0)
                    key = (economy, str(sector_title), str(fuel_label), int(year))
                    output_totals[key] = output_totals.get(key, 0.0) + year_value

        lookup = {}
        for sector_key, fuel_set in sector_fuels.items():
            fuel_labels = sorted(fuel_set)
            if not fuel_labels:
                continue
            economy, sector_title = sector_key
            sector_bucket = lookup.setdefault((economy, sector_title), {})
            last_nonzero_share = None
            for year in years:
                fuel_totals = {
                    fuel_label: float(
                        output_totals.get((economy, sector_title, fuel_label, int(year)), 0.0)
                    )
                    for fuel_label in fuel_labels
                }
                sector_total = float(sum(fuel_totals.values()))
                if sector_total > value_tolerance:
                    raw_shares = {
                        fuel_label: (fuel_totals[fuel_label] / sector_total) * 100.0
                        for fuel_label in fuel_labels
                    }
                    share_by_label = _normalize_share_percentages(raw_shares)
                    last_nonzero_share = dict(share_by_label)
                elif last_nonzero_share is not None:
                    share_by_label = dict(last_nonzero_share)
                else:
                    equal_share = 100.0 / float(len(fuel_labels))
                    share_by_label = _normalize_share_percentages(
                        {fuel_label: equal_share for fuel_label in fuel_labels}
                    )
                for fuel_label in fuel_labels:
                    sector_bucket.setdefault(fuel_label, {})[int(year)] = float(
                        share_by_label.get(fuel_label, 0.0)
                    )
        return lookup
    except Exception as exc:
        print(f"Failed to build output share lookup: {exc}")
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
        process_feedstock_labels = _build_feedstock_label_universe(process_records)
        process_share_lookup = _build_process_share_lookup(
            process_records,
            code_to_name_mapping,
            base_year,
            final_year,
        )
        output_share_lookup = _build_output_share_lookup(
            process_records,
            code_to_name_mapping,
            base_year,
            final_year,
        )
        emitted_output_share_sectors = set()
        scenario_base_year, scenario_final_year = resolve_scenario_year_range(
            base_year,
            final_year,
            scenario_config,
        )
        for record in process_records:
            economy = str(record.get("economy") or "").strip()
            sector_title = map_code_label(record.get("sector_title"), code_to_name_mapping)
            process_name = map_code_label(record.get("process_name"), code_to_name_mapping)
            output_values = record.get("output_values", {})
            feedstock_shares = record.get("feedstock_shares", {})
            auxiliary_ratios = record.get("auxiliary_ratios", {})
            efficiency = record.get("efficiency")
            sector_key = (economy, str(sector_title))
            if sector_key not in emitted_output_share_sectors:
                sector_output_shares = output_share_lookup.get(sector_key, {})
                for fuel_label, value in sorted(sector_output_shares.items(), key=lambda item: str(item[0])):
                    value_by_year = clip_value_by_year_range(
                        value,
                        scenario_base_year,
                        scenario_final_year,
                    )
                    branch_path = build_branch_path(
                        [
                            "Transformation",
                            sector_title,
                            "Output Fuels",
                            str(fuel_label),
                        ]
                    )
                    rows.extend(
                        build_year_rows(
                            branch_path,
                            "Output Share",
                            scenario,
                            value_by_year,
                            "Share",
                            "%",
                            "",
                        )
                    )
                emitted_output_share_sectors.add(sector_key)

            scenario_process_share = record.get("process_share_by_year")
            if isinstance(scenario_process_share, dict) and scenario_process_share:
                process_share = scenario_process_share
            else:
                process_share = process_share_lookup.get(
                    (economy, str(sector_title), str(process_name)),
                    {},
                )
            if isinstance(process_share, dict):
                process_share_by_year = clip_value_by_year_range(
                    process_share,
                    scenario_base_year,
                    scenario_final_year,
                )
            else:
                process_share_by_year = build_value_by_year(
                    process_share,
                    scenario_base_year,
                    scenario_final_year,
                )
            process_branch_path = build_branch_path(
                ["Transformation", sector_title, "Processes", str(process_name)]
            )
            rows.extend(
                build_year_rows(
                    process_branch_path,
                    "Process Share",
                    scenario,
                    process_share_by_year,
                    "Share",
                    "%",
                    "",
                )
            )

            capacity_units = str(record.get("capacity_units") or "Gigajoules/Year")
            capacity_scale = str(record.get("capacity_scale") or "")
            historical_production_by_year = record.get("historical_production_by_year")
            if isinstance(historical_production_by_year, dict) and historical_production_by_year:
                historical_values = clip_value_by_year_range(
                    historical_production_by_year,
                    scenario_base_year,
                    scenario_final_year,
                )
                rows.extend(
                    build_year_rows(
                        process_branch_path,
                        "Historical Production",
                        scenario,
                        historical_values,
                        "Petajoule",
                        "",
                        "",
                    )
                )

            exogenous_capacity_by_year = record.get("exogenous_capacity_by_year")
            if isinstance(exogenous_capacity_by_year, dict) and exogenous_capacity_by_year:
                exogenous_values = clip_value_by_year_range(
                    exogenous_capacity_by_year,
                    scenario_base_year,
                    scenario_final_year,
                )
                rows.extend(
                    build_year_rows(
                        process_branch_path,
                        "Exogenous Capacity",
                        scenario,
                        exogenous_values,
                        capacity_units,
                        capacity_scale,
                        "",
                    )
                )

            endogenous_capacity_by_year = record.get("endogenous_capacity_by_year")
            if isinstance(endogenous_capacity_by_year, dict) and endogenous_capacity_by_year:
                endogenous_values = clip_value_by_year_range(
                    endogenous_capacity_by_year,
                    scenario_base_year,
                    scenario_final_year,
                )
                rows.extend(
                    build_year_rows(
                        process_branch_path,
                        "Endogenous Capacity",
                        scenario,
                        endogenous_values,
                        capacity_units,
                        capacity_scale,
                        "",
                    )
                )

            max_availability_by_year = record.get("maximum_availability_by_year")
            if isinstance(max_availability_by_year, dict) and max_availability_by_year:
                availability_values = clip_value_by_year_range(
                    max_availability_by_year,
                    scenario_base_year,
                    scenario_final_year,
                )
                rows.extend(
                    build_year_rows(
                        process_branch_path,
                        "Maximum Availability",
                        scenario,
                        availability_values,
                        "Percent",
                        "",
                        "",
                    )
                )

            capacity_credit_by_year = record.get("capacity_credit_by_year")
            if isinstance(capacity_credit_by_year, dict) and capacity_credit_by_year:
                credit_values = clip_value_by_year_range(
                    capacity_credit_by_year,
                    scenario_base_year,
                    scenario_final_year,
                )
                rows.extend(
                    build_year_rows(
                        process_branch_path,
                        "Capacity Credit",
                        scenario,
                        credit_values,
                        "Percent",
                        "",
                        "",
                    )
                )

            if (
                INCLUDE_OUTPUT_SERIES_IN_LEAP_EXPORT
                and TRANSFORMATION_OUTPUT_VARIABLES.get("output")
            ):
                for label, value in output_values.items():
                    value_by_year = coerce_value_by_year(value, base_year, final_year)
                    value_by_year = clip_value_by_year_range(
                        value_by_year,
                        scenario_base_year,
                        scenario_final_year,
                    )
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
                    value_by_year = clip_value_by_year_range(
                        value_by_year,
                        scenario_base_year,
                        scenario_final_year,
                    )
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
                    value_by_year = clip_value_by_year_range(
                        value_by_year,
                        scenario_base_year,
                        scenario_final_year,
                    )
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
                efficiency_by_year = clip_value_by_year_range(
                    efficiency_by_year,
                    scenario_base_year,
                    scenario_final_year,
                )
                efficiency_by_year = normalize_process_efficiency_to_percent(
                    efficiency_by_year,
                    input_scale=record.get("efficiency_scale", "ratio"),
                )
                rows.extend(
                    build_year_rows(
                        process_branch_path,
                        "Process Efficiency",
                        scenario,
                        efficiency_by_year,
                        DEFAULT_EFFICIENCY_UNITS,
                        "",
                        "",
                    )
                )

            if TRANSFORMATION_OUTPUT_VARIABLES.get("feedstock_share"):
                process_key = (
                    str(record.get("sector_title") or "").strip(),
                    str(record.get("process_name") or "").strip(),
                )
                normalized_feedstock_shares = prepare_feedstock_shares_for_export(
                    feedstock_shares,
                    record.get("feedstock_values", {}),
                    process_feedstock_labels.get(process_key, []),
                    scenario_base_year,
                    scenario_final_year,
                )
                for label, value_by_year in normalized_feedstock_shares.items():
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
                    auxiliary_leaf = map_code_label(label, code_to_name_mapping)
                    value_by_year = coerce_value_by_year(value, base_year, final_year)
                    value_by_year = clip_value_by_year_range(
                        value_by_year,
                        scenario_base_year,
                        scenario_final_year,
                    )
                    branch_path = build_branch_path(
                        [
                            "Transformation",
                            sector_title,
                            "Processes",
                            str(process_name),
                            "Auxiliary Fuels",
                            auxiliary_leaf,
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
                scenario_base_year,
                scenario_final_year,
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
        scenario_name = str(row.get("Scenario", "")).strip().lower()
        is_current_accounts = scenario_name in {"current accounts", "current account"}
        parts = []
        for year in year_cols:
            value = row.get(year)
            if value is None or pd.isna(value):
                continue
            parts.append(f"{int(year)},{float(value)}")
        if is_current_accounts and not parts:
            fallback_year = int(year_cols[0]) if year_cols else int(EXPORT_BASE_YEAR)
            parts.append(f"{fallback_year},0.0")
        if not is_current_accounts and not parts:
            fallback_year = int(year_cols[0]) if year_cols else int(EXPORT_BASE_YEAR)
            parts.append(f"{fallback_year},0.0")
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


def _sum_year_dicts(series_list):
    """Sum year->value dicts, aligning years."""
    totals = {}
    for series in series_list:
        if not series:
            continue
        for year, value in series.items():
            if value is None or pd.isna(value):
                continue
            year_int = int(year)
            totals[year_int] = totals.get(year_int, 0.0) + float(value)
    return totals


def consolidate_transformation_output_rows(
    process_records,
    include_output_series=False,
    use_output_targets=False,
):
    """Aggregate output values/targets to avoid duplicate LEAP rows."""
    if not process_records or not (include_output_series or use_output_targets):
        return
    grouped = {}
    for record in process_records:
        key = (record.get("economy"), record.get("sector_title"))
        grouped.setdefault(key, []).append(record)
    for records in grouped.values():
        if len(records) < 2:
            continue
        output_values_by_label = {}
        import_targets_by_label = {}
        export_targets_by_label = {}
        for record in records:
            for label, values in (record.get("output_values") or {}).items():
                output_values_by_label.setdefault(label, []).append(values)
            for label, values in (record.get("output_import_targets") or {}).items():
                import_targets_by_label.setdefault(label, []).append(values)
            for label, values in (record.get("output_export_targets") or {}).items():
                export_targets_by_label.setdefault(label, []).append(values)
        aggregated_outputs = {
            label: _sum_year_dicts(values)
            for label, values in output_values_by_label.items()
            if values
        }
        aggregated_imports = {
            label: _sum_year_dicts(values)
            for label, values in import_targets_by_label.items()
            if values
        }
        aggregated_exports = {
            label: _sum_year_dicts(values)
            for label, values in export_targets_by_label.items()
            if values
        }
        carrier = records[0]
        if include_output_series:
            carrier["output_values"] = aggregated_outputs
        if use_output_targets:
            carrier["output_import_targets"] = aggregated_imports
            carrier["output_export_targets"] = aggregated_exports
        for record in records[1:]:
            if include_output_series:
                record["output_values"] = {}
            if use_output_targets:
                record["output_import_targets"] = {}
                record["output_export_targets"] = {}


def build_aux_fuel_zero_rows(
    existing_rows,
    full_branch_catalog_df,
    scenarios,
    base_year,
    final_year,
    in_scope_sector_titles: set[str] | None = None,
):
    """
    Return zero-value rows for any catalog fuel branches not already set.

    Covers Auxiliary Fuels (Auxiliary Fuel Use) and Feedstock Fuels (Feedstock Fuel Share).

    Two tiers of zero-fill:

    1. Measure-specific process prefixes (derived from existing_rows): only branches under
       processes where we actually wrote that specific measure this run.  Feedstock Fuel Share
       zero-fill also applies the 100.0 fallback so LEAP doesn't reject an all-zero process.
       Data written by other researchers is never touched here.

    2. In-scope sector titles (from in_scope_sector_titles): sectors this workflow is
       responsible for but where some or all economies had no ESTO data (so no process records
       were produced).  All catalog branches under these sectors that weren't written in tier 1
       are cleared to 0.  No 100.0 fallback here — we have no data to anchor a share.

    NOTE (part-2 limitation): On re-runs after a LEAP model solve, auxiliary fuel use
    values are reported by LEAP as transformation inputs in the energy balance, not as
    separate own-use rows.  There is currently no reliable way to extract them back out
    and feed them into the ESTO-based auxiliary ratio calculation for the next iteration.
    This is a known gap — the auxiliary ratios used here always derive from ESTO loss data
    regardless of what LEAP reports.
    """
    if full_branch_catalog_df is None or (
        hasattr(full_branch_catalog_df, "empty") and full_branch_catalog_df.empty
    ):
        return []

    # Derive allowed process branch prefixes per measure from rows already written this run.
    # A branch path containing "\Processes\" gives us the process prefix up to and
    # including the process name, e.g. "Transformation\LNG\Processes\Liquefaction".
    # We use measure-specific sets so that, e.g., writing Process Share for a sector
    # does NOT entitle us to zero-fill that sector's Feedstock Fuel Share branches —
    # only processes for which we actually wrote Feedstock Fuel Share rows are touched.
    _processes_marker = "\\Processes\\"

    def _extract_process_prefix(bp: str) -> str | None:
        idx = bp.find(_processes_marker)
        if idx == -1:
            return None
        after_marker = idx + len(_processes_marker)
        next_sep = bp.find("\\", after_marker)
        return bp[:next_sep] if next_sep != -1 else bp

    # measure_name → set of process prefixes we wrote rows for
    allowed_prefixes_by_measure: dict[str, set[str]] = {}
    for row in existing_rows:
        measure = str(row.get("Measure", "")).strip()
        bp = str(row.get("Branch_Path", "")).strip()
        prefix = _extract_process_prefix(bp)
        if prefix:
            allowed_prefixes_by_measure.setdefault(measure, set()).add(prefix)

    # Map fuel_group label → (LEAP variable name, units, scale, per)
    fuel_group_spec = {
        "Auxiliary Fuels": ("Auxiliary Fuel Use", DEFAULT_AUXILIARY_UNITS, "", DEFAULT_AUXILIARY_PER),
        "Feedstock Fuels": ("Feedstock Fuel Share", DEFAULT_FEEDSTOCK_UNITS, "", ""),
    }

    # Collect (measure, scenario, branch_path) triples already written.
    existing: set[tuple[str, str, str]] = set()
    for row in existing_rows:
        existing.add((
            str(row.get("Measure", "")).strip(),
            str(row.get("Scenario", "")),
            str(row.get("Branch_Path", "")),
        ))

    zero_rows = []
    years = list(range(int(base_year), int(final_year) + 1))
    for group, (measure, units, scale, per) in fuel_group_spec.items():
        # Only operate on processes where we actually wrote this specific measure.
        # This prevents zeroing feedstock branches for sectors where we only wrote
        # Process Share / Efficiency, and avoids touching any data set by others.
        allowed_prefixes = allowed_prefixes_by_measure.get(measure, set())
        if not allowed_prefixes:
            continue

        group_catalog = full_branch_catalog_df[
            full_branch_catalog_df["fuel_group"].astype(str).str.strip() == group
        ]
        if group_catalog.empty:
            continue

        is_feedstock = measure == "Feedstock Fuel Share"

        # Pre-compute which process prefixes have at least one written row for this measure,
        # keyed by (scenario). This catches branches written but absent from the catalog
        # (e.g. Lignite for BKB when only Other bituminous coal is in the catalog), so
        # those processes are not incorrectly treated as having no written rows.
        written_prefixes_by_scenario: dict[str, set[str]] = {}
        for ex_m, ex_sc, ex_bp in existing:
            if ex_m != measure:
                continue
            ex_prefix = _extract_process_prefix(ex_bp)
            if ex_prefix:
                written_prefixes_by_scenario.setdefault(ex_sc, set()).add(ex_prefix)

        for scenario in scenarios:
            # process_prefix → list of (branch_path, is_already_set) — for feedstock grouping
            process_branch_map: dict[str, list[tuple[str, bool]]] = {}

            for _, catalog_row in group_catalog.iterrows():
                branch_path = str(catalog_row.get("branch_path", "")).strip()
                if not branch_path:
                    continue
                matched_prefix = next(
                    (p for p in allowed_prefixes if branch_path.startswith(p + "\\")),
                    None,
                )
                if matched_prefix is None:
                    continue
                already_set = (measure, scenario, branch_path) in existing
                if is_feedstock:
                    process_branch_map.setdefault(matched_prefix, []).append(
                        (branch_path, already_set)
                    )
                elif not already_set:
                    for year in years:
                        zero_rows.append(
                            {
                                "Branch_Path": branch_path,
                                "Scenario": scenario,
                                "Measure": measure,
                                "Units": units,
                                "Scale": scale,
                                "Per...": per,
                                "Date": year,
                                "Value": 0.0,
                            }
                        )

            if is_feedstock:
                for prefix, branches in process_branch_map.items():
                    # any_already_set is True when at least one catalog branch was written
                    # OR when any row for this process prefix was written (even branches
                    # absent from the catalog — e.g. Lignite for BKB when only Other
                    # bituminous coal is catalogued). Without the second check, the first
                    # catalog branch incorrectly receives the 100.0 anchor.
                    any_already_set = (
                        any(already_set for _, already_set in branches)
                        or prefix in written_prefixes_by_scenario.get(scenario, set())
                    )
                    first_unset = True
                    for branch_path, already_set in branches:
                        if already_set:
                            continue
                        if not any_already_set and first_unset:
                            value = 100.0
                            first_unset = False
                        else:
                            value = 0.0
                        for year in years:
                            zero_rows.append(
                                {
                                    "Branch_Path": branch_path,
                                    "Scenario": scenario,
                                    "Measure": measure,
                                    "Units": units,
                                    "Scale": scale,
                                    "Per...": per,
                                    "Date": year,
                                    "Value": value,
                                }
                            )

    # Tier-2: clear branches for in-scope sectors where we had no data this run.
    # These are sectors this workflow owns but produced no process records for.
    # All catalog branches under them that weren't already handled in tier 1 are zeroed,
    # except Feedstock Fuel Share which gets the same 100.0 anchor as tier 1 so LEAP
    # doesn't reject an all-zero process.
    if in_scope_sector_titles:
        for group, (measure, units, scale, per) in fuel_group_spec.items():
            group_catalog = full_branch_catalog_df[
                full_branch_catalog_df["fuel_group"].astype(str).str.strip() == group
            ]
            if group_catalog.empty:
                continue
            tier1_prefixes = allowed_prefixes_by_measure.get(measure, set())
            is_feedstock = measure == "Feedstock Fuel Share"
            for scenario in scenarios:
                # For feedstock, group by process prefix so we can anchor the first branch
                # at 100.0.  For non-feedstock, collect directly for zeroing.
                tier2_process_map: dict[str, list[str]] = {}
                for _, catalog_row in group_catalog.iterrows():
                    branch_path = str(catalog_row.get("branch_path", "")).strip()
                    if not branch_path:
                        continue
                    # Skip if already handled by tier 1 (under a process we wrote).
                    if any(branch_path.startswith(p + "\\") for p in tier1_prefixes):
                        continue
                    # Skip if already explicitly set this run.
                    if (measure, scenario, branch_path) in existing:
                        continue
                    # Check if this branch is under an in-scope sector.
                    under_in_scope = any(
                        branch_path.startswith("Transformation\\" + title + "\\")
                        for title in in_scope_sector_titles
                    )
                    if not under_in_scope:
                        continue
                    if is_feedstock:
                        prefix = _extract_process_prefix(branch_path)
                        if prefix:
                            tier2_process_map.setdefault(prefix, []).append(branch_path)
                            continue
                    for year in years:
                        zero_rows.append(
                            {
                                "Branch_Path": branch_path,
                                "Scenario": scenario,
                                "Measure": measure,
                                "Units": units,
                                "Scale": scale,
                                "Per...": per,
                                "Date": year,
                                "Value": 0.0,
                            }
                        )

                if is_feedstock:
                    for prefix, branches in tier2_process_map.items():
                        for i, branch_path in enumerate(branches):
                            value = 100.0 if i == 0 else 0.0
                            for year in years:
                                zero_rows.append(
                                    {
                                        "Branch_Path": branch_path,
                                        "Scenario": scenario,
                                        "Measure": measure,
                                        "Units": units,
                                        "Scale": scale,
                                        "Per...": per,
                                        "Date": year,
                                        "Value": value,
                                    }
                                )

    return zero_rows


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
    full_branch_catalog_df=None,
    in_scope_sector_titles: set[str] | None = None,
):
    """Save a LEAP import file built from process records across scenarios."""
    try:
        if not process_records:
            print("No process records available for LEAP export.")
            return None
        scenario_configs = {
            scenario: get_scenario_export_config(
                scenario,
                default_base_year=base_year,
                default_final_year=final_year,
            )
            for scenario in scenarios
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
        if full_branch_catalog_df is not None:
            zero_rows = build_aux_fuel_zero_rows(
                combined_rows,
                full_branch_catalog_df,
                scenarios,
                combined_base_year,
                combined_final_year,
                in_scope_sector_titles=in_scope_sector_titles,
            )
            if zero_rows:
                combined_rows = zero_rows + combined_rows
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
        export_base_year = EXPORT_BASE_YEAR
        export_final_year = EXPORT_FINAL_YEAR
        export_years = list(range(export_base_year, export_final_year + 1))
        print_sector_rows(
            esto_data,
            "LNG liquefaction/regas rows",
            {"economy": economy, "sub2sectors": lng_sub2},
            year_cols,
            start_year,
            code_to_name_mapping,
        )

        def _zero_series():
            return pd.Series({year: 0.0 for year in export_years}, dtype=float)

        def _mask_loss_values_by_year(loss_values_by_year, mask_series):
            masked = {}
            if not loss_values_by_year:
                return masked
            for label, series in loss_values_by_year.items():
                if not isinstance(series, dict):
                    continue
                selected = {}
                for year in export_years:
                    if not bool(mask_series.get(year, False)):
                        continue
                    value = series.get(year, series.get(str(year), 0.0))
                    if value is None or pd.isna(value):
                        continue
                    selected[year] = abs(float(value))
                if selected:
                    masked[label] = selected
            return masked

        def _accumulate_loss_values(target, updates):
            if not updates:
                return
            for label, series in updates.items():
                if not isinstance(series, dict):
                    continue
                destination = target.setdefault(label, {})
                for year, value in series.items():
                    if value is None or pd.isna(value):
                        continue
                    year_int = int(year)
                    destination[year_int] = destination.get(year_int, 0.0) + abs(float(value))

        feedstock_method = resolve_feedstock_method()
        regas_output_series = _zero_series()
        liquefaction_output_series = _zero_series()
        regas_input_series_by_label = {}
        liq_input_series_by_label = {}
        regas_loss_values_by_year = {}
        liq_loss_values_by_year = {}

        normalized_economy = _normalize_economy_value(economy).upper()
        aggregate_mode = "ALL" in normalized_economy or normalized_economy == "00APEC"
        if not aggregate_mode and "economy" in esto_data.columns:
            year_cols_from_start = get_years_from(year_cols, start_year)
            aggregate_rows = select_rows(
                esto_data,
                {"economy": economy, "sub2sectors": lng_sub2},
            )
            non_aggregate_rows = esto_data[
                esto_data["economy"]
                .apply(_normalize_economy_value)
                .ne(_normalize_economy_value(economy))
            ]
            non_aggregate_rows = select_rows(
                non_aggregate_rows,
                {"sub2sectors": lng_sub2},
            )
            aggregate_total = sum_years(aggregate_rows, year_cols_from_start)
            non_aggregate_total = sum_years(non_aggregate_rows, year_cols_from_start)
            tolerance = 1e-6 * max(1.0, abs(aggregate_total), abs(non_aggregate_total))
            if abs(aggregate_total - non_aggregate_total) <= tolerance:
                aggregate_mode = True
        component_economies = [economy]
        if aggregate_mode and "economy" in esto_data.columns:
            available_economies = sorted(esto_data["economy"].dropna().astype(str).unique())
            non_aggregate = [
                value
                for value in available_economies
                if _normalize_economy_value(value).upper() != normalized_economy
            ]
            if non_aggregate:
                component_economies = non_aggregate
                print(
                    "LNG aggregate mode: classifying each economy-year before summing "
                    f"({len(component_economies)} economies)."
                )

        regas_direction_count = 0
        liq_direction_count = 0
        ambiguous_direction_count = 0

        def _accumulate_series_map(target, label, series):
            if series is None or series.empty:
                return
            existing = target.get(label)
            if existing is None or existing.empty:
                target[label] = series.copy()
            else:
                target[label] = existing.add(series, fill_value=0.0)

        for component_economy in component_economies:
            ng_series_raw = ensure_full_year_series(
                sum_years_by_year(
                    select_rows(
                        esto_data,
                        {
                            "economy": component_economy,
                            "sub2sectors": lng_sub2,
                            "subfuels": fuel_codes["natural_gas"],
                        },
                    ),
                    year_cols,
                    start_year,
                ),
                export_base_year,
                export_final_year,
            )
            lng_series_raw = ensure_full_year_series(
                sum_years_by_year(
                    select_rows(
                        esto_data,
                        {
                            "economy": component_economy,
                            "sub2sectors": lng_sub2,
                            "subfuels": fuel_codes["lng"],
                        },
                    ),
                    year_cols,
                    start_year,
                ),
                export_base_year,
                export_final_year,
            )
            process_rows = select_rows(
                esto_data,
                {"economy": component_economy, "sub2sectors": lng_sub2},
            )
            timeseries, _ = summarize_fuel_timeseries(
                process_rows,
                year_cols,
                start_year,
                allow_all_years_fallback=True,
            )

            regas_mask = (ng_series_raw > 0) & (lng_series_raw < 0)
            liq_mask = (lng_series_raw > 0) & (ng_series_raw < 0)
            activity_mask = ng_series_raw.ne(0) | lng_series_raw.ne(0)
            ambiguous_mask = activity_mask & ~(regas_mask | liq_mask)

            regas_direction_count += int(regas_mask.sum())
            liq_direction_count += int(liq_mask.sum())
            ambiguous_direction_count += int(ambiguous_mask.sum())

            regas_output_series = regas_output_series.add(
                to_output_only_series(ng_series_raw.where(regas_mask, 0.0)),
                fill_value=0.0,
            )

            liquefaction_output_series = liquefaction_output_series.add(
                to_output_only_series(lng_series_raw.where(liq_mask, 0.0)),
                fill_value=0.0,
            )
            if timeseries is not None and not timeseries.empty:
                for label in timeseries.index:
                    series_full = ensure_full_year_series(
                        timeseries.loc[label],
                        export_base_year,
                        export_final_year,
                    )
                    regas_input = to_input_only_series(series_full.where(regas_mask, 0.0))
                    liq_input = to_input_only_series(series_full.where(liq_mask, 0.0))
                    if regas_input.sum() > 0:
                        _accumulate_series_map(regas_input_series_by_label, label, regas_input)
                    if liq_input.sum() > 0:
                        _accumulate_series_map(liq_input_series_by_label, label, liq_input)

            if not (regas_mask.any() or liq_mask.any()):
                continue

            _, _, _, loss_values_by_year = build_loss_context(
                loss_data,
                loss_year_cols,
                start_year,
                component_economy,
                "lng",
                lng_sub2,
            )
            _accumulate_loss_values(
                regas_loss_values_by_year,
                _mask_loss_values_by_year(loss_values_by_year, regas_mask),
            )
            _accumulate_loss_values(
                liq_loss_values_by_year,
                _mask_loss_values_by_year(loss_values_by_year, liq_mask),
            )

        print(
            "LNG direction summary: "
            f"regas economy-years={regas_direction_count}, "
            f"liq economy-years={liq_direction_count}, "
            f"ambiguous economy-years={ambiguous_direction_count}"
        )
        if ambiguous_direction_count > 0:
            print(
                "Warning: LNG ambiguous direction years were excluded from both processes."
            )

        def _aux_ratios_from_inputs(input_series_by_label, output_series, labels):
            ratios = {}
            output_series_pos = to_output_only_series(output_series)
            for label in labels:
                series = input_series_by_label.get(label)
                if series is None:
                    continue
                ratio_series = safe_divide_series(series, output_series_pos)
                ratios[label] = ratio_series.to_dict()
            return ratios

        def _build_lng_process(process_name, output_label, output_series, input_series_by_label, loss_values_by_year):
            if output_series.sum() == 0 and not input_series_by_label:
                return
            sector_title = (
                lng_config.get("regasification_title", "LNG regasification")
                if str(process_name).strip().lower() == "regasification"
                else lng_config.get("liquefaction_title", lng_config.get("title", "NG Liquefaction"))
            )
            total_input_series = build_total_input_series(input_series_by_label, export_years)
            total_loss_by_year = sum_loss_values_by_year(loss_values_by_year, export_years)
            output_import_targets_total, output_export_targets_total = gather_output_target_dicts(
                economy,
                [output_label],
                export_base_year,
                export_final_year,
                output_series_by_fuel={output_label: output_series},
            )
            if feedstock_method == FEEDSTOCK_METHOD_SPLIT:
                feedstock_labels = list(input_series_by_label.keys())
                for idx, feedstock_label in enumerate(feedstock_labels):
                    input_series = input_series_by_label[feedstock_label]
                    share_series = build_input_share_series(
                        input_series,
                        total_input_series,
                        fallback_to_one=(idx == 0),
                    )
                    allocated_output_series = output_series.mul(share_series, fill_value=0.0)
                    allocated_loss_values = allocate_loss_by_share(
                        loss_values_by_year,
                        share_series,
                    )
                    loss_total_for_eff = total_loss_by_year.mul(share_series, fill_value=0.0)
                    efficiency_series = compute_efficiency_by_year(
                        allocated_output_series,
                        input_series,
                        loss_total_for_eff,
                    )
                    auxiliary_fuels, auxiliary_ratios = build_auxiliary_from_losses_by_year(
                        allocated_loss_values,
                        allocated_output_series,
                        feedstock_labels=[feedstock_label],
                    )
                    output_import_targets = {
                        label: scale_year_dict_by_share(values, share_series)
                        for label, values in (output_import_targets_total or {}).items()
                    }
                    output_export_targets = {
                        label: scale_year_dict_by_share(values, share_series)
                        for label, values in (output_export_targets_total or {}).items()
                    }
                    process_label = (
                        process_name
                        if len(feedstock_labels) == 1
                        else f"{process_name} - {feedstock_label}"
                    )
                    print_leap_structure_block(
                        f"LNG {process_label}",
                        [output_label],
                        process_label,
                        [feedstock_label],
                        auxiliary_fuels,
                        loss_fuels=list(allocated_loss_values.keys()),
                        code_to_name_mapping=code_to_name_mapping,
                        output_fuel_values={output_label: allocated_output_series.sum()},
                        process_value=f"{efficiency_series.mean():.4f}",
                        feedstock_fuel_values={feedstock_label: input_series.sum()},
                        auxiliary_fuel_values={
                            label: summarize_numeric_value(values, summary="mean")
                            for label, values in auxiliary_ratios.items()
                        },
                        loss_fuel_values={
                            label: summarize_numeric_value(values, summary="sum")
                            for label, values in allocated_loss_values.items()
                        },
                    )
                    append_process_record(
                        process_records,
                        build_process_record(
                            economy,
                            sector_title,
                            process_label,
                            {
                                output_label: series_to_year_dict(
                                    allocated_output_series, export_base_year, export_final_year
                                )
                            },
                            {
                                feedstock_label: series_to_year_dict(
                                    input_series, export_base_year, export_final_year
                                )
                            },
                            series_to_year_dict(
                                efficiency_series, export_base_year, export_final_year
                            ),
                            auxiliary_ratios,
                            allocated_loss_values,
                            float(pd.Series(loss_total_for_eff, dtype=float).sum()),
                            loss_values_for_efficiency=series_to_year_dict(
                                loss_total_for_eff,
                                export_base_year,
                                export_final_year,
                            ),
                            feedstock_shares={feedstock_label: 1.0},
                            input_total=float(input_series.sum()),
                            output_import_targets=output_import_targets,
                            output_export_targets=output_export_targets,
                        ),
                    )
                return
            if feedstock_method == FEEDSTOCK_METHOD_MULTI:
                auxiliary_fuels, auxiliary_ratios = build_auxiliary_from_losses_by_year(
                    loss_values_by_year,
                    output_series,
                    feedstock_labels=list(input_series_by_label.keys()) if LOSS_AUX_EXCLUDE_FEEDSTOCKS else None,
                )
                efficiency_series = compute_efficiency_by_year(
                    output_series,
                    total_input_series,
                    total_loss_by_year,
                )
                feedstock_values = {
                    label: series_to_year_dict(series, export_base_year, export_final_year)
                    for label, series in input_series_by_label.items()
                }
                feedstock_shares = {
                    label: safe_divide_series(series, total_input_series).to_dict()
                    for label, series in input_series_by_label.items()
                }
                print_leap_structure_block(
                    f"LNG {process_name}",
                    [output_label],
                    process_name,
                    list(feedstock_values.keys()),
                    auxiliary_fuels,
                    loss_fuels=list(loss_values_by_year.keys()),
                    code_to_name_mapping=code_to_name_mapping,
                    output_fuel_values={output_label: output_series.sum()},
                    process_value=f"{efficiency_series.mean():.4f}",
                    feedstock_fuel_values={
                        label: summarize_numeric_value(values, summary="sum")
                        for label, values in feedstock_values.items()
                    },
                    auxiliary_fuel_values={
                        label: summarize_numeric_value(values, summary="mean")
                        for label, values in auxiliary_ratios.items()
                    },
                    loss_fuel_values={
                        label: summarize_numeric_value(values, summary="sum")
                        for label, values in loss_values_by_year.items()
                    },
                )
                append_process_record(
                    process_records,
                    build_process_record(
                        economy,
                        sector_title,
                        process_name,
                        {
                            output_label: series_to_year_dict(
                                output_series, export_base_year, export_final_year
                            )
                        },
                        feedstock_values,
                        series_to_year_dict(
                            efficiency_series, export_base_year, export_final_year
                        ),
                        auxiliary_ratios,
                        loss_values_by_year,
                        float(pd.Series(total_loss_by_year, dtype=float).sum()),
                        loss_values_for_efficiency=series_to_year_dict(
                            total_loss_by_year,
                            export_base_year,
                            export_final_year,
                        ),
                        feedstock_shares=feedstock_shares,
                        input_total=float(total_input_series.sum()),
                        output_import_targets=output_import_targets_total,
                        output_export_targets=output_export_targets_total,
                        own_use_ratios=compute_own_use_ratios_for_record(
                            loss_values_by_year, input_series_by_label, export_years
                        ),
                    ),
                )
                return

            if not input_series_by_label:
                return
            primary_input = max(
                input_series_by_label,
                key=lambda label: input_series_by_label[label].sum(),
            )
            input_series = input_series_by_label.get(primary_input)
            if input_series is None or input_series.sum() == 0:
                return
            other_feedstocks = [
                label for label in input_series_by_label.keys() if label != primary_input
            ]
            auxiliary_fuels = list(other_feedstocks)
            auxiliary_ratios = _aux_ratios_from_inputs(
                input_series_by_label,
                output_series,
                auxiliary_fuels,
            )
            auxiliary_fuels, auxiliary_ratios = merge_loss_into_auxiliary_by_year(
                auxiliary_fuels,
                auxiliary_ratios,
                loss_values_by_year,
                output_series,
                primary_input,
            )
            efficiency_series = compute_efficiency_by_year(
                output_series,
                total_input_series,
                total_loss_by_year,
            )
            print_leap_structure_block(
                f"LNG {process_name}",
                [output_label],
                process_name,
                [primary_input],
                auxiliary_fuels,
                loss_fuels=list(loss_values_by_year.keys()),
                code_to_name_mapping=code_to_name_mapping,
                output_fuel_values={output_label: output_series.sum()},
                process_value=f"{efficiency_series.mean():.4f}",
                feedstock_fuel_values={primary_input: input_series.sum()},
                auxiliary_fuel_values={
                    label: summarize_numeric_value(values, summary="mean")
                    for label, values in auxiliary_ratios.items()
                },
                loss_fuel_values={
                    label: summarize_numeric_value(values, summary="sum")
                    for label, values in loss_values_by_year.items()
                },
            )
            append_process_record(
                process_records,
                build_process_record(
                    economy,
                    sector_title,
                    process_name,
                    {
                        output_label: series_to_year_dict(
                            output_series, export_base_year, export_final_year
                        )
                    },
                    {
                        primary_input: series_to_year_dict(
                            input_series, export_base_year, export_final_year
                        )
                    },
                    series_to_year_dict(
                        efficiency_series, export_base_year, export_final_year
                    ),
                    auxiliary_ratios,
                    loss_values_by_year,
                    float(pd.Series(total_loss_by_year, dtype=float).sum()),
                    loss_values_for_efficiency=series_to_year_dict(
                        total_loss_by_year,
                        export_base_year,
                        export_final_year,
                    ),
                    feedstock_shares={primary_input: 1.0},
                    input_total=float(total_input_series.sum()),
                    output_import_targets=output_import_targets_total,
                    output_export_targets=output_export_targets_total,
                ),
            )

        regas_output_total = regas_output_series.sum()
        if regas_output_total > 0 and regas_input_series_by_label:
            _build_lng_process(
                "Regasification",
                fuel_codes["natural_gas"],
                regas_output_series,
                regas_input_series_by_label,
                regas_loss_values_by_year,
            )

        liquefaction_output_total = liquefaction_output_series.sum()
        if liquefaction_output_total > 0 and liq_input_series_by_label:
            _build_lng_process(
                "Liquefaction",
                fuel_codes["lng"],
                liquefaction_output_series,
                liq_input_series_by_label,
                liq_loss_values_by_year,
            )
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

        def _build_gas_process_records(
            sector_title,
            process_name,
            output_label,
            process_rows,
            output_rows,
            loss_sub2_code,
        ):
            if process_rows.empty:
                return
            feedstock_method = resolve_feedstock_method()
            totals, _ = summarize_fuel_totals(
                process_rows, year_cols, start_year, allow_all_years_fallback=True
            )
            timeseries, _ = summarize_fuel_timeseries(
                process_rows, year_cols, start_year, allow_all_years_fallback=True
            )
            negative = totals[totals < 0]
            if negative.empty or timeseries.empty:
                print(f"{process_name}: no negative inputs available.")
                return
            output_series = ensure_full_year_series(
                sum_years_by_year(output_rows, year_cols, start_year),
                export_base_year,
                export_final_year,
            )
            if output_series.sum() == 0:
                print(f"{process_name}: no output series available.")
                return
            input_series_map, zero_sum_labels = build_input_series_map(
                timeseries,
                list(negative.index),
                export_base_year,
                export_final_year,
            )
            if zero_sum_labels:
                log_dropped_input_fuels(economy, str(process_name), zero_sum_labels, export_base_year, export_final_year)
            if not input_series_map:
                print(f"{process_name}: no input series after normalization.")
                return
            total_input_series = build_total_input_series(input_series_map, export_years)
            loss_series, loss_total, loss_values, loss_values_by_year = build_loss_context(
                loss_data,
                loss_year_cols,
                start_year,
                economy,
                "gas_processing",
                loss_sub2_code,
            )
            total_loss_by_year = sum_loss_values_by_year(loss_values_by_year, export_years)
            output_import_targets_total, output_export_targets_total = gather_output_target_dicts(
                economy,
                [output_label],
                export_base_year,
                export_final_year,
                output_series_by_fuel={output_label: output_series},
            )

            if feedstock_method == FEEDSTOCK_METHOD_SPLIT:
                feedstock_labels = list(input_series_map.keys())
                for idx, feedstock_label in enumerate(feedstock_labels):
                    input_series = input_series_map[feedstock_label]
                    share_series = build_input_share_series(
                        input_series,
                        total_input_series,
                        fallback_to_one=(idx == 0),
                    )
                    allocated_output_series = output_series.mul(share_series, fill_value=0.0)
                    allocated_loss_values = allocate_loss_by_share(
                        loss_values_by_year,
                        share_series,
                    )
                    loss_total_for_eff = total_loss_by_year.mul(share_series, fill_value=0.0)
                    efficiency_series = compute_efficiency_by_year(
                        allocated_output_series,
                        input_series,
                        loss_total_for_eff,
                    )
                    auxiliary_fuels, auxiliary_ratios = build_auxiliary_from_losses_by_year(
                        allocated_loss_values,
                        allocated_output_series,
                        feedstock_labels=[feedstock_label],
                    )
                    output_import_targets = {
                        label: scale_year_dict_by_share(values, share_series)
                        for label, values in (output_import_targets_total or {}).items()
                    }
                    output_export_targets = {
                        label: scale_year_dict_by_share(values, share_series)
                        for label, values in (output_export_targets_total or {}).items()
                    }
                    process_label = (
                        process_name
                        if len(feedstock_labels) == 1
                        else f"{process_name} - {feedstock_label}"
                    )
                    print_leap_structure_block(
                        process_label,
                        [output_label],
                        process_label,
                        [feedstock_label],
                        auxiliary_fuels,
                        loss_fuels=list(allocated_loss_values.keys()),
                        code_to_name_mapping=code_to_name_mapping,
                        output_fuel_values={output_label: allocated_output_series.sum()},
                        process_value=f"{efficiency_series.mean():.6f}",
                        feedstock_fuel_values={feedstock_label: input_series.sum()},
                        auxiliary_fuel_values={
                            label: summarize_numeric_value(values, summary="mean")
                            for label, values in auxiliary_ratios.items()
                        },
                        loss_fuel_values={
                            label: summarize_numeric_value(values, summary="sum")
                            for label, values in allocated_loss_values.items()
                        },
                    )
                    record = build_process_record(
                        economy,
                        sector_title,
                        process_label,
                        {output_label: series_to_year_dict(allocated_output_series, export_base_year, export_final_year)},
                        {feedstock_label: series_to_year_dict(input_series, export_base_year, export_final_year)},
                        series_to_year_dict(efficiency_series, export_base_year, export_final_year),
                        auxiliary_ratios,
                        allocated_loss_values,
                        float(pd.Series(loss_total_for_eff, dtype=float).sum()),
                        loss_values_for_efficiency=series_to_year_dict(
                            loss_total_for_eff,
                            export_base_year,
                            export_final_year,
                        ),
                        feedstock_shares={feedstock_label: 1.0},
                        input_total=float(input_series.sum()),
                        output_import_targets=output_import_targets,
                        output_export_targets=output_export_targets,
                    )
                    append_process_record(process_records, record)
            elif feedstock_method == FEEDSTOCK_METHOD_MULTI:
                auxiliary_fuels, auxiliary_ratios = build_auxiliary_from_losses_by_year(
                    loss_values_by_year,
                    output_series,
                    feedstock_labels=list(input_series_map.keys()) if LOSS_AUX_EXCLUDE_FEEDSTOCKS else None,
                )
                efficiency_series = compute_efficiency_by_year(
                    output_series,
                    total_input_series,
                    total_loss_by_year,
                )
                feedstock_values = {
                    label: series_to_year_dict(series, export_base_year, export_final_year)
                    for label, series in input_series_map.items()
                }
                feedstock_shares = {
                    label: safe_divide_series(series, total_input_series).to_dict()
                    for label, series in input_series_map.items()
                }
                print_leap_structure_block(
                    process_name,
                    [output_label],
                    process_name,
                    list(feedstock_values.keys()),
                    auxiliary_fuels,
                    loss_fuels=list(loss_values_by_year.keys()),
                    code_to_name_mapping=code_to_name_mapping,
                    output_fuel_values={output_label: output_series.sum()},
                    process_value=f"{efficiency_series.mean():.6f}",
                    feedstock_fuel_values={
                        label: summarize_numeric_value(values, summary="sum")
                        for label, values in feedstock_values.items()
                    },
                    auxiliary_fuel_values={
                        label: summarize_numeric_value(values, summary="mean")
                        for label, values in auxiliary_ratios.items()
                    },
                    loss_fuel_values={
                        label: summarize_numeric_value(values, summary="sum")
                        for label, values in loss_values_by_year.items()
                    },
                )
                record = build_process_record(
                    economy,
                    sector_title,
                    process_name,
                    {output_label: series_to_year_dict(output_series, export_base_year, export_final_year)},
                    feedstock_values,
                    series_to_year_dict(efficiency_series, export_base_year, export_final_year),
                    auxiliary_ratios,
                    loss_values_by_year,
                    loss_total,
                    loss_values_for_efficiency=series_to_year_dict(
                        total_loss_by_year,
                        export_base_year,
                        export_final_year,
                    ),
                    feedstock_shares=feedstock_shares,
                    input_total=float(total_input_series.sum()),
                    output_import_targets=output_import_targets_total,
                    output_export_targets=output_export_targets_total,
                    own_use_ratios=compute_own_use_ratios_for_record(
                        loss_values_by_year, input_series_map, export_years
                    ),
                )
                append_process_record(process_records, record)
            else:
                primary_input = negative.idxmin()
                input_series = input_series_map.get(primary_input)
                if input_series is None or input_series.empty:
                    print(f"{process_name}: primary input series missing after normalization.")
                    return
                other_feedstock_fuels = [
                    label for label in input_series_map.keys() if label != primary_input
                ]
                auxiliary_fuels = list(other_feedstock_fuels)
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
                efficiency_series = compute_efficiency_by_year(
                    output_series,
                    total_input_series,
                    total_loss_by_year,
                )
                print_leap_structure_block(
                    process_name,
                    [output_label],
                    process_name,
                    [primary_input],
                    auxiliary_fuels,
                    loss_fuels=list(loss_values_by_year.keys()),
                    code_to_name_mapping=code_to_name_mapping,
                    output_fuel_values={output_label: output_series.sum()},
                    process_value=f"{efficiency_series.mean():.6f}",
                    feedstock_fuel_values={primary_input: input_series.sum()},
                    auxiliary_fuel_values={
                        label: summarize_numeric_value(values, summary="mean")
                        for label, values in auxiliary_ratios.items()
                    },
                    loss_fuel_values={
                        label: summarize_numeric_value(values, summary="sum")
                        for label, values in loss_values_by_year.items()
                    },
                )
                record = build_process_record(
                    economy,
                    sector_title,
                    process_name,
                    {output_label: series_to_year_dict(output_series, export_base_year, export_final_year)},
                    {primary_input: series_to_year_dict(input_series, export_base_year, export_final_year)},
                    series_to_year_dict(efficiency_series, export_base_year, export_final_year),
                    auxiliary_ratios,
                    loss_values_by_year,
                    loss_total,
                    loss_values_for_efficiency=series_to_year_dict(
                        total_loss_by_year,
                        export_base_year,
                        export_final_year,
                    ),
                    feedstock_shares={primary_input: 1.0},
                    input_total=float(total_input_series.sum()),
                    output_import_targets=output_import_targets_total,
                    output_export_targets=output_export_targets_total,
                )
                append_process_record(process_records, record)

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
        output_label = fuel_codes["gas_works_gas"]
        if "products" in esto_data.columns:
            output_label = product_code_gas_works_gas
        gas_works_sector_title = str(
            gas_config.get("title", "Gas works plants")
        ).strip() or "Gas works plants"
        _build_gas_process_records(
            gas_works_sector_title,
            "Gas works plants",
            output_label,
            gas_works_rows,
            gas_works_output,
            gas_works_sub2,
        )

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
        output_label = fuel_codes["natural_gas"]
        if "products" in esto_data.columns:
            output_label = product_code_natural_gas
        blending_sector_title = str(
            MAJOR_SECTOR_CONFIG.get("gas_blending", {}).get(
                "title",
                "Natural gas blending plants",
            )
        ).strip() or "Natural gas blending plants"
        _build_gas_process_records(
            blending_sector_title,
            "Natural gas blending plants",
            output_label,
            blending_rows,
            blending_output,
            blending_sub2,
        )

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

            feedstock_method = resolve_feedstock_method()
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
            output_series_by_label = {primary_output: output_series}
            input_series_map, zero_sum_labels = build_input_series_map(
                timeseries,
                list(negative.index),
                export_base_year,
                export_final_year,
            )
            if zero_sum_labels:
                log_dropped_input_fuels(economy, flow_code, zero_sum_labels, export_base_year, export_final_year)
            if not input_series_map:
                print(f"[WARN] {flow_code}: no input series available after normalization")
                continue
            total_input_series = build_total_input_series(input_series_map, export_years)
            total_loss_by_year = sum_loss_values_by_year(loss_values_by_year, export_years)

            input_name = map_code_label(primary_input, code_to_name_mapping)
            output_name = map_code_label(primary_output, code_to_name_mapping)

            print(
                f"{flow_code}: output {output_name} ({output_total:.2f}), "
                f"input {input_name} ({-input_total:.2f})"
            )

            output_import_targets_total, output_export_targets_total = gather_output_target_dicts(
                economy,
                [primary_output],
                export_base_year,
                export_final_year,
                output_series_by_fuel=output_series_by_label,
            )

            if feedstock_method == FEEDSTOCK_METHOD_SPLIT:
                feedstock_labels = list(input_series_map.keys())
                for idx, feedstock_label in enumerate(feedstock_labels):
                    input_series = input_series_map[feedstock_label]
                    share_series = build_input_share_series(
                        input_series,
                        total_input_series,
                        fallback_to_one=(idx == 0),
                    )
                    allocated_outputs = allocate_outputs_by_share(
                        output_series_by_label,
                        share_series,
                    )
                    allocated_loss_values = allocate_loss_by_share(
                        loss_values_by_year,
                        share_series,
                    )
                    loss_total_for_eff = total_loss_by_year.mul(share_series, fill_value=0.0)
                    allocated_output_series = allocated_outputs.get(primary_output, pd.Series(dtype=float))
                    efficiency_series = compute_efficiency_by_year(
                        allocated_output_series,
                        input_series,
                        loss_total_for_eff,
                    )
                    auxiliary_fuels, auxiliary_ratios = build_auxiliary_from_losses_by_year(
                        allocated_loss_values,
                        allocated_output_series,
                        feedstock_labels=[feedstock_label],
                    )
                    output_import_targets = {
                        label: scale_year_dict_by_share(values, share_series)
                        for label, values in (output_import_targets_total or {}).items()
                    }
                    output_export_targets = {
                        label: scale_year_dict_by_share(values, share_series)
                        for label, values in (output_export_targets_total or {}).items()
                    }
                    process_name = (
                        flow_code
                        if len(feedstock_labels) == 1
                        else f"{flow_code} - {feedstock_label}"
                    )
                    print_leap_structure_block(
                        f"{title} - {process_name}",
                        [primary_output],
                        process_name,
                        [feedstock_label],
                        auxiliary_fuels,
                        loss_fuels=list(allocated_loss_values.keys()),
                        code_to_name_mapping=code_to_name_mapping,
                        output_fuel_values={
                            primary_output: summarize_numeric_value(
                                series_to_year_dict(allocated_output_series, export_base_year, export_final_year),
                                summary="sum",
                            )
                        },
                        process_value=f"{efficiency_series.mean():.4f}",
                        feedstock_fuel_values={
                            feedstock_label: summarize_numeric_value(
                                series_to_year_dict(input_series, export_base_year, export_final_year),
                                summary="sum",
                            )
                        },
                        auxiliary_fuel_values={
                            label: summarize_numeric_value(values, summary="mean")
                            for label, values in auxiliary_ratios.items()
                        },
                        loss_fuel_values={
                            label: summarize_numeric_value(values, summary="sum")
                            for label, values in allocated_loss_values.items()
                        },
                    )
                    record = build_process_record(
                        economy,
                        title,
                        process_name,
                        {
                            primary_output: series_to_year_dict(
                                allocated_output_series, export_base_year, export_final_year
                            )
                        },
                        {
                            feedstock_label: series_to_year_dict(
                                input_series, export_base_year, export_final_year
                            )
                        },
                        series_to_year_dict(efficiency_series, export_base_year, export_final_year),
                        auxiliary_ratios,
                        allocated_loss_values,
                        float(pd.Series(loss_total_for_eff, dtype=float).sum()),
                        loss_values_for_efficiency=series_to_year_dict(
                            loss_total_for_eff,
                            export_base_year,
                            export_final_year,
                        ),
                        feedstock_shares={feedstock_label: 1.0},
                        input_total=float(input_series.sum()),
                        output_import_targets=output_import_targets,
                        output_export_targets=output_export_targets,
                    )
                    append_process_record(process_records, record)
            elif feedstock_method == FEEDSTOCK_METHOD_MULTI:
                auxiliary_fuels, auxiliary_ratios = build_auxiliary_from_losses_by_year(
                    loss_values_by_year,
                    output_series,
                    feedstock_labels=list(input_series_map.keys()) if LOSS_AUX_EXCLUDE_FEEDSTOCKS else None,
                )
                loss_total_for_eff = total_loss_by_year
                efficiency_series = compute_efficiency_by_year(
                    output_series,
                    total_input_series,
                    loss_total_for_eff,
                )
                feedstock_values = {
                    label: series_to_year_dict(series, export_base_year, export_final_year)
                    for label, series in input_series_map.items()
                }
                feedstock_shares = {
                    label: safe_divide_series(series, total_input_series).to_dict()
                    for label, series in input_series_map.items()
                }
                print_leap_structure_block(
                    f"{title} - {flow_code}",
                    [primary_output],
                    flow_code,
                    list(feedstock_values.keys()),
                    auxiliary_fuels,
                    loss_fuels=list(loss_values_by_year.keys()),
                    code_to_name_mapping=code_to_name_mapping,
                    output_fuel_values={primary_output: output_total},
                    process_value=f"{efficiency_series.mean():.4f}",
                    feedstock_fuel_values={
                        label: summarize_numeric_value(values, summary="sum")
                        for label, values in feedstock_values.items()
                    },
                    auxiliary_fuel_values={
                        label: summarize_numeric_value(values, summary="mean")
                        for label, values in auxiliary_ratios.items()
                    },
                    loss_fuel_values={
                        label: summarize_numeric_value(values, summary="sum")
                        for label, values in loss_values_by_year.items()
                    },
                )
                record = build_process_record(
                    economy,
                    title,
                    flow_code,
                    {primary_output: series_to_year_dict(output_series, export_base_year, export_final_year)},
                    feedstock_values,
                    series_to_year_dict(efficiency_series, export_base_year, export_final_year),
                    auxiliary_ratios,
                    loss_values_by_year,
                    float(loss_total),
                    loss_values_for_efficiency=series_to_year_dict(
                        loss_total_for_eff,
                        export_base_year,
                        export_final_year,
                    ),
                    feedstock_shares=feedstock_shares,
                    input_total=float(total_input_series.sum()),
                    output_import_targets=output_import_targets_total,
                    output_export_targets=output_export_targets_total,
                    own_use_ratios=compute_own_use_ratios_for_record(
                        loss_values_by_year, input_series_map, export_years
                    ),
                )
                append_process_record(process_records, record)
            else:
                other_feedstock_fuels = [
                    label for label in input_series_map.keys() if label != primary_input
                ]
                auxiliary_fuels = list(other_feedstock_fuels)
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
                input_series = input_series_map.get(primary_input)
                if input_series is None or input_series.empty:
                    print(f"{flow_code}: primary input series missing after normalization")
                    continue
                loss_total_for_eff = total_loss_by_year
                # Compute efficiency from total process input so auxiliary
                # feedstocks are included in the denominator.
                efficiency_series = compute_efficiency_by_year(
                    output_series,
                    total_input_series,
                    loss_total_for_eff,
                )
                print_leap_structure_block(
                    f"{title} - {flow_code}",
                    [primary_output],
                    flow_code,
                    [primary_input],
                    auxiliary_fuels,
                    loss_fuels=list(loss_values_by_year.keys()),
                    code_to_name_mapping=code_to_name_mapping,
                    output_fuel_values={primary_output: output_total},
                    process_value=f"{efficiency_series.mean():.4f}",
                    feedstock_fuel_values={primary_input: input_total},
                    auxiliary_fuel_values={
                        label: summarize_numeric_value(values, summary="mean")
                        for label, values in auxiliary_ratios.items()
                    },
                    loss_fuel_values={
                        label: summarize_numeric_value(values, summary="sum")
                        for label, values in loss_values_by_year.items()
                    },
                )
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
                    loss_values_for_efficiency=series_to_year_dict(
                        loss_total_for_eff,
                        export_base_year,
                        export_final_year,
                    ),
                    feedstock_shares={primary_input: 1.0},
                    input_total=float(total_input_series.sum()),
                    output_import_targets=output_import_targets_total,
                    output_export_targets=output_export_targets_total,
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


def _coerce_label_list(value):
    """Return a clean list[str] from a scalar/list config value."""
    try:
        if value is None:
            return []
        if isinstance(value, list):
            raw_values = value
        else:
            raw_values = [value]
        labels = []
        for raw in raw_values:
            text = str(raw).strip()
            if text:
                labels.append(text)
        return labels
    except Exception as exc:
        print(f"Failed to coerce label list from {value}: {exc}")
        try_debug_breakpoint()
        raise


def _sum_series_collection(series_list, year_index):
    """Return the sum of aligned Series objects over year_index."""
    try:
        total = pd.Series({year: 0.0 for year in year_index}, dtype=float)
        for series in series_list:
            if series is None:
                continue
            aligned = series.reindex(year_index, fill_value=0.0)
            total = total.add(aligned, fill_value=0.0)
        return total
    except Exception as exc:
        print(f"Failed to sum series collection: {exc}")
        try_debug_breakpoint()
        raise


def _merge_year_value_maps(target_map, updates):
    """Merge dict[label][year] values into target_map by summing overlaps."""
    try:
        if not updates:
            return target_map
        for label, year_map in updates.items():
            if not isinstance(year_map, dict):
                continue
            destination = target_map.setdefault(label, {})
            for year, value in year_map.items():
                if value is None or pd.isna(value):
                    continue
                year_int = int(year)
                destination[year_int] = destination.get(year_int, 0.0) + float(value)
        return target_map
    except Exception as exc:
        print(f"Failed to merge year value maps: {exc}")
        try_debug_breakpoint()
        raise


def _build_hydrogen_loss_context(
    loss_data,
    loss_year_cols,
    start_year,
    economy,
    loss_sub2_codes,
):
    """Return (loss_values_by_year, loss_total_by_year) for configured loss sub2 codes."""
    try:
        if not loss_sub2_codes:
            return {}, pd.Series(dtype=float)
        merged_loss_values = {}
        for loss_sub2 in loss_sub2_codes:
            _, _, loss_values_by_year = summarize_own_use_losses_by_year(
                loss_data,
                loss_year_cols,
                start_year,
                economy,
                loss_sub2_code=loss_sub2,
                allow_all_years_fallback=True,
            )
            _merge_year_value_maps(merged_loss_values, loss_values_by_year)
        if not merged_loss_values:
            return {}, pd.Series(dtype=float)
        loss_totals = {}
        for year_map in merged_loss_values.values():
            for year, value in year_map.items():
                if value is None or pd.isna(value):
                    continue
                year_int = int(year)
                loss_totals[year_int] = loss_totals.get(year_int, 0.0) + abs(float(value))
        return merged_loss_values, pd.Series(loss_totals, dtype=float)
    except Exception as exc:
        print(f"Failed to build hydrogen loss context: {exc}")
        try_debug_breakpoint()
        raise


def analyze_hydrogen_transformation(
    data,
    year_cols,
    start_year,
    economy,
    code_to_name_mapping,
    loss_data,
    loss_year_cols,
    sector_config=None,
    process_records=None,
):
    """Build hydrogen transformation process records from 9th sub-sector data."""
    try:
        hydrogen_config = sector_config or MAJOR_SECTOR_CONFIG["hydrogen_transformation"]
        sector_title = hydrogen_config.get("title", "09_13_hydrogen_transformation")
        transformation_sub1 = hydrogen_config.get(
            "transformation_sub1",
            "09_13_hydrogen_transformation",
        )
        output_fuels = set(
            _coerce_label_list(hydrogen_config.get("output_fuels") or ["16_others"])
        )
        default_output_subfuels = _coerce_label_list(
            hydrogen_config.get("output_subfuels") or HYDROGEN_OUTPUT_SUBFUELS
        )
        process_config = hydrogen_config.get("process_config") or HYDROGEN_PROCESS_CONFIG
        if not process_config:
            print(f"{sector_title}: no hydrogen process config found; skipping.")
            return
        print(f"\n==== {sector_title} ({economy}) ====")
        if not has_required_columns(
            data,
            [["sub1sectors", "sub2sectors", "fuels", "subfuels"]],
            sector_title,
        ):
            return

        export_base_year = EXPORT_BASE_YEAR
        export_final_year = EXPORT_FINAL_YEAR
        export_years = list(range(export_base_year, export_final_year + 1))

        enabled_processes = []
        for raw_cfg in process_config:
            if not isinstance(raw_cfg, dict):
                continue
            if not raw_cfg.get("enabled", True):
                continue
            process_code = str(raw_cfg.get("process_code", "")).strip()
            source_sub2 = str(
                raw_cfg.get("source_sub2sectors") or raw_cfg.get("sub2sectors") or process_code
            ).strip()
            if not process_code or not source_sub2:
                continue
            cfg = dict(raw_cfg)
            cfg["process_code"] = process_code
            cfg["source_sub2sectors"] = source_sub2
            cfg["input_fuel_codes"] = _coerce_label_list(cfg.get("input_fuel_codes"))
            cfg["output_subfuels"] = _coerce_label_list(
                cfg.get("output_subfuels") or default_output_subfuels
            )
            cfg["loss_sub2sectors"] = _coerce_label_list(cfg.get("loss_sub2sectors"))
            enabled_processes.append(cfg)
        if not enabled_processes:
            print(f"{sector_title}: no enabled hydrogen process configs.")
            return

        process_groups = {}
        for cfg in enabled_processes:
            process_groups.setdefault(cfg["source_sub2sectors"], []).append(cfg)

        for source_sub2, group_configs in process_groups.items():
            process_rows = select_rows(
                data,
                {
                    "economy": economy,
                    "sub1sectors": transformation_sub1,
                    "sub2sectors": source_sub2,
                },
            )
            if process_rows.empty:
                continue

            print_sector_rows_from_df(
                process_rows,
                f"{sector_title} rows ({source_sub2})",
                year_cols,
                start_year,
                code_to_name_mapping,
            )
            timeseries, _ = summarize_fuel_timeseries(
                process_rows,
                year_cols,
                start_year,
                allow_all_years_fallback=True,
            )
            totals, _ = summarize_fuel_totals(
                process_rows,
                year_cols,
                start_year,
                allow_all_years_fallback=True,
            )
            if timeseries.empty:
                continue
            feedstock_method = resolve_feedstock_method()

            output_rows = process_rows
            if "fuels" in output_rows.columns and output_fuels:
                output_rows = output_rows[output_rows["fuels"].astype(str).isin(output_fuels)]
            output_timeseries, _ = summarize_fuel_timeseries(
                output_rows,
                year_cols,
                start_year,
                allow_all_years_fallback=True,
            )
            group_output_labels = []
            for cfg in group_configs:
                for label in cfg["output_subfuels"]:
                    if label not in group_output_labels:
                        group_output_labels.append(label)
            source_output_series = {}
            for output_label in group_output_labels:
                series = ensure_full_year_series(
                    to_output_only_series(get_label_timeseries(output_timeseries, output_label)),
                    export_base_year,
                    export_final_year,
                )
                if series.sum() == 0:
                    continue
                source_output_series[output_label] = series
            if not source_output_series:
                continue

            process_inputs = {}
            negative_labels = list(totals[totals < 0].index)
            input_series_map_all, zero_sum_labels = build_input_series_map(
                timeseries,
                negative_labels,
                export_base_year,
                export_final_year,
            )
            if zero_sum_labels:
                log_dropped_input_fuels(economy, "hydrogen_transformation", zero_sum_labels, export_base_year, export_final_year)
            for cfg in group_configs:
                input_series_by_label = dict(input_series_map_all)
                input_total_series = build_total_input_series(
                    input_series_by_label,
                    export_years,
                )
                process_inputs[cfg["process_code"]] = {
                    "config": cfg,
                    "input_series_by_label": input_series_by_label,
                    "input_total_series": input_total_series,
                }

            group_total_input = _sum_series_collection(
                [
                    details["input_total_series"]
                    for details in process_inputs.values()
                ],
                export_years,
            )
            process_codes = [cfg["process_code"] for cfg in group_configs]

            for index, process_code in enumerate(process_codes):
                details = process_inputs.get(process_code, {})
                cfg = details.get("config")
                input_series_by_label = details.get("input_series_by_label", {})
                input_total_series = details.get("input_total_series")
                if cfg is None or input_total_series is None:
                    continue
                if not input_series_by_label:
                    continue

                share_series = safe_divide_series(input_total_series, group_total_input)
                if index == 0:
                    zero_mask = group_total_input.eq(0.0)
                    if zero_mask.any():
                        share_series = share_series.where(~zero_mask, 1.0)

                output_series_by_label = {}
                for output_label, series in source_output_series.items():
                    allocated = series.mul(share_series, fill_value=0.0)
                    if allocated.sum() == 0:
                        continue
                    output_series_by_label[output_label] = allocated
                total_output_series = _sum_series_collection(
                    list(output_series_by_label.values()),
                    export_years,
                )

                if total_output_series.sum() == 0 and input_total_series.sum() == 0:
                    continue

                loss_values_by_year, _ = _build_hydrogen_loss_context(
                    loss_data,
                    loss_year_cols,
                    start_year,
                    economy,
                    cfg.get("loss_sub2sectors"),
                )
                total_loss_by_year = sum_loss_values_by_year(loss_values_by_year, export_years)
                output_import_targets_total, output_export_targets_total = gather_output_target_dicts(
                    economy,
                    list(output_series_by_label.keys()),
                    export_base_year,
                    export_final_year,
                    output_series_by_fuel=output_series_by_label,
                )

                if feedstock_method == FEEDSTOCK_METHOD_SPLIT:
                    feedstock_labels = list(input_series_by_label.keys())
                    for idx, feedstock_label in enumerate(feedstock_labels):
                        input_series = input_series_by_label[feedstock_label]
                        feedstock_share_series = build_input_share_series(
                            input_series,
                            input_total_series,
                            fallback_to_one=(idx == 0),
                        )
                        allocated_outputs = allocate_outputs_by_share(
                            output_series_by_label,
                            feedstock_share_series,
                        )
                        allocated_output_total = _sum_series_collection(
                            list(allocated_outputs.values()),
                            export_years,
                        )
                        allocated_loss_values = allocate_loss_by_share(
                            loss_values_by_year,
                            feedstock_share_series,
                        )
                        loss_total_for_eff = total_loss_by_year.mul(
                            feedstock_share_series,
                            fill_value=0.0,
                        )
                        efficiency_series = compute_efficiency_by_year(
                            allocated_output_total,
                            input_series,
                            loss_total_for_eff,
                        )
                        auxiliary_fuels, auxiliary_ratios = build_auxiliary_from_losses_by_year(
                            allocated_loss_values,
                            allocated_output_total,
                            feedstock_labels=[feedstock_label],
                        )
                        output_import_targets = {
                            label: scale_year_dict_by_share(values, feedstock_share_series)
                            for label, values in (output_import_targets_total or {}).items()
                        }
                        output_export_targets = {
                            label: scale_year_dict_by_share(values, feedstock_share_series)
                            for label, values in (output_export_targets_total or {}).items()
                        }
                        process_name = (
                            cfg["process_code"]
                            if len(feedstock_labels) == 1
                            else f"{cfg['process_code']} - {feedstock_label}"
                        )
                        print_leap_structure_block(
                            f"{sector_title} - {process_name}",
                            list(allocated_outputs.keys()),
                            process_name,
                            [feedstock_label],
                            auxiliary_fuels,
                            loss_fuels=list(allocated_loss_values.keys()),
                            code_to_name_mapping=code_to_name_mapping,
                            output_fuel_values={
                                label: series.sum() for label, series in allocated_outputs.items()
                            },
                            process_value=f"{efficiency_series.mean():.4f}",
                            feedstock_fuel_values={feedstock_label: input_series.sum()},
                            auxiliary_fuel_values={
                                label: summarize_numeric_value(values, summary="mean")
                                for label, values in auxiliary_ratios.items()
                            },
                            loss_fuel_values={
                                label: summarize_numeric_value(values, summary="sum")
                                for label, values in allocated_loss_values.items()
                            },
                        )
                        append_process_record(
                            process_records,
                            build_process_record(
                                economy,
                                sector_title,
                                process_name,
                                {
                                    label: series_to_year_dict(
                                        series, export_base_year, export_final_year
                                    )
                                    for label, series in allocated_outputs.items()
                                },
                                {
                                    feedstock_label: series_to_year_dict(
                                        input_series, export_base_year, export_final_year
                                    )
                                },
                                series_to_year_dict(
                                    efficiency_series,
                                    export_base_year,
                                    export_final_year,
                                ),
                                auxiliary_ratios,
                                allocated_loss_values,
                                float(pd.Series(loss_total_for_eff, dtype=float).sum()),
                                loss_values_for_efficiency=series_to_year_dict(
                                    loss_total_for_eff,
                                    export_base_year,
                                    export_final_year,
                                ),
                                feedstock_shares={feedstock_label: 1.0},
                                input_total=input_series.sum(),
                                output_import_targets=output_import_targets,
                                output_export_targets=output_export_targets,
                            ),
                        )
                elif feedstock_method == FEEDSTOCK_METHOD_MULTI:
                    auxiliary_fuels, auxiliary_ratios = build_auxiliary_from_losses_by_year(
                        loss_values_by_year,
                        total_output_series,
                        feedstock_labels=list(input_series_by_label.keys()) if LOSS_AUX_EXCLUDE_FEEDSTOCKS else None,
                    )
                    efficiency_series = compute_efficiency_by_year(
                        total_output_series,
                        input_total_series,
                        total_loss_by_year,
                    )
                    feedstock_values = {
                        label: series_to_year_dict(series, export_base_year, export_final_year)
                        for label, series in input_series_by_label.items()
                    }
                    feedstock_shares = {
                        label: safe_divide_series(series, input_total_series).to_dict()
                        for label, series in input_series_by_label.items()
                    }
                    output_values = {
                        label: series_to_year_dict(series, export_base_year, export_final_year)
                        for label, series in output_series_by_label.items()
                    }
                    print_leap_structure_block(
                        f"{sector_title} - {cfg['process_code']}",
                        list(output_values.keys()),
                        cfg["process_code"],
                        list(feedstock_values.keys()),
                        auxiliary_fuels,
                        loss_fuels=list(loss_values_by_year.keys()),
                        code_to_name_mapping=code_to_name_mapping,
                        output_fuel_values={
                            label: series.sum() for label, series in output_series_by_label.items()
                        },
                        process_value=f"{efficiency_series.mean():.4f}",
                        feedstock_fuel_values={
                            label: series.sum() for label, series in input_series_by_label.items()
                        },
                        auxiliary_fuel_values={
                            label: summarize_numeric_value(values, summary="mean")
                            for label, values in auxiliary_ratios.items()
                        },
                        loss_fuel_values={
                            label: summarize_numeric_value(values, summary="sum")
                            for label, values in loss_values_by_year.items()
                        },
                    )
                    append_process_record(
                        process_records,
                        build_process_record(
                            economy,
                            sector_title,
                            cfg["process_code"],
                            output_values,
                            feedstock_values,
                            series_to_year_dict(
                                efficiency_series,
                                export_base_year,
                                export_final_year,
                            ),
                            auxiliary_ratios,
                            loss_values_by_year,
                            float(pd.Series(total_loss_by_year, dtype=float).sum()),
                            loss_values_for_efficiency=series_to_year_dict(
                                total_loss_by_year,
                                export_base_year,
                                export_final_year,
                            ),
                            feedstock_shares=feedstock_shares,
                            input_total=input_total_series.sum(),
                            output_import_targets=output_import_targets_total,
                            output_export_targets=output_export_targets_total,
                            own_use_ratios=compute_own_use_ratios_for_record(
                                loss_values_by_year, input_series_by_label, export_years
                            ),
                        ),
                    )
                else:
                    primary_input = max(
                        input_series_by_label,
                        key=lambda label: input_series_by_label[label].sum(),
                    )
                    input_series = input_series_by_label.get(primary_input)
                    if input_series is None or input_series.sum() == 0:
                        continue
                    auxiliary_fuels = [
                        label for label in input_series_by_label.keys() if label != primary_input
                    ]
                    auxiliary_ratios = build_auxiliary_ratios_by_year(
                        timeseries,
                        auxiliary_fuels,
                        total_output_series,
                    )
                    auxiliary_fuels, auxiliary_ratios = merge_loss_into_auxiliary_by_year(
                        auxiliary_fuels,
                        auxiliary_ratios,
                        loss_values_by_year,
                        total_output_series,
                        primary_input,
                    )
                    # Use total process input for efficiency so years where the
                    # chosen primary fuel is zero do not inflate efficiency.
                    efficiency_series = compute_efficiency_by_year(
                        total_output_series,
                        input_total_series,
                        total_loss_by_year,
                    )
                    output_values = {
                        label: series_to_year_dict(series, export_base_year, export_final_year)
                        for label, series in output_series_by_label.items()
                    }
                    print_leap_structure_block(
                        f"{sector_title} - {cfg['process_code']}",
                        list(output_values.keys()),
                        cfg["process_code"],
                        [primary_input],
                        auxiliary_fuels,
                        loss_fuels=list(loss_values_by_year.keys()),
                        code_to_name_mapping=code_to_name_mapping,
                        output_fuel_values={
                            label: series.sum() for label, series in output_series_by_label.items()
                        },
                        process_value=f"{efficiency_series.mean():.4f}",
                        feedstock_fuel_values={primary_input: input_series.sum()},
                        auxiliary_fuel_values={
                            label: summarize_numeric_value(values, summary="mean")
                            for label, values in auxiliary_ratios.items()
                        },
                        loss_fuel_values={
                            label: summarize_numeric_value(values, summary="sum")
                            for label, values in loss_values_by_year.items()
                        },
                    )
                    append_process_record(
                        process_records,
                        build_process_record(
                            economy,
                            sector_title,
                            cfg["process_code"],
                            output_values,
                            {primary_input: series_to_year_dict(input_series, export_base_year, export_final_year)},
                            series_to_year_dict(
                                efficiency_series,
                                export_base_year,
                                export_final_year,
                            ),
                            auxiliary_ratios,
                            loss_values_by_year,
                            float(pd.Series(total_loss_by_year, dtype=float).sum()),
                            loss_values_for_efficiency=series_to_year_dict(
                                total_loss_by_year,
                                export_base_year,
                                export_final_year,
                            ),
                            feedstock_shares={primary_input: 1.0},
                            input_total=input_series.sum(),
                            output_import_targets=output_import_targets_total,
                            output_export_targets=output_export_targets_total,
                        ),
                    )
    except Exception as exc:
        print(f"Hydrogen transformation analysis failed: {exc}")
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
        # Register this sector so zero-fill can clear its catalog branches even for
        # economies that had no ESTO data (no process record produced).
        sector_title = map_code_label(sector_config.get("title", ""), code_to_name_mapping)
        if sector_title:
            _analyzed_sector_titles.add(sector_title)
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


def run_hydrogen_transformation_analysis(
    data, year_cols, economy, loss_data, loss_year_cols, sector_config, process_records
):
    """Run hydrogen transformation analysis for a single economy."""
    analyze_hydrogen_transformation(
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
#%%

#%%
######### CONSTANTS (LIKELY TO CHANGE) #########
# Editable workflow settings live in `codebase/workflow_config.py`.
RUN_LNG_ANALYSIS = workflow_cfg.TRANSFORMATION_RUN_LNG_ANALYSIS
RUN_GAS_PROCESSING_ANALYSIS = workflow_cfg.TRANSFORMATION_RUN_GAS_PROCESSING_ANALYSIS
RUN_COAL_TRANSFORMATION_ANALYSIS = workflow_cfg.TRANSFORMATION_RUN_COAL_TRANSFORMATION_ANALYSIS
RUN_OTHER_TRANSFORMATION_ANALYSIS = workflow_cfg.TRANSFORMATION_RUN_OTHER_TRANSFORMATION_ANALYSIS
RUN_CHARCOAL_PROCESSING_ANALYSIS = workflow_cfg.TRANSFORMATION_RUN_CHARCOAL_PROCESSING_ANALYSIS
RUN_NONSPECIFIED_TRANSFORMATION_ANALYSIS = (
    workflow_cfg.TRANSFORMATION_RUN_NONSPECIFIED_TRANSFORMATION_ANALYSIS
)
RUN_OIL_REFINERY_ANALYSIS = workflow_cfg.TRANSFORMATION_RUN_OIL_REFINERY_ANALYSIS
RUN_HYDROGEN_TRANSFORMATION_ANALYSIS = (
    workflow_cfg.TRANSFORMATION_RUN_HYDROGEN_TRANSFORMATION_ANALYSIS
)
ALL_ECONOMY_LABEL = workflow_cfg.TRANSFORMATION_ALL_ECONOMY_LABEL
INCLUDE_ALL_FEEDSTOCKS_AS_AUXILIARY = (
    workflow_cfg.TRANSFORMATION_INCLUDE_ALL_FEEDSTOCKS_AS_AUXILIARY
)
ECONOMIES_TO_ANALYZE = list(workflow_cfg.TRANSFORMATION_ECONOMIES_TO_ANALYZE)
SAVE_ESTO_SUBTOTAL_LABELED = workflow_cfg.TRANSFORMATION_SAVE_ESTO_SUBTOTAL_LABELED
ESTO_SUBTOTAL_LABELED_OUTPUT_PATH = (
    workflow_cfg.TRANSFORMATION_ESTO_SUBTOTAL_LABELED_OUTPUT_PATH
)
BUILD_LEAP_EXPORT = workflow_cfg.TRANSFORMATION_BUILD_LEAP_EXPORT
SAVE_LEAP_EXPORT_FILE = workflow_cfg.TRANSFORMATION_SAVE_LEAP_EXPORT_FILE
SAVE_SUMMARY_TABLES = workflow_cfg.TRANSFORMATION_SAVE_SUMMARY_TABLES
EXPORT_OUTPUT_DIR = workflow_cfg.TRANSFORMATION_EXPORT_OUTPUT_DIR
EXPORT_FILENAME_FALLBACK = workflow_cfg.TRANSFORMATION_EXPORT_FILENAME_FALLBACK
EXPORT_FILENAME_TEMPLATE = workflow_cfg.TRANSFORMATION_EXPORT_FILENAME_TEMPLATE
EXPORT_MODEL_NAME = workflow_cfg.TRANSFORMATION_EXPORT_MODEL_NAME
EXPORT_REGION = workflow_cfg.TRANSFORMATION_EXPORT_REGION
SCENARIOS_TO_EXPORT = list(workflow_cfg.TRANSFORMATION_SCENARIOS_TO_EXPORT)
EXPORT_BASE_YEAR = int(workflow_cfg.TRANSFORMATION_EXPORT_BASE_YEAR)
_config_final_year = workflow_cfg.TRANSFORMATION_EXPORT_FINAL_YEAR
EXPORT_FINAL_YEAR = (
    PROJECTION_END_YEAR if _config_final_year is None else int(_config_final_year)
)
SUMMARY_OUTPUT_DIR = workflow_cfg.TRANSFORMATION_SUMMARY_OUTPUT_DIR
PROCESS_SUMMARY_FILENAME = workflow_cfg.TRANSFORMATION_PROCESS_SUMMARY_FILENAME
DETAIL_SUMMARY_FILENAME = workflow_cfg.TRANSFORMATION_DETAIL_SUMMARY_FILENAME

# Skip emitting output series rows that LEAP does not expect.
INCLUDE_OUTPUT_SERIES_IN_LEAP_EXPORT = (
    workflow_cfg.TRANSFORMATION_INCLUDE_OUTPUT_SERIES_IN_LEAP_EXPORT
)

SAVE_PROJECTION_DIAGNOSTICS = workflow_cfg.TRANSFORMATION_SAVE_PROJECTION_DIAGNOSTICS
PROJECTION_DIAGNOSTICS_PATH = (
    workflow_cfg.TRANSFORMATION_PROJECTION_DIAGNOSTICS_PATH
)

SCENARIO_EXPORT_OVERRIDES = workflow_cfg.TRANSFORMATION_SCENARIO_EXPORT_OVERRIDES


def get_scenario_export_config(scenario, default_base_year=None, default_final_year=None):
    """Return export overrides for the given scenario."""
    overrides = SCENARIO_EXPORT_OVERRIDES.get(scenario, {})
    scenario_name = str(scenario or "").strip().lower()
    is_current_accounts = scenario_name in {"current accounts", "current account"}
    resolved_base_year = (
        EXPORT_BASE_YEAR if default_base_year is None else int(default_base_year)
    )
    resolved_final_year = (
        EXPORT_FINAL_YEAR if default_final_year is None else int(default_final_year)
    )
    default_final_year = resolved_base_year if is_current_accounts else resolved_final_year
    return {
        "export_base_year": overrides.get("export_base_year", resolved_base_year),
        "export_final_year": overrides.get("export_final_year", default_final_year),
        "include_current_account_rows": overrides.get(
            "include_current_account_rows",
            is_current_accounts,
        ),
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
workflow_common.archive_config_dir_once_per_day()
esto_data_raw, ninth_data_raw = load_augmented_reference_tables(
    esto_path=MATT_DATA_PATH,
    ninth_path=ESTO_DATA_PATH,
    subtotal_mapping_path=SUBTOTAL_MAPPING_PATH,
    synthetic_rules_path=CONFIG_DIR / "synthetic_reference_rows.csv",
    cache_dir=REFERENCE_CACHE_DIR,
    apply_esto_subtotal_map=True,
    filter_esto_subtotals_flag=False,
    filter_ninth_subtotals_flag=False,
)
print(
    f"Loaded ESTO data (augmented): {esto_data_raw.shape[0]} rows, {esto_data_raw.shape[1]} columns"
)
print(
    f"Loaded 9th data (augmented): {ninth_data_raw.shape[0]} rows, {ninth_data_raw.shape[1]} columns"
)
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
# if SAVE_ESTO_SUBTOTAL_LABELED:
#     save_subtotal_labeled_data(
#         esto_data_with_subtotals,
#         ESTO_SUBTOTAL_LABELED_OUTPUT_PATH,
#         "ESTO (Matt) data",
#     )
esto_data = filter_matt_subtotals(esto_data_with_subtotals)
_should_aggregate, _aggregate_label, _ = workflow_common.resolve_aggregate_economy(
    ECONOMIES_TO_ANALYZE,
    aggregate_label=ALL_ECONOMY_LABEL,
)
if _should_aggregate:
    ninth_data = add_all_economy_total(ninth_data, ninth_year_cols, _aggregate_label)
    esto_data = add_all_economy_total(esto_data, esto_year_cols, _aggregate_label)
projection_sign_stable_flows = resolve_projection_sign_stable_flows(
    PROJECTION_SIGN_STABLE_MODE,
    SIGN_STABLE_PROJECTION_FLOWS,
)
projection_df, projection_diagnostics = build_esto_projection_table(
    ninth_data=ninth_data,
    esto_data=esto_data,
    mapping_path=NINTH_TO_ESTO_MAPPING_PATH,
    base_year=BASE_YEAR,
    projection_years=PROJECTION_YEAR_RANGE,
    sign_stable_flows=projection_sign_stable_flows,
    strict_conservation=PROJECTION_STRICT_CONSERVATION,
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
# Transformation\NG Liquefaction\Processes\Regasification\Auxiliary Fuels\Electricity Auxiliary Fuel Use Current Accounts United States Petajoule Petajoule 0.00428 Transformation NG Liquefaction Processes Regasification Auxiliary Fuels Electricity
# Transformation\NG Liquefaction\Processes\Regasification\Auxiliary Fuels\Electricity Auxiliary Fuel Use Reference United States Petajoule Petajoule 0.00428 Transformation NG Liquefaction Processes Regasification Auxiliary Fuels Electricity
# Transformation\NG Liquefaction\Processes\Regasification\Auxiliary Fuels\Electricity Auxiliary Fuel Use Target United States Petajoule Petajoule 0.00428 Transformation NG Liquefaction Processes Regasification Auxiliary Fuels Electricity
#%%

#%%
# LNG done basically. Next up is the other gas processing sectors:
# Using ESTO data here (Matt file does not include sector detail columns).
#%%
