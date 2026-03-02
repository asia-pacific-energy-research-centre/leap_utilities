# Centralized workflow configuration for interactive runs and notebooks.
from __future__ import annotations

from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]

###############################
# CORE SETTINGS (LIKELY TO EDIT)
###############################
# These drive most workflows. Set these first.

# Accepts a string or list. Use "00_APEC" (or "ALL_ECONOMIES") to trigger aggregation.
def _normalize_economies(value) -> list[str]:
    """Return a normalized economy list from a string/iterable input."""
    if value is None:
        return []
    if isinstance(value, str):
        text = value.strip()
        return [text] if text else []
    return [str(item).strip() for item in value if str(item).strip()]


def _resolve_global_aggregate(economies: list[str]) -> str:
    """Infer the aggregate economy label from a normalized economy list."""
    if len(economies) == 1 and economies[0] in {"00_APEC", "ALL_ECONOMIES", "ALL"}:
        return economies[0]
    return "ALL_ECONOMIES"


GLOBAL_ECONOMIES = _normalize_economies(["00_APEC"])
# Multiple economies -> per-economy runs (no aggregation unless a sentinel is used).
GLOBAL_SCENARIOS = ["Reference", "Target", "Current Accounts"]
GLOBAL_BASE_YEAR = 2022
# None means: fall back to the workflow/module default final year.
GLOBAL_FINAL_YEAR = None
GLOBAL_REGION = "United States of America"
GLOBAL_EXPORT_OUTPUT_DIR = REPO_ROOT / "outputs" / "leap_exports"
GLOBAL_AGGREGATE_ECONOMY_LABEL = _resolve_global_aggregate(GLOBAL_ECONOMIES)

#########################
# HELPER FUNCTIONS (RARE)
#########################
# Only needed if you want to customize file naming behavior.

def _format_scenario_token(scenarios: list[str]) -> str:
    """Return a filename-friendly scenario string."""
    sanitized = "_".join(
        "".join(ch for ch in str(scenario) if ch.isalnum()) for scenario in scenarios
    )
    return sanitized or "scenarios"


def _default_supply_export_filename(
    economies: list[str], scenarios: list[str]
) -> str:
    """Return a default supply export filename aligned to current globals."""
    economy_label = economies[0] if economies else "ALL"
    scenario_token = _format_scenario_token(scenarios)
    return f"supply_leap_imports_{economy_label}_{scenario_token}.xlsx"

###############################
# WORKFLOW SETTINGS (LESS USED)
###############################
# Everything below can be left alone unless a workflow needs a special override.

###############################
# TRANSFORMATION ANALYSIS
###############################
TRANSFORMATION_RUN_LNG_ANALYSIS = True
TRANSFORMATION_RUN_GAS_PROCESSING_ANALYSIS = True
TRANSFORMATION_RUN_COAL_TRANSFORMATION_ANALYSIS = True
TRANSFORMATION_RUN_CHARCOAL_PROCESSING_ANALYSIS = True
TRANSFORMATION_RUN_NONSPECIFIED_TRANSFORMATION_ANALYSIS = True
TRANSFORMATION_RUN_HYDROGEN_TRANSFORMATION_ANALYSIS = True
TRANSFORMATION_ALL_ECONOMY_LABEL = GLOBAL_AGGREGATE_ECONOMY_LABEL
TRANSFORMATION_INCLUDE_ALL_FEEDSTOCKS_AS_AUXILIARY = True
TRANSFORMATION_ECONOMIES_TO_ANALYZE = list(GLOBAL_ECONOMIES)
TRANSFORMATION_SAVE_ESTO_SUBTOTAL_LABELED = True
TRANSFORMATION_ESTO_SUBTOTAL_LABELED_OUTPUT_PATH = (
    "data/00APEC_2024_low_with_subtotals.csv"
)
TRANSFORMATION_BUILD_LEAP_EXPORT = True
TRANSFORMATION_SAVE_LEAP_EXPORT_FILE = True
TRANSFORMATION_SAVE_SUMMARY_TABLES = True
TRANSFORMATION_EXPORT_OUTPUT_DIR = GLOBAL_EXPORT_OUTPUT_DIR
TRANSFORMATION_EXPORT_FILENAME_FALLBACK = "transformation_leap_imports.xlsx"
TRANSFORMATION_EXPORT_FILENAME_TEMPLATE = (
    "transformation_leap_imports_{economy}_{scenario}.xlsx"
)
TRANSFORMATION_EXPORT_MODEL_NAME = "LEAP Transformation Imports"
TRANSFORMATION_EXPORT_REGION = GLOBAL_REGION
TRANSFORMATION_SCENARIOS_TO_EXPORT = list(GLOBAL_SCENARIOS)
TRANSFORMATION_EXPORT_BASE_YEAR = GLOBAL_BASE_YEAR
# None means: fall back to PROJECTION_END_YEAR in transformation_analysis_utils.
TRANSFORMATION_EXPORT_FINAL_YEAR = GLOBAL_FINAL_YEAR
TRANSFORMATION_SUMMARY_OUTPUT_DIR = TRANSFORMATION_EXPORT_OUTPUT_DIR
TRANSFORMATION_PROCESS_SUMMARY_FILENAME = "transformation_process_summary.csv"
TRANSFORMATION_DETAIL_SUMMARY_FILENAME = "transformation_detail_summary.csv"
TRANSFORMATION_INCLUDE_OUTPUT_SERIES_IN_LEAP_EXPORT = False
TRANSFORMATION_SAVE_PROJECTION_DIAGNOSTICS = False
TRANSFORMATION_PROJECTION_DIAGNOSTICS_PATH = (
    Path(TRANSFORMATION_EXPORT_OUTPUT_DIR) / "ninth_projection_allocation_fallbacks.csv"
)
TRANSFORMATION_SCENARIO_EXPORT_OVERRIDES = {
    "Current Accounts": {
        "include_current_account_rows": True,
    },
    "Current Account": {
        "include_current_account_rows": True,
    },
}

###############################
# TRANSFORMATION WORKFLOW (WRAPPER)
###############################
TRANSFORMATION_WORKFLOW_SHEET_NAME = "LEAP"
TRANSFORMATION_WORKFLOW_EXPORT_FILENAME_PREFIX = "transformation_leap_imports_"
TRANSFORMATION_WORKFLOW_DEFAULT_SCENARIOS = list(GLOBAL_SCENARIOS)

###############################
# HYDROGEN TRANSFORMATION
###############################
HYDROGEN_SHEET_NAME = "LEAP"
HYDROGEN_EXPORT_FILENAME_PREFIX = "hydrogen_transformation_leap_imports_"
HYDROGEN_EXPORT_FILENAME_TEMPLATE = (
    "hydrogen_transformation_leap_imports_{economy}_{scenario}.xlsx"
)
HYDROGEN_EXPORT_FILENAME_FALLBACK = "hydrogen_transformation_leap_imports.xlsx"
HYDROGEN_PROCESS_SUMMARY_FILENAME = "hydrogen_transformation_process_summary.csv"
HYDROGEN_DETAIL_SUMMARY_FILENAME = "hydrogen_transformation_detail_summary.csv"
HYDROGEN_DEFAULT_SCENARIOS = list(GLOBAL_SCENARIOS)

###############################
# SUPPLY DATA PIPELINE
###############################
SUPPLY_RUN_SUPPLY_ANALYSIS = True
SUPPLY_RUN_LIST_FUELS = True
SUPPLY_ALL_ECONOMY_LABEL = GLOBAL_AGGREGATE_ECONOMY_LABEL
SUPPLY_ECONOMIES_TO_ANALYZE = list(GLOBAL_ECONOMIES)
SUPPLY_SAVE_ESTO_SUBTOTAL_LABELED = False
SUPPLY_ESTO_SUBTOTAL_LABELED_OUTPUT_PATH = "data/00APEC_2024_low_with_subtotals.csv"
SUPPLY_EXPORT_DATASET_KEY = "esto"
SUPPLY_EXPORT_DIR = GLOBAL_EXPORT_OUTPUT_DIR
SUPPLY_EXPORT_FILE_NAME = _default_supply_export_filename(
    GLOBAL_ECONOMIES,
    GLOBAL_SCENARIOS,
)
SUPPLY_SCENARIO_TO_RUN = "Target"
SUPPLY_FILL_BRANCHES_FROM_EXPORT_FILE = True
SUPPLY_HANDLE_CURRENT_ACCOUNTS_TOO = True
SUPPLY_RUN_SUPPLY_LEAP_IMPORT = False
SUPPLY_SHEET_NAME = "LEAP"
# When False, the supply export will not write the "Unmet Requirements" measure,
# leaving it for manual setup in LEAP.
SUPPLY_INCLUDE_UNMET_REQUIREMENTS = False

###############################
# SUPPLY WORKFLOW (WRAPPER)
###############################
SUPPLY_WORKFLOW_DEFAULT_ECONOMIES = list(GLOBAL_ECONOMIES)
SUPPLY_WORKFLOW_DEFAULT_SCENARIOS = list(GLOBAL_SCENARIOS)
SUPPLY_NOTEBOOK_ECONOMIES = list(GLOBAL_ECONOMIES)
SUPPLY_NOTEBOOK_INCLUDE_LEAP_IMPORT = True
SUPPLY_NOTEBOOK_SCENARIOS = list(GLOBAL_SCENARIOS)

###############################
# TRANSFERS WORKFLOW
###############################
TRANSFERS_EXPORT_FILENAME_TEMPLATE = "transfer_leap_imports_{economy}_{scenario}.xlsx"
TRANSFERS_EXPORT_FILENAME_PREFIX = "transfer_leap_imports_"
TRANSFERS_SHEET_NAME = "LEAP"
TRANSFERS_DEFAULT_SCENARIOS = list(GLOBAL_SCENARIOS)
TRANSFERS_AGGREGATE_ECONOMY_LABEL = GLOBAL_AGGREGATE_ECONOMY_LABEL
# None means: fall back to core.ECONOMIES_TO_ANALYZE or DEFAULT_SCENARIOS.
TRANSFERS_NOTEBOOK_ECONOMIES = None
TRANSFERS_NOTEBOOK_SCENARIOS = None
# None means: fall back to leap_api availability in transfers_workflow.
TRANSFERS_NOTEBOOK_INCLUDE_LEAP_IMPORT = None
TRANSFERS_NOTEBOOK_CURRENT_ACCOUNTS = True

###############################
# MINOR DEMAND WORKFLOW
###############################
MINOR_DEMAND_EXPORT_FILENAME_TEMPLATE = (
    "outputs/leap_exports/minor_demand_export_{economy}_{scenario}.xlsx"
)
MINOR_DEMAND_EXPORT_MODEL_NAME = "Minor demand import"
MINOR_DEMAND_EXPORT_REGION = GLOBAL_REGION
MINOR_DEMAND_EXPORT_SCENARIOS = list(GLOBAL_SCENARIOS)
MINOR_DEMAND_NINTH_SCENARIO = "reference"
MINOR_DEMAND_MEASURE_UNITS = {
    "Activity Level": {"units": "Petajoule", "scale": "", "per": ""},
    "Final Energy Intensity": {"units": "Petajoule", "scale": "", "per": ""},
}
# Fuel-activity behavior for codebase/minor_demand_workflow.py:
# - "activity_as_energy_intensity_as_one" (default): fuel Activity Level is in
#   energy units and equals sector_activity * fuel_share, so fuel activities
#   sum back to sector activity.
#   Compatibility alias: "fuel_share" (legacy name for the same behavior).
#   Activity Level units are written as "Unspecified Unit" for both sector and
#   fuel branches.
# - "sector_share": fuel Activity Level is a dimensionless fraction (0..1) of
#   sector total energy for each year. In this mode, minor_demand_workflow
#   forces Final Energy Intensity to 1.0 for fuel branches and writes fuel
#   Activity Level units as "Share". Sector Activity Level units are written
#   as "Unspecified Unit".
# - "none": skip fuel Activity Level rows (still writes fuel intensity rows
#   where applicable).
MINOR_DEMAND_FUEL_ACTIVITY_MODE = "sector_share"
MINOR_DEMAND_INTENSITY_CALIBRATION_MODE = "disabled"
MINOR_DEMAND_RELATIVE_INTENSITY_OVERRIDES = {}
MINOR_DEMAND_AGGREGATE_ECONOMY_LABEL = GLOBAL_AGGREGATE_ECONOMY_LABEL
MINOR_DEMAND_DEMAND_ROOT_PARTS = ["Demand", "Other sector"]
# None means: fall back to BASE_YEAR in minor_demand_workflow.
MINOR_DEMAND_EXPORT_BASE_YEAR = GLOBAL_BASE_YEAR
MINOR_DEMAND_PROJECTION_START_YEAR = 2023
MINOR_DEMAND_EXPORT_FINAL_YEAR = 2060

###############################
# FULL MODEL WORKFLOW NOTEBOOK
###############################
FULL_MODEL_RUN_TRANSFORMATION_WORKFLOW = True
FULL_MODEL_RUN_HYDROGEN_TRANSFORMATION_WORKFLOW = True
FULL_MODEL_RUN_SUPPLY_WORKFLOW = True
FULL_MODEL_RUN_TRANSFERS_WORKFLOW = True
FULL_MODEL_RUN_MINOR_DEMAND_WORKFLOW = True
FULL_MODEL_RUN_INDUSTRY_MAPPING_WORKFLOW = True

# Feedstock method for transformation + hydrogen + transfers (None keeps module default).
FULL_MODEL_FEEDSTOCK_METHOD = "multi_feedstock_single_process"

# Transformation workflow config
FULL_MODEL_TRANSFORMATION_ECONOMIES = list(GLOBAL_ECONOMIES)
FULL_MODEL_TRANSFORMATION_SCENARIOS = list(GLOBAL_SCENARIOS)
# None means: fall back to LEAP API availability in full_model_workflow_notebook.
FULL_MODEL_TRANSFORMATION_INCLUDE_LEAP_IMPORT = None
FULL_MODEL_TRANSFORMATION_AGGREGATE_ECONOMY_LABEL = GLOBAL_AGGREGATE_ECONOMY_LABEL
FULL_MODEL_TRANSFORMATION_EXPORT_DIR = None
FULL_MODEL_TRANSFORMATION_FILENAME_TEMPLATE = None

# Hydrogen transformation workflow config
FULL_MODEL_HYDROGEN_TRANSFORMATION_ECONOMIES = list(GLOBAL_ECONOMIES)
FULL_MODEL_HYDROGEN_TRANSFORMATION_SCENARIOS = list(GLOBAL_SCENARIOS)
FULL_MODEL_HYDROGEN_TRANSFORMATION_INCLUDE_LEAP_IMPORT = None
FULL_MODEL_HYDROGEN_TRANSFORMATION_HANDLE_CURRENT_ACCOUNTS = True
FULL_MODEL_HYDROGEN_TRANSFORMATION_AGGREGATE_ECONOMY_LABEL = GLOBAL_AGGREGATE_ECONOMY_LABEL
FULL_MODEL_HYDROGEN_TRANSFORMATION_EXPORT_DIR = None
FULL_MODEL_HYDROGEN_TRANSFORMATION_FILENAME_TEMPLATE = None

# Supply workflow config
FULL_MODEL_SUPPLY_ECONOMIES = list(GLOBAL_ECONOMIES)
FULL_MODEL_SUPPLY_SCENARIOS = list(GLOBAL_SCENARIOS)
FULL_MODEL_SUPPLY_INCLUDE_LEAP_IMPORT = None
FULL_MODEL_SUPPLY_EXPORT_DATASET_KEY = "esto"

# Transfers workflow config
FULL_MODEL_TRANSFERS_ECONOMIES = list(GLOBAL_ECONOMIES)
FULL_MODEL_TRANSFERS_SCENARIOS = list(GLOBAL_SCENARIOS)
FULL_MODEL_TRANSFERS_INCLUDE_LEAP_IMPORT = None
FULL_MODEL_TRANSFERS_HANDLE_CURRENT_ACCOUNTS = True
FULL_MODEL_TRANSFERS_INCLUDE_OUTPUT_SERIES = False
FULL_MODEL_TRANSFERS_USE_OUTPUT_TARGETS = True
FULL_MODEL_TRANSFERS_AGGREGATE_ECONOMY_LABEL = GLOBAL_AGGREGATE_ECONOMY_LABEL

# Minor demand workflow config
FULL_MODEL_MINOR_DEMAND_ECONOMIES = list(GLOBAL_ECONOMIES)
FULL_MODEL_MINOR_DEMAND_SCENARIOS = list(GLOBAL_SCENARIOS)
FULL_MODEL_MINOR_DEMAND_REGION = GLOBAL_REGION
FULL_MODEL_MINOR_DEMAND_INCLUDE_LEAP_IMPORT = None
FULL_MODEL_MINOR_DEMAND_AGGREGATE_ECONOMY_LABEL = GLOBAL_AGGREGATE_ECONOMY_LABEL
FULL_MODEL_MINOR_DEMAND_EXPORT_FILENAME = None

# Industry mapping workflow config
FULL_MODEL_INDUSTRY_EXPORT_PATH = REPO_ROOT / "data" / "industry export.xlsx"
FULL_MODEL_INDUSTRY_SHEET_NAME = "Export"
FULL_MODEL_INDUSTRY_ECONOMY = "20_USA"
FULL_MODEL_INDUSTRY_BASE_YEAR = GLOBAL_BASE_YEAR
FULL_MODEL_INDUSTRY_SCENARIO = "Target"
FULL_MODEL_INDUSTRY_REGION = GLOBAL_REGION
