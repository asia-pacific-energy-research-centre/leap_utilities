#%%
"""
Full model workflow notebook runner.

Open this file as a notebook (VS Code: Python Interactive) and run all cells
to build a full LEAP model in one go.

After running, use `codebase/full_model_workflow_notebook_post_run_guide.md`
to complete manual LEAP checks/actions (units, skipped variables, branch gaps).
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

#%%
# --- Repo setup ---
REPO_ROOT = Path(__file__).resolve().parents[1]
if Path.cwd() != REPO_ROOT:
    os.chdir(REPO_ROOT)
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

#%%
# --- Global toggles ---
# Most user-editable settings live in `codebase/workflow_config.py`.
from codebase import leap_api
from codebase.utilities import workflow_common
from codebase.configuration import workflow_config as workflow_cfg

LEAP_API_AVAILABLE = leap_api.is_available()

RUN_TRANSFORMATION_WORKFLOW = workflow_cfg.FULL_MODEL_RUN_TRANSFORMATION_WORKFLOW
RUN_HYDROGEN_TRANSFORMATION_WORKFLOW = (
    workflow_cfg.FULL_MODEL_RUN_HYDROGEN_TRANSFORMATION_WORKFLOW
)
RUN_SUPPLY_WORKFLOW = workflow_cfg.FULL_MODEL_RUN_SUPPLY_WORKFLOW
RUN_TRANSFERS_WORKFLOW = workflow_cfg.FULL_MODEL_RUN_TRANSFERS_WORKFLOW
RUN_MINOR_DEMAND_WORKFLOW = workflow_cfg.FULL_MODEL_RUN_MINOR_DEMAND_WORKFLOW
RUN_INDUSTRY_MAPPING_WORKFLOW = workflow_cfg.FULL_MODEL_RUN_INDUSTRY_MAPPING_WORKFLOW

#%%
# --- Feedstock method (applies to transformation + hydrogen + transfers) ---
# Options:
# - "single_feedstock_aux_others"
# - "split_processes_per_feedstock"
# - "multi_feedstock_single_process"
# Use None to keep the module default.
FEEDSTOCK_METHOD = workflow_cfg.FULL_MODEL_FEEDSTOCK_METHOD


def _default_import_scenarios(scenarios: list[str]) -> list[str]:
    """Return non-current-account scenarios as lowercase labels."""
    current_accounts_labels = {"current accounts", "current account"}
    return [
        str(scenario).strip().lower()
        for scenario in scenarios
        if str(scenario).strip()
        and str(scenario).strip().lower() not in current_accounts_labels
    ]

#%%
# --- Transformation workflow config ---
TRANSFORMATION_ECONOMIES = workflow_cfg.FULL_MODEL_TRANSFORMATION_ECONOMIES
TRANSFORMATION_SCENARIOS = list(workflow_cfg.FULL_MODEL_TRANSFORMATION_SCENARIOS)
TRANSFORMATION_INCLUDE_LEAP_IMPORT = (
    LEAP_API_AVAILABLE
    if workflow_cfg.FULL_MODEL_TRANSFORMATION_INCLUDE_LEAP_IMPORT is None
    else workflow_cfg.FULL_MODEL_TRANSFORMATION_INCLUDE_LEAP_IMPORT
)
TRANSFORMATION_IMPORT_SCENARIOS = _default_import_scenarios(TRANSFORMATION_SCENARIOS)
TRANSFORMATION_AGGREGATE_ECONOMY_LABEL = (
    workflow_cfg.FULL_MODEL_TRANSFORMATION_AGGREGATE_ECONOMY_LABEL
)
TRANSFORMATION_EXPORT_DIR = workflow_cfg.FULL_MODEL_TRANSFORMATION_EXPORT_DIR
TRANSFORMATION_FILENAME_TEMPLATE = workflow_cfg.FULL_MODEL_TRANSFORMATION_FILENAME_TEMPLATE

#%%
# --- Hydrogen transformation workflow config ---
HYDROGEN_TRANSFORMATION_ECONOMIES = (
    workflow_cfg.FULL_MODEL_HYDROGEN_TRANSFORMATION_ECONOMIES
)
HYDROGEN_TRANSFORMATION_SCENARIOS = list(
    workflow_cfg.FULL_MODEL_HYDROGEN_TRANSFORMATION_SCENARIOS
)
HYDROGEN_TRANSFORMATION_INCLUDE_LEAP_IMPORT = (
    LEAP_API_AVAILABLE
    if workflow_cfg.FULL_MODEL_HYDROGEN_TRANSFORMATION_INCLUDE_LEAP_IMPORT is None
    else workflow_cfg.FULL_MODEL_HYDROGEN_TRANSFORMATION_INCLUDE_LEAP_IMPORT
)
HYDROGEN_TRANSFORMATION_IMPORT_SCENARIOS = _default_import_scenarios(HYDROGEN_TRANSFORMATION_SCENARIOS)
HYDROGEN_TRANSFORMATION_HANDLE_CURRENT_ACCOUNTS = (
    workflow_cfg.FULL_MODEL_HYDROGEN_TRANSFORMATION_HANDLE_CURRENT_ACCOUNTS
)
HYDROGEN_TRANSFORMATION_AGGREGATE_ECONOMY_LABEL = (
    workflow_cfg.FULL_MODEL_HYDROGEN_TRANSFORMATION_AGGREGATE_ECONOMY_LABEL
)
HYDROGEN_TRANSFORMATION_EXPORT_DIR = (
    workflow_cfg.FULL_MODEL_HYDROGEN_TRANSFORMATION_EXPORT_DIR
)
HYDROGEN_TRANSFORMATION_FILENAME_TEMPLATE = (
    workflow_cfg.FULL_MODEL_HYDROGEN_TRANSFORMATION_FILENAME_TEMPLATE
)

#%%
# --- Supply workflow config ---
SUPPLY_ECONOMIES = list(workflow_cfg.FULL_MODEL_SUPPLY_ECONOMIES)
SUPPLY_SCENARIOS = list(workflow_cfg.FULL_MODEL_SUPPLY_SCENARIOS)
SUPPLY_INCLUDE_LEAP_IMPORT = (
    LEAP_API_AVAILABLE
    if workflow_cfg.FULL_MODEL_SUPPLY_INCLUDE_LEAP_IMPORT is None
    else workflow_cfg.FULL_MODEL_SUPPLY_INCLUDE_LEAP_IMPORT
)
SUPPLY_IMPORT_SCENARIOS = _default_import_scenarios(SUPPLY_SCENARIOS)
SUPPLY_EXPORT_DATASET_KEY = workflow_cfg.FULL_MODEL_SUPPLY_EXPORT_DATASET_KEY

#%%
# --- Transfers workflow config ---
TRANSFERS_ECONOMIES = workflow_cfg.FULL_MODEL_TRANSFERS_ECONOMIES
TRANSFERS_SCENARIOS = list(workflow_cfg.FULL_MODEL_TRANSFERS_SCENARIOS)
TRANSFERS_INCLUDE_LEAP_IMPORT = (
    LEAP_API_AVAILABLE
    if workflow_cfg.FULL_MODEL_TRANSFERS_INCLUDE_LEAP_IMPORT is None
    else workflow_cfg.FULL_MODEL_TRANSFERS_INCLUDE_LEAP_IMPORT
)
TRANSFERS_IMPORT_SCENARIOS = _default_import_scenarios(TRANSFERS_SCENARIOS)
TRANSFERS_HANDLE_CURRENT_ACCOUNTS = workflow_cfg.FULL_MODEL_TRANSFERS_HANDLE_CURRENT_ACCOUNTS
TRANSFERS_INCLUDE_OUTPUT_SERIES = workflow_cfg.FULL_MODEL_TRANSFERS_INCLUDE_OUTPUT_SERIES
TRANSFERS_USE_OUTPUT_TARGETS = workflow_cfg.FULL_MODEL_TRANSFERS_USE_OUTPUT_TARGETS
TRANSFERS_AGGREGATE_ECONOMY_LABEL = workflow_cfg.FULL_MODEL_TRANSFERS_AGGREGATE_ECONOMY_LABEL

#%%
# --- Minor demand workflow config ---
MINOR_DEMAND_ECONOMIES = list(workflow_cfg.FULL_MODEL_MINOR_DEMAND_ECONOMIES)
MINOR_DEMAND_SCENARIOS = list(workflow_cfg.FULL_MODEL_MINOR_DEMAND_SCENARIOS)
MINOR_DEMAND_IMPORT_SCENARIOS = _default_import_scenarios(MINOR_DEMAND_SCENARIOS)
MINOR_DEMAND_REGION = workflow_cfg.FULL_MODEL_MINOR_DEMAND_REGION
MINOR_DEMAND_INCLUDE_LEAP_IMPORT = (
    LEAP_API_AVAILABLE
    if workflow_cfg.FULL_MODEL_MINOR_DEMAND_INCLUDE_LEAP_IMPORT is None
    else workflow_cfg.FULL_MODEL_MINOR_DEMAND_INCLUDE_LEAP_IMPORT
)
MINOR_DEMAND_AGGREGATE_ECONOMY_LABEL = (
    workflow_cfg.FULL_MODEL_MINOR_DEMAND_AGGREGATE_ECONOMY_LABEL
)
MINOR_DEMAND_EXPORT_FILENAME = workflow_cfg.FULL_MODEL_MINOR_DEMAND_EXPORT_FILENAME

#%%
# --- Industry mapping workflow config ---
INDUSTRY_EXPORT_PATH = workflow_cfg.FULL_MODEL_INDUSTRY_EXPORT_PATH
INDUSTRY_SHEET_NAME = workflow_cfg.FULL_MODEL_INDUSTRY_SHEET_NAME
INDUSTRY_ECONOMY = workflow_cfg.FULL_MODEL_INDUSTRY_ECONOMY
INDUSTRY_BASE_YEAR = workflow_cfg.FULL_MODEL_INDUSTRY_BASE_YEAR
INDUSTRY_SCENARIO = workflow_cfg.FULL_MODEL_INDUSTRY_SCENARIO
INDUSTRY_REGION = workflow_cfg.FULL_MODEL_INDUSTRY_REGION


def _safe_filename_segment(value: str) -> str:
    return "".join(ch if ch.isalnum() or ch in {"_", "-"} else "_" for ch in str(value).strip())

INDUSTRY_REMAP_FUELS = True
INDUSTRY_MAPPING_CSV_PATH = REPO_ROOT / "intermediate_data" / "industry_fuel_mapping.csv"
INDUSTRY_ESTO_DATA_PATH = REPO_ROOT / "data" / "00APEC_2024_low.csv"
INDUSTRY_NINTH_DATA_PATH = REPO_ROOT / "data" / "merged_file_energy_ALL_20250814_pre_trump.csv"
INDUSTRY_ESTO_SUBTOTAL_MAPPING_PATH = REPO_ROOT / "config" / "ESTO_subtotal_mapping.xlsx"
INDUSTRY_REMAP_OUTPUT_PATH = (
    REPO_ROOT
    / "outputs"
    / "leap_exports"
    / f"industry_export_remapped_{_safe_filename_segment(INDUSTRY_ECONOMY)}_{_safe_filename_segment(INDUSTRY_SCENARIO)}.xlsx"
)
INDUSTRY_REMAP_REPORT_PATH = REPO_ROOT / "intermediate_data" / "industry_fuel_remap_report.csv"
INDUSTRY_REMAP_VALIDATION_PATH = REPO_ROOT / "intermediate_data" / "industry_fuel_remap_validation.csv"
INDUSTRY_NINTH_SCENARIO = "reference"

INDUSTRY_CREATE_BRANCHES = False
INDUSTRY_FILL_BRANCHES = True
INDUSTRY_SET_UNITS = True
INDUSTRY_HANDLE_CURRENT_ACCOUNTS_TOO = True

#%%
# --- Workflow runners ---

def run_transformation_workflow():
    from codebase import transformation_entry

    return transformation_entry.run_transformation_workflow(
        economies=TRANSFORMATION_ECONOMIES,
        scenarios=TRANSFORMATION_SCENARIOS,
        include_leap_import=TRANSFORMATION_INCLUDE_LEAP_IMPORT,
        import_scenario=TRANSFORMATION_IMPORT_SCENARIOS,
        feedstock_method=FEEDSTOCK_METHOD,
        aggregate_economy_label=TRANSFORMATION_AGGREGATE_ECONOMY_LABEL,
        export_output_dir=TRANSFORMATION_EXPORT_DIR,
        filename_template=TRANSFORMATION_FILENAME_TEMPLATE,
    )


def run_hydrogen_transformation_workflow():
    from codebase import hydrogen_transformation_workflow

    return hydrogen_transformation_workflow.run_hydrogen_export_and_import(
        economies=HYDROGEN_TRANSFORMATION_ECONOMIES,
        scenarios=HYDROGEN_TRANSFORMATION_SCENARIOS,
        include_leap_import=HYDROGEN_TRANSFORMATION_INCLUDE_LEAP_IMPORT,
        import_scenario=HYDROGEN_TRANSFORMATION_IMPORT_SCENARIOS,
        handle_current_accounts=HYDROGEN_TRANSFORMATION_HANDLE_CURRENT_ACCOUNTS,
        feedstock_method=FEEDSTOCK_METHOD,
        aggregate_economy_label=HYDROGEN_TRANSFORMATION_AGGREGATE_ECONOMY_LABEL,
        export_output_dir=HYDROGEN_TRANSFORMATION_EXPORT_DIR,
        filename_template=HYDROGEN_TRANSFORMATION_FILENAME_TEMPLATE,
    )


def run_supply_workflow():
    from codebase import supply_workflow

    return supply_workflow.run_supply_export_and_import(
        economies=SUPPLY_ECONOMIES,
        export_dataset_key=SUPPLY_EXPORT_DATASET_KEY,
        scenario_names=SUPPLY_SCENARIOS,
        include_leap_import=SUPPLY_INCLUDE_LEAP_IMPORT,
        import_scenario=SUPPLY_IMPORT_SCENARIOS,
    )


def run_transfers_workflow():
    from codebase import transfers_workflow

    return transfers_workflow.run_transfer_export_and_import(
        economies=TRANSFERS_ECONOMIES,
        scenarios=TRANSFERS_SCENARIOS,
        include_leap_import=TRANSFERS_INCLUDE_LEAP_IMPORT,
        import_scenario=TRANSFERS_IMPORT_SCENARIOS,
        handle_current_accounts=TRANSFERS_HANDLE_CURRENT_ACCOUNTS,
        include_output_series=TRANSFERS_INCLUDE_OUTPUT_SERIES,
        use_output_targets=TRANSFERS_USE_OUTPUT_TARGETS,
        feedstock_method=FEEDSTOCK_METHOD,
        aggregate_economy_label=TRANSFERS_AGGREGATE_ECONOMY_LABEL,
    )


def run_minor_demand_workflow():
    from codebase import minor_demand_workflow

    run_economies = list(MINOR_DEMAND_ECONOMIES)
    should_aggregate, aggregate_label, _ = workflow_common.resolve_aggregate_economy(
        run_economies,
        aggregate_label=MINOR_DEMAND_AGGREGATE_ECONOMY_LABEL,
    )
    if should_aggregate:
        run_economies = [aggregate_label]

    outputs = []
    for economy in run_economies:
        output = minor_demand_workflow.assemble_minor_demand_workbook(
            economy=economy,
            export_filename=MINOR_DEMAND_EXPORT_FILENAME,
            include_leap_import=MINOR_DEMAND_INCLUDE_LEAP_IMPORT,
            scenarios=MINOR_DEMAND_SCENARIOS,
            import_scenario=MINOR_DEMAND_IMPORT_SCENARIOS,
            region=MINOR_DEMAND_REGION,
            aggregate_economy_label=MINOR_DEMAND_AGGREGATE_ECONOMY_LABEL,
        )
        outputs.append(output)
    return outputs


def run_industry_mapping_workflow():
    from codebase.configuration.config import (
        BRANCH_DEMAND_CATEGORY,
        BRANCH_DEMAND_TECHNOLOGY,
    )
    from codebase.functions.industry_fuel_remap import remap_industry_export_fuels
    from codebase.functions.leap_core import (
        connect_to_leap,
        create_branches_from_export_file,
        fill_branches_from_export_file,
    )

    export_path = INDUSTRY_EXPORT_PATH
    industry_sheet_name = INDUSTRY_SHEET_NAME
    if INDUSTRY_REMAP_FUELS:
        remap_industry_export_fuels(
            input_path=str(INDUSTRY_EXPORT_PATH),
            output_path=str(INDUSTRY_REMAP_OUTPUT_PATH),
            mapping_csv_path=str(INDUSTRY_MAPPING_CSV_PATH),
            esto_data_path=str(INDUSTRY_ESTO_DATA_PATH),
            ninth_data_path=str(INDUSTRY_NINTH_DATA_PATH),
            subtotal_mapping_path=str(INDUSTRY_ESTO_SUBTOTAL_MAPPING_PATH),
            economy=INDUSTRY_ECONOMY,
            base_year=INDUSTRY_BASE_YEAR,
            scenario=INDUSTRY_NINTH_SCENARIO,
            sheet_name=INDUSTRY_SHEET_NAME,
            include_extra_others=False,
            report_path=str(INDUSTRY_REMAP_REPORT_PATH),
            validation_path=str(INDUSTRY_REMAP_VALIDATION_PATH),
            ensure_base_year_from_current_accounts=True,
            enforce_base_year_presence=True,
        )
        export_path = INDUSTRY_REMAP_OUTPUT_PATH
        industry_sheet_name = "LEAP"

    if not INDUSTRY_CREATE_BRANCHES and not INDUSTRY_FILL_BRANCHES:
        return export_path
    if not LEAP_API_AVAILABLE:
        print("[WARN] LEAP API unavailable; skipping industry branch creation/fill.")
        return export_path

    L = connect_to_leap()

    if INDUSTRY_CREATE_BRANCHES:
        create_branches_from_export_file(
            L,
            str(export_path),
            sheet_name=industry_sheet_name,
            branch_path_col="Branch Path",
            scenario=INDUSTRY_SCENARIO,
            region=INDUSTRY_REGION,
            branch_type_mapping=None,
            default_branch_type=(
                BRANCH_DEMAND_CATEGORY,
                BRANCH_DEMAND_CATEGORY,
                BRANCH_DEMAND_TECHNOLOGY,
            ),
            RAISE_ERROR_ON_FAILED_BRANCH_CREATION=True,
        )

    if INDUSTRY_FILL_BRANCHES:
        fill_branches_from_export_file(
            L,
            str(export_path),
            sheet_name=industry_sheet_name,
            scenario=INDUSTRY_SCENARIO,
            region=INDUSTRY_REGION,
            RAISE_ERROR_ON_FAILED_SET=True,
            SET_UNITS=INDUSTRY_SET_UNITS,
            HANDLE_CURRENT_ACCOUNTS_TOO=INDUSTRY_HANDLE_CURRENT_ACCOUNTS_TOO,
        )

    return export_path

#%%
# --- Run all workflows ---

def run_all_workflows():
    results = {}
    start = time.time()

    if RUN_TRANSFORMATION_WORKFLOW:
        print("[1/6] Running transformation workflow...")
        results["transformation"] = run_transformation_workflow()

    if RUN_HYDROGEN_TRANSFORMATION_WORKFLOW:
        print("[2/6] Running hydrogen transformation workflow...")
        results["hydrogen_transformation"] = run_hydrogen_transformation_workflow()

    if RUN_SUPPLY_WORKFLOW:
        print("[3/6] Running supply workflow...")
        results["supply"] = run_supply_workflow()

    if RUN_TRANSFERS_WORKFLOW:
        print("[4/6] Running transfers workflow...")
        results["transfers"] = run_transfers_workflow()

    if RUN_MINOR_DEMAND_WORKFLOW:
        print("[5/6] Running minor demand workflow...")
        results["minor_demand"] = run_minor_demand_workflow()

    if RUN_INDUSTRY_MAPPING_WORKFLOW:
        print("[6/6] Running industry mapping workflow...")
        results["industry_mapping"] = run_industry_mapping_workflow()

    elapsed = time.time() - start
    print(f"Done in {elapsed:.1f}s")
    return results

#%%
# Execute everything in one click.
if __name__ == "__main__":
    run_all_workflows()
#%%