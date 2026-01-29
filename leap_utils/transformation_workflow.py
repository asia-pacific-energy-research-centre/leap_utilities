#%%
# Transformation export pipeline helpers that build workbooks and optionally fill LEAP branches.
from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Iterable, Sequence

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
CURRENT_DIR = Path.cwd()
if CURRENT_DIR != REPO_ROOT:
    os.chdir(REPO_ROOT)
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from leap_utils import transformation_analysis_utils as core
from leap_utils.config import (
    BRANCH_DEMAND_CATEGORY,
    BRANCH_DEMAND_TECHNOLOGY,
)
from leap_utils.leap_core import (
    connect_to_leap,
    create_branches_from_export_file,
    fill_branches_from_export_file,
    is_leap_api_available,
)

LEAP_API_AVAILABLE = is_leap_api_available()

#%%
SHEET_NAME = "LEAP"
EXPORT_FILENAME_PREFIX = "transformation_leap_imports_"

RUN_LNG_ANALYSIS = core.RUN_LNG_ANALYSIS
RUN_GAS_PROCESSING_ANALYSIS = core.RUN_GAS_PROCESSING_ANALYSIS
RUN_COAL_TRANSFORMATION_ANALYSIS = core.RUN_COAL_TRANSFORMATION_ANALYSIS
RUN_CHARCOAL_PROCESSING_ANALYSIS = core.RUN_CHARCOAL_PROCESSING_ANALYSIS
RUN_NONSPECIFIED_TRANSFORMATION_ANALYSIS = core.RUN_NONSPECIFIED_TRANSFORMATION_ANALYSIS

ANALYSIS_REGISTRY = [
    ("lng", core.run_lng_analysis, RUN_LNG_ANALYSIS),
    ("gas_works", core.run_gas_processing_analysis, RUN_GAS_PROCESSING_ANALYSIS),
    ("coal_coke_ovens", core.run_flow_sector_analysis, RUN_COAL_TRANSFORMATION_ANALYSIS),
    ("coal_blast_furnaces", core.run_flow_sector_analysis, RUN_COAL_TRANSFORMATION_ANALYSIS),
    ("coal_patent_fuel_plants", core.run_flow_sector_analysis, RUN_COAL_TRANSFORMATION_ANALYSIS),
    ("coal_bkb_pb_plants", core.run_flow_sector_analysis, RUN_COAL_TRANSFORMATION_ANALYSIS),
    ("coal_liquefaction", core.run_flow_sector_analysis, RUN_COAL_TRANSFORMATION_ANALYSIS),
    ("coal_mines", core.run_flow_sector_analysis, RUN_COAL_TRANSFORMATION_ANALYSIS),
    ("charcoal_processing", core.run_flow_sector_analysis, RUN_CHARCOAL_PROCESSING_ANALYSIS),
    ("nonspecified_transformation", core.run_flow_sector_analysis, RUN_NONSPECIFIED_TRANSFORMATION_ANALYSIS),
]

def _format_scenario_segment(scenarios: Sequence[str]) -> str:
    """Sanitize scenario labels that become part of the filename."""
    tokens = [core.format_filename_segment(segment) for segment in scenarios if segment]
    sanitized = "_".join(token for token in tokens if token)
    return sanitized or "scenarios"


def _build_export_filename(
    economy_label: str,
    scenarios: Sequence[str],
    template: str | None = None,
) -> str:
    """Format the workbook name consumed by `save_transformation_export`."""
    template = template or core.EXPORT_FILENAME_TEMPLATE
    scenario_segment = _format_scenario_segment(scenarios)
    economy_segment = core.format_filename_segment(economy_label)
    try:
        return template.format(economy=economy_segment, scenario=scenario_segment)
    except Exception as exc:
        print(f"Failed to format export filename: {exc}")
        return core.EXPORT_FILENAME_FALLBACK


def _infer_primary_economy(process_records: Sequence[dict]) -> str:
    """Return the first economy that appears in the generated process records."""
    for record in process_records:
        economy = record.get("economy")
        if economy:
            return economy
    if core.ECONOMIES_TO_ANALYZE:
        return core.ECONOMIES_TO_ANALYZE[0]
    return "economy"


def _collect_process_records(economies: Iterable[str] | None = None) -> list[dict]:
    """Run the configured analyses and gather every transformation process record."""
    original_economies = core.ECONOMIES_TO_ANALYZE
    override = economies is not None
    if override:
        core.ECONOMIES_TO_ANALYZE = list(economies)
    records: list[dict] = []
    try:
        for sector_key, callback, enabled in ANALYSIS_REGISTRY:
            core.run_analysis_for_sector(enabled, sector_key, callback, records)
    finally:
        if override:
            core.ECONOMIES_TO_ANALYZE = original_economies
    return records


def prepare_transformation_exports(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    export_output_dir: Path | str | None = None,
    filename_template: str | None = None,
    build_export: bool = core.BUILD_LEAP_EXPORT,
) -> list[Path]:
    """Run the analytics, persist the LEAP workbook, and return the export path(s)."""
    if not build_export:
        print("BUILD_LEAP_EXPORT is False; skipping workbook generation.")
        return []
    process_records = _collect_process_records(economies=economies)
    if not process_records:
        print("No transformation records were generated; nothing to export.")
        return []
    if core.SAVE_SUMMARY_TABLES:
        core.save_transformation_summaries(
            process_records,
            core.code_to_name_mapping,
            core.SUMMARY_OUTPUT_DIR,
            core.PROCESS_SUMMARY_FILENAME,
            core.DETAIL_SUMMARY_FILENAME,
        )
    scenario_list = list(scenarios or core.SCENARIOS_TO_EXPORT)
    economy_label = _infer_primary_economy(process_records)
    output_dir_path = Path(export_output_dir or core.EXPORT_OUTPUT_DIR)
    output_dir_path.mkdir(parents=True, exist_ok=True)
    export_filename = _build_export_filename(economy_label, scenario_list, filename_template)
    export_path = core.save_transformation_export(
        process_records,
        core.EXPORT_REGION,
        core.EXPORT_BASE_YEAR,
        core.EXPORT_FINAL_YEAR,
        core.code_to_name_mapping,
        str(output_dir_path),
        export_filename,
        core.EXPORT_MODEL_NAME,
        scenario_list,
    )
    return [Path(export_path)] if export_path else []


def run_transformation_pipeline(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    include_leap_import: bool = False,
    import_scenario: str | None = None,
    region: str | None = None,
    handle_current_accounts: bool = True,
    create_branches: bool = True,
    fill_branches: bool = True,
    **export_kwargs,
) -> list[Path]:
    """Run exports and optionally push the workbook into LEAP."""
    exports = prepare_transformation_exports(
        economies=economies,
        scenarios=scenarios,
        export_output_dir=export_kwargs.get("export_output_dir"),
        filename_template=export_kwargs.get("filename_template"),
        build_export=export_kwargs.get("build_export", core.BUILD_LEAP_EXPORT),
    )
    if not exports or not include_leap_import:
        return exports
    scenario_choice = import_scenario or (scenarios or core.SCENARIOS_TO_EXPORT)[0]
    if not LEAP_API_AVAILABLE:
        print(
            "[INFO] LEAP API unavailable in this environment; skipping branch creation/fill."
        )
        return exports
    run_transformation_leap_import(
        export_directory=exports[0].parent,
        filename=exports[0].name,
        scenario_to_run=scenario_choice,
        region=region or core.EXPORT_REGION,
        include_current_accounts=handle_current_accounts,
        create_branches=create_branches,
        fill_branches=fill_branches,
    )
    return exports


def _read_unique_column(export_path: Path, column: str) -> list[str]:
    """Return unique values from a column while preserving the order they appear in the workbook."""
    for header in (2, 0):
        try:
            df = pd.read_excel(
                export_path, sheet_name=SHEET_NAME, header=header, usecols=[column]
            )
        except Exception:
            continue
        if column not in df.columns:
            continue
        seen: list[str] = []
        for value in df[column].dropna().astype(str):
            if value not in seen:
                seen.append(value)
        if seen:
            return seen
    return []

def get_available_scenarios(export_path: Path) -> list[str]:
    """Return the Scenario column values in declaration order."""
    return _read_unique_column(export_path, "Scenario")


def ensure_region_in_export(export_path: Path, region: str) -> None:
    """Ensure the workbook contains the requested region."""
    regions = _read_unique_column(export_path, "Region")
    if not regions:
        print(f"Warning: 'Region' column missing from {export_path.name}; skipping region check.")
        return
    if region not in regions:
        raise ValueError(
            f"Requested region '{region}' not present in {export_path.name}; available: {regions}"
        )


def locate_transformation_export(
    directory: Path | str | None = None, filename: str | None = None
) -> Path:
    """Return a candidate transformation workbook, optionally using an explicit name."""
    directory_path = Path(directory or core.EXPORT_OUTPUT_DIR)
    if filename:
        candidate = directory_path / filename
        if candidate.exists():
            return candidate
        raise FileNotFoundError(f"Specified transformation export missing: {candidate}")
    matches = sorted(directory_path.glob(f"{EXPORT_FILENAME_PREFIX}*.xlsx"))
    if not matches:
        raise FileNotFoundError(f"No transformation exports detected in {directory_path}")
    return matches[-1]


def run_transformation_leap_import(
    export_directory: Path | str | None = None,
    filename: str | None = None,
    scenario_to_run: str | None = None,
    region: str | None = None,
    include_current_accounts: bool = True,
    create_branches: bool = True,
    fill_branches: bool = True,
    raise_on_missing_branch: bool = False,
) -> Path:
    """Connect to LEAP, create the branches, and fill the data from the export file."""
    export_path = locate_transformation_export(export_directory, filename)
    available = get_available_scenarios(export_path)
    scenario_choice = scenario_to_run or (available[0] if available else None)
    if scenario_choice and scenario_choice not in available:
        raise ValueError(
            f"Scenario '{scenario_choice}' not found in {export_path.name}; options {available}"
        )
    target_region = region or core.EXPORT_REGION
    ensure_region_in_export(export_path, target_region)
    
    leap_conn = connect_to_leap()
    if leap_conn is None:
        raise RuntimeError("Unable to connect to LEAP.")
    if create_branches:
        create_branches_from_export_file(
            leap_conn,
            export_path,
            sheet_name=SHEET_NAME,
            branch_root=None,
            default_branch_type=(
                BRANCH_DEMAND_CATEGORY,
                BRANCH_DEMAND_CATEGORY,
                BRANCH_DEMAND_TECHNOLOGY,
            ),
            RAISE_ERROR_ON_FAILED_BRANCH_CREATION=raise_on_missing_branch,
        )
    if fill_branches:
        
        fill_branches_from_export_file(
            leap_conn,
            export_path,
            sheet_name=SHEET_NAME,
            scenario=scenario_choice,
            region=target_region,
            HANDLE_CURRENT_ACCOUNTS_TOO=include_current_accounts,
        )
    return export_path


#%%

# Simple notebook-focused configuration block.
NOTEBOOK_SCENARIOS = ["Reference", "Target", "Current Accounts"]
NOTEBOOK_INCLUDE_LEAP_IMPORT = LEAP_API_AVAILABLE
NOTEBOOK_IMPORT_SCENARIO = "Target"
NOTEBOOK_ECONOMIES = list(core.ECONOMIES_TO_ANALYZE)
NOTEBOOK_CURRENT_ACCOUNTS = True

def run_with_notebook_config() -> list[Path]:
    """Run the transformation export/import helpers with the editable notebook constants."""
    return run_transformation_pipeline(
        economies=NOTEBOOK_ECONOMIES,
        scenarios=NOTEBOOK_SCENARIOS,
        include_leap_import=NOTEBOOK_INCLUDE_LEAP_IMPORT,
        import_scenario=NOTEBOOK_IMPORT_SCENARIO,
        handle_current_accounts=NOTEBOOK_CURRENT_ACCOUNTS,
    )

run_with_notebook_config()
#%%
#todo check why we arent getting projections for a lot of processes.