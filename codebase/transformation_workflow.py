#%%
"""
Build transformation export workbooks and optionally fill LEAP branches.

This workflow runs the configured transformation analyses, assembles process
records into LEAP import workbooks, and can push those workbooks into LEAP. It
covers the main transformation processes and should be the default entry point
for non-hydrogen transformation exports.
"""

# Transformation export pipeline helpers that build workbooks and optionally fill LEAP branches.
# Most user-editable settings live in `codebase/workflow_config.py`.
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

from codebase.functions import transformation_analysis_utils as core
from codebase.configuration import workflow_config as workflow_cfg
from codebase.functions import leap_api, leap_exports
from codebase.functions.analysis_input_write_dispatcher import (
    get_analysis_input_write_mode,
)
from codebase.configuration.config import (
    BRANCH_DEMAND_CATEGORY,
    BRANCH_DEMAND_TECHNOLOGY,
)
from codebase.utilities import workflow_common

LEAP_API_AVAILABLE = leap_api.is_available()

#%%
SHEET_NAME = workflow_cfg.TRANSFORMATION_WORKFLOW_SHEET_NAME
EXPORT_FILENAME_PREFIX = workflow_cfg.TRANSFORMATION_WORKFLOW_EXPORT_FILENAME_PREFIX
DEFAULT_SCENARIOS = list(workflow_cfg.TRANSFORMATION_WORKFLOW_DEFAULT_SCENARIOS)
# Projection allocation behavior is configured in
# `codebase/transformation_analysis_utils.py`:
# - `PROJECTION_SIGN_STABLE_MODE`: "all" | "selected" | "off"
# - `PROJECTION_STRICT_CONSERVATION`: raises if allocation no longer conserves totals.
# This workflow consumes `core.esto_data`, so those settings directly affect
# transformation exports (and transfers via shared core data).

RUN_LNG_ANALYSIS = core.RUN_LNG_ANALYSIS
RUN_GAS_PROCESSING_ANALYSIS = core.RUN_GAS_PROCESSING_ANALYSIS
RUN_COAL_TRANSFORMATION_ANALYSIS = core.RUN_COAL_TRANSFORMATION_ANALYSIS
RUN_OTHER_TRANSFORMATION_ANALYSIS = core.RUN_OTHER_TRANSFORMATION_ANALYSIS
RUN_CHARCOAL_PROCESSING_ANALYSIS = core.RUN_CHARCOAL_PROCESSING_ANALYSIS
RUN_NONSPECIFIED_TRANSFORMATION_ANALYSIS = core.RUN_NONSPECIFIED_TRANSFORMATION_ANALYSIS
RUN_HYDROGEN_TRANSFORMATION_ANALYSIS = core.RUN_HYDROGEN_TRANSFORMATION_ANALYSIS

ANALYSIS_REGISTRY = [
    ("lng", core.run_lng_analysis, RUN_LNG_ANALYSIS),
    ("gas_works", core.run_gas_processing_analysis, RUN_GAS_PROCESSING_ANALYSIS),
    ("coal_coke_ovens", core.run_flow_sector_analysis, RUN_COAL_TRANSFORMATION_ANALYSIS),
    ("coal_blast_furnaces", core.run_flow_sector_analysis, RUN_COAL_TRANSFORMATION_ANALYSIS),
    ("coal_patent_fuel_plants", core.run_flow_sector_analysis, RUN_COAL_TRANSFORMATION_ANALYSIS),
    ("coal_bkb_pb_plants", core.run_flow_sector_analysis, RUN_COAL_TRANSFORMATION_ANALYSIS),
    ("coal_liquefaction", core.run_flow_sector_analysis, RUN_COAL_TRANSFORMATION_ANALYSIS),
    # ("coal_mines", core.run_flow_sector_analysis, RUN_COAL_TRANSFORMATION_ANALYSIS),
    ("electric_boilers", core.run_flow_sector_analysis, RUN_OTHER_TRANSFORMATION_ANALYSIS),
    ("chemical_heat_for_electricity_production", core.run_flow_sector_analysis, RUN_OTHER_TRANSFORMATION_ANALYSIS),
    ("petrochemical_industry", core.run_flow_sector_analysis, RUN_OTHER_TRANSFORMATION_ANALYSIS),
    ("gas_to_liquids_plants", core.run_flow_sector_analysis, RUN_OTHER_TRANSFORMATION_ANALYSIS),
    ("biofuels_processing", core.run_flow_sector_analysis, RUN_OTHER_TRANSFORMATION_ANALYSIS),
    ("charcoal_processing", core.run_flow_sector_analysis, RUN_CHARCOAL_PROCESSING_ANALYSIS),
    ("nonspecified_transformation", core.run_flow_sector_analysis, RUN_NONSPECIFIED_TRANSFORMATION_ANALYSIS),
    ("hydrogen_transformation", core.run_hydrogen_transformation_analysis, RUN_HYDROGEN_TRANSFORMATION_ANALYSIS),
]


def _print_reset_reminder_for_import(include_leap_import: bool) -> None:
    """Remind users that standalone transformation import does not clear stale trade targets."""
    if not include_leap_import:
        return
    print(
        "[WARN] Reset reminder: standalone transformation workflow import does not perform a global "
        "supply/transformation trade reset. If you need a clean rerun, run "
        "codebase/results_supply_link_workflow.py with "
        "MAIN_RUN_RESET_SUPPLY_AND_TRANSFORMATION_IMPORT_EXPORT=True."
    )


def format_export_filename(
    economy_label: str,
    scenarios: Sequence[str],
    template: str | None = None,
) -> str:
    """Format the workbook name consumed by `save_transformation_export`."""
    template = template or core.EXPORT_FILENAME_TEMPLATE
    return leap_exports.build_workbook_filename(
        economy_label=economy_label,
        scenarios=scenarios,
        template=template,
        fallback_template=core.EXPORT_FILENAME_FALLBACK,
    )


def _infer_primary_economy(rows: Sequence[dict]) -> str:
    """Return the first economy that appears in the generated rows."""
    for row in rows:
        economy = row.get("economy")
        if economy:
            return economy
    if core.ECONOMIES_TO_ANALYZE:
        return core.ECONOMIES_TO_ANALYZE[0]
    return "economy"


def build_transformation_rows(economies: Iterable[str] | None = None) -> list[dict]:
    """Run the configured analyses and return the transformation rows."""
    original_economies = core.ECONOMIES_TO_ANALYZE
    override = economies is not None
    if override:
        core.ECONOMIES_TO_ANALYZE = list(economies)
    rows: list[dict] = []
    try:
        for sector_key, callback, enabled in ANALYSIS_REGISTRY:
            core.run_analysis_for_sector(enabled, sector_key, callback, rows)
    finally:
        if override:
            core.ECONOMIES_TO_ANALYZE = original_economies
    return rows


def collect_transformation_rows(
    economies: Iterable[str] | None = None,
    aggregate_economy_label: str | None = None,
    projection_scenario: str | None = None,
) -> list[dict]:
    """Collect process rows with optional synthetic all-economies aggregation."""
    economy_list = workflow_common.normalize_economies(economies or core.ECONOMIES_TO_ANALYZE)
    should_aggregate, aggregate_label, _ = workflow_common.resolve_aggregate_economy(
        economy_list,
        aggregate_label=aggregate_economy_label or core.ALL_ECONOMY_LABEL,
    )
    run_economies = list(economy_list)
    data_map_override = None
    import_export_data_override = None
    import_export_year_cols_override = None
    normalized_projection_scenario = str(projection_scenario or "").strip().lower()
    if normalized_projection_scenario:
        scenario_ninth = core.clean_esto_subtotals(core.ninth_data_raw, core.ninth_year_cols)
        if "scenarios" in scenario_ninth.columns:
            scenario_filtered = scenario_ninth[
                scenario_ninth["scenarios"].astype(str).str.strip().str.lower() == normalized_projection_scenario
            ].copy()
            if scenario_filtered.empty:
                raise ValueError(
                    f"No 9th rows found for projection_scenario={projection_scenario!r}"
                )
            scenario_ninth = scenario_filtered
        if "subtotal_results" in scenario_ninth.columns:
            scenario_ninth = scenario_ninth[scenario_ninth["subtotal_results"] == False].copy()
        scenario_ninth = core.filter_total_energy_rows(scenario_ninth)

        scenario_esto = core.normalize_esto_economy_codes(core.esto_data_raw.copy())
        scenario_esto = core.filter_total_energy_rows(scenario_esto)
        scenario_esto_with_subtotals = core.apply_matt_subtotal_mapping(
            scenario_esto,
            core.SUBTOTAL_MAPPING_PATH,
        )
        scenario_esto = core.filter_matt_subtotals(scenario_esto_with_subtotals)
        scenario_esto_year_cols = sorted([col for col in scenario_esto.columns if str(col).isdigit()])

        if should_aggregate:
            scenario_ninth = core.add_all_economy_total(
                scenario_ninth,
                core.ninth_year_cols,
                aggregate_label,
            )
            scenario_esto = core.add_all_economy_total(
                scenario_esto,
                scenario_esto_year_cols,
                aggregate_label,
            )
            run_economies = [aggregate_label]

        projection_sign_stable_flows = core.resolve_projection_sign_stable_flows(
            core.PROJECTION_SIGN_STABLE_MODE,
            core.SIGN_STABLE_PROJECTION_FLOWS,
        )
        try:
            projection_df, _ = core.build_esto_projection_table(
                ninth_data=scenario_ninth,
                esto_data=scenario_esto,
                mapping_path=core.NINTH_TO_ESTO_MAPPING_PATH,
                base_year=core.BASE_YEAR,
                projection_years=core.PROJECTION_YEAR_RANGE,
                scenario=normalized_projection_scenario,
                sign_stable_flows=projection_sign_stable_flows,
                strict_conservation=core.PROJECTION_STRICT_CONSERVATION,
            )
        except ValueError as exc:
            print(
                "[WARN] Projection strict-conservation check failed for "
                f"projection_scenario={projection_scenario!r}; retrying non-strict: {exc}"
            )
            projection_df, _ = core.build_esto_projection_table(
                ninth_data=scenario_ninth,
                esto_data=scenario_esto,
                mapping_path=core.NINTH_TO_ESTO_MAPPING_PATH,
                base_year=core.BASE_YEAR,
                projection_years=core.PROJECTION_YEAR_RANGE,
                scenario=normalized_projection_scenario,
                sign_stable_flows=projection_sign_stable_flows,
                strict_conservation=False,
            )
        scenario_esto = core.merge_projection_into_esto(
            scenario_esto,
            projection_df,
            core.PROJECTION_YEAR_RANGE,
        )
        scenario_esto_year_cols = sorted([col for col in scenario_esto.columns if str(col).isdigit()])
        data_map_override = core.build_dataset_map(
            scenario_esto,
            scenario_esto_year_cols,
            scenario_ninth,
            core.ninth_year_cols,
            core.esto_data_raw,
            core.esto_year_cols_raw,
        )
        import_export_data_override = scenario_esto
        import_export_year_cols_override = scenario_esto_year_cols
    elif should_aggregate:
        aggregated_ninth = core.add_all_economy_total(
            core.ninth_data,
            core.ninth_year_cols,
            aggregate_label,
        )
        aggregated_esto = core.add_all_economy_total(
            core.esto_data,
            core.esto_year_cols,
            aggregate_label,
        )
        data_map_override = core.build_dataset_map(
            aggregated_esto,
            core.esto_year_cols,
            aggregated_ninth,
            core.ninth_year_cols,
            core.esto_data_raw,
            core.esto_year_cols_raw,
        )
        import_export_data_override = aggregated_esto
        import_export_year_cols_override = core.esto_year_cols
        run_economies = [aggregate_label]
    previous_data_map = core.DATASET_MAP
    previous_import_export_data = core.ESTO_IMPORT_EXPORT_REFERENCE_DATA
    previous_import_export_years = core.ESTO_IMPORT_EXPORT_YEAR_COLS
    if data_map_override is not None:
        core.DATASET_MAP = data_map_override
        core.ESTO_IMPORT_EXPORT_REFERENCE_DATA = import_export_data_override
        core.ESTO_IMPORT_EXPORT_YEAR_COLS = import_export_year_cols_override
    try:
        rows = build_transformation_rows(economies=run_economies)
    finally:
        if data_map_override is not None:
            core.DATASET_MAP = previous_data_map
            core.ESTO_IMPORT_EXPORT_REFERENCE_DATA = previous_import_export_data
            core.ESTO_IMPORT_EXPORT_YEAR_COLS = previous_import_export_years
    return rows


def _sum_fuel_values_by_year(value_map: dict | None) -> dict[int, float]:
    """Return year totals from a fuel->(year->value) mapping."""
    totals: dict[int, float] = {}
    if not isinstance(value_map, dict):
        return totals
    for value in value_map.values():
        if not isinstance(value, dict):
            continue
        for year, raw in value.items():
            if raw is None or pd.isna(raw):
                continue
            year_int = int(year)
            totals[year_int] = totals.get(year_int, 0.0) + float(raw)
    return totals


def _coerce_year_value(value, year: int) -> float:
    """Return a numeric value for `year` from scalar/dict inputs."""
    if isinstance(value, dict):
        raw = value.get(year, value.get(str(year), 0.0))
        if raw is None or pd.isna(raw):
            return 0.0
        return float(raw)
    if value is None or pd.isna(value):
        return 0.0
    return float(value)


def _coerce_year_map(value_map: dict | None) -> dict[int, float]:
    """Return a normalized year->value mapping."""
    result: dict[int, float] = {}
    if not isinstance(value_map, dict):
        return result
    for year, value in value_map.items():
        if value is None or pd.isna(value):
            continue
        result[int(year)] = float(value)
    return result


def verify_transformation_backcalculation(
    economies: Iterable[str] | None = None,
    aggregate_economy_label: str | None = None,
    abs_tolerance: float = 1e-6,
    rel_tolerance: float = 1e-3,
    output_csv: Path | str | None = None,
) -> tuple[pd.DataFrame, dict]:
    """Back-calculate outputs from LEAP parameters and compare against source process totals.

    For each process-year:
    - source output is summed from `output_values` in process records.
    - output used by LEAP is `max(source_output, 0)` (negative years are non-exportable).
    - implied output is `efficiency * (feedstock_total + loss_total)`.
    """
    rows = collect_transformation_rows(
        economies=economies,
        aggregate_economy_label=aggregate_economy_label,
    )
    checks: list[dict] = []
    for record in rows:
        output_by_year = _sum_fuel_values_by_year(record.get("output_values"))
        feedstock_by_year = _sum_fuel_values_by_year(record.get("feedstock_values"))
        loss_by_year = _sum_fuel_values_by_year(record.get("loss_values"))
        loss_for_eff_by_year = _coerce_year_map(record.get("loss_values_for_efficiency"))
        year_set = set(output_by_year) | set(feedstock_by_year) | set(loss_by_year)
        year_set.update(loss_for_eff_by_year)
        efficiency = record.get("efficiency")
        if isinstance(efficiency, dict):
            year_set.update(int(year) for year in efficiency.keys())
        for ratio in (record.get("auxiliary_ratios") or {}).values():
            if isinstance(ratio, dict):
                year_set.update(int(year) for year in ratio.keys())
        for share in (record.get("feedstock_shares") or {}).values():
            if isinstance(share, dict):
                year_set.update(int(year) for year in share.keys())
        for year in sorted(year_set):
            output_raw = output_by_year.get(year, 0.0)
            output_for_leap = max(output_raw, 0.0)
            feedstock_total = max(feedstock_by_year.get(year, 0.0), 0.0)
            loss_total_raw = abs(loss_by_year.get(year, 0.0))
            loss_total_for_eff = abs(
                loss_for_eff_by_year.get(year, loss_total_raw)
            )
            efficiency_value = max(_coerce_year_value(efficiency, year), 0.0)
            implied_output = efficiency_value * (feedstock_total + loss_total_for_eff)
            abs_error = abs(implied_output - output_for_leap)
            rel_error = (
                abs_error / output_for_leap
                if output_for_leap > 0
                else (0.0 if abs_error <= abs_tolerance else float("inf"))
            )
            passes = abs_error <= abs_tolerance or (
                output_for_leap > 0 and rel_error <= rel_tolerance
            )
            aux_implied_total = 0.0
            for ratio in (record.get("auxiliary_ratios") or {}).values():
                ratio_value = max(_coerce_year_value(ratio, year), 0.0)
                aux_implied_total += ratio_value * output_for_leap
            share_sum = 0.0
            for share in (record.get("feedstock_shares") or {}).values():
                share_sum += max(_coerce_year_value(share, year), 0.0)
            checks.append(
                {
                    "economy": record.get("economy"),
                    "sector_title": record.get("sector_title"),
                    "process_name": record.get("process_name"),
                    "year": year,
                    "output_raw": output_raw,
                    "output_for_leap": output_for_leap,
                    "feedstock_total": feedstock_total,
                    "loss_total": loss_total_raw,
                    "loss_total_for_efficiency": loss_total_for_eff,
                    "efficiency": efficiency_value,
                    "implied_output_from_leap_params": implied_output,
                    "abs_error": abs_error,
                    "rel_error": rel_error,
                    "passes_tolerance": passes,
                    "negative_raw_output": output_raw < 0,
                    "aux_input_implied_total": aux_implied_total,
                    "feedstock_share_sum": share_sum,
                }
            )
    checks_df = pd.DataFrame(checks)
    if output_csv is not None:
        output_path = Path(output_csv)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        checks_df.to_csv(output_path, index=False)
    if checks_df.empty:
        summary = {
            "row_count": 0,
            "failed_rows": 0,
            "negative_raw_output_rows": 0,
            "max_abs_error": 0.0,
            "max_rel_error": 0.0,
        }
        return checks_df, summary
    failed_rows = int((~checks_df["passes_tolerance"]).sum())
    negative_rows = int(checks_df["negative_raw_output"].sum())
    summary = {
        "row_count": int(len(checks_df)),
        "failed_rows": failed_rows,
        "negative_raw_output_rows": negative_rows,
        "max_abs_error": float(checks_df["abs_error"].max()),
        "max_rel_error": float(checks_df["rel_error"].replace([float("inf")], pd.NA).max(skipna=True) or 0.0),
    }
    return checks_df, summary


def assemble_transformation_workbook(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    export_output_dir: Path | str | None = None,
    filename_template: str | None = None,
    feedstock_method: str | None = None,
    aggregate_economy_label: str | None = None,
    build_export: bool = core.BUILD_LEAP_EXPORT,
) -> list[Path]:
    """Build transformation rows, persist the LEAP workbook, and return the export path(s)."""
    if not build_export:
        print("BUILD_LEAP_EXPORT is False; skipping workbook generation.")
        return []
    original_feedstock_method = core.FEEDSTOCK_METHOD
    if feedstock_method is not None:
        core.FEEDSTOCK_METHOD = core.resolve_feedstock_method(feedstock_method)
    try:
        rows = collect_transformation_rows(
            economies=economies,
            aggregate_economy_label=aggregate_economy_label,
        )
        if not rows:
            print("No transformation rows were generated; nothing to export.")
            return []
        if core.SAVE_SUMMARY_TABLES:
            core.save_transformation_summaries(
                rows,
                core.code_to_name_mapping,
                core.SUMMARY_OUTPUT_DIR,
                core.PROCESS_SUMMARY_FILENAME,
                core.DETAIL_SUMMARY_FILENAME,
            )
        core.consolidate_transformation_output_rows(
            rows,
            include_output_series=core.INCLUDE_OUTPUT_SERIES_IN_LEAP_EXPORT,
            use_output_targets=bool(
                core.TRANSFORMATION_OUTPUT_VARIABLES.get("output_import_target")
                or core.TRANSFORMATION_OUTPUT_VARIABLES.get("output_export_target")
            ),
        )
        scenario_list = workflow_common.normalize_workflow_scenarios(
            scenarios,
            DEFAULT_SCENARIOS,
        )
        economy_label = _infer_primary_economy(rows)
        output_dir_path = Path(export_output_dir or core.EXPORT_OUTPUT_DIR)
        output_dir_path.mkdir(parents=True, exist_ok=True)
        export_filename = format_export_filename(economy_label, scenario_list, filename_template)
        export_path = core.save_transformation_export(
            rows,
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
    finally:
        core.FEEDSTOCK_METHOD = original_feedstock_method


def run_transformation_export_and_import(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    include_leap_import: bool = False,
    import_scenario: str | Sequence[str] | None = None,
    region: str | None = None,
    handle_current_accounts: bool = True,
    create_branches: bool = True,
    fill_branches: bool = True,
    aggregate_economy_label: str | None = None,
    feedstock_method: str | None = None,
    **export_kwargs,
) -> list[Path]:
    """Run exports and optionally push the workbook into LEAP."""
    _print_reset_reminder_for_import(include_leap_import)
    exports = assemble_transformation_workbook(
        economies=economies,
        scenarios=scenarios,
        export_output_dir=export_kwargs.get("export_output_dir"),
        filename_template=export_kwargs.get("filename_template"),
        feedstock_method=feedstock_method,
        aggregate_economy_label=aggregate_economy_label,
        build_export=export_kwargs.get("build_export", core.BUILD_LEAP_EXPORT),
    )
    if not exports or not include_leap_import:
        return exports
    scenario_list = workflow_common.normalize_workflow_scenarios(
        scenarios,
        DEFAULT_SCENARIOS,
    )
    scenario_choices = workflow_common.resolve_import_scenarios(
        scenario_list,
        import_scenario,
    )
    if get_analysis_input_write_mode() == "api" and not LEAP_API_AVAILABLE:
        print(
            "[INFO] LEAP API unavailable in this environment; skipping branch creation/fill."
        )
        return exports
    for index, scenario_choice in enumerate(scenario_choices):
        import_transformation_workbook_to_leap(
            export_directory=exports[0].parent,
            filename=exports[0].name,
            scenario_to_run=scenario_choice,
            region=region or core.EXPORT_REGION,
            include_current_accounts=handle_current_accounts and index == 0,
            create_branches=create_branches and index == 0,
            fill_branches=fill_branches,
        )
    return exports


def list_export_scenarios(export_path: Path) -> list[str]:
    """Return the Scenario column values in declaration order."""
    return leap_exports.list_scenarios(export_path, sheet_name=SHEET_NAME)


def validate_export_region(export_path: Path, region: str) -> None:
    """Ensure the workbook contains the requested region."""
    return leap_exports.validate_region(export_path, region, sheet_name=SHEET_NAME)


def find_transformation_workbook(
    directory: Path | str | None = None, filename: str | None = None
) -> Path:
    """Return a candidate transformation workbook, optionally using an explicit name."""
    directory_path = Path(directory or core.EXPORT_OUTPUT_DIR)
    return leap_exports.find_workbook(
        directory=directory_path,
        prefix=EXPORT_FILENAME_PREFIX,
        filename=filename,
    )


def import_transformation_workbook_to_leap(
    export_directory: Path | str | None = None,
    filename: str | None = None,
    scenario_to_run: str | None = None,
    region: str | None = None,
    include_current_accounts: bool = False,
    create_branches: bool = True,
    fill_branches: bool = True,
    raise_on_missing_branch: bool = False,
) -> Path:
    """Connect to LEAP, create the branches, and fill the data from the export file."""
    if (
        str(scenario_to_run or "").strip().lower() in {"current accounts", "current account"}
        and not include_current_accounts
    ):
        raise ValueError(
            "Direct transformation LEAP import for 'Current Accounts' is disabled "
            "unless include_current_accounts=True is passed explicitly."
        )
    export_path = find_transformation_workbook(export_directory, filename)
    target_region = region or core.EXPORT_REGION
    return leap_api.import_workbook(
        export_path=export_path,
        sheet_name=SHEET_NAME,
        scenario=scenario_to_run,
        region=target_region,
        create_branches=create_branches,
        fill_branches=fill_branches,
        include_current_accounts=include_current_accounts,
        default_branch_type=(
            BRANCH_DEMAND_CATEGORY,
            BRANCH_DEMAND_CATEGORY,
            BRANCH_DEMAND_TECHNOLOGY,
        ),
        raise_on_missing_branch=raise_on_missing_branch,
    )


# Legacy names kept for compatibility.
def _collect_process_records(economies: Iterable[str] | None = None) -> list[dict]:
    return build_transformation_rows(economies)


def prepare_transformation_exports(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    export_output_dir: Path | str | None = None,
    filename_template: str | None = None,
    aggregate_economy_label: str | None = None,
    build_export: bool = core.BUILD_LEAP_EXPORT,
) -> list[Path]:
    return assemble_transformation_workbook(
        economies=economies,
        scenarios=scenarios,
        export_output_dir=export_output_dir,
        filename_template=filename_template,
        aggregate_economy_label=aggregate_economy_label,
        build_export=build_export,
    )


def run_transformation_pipeline(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    include_leap_import: bool = False,
    import_scenario: str | Sequence[str] | None = None,
    region: str | None = None,
    handle_current_accounts: bool = True,
    create_branches: bool = True,
    fill_branches: bool = True,
    aggregate_economy_label: str | None = None,
    **export_kwargs,
) -> list[Path]:
    return run_transformation_export_and_import(
        economies=economies,
        scenarios=scenarios,
        include_leap_import=include_leap_import,
        import_scenario=import_scenario,
        region=region,
        handle_current_accounts=handle_current_accounts,
        create_branches=create_branches,
        fill_branches=fill_branches,
        aggregate_economy_label=aggregate_economy_label,
        **export_kwargs,
    )


def locate_transformation_export(
    directory: Path | str | None = None, filename: str | None = None
) -> Path:
    return find_transformation_workbook(directory=directory, filename=filename)


def get_available_scenarios(export_path: Path) -> list[str]:
    return list_export_scenarios(export_path)


def ensure_region_in_export(export_path: Path, region: str) -> None:
    return validate_export_region(export_path, region)


def run_transformation_leap_import(
    export_directory: Path | str | None = None,
    filename: str | None = None,
    scenario_to_run: str | None = None,
    region: str | None = None,
    include_current_accounts: bool = False,
    create_branches: bool = True,
    fill_branches: bool = True,
    raise_on_missing_branch: bool = False,
) -> Path:
    return import_transformation_workbook_to_leap(
        export_directory=export_directory,
        filename=filename,
        scenario_to_run=scenario_to_run,
        region=region,
        include_current_accounts=include_current_accounts,
        create_branches=create_branches,
        fill_branches=fill_branches,
        raise_on_missing_branch=raise_on_missing_branch,
    )


#%%

# Simple notebook-focused configuration block.
NOTEBOOK_SCENARIOS = ["Reference", "Target", "Current Accounts"]
NOTEBOOK_INCLUDE_LEAP_IMPORT = (
    LEAP_API_AVAILABLE if get_analysis_input_write_mode() == "api" else True
)
NOTEBOOK_IMPORT_SCENARIOS = [
    scenario.lower()
    for scenario in NOTEBOOK_SCENARIOS
    if scenario.lower() not in {"current accounts", "current account"}
]
NOTEBOOK_ECONOMIES = list(core.ECONOMIES_TO_ANALYZE)
NOTEBOOK_CURRENT_ACCOUNTS = True
NOTEBOOK_AGGREGATE_ECONOMY_LABEL = "ALL_ECONOMIES"

def run_with_notebook_config() -> list[Path]:
    """Run the transformation export/import helpers with the editable notebook constants."""
    return run_transformation_export_and_import(
        economies=NOTEBOOK_ECONOMIES,
        scenarios=NOTEBOOK_SCENARIOS,
        include_leap_import=NOTEBOOK_INCLUDE_LEAP_IMPORT,
        import_scenario=NOTEBOOK_IMPORT_SCENARIOS,
        handle_current_accounts=NOTEBOOK_CURRENT_ACCOUNTS,
        aggregate_economy_label=NOTEBOOK_AGGREGATE_ECONOMY_LABEL,
    )

if __name__ == "__main__":
    run_with_notebook_config()
#%%


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
