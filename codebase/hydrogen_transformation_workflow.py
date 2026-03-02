#%%
# Hydrogen-specific transformation workflow that reuses transformation core helpers.
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
from codebase.configuration.config import (
    BRANCH_DEMAND_CATEGORY,
    BRANCH_DEMAND_TECHNOLOGY,
)
from codebase.utilities import workflow_common

LEAP_API_AVAILABLE = leap_api.is_available()

SHEET_NAME = workflow_cfg.HYDROGEN_SHEET_NAME
EXPORT_FILENAME_PREFIX = workflow_cfg.HYDROGEN_EXPORT_FILENAME_PREFIX
EXPORT_FILENAME_TEMPLATE = workflow_cfg.HYDROGEN_EXPORT_FILENAME_TEMPLATE
EXPORT_FILENAME_FALLBACK = workflow_cfg.HYDROGEN_EXPORT_FILENAME_FALLBACK
PROCESS_SUMMARY_FILENAME = workflow_cfg.HYDROGEN_PROCESS_SUMMARY_FILENAME
DETAIL_SUMMARY_FILENAME = workflow_cfg.HYDROGEN_DETAIL_SUMMARY_FILENAME
DEFAULT_SCENARIOS = list(workflow_cfg.HYDROGEN_DEFAULT_SCENARIOS)

RUN_HYDROGEN_TRANSFORMATION_ANALYSIS = core.RUN_HYDROGEN_TRANSFORMATION_ANALYSIS
ANALYSIS_REGISTRY = [
    (
        "hydrogen_transformation",
        core.run_hydrogen_transformation_analysis,
        RUN_HYDROGEN_TRANSFORMATION_ANALYSIS,
    ),
]


HYDROGEN_DISPLAY_NAME_OVERRIDES = {
    "16_x_hydrogen": "Hydrogen",
    "16_x_ammonia": "Ammonia",
    "16_x_efuel": "Efuel",
    "electrolysers_non_green": "Electrolysers (non-green electricity)",
}


def _build_hydrogen_display_mapping() -> dict:
    """Return code->name mapping with hydrogen output display overrides."""
    mapping = dict(core.code_to_name_mapping or {})
    mapping.update(HYDROGEN_DISPLAY_NAME_OVERRIDES)
    return mapping


def _infer_primary_economy(rows: Sequence[dict]) -> str:
    for row in rows:
        economy = row.get("economy")
        if economy:
            return economy
    if core.ECONOMIES_TO_ANALYZE:
        return core.ECONOMIES_TO_ANALYZE[0]
    return "economy"


def _coerce_economy_label(value, fallback: str = "ALL_ECONOMIES") -> str:
    """Return a stable string economy label."""
    if value is None:
        return fallback
    if isinstance(value, bool):
        return fallback
    text = str(value).strip()
    return text or fallback


def deduplicate_hydrogen_output_targets(rows: list[dict]) -> None:
    """Keep one output import/export target per output fuel at sector level."""
    if not rows:
        return
    grouped: dict[tuple[str, str], list[dict]] = {}
    for record in rows:
        grouped.setdefault(
            (str(record.get("economy")), str(record.get("sector_title"))),
            [],
        ).append(record)
    for records in grouped.values():
        seen_import_labels: set[str] = set()
        seen_export_labels: set[str] = set()
        for record in records:
            import_targets = record.get("output_import_targets") or {}
            export_targets = record.get("output_export_targets") or {}
            filtered_import_targets = {}
            for label, values in import_targets.items():
                label_key = str(label)
                if label_key in seen_import_labels:
                    continue
                filtered_import_targets[label] = values
                seen_import_labels.add(label_key)
            filtered_export_targets = {}
            for label, values in export_targets.items():
                label_key = str(label)
                if label_key in seen_export_labels:
                    continue
                filtered_export_targets[label] = values
                seen_export_labels.add(label_key)
            record["output_import_targets"] = filtered_import_targets
            record["output_export_targets"] = filtered_export_targets


def format_export_filename(
    economy_label: str,
    scenarios: Sequence[str],
    template: str | None = None,
) -> str:
    template = template or EXPORT_FILENAME_TEMPLATE
    return leap_exports.build_workbook_filename(
        economy_label=economy_label,
        scenarios=scenarios,
        template=template,
        fallback_template=EXPORT_FILENAME_FALLBACK,
    )


def _get_hydrogen_sector_config() -> dict:
    sector_config = core.MAJOR_SECTOR_CONFIG.get("hydrogen_transformation")
    if not sector_config:
        raise KeyError("hydrogen_transformation config not found in MAJOR_SECTOR_CONFIG")
    return sector_config


def build_hydrogen_rows(
    economies: Iterable[str] | None = None,
    process_config: list[dict] | None = None,
) -> list[dict]:
    """Run configured hydrogen analyses and return process records."""
    original_economies = core.ECONOMIES_TO_ANALYZE
    override_economies = economies is not None
    sector_config = _get_hydrogen_sector_config()
    original_process_config = sector_config.get("process_config")
    override_process_config = process_config is not None
    if override_economies:
        core.ECONOMIES_TO_ANALYZE = [
            _coerce_economy_label(value, fallback="")
            for value in economies
        ]
    if override_process_config:
        sector_config["process_config"] = process_config
    rows: list[dict] = []
    try:
        for sector_key, callback, enabled in ANALYSIS_REGISTRY:
            core.run_analysis_for_sector(enabled, sector_key, callback, rows)
    finally:
        if override_economies:
            core.ECONOMIES_TO_ANALYZE = original_economies
        if override_process_config:
            sector_config["process_config"] = original_process_config
    return rows


def collect_hydrogen_rows(
    economies: Iterable[str] | None = None,
    process_config: list[dict] | None = None,
    aggregate_economy_label: str | None = None,
) -> list[dict]:
    """Collect hydrogen process rows with optional all-economies aggregation."""
    economy_list = workflow_common.normalize_economies(economies or core.ECONOMIES_TO_ANALYZE)
    aggregate_label = _coerce_economy_label(
        aggregate_economy_label,
        fallback=core.ALL_ECONOMY_LABEL,
    )
    should_aggregate, aggregate_label, _ = workflow_common.resolve_aggregate_economy(
        economy_list,
        aggregate_label=aggregate_label,
    )
    run_economies = list(economy_list)
    data_map_override = None
    import_export_data_override = None
    import_export_year_cols_override = None
    if should_aggregate:
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
        rows = build_hydrogen_rows(
            economies=run_economies,
            process_config=process_config,
        )
    finally:
        if data_map_override is not None:
            core.DATASET_MAP = previous_data_map
            core.ESTO_IMPORT_EXPORT_REFERENCE_DATA = previous_import_export_data
            core.ESTO_IMPORT_EXPORT_YEAR_COLS = previous_import_export_years
    return rows


def _sum_record_output_values(rows: Sequence[dict]) -> dict[tuple[str, str, int], float]:
    """Return summed output values keyed by (economy, subfuel, year)."""
    totals: dict[tuple[str, str, int], float] = {}
    for record in rows:
        economy = str(record.get("economy"))
        output_values = record.get("output_values") or {}
        for subfuel, value_by_year in output_values.items():
            if not isinstance(value_by_year, dict):
                continue
            for year, value in value_by_year.items():
                if value is None or pd.isna(value):
                    continue
                key = (economy, str(subfuel), int(year))
                totals[key] = totals.get(key, 0.0) + float(value)
    return totals


def _build_hydrogen_source_output_totals(
    process_config: list[dict] | None = None,
    aggregate_economy_label: str | None = None,
    economies: Iterable[str] | None = None,
) -> dict[tuple[str, str, int], float]:
    """Return source 9th output totals keyed by (economy, subfuel, year)."""
    sector_config = _get_hydrogen_sector_config()
    active_config = process_config or sector_config.get("process_config") or []
    source_sub2 = []
    output_subfuels = []
    for cfg in active_config:
        if not isinstance(cfg, dict):
            continue
        if not cfg.get("enabled", True):
            continue
        sub2 = cfg.get("source_sub2sectors") or cfg.get("sub2sectors") or cfg.get("process_code")
        if sub2 and sub2 not in source_sub2:
            source_sub2.append(str(sub2))
        for output in cfg.get("output_subfuels") or []:
            if output and output not in output_subfuels:
                output_subfuels.append(str(output))
    if not output_subfuels:
        output_subfuels = list(core.HYDROGEN_OUTPUT_SUBFUELS)
    output_fuels = set(sector_config.get("output_fuels") or ["16_others"])
    transformation_sub1 = sector_config.get(
        "transformation_sub1",
        "09_13_hydrogen_transformation",
    )

    source_data = core.ninth_data
    economy_list = workflow_common.normalize_economies(economies)
    aggregate_label = _coerce_economy_label(
        aggregate_economy_label,
        fallback=core.ALL_ECONOMY_LABEL,
    )
    should_aggregate, aggregate_label, _ = workflow_common.resolve_aggregate_economy(
        economy_list,
        aggregate_label=aggregate_label,
    )
    if should_aggregate:
        source_data = core.add_all_economy_total(
            source_data,
            core.ninth_year_cols,
            aggregate_label,
        )
    source_rows = core.select_rows(
        source_data,
        {"sub1sectors": transformation_sub1},
    )
    if source_rows.empty:
        return {}
    if economy_list and "economy" in source_rows.columns:
        economy_set = {aggregate_label} if should_aggregate else {str(econ) for econ in economy_list}
        source_rows = source_rows[source_rows["economy"].astype(str).isin(economy_set)]
    if source_sub2 and "sub2sectors" in source_rows.columns:
        source_rows = source_rows[source_rows["sub2sectors"].astype(str).isin(source_sub2)]
    if output_fuels and "fuels" in source_rows.columns:
        source_rows = source_rows[source_rows["fuels"].astype(str).isin(output_fuels)]
    if output_subfuels and "subfuels" in source_rows.columns:
        source_rows = source_rows[source_rows["subfuels"].astype(str).isin(output_subfuels)]
    if source_rows.empty:
        return {}

    year_cols = [
        year for year in core.ninth_year_cols
        if core.EXPORT_BASE_YEAR <= year <= core.EXPORT_FINAL_YEAR
    ]
    grouped = (
        source_rows.groupby(["economy", "subfuels"], dropna=False)[year_cols]
        .sum()
        .reset_index()
    )
    totals: dict[tuple[str, str, int], float] = {}
    for _, row in grouped.iterrows():
        economy = str(row["economy"])
        subfuel = str(row["subfuels"])
        for year in year_cols:
            key = (economy, subfuel, int(year))
            totals[key] = float(row[year])
    return totals


def verify_hydrogen_output_reproduction(
    economies: Iterable[str] | None = None,
    process_config: list[dict] | None = None,
    aggregate_economy_label: str | None = None,
    abs_tolerance: float = 1e-6,
    rel_tolerance: float = 1e-6,
    output_csv: Path | str | None = None,
) -> tuple[pd.DataFrame, dict]:
    """Compare modeled hydrogen outputs to 9th source outputs by economy/year/subfuel."""
    rows = collect_hydrogen_rows(
        economies=economies,
        process_config=process_config,
        aggregate_economy_label=aggregate_economy_label,
    )
    modeled_totals = _sum_record_output_values(rows)
    compare_economies = sorted({key[0] for key in modeled_totals.keys()})
    if economies is not None:
        compare_economies = [str(economy) for economy in economies]
    source_totals = _build_hydrogen_source_output_totals(
        process_config=process_config,
        aggregate_economy_label=aggregate_economy_label,
        economies=compare_economies if compare_economies else None,
    )
    display_mapping = _build_hydrogen_display_mapping()
    all_keys = sorted(set(modeled_totals) | set(source_totals))
    checks = []
    for economy, subfuel, year in all_keys:
        modeled_value = float(modeled_totals.get((economy, subfuel, year), 0.0))
        source_value = float(source_totals.get((economy, subfuel, year), 0.0))
        abs_error = abs(modeled_value - source_value)
        rel_error = (
            abs_error / abs(source_value)
            if source_value != 0.0
            else (0.0 if abs_error <= abs_tolerance else float("inf"))
        )
        passes = abs_error <= abs_tolerance or (
            source_value != 0.0 and rel_error <= rel_tolerance
        )
        checks.append(
            {
                "economy": economy,
                "subfuel": subfuel,
                "subfuel_display": display_mapping.get(subfuel, subfuel),
                "year": int(year),
                "modeled_output": modeled_value,
                "source_output": source_value,
                "abs_error": abs_error,
                "rel_error": rel_error,
                "passes_tolerance": passes,
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
            "max_abs_error": 0.0,
            "max_rel_error": 0.0,
        }
        return checks_df, summary
    failed_rows = int((~checks_df["passes_tolerance"]).sum())
    summary = {
        "row_count": int(len(checks_df)),
        "failed_rows": failed_rows,
        "max_abs_error": float(checks_df["abs_error"].max()),
        "max_rel_error": float(
            checks_df["rel_error"].replace([float("inf")], pd.NA).max(skipna=True) or 0.0
        ),
    }
    return checks_df, summary


def assemble_hydrogen_workbook(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    export_output_dir: Path | str | None = None,
    filename_template: str | None = None,
    process_config: list[dict] | None = None,
    feedstock_method: str | None = None,
    aggregate_economy_label: str | None = None,
    build_export: bool = core.BUILD_LEAP_EXPORT,
) -> list[Path]:
    """Build hydrogen process rows, persist LEAP workbook, and return export path(s)."""
    if not build_export:
        print("BUILD_LEAP_EXPORT is False; skipping workbook generation.")
        return []
    original_feedstock_method = core.FEEDSTOCK_METHOD
    if feedstock_method is not None:
        core.FEEDSTOCK_METHOD = core.resolve_feedstock_method(feedstock_method)
    try:
        rows = collect_hydrogen_rows(
            economies=economies,
            process_config=process_config,
            aggregate_economy_label=aggregate_economy_label,
        )
        if not rows:
            print("No hydrogen transformation rows were generated; nothing to export.")
            return []
        display_mapping = _build_hydrogen_display_mapping()
        if core.SAVE_SUMMARY_TABLES:
            core.save_transformation_summaries(
                rows,
                display_mapping,
                core.SUMMARY_OUTPUT_DIR,
                PROCESS_SUMMARY_FILENAME,
                DETAIL_SUMMARY_FILENAME,
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
            display_mapping,
            str(output_dir_path),
            export_filename,
            core.EXPORT_MODEL_NAME,
            scenario_list,
        )
        return [Path(export_path)] if export_path else []
    finally:
        core.FEEDSTOCK_METHOD = original_feedstock_method


def run_hydrogen_export_and_import(
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
    """Run hydrogen exports and optionally push workbook into LEAP."""
    exports = assemble_hydrogen_workbook(
        economies=economies,
        scenarios=scenarios,
        export_output_dir=export_kwargs.get("export_output_dir"),
        filename_template=export_kwargs.get("filename_template"),
        process_config=export_kwargs.get("process_config"),
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
    if not LEAP_API_AVAILABLE:
        print("[INFO] LEAP API unavailable in this environment; skipping branch creation/fill.")
        return exports
    for index, scenario_choice in enumerate(scenario_choices):
        import_hydrogen_workbook_to_leap(
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
    return leap_exports.list_scenarios(export_path, sheet_name=SHEET_NAME)


def validate_export_region(export_path: Path, region: str) -> None:
    return leap_exports.validate_region(export_path, region, sheet_name=SHEET_NAME)


def find_hydrogen_workbook(
    directory: Path | str | None = None, filename: str | None = None
) -> Path:
    directory_path = Path(directory or core.EXPORT_OUTPUT_DIR)
    return leap_exports.find_workbook(
        directory=directory_path,
        prefix=EXPORT_FILENAME_PREFIX,
        filename=filename,
    )


def import_hydrogen_workbook_to_leap(
    export_directory: Path | str | None = None,
    filename: str | None = None,
    scenario_to_run: str | None = None,
    region: str | None = None,
    include_current_accounts: bool = True,
    create_branches: bool = True,
    fill_branches: bool = True,
    raise_on_missing_branch: bool = False,
) -> Path:
    """Connect to LEAP, create branches, and fill data from hydrogen export."""
    export_path = find_hydrogen_workbook(export_directory, filename)
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
def prepare_hydrogen_exports(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    export_output_dir: Path | str | None = None,
    filename_template: str | None = None,
    process_config: list[dict] | None = None,
    feedstock_method: str | None = None,
    aggregate_economy_label: str | None = None,
    build_export: bool = core.BUILD_LEAP_EXPORT,
) -> list[Path]:
    return assemble_hydrogen_workbook(
        economies=economies,
        scenarios=scenarios,
        export_output_dir=export_output_dir,
        filename_template=filename_template,
        process_config=process_config,
        feedstock_method=feedstock_method,
        aggregate_economy_label=aggregate_economy_label,
        build_export=build_export,
    )


def run_hydrogen_pipeline(
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
    return run_hydrogen_export_and_import(
        economies=economies,
        scenarios=scenarios,
        include_leap_import=include_leap_import,
        import_scenario=import_scenario,
        region=region,
        handle_current_accounts=handle_current_accounts,
        create_branches=create_branches,
        fill_branches=fill_branches,
        aggregate_economy_label=aggregate_economy_label,
        feedstock_method=feedstock_method,
        **export_kwargs,
    )


def locate_hydrogen_export(
    directory: Path | str | None = None, filename: str | None = None
) -> Path:
    return find_hydrogen_workbook(directory=directory, filename=filename)


def get_available_scenarios(export_path: Path) -> list[str]:
    return list_export_scenarios(export_path)


def ensure_region_in_export(export_path: Path, region: str) -> None:
    return validate_export_region(export_path, region)


def run_hydrogen_leap_import(
    export_directory: Path | str | None = None,
    filename: str | None = None,
    scenario_to_run: str | None = None,
    region: str | None = None,
    include_current_accounts: bool = True,
    create_branches: bool = True,
    fill_branches: bool = True,
    raise_on_missing_branch: bool = False,
) -> Path:
    return import_hydrogen_workbook_to_leap(
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
ECONOMIES = list(core.ECONOMIES_TO_ANALYZE)
SCENARIOS = list(DEFAULT_SCENARIOS)
INCLUDE_LEAP_IMPORT = LEAP_API_AVAILABLE
IMPORT_SCENARIOS = [
    scenario.lower()
    for scenario in SCENARIOS
    if scenario.lower() not in {"current accounts", "current account"}
]
CURRENT_ACCOUNTS = True
AGGREGATE_ECONOMY_LABEL = "ALL_ECONOMIES"

#%%
if __name__ == "__main__":
    exports = run_hydrogen_export_and_import(
        economies=ECONOMIES,
        scenarios=SCENARIOS,
        include_leap_import=INCLUDE_LEAP_IMPORT,
        import_scenario=IMPORT_SCENARIOS,
        handle_current_accounts=CURRENT_ACCOUNTS,
        aggregate_economy_label=AGGREGATE_ECONOMY_LABEL,
    )
    if exports:
        print(f"Hydrogen transformation export saved to: {exports[0]}")
#%%
