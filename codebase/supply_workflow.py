#%%
# High-level supply workflow helpers that stay under 300 lines, delegating heavy logic to `supply_data_pipeline.py`.
# Most user-editable settings live in `codebase/workflow_config.py`.
#%%

#%%
from __future__ import annotations

import sys
from pathlib import Path
from typing import Iterable, Sequence

REPO_ROOT = Path(__file__).resolve().parents[1]
try:
    if str(REPO_ROOT) not in sys.path:
        sys.path.insert(0, str(REPO_ROOT))
except Exception as exc:
    print(f"Failed to add repo root to sys.path: {exc}")

from codebase.functions import supply_data_pipeline
from codebase.configuration import workflow_config as workflow_cfg
from codebase.utilities import workflow_common

#%%
DEFAULT_ECONOMIES = list(workflow_cfg.SUPPLY_WORKFLOW_DEFAULT_ECONOMIES)
DEFAULT_SCENARIOS = list(workflow_cfg.SUPPLY_WORKFLOW_DEFAULT_SCENARIOS)


def normalize_economies(economies: Iterable[str] | None = None) -> list[str]:
    """Return a concrete list of economies for the quick entrypoint."""
    if economies:
        return list(economies)
    return list(DEFAULT_ECONOMIES)


def assemble_supply_workbooks(
    economies: Iterable[str] | None = None,
    export_dataset_key: str = "esto",
    scenario_names: Sequence[str] | None = None,
) -> list[supply_data_pipeline.Path]:
    """
    Run the simplified supply export pipeline.

    Args:
        economies: Optional list of economy codes; defaults to `["ALL"]`.
        export_dataset_key: Dataset key (usually `"esto"`).
        scenario_names: Scenario labels to export; defaults to `DEFAULT_SCENARIOS`.

    Returns:
        A list of paths to the generated export workbooks.
    """
    scenarios = scenario_names or list(DEFAULT_SCENARIOS)
    run_economies = normalize_economies(economies)
    should_aggregate, aggregate_label, _ = workflow_common.resolve_aggregate_economy(
        run_economies,
        aggregate_label=workflow_cfg.SUPPLY_ALL_ECONOMY_LABEL,
    )
    if should_aggregate:
        run_economies = [aggregate_label]
    assets = supply_data_pipeline.prepare_supply_assets(economies=run_economies)
    dataset_map, sector_config, code_to_name_mapping, _, _ = assets
    export_paths = supply_data_pipeline.generate_supply_exports(
        dataset_map,
        sector_config,
        code_to_name_mapping,
        projection_years=supply_data_pipeline.PROJECTION_YEAR_RANGE,
        dataset_key=export_dataset_key,
        economies=run_economies,
        scenario_names=scenarios,
        export_output_dir=supply_data_pipeline.EXPORT_OUTPUT_DIR,
        filename_template=supply_data_pipeline.EXPORT_FILENAME_TEMPLATE,
    )
    return [path for _, path in export_paths]


def run_supply_export_and_import(
    economies: Iterable[str] | None = None,
    export_dataset_key: str = "esto",
    scenario_names: Sequence[str] | None = None,
    include_leap_import: bool = False,
    import_scenario: str | Sequence[str] | None = None,
) -> list[supply_data_pipeline.Path]:
    """Run the export preparation and optionally fill LEAP using the generated workbooks."""
    scenarios = workflow_common.normalize_workflow_scenarios(
        scenario_names,
        DEFAULT_SCENARIOS,
    )
    exports = assemble_supply_workbooks(
        economies=economies,
        export_dataset_key=export_dataset_key,
        scenario_names=scenarios,
    )
    if include_leap_import:
        scenario_choices = workflow_common.resolve_import_scenarios(
            scenarios,
            import_scenario,
        )
        for export_path in exports:
            for index, scenario_to_run in enumerate(scenario_choices):
                supply_data_pipeline.run_supply_leap_import(
                    export_directory=supply_data_pipeline.EXPORT_DIR,
                    filename=export_path.name,
                    scenario_to_run=scenario_to_run,
                    handle_current_accounts=index == 0,
                    fill_branches=True,
                )
    return exports


# Legacy names kept for compatibility.
def quick_supply_export(
    economies: Iterable[str] | None = None,
    export_dataset_key: str = "esto",
    scenario_names: Sequence[str] | None = None,
) -> list[supply_data_pipeline.Path]:
    return assemble_supply_workbooks(
        economies=economies,
        export_dataset_key=export_dataset_key,
        scenario_names=scenario_names,
    )


def run_supply_pipeline(
    economies: Iterable[str] | None = None,
    export_dataset_key: str = "esto",
    scenario_names: Sequence[str] | None = None,
    include_leap_import: bool = False,
    import_scenario: str | Sequence[str] | None = None,
) -> list[supply_data_pipeline.Path]:
    return run_supply_export_and_import(
        economies=economies,
        export_dataset_key=export_dataset_key,
        scenario_names=scenario_names,
        include_leap_import=include_leap_import,
        import_scenario=import_scenario,
    )

#%%
#----------------------------------------------------------------------------
# Simple configuration block for notebook/interactive usage.
#----------------------------------------------------------------------------
NOTEBOOK_WORKFLOW_ECONOMIES = list(workflow_cfg.SUPPLY_NOTEBOOK_ECONOMIES)
NOTEBOOK_INCLUDE_LEAP_IMPORT = workflow_cfg.SUPPLY_NOTEBOOK_INCLUDE_LEAP_IMPORT
NOTEBOOK_SCENARIOS = list(workflow_cfg.SUPPLY_NOTEBOOK_SCENARIOS)
NOTEBOOK_IMPORT_SCENARIOS = [
    scenario.lower()
    for scenario in NOTEBOOK_SCENARIOS
    if scenario.lower() not in {"current accounts", "current account"}
]


def run_with_config() -> list[supply_data_pipeline.Path]:
    """Run the supply workflow using the editable constants in this file."""
    return run_supply_export_and_import(
        economies=NOTEBOOK_WORKFLOW_ECONOMIES,
        include_leap_import=NOTEBOOK_INCLUDE_LEAP_IMPORT,
        scenario_names=NOTEBOOK_SCENARIOS,
        import_scenario=NOTEBOOK_IMPORT_SCENARIOS,
    )

#%%
if __name__ == "__main__":
    run_with_config()
#%%
