#%%
# High-level supply workflow helpers that stay under 300 lines, delegating heavy logic to `supply_data_pipeline.py`.
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

from leap_utils import supply_data_pipeline

#%%
DEFAULT_ECONOMIES = ["ALL"]
DEFAULT_SCENARIOS = ["Current Accounts", "Reference", "Target"]


def normalize_economies(economies: Iterable[str] | None = None) -> list[str]:
    """Return a concrete list of economies for the quick entrypoint."""
    if economies:
        return list(economies)
    return list(DEFAULT_ECONOMIES)


def quick_supply_export(
    economies: Iterable[str] | None = None,
    include_all_economies: bool = True,
    export_dataset_key: str = "esto",
    scenario_names: Sequence[str] | None = None,
) -> list[supply_data_pipeline.Path]:
    """
    Run the simplified supply export pipeline.

    Args:
        economies: Optional list of economy codes; defaults to `["ALL"]`.
        include_all_economies: Whether to add synthetic `ALL` rows.
        export_dataset_key: Dataset key (usually `"esto"`).
        scenario_names: Scenario labels to export; defaults to `DEFAULT_SCENARIOS`.

    Returns:
        A list of paths to the generated export workbooks.
    """
    scenarios = scenario_names or list(DEFAULT_SCENARIOS)
    run_economies = normalize_economies(economies)
    assets = supply_data_pipeline.prepare_supply_assets(
        include_all_economies=include_all_economies
    )
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


def run_supply_pipeline(
    economies: Iterable[str] | None = None,
    include_all_economies: bool = True,
    export_dataset_key: str = "esto",
    scenario_names: Sequence[str] | None = None,
    include_leap_import: bool = False,
    import_scenario: str | None = None,
) -> list[supply_data_pipeline.Path]:
    """Run the export preparation and optionally fill LEAP using the generated workbooks."""
    scenarios = scenario_names or list(DEFAULT_SCENARIOS)
    exports = quick_supply_export(
        economies=economies,
        include_all_economies=include_all_economies,
        export_dataset_key=export_dataset_key,
        scenario_names=scenarios,
    )
    if include_leap_import:
        scenario_to_run = import_scenario or (scenarios[0] if scenarios else supply_data_pipeline.SCENARIO_TO_RUN)
        for export_path in exports:
            supply_data_pipeline.run_supply_leap_import(
                export_directory=supply_data_pipeline.EXPORT_DIR,
                filename=export_path.name,
                scenario_to_run=scenario_to_run,
                fill_branches=True,
            )
    return exports

#%%
#----------------------------------------------------------------------------
# Simple configuration block for notebook/interactive usage.
#----------------------------------------------------------------------------
NOTEBOOK_WORKFLOW_ECONOMIES = ["ALL"]
NOTEBOOK_INCLUDE_LEAP_IMPORT = True
NOTEBOOK_IMPORT_SCENARIO = "Target"
NOTEBOOK_SCENARIOS = ["Current Accounts", "Reference", "Target"]


def run_with_config() -> list[supply_data_pipeline.Path]:
    """Run the supply workflow using the editable constants in this file."""
    return run_supply_pipeline(
        economies=NOTEBOOK_WORKFLOW_ECONOMIES,
        include_leap_import=NOTEBOOK_INCLUDE_LEAP_IMPORT,
        scenario_names=NOTEBOOK_SCENARIOS,
        import_scenario=NOTEBOOK_IMPORT_SCENARIO,
    )

#%%
run_with_config() 
#%%
