#%%
"""
Provide notebook-friendly entry points for transformation workbook exports.

This module wraps `transformation_workflow` with smaller helper functions that
are convenient to call from notebooks or ad hoc scripts. Use it when you want a
quick transformation export/import without stepping through the full workflow
module configuration.
"""

# Simplified transformation workflow for notebooks: exports the LEAP workbook and optionally runs the LEAP import.
# Most user-editable settings live in `codebase/workflow_config.py`.
from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Iterable, Sequence

# REPO_ROOT = Path(__file__).resolve().parents[1]
# CURRENT_DIR = Path.cwd()
# if CURRENT_DIR != REPO_ROOT:
#     os.chdir(REPO_ROOT)
# if str(CURRENT_DIR) not in sys.path:
#     sys.path.insert(0, str(CURRENT_DIR))

from codebase import transformation_workflow as pipeline
from codebase.transformation_workflow import core

#%%
DEFAULT_SCENARIOS = ["Reference", "Target", "Current Accounts"]
DEFAULT_ECONOMIES = list(core.ECONOMIES_TO_ANALYZE)


def quick_transformation_export(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    export_output_dir: Path | str | None = None,
    filename_template: str | None = None,
    feedstock_method: str | None = None,
    aggregate_economy_label: str | None = None,
) -> list[Path]:
    """Re-run the transformation analytics pipeline and emit the LEAP workbook."""
    return pipeline.assemble_transformation_workbook(
        economies=economies,
        scenarios=scenarios,
        export_output_dir=export_output_dir,
        filename_template=filename_template,
        feedstock_method=feedstock_method,
        aggregate_economy_label=aggregate_economy_label,
    )


def run_transformation_workflow(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    include_leap_import: bool = False,
    import_scenario: str | Sequence[str] | None = None,
    feedstock_method: str | None = None,
    aggregate_economy_label: str | None = None,
    **export_kwargs,
) -> list[Path]:
    """Convenience wrapper that optionally performs the LEAP import after exporting."""
    return pipeline.run_transformation_export_and_import(
        economies=economies,
        scenarios=scenarios,
        include_leap_import=include_leap_import,
        import_scenario=import_scenario,
        feedstock_method=feedstock_method,
        aggregate_economy_label=aggregate_economy_label,
        **export_kwargs,
    )


def run_leap_import(
    export_directory: Path | str | None = None,
    filename: str | None = None,
    scenario_to_run: str | None = None,
    **import_kwargs,
) -> Path:
    """Run only the LEAP import against an existing export workbook."""
    return pipeline.import_transformation_workbook_to_leap(
        export_directory=export_directory,
        filename=filename,
        scenario_to_run=scenario_to_run,
        **import_kwargs,
    )


def run_backcalc_verification(
    economies: Iterable[str] | None = None,
    aggregate_economy_label: str | None = None,
    abs_tolerance: float = 1e-6,
    rel_tolerance: float = 1e-3,
    output_csv: Path | str | None = None,
):
    """Verify LEAP-parameter back-calculation against process totals."""
    return pipeline.verify_transformation_backcalculation(
        economies=economies,
        aggregate_economy_label=aggregate_economy_label,
        abs_tolerance=abs_tolerance,
        rel_tolerance=rel_tolerance,
        output_csv=output_csv,
    )


#%%
# Notebook toggles
NOTEBOOK_ECONOMIES = DEFAULT_ECONOMIES
NOTEBOOK_SCENARIOS = DEFAULT_SCENARIOS
NOTEBOOK_INCLUDE_LEAP_IMPORT = True
NOTEBOOK_IMPORT_SCENARIOS = [
    scenario.lower()
    for scenario in NOTEBOOK_SCENARIOS
    if scenario.lower() not in {"current accounts", "current account"}
]
NOTEBOOK_EXPORT_DIR = None
NOTEBOOK_FILENAME_TEMPLATE = None
NOTEBOOK_AGGREGATE_ECONOMY_LABEL = "ALL_ECONOMIES"
NOTEBOOK_FEEDSTOCK_METHOD = None


def run_with_notebook_config() -> list[Path]:
    """Run the workflow using the notebook-friendly constants."""
    return run_transformation_workflow(
        economies=NOTEBOOK_ECONOMIES,
        scenarios=NOTEBOOK_SCENARIOS,
        include_leap_import=NOTEBOOK_INCLUDE_LEAP_IMPORT,
        import_scenario=NOTEBOOK_IMPORT_SCENARIOS,
        feedstock_method=NOTEBOOK_FEEDSTOCK_METHOD,
        aggregate_economy_label=NOTEBOOK_AGGREGATE_ECONOMY_LABEL,
        export_output_dir=NOTEBOOK_EXPORT_DIR,
        filename_template=NOTEBOOK_FILENAME_TEMPLATE,
    )


#%%
