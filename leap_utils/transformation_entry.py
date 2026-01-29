#%%
# Simplified transformation workflow for notebooks: exports the LEAP workbook and optionally runs the LEAP import.
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

from leap_utils import transformation_workflow as pipeline
from leap_utils.transformation_workflow import core

#%%
DEFAULT_SCENARIOS = ["Reference", "Target", "Current Accounts"]
DEFAULT_ECONOMIES = list(core.ECONOMIES_TO_ANALYZE)


def quick_transformation_export(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    export_output_dir: Path | str | None = None,
    filename_template: str | None = None,
) -> list[Path]:
    """Re-run the transformation analytics pipeline and emit the LEAP workbook."""
    return pipeline.prepare_transformation_exports(
        economies=economies,
        scenarios=scenarios,
        export_output_dir=export_output_dir,
        filename_template=filename_template,
    )


def run_transformation_workflow(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    include_leap_import: bool = False,
    import_scenario: str | None = None,
    **export_kwargs,
) -> list[Path]:
    """Convenience wrapper that optionally performs the LEAP import after exporting."""
    return pipeline.run_transformation_pipeline(
        economies=economies,
        scenarios=scenarios,
        include_leap_import=include_leap_import,
        import_scenario=import_scenario,
        **export_kwargs,
    )


def run_leap_import(
    export_directory: Path | str | None = None,
    filename: str | None = None,
    scenario_to_run: str | None = None,
    **import_kwargs,
) -> Path:
    """Run only the LEAP import against an existing export workbook."""
    return pipeline.run_transformation_leap_import(
        export_directory=export_directory,
        filename=filename,
        scenario_to_run=scenario_to_run,
        **import_kwargs,
    )


#%%
# Notebook toggles
NOTEBOOK_ECONOMIES = DEFAULT_ECONOMIES
NOTEBOOK_SCENARIOS = DEFAULT_SCENARIOS
NOTEBOOK_INCLUDE_LEAP_IMPORT = True
NOTEBOOK_IMPORT_SCENARIO = "Target"
NOTEBOOK_EXPORT_DIR = None
NOTEBOOK_FILENAME_TEMPLATE = None


def run_with_notebook_config() -> list[Path]:
    """Run the workflow using the notebook-friendly constants."""
    return run_transformation_workflow(
        economies=NOTEBOOK_ECONOMIES,
        scenarios=NOTEBOOK_SCENARIOS,
        include_leap_import=NOTEBOOK_INCLUDE_LEAP_IMPORT,
        import_scenario=NOTEBOOK_IMPORT_SCENARIO,
        export_output_dir=NOTEBOOK_EXPORT_DIR,
        filename_template=NOTEBOOK_FILENAME_TEMPLATE,
    )


#%%
