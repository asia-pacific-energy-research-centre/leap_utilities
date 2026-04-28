#%%
"""
Utility functions for pulling LEAP Results tables via the COM API (v2.3).

These helpers keep the flow Jupyter-friendly: connect to LEAP, set context
(area/scenario/region/year/unit/branch/variable), optionally activate a
favorite to lock the view layout, then export the Results table to CSV for
post-processing in pandas.

Notes:
- COM access works only on Windows with pywin32. WSL will fail at import.
- ExportResultsCSV writes directly to a file path; ExportResultsXLS in the
  type library does not accept a filename, so we standardise on CSV then
  convert to Excel in Python.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Iterable, Optional

try:
    import win32com.client  # type: ignore
except ImportError:
    win32com = None  # sentinel for environments without COM


# Stable constants
REPO_ROOT = Path(__file__).resolve().parents[2]
from codebase.functions.leap_api_guard import ensure_leap_api_allowed
from codebase.functions.leap_session import (
    get_live_pinned_leap_app,
    pin_leap_app,
)


def ensure_repo_on_path() -> None:
    """Ensure repository root is importable when running from notebooks."""
    if str(REPO_ROOT) not in sys.path:
        sys.path.insert(0, str(REPO_ROOT))


def connect_leap(*, visible: bool = False, reuse_running: bool = True):
    """Return a LEAPApplication COM object, reusing a running instance when possible."""
    ensure_leap_api_allowed("leap_results_functions.connect_leap")
    if win32com is None:
        raise ImportError("pywin32 is required for LEAP COM access (Windows only)")
    pinned = get_live_pinned_leap_app()
    if pinned is not None:
        if visible:
            pinned.Visible = True
        return pinned
    app = None
    if reuse_running:
        try:
            app = win32com.client.GetActiveObject("Leap.LEAPApplication")
        except Exception:
            app = None
    if app is None:
        app = win32com.client.Dispatch("Leap.LEAPApplication")
    pin_leap_app(app)
    # Only change visibility when the caller explicitly asks; avoid hiding an existing UI.
    if visible:
        app.Visible = True
    return app


def select_area(app, area_name: Optional[str]) -> None:
    """Set ActiveArea if a name is provided."""
    if area_name:
        app.ActiveArea = area_name


def ensure_calculated(app, *, force: bool = False) -> None:
    """Calculate results when needed or explicitly forced."""
    if force:
        app.ForceCalculation()
        app.Calculate()
        return

    try:
        scenario = app.ActiveScenario
        needs_calc = getattr(scenario, "NeedsCalculation", False)
    except Exception:
        needs_calc = False

    if needs_calc:
        app.Calculate()


def activate_favorite(app, favorite_name: Optional[str]) -> Optional[str]:
    """Activate a saved Results favorite by name (Folder#Favorite)."""
    if not favorite_name:
        return None
    try:
        fave = app.Favorites.Item(favorite_name)
        fave.Activate()
        return str(fave.Name)
    except Exception as exc:  # noqa: BLE001
        return f"Favorite not found/activated: {favorite_name} ({exc})"


def set_axes(app, *, x_axis: Optional[str] = None, legend: Optional[str] = None) -> None:
    """Assign Results X-axis and legend dimensions when provided."""
    if x_axis:
        app.ResultsXAxis = x_axis
    if legend:
        app.ResultsLegend = legend


def set_context(
    app,
    *,
    scenario: Optional[str] = None,
    region: Optional[str] = None,
    year: Optional[int] = None,
    unit: Optional[str] = None,
    branch_path: Optional[str] = None,
    variable_name: Optional[str] = None,
) -> None:
    """Set the active scenario/region/year/unit/branch/variable if given."""
    if scenario:
        app.ActiveScenario = scenario
    if region:
        app.ActiveRegion = region
    if year is not None:
        app.ActiveYear = int(year)
    if unit:
        app.ActiveUnit = unit
    if branch_path:
        app.ActiveBranch = branch_path
    if variable_name:
        app.ActiveVariable = variable_name


def list_dimensions(app) -> list[str]:
    """Return the names of dimensions currently available in Results view."""
    try:
        dims = app.Dimensions
        return [str(d.Name) for d in dims]
    except Exception:
        return []


def list_scenarios(app) -> list[str]:
    try:
        return [str(s.Name) for s in app.Scenarios]
    except Exception:
        return []


def list_regions(app) -> list[str]:
    try:
        return [str(r.Name) for r in app.Regions]
    except Exception:
        return []


def export_results_csv(app, output_path: Path) -> Path:
    """Export the current Results table to CSV at output_path."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    app.ShowResultsViewTable()
    app.ExportResultsCSV(str(output_path))
    return output_path


def fetch_branchvar_value(
    app,
    *,
    branch_var: str,
    year: int,
    unit: str = "",
) -> float:
    """Retrieve a single result via ResultValue(branchVarName, year, unit)."""
    return float(app.ResultValue(branch_var, year, unit))


def fetch_values_rs(
    variable_obj,
    *,
    region: str,
    scenario: str,
    year: int,
    unit: str = "",
    filter_str: str = "",
) -> float:
    """Retrieve a result using Variable.ValueRS with an optional filter."""
    return float(variable_obj.ValueRS(region, scenario, year, unit, filter_str))


def ensure_parent_dirs(paths: Iterable[Path]) -> None:
    for path in paths:
        Path(path).parent.mkdir(parents=True, exist_ok=True)


#%%
# Final cell marker for Jupyter runners
#%%
