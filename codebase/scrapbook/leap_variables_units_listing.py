#%%
"""
List all LEAP variables and their default result units for the active area.

Outputs:
- CSV at outputs/scrapbook/leap_variables_units.csv with columns:
  branch_full, variable_name, branch_variable_name, default_result_unit, is_data, branch_type, branch_type_name

Notes:
- Windows + pywin32 only (COM). Will fail in WSL.
- Uses LEAP's default units (does not try to set ActiveUnit to avoid errors).
- Leaves LEAP visible=False by default; flip LEAP_VISIBLE if you want to watch.
"""
from __future__ import annotations

import csv
import os
from pathlib import Path

try:
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover
    win32com = None  # sentinel for environments without COM

from codebase.functions.leap_api_guard import ensure_leap_api_allowed

#%% user-editable toggles
LEAP_VISIBLE = False
OUTPUT_PATH = Path("outputs/scrapbook/leap_variables_units.csv")

#%% helpers
def ensure_repo_root() -> Path:
    root = Path(__file__).resolve().parents[2]
    if Path.cwd() != root:
        os.chdir(root)
    return root


def connect_leap(visible: bool = False):
    ensure_leap_api_allowed("scrapbook.leap_variables_units_listing.connect_leap")
    if win32com is None:
        raise SystemExit("pywin32 is required to run this script (Windows only).")
    app = win32com.client.Dispatch("Leap.LEAPApplication")
    app.Visible = bool(visible)
    return app


def iter_branch_tree(branch):
    """Yield branch and all descendants (pre-order)."""
    yield branch
    try:
        children = branch.Children
        for i in range(1, children.Count + 1):
            yield from iter_branch_tree(children.Item(i))
    except Exception:
        return


def list_variables_with_units(app) -> list[dict]:
    records = []
    top = app.Branches
    for i in range(1, top.Count + 1):
        branch = top.Item(i)
        for b in iter_branch_tree(branch):
            full_name = str(b.FullName)
            branch_type = None
            try:
                branch_type = b.BranchType
            except Exception:
                pass
            branch_type_name = None
            try:
                branch_type_name = b.BranchTypeName
            except Exception:
                pass
            vars_obj = b.Variables
            for v_idx in range(1, vars_obj.Count + 1):
                v = vars_obj.Item(v_idx)
                try:
                    default_unit = v.DefaultResultUnit.Name
                except Exception:
                    default_unit = ""
                try:
                    is_data = bool(v.IsData)
                except Exception:
                    is_data = None
                try:
                    bv_name = v.BranchVariableName
                except Exception:
                    bv_name = ""
                records.append(
                    {
                        "branch_full": full_name,
                        "variable_name": str(v.Name),
                        "branch_variable_name": str(bv_name),
                        "default_result_unit": str(default_unit),
                        "is_data": is_data,
                        "branch_type": branch_type,
                        "branch_type_name": branch_type_name,
                    }
                )
    return records


def write_csv(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        return
    fieldnames = list(rows[0].keys())
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


#%% run
if __name__ == "__main__":
    ensure_repo_root()
    app = connect_leap(visible=LEAP_VISIBLE)
    records = list_variables_with_units(app)
    write_csv(OUTPUT_PATH, records)
    print(f"Wrote {len(records)} rows to {OUTPUT_PATH}")

#%%
