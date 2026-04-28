#%%
# Power sector mapping example using LEAP export-driven branch creation and fill.
#%%

# --- Imports ---
import sys
from pathlib import Path

# Allow repo root on sys.path so code imports resolve without install
REPO_ROOT = Path(__file__).resolve().parents[1]
if REPO_ROOT.exists() and str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.configuration.config import (
    BRANCH_DEMAND_CATEGORY,
    BRANCH_DEMAND_TECHNOLOGY,
    BRANCH_DEMAND_FUEL,
    BRANCH_KEY_ASSUMPTION_BRANCH,
    BRANCH_KEY_ASSUMPTION_CATEGORY,
)


# --- Functions ---
def _ensure_repo_on_path():
    """Allow repo root on sys.path so code imports resolve without install."""
    repo_root = Path(__file__).resolve().parents[1]
    if repo_root.exists() and str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))


def run_power_mapping_example():
    """Create and fill LEAP branches from an export file for power sector workflows."""
    _ensure_repo_on_path()

    from codebase.functions.leap_core import (
        fill_branches_from_export_file,
        create_branches_from_export_file,
        connect_to_leap,
    )
    from codebase.functions.analysis_input_write_dispatcher import (
        dispatch_analysis_input_write,
        get_analysis_input_write_mode,
    )

    write_mode = get_analysis_input_write_mode()
    if write_mode == "workbook":
        dispatch_analysis_input_write(
            export_path=Path(LEAP_EXPORT_FILENAME),
            sheet_name=SHEET_NAME,
            scenario=SCENARIO,
            region=REGION,
            context_label="examples.power_mapping_example",
        )
        return

    L = connect_to_leap()

    if CREATE_BRANCHES_FROM_EXPORT_FILE:
        create_branches_from_export_file(
            L,
            LEAP_EXPORT_FILENAME,
            sheet_name=SHEET_NAME,
            branch_path_col="Branch Path",
            branch_root=ROOT,
            scenario=SCENARIO,
            region=REGION,
            branch_type_mapping=None,
            default_branch_type=(
                BRANCH_DEMAND_CATEGORY,
                BRANCH_DEMAND_CATEGORY,
                BRANCH_DEMAND_TECHNOLOGY,
            ),
            RAISE_ERROR_ON_FAILED_BRANCH_CREATION=True,
        )

    if FILL_BRANCHES_FROM_EXPORT_FILE:
        fill_branches_from_export_file(
            L,
            LEAP_EXPORT_FILENAME,
            sheet_name=SHEET_NAME,
            scenario=SCENARIO,
            region=REGION,
            RAISE_ERROR_ON_FAILED_SET=True,
            HANDLE_CURRENT_ACCOUNTS_TOO=HANDLE_CURRENT_ACCOUNTS_TOO,
        )

# --- Constants to toggle blocks ---
LEAP_EXPORT_FILENAME = "../data/USA_power_leap_import_REF.xlsx"
SHEET_NAME = "Export"
SCENARIO = "Target"
REGION = "Region 1"
ROOT = r""

CREATE_BRANCHES_FROM_EXPORT_FILE = False
FILL_BRANCHES_FROM_EXPORT_FILE = True
HANDLE_CURRENT_ACCOUNTS_TOO = True

# --- Run blocks ---
run_power_mapping_example()

#%%
