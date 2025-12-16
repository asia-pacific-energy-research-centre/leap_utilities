BRANCH_DEMAND_CATEGORY = 1
BRANCH_DEMAND_TECHNOLOGY = 4
BRANCH_DEMAND_FUEL = 36
BRANCH_KEY_ASSUMPTION_BRANCH = 9#contains number
BRANCH_KEY_ASSUMPTION_CATEGORY = 10#contains many sub-branches
import sys
from pathlib import Path
from traitlets import Tuple

# Allow repo root on sys.path so leap_utils imports resolve without install
REPO_ROOT = Path(__file__).resolve().parents[1]
if REPO_ROOT.exists() and str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from leap_utils.leap_core import (
    fill_branches_from_export_file,
    create_branches_from_export_file,
    connect_to_leap
)
from leap_utils.leap_excel_io import (
    copy_energy_spreadsheet_into_leap_import_file,
)
# Connect to LEAP
L = connect_to_leap()
leap_export_filename = '../results/leap_balances_export_file.xlsx'
sheet_name = "Energy_Balances"
CREATE_BRANCHES_FROM_EXPORT_FILE = True

# Define parameters
leap_export_filename = '../data/industry export.xlsx'
ECONOMY = '20_USA'
BASE_YEAR = 2022
SUBTOTAL_COLUMN = 'subtotal_layout'
SCENARIO = "Reference"
ROOT = r""
REGION = "United States of America"
sheet_name = "Export"

if CREATE_BRANCHES_FROM_EXPORT_FILE:
    # Create branches from export file
    create_branches_from_export_file(
        L,
        leap_export_filename,
        sheet_name=sheet_name,
        branch_path_col="Branch Path",
        branch_root=ROOT,
        scenario=SCENARIO,
        region=REGION,
        branch_type_mapping=None,
        default_branch_type=(BRANCH_DEMAND_CATEGORY, BRANCH_DEMAND_CATEGORY, BRANCH_DEMAND_TECHNOLOGY),
        RAISE_ERROR_ON_FAILED_BRANCH_CREATION=True,
    )
#%%
FILL_BRANCHES_FROM_EXPORT_FILE = True
if FILL_BRANCHES_FROM_EXPORT_FILE:
    # Fill branches with data from export file
    fill_branches_from_export_file(
        L,
        leap_export_filename,
        sheet_name=sheet_name,
        scenario=SCENARIO,
        region=REGION,
        RAISE_ERROR_ON_FAILED_SET=True,
    )

#%%
