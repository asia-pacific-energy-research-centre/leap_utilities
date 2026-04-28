#%%
# Constants mapped to LEAP BranchType enumeration values
# According to LEAP TypeLib: 1 = DemandCategoryBranchType,
# 4 = DemandTechnologyBranchType, 36 = DemandFuelBranchType
import sys
from pathlib import Path
from traitlets import Tuple

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

from codebase.functions.leap_core import (
    fill_branches_from_export_file,
    create_branches_from_export_file,
    connect_to_leap
)
from codebase.functions.analysis_input_write_dispatcher import (
    dispatch_analysis_input_write,
    get_analysis_input_write_mode,
)
from codebase.functions.leap_excel_io import (
    copy_energy_spreadsheet_into_leap_import_file,
)
# Connect to LEAP only in API mode.
WRITE_MODE = get_analysis_input_write_mode()
L = connect_to_leap() if WRITE_MODE == "api" else None

# Define parameters
leap_export_filename = '../outputs/leap_balances_export_file.xlsx'
energy_spreadsheet_filename = '../data/merged_file_energy_ALL_20250814_pre_trump.csv'
# Use merged_file_energy_ALL_20251106.csv and merged_file_energy_00_APEC_20251106 for exact 9th edition projection matching.
ECONOMY = '20_USA'
BASE_YEAR = 2022
SUBTOTAL_COLUMN = 'subtotal_layout'
SCENARIO = "reference"
ROOT = r"Key Assumptions\Energy Balances"
REGION = "Region 1"
DROP_ZERO_BRANCHES = True
sheet_name = "Energy_Balances"
variable_col_value="Activity Level"#turns out that if u are doing key assumptions, u need to specify the variable col value as "Activity Level" even if it is some other measure, like energy.
units = "PJ"
filters_dict = {
    "sectors": ["15_transport_sector"]
}
#%%
# Copy energy spreadsheet into LEAP import file
COPY_ENERGY_SPREADSHEET_INTO_LEAP_IMPORT_FILE = False
if COPY_ENERGY_SPREADSHEET_INTO_LEAP_IMPORT_FILE:
    copy_energy_spreadsheet_into_leap_import_file(
        leap_export_filename=leap_export_filename,
        energy_spreadsheet_filename=energy_spreadsheet_filename,
        ECONOMY=ECONOMY,
        BASE_YEAR=BASE_YEAR,
        SUBTOTAL_COLUMN=SUBTOTAL_COLUMN,
        SCENARIO=SCENARIO,
        ROOT=ROOT,
        REGION=REGION,
        DROP_ZERO_BRANCHES=DROP_ZERO_BRANCHES,
        sheet_name=sheet_name,
        variable_col_value=variable_col_value,
        units=units,
        filters_dict=filters_dict,
    )

#%%
CREATE_BRANCHES_FROM_EXPORT_FILE = True
FILL_BRANCHES_FROM_EXPORT_FILE = True

if WRITE_MODE == "workbook" and (CREATE_BRANCHES_FROM_EXPORT_FILE or FILL_BRANCHES_FROM_EXPORT_FILE):
    dispatch_analysis_input_write(
        export_path=Path(leap_export_filename),
        sheet_name=sheet_name,
        scenario=SCENARIO,
        region=REGION,
        context_label="examples.balance_tables_example",
    )

if CREATE_BRANCHES_FROM_EXPORT_FILE and WRITE_MODE == "api":
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
        default_branch_type=(BRANCH_KEY_ASSUMPTION_CATEGORY, BRANCH_KEY_ASSUMPTION_CATEGORY, BRANCH_KEY_ASSUMPTION_BRANCH),
        RAISE_ERROR_ON_FAILED_BRANCH_CREATION=True,
    )
#%%
if FILL_BRANCHES_FROM_EXPORT_FILE and WRITE_MODE == "api":
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
