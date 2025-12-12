#%%
# Constants mapped to LEAP BranchType enumeration values
# According to LEAP TypeLib: 1 = DemandCategoryBranchType,
# 4 = DemandTechnologyBranchType, 36 = DemandFuelBranchType
BRANCH_DEMAND_CATEGORY = 1
BRANCH_DEMAND_TECHNOLOGY = 4
BRANCH_DEMAND_FUEL = 36
BRANCH_KEY_ASSUMPTION_BRANCH = 9#contains number
BRANCH_KEY_ASSUMPTION_CATEGORY = 10#contains many sub-branches
# Hypothetical value for key assumptions
#below are all teh unique values from the leap typelib for branch types
#  1=DemandCategoryBranchType, 2=TransformationModuleBranchType, 3=TransformationProcessBranchType, 4=DemandTechnologyBranchType, 5=TransformationProcessCategoryType, 6=TransformationOutputCategoryType, 7=TransformationOutputBranchType, 9=KeyAssumptionCategoryType, 10=KeyAssumptionBranchType, 11=ResourceRootType, 12=PrimaryBranchCategoryType, 13=SecondaryBranchCategoryType, 15=ResourceBranchType, 16=ResourceDisagType, 18=StatDiffRootType, 19=StockChangeRootType, 20=StatDiffPrimaryCategoryType, 21= StatDiffSecondaryCategoryType, 22=StockChangePrimaryCategoryType, 23=StockChangeSecondaryCategoryType, 24=StatDiffBranchType, 25=  StockChangeBranchType, 26=NonEnergyCategoryType, 27=NonEnergyBranchType, 30=AuxCategoryType, 31=AuxBranchType, 32=FeedstockCategoryType, 33= FeedstockBranchType, 34=DMDPollutionBranchType, 35=TransformationPollutionBranchType, 36=DemandFuelBranchType, 37=IndicatorCategoryType, 38=IndicatorBranchType, 39=EmissionConstraintBranchType"
#these can be looked up in config/TypeLib_LEAP_API_full.txt
# e.g.         
# [id(0x0000012a), propget, helpstring("Adds a new key assumption branch with name BName and the specified scale and units below branch ParentID.")]
# HRESULT AddKeyAssumption(
#                 [in] int ParentID, 
#                 [in] VARIANT BName, 
#                 [in] VARIANT Scale, 
#                 [in] VARIANT KUnit, 
#                 [out, retval] ILEAPBranch** Value);
# [id(0x0000012e), propget, helpstring("Adds a new key assumption category branch with name BName below branch ParentID.")]
# HRESULT AddKeyAssumptionCategory(
#                 [in] int ParentID, 
#                 [in] VARIANT BName, 
#                 [out, retval] ILEAPBranch** Value);

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

# Define parameters
leap_export_filename = '../results/leap_balances_export_file.xlsx'
energy_spreadsheet_filename = '../data/merged_file_energy_ALL_20250814.csv'
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
