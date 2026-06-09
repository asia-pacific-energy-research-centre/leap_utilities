#%%
# Configuration constants for LEAP utilities, including branch types and metadata.
#%%

#%%
# --- Branch Type Constants ---
BRANCH_DEMAND_CATEGORY = 1
BRANCH_DEMAND_TECHNOLOGY = 4
BRANCH_DEMAND_FUEL = 36
BRANCH_KEY_ASSUMPTION_BRANCH = 9
BRANCH_KEY_ASSUMPTION_CATEGORY = 10
BRANCH_TRANSFORMATION_MODULE = 2
BRANCH_TRANSFORMATION_PROCESS = 3
BRANCH_PROCESS_CATEGORY = 5
BRANCH_OUTPUT_CATEGORY = 6
BRANCH_OUTPUT = 7
BRANCH_RESOURCE_ROOT = 11
BRANCH_RESOURCE_PRIMARY_CATEGORY = 12
BRANCH_RESOURCE_SECONDARY_CATEGORY = 13
BRANCH_RESOURCE_BRANCH = 15
BRANCH_RESOURCE_DISAG = 16
BRANCH_FEEDSTOCK_CATEGORY = 32
BRANCH_FEEDSTOCK_BRANCH = 33
BRANCH_AUX_CATEGORY = 30
BRANCH_AUX_BRANCH = 31
# Base year for derived share calculations and projections.
BASE_YEAR = 2022
#%%

#%%
# --- LEAP API hints (branch creation & data access) ---
LEAP_API_BRANCH_CREATION_METHODS = [
    {
        "name": "AddCategory",
        "context": "Demand tree",
        "description": (
            "Adds a demand category branch (Activity Level) under `ParentID`. "
            "Pass the new branch name, scale, and activity unit strings."
        ),
        "use_cases": ["High-level demand categories", "Organization nodes for technologies"],
    },
    {
        "name": "AddTechnology",
        "context": "Demand tree (Activity Analysis)",
        "description": (
            "Creates a demand technology tied to a single fuel. "
            "Needs a category parent and explicit fuel/energy unit."
        ),
        "use_cases": ["Technology nodes with Activity Level variables"],
    },
    {
        "name": "AddProcess",
        "context": "Transformation tree",
        "description": (
            "Creates a transformation process under a process category. "
            "Supply the first feedstock fuel name so the branch links to the "
            "correct fuel object."
        ),
        "use_cases": ["Transformation branches like Blast furnace process"],
    },
    {
        "name": "AddOutput",
        "context": "Transformation tree",
        "description": "Adds an output fuel branch beneath an output category node.",
        "use_cases": ["Register process outputs"],
    },
    {
        "name": "AddFeedstock",
        "context": "Transformation tree",
        "description": "Attaches feedstock or auxiliary fuels to a process branch.",
        "use_cases": ["Feedstock/auxiliary branches per process"],
    },
    {
        "name": "AddSimpleProcess",
        "context": "Transformation tree",
        "description": (
            "Shortcut for simple process creation with one feedstock and one "
            "output fuel."
        ),
        "use_cases": ["Quick single input/output processes"],
    },
    {
        "name": "AddModule",
        "context": "Transformation module",
        "description": (
            "Defines module metadata (application type, capacity, co-products, "
            "auxiliary handling) for advanced transformation modelling."
        ),
        "use_cases": ["Modules with custom efficiencies or load curves"],
    },
    {
        "name": "AddKeyAssumption",
        "context": "Key assumptions tree",
        "description": (
            "Adds a key assumption branch with scale/unit text below an assumption "
            "category."
        ),
        "use_cases": ["Assumption variables per sector or scenario"],
    },
    {
        "name": "AddKeyAssumptionCategory",
        "context": "Key assumptions tree",
        "description": "Creates a grouping node for key assumptions.",
        "use_cases": ["Organizing assumption branches"],
    },
]

LEAP_API_VARIABLE_ACCESS = [
    {
        "path": "Branch(branch_path).Variables(variable_name)",
        "description": (
            "Returns a `LEAPVariable` for the branch/variable combination. "
            "Equivalent to `BranchVariable('Branch:Variable')`."
        ),
        "actions": [
            "Use `variable.Expression` to read/write interpolated data.",
            "Use `variable.InheritedExpression` to check scenario inheritance.",
        ],
    },
    {
        "path": "BranchVariable('branch:variable')",
        "description": "Alternative way to target the same `LEAPVariable` object.",
        "actions": ["Set `.Expression` or inspect `.ExpressionRS` for scenario data."],
    },
    {
        "path": "LEAPVariable.Expression",
        "description": "Controls the Data()/Interp() expression that drives the variable.",
        "actions": [
            "Write expressions built from year-value pairs (our export uses Data()).",
            "Read expressions for debugging or verifying auto-generated values.",
        ],
    },
]

LEAP_API_NOTES = (
    "See `config/TypeLib_LEAP_API_full.txt` for the full IDL reference. The "
    "collections above capture the most useful branch creation and variable "
    "methods we touch from that file."
)

#%%
# --- Scenario/Region Constants ---
scenario_dict = {
    "Current Accounts": {
        "scenario_name": "Current Accounts",
        "scenario_code": "CA",
        "scenario_id": 1,
    },
    "Target": {
        "scenario_name": "Target",
        "scenario_code": "TGT",
        "scenario_id": 3,
    },
    "Reference": {
        "scenario_name": "Reference",
        "scenario_code": "REF",
        "scenario_id": 4,
    },
}
region_id_name_dict = {
    "12_NZ": {
        "region_id": 2,
        "region_name": "New Zealand",
        "region_code": "12_NZ",
    },
    "20_USA": {
        "region_id": 1,
        "region_name": "United States",
        "region_code": "20_USA",
    },
}
#%%

# --- LEAP Units by ID ---
#add config to path
import os
import sys
from pathlib import Path

_env_root = os.environ.get("LEAP_UTILITIES_REPO_ROOT", "").strip()
REPO_ROOT = Path(_env_root) if _env_root else Path(__file__).resolve().parents[2]
#add paths to sys.path so we can import config files. This is complicated by the way other files from different repos import this and run it, so we have to add paths to sys.path instead of changing cwd.

#Path(__file__).resolve().parents[1]
# CURRENT_DIR = Path.cwd()
# if CURRENT_DIR != REPO_ROOT:
#     os.chdir(REPO_ROOT)
# if str(CURRENT_DIR) not in sys.path:
sys.path.insert(0, str(REPO_ROOT))
#add config to path --- END ---
CONFIG_DIR = REPO_ROOT / "config"
sys.path.insert(0, str(CONFIG_DIR))
from LEAP_API_helpers import LEAP_UNITS_BY_ID

# %%
