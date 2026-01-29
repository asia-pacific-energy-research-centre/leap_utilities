# ============================================================
# LEAP_core.py
# ============================================================
# Core helper functions for LEAP transport data integration.
# Provides connection, diagnostics, normalization, logging,
# and activity level utilities shared by loader scripts.
# ============================================================

import contextlib
import io
import math
import re

import pandas as pd

try:  # pragma: no cover - windows-only
    from win32com.client import Dispatch, GetActiveObject, gencache
    _WIN32COM_AVAILABLE = True
except ImportError:  # pragma: no cover - windows-only
    Dispatch = GetActiveObject = gencache = None
    _WIN32COM_AVAILABLE = False

from .config import (
    BRANCH_DEMAND_CATEGORY,
    BRANCH_DEMAND_TECHNOLOGY,
    BRANCH_DEMAND_FUEL,
    BRANCH_KEY_ASSUMPTION_BRANCH,
    BRANCH_KEY_ASSUMPTION_CATEGORY,
    BRANCH_TRANSFORMATION_MODULE,
    BRANCH_TRANSFORMATION_PROCESS,
    BRANCH_PROCESS_CATEGORY,
    BRANCH_OUTPUT_CATEGORY,
    BRANCH_OUTPUT,
    BRANCH_FEEDSTOCK_CATEGORY,
    BRANCH_FEEDSTOCK_BRANCH,
    BRANCH_RESOURCE_ROOT,
    BRANCH_RESOURCE_PRIMARY_CATEGORY,
    BRANCH_RESOURCE_SECONDARY_CATEGORY,
    BRANCH_RESOURCE_BRANCH,
    BRANCH_RESOURCE_DISAG,
    BRANCH_AUX_CATEGORY,
    BRANCH_AUX_BRANCH,
    LEAP_UNITS_BY_ID
)

# Optional transport-specific mappings; if unavailable, functions accept injected mappings instead
try:  # pragma: no cover - optional dependency
    from transport_branch_mappings import (
        ESTO_SECTOR_FUEL_TO_LEAP_BRANCH_MAP,
        LEAP_BRANCH_TO_SOURCE_MAP,
        SHORTNAME_TO_LEAP_BRANCHES,
        LEAP_MEASURE_CONFIG,
    )
except Exception:  # pragma: no cover - optional dependency
    ESTO_SECTOR_FUEL_TO_LEAP_BRANCH_MAP = None
    LEAP_BRANCH_TO_SOURCE_MAP = None
    SHORTNAME_TO_LEAP_BRANCHES = None
    LEAP_MEASURE_CONFIG = None

try:  # pragma: no cover - optional dependency
    from transport_measure_metadata import SHORTNAME_TO_ANALYSIS_TYPE
except Exception:  # pragma: no cover - optional dependency
    SHORTNAME_TO_ANALYSIS_TYPE = None

try:  # pragma: no cover - optional dependency
    from transport_measure_catalog import LEAP_BRANCH_TO_ANALYSIS_TYPE_MAP
except Exception:  # pragma: no cover - optional dependency
    LEAP_BRANCH_TO_ANALYSIS_TYPE_MAP = None

try:  # pragma: no cover - optional dependency
    from transport_branch_expression_mapping import (
        LEAP_BRANCH_TO_EXPRESSION_MAPPING,
        ALL_YEARS,
    )
except Exception:  # pragma: no cover - optional dependency
    LEAP_BRANCH_TO_EXPRESSION_MAPPING = None
    ALL_YEARS = None

# Prompt on branch-creation warnings to pause execution (default on).
ASK_ON_MISSING_BRANCH_CREATION = True


def _prompt_on_missing_branch_creation(message: str) -> None:
    """Pause execution on branch-creation warnings so the user can decide to continue."""
    if not ASK_ON_MISSING_BRANCH_CREATION:
        return
    prompt = f"{message}\nEnter 'c' to continue or 'b' to break: "
    while True:
        choice = input(prompt).strip().lower()
        if choice in ("c", "continue", ""):
            return
        if choice in ("b", "break", "q", "quit"):
            raise RuntimeError(f"User aborted after warning: {message}")
        print("Please enter 'c' to continue or 'b' to break.")


def _require_global(name: str, val):
    if val is None:
        raise ImportError(
            f"{name} not available; pass it explicitly to this function or install the transport mappings."
        )
    return val

# ------------------------------------------------------------
# Connection & Core Helpers
# ------------------------------------------------------------


def is_leap_api_available():
    """Return True when win32com/LEAP API is importable."""
    return _WIN32COM_AVAILABLE

_LEAP_TYPELIB = ("{6161465F-91BE-4B6B-8BB0-361F5BFA612A}", 0, 2, 3)


def _quiet_com_cache_refresh():
    """Rebuild win32com cache and ensure LEAP typelibs, suppressing noisy makepy output."""
    if gencache is None:
        return
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer), contextlib.redirect_stderr(buffer):
        try:
            gencache.Rebuild()
        except Exception:
            pass
        try:
            gencache.EnsureModule(*_LEAP_TYPELIB)
        except Exception:
            pass


def _ensure_leap_com_wrappers():
    """Ensure LEAP COM wrappers exist; rebuild cache quietly on failure."""
    if gencache is None:
        return
    try:
        gencache.EnsureDispatch("LEAP.LEAPApplication")
    except Exception:
        _quiet_com_cache_refresh()
        gencache.EnsureDispatch("LEAP.LEAPApplication")


def connect_to_leap(force_rebuild: bool = True):
    """Enhanced LEAP connection with project readiness checks."""
    if not _WIN32COM_AVAILABLE:
        raise RuntimeError(
            "LEAP API (`win32com`) is unavailable in this environment (Linux/WSL). "
            "Run this script from Windows with pywin32 installed to reach LEAP."
        )
    print("[INFO] Connecting to LEAP...")
    
    try:
        if force_rebuild:
            _quiet_com_cache_refresh()
        _ensure_leap_com_wrappers()
        try:
            leap_app = GetActiveObject("LEAP.LEAPApplication")
            print("[SUCCESS] Connected to existing LEAP instance")
        except:
            leap_app = Dispatch("LEAP.LEAPApplication")
            print("[SUCCESS] Created new LEAP instance")
        
        # Check if LEAP is ready for Branch() calls
        try:
            areas = leap_app.Areas
            if areas.Count == 0:
                print("[WARN] LEAP has no project loaded - Branch() calls will fail")
                print("[WARN] Please load a project in LEAP first")
            else:
                active_area = leap_app.ActiveArea
                print(f"[INFO] LEAP ready - Active area: '{active_area}' with {areas.Count} area(s)")
        except Exception as e:
            print(f"[WARN] Cannot check LEAP project state: {e}")
        
        return leap_app
        
    except Exception as e:
        print(f"[ERROR] LEAP connection failed: {e}")
        return None

def safe_branch_call(leap_obj, branch_path, AUTO_SET_MISSING_BRANCHES=False, THROW_ERROR_ON_MISSING=True, timeout_msg=True):
    """
    Safe Branch() call that won't hang - use this instead of L.Branch() directly.
    
    Args:
        leap_obj: LEAP application object
        branch_path: string path to branch (e.g., "Demand", "Key\\Population")
        timeout_msg: whether to print timeout messages
        
    Returns:
        branch object if successful, None if failed
        
    Usage:
        L = connect_to_leap()
        branch = safe_branch_call(L, "Demand")
        if branch:
            variables = branch.Variables
        else:
            print("Branch not found")
    """
    if leap_obj is None:
        return None
    
    branches = leap_obj.Branches
    try:
        exists = branches.Exists(branch_path)
    except Exception as e:
        breakpoint()
        raise Exception(f"Branches.Exists failed for '{branch_path}': {e}")

    if not exists:
        if AUTO_SET_MISSING_BRANCHES:
            print(f"[INFO] AUTO_SET_MISSING_BRANCHES is set to true. The branch will be auto-created: {branch_path}")
            #set it 
        elif THROW_ERROR_ON_MISSING:
            breakpoint()
            raise Exception(f"Branches.Exists returned false for '{branch_path}'. AUTO_SET_MISSING_BRANCHES is False and THROW_ERROR_ON_MISSING is true so throwing an error.")
        else:
            pass# THROW_ERROR_ON_MISSING is false so we just want to return None
        return None

    branch = leap_obj.Branch(branch_path)
    return branch
    # except Exception as e:
    #     if timeout_msg:
    #         error_str = str(e)
    #         if len(error_str) > 60:
    #             error_str = error_str[:60] + "..."
    #         print(f"[INFO] Branch '{branch_path}' not accessible: {error_str}")
    #     return None


def build_expr(points, expression_type="Interp"):
    """Build a LEAP-compatible Interp() expression."""
    if not points:
        return None
    df = pd.DataFrame(points, columns=["year", "value"]).dropna(subset=["year", "value"])
    if df["year"].duplicated().any():
        breakpoint()
    df = df.sort_values("year")
    pts = list(zip(df["year"].astype(int), df["value"].astype(float)))
    if len(pts) == 1:
        return str(pts[0][1])
    if expression_type == "":
        raise ValueError("expression_type cannot be empty string if the number of points is greater than 1.")
    return f"{expression_type}(" + ", ".join(f"{y}, {v:.6g}" for y, v in pts) + ")"


def safe_set_variable(L, obj, varname, expr, unit_name=None, context=""):
    """Safely assign expressions to LEAP variables with logging."""
        
    try:
        var = obj.Variable(varname)
        if var is None:
            print(f"[WARN] Missing variable '{varname}' on {context} within LEAP.")
            return False
        prev_expr = var.Expression
        if prev_expr and prev_expr.strip():
            print(f"[INFO] Clearing previous expression for '{varname}' on {context}")
            var.Expression = ""
            try:
                obj.Application.RefreshBranches()
            except Exception:
                pass
        var.Expression = expr
        #check that the expression is a string
        short_expr = var.Expression[:80] + ("..." if len(var.Expression) > 80 else "")
        print(f"[SET] {context} → {varname} = {short_expr}")
        
        _set_variable_unit(L, var, unit_name, context=context)
        # Set scale if provided #NOTE i tried to set scale here but it didnt work. cannot access Scales from var.
        # if scale_value is None:
        #     return True
        # breakpoint()#is there a L.Scales var? how to set scale? its important for % especially
        # scales = L.Scales
        ########
        return True
    except Exception as e:
        print(f"[ERROR] Failed setting {varname} on {context}: {e}")
        return False


def _set_variable_unit(L, var, unit_name, context=""):
    """Assign DataUnit when available; warn or raise when missing."""
    if unit_name is None:
        return True
    if isinstance(unit_name, float) and math.isnan(unit_name):
        return True
    if isinstance(unit_name, str) and unit_name.strip().lower() in {"", "nan"}:
        return True
    units = L.Units
    if not units.Exists(unit_name):
        raise ValueError(f"Unit not found: {unit_name}. Unit not set.")
    unit = units.Item(unit_name)  # returns ILEAPUnit
    if unit is None:
        known_units = {u.get("name") for u in LEAP_UNITS_BY_ID.values()}
        if unit_name not in known_units:
            print(
                f"[WARN] Unit name '{unit_name}' not found in LEAP units list. Proceeding without setting unit."
            )
            return True
        raise ValueError(
            f"Unit name '{unit_name}' found in LEAP_UNITS_BY_ID but ILEAPUnit not found in LEAP. Cannot set unit."
        )
    var.DataUnit = unit  # or: var.DataUnitID = unit.ID
    return True

def define_value_based_on_src_tuple(meta_values, src_tuple):
    ttype, medium, vtype, drive, fuel = tuple(list(src_tuple) + [None] * (5 - len(src_tuple)))[:5]
    for col in ['LEAP_units', 'LEAP_Scale', 'LEAP_Per']:
        val = meta_values.get(col)
        if val is not None and isinstance(val, str) and '$' in val:
            # extract the options. if there are multiple $'s throw an error, code is not designed for that
            parts = val.split('$')
            if len(parts) != 2:
                raise ValueError(f"Unexpected format for metadata value: {val}")
            #now we have special code based on what the pklaceholder is
            if val == 'Passenger-km$Tonne-km':
                if 'passenger' in ttype:
                    resolved_value = 'Passenger-km'
                elif 'freight' in ttype:
                    resolved_value = 'Tonne-km'
                else:
                    raise ValueError(f"Unexpected ttype for resolving Passenger-km$Tonne-km: {ttype}")
                meta_values[col] = resolved_value
            elif val == 'of Tonne-km$of Passenger-km':
                if 'passenger' in ttype:
                    resolved_value = 'of Passenger-km'
                elif 'freight' in ttype:
                    resolved_value = 'of Tonne-km'
                else:
                    raise ValueError(f"Unexpected ttype for resolving of Tonne-km$of Passenger-km: {ttype}")
                meta_values[col] = resolved_value
            else:
                raise ValueError(f"Unknown placeholder in metadata value: {val}")
    return meta_values
# ------------------------------------------------------------
# Activity Levels
# ------------------------------------------------------------
# def ensure_activity_levels(L, TRANSPORT_ROOT=r"Demand"):
#     """Ensure 'Activity Level' variables exist in all transport branches."""
#     print("\n=== Checking and fixing Activity Levels ===")
#     try:
#         transport_branch = safe_branch_call(L, TRANSPORT_ROOT, , AUTO_SET_MISSING_BRANCHES=AUTO_SET_MISSING_BRANCHES)
#         if transport_branch:
#             if not transport_branch.Variable("Activity Level").Expression:
#                 transport_branch.Variable("Activity Level").Expression = "100"
#             for sub in ["Passenger", "Freight"]:
#                 try:
#                     b = L.Branch(f"{TRANSPORT_ROOT}\\{sub}")
#                     if not b.Variable("Activity Level").Expression:
#                         b.Variable("Activity Level").Expression = "50"
#                 except Exception:
#                     print(f"[WARN] Could not access {TRANSPORT_ROOT}\\{sub}")
#         else:
#             print("[WARN] Could not access Demand branch - skipping Activity Level setup")
#     except Exception as e:
#         print(f"[ERROR] Activity Level setup failed: {e}")
#     print("==============================================\n")



# ------------------------------------------------------------
# Logging
# ------------------------------------------------------------
def create_transport_export_df():
    """Initialize DataFrame to log all data written to LEAP."""
    return pd.DataFrame(columns=[
        'Date', 'Transport_Type', 'Medium', 'Vehicle_Type', 'Technology', 'Fuel',
        'Measure', 'Value', 'Branch_Path', 'LEAP_Tuple', 'Source_Tuple'
    ])

def write_row_to_leap_export_df(export_df, leap_tuple, src_tuple, branch_path, measure, df_m):
    """Add processed measure data to the export DataFrame."""
    new_rows = []
    for _, row in df_m.iterrows():
        if pd.notna(row[measure]):
            new_rows.append({
                'Date': int(row["Date"]),
                'Transport_Type': leap_tuple[0] if len(leap_tuple) > 0 else pd.NA,
                'Medium': leap_tuple[1] if len(leap_tuple) > 1 else pd.NA,
                'Vehicle_Type': leap_tuple[2] if len(leap_tuple) > 2 else pd.NA,
                'Technology': leap_tuple[3] if len(leap_tuple) > 3 else pd.NA,
                'Fuel': leap_tuple[4] if len(leap_tuple) > 4 else pd.NA,
                'Measure': measure,
                'Value': float(row[measure]),
                'Branch_Path': branch_path,
                'LEAP_Tuple': str(leap_tuple),
                'Source_Tuple': str(src_tuple)
            })
    if new_rows:
        new_df = pd.DataFrame(new_rows)
        export_df = pd.concat([export_df, new_df], ignore_index=True) if not export_df.empty else new_df.copy()
    return export_df


def save_leap_export_df(export_df, filename="leap_export.xlsx"):#, log_tuple=None):
    """Save the complete LEAP data log to Excel with summaries."""
    print(f"\n=== Saving LEAP Data for exporting to LEAP to {filename} ===")
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        export_df.to_excel(writer, sheet_name='All_Data', index=False)
    print(f"✅ Saved {len(export_df)} data points to {filename}")
    print("=" * 50)


def build_expression_from_mapping(
    branch_tuple,
    df_m,
    measure,
    mapping=None,
    all_years=None,
):
    """
    Builds the correct LEAP expression for a branch based on LEAP_BRANCH_TO_EXPRESSION_MAPPING.
    
    Parameters:
    - branch_tuple: tuple key from LEAP_BRANCH_TO_EXPRESSION_MAPPING
    - df_m: DataFrame containing 'Date' and the measure column
    - measure: measure name string (e.g., 'Stock Share', 'Activity Level')

    Returns:
    - expr: string suitable for LEAP variable.Expression
    """
    mapping = mapping or _require_global(
        "LEAP_BRANCH_TO_EXPRESSION_MAPPING", LEAP_BRANCH_TO_EXPRESSION_MAPPING
    )
    all_years = all_years or _require_global("ALL_YEARS", ALL_YEARS)

    entry = (measure,) + branch_tuple
    mapping_entry = mapping.get(entry, ("Data", all_years))
    mode, arg = mapping_entry
    #check if there is only one value, in which case set to SingleValue
    if mode != 'SingleValue':
        valid = df_m[pd.notna(df_m['Value'])]
        if len(valid) == 1:
            mode = 'SingleValue'
    # Default: Data from all available years
    if mode == 'Data':
        pts = [
            (int(r["Date"]), float(r['Value']))
            for _, r in df_m.iterrows()
            if pd.notna(r['Value'])
        ]
        return build_expr(pts, "Data") if pts else None, 'Data'

    # Interp between given years
    elif mode == 'Interp':
        start, end = arg[0], arg[-1]
        df_filtered = df_m[(df_m["Date"] >= start) & (df_m["Date"] <= end)]
        pts = [
            (int(r["Date"]), float(r['Value']))
            for _, r in df_filtered.iterrows()
            if pd.notna(r['Value'])
        ]
        return build_expr(pts, "Interp") if pts else None, 'Interp'

    # Flat value (constant for a single year)
    elif mode == 'Flat':
        year = arg[0]
        val = df_m.loc[df_m["Date"] == year, measure].mean()
        return str(float(val)) if pd.notna(val) else None, 'Flat'

    # Return only the number if exactly one data point exists
    elif mode == 'SingleValue':
        valid = df_m[pd.notna(df_m['Value'])]
        if len(valid) == 1:
            return str(float(valid['Value'].iloc[0])), 'SingleValue'
        print(f"[WARN] Expected single value for {branch_tuple} but found {len(valid)} rows. Falling back to Data.")
        pts = [
            (int(r["Date"]), float(r['Value']))
            for _, r in valid.iterrows()
        ]
        return build_expr(pts, "Data") if pts else None, 'Data'

    # Custom function for special logic
    elif mode == 'Custom':
        func_name = arg
        try:
            func = globals().get(func_name)
            if callable(func):
                return func(branch_tuple, df_m, measure), 'Custom'
            else:
                print(f"[WARN] Custom function '{func_name}' not found for {branch_tuple}")
                return None, None
        except Exception as e:
            print(f"[ERROR] Custom expression failed for {branch_tuple}: {e}")
            return None, None

    # Default fallback
    else:
        print(f"[WARN] Unknown mode '{mode}' for {branch_tuple}. Using raw data.")
        pts = [
            (int(r["Date"]), float(r['Value']))
            for _, r in df_m.iterrows()
            if pd.notna(r['Value'])
        ]
        return build_expr(pts, "Data") if pts else None, 'Data'

#%%

#################################################
# Auto-Creation of LEAP Branches
#################################################
# ------------------------------------------------------------
# Constants mapped to LEAP BranchType enumeration values
# According to LEAP TypeLib: 1 = DemandCategoryBranchType,
# 4 = DemandTechnologyBranchType, 36 = DemandFuelBranchType
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


def _choose_branch_type_for_segment(current_path, segment_name, branch_tuple, shortname_to_leap_branches=None):
    """
    Decide what LEAP branch type to use when auto-creating a missing segment.

    Parameters
    ----------
    current_path : str
        Full path up to (but not including) this segment.
    segment_name : str
        The missing branch name we are about to create.
    branch_tuple : any
        One of the tuples stored in SHORTNAME_TO_LEAP_BRANCHES[key].
        We infer 'shortname' and branch type rules from this.
    """

    # First identify what type of branch_tuple we have by going through
    # all the keys in SHORTNAME_TO_LEAP_BRANCHES and seeing if the
    # branch_tuple matches any of the values.
    shortname_to_leap_branches = shortname_to_leap_branches or _require_global(
        "SHORTNAME_TO_LEAP_BRANCHES", SHORTNAME_TO_LEAP_BRANCHES
    )

    shortname = None
    for key, values in shortname_to_leap_branches.items():
        if branch_tuple in values:
            shortname = key
            break
    
    if shortname is None:
        raise ValueError(f"Branch tuple {branch_tuple} not found in SHORTNAME_TO_LEAP_BRANCHES.")

    short_lower = shortname.lower()

    # ------------------------------------------------------------------
    # STOCK-BASED BRANCHES (contain '(road)' in the shortname)
    # ------------------------------------------------------------------
    # If shortname has (road) in it, it is a stock-based branch and we
    # cannot set its technology-based branches (DemandTechnologyBranchType=4)
    # within the LEAP API. However, we can set its fuel-based branches.
    #
    # So:
    #   - If shortname == 'Fuel (road)': set as DemandFuelBranchType (36)
    #   - Otherwise: raise, user must manually create that branch in LEAP
    # ------------------------------------------------------------------
    if "(road)" in short_lower:
        if shortname == "Fuel (road)":
            return BRANCH_DEMAND_FUEL
        else:
            raise RuntimeError(
                "Attempted to auto-create a stock-based ('(road)') branch that is "
                "not 'Fuel (road)'. LEAP requires these technology/category "
                "branches to be created manually in the UI.\n"
                f"  shortname: {shortname}\n"
                f"  path: {current_path}\\{segment_name}"
            )

    # ------------------------------------------------------------------
    # INTENSITY-BASED BRANCHES (no '(road)' in the shortname)
    # ------------------------------------------------------------------
    # If the shortname is not stock based, then it is intensity based and
    # we have to identify whether it is a technology branch.
    #
    # This is done by checking if the shortname is in:
    #   ['Others (level 2)', 'Fuel (non-road)']
    #
    # Since intensity-based branches don't have fuel branches at the end,
    # only technology branches, 'Fuel (non-road)' is treated as a *technology*.
    #
    # If so, we can set it as a DemandTechnologyBranchType (4).
    # Otherwise, we can set it as a DemandCategoryBranchType (1).
    # ------------------------------------------------------------------
    if shortname in ["Others (level 2)", "Fuel (non-road)"]:
        # Intensity-based technology branch
        return BRANCH_DEMAND_TECHNOLOGY

    # Fallback: generic intensity-based category
    return BRANCH_DEMAND_CATEGORY

def ensure_branch_exists(
    L,
    full_path,
    branch_tuple,
    AUTO_SET_MISSING_BRANCHES=True,
    branch_type_mapping=None,
    shortname_to_leap_branches=None,
):
    """
    Ensures a LEAP branch exists at full_path, creating any missing segments
    using _choose_branch_type_for_segment() and LEAPApplication Add* methods.

    Parameters
    ----------
    L : LEAPApplication COM object
    full_path : str
        Example: "Demand\\Freight non road\\Air\\Aviation gasoline"
    branch_tuple : tuple
        One of the tuples stored in SHORTNAME_TO_LEAP_BRANCHES for this
        logical branch type. Used to infer whether this path is stock-based
        vs intensity-based, and whether a missing segment is a category
        vs technology.
    """
    parts = [p for p in full_path.split("\\") if p]
    parent_branch = None

    for i, part in enumerate(parts):
        current_path = "\\".join(parts[:i+1])
        # Try to get the branch via your safe helper
        br = safe_branch_call(L, current_path, AUTO_SET_MISSING_BRANCHES=AUTO_SET_MISSING_BRANCHES)
        if br is not None:
            parent_branch = br
            continue

        # Branch is missing: decide what type it should be
        parent_path = "\\".join(parts[:i]) if i > 0 else ""
        # Allow user to override branch type selection
        # If branch_tuple is a dict with 'branch_type' key, use that
        # Otherwise fall back to automatic inference
        if isinstance(branch_tuple, dict) and 'branch_type' in branch_tuple:
            branch_type = branch_tuple['branch_type']
        else:
            branch_type = _choose_branch_type_for_segment(
                current_path=parent_path,
                segment_name=part,
                branch_tuple=branch_tuple,
                shortname_to_leap_branches=shortname_to_leap_branches,
            )
        if AUTO_SET_MISSING_BRANCHES:
            # Create the new branch with LEAPApplication methods
            new_branch = _create_child_branch(L, parent_branch, part, branch_type)
        else:
            breakpoint()#not sure how this will behave
            new_branch = None
        parent_branch = new_branch

    return parent_branch

def _create_child_branch(L, parent_branch, name, branch_type):
    """
    Create a new LEAP branch under parent_branch, using LEAPApplication
    methods (AddCategory, AddTechnology, etc.).

    NOTE:
    - LEAP has no AddDemandFuel API. Demand fuel branches (type 36) are
      created implicitly when you create technologies with a fuel.
    """
    
    if parent_branch is None:
        breakpoint()
        raise RuntimeError(
            f"Cannot create top-level branch '{name}' without an existing parent. "
            "In practice, roots like 'Demand' must already exist."
        )

    # Get the parent ID from the branch
    parent_id = parent_branch.ID  # COM property: Branch.ID

    # Category: use AddCategory(parent_id, name, Scale, AcUnit)
    if branch_type == BRANCH_DEMAND_CATEGORY:
        # Use blank defaults for scale and activity unit; user can edit later.
        # AddCategory(ParentID, BName, Scale, AcUnit) :contentReference[oaicite:2]{index=2}
        return L.AddCategory(parent_id, name, "", "")

    # Technology (Activity method): use AddTechnology(...)
    if branch_type == BRANCH_DEMAND_TECHNOLOGY:
        # AddTechnology(ParentID, BName, Scale, AcUnit, Fuel, EnergyUnit) :contentReference[oaicite:3]{index=3}
        # We don't know the actual defaults from here, so use empty strings. The user will need to set them manually... they may also get set by the imported data.
        
        # and let the user fill in fuel & units in LEAP later.
        #AddTechnology(ParentID, BName, Scale, AcUnit, Fuel, EnergyUnit)
        print(f"Creating technology branch '{name}' under parent ID {parent_id}. Remember to set units manually in LEAP.")
        return L.AddTechnology(parent_id, name, "", "", name, "")

    # Demand fuel branches: LEAP exposes BranchType=36 but no AddDemandFuel.
    # These are normally created when you define a technology with an
    # associated fuel, not directly via API.
    if branch_type == BRANCH_DEMAND_FUEL:
        breakpoint()
        raise RuntimeError(
            f"Cannot auto-create demand fuel branch '{name}': LEAP API "
            "does not expose an AddDemandFuel method. Create the associated "
            "technology (with its fuel) in LEAP, or handle this branch manually."
        )

    raise RuntimeError(f"Unsupported branch_type={branch_type} for '{name}'.")



# ------------------------------------------------------------
def diagnose_measures_in_leap_branch(L, branch_path, leap_tuple, expected_vars=None, verbose=False):
    """Diagnose variables available in a LEAP branch."""
    branch = safe_branch_call(L, branch_path)
    if branch is None:
        print(f"[ERROR] Could not access branch {branch_path}")
        print("=" * 50)
        return

    try:
        var_count = branch.Variables.Count
        available_vars = [branch.Variables.Item(i + 1).Name for i in range(var_count)]

        if verbose:
            print(f"\n=== Diagnosing Branch: {leap_tuple} ===")
            print(f"Available variables: {sorted(available_vars)}")

        if expected_vars:
            missing = set(expected_vars) - set(available_vars)
            if missing:
                print(f"Missing expected variables: {sorted(missing)}")

    except Exception as e:
        print(f"[ERROR] Could not enumerate variables in '{branch_path}': {e}")

    print("=" * 50)
    return


# ------------------------------------------------------------
# Transformation & Resource helpers
# ------------------------------------------------------------

def _resolve_branch_reference(L, branch_reference, description="branch"):
    """Return a LEAP branch object given a path or existing branch."""
    if isinstance(branch_reference, str):
        branch = safe_branch_call(
            L, branch_reference, AUTO_SET_MISSING_BRANCHES=False, THROW_ERROR_ON_MISSING=False
        )
        if branch is None:
            raise RuntimeError(f"Could not locate {description} at '{branch_reference}'.")
        return branch
    if branch_reference is None:
        raise ValueError(f"{description.capitalize()} reference cannot be None.")
    return branch_reference


def create_transformation_module(
    L,
    parent_branch,
    module_name,
    *,
    is_simple=True,
    use_efficiencies=True,
    use_capacities=False,
    use_load_curve=False,
    use_co_prod=False,
    use_output_shares=False,
    meet_aux_from_outputs=True,
    coproduct_fuel=None,
    output_scale="",
    output_unit="",
    capacity_scale="",
    capacity_unit="",
):
    """Create a transformation module (LEAP.AddModule) beneath the given parent."""
    try:
        parent = _resolve_branch_reference(L, parent_branch, description="module parent")
        return L.AddModule(
            module_name,
            bool(is_simple),
            bool(use_efficiencies),
            bool(use_capacities),
            bool(use_load_curve),
            bool(use_co_prod),
            bool(use_output_shares),
            bool(meet_aux_from_outputs),
            coproduct_fuel or "",
            output_scale,
            output_unit,
            capacity_scale,
            capacity_unit,
        )
    except Exception as exc:
        print(f"[ERROR] Failed to create transformation module '{module_name}': {exc}")
        breakpoint()
        raise


def create_transformation_process(
    L,
    parent_branch,
    process_name,
    feedstock_fuel="",
    dispatch_rule=0,
):
    """Create a transformation process branch (LEAP.AddProcess) under a process category."""
    try:
        parent = _resolve_branch_reference(L, parent_branch, description="process parent")
        sanitized_feedstock = sanitize_leap_name(feedstock_fuel) if feedstock_fuel else ""
        if sanitized_feedstock:
            ensure_fuel_exists(L, sanitized_feedstock)
        return L.AddProcess(parent.ID, process_name, sanitized_feedstock or "", int(dispatch_rule))
    except Exception as exc:
        print(f"[ERROR] Failed to create transformation process '{process_name}': {exc}")
        breakpoint()
        raise


def create_transformation_output(
    L,
    parent_branch,
    fuel_name,
    shortfall_import=0,
    surplus_export=0,
    domestic_priority=0,
    is_priority_fuel=False,
):
    """Attach an output fuel branch to a transformation process."""
    try:
        parent = _resolve_branch_reference(L, parent_branch, description="output parent")
        sanitized_fuel = sanitize_leap_name(fuel_name)
        ensure_fuel_exists(L, sanitized_fuel)
        return L.AddOutput(
            parent.ID,
            sanitized_fuel,
            int(shortfall_import),
            int(surplus_export),
            int(domestic_priority),
            bool(is_priority_fuel),
        )
    except Exception as exc:
        print(f"[ERROR] Failed to add transformation output '{fuel_name}': {exc}")
        breakpoint()
        raise


def create_transformation_feedstock(L, parent_branch, fuel_name):
    """Attach a feedstock fuel branch under the specified transformation process."""
    try:
        parent = _resolve_branch_reference(L, parent_branch, description="feedstock parent")
        sanitized_fuel = sanitize_leap_name(fuel_name)
        ensure_fuel_exists(L, sanitized_fuel)
        return L.AddFeedstock(parent.ID, sanitized_fuel)
    except Exception as exc:
        print(f"[ERROR] Failed to add transformation feedstock '{fuel_name}': {exc}")
        breakpoint()
        raise


def create_simple_transformation_process(
    L, parent_branch, process_name, input_fuel, output_fuel
):
    """Simplified constructor for single-input single-output transformation processes."""
    try:
        parent = _resolve_branch_reference(L, parent_branch, description="process parent")
        sanitized_input = sanitize_leap_name(input_fuel)
        sanitized_output = sanitize_leap_name(output_fuel)
        ensure_fuel_exists(L, sanitized_input)
        ensure_fuel_exists(L, sanitized_output)
        return L.AddSimpleProcess(
            parent.ID, process_name, sanitized_input, sanitized_output
        )
    except Exception as exc:
        print(f"[ERROR] Failed to create simple process '{process_name}': {exc}")
        breakpoint()
        raise


def get_resource_branch_for_fuel(L, fuel_name):
    """Returns the supply branch assigned to the fuel (Resources → Primary/Secondary)."""
    try:
        sanitized_fuel = sanitize_leap_name(fuel_name)
        branch = L.ResourceBranchFromFuel(sanitized_fuel)
        if branch is None:
            raise RuntimeError(
                f"No resource branch found for fuel '{fuel_name}'."
            )
        return branch
    except Exception as exc:
        print(f"[ERROR] Failed to retrieve resource branch for '{fuel_name}': {exc}")
        breakpoint()
        raise


def ensure_fuel_exists(L, fuel_name, copy_from=None, fuel_state=2):
    """Create a new fuel entry if one does not already exist."""
    sanitized_fuel = sanitize_leap_name(fuel_name)
    if not sanitized_fuel:
        raise ValueError(f"Fuel name '{fuel_name}' is empty after sanitization.")
    sanitized_copy_from = sanitize_leap_name(copy_from) if copy_from else ""
    try:
        fuels = L.Fuels
        if fuels.Exists(sanitized_fuel):
            return fuels.Item(sanitized_fuel)
        # breakpoint()
        return fuels.Add(sanitized_fuel, sanitized_copy_from or "", int(fuel_state))
    except Exception as exc:
        print(f"[ERROR] Could not create fuel '{fuel_name}': {exc}")
        breakpoint()
        raise


def ensure_unit_exists(L, unit_name):
    """Return a Unit object, raising if it does not exist."""
    try:
        units = L.Units
        if units.Exists(unit_name):
            return units.Item(unit_name)
        raise ValueError(f"Unit not found in LEAP: {unit_name}")
    except Exception as exc:
        print(f"[ERROR] Could not load unit '{unit_name}': {exc}")
        breakpoint()
        raise


# ------------------------------------------------------------
# Branch creation from an export spreadsheet
# ------------------------------------------------------------

def identify_branch_type_from_mapping(bp, other_branch_paths, branch_root, branch_type_mapping, default_branch_type):
    branch_tuple = tuple(bp.split('\\'))
    #if the root branch type is provided in the mapping then create a version of teh branch tuplewhich does not include the root branch
    if branch_root is not None:
        branch_root_tuple = tuple(branch_root.split('\\'))
        branch_tuple_no_root = branch_tuple[len(branch_root_tuple):]
        
    #test if we can find the branch type directly from the mapping
    branch_type = branch_type_mapping.get(branch_tuple)
    if branch_type is None and branch_root is not None:
        branch_type = branch_type_mapping.get(branch_tuple_no_root)
    
    if branch_type is not None:
        pass#we have identified the branch type to use
    else:
        #identify the branch type to use. have to do this by finding all branches that contain this branch path... if there are other branches with this path then we need to identify if this is the last, 2nd to last or other segment in the path
        
        #find branch paths in branch_paths_copy
        matching_branch_paths = [b for b in other_branch_paths if b.startswith(bp)]
        branch_paths_with_one_more_segment = [b for b in matching_branch_paths if len(b.split("\\")) - len(bp.split("\\")) == 1]
        branch_paths_with_two_more_segments = [b for b in matching_branch_paths if len(b.split("\\")) - len(bp.split("\\")) == 2]
        if len(branch_paths_with_two_more_segments)>0:
            branch_type = default_branch_type[0]#not last or 2nd to last segment
        elif len(branch_paths_with_one_more_segment)>0:
            branch_type = default_branch_type[1]#2nd to last segment
        else:
            branch_type = default_branch_type[2]#last segment
    
    return branch_type


def _find_first_feedstock_for_process(process_path, branch_paths):
    """Return the first feedstock fuel name discovered under a process path."""
    if not branch_paths:
        return ""
    feedstock_prefix = process_path + "\\Feedstock Fuels\\"
    for path in branch_paths:
        if path.startswith(feedstock_prefix):
            parts = [p for p in path.split("\\") if p]
            if parts:
                return parts[-1]
    return ""


def _normalize_label(raw: str) -> str:
    """Normalize branch labels for comparison (replace '-' with space)."""
    if raw is None:
        return ""
    return raw.replace("-", " ").strip().lower()


_LEAP_NAME_SANITIZE_RE = re.compile(r"[^A-Za-z0-9\s]+")
_LEAP_AND_REPLACEMENTS = {
    "&": " and ",
    "/": " and ",
    "-": " ",
}


def sanitize_leap_name(raw: str | None) -> str:
    """
    Normalize names before sending them to LEAP to avoid unsupported characters.

    Args:
        raw: The user-provided label.

    Returns:
        A version of the label where ``&`` is replaced with `` and ``, ``-`` becomes
        `` and ``, ``/`` becomes `` and ``, and other non-alphanumeric characters are stripped.
    """
    if not raw:
        return ""
    normalized = str(raw)
    for target, replacement in _LEAP_AND_REPLACEMENTS.items():
        normalized = normalized.replace(target, replacement)
    normalized = _LEAP_NAME_SANITIZE_RE.sub(" ", normalized)
    return " ".join(normalized.split()).strip()


def sanitize_leap_branch_path(raw: str | None) -> str:
    """Return a branch path with each segment sanitized for LEAP compatibility."""
    if not raw:
        return ""
    segments = [segment.strip() for segment in str(raw).split("\\") if segment.strip()]
    sanitized_segments = [sanitize_leap_name(segment) for segment in segments]
    sanitized_segments = [segment for segment in sanitized_segments if segment]
    return "\\".join(sanitized_segments)


def _collect_transformation_category_paths(branch_paths):
    """Return transformation-specific category paths (process/output/feedstock/aux)."""
    category_types = {
        BRANCH_PROCESS_CATEGORY,
        BRANCH_OUTPUT_CATEGORY,
        BRANCH_FEEDSTOCK_CATEGORY,
        BRANCH_AUX_CATEGORY,
    }
    categories = set()
    for path in branch_paths:
        parts = [p for p in path.split("\\") if p]
        if not parts:
            continue
        branch_type = _guess_branch_type_for_segment(parts, len(parts) - 1)
        if branch_type in category_types:
            categories.add(path)
    return categories


def _guess_branch_type_for_segment(parts, index):
    """Guess special branch types for transformation and resources based on path parts."""
    if not parts or index >= len(parts):
        return None
    root = _normalize_label(parts[0])
    name = parts[index]
    name_lower = _normalize_label(name)
    parent = _normalize_label(parts[index - 1]) if index > 0 else ""
    if root == "transformation":
        if name_lower == "output fuels":
            return BRANCH_OUTPUT_CATEGORY
        if parent == "output fuels":
            return BRANCH_OUTPUT
        if name_lower == "processes":
            return BRANCH_PROCESS_CATEGORY
        if parent == "processes":
            return BRANCH_TRANSFORMATION_PROCESS
        if name_lower == "auxiliary fuels":
            return BRANCH_AUX_CATEGORY
        if parent == "auxiliary fuels":
            return BRANCH_AUX_BRANCH
        if name_lower == "feedstock fuels":
            return BRANCH_FEEDSTOCK_CATEGORY
        if parent == "feedstock fuels":
            return BRANCH_FEEDSTOCK_BRANCH
        if index == 1:
            # Transformation/<Sector> must be process category level
            return BRANCH_PROCESS_CATEGORY

    if root == "resources":
        if index == 0:
            return BRANCH_RESOURCE_ROOT
        if index == 1:
            if name_lower == "primary":
                return BRANCH_RESOURCE_PRIMARY_CATEGORY
            if name_lower == "secondary":
                return BRANCH_RESOURCE_SECONDARY_CATEGORY
            # Anything directly under Resources is a resource category
            return BRANCH_RESOURCE_BRANCH
        if index >= 2:
            return BRANCH_RESOURCE_BRANCH

    return None


def build_branch_type_mapping_from_paths(branch_paths):
    """
    Build a fallback branch_type_mapping by inferring branch_type from path segments.

    This mirrors the structure exposed in `leap_utils/transformation_analysis_utils.py`
    where processes, outputs, feedstocks and auxiliaries live under the Transformation root.
    """
    mapping = {}
    for path in branch_paths:
        parts = [p for p in path.split("\\") if p]
        if not parts:
            continue
        branch_type = _guess_branch_type_for_segment(parts, len(parts) - 1)
        if branch_type:
            mapping[path] = branch_type
    return mapping


def _ensure_path_exists_create_if_not(
    L,
    full_path,
    branch_root,
    other_branch_paths,
    branch_type_mapping,
    default_branch_type,
    missing_process_categories,
    SCALE=1,
    UNIT="PJ",
):
    """
    NOTE THAT THIS FUCTION HAS BEEN BUILT TO WORK INDEPENDTLY OF THE TRANSPORT BASED SYSTEM.
    Create a chain of key assumption style categories (just one number, no inference of the kind of category it is, e.g. technology/fuel/stock/intensity style branches)."""
    parts = [p for p in full_path.split("\\") if p]
    parent_branch = None
    for i, part in enumerate(parts):
        current_path = "\\".join(parts[:i + 1])
        br = safe_branch_call(L, current_path, AUTO_SET_MISSING_BRANCHES=False, THROW_ERROR_ON_MISSING=False)
        if br is not None:
            parent_branch = br
            continue
        if parent_branch is None:
            print(f"[WARN] Cannot create '{current_path}' because its parent is missing. Ensure root branches exist.")
            return None
        
        branch_type = identify_branch_type_from_mapping(
            current_path, other_branch_paths, branch_root, branch_type_mapping, default_branch_type
        )
        guessed = _guess_branch_type_for_segment(parts[:i + 1], i)
        branch_type = guessed or branch_type

        if branch_type == BRANCH_DEMAND_CATEGORY:
            parent_branch = L.AddCategory(parent_branch.ID, part, "", "")
        elif branch_type == BRANCH_DEMAND_TECHNOLOGY:
            print(f"[INFO] Creating technology branch '{part}' under parent ID {parent_branch.ID}. Remember to set units manually in LEAP.")
            parent_branch = L.AddTechnology(parent_branch.ID, part, "", "", part, "")
        elif branch_type == BRANCH_KEY_ASSUMPTION_BRANCH:
            parent_branch = L.AddKeyAssumption(parent_branch.ID, part, SCALE, UNIT)
        elif branch_type == BRANCH_KEY_ASSUMPTION_CATEGORY:
            parent_branch = L.AddKeyAssumptionCategory(parent_branch.ID, part)
        elif branch_type in (
            BRANCH_PROCESS_CATEGORY,
            BRANCH_OUTPUT_CATEGORY,
            BRANCH_FEEDSTOCK_CATEGORY,
            BRANCH_AUX_CATEGORY,
        ):
            warning = (
                f"Transformation category '{current_path}' must exist in LEAP "
                "before child branches can be created."
            )
            print(f"[WARN] {warning}")
            _prompt_on_missing_branch_creation(warning)
            if missing_process_categories is not None:
                missing_process_categories.add(current_path)
            return None
        elif branch_type in (
            BRANCH_TRANSFORMATION_MODULE,
            BRANCH_RESOURCE_ROOT,
            BRANCH_RESOURCE_PRIMARY_CATEGORY,
            BRANCH_RESOURCE_SECONDARY_CATEGORY,
        ):
            parent_branch = L.AddCategory(parent_branch.ID, part, "", "")
        elif branch_type == BRANCH_TRANSFORMATION_PROCESS:
            try:
                parent_branch = create_transformation_process(
                    L,
                    parent_branch,
                    part,
                    feedstock_fuel=_find_first_feedstock_for_process(
                        current_path, other_branch_paths
                    ),
                )
            except Exception:
                warning = (
                    f"Transformation process '{current_path}' could not be created via the API. "
                    "Please add it manually so its child branches can be filled."
                )
                print(f"[WARN] {warning}")
                _prompt_on_missing_branch_creation(warning)
                if missing_process_categories is not None:
                    missing_process_categories.add(current_path)
                return None
        elif branch_type == BRANCH_OUTPUT:
            parent_branch = create_transformation_output(L, parent_branch, part)
        elif branch_type == BRANCH_FEEDSTOCK_BRANCH:
            parent_branch = create_transformation_feedstock(L, parent_branch, part)
        elif branch_type == BRANCH_AUX_BRANCH:
            sanitized_aux = sanitize_leap_name(part)
            ensure_fuel_exists(L, sanitized_aux)
            unit_obj = ensure_unit_exists(L, "Gigajoule")
            parent_branch = L.AddAuxiliary(
                parent_branch.ID, sanitized_aux, unit_obj, unit_obj, 1
            )
        elif branch_type == BRANCH_DEMAND_FUEL:
            breakpoint()
            raise RuntimeError(
                f"Cannot auto-create demand fuel branch '{current_path}': LEAP API "
                "does not expose an AddDemandFuel method. Create the associated "
                "technology (with its fuel) in LEAP, or handle this branch manually."
            )
        else:
            print(f"[WARN] Unsupported branch_type {branch_type} for '{current_path}'. Skipping creation.")
            return None
        if parent_branch is None:
            breakpoint()
            print(f"[WARN] Failed to create branch at '{parent_branch}'.")
    return parent_branch

def create_branches_from_export_file(
    L,
    leap_export_filename,
    sheet_name="LEAP",
    branch_path_col="Branch Path",
    scenario=None,
    region=None,
    branch_root=None, 
    branch_type_mapping=None,
    default_branch_type=(BRANCH_KEY_ASSUMPTION_CATEGORY, BRANCH_KEY_ASSUMPTION_CATEGORY, BRANCH_KEY_ASSUMPTION_BRANCH),
    RAISE_ERROR_ON_FAILED_BRANCH_CREATION=True,
):
    """
    NOTE THAT THIS FUNCTION HAS BEEN BUILT TO WORK INDEPENDENTLY OF THE TRANSPORT BASED SYSTEM.
    Create LEAP branches listed in an export/import spreadsheet.

    Parameters:
    -----------
    L : LEAP application object
        Connected LEAP instance
    leap_export_filename : str
        Path to Excel file containing branch paths
    sheet_name : str
        Sheet name to read from (default 'LEAP')
    branch_path_col : str
        Column name containing branch paths (default 'Branch Path')
    scenario : str, optional
        Filter by scenario if column exists
    region : str, optional
        Filter by region if column exists
    branch_root : str, optional
        Root path which prepends branches paths. Example: "Key Assumptions/Energy Balances". This is not used in branch creation but can be used in branch_type_mapping, for exmaple in the Energy Balances context you dont want to look up mappings for "Key Assumptions/Energy Balances/XYZ", instead you jsut want to look up "XYZ"
    branch_type_mapping : dict, optional
        Maps branch paths to specific branch types. Example:
        {"Key\\Population": BRANCH_KEY_ASSUMPTION_BRANCH}
    default_branch_type : tuple
        Three-element tuple (non_leaf, second_to_leaf, leaf) defining branch types
        for different positions in the path hierarchy.bused when branch_type_mapping does not provide a type.
        Default: (BRANCH_KEY_ASSUMPTION_CATEGORY, BRANCH_KEY_ASSUMPTION_CATEGORY, BRANCH_KEY_ASSUMPTION_BRANCH)
    RAISE_ERROR_ON_FAILED_BRANCH_CREATION : bool
        If True, raises error when branch creation fails. If False, logs warning.
        
    Returns:
    --------
    dict with keys 'created', 'skipped', 'failed' containing lists of branch paths
    
    Notes:
    ------
    - Reads branch paths from Excel and creates missing branches hierarchically
    - Supports both single header (row 0) and double header (row 2) formats
    - default_branch_type uses position-based logic:
        * First element: for branches with 2+ children below them
        * Second element: for branches with exactly 1 child below them
        * Third element: for leaf branches (no children)
        > this arg is used when branch_type_mapping does not provide a type for a given path. this is currently only used for demand branches. in future it would be good to shift to only using branch_type_mapping for all branch types but its a bit difficult to do that right now without breaking existing functionality.
    - branch_type_mapping overrides default_branch_type for specific paths. set in create_branches_from_export_file 
    """
    if L is None:
        raise RuntimeError("LEAP application instance (L) is required to create branches.")

    def _read_sheet(path, header_guess):
        try:
            return pd.read_excel(path, sheet_name=sheet_name, header=header_guess)
        except Exception as e:
            print(f"[WARN] Failed reading sheet '{sheet_name}' with header={header_guess}: {e}")
            return None

    df = _read_sheet(leap_export_filename, header_guess=0)
    if df is None or branch_path_col not in df.columns:
        df = _read_sheet(leap_export_filename, header_guess=2)
    if df is None or branch_path_col not in df.columns:
        raise ValueError(f"Column '{branch_path_col}' not found in {leap_export_filename} (sheet '{sheet_name}').")

    if scenario is not None and "Scenario" in df.columns:
        df = df[df["Scenario"] == scenario]
        if len(df) == 0:
            breakpoint()
            raise ValueError(f"No rows found for scenario '{scenario}' in {leap_export_filename} (sheet '{sheet_name}').")
    if region is not None and "Region" in df.columns:
        df = df[df["Region"] == region]
        if len(df) == 0:
            breakpoint()
            raise ValueError(f"No rows found for region '{region}' in {leap_export_filename} (sheet '{sheet_name}').")

    branch_paths_raw = [bp for bp in df[branch_path_col].dropna().unique() if isinstance(bp, str)]
    branch_paths = []
    seen = set()
    for bp in branch_paths_raw:
        sanitized_bp = sanitize_leap_branch_path(bp)
        if not sanitized_bp or sanitized_bp in seen:
            continue
        seen.add(sanitized_bp)
        branch_paths.append(sanitized_bp)
    branch_paths = sorted(branch_paths, key=lambda x: len(x.split("\\")))

    created = []
    skipped = []
    failed = []
    branch_type_mapping = branch_type_mapping or {}#if we were provided a branchtype mapping then the branch types will be inferred from that where possible
    branch_paths_copy = branch_paths.copy()
    inferred_mapping = build_branch_type_mapping_from_paths(branch_paths)
    for path, bt in inferred_mapping.items():
        branch_type_mapping.setdefault(path, bt)
    missing_process_categories = set()
    transformation_categories = _collect_transformation_category_paths(branch_paths)
    for category in transformation_categories:
        if safe_branch_call(
            L, category, AUTO_SET_MISSING_BRANCHES=False, THROW_ERROR_ON_MISSING=False
        ) is None:
            missing_process_categories.add(category)
    # breakpoint()#investiage how to handle tranformation process categories sych as Transformation\Transfers\Processes\Upstream liquids transfers
    for bp in branch_paths:
        if safe_branch_call(L, bp, AUTO_SET_MISSING_BRANCHES=False, THROW_ERROR_ON_MISSING=False) is not None:
            skipped.append(bp)
            continue
                
        node = _ensure_path_exists_create_if_not(
            L,
            bp,
            branch_root,
            branch_paths_copy,
            branch_type_mapping,
            default_branch_type,
            missing_process_categories,
        )
        
        if node:
            created.append(bp)
            continue
        if missing_process_categories and any(
            bp.startswith(f"{cat}\\") or bp == cat for cat in missing_process_categories
        ):
            skipped.append(bp)
            continue
        if RAISE_ERROR_ON_FAILED_BRANCH_CREATION:
            breakpoint()
            raise RuntimeError(f"Failed to create branch at '{bp}'.")
        failed.append(bp)
        print(f"[WARN] Failed to create branch at '{bp}'.")

    print(f"[INFO] Branch creation complete. Created {len(created)}, skipped existing {len(skipped)}.")
    if missing_process_categories:
        print(
            "[WARN] The following Transformation process categories need to be "
            "created manually in LEAP before importing child data:"
        )
        for category in sorted(missing_process_categories):
            print(f"  - {category}")
    return {"created": created, "skipped": skipped, "failed": failed}


def fill_branches_from_export_file(
    L,
    leap_export_filename,
    sheet_name="LEAP",
    scenario=None,
    region=None,
    RAISE_ERROR_ON_FAILED_SET=True,
    SET_UNITS=True,
    HANDLE_CURRENT_ACCOUNTS_TOO=False,
    # SET_SCALE=True,
):
    """
    NOTE THAT THIS FUCTION HAS BEEN BUILT TO WORK INDEPENDTLY OF THE TRANSPORT BASED SYSTEM.
    Fill LEAP branch variables with data from an export/import spreadsheet.
    
    - Reads data from Excel file (default sheet 'LEAP')
    - Expects LEAP import format with branch paths, variables, and year columns
    - Sets expressions on branches using Data() interpolation
    - Handles both single header (row 0) and double header (row 2) formats
    
    Parameters:
    -----------
    L : LEAP application object
    leap_export_filename : str
        Path to Excel file containing LEAP data
    sheet_name : str
        Sheet name to read from (default 'LEAP')
    branch_path_col : str
        Column name for branch paths
    variable_col : str
        Column name for variable names
    scenario : str, optional
        Filter by scenario if column exists
    region : str, optional
        Filter by region if column exists
    HANDLE_CURRENT_ACCOUNTS_TOO : bool
        If True and a scenario is provided, also process "Current Accounts"
    RAISE_ERROR_ON_FAILED_SET : bool
        Whether to raise error if setting a variable fails
        
    Returns:
    --------
    dict with keys 'success', 'failed' containing lists of (branch_path, variable) tuples
    """
    if L is None:
        raise RuntimeError("LEAP application instance (L) is required to fill branches.")

    def _read_sheet(path, header_guess):
        try:
            return pd.read_excel(path, sheet_name=sheet_name, header=header_guess)
        except Exception as e:
            print(f"[WARN] Failed reading sheet '{sheet_name}' with header={header_guess}: {e}")
            return None

    # Try reading with different header rows
    df = _read_sheet(leap_export_filename, header_guess=0)
    if df is None or "Branch Path" not in df.columns:
        df = _read_sheet(leap_export_filename, header_guess=2)
    if df is None or "Branch Path" not in df.columns:
        raise ValueError(f"Columns 'Branch Path' or 'Variable' not found in {leap_export_filename} (sheet '{sheet_name}').")

    def _fill_from_df(df_in):
        # Filter by scenario/region if specified
        if scenario is not None and "Scenario" in df_in.columns:
            df_in = df_in[df_in["Scenario"] == scenario]
            if len(df_in) == 0:
                breakpoint()
                raise ValueError(f"No rows found for scenario '{scenario}' in {leap_export_filename} (sheet '{sheet_name}').")
        if region is not None and "Region" in df_in.columns:
            df_in = df_in[df_in["Region"] == region]
            if len(df_in) == 0:
                breakpoint()
                raise ValueError(f"No rows found for region '{region}' in {leap_export_filename} (sheet '{sheet_name}').")
        
        df_in = df_in.copy()
        df_in["_Sanitized Branch Path"] = df_in["Branch Path"].apply(
            sanitize_leap_branch_path
        )
        df_in = df_in[df_in["_Sanitized Branch Path"] != ""]

        #if the df contains year cols then we use those instead of expression cols.
        if 'Expression' in df_in.columns:
            print(f"[INFO] 'Expression' column found in {leap_export_filename}, using it to set variable expressions directly.")
            year_cols = ['Expression']
        else:
            # Identify year columns (numeric or str columns that have 4 digits)
            year_cols = [col for col in df_in.columns if len(str(col)) == 4 and str(col).isdigit()]
        
        if not year_cols:
            breakpoint()
            raise ValueError(f"No year columns found in {leap_export_filename}")

        success = []
        failed = []
        
        # Group by branch path and variable
        
        for (sanitized_bp, var), group in df_in.groupby(
            ["_Sanitized Branch Path", "Variable"]
        ):
            branch_path_used = sanitized_bp
            branch = safe_branch_call(
                L,
                sanitized_bp,
                AUTO_SET_MISSING_BRANCHES=False,
                THROW_ERROR_ON_MISSING=False,
            )
            if branch is None:
                original_bp = group["Branch Path"].iloc[0]
                branch = safe_branch_call(
                    L,
                    original_bp,
                    AUTO_SET_MISSING_BRANCHES=False,
                    THROW_ERROR_ON_MISSING=False,
                )
                branch_path_used = original_bp
            
            if branch is None:
                original_bp = group["Branch Path"].iloc[0]
                msg = (
                    f"Branch '{original_bp}' (sanitized '{sanitized_bp}') not found - "
                    f"skipping variable '{var}'"
                )
                if RAISE_ERROR_ON_FAILED_SET:
                    breakpoint()
                    raise RuntimeError(msg)
                else:
                    print(f"[WARN] {msg}")
                    failed.append((branch_path_used, var))
                    continue
            if ['Expression'] == year_cols:
                #we just need to set the expression directly
                expr = group['Expression'].iloc[0]
            else:
                # Extract year-value pairs
                points = []
                for year in year_cols:
                    val = group[year].iloc[0]
                    if pd.notna(val):
                        try:
                            points.append((int(year), float(val)))
                        except (ValueError, TypeError):
                            print(
                                f"[WARN] Invalid value for {branch_path_used}\\{var} "
                                f"in year {year}: {val}"
                            )
                            continue

                if not points:
                    print(
                        f"[WARN] No valid data points for {branch_path_used}\\{var}"
                    )
                    failed.append((branch_path_used, var))
                    continue

                # Build expression
                expr = build_expr(points, expression_type="")
            
            if expr is None:
                if RAISE_ERROR_ON_FAILED_SET:
                    breakpoint()
                    raise RuntimeError(
                        f"Failed to build expression for {branch_path_used}\\{var}"
                    )
                print(
                    f"[WARN] Failed to build expression for {branch_path_used}\\{var}"
                )
                failed.append((branch_path_used, var))
                continue
            
            unit_name = None
            # scale_value = None
            if SET_UNITS:
                unit_name = group['Units'].iloc[0] if 'Units' in group.columns else None
            # if SET_SCALE:#kept this here in case someone wants to try again to insert scale value.. its also kind of proof that it wont work but was considered
            #     #if the scale column exists and the value is not na then we set the scale
            #     scale_value = group['Scale'].iloc[0] if 'Scale' in group.columns else None   
            #     if pd.isna(scale_value):
            #         scale_value = None           
            # Set the variable
            # if var == "Process Efficiency":
            #     breakpoint()#trying to track down cause of [ERROR] Failed setting Process Efficiency on Transformation\Patent fuel plants\Processes\Patent fuel plants: (-2147352571, 'Type mismatch.', None, 1)
            set_success = safe_set_variable(
                L, branch, var, expr, unit_name=unit_name, context=branch_path_used
            )
            
            if set_success:
                success.append((branch_path_used, var))
            else:
                # breakpoint()
                if RAISE_ERROR_ON_FAILED_SET:
                    breakpoint()
                    raise RuntimeError(
                        f"Failed to set variable '{var}' on branch '{branch_path_used}'"
                    )
                failed.append((branch_path_used, var))

        print(f"[INFO] Data fill complete. Success: {len(success)}, Failed: {len(failed)}")
        return {"success": success, "failed": failed}

    results = []
    scenarios_to_run = [scenario]
    if HANDLE_CURRENT_ACCOUNTS_TOO and scenario and scenario != "Current Accounts":
        scenarios_to_run.append("Current Accounts")

    for scen in scenarios_to_run:
        if scen is not None:
            print(f"[INFO] Filling data for scenario '{scen}'.")
        prev_scenario = scenario
        scenario = scen
        results.append(_fill_from_df(df.copy()))
        scenario = prev_scenario

    combined = {"success": [], "failed": []}
    for res in results:
        combined["success"].extend(res["success"])
        combined["failed"].extend(res["failed"])
    return combined
