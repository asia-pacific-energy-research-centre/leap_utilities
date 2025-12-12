# ============================================================
# LEAP_core.py
# ============================================================
# Core helper functions for LEAP transport data integration.
# Provides connection, diagnostics, normalization, logging,
# and activity level utilities shared by loader scripts.
# ============================================================

import pandas as pd
from win32com.client import Dispatch, GetActiveObject, gencache

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


def _require_global(name: str, val):
    if val is None:
        raise ImportError(
            f"{name} not available; pass it explicitly to this function or install the transport mappings."
        )
    return val

# ------------------------------------------------------------
# Connection & Core Helpers
# ------------------------------------------------------------

def connect_to_leap():
    """Enhanced LEAP connection with project readiness checks."""
    print("[INFO] Connecting to LEAP...")
    
    try:
        # Clear win32com cache to fix corrupted type library
        import shutil
        import tempfile
        gen_py_path = gencache.GetGeneratePath()
        if gen_py_path:
            try:
                shutil.rmtree(gen_py_path)
                print("[INFO] Cleared win32com cache")
            except Exception as e:
                print(f"[WARN] Could not clear cache: {e}")
        
        gencache.EnsureDispatch("LEAP.LEAPApplication")
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
        short_expr = expr[:80] + ("..." if len(expr) > 80 else "")
        print(f"[SET] {context} → {varname} = {short_expr}")
        
        ########
        
        # Set units if provided
        if unit_name is None:
            return True
        units = L.Units
        if units.Exists(unit_name):
            unit = units.Item(unit_name)      # returns ILEAPUnit
            var.DataUnit = unit               # or: var.DataUnitID = unit.ID
            # optional: var.Scale = <scale_id>  # if you also need to set the scale (config lines 1299-1302)
        else:
            breakpoint()
            THROW_ERROR_ON_MISSING=True
            if THROW_ERROR_ON_MISSING:
                raise ValueError(f"Unit not found: {unit_name}")
            else:
                print(f"[WARN] Unit not found: {unit_name}. Proceeding without setting unit.")
                
        ########
        return True
    except Exception as e:
        print(f"[ERROR] Failed setting {varname} on {context}: {e}")
        return False

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


def _ensure_path_exists_create_if_not(L, full_path, branch_root, other_branch_paths, branch_type_mapping, default_branch_type, SCALE=1, UNIT="PJ"):
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
        
        branch_type = identify_branch_type_from_mapping(current_path, other_branch_paths, branch_root, branch_type_mapping, default_branch_type)
        
        if branch_type == BRANCH_DEMAND_CATEGORY:
            parent_branch = L.AddCategory(parent_branch.ID, part, "", "")
        elif branch_type == BRANCH_DEMAND_TECHNOLOGY:
            print(f"[INFO] Creating technology branch '{part}' under parent ID {parent_branch.ID}. Remember to set units manually in LEAP.")
            parent_branch = L.AddTechnology(parent_branch.ID, part, "", "", part, "")
        elif branch_type == BRANCH_DEMAND_FUEL:
            breakpoint()
            raise RuntimeError(
                f"Cannot auto-create demand fuel branch '{current_path}': LEAP API "
                "does not expose an AddDemandFuel method. Create the associated "
                "technology (with its fuel) in LEAP, or handle this branch manually."
            )
        elif branch_type == BRANCH_KEY_ASSUMPTION_BRANCH:
            parent_branch = L.AddKeyAssumption(parent_branch.ID, part, SCALE, UNIT)
        elif branch_type == BRANCH_KEY_ASSUMPTION_CATEGORY:
            parent_branch = L.AddKeyAssumptionCategory(parent_branch.ID, part)
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
    branch_type_mapping : dict, optional
        Maps branch paths to specific branch types. Example:
        {"Key\\Population": BRANCH_KEY_ASSUMPTION_BRANCH}
    default_branch_type : tuple
        Three-element tuple (non_leaf, second_to_leaf, leaf) defining branch types
        for different positions in the path hierarchy.
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
    - branch_type_mapping overrides default_branch_type for specific paths
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
    if region is not None and "Region" in df.columns:
        df = df[df["Region"] == region]

    branch_paths = [bp for bp in df[branch_path_col].dropna().unique() if isinstance(bp, str)]
    branch_paths = sorted(branch_paths, key=lambda x: len(x.split("\\")))

    created = []
    skipped = []
    failed = []
    branch_type_mapping = branch_type_mapping or {}#if we were provided a branchtype mapping then the branch types will be inferred from that where possible
    branch_paths_copy = branch_paths.copy()
    for bp in branch_paths:
        if safe_branch_call(L, bp, AUTO_SET_MISSING_BRANCHES=False, THROW_ERROR_ON_MISSING=False) is not None:
            skipped.append(bp)
            continue
                
        node = _ensure_path_exists_create_if_not(L, bp, branch_root, branch_paths_copy, branch_type_mapping, default_branch_type)
        
        if node:
            created.append(bp)
        else:
            if RAISE_ERROR_ON_FAILED_BRANCH_CREATION:
                breakpoint()
                raise RuntimeError(f"Failed to create branch at '{bp}'.")
            else:
                failed.append(bp)
                print(f"[WARN] Failed to create branch at '{bp}'.")

    print(f"[INFO] Branch creation complete. Created {len(created)}, skipped existing {len(skipped)}.")
    return {"created": created, "skipped": skipped, "failed": failed}


def fill_branches_from_export_file(
    L,
    leap_export_filename,
    sheet_name="LEAP",
    scenario=None,
    region=None,
    RAISE_ERROR_ON_FAILED_SET=True,
    SET_UNITS=True,
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

    # Filter by scenario/region if specified
    if scenario is not None and "Scenario" in df.columns:
        df = df[df["Scenario"] == scenario]
    if region is not None and "Region" in df.columns:
        df = df[df["Region"] == region]

    # Identify year columns (numeric or str columns that have 4 digits)
    year_cols = [col for col in df.columns if len(str(col)) == 4 and str(col).isdigit()]
    
    if not year_cols:
        raise ValueError(f"No year columns found in {leap_export_filename}")

    success = []
    failed = []

    # Group by branch path and variable
    for (bp, var), group in df.groupby(["Branch Path", "Variable"]):
        branch = safe_branch_call(L, bp, AUTO_SET_MISSING_BRANCHES=False, THROW_ERROR_ON_MISSING=False)
        
        if branch is None:
            msg = f"Branch '{bp}' not found - skipping variable '{var}'"
            if RAISE_ERROR_ON_FAILED_SET:
                raise RuntimeError(msg)
            else:
                print(f"[WARN] {msg}")
                failed.append((bp, var))
                continue

        # Extract year-value pairs
        points = []
        for year in year_cols:
            val = group[year].iloc[0]
            if pd.notna(val):
                try:
                    points.append((int(year), float(val)))
                except (ValueError, TypeError):
                    print(f"[WARN] Invalid value for {bp}\\{var} in year {year}: {val}")
                    continue

        if not points:
            print(f"[WARN] No valid data points for {bp}\\{var}")
            failed.append((bp, var))
            continue

        # Build expression
        expr = build_expr(points, expression_type="")
        
        if expr is None:
            if RAISE_ERROR_ON_FAILED_SET:
                breakpoint()
                raise RuntimeError(f"Failed to build expression for {bp}\\{var}")
            print(f"[WARN] Failed to build expression for {bp}\\{var}")
            failed.append((bp, var))
            continue
        
        unit_name = None
        if SET_UNITS:
            unit_name = group['Units'].iloc[0] if 'Units' in group.columns else None
        # Set the variable
        set_success = safe_set_variable(L, branch, var, expr,unit_name=unit_name, context=bp)
        
        if set_success:
            success.append((bp, var))
        else:
            if RAISE_ERROR_ON_FAILED_SET:
                breakpoint()
                raise RuntimeError(f"Failed to set variable '{var}' on branch '{bp}'")
            else:
                failed.append((bp, var))

    print(f"[INFO] Data fill complete. Success: {len(success)}, Failed: {len(failed)}")
    return {"success": success, "failed": failed}
