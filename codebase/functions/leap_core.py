# ============================================================
# LEAP_core.py
# ============================================================
# Core helper functions for LEAP transport data integration.
# Provides connection, diagnostics, normalization, logging,
# and activity level utilities shared by loader scripts.
# ============================================================

import contextlib
import inspect
import io
import math
import os
import re

import pandas as pd
from codebase.utilities import fuel_catalog_preflight
from codebase.functions.analysis_input_write_dispatcher import (
    ensure_analysis_view_api_read_allowed,
    ensure_api_write_allowed,
)
from codebase.functions.leap_api_guard import (
    ensure_leap_api_allowed,
    is_leap_api_allowed,
)
from codebase.functions.leap_session import (
    get_live_pinned_leap_app,
    pin_leap_app,
)

try:  # pragma: no cover - windows-only
    from win32com.client import Dispatch, GetActiveObject, gencache
    _WIN32COM_AVAILABLE = True
except ImportError:  # pragma: no cover - windows-only
    Dispatch = GetActiveObject = gencache = None
    _WIN32COM_AVAILABLE = False

from codebase.configuration.config import (
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
# Debug toggle for AddAuxiliary failures.
# Set to True here, or set env var LEAP_BREAK_ON_ADDAUXILIARY_ERROR=1.
BREAK_ON_ADD_AUXILIARY_ERROR = True
_LEAP_KNOWN_UNIT_NAMES = {
    str(unit.get("name")).strip().lower()
    for unit in LEAP_UNITS_BY_ID.values()
    if unit.get("name")
}
_SKIP_UNIT_ASSIGNMENT_ALIASES = {"percent", "%", "percentage"}
LEAP_IMPORT_LOG_LEVEL_DEFAULT = "detailed"
LEAP_IMPORT_WARNING_PRINT_LIMIT_DEFAULT = 25


def _env_flag(name: str) -> bool:
    value = os.getenv(name, "")
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def _leap_import_log_level() -> str:
    """Return LEAP import log verbosity level."""
    value = os.getenv("LEAP_IMPORT_LOG_LEVEL", LEAP_IMPORT_LOG_LEVEL_DEFAULT)
    level = str(value or "").strip().lower()
    return level if level else LEAP_IMPORT_LOG_LEVEL_DEFAULT


def _is_summary_logging() -> bool:
    """Return True when row-level import logs should be suppressed."""
    return _leap_import_log_level() in {"summary", "quiet", "minimal"}


def _warning_print_limit() -> int:
    """Max warning lines per repeated warning kind."""
    try:
        value = int(
            os.getenv(
                "LEAP_IMPORT_WARNING_PRINT_LIMIT",
                str(LEAP_IMPORT_WARNING_PRINT_LIMIT_DEFAULT),
            )
        )
    except Exception:
        value = LEAP_IMPORT_WARNING_PRINT_LIMIT_DEFAULT
    return max(value, 1)


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
    return _WIN32COM_AVAILABLE and is_leap_api_allowed()

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
    ensure_leap_api_allowed("leap_core.connect_to_leap")
    if not _WIN32COM_AVAILABLE:
        raise RuntimeError(
            "LEAP API (`win32com`) is unavailable in this environment (Linux/WSL). "
            "Run this script from Windows with pywin32 installed to reach LEAP."
        )
    pinned = get_live_pinned_leap_app()
    if pinned is not None:
        print("[INFO] Reusing pinned LEAP instance")
        return pinned
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
        pin_leap_app(leap_app)
        
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
    def _clean_branch_path(path):
        return "\\".join([part.strip() for part in str(path or "").split("\\") if str(part).strip()])

    def _transport_branch_candidates(path):
        cleaned = _clean_branch_path(path)
        if not cleaned:
            return []
        parts = cleaned.split("\\")
        if not parts or parts[0] != "Demand":
            return [cleaned]
        if len(parts) < 2:
            return [cleaned]
        non_road_roots = {
            "Freight non road",
            "Passenger non road",
            "International transport",
            "Pipeline transport",
            "Nonspecified transport",
        }
        if parts[1] == "Transport non road":
            logical_parts = parts[2:]
            if logical_parts and logical_parts[0] in non_road_roots:
                legacy = "\\".join(["Demand", *logical_parts])
                return [cleaned, legacy]
            return [cleaned]
        if parts[1] in non_road_roots:
            canonical = "\\".join(["Demand", "Transport non road", *parts[1:]])
            return [canonical, cleaned]
        return [cleaned]

    ensure_analysis_view_api_read_allowed("leap_core.safe_branch_call")

    if leap_obj is None:
        return None

    branch_candidates = _transport_branch_candidates(branch_path)
    if not branch_candidates:
        return None

    branches = leap_obj.Branches
    for candidate_path in branch_candidates:
        try:
            exists = branches.Exists(candidate_path)
        except Exception as e:
            breakpoint()
            raise Exception(f"Branches.Exists failed for '{candidate_path}': {e}")

        if exists:
            return leap_obj.Branch(candidate_path)

    requested_path = branch_candidates[0]
    if AUTO_SET_MISSING_BRANCHES:
        print(f"[INFO] AUTO_SET_MISSING_BRANCHES is set to true. The branch will be auto-created: {requested_path}")
    elif THROW_ERROR_ON_MISSING:
        breakpoint()
        raise Exception(
            f"Branches.Exists returned false for '{requested_path}'. "
            "AUTO_SET_MISSING_BRANCHES is False and THROW_ERROR_ON_MISSING is true so throwing an error."
        )
    return None
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


def safe_set_variable(
    L,
    obj,
    varname,
    expr,
    unit_name=None,
    context="",
):
    """Safely assign expressions to LEAP variables with logging."""
    ensure_api_write_allowed("leap_core.safe_set_variable")
    try:
        var = obj.Variable(varname)
        if var is None:
            if not _is_summary_logging():
                print(f"[WARN] Missing variable '{varname}' on {context} within LEAP.")
            return False
        var.Expression = expr
        #check that the expression is a string
        short_expr = str(expr)[:80] + ("..." if len(str(expr)) > 80 else "")
        if not _is_summary_logging():
            print(f"[SET] {context} → {varname} = {short_expr}")
        
        try:
            _set_variable_unit(L, var, unit_name, context=context)
        except Exception as exc:
            if not _is_summary_logging():
                print(
                    f"[WARN] Unit assignment skipped for '{varname}' on {context}: {exc}"
                )
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
    normalized_unit_name = str(unit_name).strip()
    if normalized_unit_name.lower() in _SKIP_UNIT_ASSIGNMENT_ALIASES:
        # LEAP does not expose a generic "Percent" unit in all models; keep expression only.
        return True
    units = L.Units
    if not units.Exists(normalized_unit_name):
        raise ValueError(f"Unit not found: {normalized_unit_name}. Unit not set.")
    unit = units.Item(normalized_unit_name)  # returns ILEAPUnit
    if unit is None:
        if normalized_unit_name.lower() not in _LEAP_KNOWN_UNIT_NAMES:
            print(
                f"[WARN] Unit name '{normalized_unit_name}' not found in LEAP units list. Proceeding without setting unit."
            )
            return True
        raise ValueError(
            f"Unit name '{normalized_unit_name}' found in LEAP_UNITS_BY_ID but ILEAPUnit not found in LEAP. Cannot set unit."
        )
    try:
        current_unit = var.DataUnit
    except Exception:
        current_unit = None
    if current_unit is not None:
        try:
            current_name = str(current_unit.Name).strip()
        except Exception:
            current_name = ""
        if current_name and current_name.lower() == normalized_unit_name.lower():
            return True
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
        ensure_fuel_exists(L, name)
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
    def _find_existing_process(parent, target_name):
        target = _normalize_label(target_name)
        if not target:
            return None
        try:
            children = parent.Children
            child_count = int(children.Count)
        except Exception:
            child_count = 0
        for idx in range(1, child_count + 1):
            try:
                child = children.Item(idx)
                child_name = str(child.Name).strip()
            except Exception:
                continue
            if _normalize_label(child_name) == target:
                return child
        try:
            parent_full = str(parent.FullName).strip()
        except Exception:
            parent_full = ""
        if parent_full:
            return safe_branch_call(
                L,
                parent_full + "\\" + target_name,
                AUTO_SET_MISSING_BRANCHES=False,
                THROW_ERROR_ON_MISSING=False,
            )
        return None

    try:
        parent = _resolve_branch_reference(L, parent_branch, description="process parent")
        sanitized_feedstock = sanitize_leap_name(feedstock_fuel) if feedstock_fuel else ""
        if sanitized_feedstock:
            ensure_fuel_exists(L, sanitized_feedstock)
        existing = _find_existing_process(parent, process_name)
        if existing is not None:
            return existing

        # LEAP COM signatures can vary by build; try the common variants.
        attempts = [
            lambda: L.AddProcess(parent.ID, process_name, sanitized_feedstock or "", int(dispatch_rule)),
            lambda: L.AddProcess(parent.ID, process_name, int(dispatch_rule), sanitized_feedstock or ""),
            lambda: L.AddProcess(parent.ID, process_name, int(dispatch_rule)),
            lambda: L.AddProcess(parent.ID, process_name, sanitized_feedstock or ""),
        ]
        errors = []
        for attempt in attempts:
            try:
                created = attempt()
                if created is not None:
                    if sanitized_feedstock:
                        try:
                            create_transformation_feedstock(L, created, sanitized_feedstock)
                        except Exception:
                            pass
                    return created
            except Exception as exc:
                errors.append(str(exc))
                continue

        # One last check in case LEAP created the process despite COM arg ambiguity.
        existing_after = _find_existing_process(parent, process_name)
        if existing_after is not None:
            if sanitized_feedstock:
                try:
                    create_transformation_feedstock(L, existing_after, sanitized_feedstock)
                except Exception:
                    pass
            return existing_after
        raise RuntimeError("; ".join(errors) if errors else "Unknown AddProcess failure")
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
    def _normalize_feedstock_label(text):
        return sanitize_leap_name(text).strip().lower()

    def _child_fuel_label(child):
        """Return the fuel label linked to a feedstock child branch when available."""
        try:
            fuel_obj = child.Fuel
        except Exception:
            return ""
        try:
            name = str(fuel_obj.Name).strip()
            if name:
                return name
        except Exception:
            pass
        try:
            text = str(fuel_obj).strip()
            if text:
                return text
        except Exception:
            pass
        return ""

    def _find_existing_feedstock(parent, target_name):
        target_norm = _normalize_feedstock_label(target_name)
        if not target_norm:
            return None
        feedstock_category = None
        try:
            parent_name = str(parent.Name).strip()
        except Exception:
            parent_name = ""
        if _normalize_label(parent_name) == "feedstock fuels":
            feedstock_category = parent
        try:
            children = parent.Children
            child_count = int(children.Count)
        except Exception:
            child_count = 0
        for idx in range(1, child_count + 1):
            try:
                child = children.Item(idx)
            except Exception:
                continue
            try:
                child_name = str(child.Name).strip()
            except Exception:
                child_name = ""
            if feedstock_category is None and _normalize_label(child_name) == "feedstock fuels":
                feedstock_category = child
                break
        if feedstock_category is None:
            try:
                parent_full = str(parent.FullName).strip()
            except Exception:
                parent_full = ""
            if parent_full:
                feedstock_category = safe_branch_call(
                    L,
                    parent_full + "\\Feedstock Fuels",
                    AUTO_SET_MISSING_BRANCHES=False,
                    THROW_ERROR_ON_MISSING=False,
                )
        if feedstock_category is None:
            return None
        try:
            children = feedstock_category.Children
            child_count = int(children.Count)
        except Exception:
            child_count = 0
        for idx in range(1, child_count + 1):
            try:
                child = children.Item(idx)
            except Exception:
                continue
            try:
                child_name = str(child.Name).strip()
            except Exception:
                child_name = ""
            if not child_name:
                try:
                    full_name = str(child.FullName).strip()
                except Exception:
                    full_name = ""
                if full_name and "\\" in full_name:
                    child_name = full_name.rsplit("\\", 1)[-1].strip()
            if not child_name:
                continue
            if _normalize_feedstock_label(child_name) == target_norm:
                return child
            linked_fuel = _child_fuel_label(child)
            if linked_fuel and _normalize_feedstock_label(linked_fuel) == target_norm:
                return child
        try:
            category_full = str(feedstock_category.FullName).strip()
        except Exception:
            category_full = ""
        if category_full:
            candidate = safe_branch_call(
                L,
                category_full + "\\" + target_name,
                AUTO_SET_MISSING_BRANCHES=False,
                THROW_ERROR_ON_MISSING=False,
            )
            if candidate is not None:
                return candidate
        return None

    try:
        parent = _resolve_branch_reference(L, parent_branch, description="feedstock parent")
        sanitized_fuel = sanitize_leap_name(fuel_name)
        ensure_fuel_exists(L, sanitized_fuel)
        existing = _find_existing_feedstock(parent, sanitized_fuel)
        if existing is not None:
            return existing
        return L.AddFeedstock(parent.ID, sanitized_fuel)
    except Exception as exc:
        error_text = str(exc).lower()
        if "already has a child branch" in error_text or "addfeedstock" in error_text:
            existing = _find_existing_feedstock(parent, sanitized_fuel)
            if existing is not None:
                return existing
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
    ensure_analysis_view_api_read_allowed("leap_core.get_resource_branch_for_fuel")
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
    ensure_api_write_allowed("leap_core.ensure_fuel_exists")
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


def _add_auxiliary_branch(
    L,
    parent_branch,
    fuel_name,
    numerator_unit_name,
    denominator_unit_name,
    method=1,
    current_path="",
):
    """Attach an auxiliary fuel branch using the provided units."""
    try:
        parent = _resolve_branch_reference(L, parent_branch, description="aux parent")
        sanitized_aux = sanitize_leap_name(fuel_name)
        ensure_fuel_exists(L, sanitized_aux)

        unit_name = str(numerator_unit_name).strip() if numerator_unit_name is not None else ""
        per_unit_name = str(denominator_unit_name).strip() if denominator_unit_name is not None else ""
        if not unit_name:
            context = current_path or sanitized_aux
            raise ValueError(
                "Missing Units for auxiliary branch "
                f"'{sanitized_aux}' at '{context}'. "
                "Provide Units (and Per... if different) in the export spreadsheet."
            )
        if not per_unit_name:
            per_unit_name = unit_name

        numerator_unit = ensure_unit_exists(L, unit_name)
        denominator_unit = ensure_unit_exists(L, per_unit_name)
        return L.AddAuxiliary(
            parent.ID,
            sanitized_aux,
            unit_name,
            per_unit_name,
            int(method),
        )
    except Exception as exc:
        print(
            f"[ERROR] Failed to add auxiliary fuel '{fuel_name}': {exc}"
        )
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


def is_placeholder_branch_path(raw: str | None) -> bool:
    """Return True when a branch-path token is placeholder text."""
    text = str(raw or "").strip().lower()
    return text in {"", "none", "nan", "<na>", "null"}


def ensure_analysis_view_context(
    L,
    *,
    context_label: str = "",
    preferred_branches: tuple[str, ...] = ("Transformation", "Demand", "Resources", "Key"),
) -> bool:
    """
    Best-effort switch LEAP to Analysis view and a safe active branch.

    Some LEAP API operations can fail or behave inconsistently when LEAP is in
    Results view; this helper is defensive across LEAP UI/API variants.
    """
    ensure_analysis_view_api_read_allowed("leap_core.ensure_analysis_view_context")
    if L is None:
        return False
    prefix = f"{context_label}: " if str(context_label).strip() else ""
    switched = False

    for attr_name, attr_value in (
        ("ActiveView", "Analysis"),
        ("View", "Analysis"),
        ("Mode", "Analysis"),
    ):
        if not hasattr(L, attr_name):
            continue
        try:
            setattr(L, attr_name, attr_value)
            print(f"[INFO] {prefix}LEAP view set via {attr_name}='{attr_value}'.")
            switched = True
            break
        except Exception:
            continue

    for method_name in (
        "ShowAnalysisView",
        "ShowAnalysisViewTree",
        "ShowAnalysisViewBranches",
        "ShowAnalysisViewTable",
        "ShowAnalysis",
    ):
        method = getattr(L, method_name, None)
        if not callable(method):
            continue
        try:
            method()
            print(f"[INFO] {prefix}LEAP view reset via {method_name}().")
            switched = True
            break
        except Exception:
            continue

    branch_set = False
    for branch_name in preferred_branches:
        try:
            if not bool(L.Branches.Exists(branch_name)):
                continue
            L.ActiveBranch = branch_name
            print(f"[INFO] {prefix}LEAP active branch reset to '{branch_name}'.")
            branch_set = True
            break
        except Exception:
            continue

    if not switched and not branch_set:
        print(
            f"[WARN] {prefix}Could not confirm Analysis-view reset. "
            "If imports fail, switch LEAP UI to Analysis and retry."
        )
    return switched or branch_set


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

    This mirrors the structure exposed in `codebase/transformation_analysis_utils.py`
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
    auxiliary_units_by_path,
    missing_process_categories,
    SCALE=1,
    UNIT="PJ",
):
    """
    NOTE THAT THIS FUCTION HAS BEEN BUILT TO WORK INDEPENDTLY OF THE TRANSPORT BASED SYSTEM.
    Create a chain of key assumption style categories (just one number, no inference of the kind of category it is, e.g. technology/fuel/stock/intensity style branches)."""
    parts = [p for p in full_path.split("\\") if p]
    parent_branch = None
    
    # breakpoint()#why is per not being set
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
            ensure_fuel_exists(L, part)
            parent_branch = L.AddTechnology(parent_branch.ID, part, "", "", part, "")
        elif branch_type == BRANCH_KEY_ASSUMPTION_BRANCH:
            parent_branch = L.AddKeyAssumption(parent_branch.ID, part, SCALE, UNIT)
        elif branch_type == BRANCH_KEY_ASSUMPTION_CATEGORY:
            parent_branch = L.AddKeyAssumptionCategory(parent_branch.ID, part)
        elif branch_type in (
            BRANCH_PROCESS_CATEGORY,
            BRANCH_OUTPUT_CATEGORY,
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
        elif branch_type == BRANCH_FEEDSTOCK_CATEGORY:
            # Feedstock category is typically created implicitly with processes/feedstocks.
            # Keep parent unchanged and continue to create the child feedstock branch.
            continue
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
            
            aux_units = None
            aux_per = None
            if auxiliary_units_by_path:
                aux_entry = auxiliary_units_by_path.get(current_path)
                if aux_entry:
                    aux_units = aux_entry.get("units")
                    aux_per = aux_entry.get("per")
            parent_branch = _add_auxiliary_branch(
                L,
                parent_branch,
                part,
                aux_units,
                aux_per,
                method=1,
                current_path=current_path,
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


def _build_auxiliary_units_by_path(df, branch_path_col):
    """Return a mapping of sanitized branch path -> {units, per} for aux branches."""
    if df is None or df.empty:
        return {}
    units_col = "Units" if "Units" in df.columns else None
    per_col = "Per..." if "Per..." in df.columns else ("Per" if "Per" in df.columns else None)
    if not units_col or not per_col:
        return {}
    var_col = "Variable" if "Variable" in df.columns else ("Measure" if "Measure" in df.columns else None)
    aux_df = df
    if var_col:
        aux_df = df[df[var_col] == "Auxiliary Fuel Use"]
    mapping = {}
    for _, row in aux_df.iterrows():
        raw_path = row.get(branch_path_col)
        if not isinstance(raw_path, str) or not raw_path.strip():
            continue
        sanitized_path = sanitize_leap_branch_path(raw_path)
        if not sanitized_path:
            continue
        units = row.get(units_col)
        per_value = row.get(per_col)
        if pd.isna(units):
            units = None
        if pd.isna(per_value):
            per_value = None
        if isinstance(units, str):
            units = units.strip() or None
        if isinstance(per_value, str):
            per_value = per_value.strip() or None
        if units is None and per_value is None:
            continue
        if sanitized_path in mapping:
            existing = mapping[sanitized_path]
            units_mismatch = units and existing.get("units") and units != existing["units"]
            per_mismatch = per_value and existing.get("per") and per_value != existing["per"]
            if units_mismatch or per_mismatch:
                print(
                    "[WARN] Multiple Units/Per... values for auxiliary branch "
                    f"'{sanitized_path}'. Using '{existing.get('units')}/{existing.get('per')}', "
                    f"ignoring '{units}/{per_value}'."
                )
            continue
        mapping[sanitized_path] = {"units": units, "per": per_value}
    return mapping

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
    - Auxiliary fuel branches require Units and Per... columns to supply numerator/denominator units
    - default_branch_type uses position-based logic:
        * First element: for branches with 2+ children below them
        * Second element: for branches with exactly 1 child below them
        * Third element: for leaf branches (no children)
        > this arg is used when branch_type_mapping does not provide a type for a given path. this is currently only used for demand branches. in future it would be good to shift to only using branch_type_mapping for all branch types but its a bit difficult to do that right now without breaking existing functionality.
    - branch_type_mapping overrides default_branch_type for specific paths. set in create_branches_from_export_file 
    """
    ensure_api_write_allowed("leap_core.create_branches_from_export_file")
    if L is None:
        raise RuntimeError("LEAP application instance (L) is required to create branches.")
    ensure_analysis_view_context(
        L,
        context_label="create_branches_from_export_file",
    )

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

    auxiliary_units_by_path = _build_auxiliary_units_by_path(df, branch_path_col)

    branch_paths_raw = [
        bp
        for bp in df[branch_path_col].dropna().unique()
        if isinstance(bp, str) and not is_placeholder_branch_path(bp)
    ]
    branch_paths = []
    seen = set()
    for bp in branch_paths_raw:
        sanitized_bp = sanitize_leap_branch_path(bp)
        if is_placeholder_branch_path(sanitized_bp) or sanitized_bp in seen:
            continue
        seen.add(sanitized_bp)
        branch_paths.append(sanitized_bp)
    branch_paths = sorted(branch_paths, key=lambda x: len(x.split("\\")))

    created = []
    skipped = []
    skipped_existing = []
    skipped_locked = []
    failed = []
    branch_type_mapping = branch_type_mapping or {}#if we were provided a branchtype mapping then the branch types will be inferred from that where possible
    branch_paths_copy = branch_paths.copy()
    def _is_locked_transformation_fuel_child(path: str | None) -> bool:
        text = str(path or "").strip()
        return text.startswith("Transformation\\") and (
            "\\Auxiliary Fuels\\" in text
            or "\\Feedstock Fuels\\" in text
            or "\\Output Fuels\\" in text
        )
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
        if _is_locked_transformation_fuel_child(bp):
            skipped_locked.append(bp)
            skipped.append(bp)
            continue
        if safe_branch_call(L, bp, AUTO_SET_MISSING_BRANCHES=False, THROW_ERROR_ON_MISSING=False) is not None:
            skipped_existing.append(bp)
            skipped.append(bp)
            continue
                
        node = _ensure_path_exists_create_if_not(
            L,
            bp,
            branch_root,
            branch_paths_copy,
            branch_type_mapping,
            default_branch_type,
            auxiliary_units_by_path,
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

    print(
        "[INFO] Branch creation complete. "
        f"Created {len(created)}, skipped existing {len(skipped_existing)}, "
        f"skipped locked {len(skipped_locked)}."
    )
    if missing_process_categories:
        print(
            "[WARN] The following Transformation process categories need to be "
            "created manually in LEAP before importing child data:"
        )
        for category in sorted(missing_process_categories):
            print(f"  - {category}")
    return {
        "created": created,
        "skipped": skipped,
        "skipped_existing": skipped_existing,
        "skipped_locked": skipped_locked,
        "failed": failed,
    }


def fill_branches_from_export_file(
    L,
    leap_export_filename,
    sheet_name="LEAP",
    scenario=None,
    region=None,
    RAISE_ERROR_ON_FAILED_SET=False,
    RAISE_ERROR_ON_MISSING_SCENARIO=True,
    SET_UNITS=True,
    HANDLE_CURRENT_ACCOUNTS_TOO=False,
    NORMALIZE_SHARE_COLUMNS=False,
    PROMPT_ON_SHARE_MISMATCH=True,
    SHARE_NORMALIZATION_TOLERANCE=1e-3,
    CHECK_STALE_CHILD_BRANCHES=False,
    PROMPT_DELETE_STALE_BRANCHES=True,
    AUTO_DELETE_STALE_BRANCHES=False,
    STALE_BRANCH_MIN_DEPTH=3,
    STALE_BRANCH_REQUIRE_PARENT_IN_IMPORT=True,
    SKIP_VARIABLES=None,
    RUN_FUEL_CATALOG_PREFLIGHT=True,
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
    NORMALIZE_SHARE_COLUMNS : bool
        When PROMPT_ON_SHARE_MISMATCH is False, controls whether detected
        share/saturation mismatches are auto-normalized.
        Default is False to avoid silent value changes.
    PROMPT_ON_SHARE_MISMATCH : bool
        If True, prompt before normalizing share/saturation mismatches.
        If False, use NORMALIZE_SHARE_COLUMNS behavior.
    SHARE_NORMALIZATION_TOLERANCE : float
        Absolute tolerance used when checking whether share totals already
        match the target sum.
    CHECK_STALE_CHILD_BRANCHES : bool
        If True, compare existing LEAP child branches against imported child
        branches for imported parents and flag stale branches.
    PROMPT_DELETE_STALE_BRANCHES : bool
        If True (and stale checking is enabled), prompt whether to delete each
        stale branch.
    AUTO_DELETE_STALE_BRANCHES : bool
        If True, delete stale branches without prompting.
    STALE_BRANCH_MIN_DEPTH : int
        Minimum parent depth (number of path segments) for stale checks.
        Default 3 avoids high-level root checks (e.g., only sector-level+).
    STALE_BRANCH_REQUIRE_PARENT_IN_IMPORT : bool
        If True, only check parents that also appear directly in import data.
    RAISE_ERROR_ON_FAILED_SET : bool
        Whether to raise error if setting a variable fails
    RAISE_ERROR_ON_MISSING_SCENARIO : bool
        Whether to raise if a requested scenario is not present in LEAP.
        If False, missing scenarios are logged and skipped.
    SKIP_VARIABLES : list[str] | set[str] | None
        Variable names to skip setting (case-insensitive). Useful for variables
        that are not allowed or applicable in a given LEAP branch.
    RUN_FUEL_CATALOG_PREFLIGHT : bool
        If True, run shared fuel-catalog preflight checks before filling values.
        
    Returns:
    --------
    dict with keys 'success', 'failed', 'skipped' containing lists of (branch_path, variable) tuples
    """
    ensure_api_write_allowed("leap_core.fill_branches_from_export_file")
    if L is None:
        raise RuntimeError("LEAP application instance (L) is required to fill branches.")

    if RUN_FUEL_CATALOG_PREFLIGHT:
        fuel_catalog_preflight.run_fuel_catalog_preflight(
            export_path=leap_export_filename,
            sheet_name=sheet_name,
            scenario=scenario,
            context="fill_branches_from_export_file",
            leap_app=L,
        )

    skip_variables = {str(v).strip().lower() for v in (SKIP_VARIABLES or []) if str(v).strip()}

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
    available_import_scenarios: set[str] = set()
    if "Scenario" in df.columns:
        available_import_scenarios = {
            str(value).strip().lower()
            for value in df["Scenario"].dropna().astype(str).tolist()
            if str(value).strip()
        }

    stale_branch_decisions = {}
    stale_delete_all = False
    stale_keep_all = False
    stale_check_has_run = False

    def _branch_depth(path: str) -> int:
        if path is None:
            return 0
        text = str(path).strip()
        if not text:
            return 0
        return len([part for part in text.split("\\") if part])

    def _split_parent_child(path: str):
        if path is None:
            return None, None
        text = str(path).strip()
        if not text or "\\" not in text:
            return None, None
        parent, child = text.rsplit("\\", 1)
        parent = parent.strip()
        child = child.strip()
        if not parent or not child:
            return None, None
        return parent, child

    def _resolve_branch_by_paths(sanitized_path: str, original_path: str | None = None):
        branch = safe_branch_call(
            L,
            sanitized_path,
            AUTO_SET_MISSING_BRANCHES=False,
            THROW_ERROR_ON_MISSING=False,
        )
        if (
            branch is None
            and original_path
            and str(original_path).strip()
            and str(original_path).strip() != str(sanitized_path).strip()
        ):
            branch = safe_branch_call(
                L,
                original_path,
                AUTO_SET_MISSING_BRANCHES=False,
                THROW_ERROR_ON_MISSING=False,
            )
        return branch

    def _swap_resources_primary_secondary(path: str | None) -> str | None:
        text = str(path or "").strip()
        if not text:
            return None
        parts = [segment.strip() for segment in text.split("\\") if segment.strip()]
        if len(parts) < 2:
            return None
        if parts[0].lower() != "resources":
            return None
        level = parts[1].lower()
        if level == "primary":
            parts[1] = "Secondary"
            return "\\".join(parts)
        if level == "secondary":
            parts[1] = "Primary"
            return "\\".join(parts)
        return None

    def _resolve_branch_with_resource_fallback(
        sanitized_path: str,
        original_path: str | None = None,
    ) -> tuple[object | None, str]:
        branch = safe_branch_call(
            L,
            sanitized_path,
            AUTO_SET_MISSING_BRANCHES=False,
            THROW_ERROR_ON_MISSING=False,
        )
        if branch is not None:
            return branch, sanitized_path

        original_text = str(original_path or "").strip()
        if original_text and original_text != str(sanitized_path).strip():
            branch = safe_branch_call(
                L,
                original_text,
                AUTO_SET_MISSING_BRANCHES=False,
                THROW_ERROR_ON_MISSING=False,
            )
            if branch is not None:
                return branch, original_text

        candidates = [
            _swap_resources_primary_secondary(sanitized_path),
            _swap_resources_primary_secondary(original_text),
        ]
        seen: set[str] = set()
        for candidate in candidates:
            if not candidate:
                continue
            token = str(candidate).strip()
            if not token or token in seen:
                continue
            seen.add(token)
            branch = safe_branch_call(
                L,
                token,
                AUTO_SET_MISSING_BRANCHES=False,
                THROW_ERROR_ON_MISSING=False,
            )
            if branch is not None:
                source = original_text or str(sanitized_path).strip()
                print(
                    f"[INFO] Branch fallback matched '{source}' to existing '{token}'."
                )
                return branch, token
        return None, original_text or str(sanitized_path).strip()

    def _list_child_branches(parent_branch):
        rows = []
        try:
            children = parent_branch.Children
            child_count = int(children.Count)
        except Exception as exc:
            print(f"[WARN] Could not read child branches: {exc}")
            return rows

        for idx in range(1, child_count + 1):
            try:
                child = children.Item(idx)
            except Exception:
                continue
            try:
                child_name = str(child.Name).strip()
            except Exception:
                child_name = ""
            try:
                full_name = str(child.FullName).strip()
            except Exception:
                full_name = ""
            if not child_name and full_name and "\\" in full_name:
                child_name = full_name.rsplit("\\", 1)[-1].strip()
            rows.append((child_name, full_name))
        return rows

    def _delete_branch_if_exists(branch_path: str) -> bool:
        branch = safe_branch_call(
            L,
            branch_path,
            AUTO_SET_MISSING_BRANCHES=False,
            THROW_ERROR_ON_MISSING=False,
        )
        if branch is None:
            print(f"[INFO] Stale branch already missing: {branch_path}")
            return True
        try:
            branch.Delete()
            try:
                L.RefreshBranches()
            except Exception:
                pass
            print(f"[INFO] Deleted stale branch: {branch_path}")
            return True
        except Exception as exc:
            print(f"[WARN] Failed to delete stale branch '{branch_path}': {exc}")
            return False

    def _prompt_stale_branch_action(branch_path: str, parent_path: str) -> str:
        nonlocal stale_delete_all, stale_keep_all
        if AUTO_DELETE_STALE_BRANCHES:
            return "delete"
        if stale_delete_all:
            return "delete"
        if stale_keep_all:
            return "keep"
        if not PROMPT_DELETE_STALE_BRANCHES:
            return "keep"

        prompt = (
            f"[PROMPT] Branch '{branch_path}' exists under '{parent_path}' but is not in import data.\n"
            "Delete this branch? [y]es/[n]o/[a]ll-delete/[s]kip-all/[q]uit: "
        )
        while True:
            choice = input(prompt).strip().lower()
            if choice in {"y", "yes"}:
                return "delete"
            if choice in {"n", "no", ""}:
                return "keep"
            if choice in {"a", "all", "all-delete"}:
                stale_delete_all = True
                return "delete"
            if choice in {"s", "skip-all"}:
                stale_keep_all = True
                return "keep"
            if choice in {"q", "quit", "abort", "b", "break"}:
                raise RuntimeError("User aborted while handling stale branches.")
            print("Please enter y, n, a, s, or q.")

    def _check_and_handle_stale_children(df_in):
        nonlocal stale_check_has_run
        if not CHECK_STALE_CHILD_BRANCHES:
            return
        if stale_check_has_run:
            return
        stale_check_has_run = True

        expected_children_by_parent = {}
        sanitized_paths_in_import = set()
        parent_original_by_sanitized = {}

        path_rows = (
            df_in[["Branch Path", "_Sanitized Branch Path"]]
            .dropna(subset=["_Sanitized Branch Path"])
            .drop_duplicates(subset=["_Sanitized Branch Path"])
        )

        for _, row in path_rows.iterrows():
            sanitized_path = str(row["_Sanitized Branch Path"]).strip()
            if not sanitized_path:
                continue
            sanitized_paths_in_import.add(sanitized_path)

            sanitized_parent, child_name = _split_parent_child(sanitized_path)
            if sanitized_parent is None:
                continue
            if _branch_depth(sanitized_parent) < int(STALE_BRANCH_MIN_DEPTH):
                continue
            expected_children_by_parent.setdefault(sanitized_parent, set()).add(
                child_name.lower()
            )

            original_path = str(row["Branch Path"]).strip() if "Branch Path" in row else ""
            original_parent, _ = _split_parent_child(original_path)
            if original_parent and sanitized_parent not in parent_original_by_sanitized:
                parent_original_by_sanitized[sanitized_parent] = original_parent

        if STALE_BRANCH_REQUIRE_PARENT_IN_IMPORT:
            expected_children_by_parent = {
                parent_path: children
                for parent_path, children in expected_children_by_parent.items()
                if parent_path in sanitized_paths_in_import
            }

        stale_candidates = []
        for parent_path in sorted(expected_children_by_parent):
            original_parent = parent_original_by_sanitized.get(parent_path)
            parent_branch = _resolve_branch_by_paths(parent_path, original_parent)
            if parent_branch is None:
                print(
                    f"[WARN] Could not locate parent branch for stale check: "
                    f"'{parent_path}' (original: '{original_parent}')"
                )
                continue

            expected_child_names = expected_children_by_parent[parent_path]
            for child_name, full_name in _list_child_branches(parent_branch):
                if not child_name:
                    continue
                if child_name.lower() in expected_child_names:
                    continue
                stale_path = full_name if full_name else f"{parent_path}\\{child_name}"
                stale_candidates.append((parent_path, stale_path))

        if not stale_candidates:
            print("[INFO] Stale-branch check: no unexpected child branches found.")
            return

        print(
            f"[WARN] Stale-branch check found {len(stale_candidates)} "
            "unexpected child branch(es)."
        )

        for parent_path, stale_path in stale_candidates:
            existing_decision = stale_branch_decisions.get(stale_path)
            if existing_decision in {"kept", "deleted"}:
                continue

            action = _prompt_stale_branch_action(stale_path, parent_path)
            if action == "keep":
                stale_branch_decisions[stale_path] = "kept"
                print(f"[INFO] Keeping stale branch: {stale_path}")
                continue

            if _delete_branch_if_exists(stale_path):
                stale_branch_decisions[stale_path] = "deleted"
            else:
                stale_branch_decisions[stale_path] = "delete_failed"
                print(
                    "[WARN] Could not delete stale branch via API. "
                    f"Please delete manually in LEAP: {stale_path}"
                )

    def _normalize_share_columns_wide(df_in, year_cols, tol=1e-3, apply_changes=True):
        """
        Normalize wide share rows so siblings under the same parent sum to target.

        The target is 100 by default; paths containing "nonspecified" use 1.
        """
        summary = {
            "issue_group_years": 0,
            "issue_cells": 0,
            "normalized_group_years": 0,
            "normalized_cells": 0,
            "sample": [],
        }
        if df_in is None or df_in.empty:
            return df_in, summary
        if "Variable" not in df_in.columns or "Branch Path" not in df_in.columns:
            return df_in, summary
        if not year_cols:
            return df_in, summary

        working = df_in.copy()
        if "_Sanitized Branch Path" not in working.columns:
            working["_Sanitized Branch Path"] = working["Branch Path"].apply(
                sanitize_leap_branch_path
            )
        working["_Parent Path"] = working["_Sanitized Branch Path"].apply(
            lambda x: "\\".join(str(x).split("\\")[:-1]) if isinstance(x, str) and x else ""
        )

        variable_text = working["Variable"].astype(str).str.strip().str.lower()
        if "Units" in working.columns:
            units_text = working["Units"].astype(str).str.strip().str.lower()
        else:
            units_text = pd.Series("", index=working.index)
        share_mask = (
            variable_text.str.contains("share", na=False)
            | variable_text.str.contains("saturation", na=False)
            | units_text.isin({"share", "saturation"})
        )
        # Respect SKIP_VARIABLES during diagnostics/normalization as well as fill.
        if skip_variables:
            share_mask = share_mask & ~variable_text.isin(skip_variables)
        if not share_mask.any():
            return working, summary

        group_cols = ["_Parent Path", "Variable"]
        for optional_col in ("Scenario", "Region"):
            if optional_col in working.columns:
                group_cols.append(optional_col)

        subset = working.loc[share_mask]
        grouped = subset.groupby(group_cols, dropna=False).groups

        for _, idx_values in grouped.items():
            idx = pd.Index(idx_values)
            if len(idx) <= 1:
                continue
            # Ignore duplicate rows for a single branch path; normalize siblings only.
            if working.loc[idx, "_Sanitized Branch Path"].nunique() <= 1:
                continue

            parent_path = str(working.loc[idx, "_Parent Path"].iloc[0])
            target = 1.0 if "nonspecified" in parent_path.lower() else 100.0
            variable_name = str(working.loc[idx, "Variable"].iloc[0])

            for year in year_cols:
                numeric_values = pd.to_numeric(working.loc[idx, year], errors="coerce")
                valid_idx = numeric_values[numeric_values.notna()].index
                valid_count = len(valid_idx)
                if valid_count <= 1:
                    continue

                total = float(numeric_values.loc[valid_idx].sum())
                if abs(total - target) <= tol:
                    continue

                summary["issue_group_years"] += 1
                summary["issue_cells"] += valid_count

                if apply_changes:
                    if abs(total) <= tol:
                        replacement = pd.Series(
                            target / valid_count,
                            index=valid_idx,
                            dtype=float,
                        )
                    else:
                        replacement = numeric_values.loc[valid_idx] * (target / total)
                    working.loc[valid_idx, year] = replacement
                    summary["normalized_group_years"] += 1
                    summary["normalized_cells"] += valid_count
                if len(summary["sample"]) < 5:
                    summary["sample"].append(
                        {
                            "variable": variable_name,
                            "parent_path": parent_path,
                            "year": int(year),
                            "total_before": total,
                            "target": target,
                        }
                    )

        working = working.drop(columns=["_Parent Path"], errors="ignore")
        return working, summary

    def _prompt_yes_no(message, default=False):
        hint = "[Y/n]" if default else "[y/N]"
        while True:
            try:
                choice = input(f"{message} {hint}: ").strip().lower()
            except EOFError:
                print("[WARN] Prompt input unavailable; defaulting to 'No'.")
                return False if not default else True
            if choice == "":
                return default
            if choice in {"y", "yes"}:
                return True
            if choice in {"n", "no"}:
                return False
            print("Please enter 'y' or 'n'.")

    def _is_leap_locked_transformation_fuel_child(
        path: str | None,
        variable_name: str | None = None,
    ) -> bool:
        text = str(path or "").strip()
        is_fuel_child = text.startswith("Transformation\\") and (
            "\\Auxiliary Fuels\\" in text
            or "\\Feedstock Fuels\\" in text
            or "\\Output Fuels\\" in text
        )
        if not is_fuel_child:
            return False
        # Import/Export targets on Output Fuels must still be writable so we can
        # explicitly clear stale trade settings in LEAP.
        var_text = str(variable_name or "").strip().lower()
        if var_text in {"import target", "export target"}:
            return False
        return True

    def _fill_from_df(df_in):
        current_accounts_scenario_active = (
            str(scenario or "").strip().lower() in {"current accounts", "current account"}
        )
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
        if "Branch Path" in df_in.columns:
            before_count = len(df_in)
            df_in = df_in[~df_in["Branch Path"].apply(is_placeholder_branch_path)]
            dropped = before_count - len(df_in)
            if dropped > 0:
                print(f"[INFO] Dropped {dropped} row(s) with placeholder Branch Path values.")
        df_in["_Sanitized Branch Path"] = df_in["Branch Path"].apply(
            sanitize_leap_branch_path
        )
        df_in = df_in[df_in["_Sanitized Branch Path"] != ""]
        df_in = df_in[~df_in["_Sanitized Branch Path"].apply(is_placeholder_branch_path)]
        _check_and_handle_stale_children(df_in)

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

        if year_cols == ["Expression"] and (NORMALIZE_SHARE_COLUMNS or PROMPT_ON_SHARE_MISMATCH):
            print(
                "[WARN] Share normalization skipped because workbook uses 'Expression' "
                "instead of explicit year columns."
            )
        elif year_cols != ["Expression"]:
            _, normalization_summary = _normalize_share_columns_wide(
                df_in,
                year_cols=year_cols,
                tol=SHARE_NORMALIZATION_TOLERANCE,
                apply_changes=False,
            )
            if normalization_summary["issue_group_years"] > 0:
                print(
                    "[WARN] Share/saturation totals outside tolerance for "
                    f"{normalization_summary['issue_group_years']} parent-year groups "
                    f"({normalization_summary['issue_cells']} cells)."
                )
                for sample in normalization_summary["sample"]:
                    parent_path = sample["parent_path"] or "<root>"
                    print(
                        f"[WARN]   {sample['variable']} | {parent_path} | {sample['year']}: "
                        f"{sample['total_before']:.6g} vs target {sample['target']:.6g}"
                    )
                if PROMPT_ON_SHARE_MISMATCH:
                    do_normalize = _prompt_yes_no(
                        "Normalize these share/saturation values before importing?",
                        default=False,
                    )
                else:
                    do_normalize = bool(NORMALIZE_SHARE_COLUMNS)
                if do_normalize:
                    df_in, applied_summary = _normalize_share_columns_wide(
                        df_in,
                        year_cols=year_cols,
                        tol=SHARE_NORMALIZATION_TOLERANCE,
                        apply_changes=True,
                    )
                    print(
                        "[INFO] Share normalization applied to "
                        f"{applied_summary['normalized_group_years']} parent-year groups "
                        f"({applied_summary['normalized_cells']} cells)."
                    )
                else:
                    print("[INFO] Share normalization skipped by user choice/config.")
            else:
                print(
                    "[INFO] Share/saturation totals already within tolerance; "
                    "no normalization needed."
                )

        success = []
        failed = []
        skipped = []
        warning_counts: dict[str, int] = {}
        suppressed_warning_kinds: set[str] = set()

        def _warn_limited(kind: str, message: str) -> None:
            count = warning_counts.get(kind, 0) + 1
            warning_counts[kind] = count
            limit = _warning_print_limit()
            if _is_summary_logging():
                limit = min(limit, 5)
            if count <= limit:
                print(message)
            elif kind not in suppressed_warning_kinds:
                suppressed_warning_kinds.add(kind)
                print(
                    f"[WARN] Suppressing additional '{kind}' warnings after {limit} messages. "
                    "Counts will be shown in summary."
                )

        missing_feedstock_share_rows = []
        try:
            variable_text = df_in["Variable"].astype(str).str.strip().str.lower()
            feedstock_mask = variable_text.eq("feedstock fuel share")
            if feedstock_mask.any():
                feedstock_df = df_in.loc[feedstock_mask].copy()
                feedstock_df["_Parent Path"] = feedstock_df["_Sanitized Branch Path"].apply(
                    lambda x: "\\".join(str(x).split("\\")[:-1])
                    if isinstance(x, str) and x
                    else ""
                )
                feedstock_groups = feedstock_df.groupby("_Parent Path", dropna=False)
                for parent_path, group in feedstock_groups:
                    parent_path = str(parent_path or "").strip()
                    if not parent_path:
                        continue
                    expected_children = set(
                        group["_Sanitized Branch Path"].astype(str).str.strip().str.lower()
                    )
                    if not expected_children:
                        continue
                    parent_branch, parent_used = _resolve_branch_with_resource_fallback(
                        parent_path,
                        parent_path,
                    )
                    if parent_branch is None:
                        continue
                    for child_name, full_name in _list_child_branches(parent_branch):
                        child_path = full_name if full_name else f"{parent_used}\\{child_name}"
                        child_path = sanitize_leap_branch_path(child_path)
                        if not child_path:
                            continue
                        if child_path.lower() in expected_children:
                            continue
                        missing_feedstock_share_rows.append(
                            {
                                "branch_path": child_path,
                                "variable": "Feedstock Fuel Share",
                                "units": group["Units"].iloc[0] if "Units" in group.columns else None,
                            }
                        )
                df_in = df_in.drop(columns=["_Parent Path"], errors="ignore")
        except Exception as exc:
            print(f"[WARN] Failed to prepare missing feedstock share rows: {exc}")
        
        # Group by branch path and variable
        
        for (sanitized_bp, var), group in df_in.groupby(
            ["_Sanitized Branch Path", "Variable"]
        ):
            if skip_variables and str(var).strip().lower() in skip_variables:
                original_bp = group["Branch Path"].iloc[0]
                skipped.append((original_bp, var))
                continue
            original_bp = group["Branch Path"].iloc[0]
            if _is_leap_locked_transformation_fuel_child(sanitized_bp, variable_name=var):
                _warn_limited(
                    "locked_transformation_fuel_child",
                    "[WARN] Skipping LEAP-locked transformation fuel-child import row "
                    f"'{original_bp}' / '{var}'.",
                )
                skipped.append((original_bp, var))
                continue
            branch, branch_path_used = _resolve_branch_with_resource_fallback(
                sanitized_bp,
                original_bp,
            )
            
            if branch is None:
                msg = (
                    f"Branch '{original_bp}' (sanitized '{sanitized_bp}') not found - "
                    f"skipping variable '{var}'"
                )
                if RAISE_ERROR_ON_FAILED_SET:
                    breakpoint()
                    raise RuntimeError(msg)
                else:
                    _warn_limited("missing_branch", f"[WARN] {msg}")
                    failed.append((branch_path_used, var))
                    continue
            if year_cols == ["Expression"]:
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

                # For Current Accounts, preserve a single base-year point as Data(year, value)
                # so LEAP does not treat it as a constant across future years.
                if current_accounts_scenario_active and len(points) == 1:
                    single_year, single_value = points[0]
                    expr = f"Data({int(single_year)}, {float(single_value):.6g})"
                else:
                    # Build expression for multi-year time series values.
                    expr = build_expr(points, expression_type="Interp")
            
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
                L,
                branch,
                var,
                expr,
                unit_name=unit_name,
                context=branch_path_used,
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

        if missing_feedstock_share_rows:
            zero_expr = (
                "0"
                if year_cols == ["Expression"]
                else build_expr([(int(year), 0.0) for year in year_cols], expression_type="Interp")
            )
            for item in missing_feedstock_share_rows:
                if _is_leap_locked_transformation_fuel_child(
                    item["branch_path"],
                    variable_name=item.get("variable"),
                ):
                    _warn_limited(
                        "locked_feedstock_backfill",
                        "[WARN] Skipping zero Feedstock Fuel Share backfill for "
                        f"LEAP-locked branch '{item['branch_path']}'.",
                    )
                    skipped.append((item["branch_path"], item["variable"]))
                    continue
                branch, branch_path_used = _resolve_branch_with_resource_fallback(
                    item["branch_path"],
                    item["branch_path"],
                )
                if branch is None:
                    continue
                set_success = safe_set_variable(
                    L,
                    branch,
                    item["variable"],
                    zero_expr,
                    unit_name=item.get("units") if SET_UNITS else None,
                    context=branch_path_used,
                )
                if set_success:
                    success.append((branch_path_used, item["variable"]))
                else:
                    failed.append((branch_path_used, item["variable"]))
            print(
                "[INFO] Applied zero feedstock shares to "
                f"{len(missing_feedstock_share_rows)} stale sibling branch(es)."
            )

        print(
            f"[INFO] Data fill complete. Success: {len(success)}, "
            f"Failed: {len(failed)}, Skipped: {len(skipped)}"
        )
        if warning_counts:
            summary_text = ", ".join(
                f"{kind}={count}" for kind, count in sorted(warning_counts.items())
            )
            print(f"[INFO] Warning summary: {summary_text}")
        return {"success": success, "failed": failed, "skipped": skipped}

    def _list_scenarios():
        scenarios = []
        try:
            coll = L.Scenarios
            count = int(coll.Count)
            for idx in range(1, count + 1):
                try:
                    item = coll.Item(idx)
                    scenarios.append(item)
                except Exception:
                    continue
        except Exception:
            return []
        return scenarios

    def _resolve_target_scenario_name(name):
        if not name:
            return None
        scenario_items = _list_scenarios()
        if not scenario_items:
            return None
        if name == "Current Accounts":
            for item in scenario_items:
                try:
                    if bool(item.IsCurrentAccounts):
                        return str(item.Name)
                except Exception:
                    continue
        requested = str(name).strip().lower()
        for item in scenario_items:
            try:
                if str(item.Name).strip().lower() == requested:
                    return str(item.Name)
            except Exception:
                continue
        for item in scenario_items:
            try:
                if str(item.Abbreviation).strip().lower() == requested:
                    return str(item.Name)
            except Exception:
                continue
        return None

    def _list_scenario_names():
        names = []
        for item in _list_scenarios():
            try:
                name = str(item.Name).strip()
            except Exception:
                continue
            if name and name not in names:
                names.append(name)
        return names

    def _activate_scenario(name):
        if not name:
            return True
        resolved = _resolve_target_scenario_name(name)
        if resolved is None:
            available = _list_scenario_names()
            msg = (
                f"Scenario '{name}' was not found in LEAP."
                f" Available LEAP scenarios: {available}."
            )
            if RAISE_ERROR_ON_MISSING_SCENARIO:
                raise ValueError(msg)
            print(f"[WARN] {msg} Skipping fill for this scenario.")
            return False
        try:
            # COM binding is more reliable when assigning by name than by object.
            L.ActiveScenario = resolved
            return True
        except Exception as exc:
            msg = f"Failed to activate LEAP scenario '{resolved}': {exc}"
            if RAISE_ERROR_ON_MISSING_SCENARIO:
                raise RuntimeError(msg) from exc
            print(f"[WARN] {msg}")
            return False

    def _activate_region(name):
        if not name:
            return True
        try:
            regions = L.Regions
            if regions.Exists(name):
                # COM binding is more reliable when assigning by name than by object.
                L.ActiveRegion = name
                return True
            # Try case-insensitive matching when direct exists fails.
            count = int(regions.Count)
            for idx in range(1, count + 1):
                item = regions.Item(idx)
                if str(item.Name).strip().lower() == str(name).strip().lower():
                    L.ActiveRegion = str(item.Name)
                    return True
            # If only one region exists in LEAP, use it as a safe fallback.
            if int(regions.Count) == 1:
                only_region = regions.Item(1)
                only_region_name = str(only_region.Name)
                print(
                    f"[WARN] Region '{name}' not found; using only LEAP region '{only_region_name}'."
                )
                L.ActiveRegion = only_region_name
                return True
        except Exception as exc:
            print(f"[WARN] Failed checking/activating region '{name}': {exc}")
            return False
        print(f"[WARN] Region '{name}' was not found in LEAP; skipping fill.")
        return False

    def _reset_ui_context():
        # LEAP can keep an invalid branch/view selection when scenario or region changes.
        # Reset to a neutral analysis context before branch lookups and writes.
        return ensure_analysis_view_context(
            L,
            context_label="fill_branches_from_export_file",
        )

    results = []
    scenarios_to_run = [scenario]
    current_accounts_in_workbook = any(
        label in available_import_scenarios
        for label in {"current accounts", "current account"}
    )
    print(
        "[INFO] fill_branches_from_export_file scenario planning: "
        f"workbook='{leap_export_filename}', "
        f"requested_scenario={scenario!r}, "
        f"HANDLE_CURRENT_ACCOUNTS_TOO={HANDLE_CURRENT_ACCOUNTS_TOO}, "
        f"current_accounts_in_workbook={current_accounts_in_workbook}, "
        f"available_import_scenarios={sorted(available_import_scenarios)}"
    )
    if HANDLE_CURRENT_ACCOUNTS_TOO and not current_accounts_in_workbook:
        print(
            "[INFO] Skipping appended Current Accounts pass because the workbook "
            f"does not contain it: {sorted(available_import_scenarios)}"
        )
    if (
        HANDLE_CURRENT_ACCOUNTS_TOO
        and current_accounts_in_workbook
        and scenario
        and scenario != "Current Accounts"
    ):
        caller_frames = [
            f"{os.path.basename(frame.filename)}:{frame.lineno}:{frame.function}"
            for frame in inspect.stack()[1:6]
        ]
        print(
            "[INFO] Appending Current Accounts pass. "
            f"Caller stack: {caller_frames}"
        )
        scenarios_to_run.append("Current Accounts")

    for scen in scenarios_to_run:
        if scen is not None:
            print(f"[INFO] Filling data for scenario '{scen}'.")
        prev_scenario = scenario
        scenario = scen
        try:
            if not _activate_scenario(scen):
                continue
            if not _activate_region(region):
                continue
            _reset_ui_context()
            results.append(_fill_from_df(df.copy()))
        finally:
            scenario = prev_scenario

    combined = {"success": [], "failed": []}
    for res in results:
        combined["success"].extend(res["success"])
        combined["failed"].extend(res["failed"])
    return combined
