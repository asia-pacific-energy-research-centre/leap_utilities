#%%
# Summary: Build LEAP supply import/export values by fuel using ESTO/9th datasets.
# Most user-editable settings live in `codebase/workflow_config.py`.
# How it works:
# - Loads ESTO/9th data and normalizes year columns.
# - Filters 9th data to chosen scenario only.
# - For each fuel, selects import/export rows based on flow labels.
# - Uses 2022 ESTO base-year values plus 9th projections for 2023+.
# - Prints import/export totals for each fuel and economy.
import os
import re
import sys
from pathlib import Path
from typing import Iterable

import pandas as pd

from codebase.utilities.master_config import config_table_exists, read_config_table

# Ensure the repository root is importable for scripts executed from any location.
REPO_ROOT = Path(__file__).resolve().parents[2]
try:
    if str(REPO_ROOT) not in sys.path:
        sys.path.insert(0, str(REPO_ROOT))
except Exception as exc:
    print(f"Failed to add repo root to sys.path: {exc}")

from codebase.configuration.config import scenario_dict
from codebase.utilities import workflow_common
from codebase.utilities import fuel_catalog_preflight
from codebase.configuration import workflow_config as workflow_cfg
from codebase.functions.analysis_input_write_dispatcher import (
    dispatch_analysis_input_write,
)
from codebase.functions.leap_core import (
    connect_to_leap,
    ensure_fuel_exists,
    fill_branches_from_export_file,
    sanitize_leap_name,
)
from codebase.scrapbook.utilities import (
    apply_matt_subtotal_mapping,
    filter_matt_subtotals,
    load_augmented_reference_tables,
    save_subtotal_labeled_data,
)
from codebase.functions.leap_excel_io import finalise_export_df, save_export_files
from codebase.configuration.all_products_and_flows import ESTO_PRODUCT_LIST
from codebase.functions.ninth_projection_mapping import (
    build_esto_projection_table,
    build_projection_lookup,
    normalize_economy_key,
)
#%%

#%%
######### CONSTANTS (UNLIKELY TO CHANGE) #########
DATA_DIR = REPO_ROOT / "data"
ESTO_DATA_PATH = DATA_DIR / "merged_file_energy_ALL_20250814_pre_trump.csv"
# Use merged_file_energy_ALL_20251106.csv and merged_file_energy_00_APEC_20251106 for exact 9th edition projection matching.
MATT_DATA_PATH = DATA_DIR / "00APEC_2024_low.csv"
CONFIG_DIR = REPO_ROOT / "config"
SUBTOTAL_MAPPING_PATH = CONFIG_DIR / "ESTO_subtotal_mapping.xlsx"
NINTH_TO_ESTO_MAPPING_PATH = CONFIG_DIR / "ninth_pairs_to_esto_pairs.xlsx"
CODE_TO_NAME_PATHS = [
    CONFIG_DIR / "sector_fuel_codes_to_names.updated.xlsx",
    CONFIG_DIR / "sector_fuel_codes_to_names.xlsx",
]

BASE_YEAR = 2022
PROJECTION_START_YEAR = 2023
PROJECTION_END_YEAR = 2060
PROJECTION_YEAR_RANGE = list(range(PROJECTION_START_YEAR, PROJECTION_END_YEAR + 1))
REFERENCE_CACHE_DIR = DATA_DIR / ".cache" / "supply_reference_tables"
ENABLE_DEBUG_BREAKPOINTS = True
PRINT_FUEL_ROWS = True
PRINT_ONLY_NONZERO_ROWS = True
PRINT_TOP_ROWS = 10
USE_CODE_TO_NAME_MAPPING = True

FLOW_CODES_BY_DATASET = {
    "esto": {
        "production": "01 Production",
        "imports": "02 Imports",
        "exports": "03 Exports",
        "stock_changes": "06 Stock changes",
        "tpes": "07 Total primary energy supply",
    },
    "ninth": {
        "production": "01_production",
        "imports": "02_imports",
        "exports": "03_exports",
        "stock_changes": "06_stock_changes",
        "tpes": "07_total_primary_energy_supply",
    },
}

######### FLOW SIGN RULES #########
# Supply data follows the LEAP/Balance convention: outputs are reported as positives or
# negatives (e.g., exports are negative in the source CSV), inputs stay positive, so we
# normalize output flows here to keep the LEAP import positive.
OUTPUT_FLOW_KEYS = {"exports"}


def is_output_flow(flow_key):
    """Return True when a flow key represents an output that should be positive in LEAP."""
    if not flow_key:
        return False
    return str(flow_key).strip().lower() in OUTPUT_FLOW_KEYS


def normalize_supply_flow_total(flow_key, total_value):
    """Normalize the sign of a supply flow total based on its LEAP meaning."""
    try:
        if is_output_flow(flow_key):
            return abs(total_value)
        return total_value
    except Exception as exc:
        print(f"Failed to normalize flow total for {flow_key}: {exc}")
        try_debug_breakpoint()
        raise

EXCLUDED_ESTO_PREFIXES = ["19", "20", "21"]
SUPPLY_MEASURES = [
    {"name": "Imports", "flow_key": "imports", "units": "Petajoule", "per": ""},
    {"name": "Exports", "flow_key": "exports", "units": "Petajoule", "per": ""},
    {
        "name": "Unmet Requirements",
        "flow_key": None,
        "units": "Percent",
        "per": "MeetWithImports",
        "value": 0.0,
    },
]
if not getattr(workflow_cfg, "SUPPLY_INCLUDE_UNMET_REQUIREMENTS", False):
    SUPPLY_MEASURES = [
        measure for measure in SUPPLY_MEASURES if measure.get("name") != "Unmet Requirements"
    ]
EXPORT_SCENARIOS = ["Current Accounts", "Reference", "Target"]
DEFAULT_EXPORT_OUTPUT_DIR = REPO_ROOT / "outputs" / "leap_exports"
EXPORT_OUTPUT_DIR = Path(
    os.environ.get("SUPPLY_LEAP_EXPORT_DIR", str(DEFAULT_EXPORT_OUTPUT_DIR))
)
EXPORT_FILENAME_TEMPLATE = "supply_leap_imports_{economy}_{scenarios}.xlsx"
EXPORT_FILENAME_REGEX = re.compile(
    r"supply_leap_imports_(?P<economy>[^_]+)_(?P<scenarios>.+)\.xlsx",
    re.IGNORECASE,
)
EXPORT_MODEL_NAME = "USA transport supply imports"
EXPORT_REGION = "United States of America"
EXPORT_BASE_YEAR = BASE_YEAR
EXPORT_FINAL_YEAR = PROJECTION_END_YEAR
EXPORT_ECONOMY_REGION_OVERRIDES = {"20USA": EXPORT_REGION}
SAVE_PROJECTION_DIAGNOSTICS = False
PROJECTION_DIAGNOSTICS_PATH = REPO_ROOT / "outputs" / "ninth_supply_projection_fallbacks.csv"
SUPPLY_PROJECTION_LOOKUP = None
SUPPLY_ROOT_CLASSIFICATION_SOURCE_PATH = Path(
    getattr(
        workflow_cfg,
        "SUPPLY_ROOT_CLASSIFICATION_SOURCE_PATH",
        REPO_ROOT / "data" / "full model export.xlsx",
    )
)
SUPPLY_ROOT_CLASSIFICATION_SOURCE_SHEET = str(
    getattr(
        workflow_cfg,
        "SUPPLY_ROOT_CLASSIFICATION_SOURCE_SHEET",
        "Export",
    )
)
SUPPLY_ROOT_CLASSIFICATION_STRICT = bool(
    getattr(
        workflow_cfg,
        "SUPPLY_ROOT_CLASSIFICATION_STRICT",
        False,
    )
)

_SUPPLY_ROOT_LOOKUP_CACHE: dict[str, str] | None = None
_SUPPLY_ROOT_LOOKUP_SOURCE_INFO: dict[str, object] | None = None
_SUPPLY_ROOT_LOOKUP_MISS_WARNED: set[str] = set()
_SUPPLY_BRANCH_PATH_LOOKUP_CACHE: set[str] | None = None
_SUPPLY_BRANCH_PATH_LOOKUP_SOURCE_INFO: dict[str, object] | None = None
_SUPPLY_BRANCH_PATH_MISS_WARNED: set[str] = set()

SECONDARY_ESTO_PRODUCT_MAJOR_CODES = {"02", "04", "07"}
SECONDARY_ESTO_PRODUCT_EXACT = {
    "06.03 Refinery feedstocks",
    "06.04 Additives/  oxygenates",
    "06.05 Other hydrocarbons",
    "08.02 LNG",
    "08.03 Gas works gas",
    "15.03 Charcoal",
    "15.04 Black liqour",
    "16.05 Biogasoline",
    "16.06 Biodiesel",
    "16.07 Bio jet kerosene",
    "16.08 Other liquid biofuels",
    "17 Electricity",
    "18 Heat",
}


def _esto_product_major_code(product):
    """Return the two-digit ESTO major product code when present."""
    text = str(product or "").strip()
    match = re.match(r"^(\d{2})(?:[.\s]|$)", text)
    if match:
        return match.group(1)
    return ""


def _is_secondary_esto_product(product):
    """Return True for ESTO products that originate from transformation/refinement."""
    major_code = _esto_product_major_code(product)
    if major_code in {"19", "20", "21"}:
        return False
    if product in SECONDARY_ESTO_PRODUCT_EXACT:
        return True
    return major_code in SECONDARY_ESTO_PRODUCT_MAJOR_CODES


ESTO_PRODUCT_CLASSIFICATION = {
    product: ("secondary" if _is_secondary_esto_product(product) else "primary")
    for product in ESTO_PRODUCT_LIST
    if _esto_product_major_code(product) not in {"19", "20", "21"}
}


def _normalize_supply_lookup_fuel_name(value):
    """Normalize a fuel label into a stable lookup key for branch-root resolution."""
    text = str(value or "").strip()
    if not text:
        return ""
    sanitized = sanitize_leap_name(text)
    normalized = " ".join(str(sanitized or "").strip().lower().split())
    return normalized


def _read_branch_variable_rows_from_workbook(
    source_path,
    sheet_name="Export",
):
    """Read workbook rows by auto-detecting the header row containing Branch Path + Variable."""
    path = Path(str(source_path).replace("\\", "/"))
    if not path.is_absolute():
        path = REPO_ROOT / path
    if not path.exists():
        return pd.DataFrame()
    raw = read_config_table(path, sheet_name=sheet_name, header=None)
    header_row = None
    for idx in range(len(raw.index)):
        values = {
            str(item).strip().lower()
            for item in raw.iloc[idx].tolist()
            if str(item).strip() and str(item).lower() != "nan"
        }
        if "branch path" in values and "variable" in values:
            header_row = int(idx)
            break
    if header_row is None:
        return pd.DataFrame()
    data = raw.iloc[header_row + 1 :].copy()
    data.columns = raw.iloc[header_row].tolist()
    if "Branch Path" not in data.columns:
        return pd.DataFrame()
    data = data[data["Branch Path"].notna()].copy()
    return data


def _load_supply_root_lookup_from_export():
    """Load and cache supply root (Primary/Secondary) lookup from the canonical LEAP export."""
    global _SUPPLY_ROOT_LOOKUP_CACHE
    global _SUPPLY_ROOT_LOOKUP_SOURCE_INFO
    if _SUPPLY_ROOT_LOOKUP_CACHE is not None:
        return _SUPPLY_ROOT_LOOKUP_CACHE

    path = Path(str(SUPPLY_ROOT_CLASSIFICATION_SOURCE_PATH).replace("\\", "/"))
    if not path.is_absolute():
        path = REPO_ROOT / path
    sheet = str(SUPPLY_ROOT_CLASSIFICATION_SOURCE_SHEET).strip() or "Export"

    if not path.exists():
        print(
            "[WARN] Supply root classification source workbook not found: "
            f"{path}. Falling back to legacy ESTO-based classification."
        )
        _SUPPLY_ROOT_LOOKUP_SOURCE_INFO = {
            "path": str(path),
            "sheet": sheet,
            "status": "missing",
            "lookup_size": 0,
        }
        _SUPPLY_ROOT_LOOKUP_CACHE = {}
        return _SUPPLY_ROOT_LOOKUP_CACHE

    try:
        rows = _read_branch_variable_rows_from_workbook(path, sheet_name=sheet)
    except Exception as exc:
        print(
            "[WARN] Failed reading supply root classification source workbook "
            f"{path} (sheet={sheet}): {exc}. Falling back to legacy ESTO-based classification."
        )
        _SUPPLY_ROOT_LOOKUP_SOURCE_INFO = {
            "path": str(path),
            "sheet": sheet,
            "status": "read_failed",
            "lookup_size": 0,
        }
        _SUPPLY_ROOT_LOOKUP_CACHE = {}
        return _SUPPLY_ROOT_LOOKUP_CACHE

    root_counts_by_fuel: dict[str, dict[str, int]] = {}
    for branch_path in rows.get("Branch Path", pd.Series(dtype=str)).astype(str):
        parts = [part.strip() for part in str(branch_path or "").split("\\") if part.strip()]
        if len(parts) < 3:
            continue
        if parts[0].lower() != "resources":
            continue
        root = parts[1].strip().title()
        if root not in {"Primary", "Secondary"}:
            continue
        fuel_key = _normalize_supply_lookup_fuel_name(parts[2])
        if not fuel_key:
            continue
        bucket = root_counts_by_fuel.setdefault(fuel_key, {"Primary": 0, "Secondary": 0})
        bucket[root] = int(bucket.get(root, 0)) + 1

    lookup: dict[str, str] = {}
    conflicts: list[tuple[str, dict[str, int]]] = []
    for fuel_key, counts in root_counts_by_fuel.items():
        primary_count = int(counts.get("Primary", 0))
        secondary_count = int(counts.get("Secondary", 0))
        if primary_count and secondary_count:
            conflicts.append((fuel_key, counts))
            chosen = "Primary" if primary_count >= secondary_count else "Secondary"
            lookup[fuel_key] = chosen
            continue
        lookup[fuel_key] = "Secondary" if secondary_count > 0 else "Primary"

    if conflicts:
        preview = ", ".join(
            [
                f"{fuel} (Primary={counts.get('Primary', 0)}, Secondary={counts.get('Secondary', 0)})"
                for fuel, counts in conflicts[:20]
            ]
        )
        print(
            "[WARN] Supply root source has fuel(s) mapped to both Primary and Secondary; "
            f"using majority root for {len(conflicts)} fuel(s). Sample: {preview}"
        )

    print(
        "[INFO] Loaded supply root classification lookup from LEAP export source: "
        f"{path} (sheet={sheet}, fuels={len(lookup)})."
    )
    _SUPPLY_ROOT_LOOKUP_SOURCE_INFO = {
        "path": str(path),
        "sheet": sheet,
        "status": "loaded",
        "lookup_size": int(len(lookup)),
        "conflicts": int(len(conflicts)),
    }
    _SUPPLY_ROOT_LOOKUP_CACHE = lookup
    return _SUPPLY_ROOT_LOOKUP_CACHE


def _resolve_supply_root_from_export_lookup(*candidates):
    """Resolve Primary/Secondary from workbook lookup using candidate labels."""
    lookup = _load_supply_root_lookup_from_export()
    if not lookup:
        return None
    for candidate in candidates:
        key = _normalize_supply_lookup_fuel_name(candidate)
        if not key:
            continue
        root = lookup.get(key)
        if root in {"Primary", "Secondary"}:
            return root
    return None


def _normalize_supply_branch_path_for_lookup(branch_path):
    """Normalize Resources branch paths for existence checks against export source."""
    parts = [part.strip() for part in str(branch_path or "").split("\\") if part.strip()]
    if len(parts) < 3:
        return ""
    if parts[0].lower() != "resources":
        return ""
    root = parts[1].strip().lower()
    if root not in {"primary", "secondary"}:
        return ""
    fuel = _normalize_supply_lookup_fuel_name(parts[2])
    if not fuel:
        return ""
    return f"resources\\{root}\\{fuel}"


def _load_supply_branch_path_lookup_from_export():
    """Load existing Resources branch paths from canonical LEAP export source."""
    global _SUPPLY_BRANCH_PATH_LOOKUP_CACHE
    global _SUPPLY_BRANCH_PATH_LOOKUP_SOURCE_INFO
    if _SUPPLY_BRANCH_PATH_LOOKUP_CACHE is not None:
        return _SUPPLY_BRANCH_PATH_LOOKUP_CACHE

    path = Path(str(SUPPLY_ROOT_CLASSIFICATION_SOURCE_PATH).replace("\\", "/"))
    if not path.is_absolute():
        path = REPO_ROOT / path
    sheet = str(SUPPLY_ROOT_CLASSIFICATION_SOURCE_SHEET).strip() or "Export"
    if not path.exists():
        _SUPPLY_BRANCH_PATH_LOOKUP_SOURCE_INFO = {
            "path": str(path),
            "sheet": sheet,
            "status": "missing",
            "lookup_size": 0,
        }
        _SUPPLY_BRANCH_PATH_LOOKUP_CACHE = set()
        return _SUPPLY_BRANCH_PATH_LOOKUP_CACHE
    try:
        rows = _read_branch_variable_rows_from_workbook(path, sheet_name=sheet)
    except Exception as exc:
        print(
            "[WARN] Failed reading supply branch-path lookup from LEAP export source "
            f"{path} (sheet={sheet}): {exc}. Branch existence checks disabled."
        )
        _SUPPLY_BRANCH_PATH_LOOKUP_SOURCE_INFO = {
            "path": str(path),
            "sheet": sheet,
            "status": "read_failed",
            "lookup_size": 0,
        }
        _SUPPLY_BRANCH_PATH_LOOKUP_CACHE = set()
        return _SUPPLY_BRANCH_PATH_LOOKUP_CACHE
    lookup: set[str] = set()
    for branch_path in rows.get("Branch Path", pd.Series(dtype=str)).astype(str):
        token = _normalize_supply_branch_path_for_lookup(branch_path)
        if token:
            lookup.add(token)
    _SUPPLY_BRANCH_PATH_LOOKUP_SOURCE_INFO = {
        "path": str(path),
        "sheet": sheet,
        "status": "loaded",
        "lookup_size": int(len(lookup)),
    }
    _SUPPLY_BRANCH_PATH_LOOKUP_CACHE = lookup
    return _SUPPLY_BRANCH_PATH_LOOKUP_CACHE


def _supply_branch_exists_in_export_source(branch_path):
    """Return True when branch path exists in canonical export source (or source unavailable)."""
    lookup = _load_supply_branch_path_lookup_from_export()
    info = _SUPPLY_BRANCH_PATH_LOOKUP_SOURCE_INFO or {}
    if str(info.get("status") or "").lower() != "loaded":
        # Do not block exports when source is unavailable.
        return True
    token = _normalize_supply_branch_path_for_lookup(branch_path)
    if not token:
        return False
    return token in lookup


def _classify_supply_root_for_product(product_label):
    """Return the supply root classification ('primary' or 'secondary') for an ESTO product."""
    label = str(product_label or "").strip()
    if not label:
        return "primary"
    mapped = ESTO_PRODUCT_CLASSIFICATION.get(label)
    if mapped in {"primary", "secondary"}:
        return mapped
    # Fallback to prefix-based rule for products not present in ESTO_PRODUCT_LIST.
    return "secondary" if _is_secondary_esto_product(label) else "primary"


def _get_supply_branch_roots_for_entry(fuel_key, fuel_entry):
    """Resolve the LEAP supply branch root(s) for one export fuel entry."""
    entry = fuel_entry or {}
    fuel_name = str(entry.get("fuel_name") or "").strip()
    esto_label = str(entry.get("fuel_label_esto") or "").strip()
    key_label = str(fuel_key or "").strip()

    root_from_export = _resolve_supply_root_from_export_lookup(
        fuel_name,
        esto_label,
        key_label,
    )
    if root_from_export in {"Primary", "Secondary"}:
        return [["Resources", root_from_export]]

    missing_key = _normalize_supply_lookup_fuel_name(fuel_name or esto_label or key_label)
    if missing_key and missing_key not in _SUPPLY_ROOT_LOOKUP_MISS_WARNED:
        _SUPPLY_ROOT_LOOKUP_MISS_WARNED.add(missing_key)
        if SUPPLY_ROOT_CLASSIFICATION_STRICT:
            raise ValueError(
                "Supply root classification missing from LEAP export source for fuel "
                f"'{fuel_name or esto_label or key_label}'. "
                f"Source={SUPPLY_ROOT_CLASSIFICATION_SOURCE_PATH} "
                f"(sheet={SUPPLY_ROOT_CLASSIFICATION_SOURCE_SHEET})."
            )
        print(
            "[WARN] Supply root classification not found in LEAP export source for fuel "
            f"'{fuel_name or esto_label or key_label}'. "
            "Falling back to legacy ESTO-based classification."
        )

    classification = _classify_supply_root_for_product(esto_label or fuel_name or key_label)
    if classification == "secondary":
        return [["Resources", "Secondary"]]
    return [["Resources", "Primary"]]

# MAJOR_SECTOR_CONFIG uses ESTO labels for filtering, but display names can be
# filled from sector_fuel_codes_to_names.xlsx (code_to_name) via mapping below.
#%%

#%%
######### FUNCTIONS #########
def try_debug_breakpoint():
    """Trigger a debug breakpoint when enabled (safe to call anywhere)."""
    if not ENABLE_DEBUG_BREAKPOINTS:
        return
    try:
        breakpoint()
    except Exception as breakpoint_exc:
        print(f"Debug breakpoint failed: {breakpoint_exc}")


def ensure_repo_root():
    """Move to repo root if running from the scrapbook folder."""
    try:
        if os.getcwd().endswith("scrapbook"):
            os.chdir("../../")
    except Exception as exc:
        print(f"Failed to set repo root: {exc}")
        try_debug_breakpoint()
        raise


def load_csv_data(path, label):
    """Load a CSV file and return a pandas DataFrame."""
    try:
        df = read_config_table(path)
        print(f"Loaded {label}: {df.shape[0]} rows, {df.shape[1]} columns")
        return df
    except Exception as exc:
        print(f"Failed to load {label} from {path}: {exc}")
        try_debug_breakpoint()
        raise


def normalize_year_columns(df):
    """Convert year-like columns to int and return (df, year_cols)."""
    try:
        year_cols = [int(col) for col in df.columns if str(col).isdigit()]
        df.columns = [int(col) if str(col).isdigit() else col for col in df.columns]
        return df, year_cols
    except Exception as exc:
        print(f"Failed to normalize year columns: {exc}")
        try_debug_breakpoint()
        raise


def get_years_from(year_cols, base_year):
    """Return a list with the base year column when available."""
    try:
        if base_year in year_cols:
            return [base_year]
        return []
    except Exception as exc:
        print(f"Failed to filter year columns from {base_year}: {exc}")
        try_debug_breakpoint()
        raise


def filter_reference_scenario(df, label):
    """Filter to the reference scenario when a scenarios column is present."""
    try:
        if "scenarios" not in df.columns:
            return df
        scenarios = df["scenarios"].astype(str).str.strip().str.lower()
        filtered = df[scenarios == "reference"].copy()
        unique_vals = sorted(set(scenarios.unique()))
        print(
            f"{label}: filtering to scenarios=reference "
            f"(available={unique_vals}, rows={filtered.shape[0]})"
        )
        return filtered
    except Exception as exc:
        print(f"Failed to filter scenarios for {label}: {exc}")
        try_debug_breakpoint()
        raise


def _extract_numeric_segments(value):
    """Return a list of numeric code segments from a code or label."""
    try:
        if value is None:
            return []
        text = str(value)
        if text == "x":
            return []
        segments = []
        for chunk in text.replace(".", "_").split("_"):
            match = re.match(r"^(\d+)", chunk)
            if not match:
                break
            segments.append(match.group(1).zfill(2))
        return segments
    except Exception as exc:
        print(f"Failed to extract numeric segments from {value}: {exc}")
        try_debug_breakpoint()
        raise


def _match_code_prefix(label_value, code_value):
    """Check if label_value shares the numeric prefix of code_value."""
    try:
        code_segments = _extract_numeric_segments(code_value)
        if not code_segments:
            return False
        label_segments = _extract_numeric_segments(label_value)
        if len(label_segments) < len(code_segments):
            return False
        return label_segments[: len(code_segments)] == code_segments
    except Exception as exc:
        print(f"Failed to match code prefix for {code_value}: {exc}")
        try_debug_breakpoint()
        raise


def select_rows(df, filters):
    """Return filtered rows based on a dict of column -> value."""
    try:
        mask = pd.Series(True, index=df.index)
        for column, value in filters.items():
            if column in df.columns:
                mask &= df[column].eq(value)
                continue
            mask &= False
        return df.loc[mask]
    except Exception as exc:
        print(f"Failed to filter rows with {filters}: {exc}")
        try_debug_breakpoint()
        raise


def select_flow_rows(df, economy, flow_value):
    """Select rows for a flow value using flows or sectors column."""
    try:
        if "flows" in df.columns:
            return select_rows(df, {"economy": economy, "flows": flow_value})
        if "sectors" in df.columns:
            return select_rows(df, {"economy": economy, "sectors": flow_value})
        return df.iloc[0:0]
    except Exception as exc:
        print(f"Failed to select flow rows for {flow_value}: {exc}")
        try_debug_breakpoint()
        raise


def select_fuel_rows(
    df,
    fuel_code_ninth,
    fuel_label_esto,
    fuel_name=None,
    code_to_name_mapping=None,
):
    """Select rows for a fuel using products or fuels/subfuels."""
    try:
        if "products" in df.columns:
            if code_to_name_mapping and fuel_name:
                mapped_products = df["products"].apply(
                    lambda value: map_code_label(value, code_to_name_mapping)
                )
                matched = df[mapped_products.eq(fuel_name)]
                if not matched.empty:
                    return matched
            return df[df["products"].apply(lambda value: _match_code_prefix(value, fuel_label_esto))]
        if "subfuels" in df.columns:
            matched_subfuels = df[
                df["subfuels"].apply(lambda value: _match_code_prefix(value, fuel_code_ninth))
            ]
            if not matched_subfuels.empty:
                return matched_subfuels
            # Some 9th rows carry the canonical fuel code in `fuels` with `subfuels=x`
            # (for example electricity exports). Fall back to fuels in that case.
            if "fuels" in df.columns:
                matched_fuels = df[
                    df["fuels"].apply(lambda value: _match_code_prefix(value, fuel_code_ninth))
                ]
                if not matched_fuels.empty:
                    return matched_fuels
            return matched_subfuels
        if "fuels" in df.columns:
            return df[df["fuels"].apply(lambda value: _match_code_prefix(value, fuel_code_ninth))]
        return df.iloc[0:0]
    except Exception as exc:
        print(f"Failed to select fuel rows: {exc}")
        try_debug_breakpoint()
        raise


def sum_years(df, year_cols):
    """Sum values over year columns, returning a float."""
    try:
        if df.empty:
            return 0.0
        return df[year_cols].sum().sum()
    except Exception as exc:
        print(f"Failed to sum years for frame: {exc}")
        try_debug_breakpoint()
        raise


def print_flow_rows(df, label, year_cols):
    """Print flow rows for debugging."""
    try:
        if not PRINT_FUEL_ROWS:
            return
        if df.empty:
            print(f"{label}: no rows found")
            return
        summary = df.copy()
        summary["total_base_year"] = summary[year_cols].sum(axis=1)
        if PRINT_ONLY_NONZERO_ROWS:
            summary = summary[summary["total_base_year"] != 0]
        if summary.empty:
            print(f"{label}: no nonzero rows after filtering")
            return
        columns_to_show = [
            "scenarios",
            "economy",
            "sectors",
            "flows",
            "fuels",
            "subfuels",
            "products",
            "total_base_year",
        ]
        columns_to_show = [col for col in columns_to_show if col in summary.columns]
        print(f"{label}: rows {summary.shape[0]}")
        print(summary[columns_to_show].head(PRINT_TOP_ROWS).to_string(index=False))
    except Exception as exc:
        print(f"Failed to print flow rows: {exc}")
        try_debug_breakpoint()
        raise


def build_dataset_map(esto_data, esto_year_cols, ninth_data, ninth_year_cols, matt_data, matt_year_cols):
    """Return a dataset map keyed by dataset_key."""
    try:
        return {
            "esto": (esto_data, esto_year_cols),
            "ninth": (ninth_data, ninth_year_cols),
            "matt": (matt_data, matt_year_cols),
        }
    except Exception as exc:
        print(f"Failed to build dataset map: {exc}")
        try_debug_breakpoint()
        raise


def resolve_dataset(dataset_map, dataset_key):
    """Return (data, year_cols) for a dataset key."""
    try:
        if dataset_key not in dataset_map:
            raise KeyError(f"Unknown dataset key: {dataset_key}")
        return dataset_map[dataset_key]
    except Exception as exc:
        print(f"Failed to resolve dataset key {dataset_key}: {exc}")
        try_debug_breakpoint()
        raise


def load_code_to_name_mapping(path_candidates):
    """Load a code-to-name mapping from the first available workbook.

    Inputs:
        path_candidates: List of file paths to check.
    Outputs:
        Dict mapping codebase/label to name.
    Side effects:
        Reads from disk.
    """
    try:
        for path in path_candidates:
            if not config_table_exists(path, sheet_name="code_to_name"):
                continue
            mapping_df = read_config_table(path, sheet_name="code_to_name", dtype=str).fillna("")
            mapping = {}

            if "code" in mapping_df.columns and "name" in mapping_df.columns:
                working = mapping_df.copy()
                if "source_sheet" in working.columns:
                    working["source_sheet"] = working["source_sheet"].fillna("").astype(str)
                    working["source_priority"] = working["source_sheet"].ne("9th").astype(int)
                    working = (
                        working.sort_values("source_priority")
                        .drop_duplicates(subset=["code"], keep="first")
                        .drop(columns=["source_priority"])
                    )
                else:
                    working = working.drop_duplicates(subset=["code"], keep="first")

                mapping.update(
                    dict(
                        zip(
                            working["code"].astype(str).str.strip(),
                            working["name"].astype(str).str.strip(),
                        )
                    )
                )

            if "9th_label" in mapping_df.columns and "name" in mapping_df.columns:
                ninth_labels = mapping_df["9th_label"].astype(str).str.strip()
                names = mapping_df["name"].astype(str).str.strip()
                mapping.update({label: name for label, name in zip(ninth_labels, names) if label})

            if "esto_label" in mapping_df.columns and "name" in mapping_df.columns:
                esto_labels = mapping_df["esto_label"].astype(str).str.strip()
                names = mapping_df["name"].astype(str).str.strip()
                mapping.update({label: name for label, name in zip(esto_labels, names) if label})

            if mapping:
                print(f"Loaded code-to-name mapping from {path}: {len(mapping)} entries")
                return mapping

        print("Code-to-name mapping not found; using labels as-is.")
        return {}
    except Exception as exc:
        print(f"Failed to load code-to-name mapping: {exc}")
        try_debug_breakpoint()
        raise


def map_code_label(label, code_to_name_mapping):
    """Return a label mapped to a human-readable name when available.

    Inputs:
        label: Code or label to map.
        code_to_name_mapping: Dict of codebase/label -> name.
    Outputs:
        Mapped label or original label.
    Side effects:
        None.
    """
    try:
        if not code_to_name_mapping:
            return label
        if label is None:
            return label
        if isinstance(label, float) and pd.isna(label):
            return label
        return code_to_name_mapping.get(str(label), label)
    except Exception as exc:
        print(f"Failed to map label {label}: {exc}")
        try_debug_breakpoint()
        raise


def apply_code_to_name_mapping(major_sector_config, code_to_name_mapping):
    """Apply code-to-name mapping to build display names in sector config.

    Inputs:
        major_sector_config: Dict of sector config entries.
        code_to_name_mapping: Dict of codebase/label -> name.
    Outputs:
        Updated config with fuel_name populated.
    Side effects:
        None.
    """
    try:
        updated = {}
        for fuel_key, fuel_config in major_sector_config.items():
            updated_config = fuel_config.copy()
            mapped_name = map_code_label(fuel_key, code_to_name_mapping)
            if mapped_name == fuel_key:
                mapped_name = updated_config.get("fuel_label_esto", fuel_key)
            updated_config["fuel_name"] = mapped_name
            updated[fuel_key] = updated_config
        return updated
    except Exception as exc:
        print(f"Failed to apply code-to-name mapping: {exc}")
        try_debug_breakpoint()
        raise


def get_economy_list(df, requested_economies=None):
    """Return a list of economies to analyze."""
    try:
        available = sorted(df["economy"].dropna().unique())
        if requested_economies:
            requested = [econ for econ in requested_economies if econ in available]
            missing = [econ for econ in requested_economies if econ not in available]
            if missing:
                print(
                    "Warning: requested economies not found in this dataset: "
                    f"{', '.join(missing)}"
                )
            if requested:
                return requested
            print("Warning: no requested economies found; using all available economies.")
        return available
    except Exception as exc:
        print(f"Failed to build economy list: {exc}")
        try_debug_breakpoint()
        raise


def add_all_economy_total(df, year_cols, economy_label="ALL"):
    """Append an all-economy total row set to a dataset."""
    try:
        if "economy" not in df.columns or df.empty:
            return df
        if df["economy"].astype(str).eq(economy_label).any():
            return df
        group_cols = [
            col for col in df.columns if col not in year_cols and col != "economy"
        ]
        totals = (
            df.groupby(group_cols, dropna=False)[year_cols]
            .sum()
            .reset_index()
        )
        totals["economy"] = economy_label
        totals = totals[df.columns.tolist()]
        return pd.concat([df, totals], ignore_index=True)
    except Exception as exc:
        print(f"Failed to add all-economy totals: {exc}")
        try_debug_breakpoint()
        raise


def summarize_supply_for_fuel(
    data,
    year_cols,
    economy,
    fuel_config,
    flow_codes,
    base_year,
    code_to_name_mapping=None,
):
    """Print import/export totals for a fuel and economy."""
    try:
        year_cols_from_base = get_years_from(year_cols, base_year)
        display_name = fuel_config.get("fuel_name", fuel_config["fuel_label_esto"])
        fuel_rows = select_fuel_rows(
            data,
            fuel_config["fuel_code_ninth"],
            fuel_config["fuel_label_esto"],
            fuel_name=fuel_config.get("fuel_name"),
            code_to_name_mapping=code_to_name_mapping,
        )
        if fuel_rows.empty:
            print(f"{display_name}: no fuel rows found")
            return

        imports_rows = select_flow_rows(fuel_rows, economy, flow_codes["imports"])
        exports_rows = select_flow_rows(fuel_rows, economy, flow_codes["exports"])

        print_flow_rows(imports_rows, f"{economy} imports", year_cols_from_base)
        print_flow_rows(exports_rows, f"{economy} exports", year_cols_from_base)

        imports_total = sum_years(imports_rows, year_cols_from_base)
        exports_total = sum_years(exports_rows, year_cols_from_base)

        print(
            f"{economy} {display_name} (base {base_year}): "
            f"imports {imports_total:.3f}, exports {exports_total:.3f}"
        )
    except Exception as exc:
        print(f"Failed to summarize supply for {fuel_config}: {exc}")
        try_debug_breakpoint()
        raise
#%%

#%%
def list_unique_fuels_and_products(ninth_data, esto_data):
    """Print unique fuel/subfuel combos (9th) and products (ESTO)."""
    try:
        if "fuels" in ninth_data.columns and "subfuels" in ninth_data.columns:
            fuels = (
                ninth_data[["fuels", "subfuels"]]
                .dropna()
                .drop_duplicates()
                .sort_values(["fuels", "subfuels"])
            )
            fuel_pairs = list(fuels.itertuples(index=False, name=None))
            print(f"9th fuel/subfuel combos: {len(fuel_pairs)}")
            for fuel, subfuel in fuel_pairs:
                print(f"- {fuel} / {subfuel}")
        else:
            print("9th data missing fuels/subfuels columns")

        if "products" in esto_data.columns:
            products = (
                esto_data[["products"]]
                .dropna()
                .drop_duplicates()
                .sort_values(["products"])
            )
            product_list = products["products"].astype(str).tolist()
            print(f"ESTO products: {len(product_list)}")
            for product in product_list:
                print(f"- {product}")
        else:
            print("ESTO data missing products column")
    except Exception as exc:
        print(f"Failed to list unique fuels/products: {exc}")
        try_debug_breakpoint()
        raise
#%%

#%%
def find_first_existing_file(path_candidates):
    """Return the first path that exists from the provided candidates."""
    try:
        for path in path_candidates:
            if config_table_exists(path):
                return path
        print(f"No code-to-name workbook found in {path_candidates}")
        return None
    except Exception as exc:
        print(f"Failed to locate workbook from {path_candidates}: {exc}")
        try_debug_breakpoint()
        raise


def build_supply_sector_config(
    code_to_name_paths,
    exclude_prefixes=None,
    dataset_key="esto",
):
    """Build sector config entries for every ESTO fuel product."""
    workbook_path = None
    for path in code_to_name_paths:
        if config_table_exists(path, sheet_name="ESTO"):
            workbook_path = path
            break
    if not workbook_path:
        print(
            "Warning: code-to-name workbook is missing; supply export will run with an empty sector config."
        )
        return {}
    try:
        df = read_config_table(workbook_path, sheet_name="ESTO", dtype=str)
        df = df[df["products"].notna()].copy()
        df["products"] = df["products"].astype(str).str.strip()
        if exclude_prefixes:
            mask = ~df["products"].str.startswith(tuple(exclude_prefixes), na=False)
            df = df[mask]
        mapping_df = read_config_table(workbook_path, sheet_name="code_to_name", dtype=str).fillna("")
        lookup = {
            str(row.get("esto_label") or "").strip(): row.to_dict()
            for _, row in mapping_df.iterrows()
            if str(row.get("esto_label") or "").strip()
        }

        config = {}
        for product in sorted(df["products"].unique()):
            entry = lookup.get(product, {})
            config[product] = {
                "dataset_key": dataset_key,
                "fuel_label_esto": product,
                "fuel_code_ninth": entry.get("9th_label") or None,
                "fuel_name": entry.get("name") or product,
            }
        print(
            f"Built supply config for {len(config)} ESTO products "
            f"(excluding prefixes {exclude_prefixes})."
        )
        return config
    except Exception as exc:
        print(f"Failed to build supply sector config: {exc}")
        try_debug_breakpoint()
        raise


def sanitize_leap_label(value):
    """Make fuel/branch labels safe for LEAP imports."""
    try:
        if value is None:
            return value
        return sanitize_leap_name(str(value))
    except Exception as exc:
        print(f"Failed to sanitize label {value}: {exc}")
        try_debug_breakpoint()
        raise


def build_branch_path(parts):
    """Return a LEAP branch path from a list of parts."""
    try:
        return "\\".join([str(part) for part in parts if part])
    except Exception as exc:
        print(f"Failed to build branch path from {parts}: {exc}")
        try_debug_breakpoint()
        raise


def coerce_value_by_year(value, base_year, final_year):
    """Convert a scalar, dict, or Series into a year->value dict."""
    try:
        if isinstance(value, dict):
            return {int(year): float(val) for year, val in value.items()}
        if isinstance(value, pd.Series):
            return {int(year): float(val) for year, val in value.items()}
        return {
            int(year): float(value if value not in (None, "") else 0.0)
            for year in range(base_year, final_year + 1)
        }
    except Exception as exc:
        print(f"Failed to coerce value to year mapping: {exc}")
        try_debug_breakpoint()
        raise


def build_year_rows(
    branch_path,
    measure,
    scenario,
    value_by_year,
    units,
    scale,
    per_value,
):
    """Return log-style rows for a LEAP import file."""
    try:
        rows = []
        for year, value in sorted(value_by_year.items()):
            safe_value = 0.0 if value is None else float(value)
            rows.append(
                {
                    "Branch_Path": branch_path,
                    "Scenario": scenario,
                    "Measure": measure,
                    "Units": units,
                    "Scale": scale,
                    "Per...": per_value,
                    "Date": int(year),
                    "Value": safe_value,
                }
            )
        return rows
    except Exception as exc:
        print(f"Failed to build year rows for {branch_path}: {exc}")
        try_debug_breakpoint()
        raise


def _normalize_override_year_map(
    value_by_year,
    base_year,
    final_year,
):
    """Return a clipped year->float mapping from an override payload."""
    if value_by_year is None:
        return None
    normalized = coerce_value_by_year(value_by_year, base_year, final_year)
    return {
        int(year): float(normalized.get(year, 0.0))
        for year in range(base_year, final_year + 1)
    }


def _resolve_supply_override(
    flow_value_overrides,
    scenario,
    fuel_key,
    fuel_entry,
    flow_key,
    base_year,
    final_year,
):
    """Return an override year map for a scenario/fuel/flow when present."""
    if not isinstance(flow_value_overrides, dict):
        return None
    scenario_payload = flow_value_overrides.get(scenario)
    if not isinstance(scenario_payload, dict):
        scenario_payload = flow_value_overrides.get(str(scenario))
    scenario_key = str(scenario or "").strip().lower()
    lower_payloads = {
        str(key).strip().lower(): value
        for key, value in flow_value_overrides.items()
        if isinstance(value, dict)
    }
    if not isinstance(scenario_payload, dict) and scenario_key:
        scenario_payload = lower_payloads.get(scenario_key)
    if (
        not isinstance(scenario_payload, dict)
        and scenario_key in {"current accounts", "current account"}
    ):
        # Keep Current Accounts aligned with reset behavior by falling back
        # to Reference overrides when no explicit CA payload exists.
        scenario_payload = lower_payloads.get("reference")
    if not isinstance(scenario_payload, dict):
        return None

    candidate_keys = [
        fuel_key,
        fuel_entry.get("fuel_label_esto"),
        fuel_entry.get("fuel_name"),
    ]
    fuel_payload = None
    for candidate in candidate_keys:
        if candidate in scenario_payload and isinstance(scenario_payload[candidate], dict):
            fuel_payload = scenario_payload[candidate]
            break
    if not isinstance(fuel_payload, dict):
        return None
    return _normalize_override_year_map(
        fuel_payload.get(flow_key),
        base_year,
        final_year,
    )


def get_flow_total_for_fuel(
    data,
    year_cols,
    base_year,
    economy,
    fuel_config,
    flow_key,
    flow_value,
    code_to_name_mapping=None,
):
    """Sum the base-year value for a flow/fuel/economy combination and normalize the sign."""
    try:
        if flow_value is None:
            return 0.0
        if base_year not in year_cols:
            print(f"Warning: base year {base_year} missing for economy {economy}")
            return 0.0
        fuel_rows = select_fuel_rows(
            data,
            fuel_config.get("fuel_code_ninth"),
            fuel_config["fuel_label_esto"],
            fuel_name=fuel_config.get("fuel_name"),
            code_to_name_mapping=code_to_name_mapping,
        )
        flow_rows = select_flow_rows(fuel_rows, economy, flow_value)
        total = sum_years(flow_rows, [base_year])
        return normalize_supply_flow_total(flow_key, total)
    except Exception as exc:
        print(f"Failed to sum flow {flow_key} for fuel {fuel_config}: {exc}")
        try_debug_breakpoint()
        raise


def get_flow_series_for_fuel(
    data,
    year_cols,
    economy,
    fuel_config,
    flow_key,
    flow_value,
    base_year,
    final_year,
    code_to_name_mapping=None,
):
    """Return year->value from source rows for one flow/fuel/economy."""
    try:
        if flow_value is None:
            return {year: 0.0 for year in range(base_year, final_year + 1)}
        fuel_rows = select_fuel_rows(
            data,
            fuel_config.get("fuel_code_ninth"),
            fuel_config["fuel_label_esto"],
            fuel_name=fuel_config.get("fuel_name"),
            code_to_name_mapping=code_to_name_mapping,
        )
        flow_rows = select_flow_rows(fuel_rows, economy, flow_value)
        out = {year: 0.0 for year in range(base_year, final_year + 1)}
        if flow_rows.empty:
            return out
        for year in range(base_year, final_year + 1):
            if year not in year_cols:
                continue
            value = pd.to_numeric(flow_rows[year], errors="coerce").fillna(0.0).sum()
            out[year] = normalize_supply_flow_total(flow_key, float(value))
        return out
    except Exception as exc:
        print(
            f"Failed to build flow series for {economy}, {fuel_config}, {flow_key}: {exc}"
        )
        try_debug_breakpoint()
        raise


def _get_projection_series(
    projection_lookup,
    economy,
    flow_value,
    product_value,
    projection_years,
):
    """Return a projection series for an ESTO flow/product pair."""
    if projection_lookup is None or not projection_years:
        return {year: 0.0 for year in projection_years}
    econ_key = normalize_economy_key(economy)
    key = (econ_key, str(flow_value).strip(), str(product_value).strip())
    if key not in projection_lookup.index:
        return {year: 0.0 for year in projection_years}
    row = projection_lookup.loc[key]
    if isinstance(row, pd.DataFrame):
        row = row.sum()
    return {year: float(row.get(year, 0.0)) for year in projection_years}


def build_supply_value_by_year(
    data,
    year_cols,
    economy,
    fuel_config,
    flow_key,
    flow_value,
    base_year,
    final_year,
    projection_lookup=None,
    projection_years=None,
    code_to_name_mapping=None,
):
    """Return a full year mapping for one economy/fuel/flow.

    - For 9th-style datasets (sector-coded), read all years directly from source rows.
    - For ESTO-style datasets, use base-year source plus projected years from lookup.
    """
    dataset_is_ninth_style = "sectors" in data.columns and "flows" not in data.columns
    if dataset_is_ninth_style:
        return get_flow_series_for_fuel(
            data,
            year_cols,
            economy,
            fuel_config,
            flow_key,
            flow_value,
            base_year,
            final_year,
            code_to_name_mapping=code_to_name_mapping,
        )

    projection_years = [
        year for year in (projection_years or []) if year <= final_year
    ]
    base_value = get_flow_total_for_fuel(
        data,
        year_cols,
        base_year,
        economy,
        fuel_config,
        flow_key,
        flow_value,
        code_to_name_mapping=code_to_name_mapping,
    )
    projected = _get_projection_series(
        projection_lookup,
        economy,
        flow_value,
        fuel_config["fuel_label_esto"],
        projection_years,
    )
    if is_output_flow(flow_key):
        projected = {year: abs(value) for year, value in projected.items()}
    value_by_year = {year: 0.0 for year in range(base_year, final_year + 1)}
    value_by_year[base_year] = base_value
    for year, value in projected.items():
        value_by_year[int(year)] = float(value)
    if is_output_flow(flow_key):
        value_by_year = {
            year: abs(value) if value is not None else 0.0
            for year, value in value_by_year.items()
        }
    return value_by_year


def get_region_for_economy(economy_code):
    """Return the LEAP region name that should be used for an economy."""
    try:
        return EXPORT_ECONOMY_REGION_OVERRIDES.get(economy_code, EXPORT_REGION)
    except Exception as exc:
        print(f"Failed to resolve region for {economy_code}: {exc}")
        try_debug_breakpoint()
        raise


def format_scenario_label_for_filename(scenarios):
    """Return a filename-friendly scenario string."""
    try:
        sanitized = "_".join(
            "".join(ch for ch in scenario if ch.isalnum())
            for scenario in scenarios
        )
        return sanitized or "scenarios"
    except Exception as exc:
        print(f"Failed to build filename-safe scenario label: {exc}")
        try_debug_breakpoint()
        raise


def build_supply_log_rows(
    data,
    year_cols,
    economy,
    fuel_config,
    flow_codes,
    scenario_names,
    base_year,
    final_year,
    code_to_name_mapping=None,
    projection_lookup=None,
    projection_years=None,
    flow_value_overrides=None,
    supply_measures=None,
):
    """Build log entries for supply measures per fuel."""
    try:
        if not fuel_config:
            print("Warning: no supply fuels available for export.")
            return []
        measures = supply_measures if isinstance(supply_measures, list) and supply_measures else SUPPLY_MEASURES
        rows = []
        for fuel_key in sorted(fuel_config):
            entry = fuel_config[fuel_key]
            display_name = entry.get("fuel_name") or entry["fuel_label_esto"]
            safe_name = sanitize_leap_label(display_name)
            branch_roots = _get_supply_branch_roots_for_entry(fuel_key, entry)
            required_flow_keys = {
                str(measure.get("flow_key") or "").strip()
                for measure in measures
                if str(measure.get("flow_key") or "").strip()
            }
            default_flow_values_by_year = {}
            for flow_key in sorted(required_flow_keys):
                source_flow_key = "production" if flow_key == "max_production" else flow_key
                flow_value = flow_codes.get(source_flow_key)
                default_flow_values_by_year[flow_key] = build_supply_value_by_year(
                    data,
                    year_cols,
                    economy,
                    entry,
                    source_flow_key,
                    flow_value,
                    base_year,
                    final_year,
                    projection_lookup=projection_lookup,
                    projection_years=projection_years,
                    code_to_name_mapping=code_to_name_mapping,
                )
            for scenario in scenario_names:
                for branch_root in branch_roots:
                    branch_path = build_branch_path(branch_root + [safe_name])
                    if not _supply_branch_exists_in_export_source(branch_path):
                        miss_key = f"{economy}|{scenario}|{branch_path}"
                        if miss_key not in _SUPPLY_BRANCH_PATH_MISS_WARNED:
                            _SUPPLY_BRANCH_PATH_MISS_WARNED.add(miss_key)
                            print(
                                "[WARN] Skipping supply export row for branch not present in "
                                "canonical full-model export source: "
                                f"{branch_path} (economy={economy}, scenario={scenario}, fuel={display_name})"
                            )
                        continue
                    branch_type = str(branch_root[-1] if branch_root else "").strip().lower()
                    for measure in measures:
                        root_filter = str(measure.get("branch_root") or "").strip().lower()
                        if root_filter and root_filter not in {"all", branch_type}:
                            continue
                        flow_key = measure.get("flow_key")
                        if flow_key:
                            override_value_by_year = _resolve_supply_override(
                                flow_value_overrides,
                                scenario,
                                fuel_key,
                                entry,
                                flow_key,
                                base_year,
                                final_year,
                            )
                            value_by_year = override_value_by_year or default_flow_values_by_year.get(
                                flow_key, {year: 0.0 for year in range(base_year, final_year + 1)}
                            )
                        else:
                            value_by_year = coerce_value_by_year(
                                measure.get("value", 0.0), base_year, final_year
                            )
                        rows.extend(
                            build_year_rows(
                                branch_path,
                                measure["name"],
                                scenario,
                                value_by_year,
                                measure["units"],
                                "",
                                measure["per"],
                            )
                        )
        return rows
    except Exception as exc:
        print(f"Failed to build supply log rows for {economy}: {exc}")
        try_debug_breakpoint()
        raise

#%%
######### WORKFLOW CONTROLS #########
RUN_SUPPLY_ANALYSIS = workflow_cfg.SUPPLY_RUN_SUPPLY_ANALYSIS
RUN_LIST_FUELS = workflow_cfg.SUPPLY_RUN_LIST_FUELS
ALL_ECONOMY_LABEL = workflow_cfg.SUPPLY_ALL_ECONOMY_LABEL
ECONOMIES_TO_ANALYZE = list(workflow_cfg.SUPPLY_ECONOMIES_TO_ANALYZE)
SAVE_ESTO_SUBTOTAL_LABELED = workflow_cfg.SUPPLY_SAVE_ESTO_SUBTOTAL_LABELED
ESTO_SUBTOTAL_LABELED_OUTPUT_PATH = workflow_cfg.SUPPLY_ESTO_SUBTOTAL_LABELED_OUTPUT_PATH
EXPORT_DATASET_KEY = workflow_cfg.SUPPLY_EXPORT_DATASET_KEY
EXPORT_DIR = workflow_cfg.SUPPLY_EXPORT_DIR
EXPORT_FILE_NAME = workflow_cfg.SUPPLY_EXPORT_FILE_NAME
SCENARIO_TO_RUN = workflow_cfg.SUPPLY_SCENARIO_TO_RUN
FILL_BRANCHES_FROM_EXPORT_FILE = workflow_cfg.SUPPLY_FILL_BRANCHES_FROM_EXPORT_FILE
HANDLE_CURRENT_ACCOUNTS_TOO = workflow_cfg.SUPPLY_HANDLE_CURRENT_ACCOUNTS_TOO
RUN_SUPPLY_LEAP_IMPORT = workflow_cfg.SUPPLY_RUN_SUPPLY_LEAP_IMPORT
SHEET_NAME = workflow_cfg.SUPPLY_SHEET_NAME
#%%


def prepare_supply_assets(
    economies: Iterable[str] | None = None,
    aggregate_economy_label: str | None = None,
    save_subtotal_labeled: bool = SAVE_ESTO_SUBTOTAL_LABELED,
    subtotal_output_path: str = ESTO_SUBTOTAL_LABELED_OUTPUT_PATH,
):
    """Load the supply datasets and build the required mappings."""
    ensure_repo_root()
    sector_config = build_supply_sector_config(
        CODE_TO_NAME_PATHS,
        exclude_prefixes=EXCLUDED_ESTO_PREFIXES,
    )
    code_to_name_mapping = (
        load_code_to_name_mapping(CODE_TO_NAME_PATHS) if USE_CODE_TO_NAME_MAPPING else {}
    )
    if code_to_name_mapping:
        sector_config = apply_code_to_name_mapping(
            sector_config, code_to_name_mapping
        )

    workflow_common.archive_config_dir_once_per_day()
    esto_data_raw, ninth_data_raw = load_augmented_reference_tables(
        esto_path=MATT_DATA_PATH,
        ninth_path=ESTO_DATA_PATH,
        subtotal_mapping_path=SUBTOTAL_MAPPING_PATH,
        synthetic_rules_path=CONFIG_DIR / "synthetic_reference_rows.csv",
        cache_dir=REFERENCE_CACHE_DIR,
        apply_esto_subtotal_map=True,
        filter_esto_subtotals_flag=False,
        filter_ninth_subtotals_flag=False,
    )
    print(
        f"Loaded ESTO data (augmented): {esto_data_raw.shape[0]} rows, {esto_data_raw.shape[1]} columns"
    )
    print(
        f"Loaded 9th data (augmented): {ninth_data_raw.shape[0]} rows, {ninth_data_raw.shape[1]} columns"
    )
    ninth_data_raw, ninth_year_cols = normalize_year_columns(ninth_data_raw)
    esto_data_raw, esto_year_cols = normalize_year_columns(esto_data_raw)

    ninth_data = filter_reference_scenario(ninth_data_raw, "9th data")
    if "subtotal_results" in ninth_data.columns:
        ninth_data = ninth_data[ninth_data["subtotal_results"] == False].copy()
    esto_data_with_subtotals = apply_matt_subtotal_mapping(
        esto_data_raw, SUBTOTAL_MAPPING_PATH
    )
    # if save_subtotal_labeled:
    #     save_subtotal_labeled_data(
    #         esto_data_with_subtotals,
    #         subtotal_output_path,
    #         "ESTO (Matt) data",
    #     )
    esto_data = filter_matt_subtotals(esto_data_with_subtotals)

    economy_list = workflow_common.normalize_economies(economies or ECONOMIES_TO_ANALYZE)
    should_aggregate, aggregate_label, _ = workflow_common.resolve_aggregate_economy(
        economy_list,
        aggregate_label=aggregate_economy_label or ALL_ECONOMY_LABEL,
    )
    if should_aggregate:
        ninth_data = add_all_economy_total(
            ninth_data, ninth_year_cols, aggregate_label
        )
        esto_data = add_all_economy_total(
            esto_data, esto_year_cols, aggregate_label
        )

    projection_df, projection_diagnostics = build_esto_projection_table(
        ninth_data=ninth_data,
        esto_data=esto_data,
        mapping_path=NINTH_TO_ESTO_MAPPING_PATH,
        base_year=BASE_YEAR,
        projection_years=PROJECTION_YEAR_RANGE,
    )
    global SUPPLY_PROJECTION_LOOKUP
    SUPPLY_PROJECTION_LOOKUP = build_projection_lookup(projection_df)
    if SAVE_PROJECTION_DIAGNOSTICS and projection_diagnostics is not None:
        if not projection_diagnostics.empty:
            PROJECTION_DIAGNOSTICS_PATH.parent.mkdir(parents=True, exist_ok=True)
            projection_diagnostics.to_csv(PROJECTION_DIAGNOSTICS_PATH, index=False)
            print(f"Saved projection fallback report to {PROJECTION_DIAGNOSTICS_PATH}")

    dataset_map = build_dataset_map(
        esto_data,
        esto_year_cols,
        ninth_data,
        ninth_year_cols,
        esto_data_raw,
        esto_year_cols,
    )
    return dataset_map, sector_config, code_to_name_mapping, ninth_data, esto_data


def generate_supply_exports(
    dataset_map,
    fuel_config,
    code_to_name_mapping,
    projection_lookup=None,
    projection_years=None,
    dataset_key: str = EXPORT_DATASET_KEY,
    economies: list[str] | None = None,
    scenario_names=EXPORT_SCENARIOS,
    base_year=EXPORT_BASE_YEAR,
    final_year=EXPORT_FINAL_YEAR,
    export_output_dir: Path | str = EXPORT_OUTPUT_DIR,
    filename_template: str = EXPORT_FILENAME_TEMPLATE,
    flow_value_overrides_by_economy: dict | None = None,
    supply_measures: list[dict] | None = None,
    keep_all_zero_rows: bool = False,
):
    """Generate LEAP-ready supply exports for the requested economies."""
    data, year_cols = resolve_dataset(dataset_map, dataset_key)
    flow_codes = FLOW_CODES_BY_DATASET.get(dataset_key)
    if not flow_codes:
        raise KeyError(f"Unknown dataset key for flow codes: {dataset_key}")
    if projection_lookup is None:
        projection_lookup = SUPPLY_PROJECTION_LOOKUP
    target_economies = economies or get_economy_list(data, ECONOMIES_TO_ANALYZE)
    scenario_label = ", ".join(scenario_names)
    scenario_filename = format_scenario_label_for_filename(scenario_names)
    saved_exports: list[tuple[str, Path]] = []

    for economy in target_economies:
        economy_flow_overrides = None
        if isinstance(flow_value_overrides_by_economy, dict):
            economy_flow_overrides = flow_value_overrides_by_economy.get(economy)
        log_rows = build_supply_log_rows(
            data,
            year_cols,
            economy,
            fuel_config,
            flow_codes,
            scenario_names,
            base_year,
            final_year,
            code_to_name_mapping=code_to_name_mapping,
            projection_lookup=projection_lookup,
            projection_years=projection_years,
            flow_value_overrides=economy_flow_overrides,
            supply_measures=supply_measures,
        )
        if not log_rows:
            print(f"No supply rows generated for {economy}")
            continue
        log_df = pd.DataFrame(log_rows)
        region_name = get_region_for_economy(economy)
        export_df = finalise_export_df(
            log_df, scenario_label, region_name, base_year, final_year
        )
        if export_df is None:
            print(f"Skipping export for {economy} because no data survived pivot.")
            continue
        year_columns = [
            column for column in export_df.columns if isinstance(column, int)
        ]
        if year_columns and not keep_all_zero_rows:
            numeric_years = (
                export_df[year_columns]
                .apply(pd.to_numeric, errors="coerce")
                .fillna(0.0)
            )
            nonzero_mask = numeric_years.abs().sum(axis=1) > 0.0
            dropped_count = int((~nonzero_mask).sum())
            if dropped_count:
                print(
                    f"[INFO] Dropping {dropped_count} all-zero supply rows from export for {economy}."
                )
            export_df = export_df.loc[nonzero_mask].copy()
        if export_df.empty:
            print(
                f"Skipping export for {economy} because all supply rows are zero after filtering."
            )
            continue
        os.makedirs(export_output_dir, exist_ok=True)
        export_path = Path(export_output_dir) / filename_template.format(
            economy=economy, scenarios=scenario_filename
        )
        save_export_files(
            export_df,
            export_df,
            export_path,
            base_year,
            final_year,
            EXPORT_MODEL_NAME,
        )
        saved_exports.append((economy, export_path))
        print(f"Saved supply LEAP import for {economy} at {export_path}")

    return saved_exports


def run_supply_pipeline(
    run_list_fuels: bool = RUN_LIST_FUELS,
    run_supply_analysis: bool = RUN_SUPPLY_ANALYSIS,
    dataset_key: str = EXPORT_DATASET_KEY,
    economies: list[str] | None = None,
):
    """Orchestrate the supply export analysis workflow."""
    assets = None
    if run_list_fuels or run_supply_analysis:
        assets = prepare_supply_assets(economies=economies)
    if run_list_fuels and assets:
        _, _, _, ninth_data, esto_data = assets
        list_unique_fuels_and_products(ninth_data, esto_data)

    export_paths: list[Path] = []
    if run_supply_analysis and assets:
        dataset_map, sector_config, code_to_name_mapping, _, _ = assets
        exports = generate_supply_exports(
            dataset_map,
            sector_config,
            code_to_name_mapping,
            projection_years=PROJECTION_YEAR_RANGE,
            dataset_key=dataset_key,
            economies=economies,
        )
        export_paths = [path for _, path in exports]
    return export_paths


def _normalize_token(token: str) -> str:
    """Return a lowercase alphanumeric key suitable for scenario matching."""
    return "".join(ch.lower() for ch in token if ch.isalnum())


def _match_scenario_token(token: str) -> str | None:
    """Match a token from the filename to the configured scenario dictionary."""
    token_key = _normalize_token(token)
    for scenario in scenario_dict:
        if token_key == _normalize_token(scenario):
            return scenario
    return None


def locate_supply_export(directory: Path, filename: str | None = None) -> Path:
    """Return the most recent supply export, optionally using an explicit name."""
    try:
        if filename:
            candidate = directory / filename
            if candidate.exists():
                return candidate
            raise FileNotFoundError(f"Expected supply export not found: {candidate}")
        matches = sorted(directory.glob("supply_leap_imports_*.xlsx"))
        if not matches:
            raise FileNotFoundError(
                f"No supply export files detected in {directory}"
            )
        return matches[-1]
    except Exception as exc:
        print(f"[ERROR] Unable to locate supply export: {exc}")
        try_debug_breakpoint()
        raise


def extract_export_metadata(export_path: Path) -> list[str]:
    """Parse the export filename to recover declared scenario tokens."""
    match = EXPORT_FILENAME_REGEX.match(export_path.name)
    if not match:
        raise ValueError(
            f"Supply export filename '{export_path.name}' does not match the expected pattern."
        )
    tokens = [tok for tok in match.group("scenarios").split("_") if tok]
    normalized = []
    for token in tokens:
        label = token.replace("-", " ").strip()
        scenario_name = _match_scenario_token(label)
        normalized.append(scenario_name or label)
    return normalized


def _read_unique_column(
    export_path: Path, column: str, sheet_name: str = SHEET_NAME
) -> list[str]:
    """Read a single column from the export and preserve the order of unique values."""
    try:
        df = read_config_table(export_path, sheet_name=sheet_name, header=2, usecols=[column])
    except Exception as exc:
        print(f"[ERROR] Failed to read column '{column}' from {export_path}: {exc}")
        try_debug_breakpoint()
        raise
    values = df[column].dropna().astype(str).tolist()
    seen = []
    for value in values:
        if value not in seen:
            seen.append(value)
    return seen


def get_available_scenarios(export_path: Path) -> list[str]:
    """Return the scenario labels present in the export workbook."""
    return _read_unique_column(export_path, "Scenario")


def ensure_region_in_export(export_path: Path, region: str) -> None:
    """Raise if the configured region does not appear in the export file."""
    regions = _read_unique_column(export_path, "Region")
    if region not in regions:
        raise ValueError(
            f"Region '{region}' not found in export file; available values: {regions}"
        )


def _extract_fuel_from_branch_path(branch_path: str) -> str | None:
    """Return the final segment of a branch path when it represents a fuel name."""
    components = [segment.strip() for segment in branch_path.split("\\") if segment.strip()]
    if len(components) < 3:
        return None
    return components[-1]


def get_supply_fuels_from_export(export_path: Path) -> list[str]:
    """Read branch paths from the export to recover fuel names."""
    branch_paths = _read_unique_column(export_path, "Branch Path")
    fuels: list[str] = []
    seen: set[str] = set()
    for branch_path in branch_paths:
        fuel = _extract_fuel_from_branch_path(branch_path)
        if not fuel or fuel in seen:
            continue
        seen.add(fuel)
        fuels.append(fuel)
    return fuels


def ensure_supply_fuel_exists(
    L,
    fuel_name: str,
    copy_from: str | None = None,
    fuel_state: int = 2,
) -> object:
    """Wrapper around LEAP's ensure_fuel_exists so fuels appear before the fill."""
    return ensure_fuel_exists(
        L,
        fuel_name,
        copy_from=copy_from,
        fuel_state=fuel_state,
    )


def ensure_supply_fuels_from_export(L, export_path: Path) -> None:
    """Ensure every fuel referenced in the export exists in LEAP before filling."""
    try:
        fuels = get_supply_fuels_from_export(export_path)
    except Exception as exc:
        print(f"[ERROR] Unable to determine supply fuels from export: {exc}")
        try_debug_breakpoint()
        raise
    if not fuels:
        print("[INFO] No supply fuels detected in export; nothing to ensure.")
        return
    print(f"[INFO] Ensuring {len(fuels)} supply fuel(s) exist before branch fill.")
    for fuel in fuels:
        ensure_supply_fuel_exists(L, fuel_name=fuel)


def run_branch_fill(
    L,
    export_path: Path,
    scenario: str,
    region: str,
    handle_current_accounts: bool,
    raise_on_missing_branch: bool = True,
) -> None:
    """Load data into supply branches from the export workbook."""
    try:
        outcome = fill_branches_from_export_file(
            L,
            export_path,
            sheet_name=SHEET_NAME,
            scenario=scenario,
            region=region,
            RAISE_ERROR_ON_FAILED_SET=raise_on_missing_branch,
            SET_UNITS=True,
            HANDLE_CURRENT_ACCOUNTS_TOO=handle_current_accounts,
            RUN_FUEL_CATALOG_PREFLIGHT=False,
        )
        print(f"[INFO] Supply branch fill result: {outcome}")
        _print_supply_missing_both_primary_secondary_summary(outcome)
    except Exception as exc:
        print(f"[ERROR] Supply branch fill failed: {exc}")
        try_debug_breakpoint()
        raise


def _print_supply_missing_both_primary_secondary_summary(outcome: dict | None) -> None:
    """Print fuels that failed in all attempted Resources roots during branch fill."""
    if not isinstance(outcome, dict):
        return

    def _extract_root_and_fuel(branch_path: str | None) -> tuple[str | None, str | None]:
        parts = [part.strip() for part in str(branch_path or "").split("\\") if part and str(part).strip()]
        if len(parts) < 3:
            return None, None
        if parts[0].lower() != "resources":
            return None, None
        root = parts[1].strip()
        if root.lower() not in {"primary", "secondary"}:
            return None, None
        fuel = parts[2].strip()
        if not fuel:
            return None, None
        return root.title(), fuel

    success_roots_by_fuel: dict[str, set[str]] = {}
    failed_roots_by_fuel: dict[str, set[str]] = {}

    for branch_path, variable in outcome.get("success", []) or []:
        var_name = str(variable or "").strip().lower()
        if var_name not in {"imports", "exports"}:
            continue
        root, fuel = _extract_root_and_fuel(branch_path)
        if not root or not fuel:
            continue
        success_roots_by_fuel.setdefault(fuel, set()).add(root)

    for branch_path, variable in outcome.get("failed", []) or []:
        var_name = str(variable or "").strip().lower()
        if var_name not in {"imports", "exports"}:
            continue
        root, fuel = _extract_root_and_fuel(branch_path)
        if not root or not fuel:
            continue
        failed_roots_by_fuel.setdefault(fuel, set()).add(root)

    fuels_missing_all_attempted: list[str] = []
    for fuel, failed_roots in sorted(failed_roots_by_fuel.items()):
        if success_roots_by_fuel.get(fuel):
            continue
        if failed_roots:
            fuels_missing_all_attempted.append(
                f"{fuel} (attempted: {', '.join(sorted(failed_roots))})"
            )

    if fuels_missing_all_attempted:
        print(
            "[WARN] Supply fuels not found in attempted Resources root(s) "
            f"({len(fuels_missing_all_attempted)}): {', '.join(fuels_missing_all_attempted)}"
        )
    else:
        print(
            "[INFO] Supply branch lookup summary: no fuels were missing in their "
            "attempted Resources root(s)."
        )


def run_supply_leap_import(
    export_directory: Path = EXPORT_DIR,
    filename: str | None = EXPORT_FILE_NAME,
    scenario_to_run: str = SCENARIO_TO_RUN,
    region: str = EXPORT_REGION,
    handle_current_accounts: bool = HANDLE_CURRENT_ACCOUNTS_TOO,
    fill_branches: bool = FILL_BRANCHES_FROM_EXPORT_FILE,
) -> Path:
    """Locate the supply export and optionally fill the matching LEAP branches."""
    export_path = locate_supply_export(export_directory, filename)
    declared_scenarios = extract_export_metadata(export_path)
    available_scenarios = get_available_scenarios(export_path)
    print(
        f"[INFO] Preparing supply import from '{export_path.name}', declared scenarios "
        f"{declared_scenarios}, available scenarios {available_scenarios}."
    )
    if scenario_to_run not in available_scenarios:
        raise ValueError(
            f"Desired scenario '{scenario_to_run}' not present; available: {available_scenarios}"
        )
    ensure_region_in_export(export_path, region)

    dispatch_result = dispatch_analysis_input_write(
        export_path=export_path,
        sheet_name=SHEET_NAME,
        scenario=scenario_to_run,
        region=region,
        context_label="supply_data_pipeline.run_supply_leap_import",
    )
    if dispatch_result.get("mode") == "workbook":
        return export_path

    L = connect_to_leap()
    if L is None:
        raise RuntimeError("Failed to connect to LEAP.")
    fuel_catalog_preflight.run_fuel_catalog_preflight(
        export_path=export_path,
        sheet_name=SHEET_NAME,
        scenario=scenario_to_run,
        context="supply_data_pipeline.run_supply_leap_import",
        leap_app=L,
    )

    if fill_branches:
        print(
            "[INFO] Supply branches under Resources auto-create when their fuels "
            "are first used in Transformation/Demand and can be skipped until LEAP "
            "creates them."
        )
        ensure_supply_fuels_from_export(L, export_path)
        run_branch_fill(
            L,
            export_path,
            scenario_to_run,
            region,
            handle_current_accounts,
            raise_on_missing_branch=False,
        )
    return export_path


if __name__ == "__main__":
    run_supply_pipeline()
    if RUN_SUPPLY_LEAP_IMPORT:
        run_supply_leap_import()

#%%
