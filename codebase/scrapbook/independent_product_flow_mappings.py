#%%
# Summary: Clean independent product/flow mappings and write to outputs.
import os
import pandas as pd

from codebase.utilities.master_config import read_config_table
#%%

#%%
######### CONSTANTS (UNLIKELY TO CHANGE) #########
ENABLE_DEBUG_BREAKPOINTS = True
NINTH_SECTOR_COLUMNS = [
    "sectors",
    "sub1sectors",
    "sub2sectors",
    "sub3sectors",
    "sub4sectors",
]
NINTH_FUEL_COLUMNS = ["fuels", "subfuels"]
#%%

#%%
######### FUNCTIONS #########

def ensure_repo_root():
    """Move to repo root if running from the scrapbook folder."""
    try:
        if os.getcwd().endswith("scrapbook"):
            os.chdir("../../")
    except Exception as exc:
        print(f"Failed to set repo root: {exc}")
        try_debug_breakpoint()
        raise


def try_debug_breakpoint():
    """Trigger a debug breakpoint when enabled (safe to call anywhere)."""
    if not ENABLE_DEBUG_BREAKPOINTS:
        return
    try:
        breakpoint()
    except Exception as breakpoint_exc:
        print(f"Debug breakpoint failed: {breakpoint_exc}")


def load_workbook_sheets(workbook_path):
    """Load every sheet into a dict of dataframes."""
    try:
        return {
            sheet_name: read_config_table(workbook_path, sheet_name=sheet_name, dtype=str)
            for sheet_name in ["product", "flow"]
        }
    except Exception as exc:
        print(f"Failed to read workbook {workbook_path}: {exc}")
        try_debug_breakpoint()
        raise


def _normalize_series(series, drop_x=False):
    values = series.fillna("").astype(str).str.strip()
    if drop_x:
        values = values[(values != "") & (values.str.lower() != "x")]
    else:
        values = values[values != ""]
    return values


def collect_reference_labels(workbook_path):
    """Collect reference labels from the sector_fuel_codes_to_names workbook."""
    try:
        ninth_df = read_config_table(workbook_path, sheet_name="9th", dtype=str)
        esto_df = read_config_table(workbook_path, sheet_name="ESTO", dtype=str)

        ninth_sectors = pd.concat(
            [_normalize_series(ninth_df[col], drop_x=True) for col in NINTH_SECTOR_COLUMNS if col in ninth_df],
            ignore_index=True,
        ).drop_duplicates()
        ninth_fuels = pd.concat(
            [_normalize_series(ninth_df[col], drop_x=True) for col in NINTH_FUEL_COLUMNS if col in ninth_df],
            ignore_index=True,
        ).drop_duplicates()

        esto_flows = _normalize_series(esto_df.get("flows", pd.Series([], dtype=str)))
        esto_products = _normalize_series(esto_df.get("products", pd.Series([], dtype=str)))

        return {
            "9th_sectors": set(ninth_sectors.tolist()),
            "9th_fuels": set(ninth_fuels.tolist()),
            "esto_flows": set(esto_flows.tolist()),
            "esto_products": set(esto_products.tolist()),
        }
    except Exception as exc:
        print(f"Failed to collect reference labels from {workbook_path}: {exc}")
        try_debug_breakpoint()
        raise


def dedupe_sheet(df):
    """Drop exact duplicate rows while preserving original order."""
    if df is None or df.empty:
        return df.copy(), 0
    before = len(df)
    deduped = df.drop_duplicates().copy()
    removed = before - len(deduped)
    return deduped, removed


def add_missing_labels(df, left_col, right_col, left_values, right_values, drop_x_left=False):
    """Append missing labels on each side with blank counterparts."""
    if left_col not in df.columns or right_col not in df.columns:
        raise ValueError(f"Missing required columns: {left_col}, {right_col}")

    updated = df.copy()
    existing_left = _normalize_series(updated[left_col], drop_x=drop_x_left)
    existing_right = _normalize_series(updated[right_col])

    missing_left = sorted(set(left_values) - set(existing_left.tolist()))
    missing_right = sorted(set(right_values) - set(existing_right.tolist()))

    if not missing_left and not missing_right:
        return updated, 0, 0

    new_rows = []
    for value in missing_left:
        row = {col: "" for col in updated.columns}
        row[left_col] = value
        new_rows.append(row)

    for value in missing_right:
        row = {col: "" for col in updated.columns}
        row[right_col] = value
        new_rows.append(row)

    updated = pd.concat([updated, pd.DataFrame(new_rows)], ignore_index=True)
    return updated, len(missing_left), len(missing_right)


def write_workbook(workbook_path, sheet_data):
    """Write the workbook with the provided sheet dataframes."""
    try:
        output_dir = os.path.dirname(workbook_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        with pd.ExcelWriter(workbook_path, engine="openpyxl", mode="w") as writer:
            for sheet_name, df in sheet_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception as exc:
        print(f"Failed to write workbook {workbook_path}: {exc}")
        try_debug_breakpoint()
        raise
#%%

#%%
######### CONSTANTS (LIKELY TO CHANGE) #########
SOURCE_WORKBOOK_PATH = "config/independent product flow mappings.xlsx"
REFERENCE_WORKBOOK_PATH = "config/sector_fuel_codes_to_names.xlsx"
OUTPUT_WORKBOOK_PATH = "outputs/independent product flow mappings.xlsx"
PRODUCT_SHEET_NAME = "product"
FLOW_SHEET_NAME = "flow"
PRODUCT_COLUMNS = {"esto_product", "9th_fuel"}
FLOW_COLUMNS = {"9th_sector", "esto_flow"}
RUN_MAPPING_CLEANUP = True
#%%

#%%
######### RUN CLEANUP (TOGGLE) #########
if RUN_MAPPING_CLEANUP:
    try:
        ensure_repo_root()
        sheets = load_workbook_sheets(SOURCE_WORKBOOK_PATH)
        references = collect_reference_labels(REFERENCE_WORKBOOK_PATH)

        updated = {
            sheet_name: df
            for sheet_name, df in sheets.items()
            if sheet_name not in {PRODUCT_SHEET_NAME, FLOW_SHEET_NAME}
        }
        total_removed = 0

        product_df = sheets.get(PRODUCT_SHEET_NAME)
        if product_df is None:
            raise ValueError(f"Missing sheet: {PRODUCT_SHEET_NAME}")
        if not PRODUCT_COLUMNS.issubset(product_df.columns):
            raise ValueError(f"{PRODUCT_SHEET_NAME} sheet missing columns: {sorted(PRODUCT_COLUMNS)}")

        product_df, removed = dedupe_sheet(product_df)
        total_removed += removed
        product_df, added_9th, added_esto = add_missing_labels(
            product_df,
            "9th_fuel",
            "esto_product",
            references["9th_fuels"],
            references["esto_products"],
            drop_x_left=True,
        )
        product_df, removed_after = dedupe_sheet(product_df)
        total_removed += removed_after
        updated[PRODUCT_SHEET_NAME] = product_df
        print(f"{PRODUCT_SHEET_NAME}: removed {removed + removed_after} duplicate rows")
        print(f"{PRODUCT_SHEET_NAME}: added {added_9th} 9th fuels, {added_esto} ESTO products")

        flow_df = sheets.get(FLOW_SHEET_NAME)
        if flow_df is None:
            raise ValueError(f"Missing sheet: {FLOW_SHEET_NAME}")
        if not FLOW_COLUMNS.issubset(flow_df.columns):
            raise ValueError(f"{FLOW_SHEET_NAME} sheet missing columns: {sorted(FLOW_COLUMNS)}")

        flow_df, removed = dedupe_sheet(flow_df)
        total_removed += removed
        flow_df, added_9th, added_esto = add_missing_labels(
            flow_df,
            "9th_sector",
            "esto_flow",
            references["9th_sectors"],
            references["esto_flows"],
            drop_x_left=True,
        )
        flow_df, removed_after = dedupe_sheet(flow_df)
        total_removed += removed_after
        updated[FLOW_SHEET_NAME] = flow_df
        print(f"{FLOW_SHEET_NAME}: removed {removed + removed_after} duplicate rows")
        print(f"{FLOW_SHEET_NAME}: added {added_9th} 9th sectors, {added_esto} ESTO flows")

        write_workbook(OUTPUT_WORKBOOK_PATH, updated)
        print(f"Done. Total duplicate rows removed: {total_removed}")
        print(f"Output written to: {OUTPUT_WORKBOOK_PATH}")
    except Exception as exc:
        print(f"Failed to clean mappings: {exc}")
        try_debug_breakpoint()
#%%
