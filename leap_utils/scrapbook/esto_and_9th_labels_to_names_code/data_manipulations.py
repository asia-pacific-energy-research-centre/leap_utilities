#%%
# Summary: Build a merged 9th/ESTO code-to-name sheet for the sector/fuel mapping workbook.
import os
import pandas as pd
#%%

#%%
######### CONSTANTS (UNLIKELY TO CHANGE) #########
ENABLE_DEBUG_BREAKPOINTS = True
ESTO_COLUMN_MAP = {
    "FLOWS": "flows",
    "PRODUCTS": "products",
}
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


def load_code_to_name_sheet(workbook_path, sheet_name):
    """Load a worksheet into a dataframe.

    Inputs:
        workbook_path: Path to the Excel workbook.
        sheet_name: Sheet name to read.
    Outputs:
        Dataframe with the sheet contents.
    Side effects:
        Reads from disk.
    """
    try:
        return pd.read_excel(workbook_path, sheet_name=sheet_name)
    except Exception as exc:
        print(f"Failed to read {sheet_name} from {workbook_path}: {exc}")
        try_debug_breakpoint()
        raise


def build_code_to_name_columns(code_to_name_df):
    """Build a 9th/ESTO merged mapping with label/column fields.

    Inputs:
        code_to_name_df: Dataframe from the existing code_to_name sheet.
    Outputs:
        Dataframe with columns: 9th_label, 9th_column, esto_label, esto_column, name.
    Side effects:
        None.
    """
    try:
        required_cols = {"code", "name", "source_sheet", "source_column", "source_type"}
        missing = required_cols - set(code_to_name_df.columns)
        if missing:
            raise ValueError(f"Missing required columns: {sorted(missing)}")

        working = code_to_name_df.copy()
        working["sheet_key"] = working["source_sheet"].astype(str).str.strip().str.lower()
        working["label"] = working["code"].astype(str).str.strip()
        working["column_value"] = working["source_column"].astype(str).str.strip()

        # ESTO uses source_type to indicate flows vs products; map to column names.
        esto_mask = working["sheet_key"] == "esto"
        working.loc[esto_mask, "column_value"] = (
            working.loc[esto_mask, "source_type"]
            .astype(str)
            .str.strip()
            .str.upper()
            .map(ESTO_COLUMN_MAP)
        )
        working["column_value"] = working["column_value"].fillna("")

        long_df = working.melt(
            id_vars=["name", "sheet_key"],
            value_vars=["label", "column_value"],
            var_name="value_type",
            value_name="value",
        )

        pivoted = (
            long_df.pivot_table(
                index=["name", "value_type"],
                columns="sheet_key",
                values="value",
                aggfunc="first",
            )
            .reset_index()
            .rename_axis(None, axis=1)
        )

        labels = pivoted[pivoted["value_type"] == "label"].copy()
        labels = labels.rename(
            columns={
                "9th": "9th_label",
                "esto": "esto_label",
            }
        )

        columns = pivoted[pivoted["value_type"] == "column_value"].copy()
        columns = columns.rename(
            columns={
                "9th": "9th_column",
                "esto": "esto_column",
            }
        )

        merged = labels.merge(columns, on="name", how="outer")
        merged = merged.drop(columns=["value_type_x", "value_type_y"], errors="ignore")

        final_cols = ["9th_label", "9th_column", "esto_label", "esto_column", "name"]
        for col in final_cols:
            if col not in merged.columns:
                merged[col] = ""
        merged = merged[final_cols]
        return merged
    except Exception as exc:
        print(f"Failed to build the merged code-to-name sheet: {exc}")
        try_debug_breakpoint()
        raise


def write_code_to_name_sheet(workbook_path, output_df, output_sheet):
    """Write the new sheet into the workbook.

    Inputs:
        workbook_path: Excel workbook path.
        output_df: Dataframe to write.
        output_sheet: Sheet name to create/replace.
    Outputs:
        None.
    Side effects:
        Writes to disk.
    """
    try:
        with pd.ExcelWriter(
            workbook_path, mode="a", if_sheet_exists="replace", engine="openpyxl"
        ) as writer:
            output_df.to_excel(writer, sheet_name=output_sheet, index=False)
        print(f"Wrote {output_df.shape[0]} rows to {workbook_path}:{output_sheet}")
    except Exception as exc:
        print(f"Failed to write {output_sheet} to {workbook_path}: {exc}")
        try_debug_breakpoint()
        raise
#%%

#%%
######### CONSTANTS (LIKELY TO CHANGE) #########
WORKBOOK_PATH = "config/sector_fuel_codes_to_names.xlsx"
SOURCE_SHEET_NAME = "code_to_name"
OUTPUT_SHEET_NAME = "code_to_name_columns"
RUN_BUILD_CODE_TO_NAME_COLUMNS = False
#%%

#%%
######### RUN BUILD (TOGGLE) #########
if RUN_BUILD_CODE_TO_NAME_COLUMNS:
    try:
        ensure_repo_root()
        source_df = load_code_to_name_sheet(WORKBOOK_PATH, SOURCE_SHEET_NAME)
        output_df = build_code_to_name_columns(source_df)
        write_code_to_name_sheet(WORKBOOK_PATH, output_df, OUTPUT_SHEET_NAME)
    except Exception as exc:
        print(f"Failed to build {OUTPUT_SHEET_NAME}: {exc}")
        try_debug_breakpoint()
#%%
