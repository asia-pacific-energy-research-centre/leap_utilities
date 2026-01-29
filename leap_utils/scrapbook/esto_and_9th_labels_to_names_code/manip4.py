#%%
# Summary: Extract ESTO label text into a new column for code_to_name mappings.
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
#%%

#%%
######### CONSTANTS (UNLIKELY TO CHANGE) #########
ENABLE_DEBUG_BREAKPOINTS = True
ESTO_PREFIX_PATTERN = re.compile(r"^\d+(?:\.\d+)*\s*")
REQUIRED_COLUMNS = {"esto_label", "name"}
NEW_COLUMN_NAME = "esto_label_text"
#%%

#%%
######### FUNCTIONS #########
def try_debug_breakpoint() -> None:
    """Trigger a debug breakpoint when enabled (safe to call anywhere).

    Inputs:
        None.
    Outputs:
        None.
    Side effects:
        May pause execution in a debugger.
    """
    if not ENABLE_DEBUG_BREAKPOINTS:
        return
    try:
        breakpoint()
    except Exception as breakpoint_exc:
        print(f"Debug breakpoint failed: {breakpoint_exc}")


def load_code_to_name_sheet(workbook_path: Path, sheet_name: str) -> pd.DataFrame:
    """Load the code_to_name sheet.

    Inputs:
        workbook_path: Path to the Excel workbook.
        sheet_name: Sheet name to read.
    Outputs:
        Dataframe with code_to_name data.
    Side effects:
        Reads from disk.
    """
    try:
        df = pd.read_excel(workbook_path, sheet_name=sheet_name, dtype=str).fillna("")
        missing = REQUIRED_COLUMNS - set(df.columns)
        if missing:
            raise ValueError(f"Missing required columns: {sorted(missing)}")
        return df
    except Exception as exc:
        print(f"Failed to read {sheet_name} from {workbook_path}: {exc}")
        try_debug_breakpoint()
        raise


def extract_esto_text(label: str) -> str:
    """Return the text part of an ESTO label (strip numeric code prefix).

    Inputs:
        label: ESTO label string, e.g., "07.17 Other products".
    Outputs:
        Text-only label, e.g., "Other products".
    Side effects:
        None.
    """
    if label is None:
        return ""
    text = str(label).strip()
    if not text:
        return ""
    return ESTO_PREFIX_PATTERN.sub("", text).strip() or text


def add_esto_label_text_column(df: pd.DataFrame) -> pd.DataFrame:
    """Add a column with text-only ESTO labels.

    Inputs:
        df: code_to_name dataframe.
    Outputs:
        Dataframe with NEW_COLUMN_NAME populated.
    Side effects:
        None.
    """
    try:
        output = df.copy()
        esto_labels = output["esto_label"].fillna("").astype(str).str.strip()
        names = output["name"].fillna("").astype(str).str.strip()

        extracted = esto_labels.apply(extract_esto_text)
        use_name = esto_labels == ""
        extracted = extracted.where(~use_name, names)

        output[NEW_COLUMN_NAME] = extracted
        return output
    except Exception as exc:
        print(f"Failed to add {NEW_COLUMN_NAME}: {exc}")
        try_debug_breakpoint()
        raise


def save_output_csv(df: pd.DataFrame, output_path: Path) -> None:
    """Save the updated mapping to CSV.

    Inputs:
        df: Dataframe to write.
        output_path: CSV output path.
    Outputs:
        None.
    Side effects:
        Writes to disk.
    """
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_csv(output_path, index=False)
        print(f"Saved output to {output_path}")
        print(f"Total rows: {len(df)}")
    except Exception as exc:
        print(f"Failed to write CSV to {output_path}: {exc}")
        try_debug_breakpoint()
        raise
#%%

#%%
######### CONSTANTS (LIKELY TO CHANGE) #########
WORKBOOK_PATH = Path(r"C:\\Users\\Work\\github\\leap_utilities\\config\\sector_fuel_codes_to_names.xlsx")
SOURCE_SHEET_NAME = "code_to_name"
OUTPUT_CSV_PATH = Path(r"C:\\Users\\Work\\github\\leap_utilities\\outputs\\code_to_name_with_esto_text.csv")
RUN_EXTRACT = True
#%%

#%%
######### RUN EXTRACTION (TOGGLE) #########
if RUN_EXTRACT:
    try:
        source_df = load_code_to_name_sheet(WORKBOOK_PATH, SOURCE_SHEET_NAME)
        updated_df = add_esto_label_text_column(source_df)
        save_output_csv(updated_df, OUTPUT_CSV_PATH)
    except Exception as exc:
        print(f"Failed to extract ESTO label text: {exc}")
        try_debug_breakpoint()
#%%
