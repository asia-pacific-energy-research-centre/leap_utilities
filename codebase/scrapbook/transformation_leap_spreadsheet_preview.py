#%%
# Summary: Prototype builder for LEAP transformation spreadsheet exports.
# How it works:
# - Builds a log-style dataframe with Branch Path, Variable, Scenario, and yearly values.
# - Uses leap_excel_io.finalise_export_df to pivot into LEAP import format.
# - Prints a preview so we can verify paths/variables/years before wiring real data.
import os
import sys
import pandas as pd

# Ensure repo root is on sys.path when running from the scrapbook folder.
try:
    REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    if REPO_ROOT not in sys.path:
        sys.path.insert(0, REPO_ROOT)
except Exception as exc:
    print(f"Failed to add repo root to sys.path: {exc}")

from codebase.functions.leap_excel_io import finalise_export_df, save_export_files
#%%

#%%
######### CONSTANTS (UNLIKELY TO CHANGE) #########
BASE_YEAR = 2023
FINAL_YEAR = 2023
DEFAULT_REGION = "United States"
DEFAULT_SCENARIO = "Current Accounts"
DEFAULT_UNITS = "Petajoule"
DEFAULT_SCALE = ""
DEFAULT_PER = ""
DEFAULT_MODEL_NAME = "LEAP Transformation Preview"
DEFAULT_OUTPUT_DIR = os.path.join("outputs", "leap_previews")
DEFAULT_OUTPUT_FILENAME = "transformation_leap_preview.xlsx"
#%%

#%%
######### FUNCTIONS #########
def build_year_rows(branch_path, measure, scenario, value_by_year, units, scale, per_value):
    """Return log-style rows for a LEAP import file.

    Inputs:
        branch_path: LEAP branch path string (e.g., Transformation\\Coke ovens)
        measure: LEAP variable name (e.g., "Process Efficiency")
        scenario: scenario label (e.g., "Current Accounts")
        value_by_year: dict of year -> value
        units: units string for LEAP
        scale: scale string for LEAP
        per_value: per... string for LEAP

    Outputs:
        List[dict] with fields expected by finalise_export_df.

    Side effects:
        None.
    """
    try:
        rows = []
        for year, value in sorted(value_by_year.items()):
            rows.append(
                {
                    "Branch_Path": branch_path,
                    "Scenario": scenario,
                    "Measure": measure,
                    "Units": units,
                    "Scale": scale,
                    "Per...": per_value,
                    "Date": int(year),
                    "Value": float(value),
                }
            )
        return rows
    except Exception as exc:
        print(f"Failed to build year rows for {branch_path}: {exc}")
        raise


def build_transformation_log_df(entries, scenario, base_year, final_year):
    """Build a log-style dataframe for LEAP transformation imports.

    Inputs:
        entries: list of dicts with keys branch_path, measure, value
        scenario: scenario label
        base_year: first year to include
        final_year: last year to include

    Outputs:
        pd.DataFrame compatible with finalise_export_df.

    Side effects:
        None.
    """
    try:
        rows = []
        for entry in entries:
            value_by_year = {year: entry["value"] for year in range(base_year, final_year + 1)}
            rows.extend(
                build_year_rows(
                    entry["branch_path"],
                    entry["measure"],
                    scenario,
                    value_by_year,
                    entry.get("units", DEFAULT_UNITS),
                    entry.get("scale", DEFAULT_SCALE),
                    entry.get("per_value", DEFAULT_PER),
                )
            )
        return pd.DataFrame(rows)
    except Exception as exc:
        print(f"Failed to build transformation log df: {exc}")
        raise


def build_transformation_preview_export(entries, scenario, region, base_year, final_year):
    """Return the LEAP export dataframe for previewing.

    Inputs:
        entries: list of dicts with transformation entries
        scenario: scenario label
        region: region label
        base_year: first year
        final_year: last year

    Outputs:
        pd.DataFrame in LEAP import format.

    Side effects:
        None.
    """
    try:
        log_df = build_transformation_log_df(entries, scenario, base_year, final_year)
        export_df = finalise_export_df(log_df, scenario, region, base_year, final_year)
        return export_df
    except Exception as exc:
        print(f"Failed to build transformation preview export: {exc}")
        raise
#%%

#%%
def ensure_repo_root():
    """Move to repo root if running from the scrapbook folder."""
    try:
        if os.getcwd().endswith("scrapbook"):
            os.chdir("../../")
    except Exception as exc:
        print(f"Failed to set repo root: {exc}")
        raise
#%%

#%%
def build_preview_output_path(output_dir, output_filename):
    """Return the output file path for the preview export.

    Inputs:
        output_dir: output directory relative to repo root
        output_filename: output filename for the preview export

    Outputs:
        Full output path string.

    Side effects:
        None.
    """
    try:
        return os.path.join(output_dir, output_filename)
    except Exception as exc:
        print(f"Failed to build preview output path: {exc}")
        raise
#%%

#%%
######### CONSTANTS (LIKELY TO CHANGE) #########
RUN_PREVIEW = True
SAVE_PREVIEW_FILE = True
OUTPUT_DIR = DEFAULT_OUTPUT_DIR
OUTPUT_FILENAME = DEFAULT_OUTPUT_FILENAME
MODEL_NAME = DEFAULT_MODEL_NAME
#%%

#%%
######### PREVIEW INPUTS #########
PREVIEW_TRANSFORMATION_ROWS = [
    {
        "branch_path": "Transformation\\Coal transformation - Coke ovens\\Output Fuels\\02.01 Coke oven coke",
        "measure": "Output",
        "value": 298.35,
        "units": "Petajoule",
    },
    {
        "branch_path": "Transformation\\Coal transformation - Coke ovens\\Processes\\09.08.01 Coke ovens",
        "measure": "Process Efficiency",
        "value": 0.6097,
        "units": "Percent",
    },
    {
        "branch_path": "Transformation\\Coal transformation - Coke ovens\\Processes\\09.08.01 Coke ovens\\Feedstock Fuels\\01.01 Coking coal",
        "measure": "Feedstock Fuel Share",
        "value": 1.0,
        "units": "",
    },
    {
        "branch_path": "Transformation\\Coal transformation - Coke ovens\\Processes\\09.08.01 Coke ovens\\Auxiliary Fuels\\17 Electricity",
        "measure": "Auxiliary Fuel Use",
        "value": 0.01,
        "units": "Gigajoule",
        "per_value": "Gigajoule",
    },
]
#%%

#%%
######### RUN PREVIEW #########
if RUN_PREVIEW:
    ensure_repo_root()
    preview_export = build_transformation_preview_export(
        PREVIEW_TRANSFORMATION_ROWS,
        DEFAULT_SCENARIO,
        DEFAULT_REGION,
        BASE_YEAR,
        FINAL_YEAR,
    )
    if preview_export is not None:
        print("\n=== Transformation LEAP preview ===")
        print(preview_export.head(10).to_string(index=False))
        if SAVE_PREVIEW_FILE:
            output_path = build_preview_output_path(OUTPUT_DIR, OUTPUT_FILENAME)
            save_export_files(
                preview_export,
                preview_export,
                output_path,
                BASE_YEAR,
                FINAL_YEAR,
                MODEL_NAME,
            )
#%%
