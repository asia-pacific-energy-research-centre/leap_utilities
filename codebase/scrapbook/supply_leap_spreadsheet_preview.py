#%%
# Summary: Prototype builder for LEAP supply import/export spreadsheets.
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

from codebase.functions.leap_excel_io import finalise_export_df
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
#%%

#%%
######### FUNCTIONS #########
def build_year_rows(branch_path, measure, scenario, value_by_year, units, scale, per_value):
    """Return log-style rows for a LEAP import file.

    Inputs:
        branch_path: LEAP branch path string (e.g., Supply\\Imports\\01.01 Coking coal)
        measure: LEAP variable name (e.g., "Imports")
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


def build_supply_log_df(supply_rows, scenario, region, base_year, final_year):
    """Build a log-style dataframe for LEAP supply imports/exports.

    Inputs:
        supply_rows: list of dicts with keys branch_path, measure, value
        scenario: scenario label
        region: region label (used downstream in finalise_export_df)
        base_year: first year to include
        final_year: last year to include

    Outputs:
        pd.DataFrame compatible with finalise_export_df.

    Side effects:
        None.
    """
    try:
        rows = []
        for row in supply_rows:
            value_by_year = {year: row["value"] for year in range(base_year, final_year + 1)}
            rows.extend(
                build_year_rows(
                    row["branch_path"],
                    row["measure"],
                    scenario,
                    value_by_year,
                    row.get("units", DEFAULT_UNITS),
                    row.get("scale", DEFAULT_SCALE),
                    row.get("per_value", DEFAULT_PER),
                )
            )
        return pd.DataFrame(rows)
    except Exception as exc:
        print(f"Failed to build supply log df: {exc}")
        raise


def build_supply_preview_export(supply_rows, scenario, region, base_year, final_year):
    """Return the LEAP export dataframe for previewing.

    Inputs:
        supply_rows: list of dicts with supply entries
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
        log_df = build_supply_log_df(supply_rows, scenario, region, base_year, final_year)
        export_df = finalise_export_df(log_df, scenario, region, base_year, final_year)
        return export_df
    except Exception as exc:
        print(f"Failed to build supply preview export: {exc}")
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
######### CONSTANTS (LIKELY TO CHANGE) #########
RUN_PREVIEW = True
#%%

#%%
######### PREVIEW INPUTS #########
PREVIEW_SUPPLY_ROWS = [
    {
        "branch_path": "Supply\\Imports\\01.01 Coking coal",
        "measure": "Imports",
        "value": 123.456,
    },
    {
        "branch_path": "Supply\\Exports\\01.01 Coking coal",
        "measure": "Exports",
        "value": 78.9,
    },
]
#%%

#%%
######### RUN PREVIEW #########
if RUN_PREVIEW:
    ensure_repo_root()
    preview_export = build_supply_preview_export(
        PREVIEW_SUPPLY_ROWS,
        DEFAULT_SCENARIO,
        DEFAULT_REGION,
        BASE_YEAR,
        FINAL_YEAR,
    )
    if preview_export is not None:
        print("\n=== Supply LEAP preview ===")
        print(preview_export.head(10).to_string(index=False))
#%%
