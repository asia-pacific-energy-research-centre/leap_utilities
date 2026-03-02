#%%
# Summary: Draft schema tooling for LEAP import templates (Expression-focused).
# How it works:
# - Reads a LEAP export/import template and infers key columns + Expression column.
# - Builds a schema dict with column roles, validation rules, and sample rows.
# - Prints a preview of the schema so future scripts can fill Expression reliably.
import json
import os
from pathlib import Path

import pandas as pd
#%%

#%%
######### CONSTANTS (UNLIKELY TO CHANGE) #########
DEFAULT_SHEET_NAME = "Export"
DEFAULT_HEADER_ROW = 2
SCHEMA_OUTPUT_PATH = Path("config/leap_template_schema.json")
REFERENCE_TEMPLATE_PATH = Path("data/industry export.xlsx")
#%%

#%%
######### FUNCTIONS #########
def build_template_schema(template_path, sheet_name, header_row):
    """Infer a schema dict from a LEAP template spreadsheet.

    Inputs:
        template_path: Path to the template Excel file
        sheet_name: Sheet to inspect
        header_row: Row index (0-based) containing headers

    Outputs:
        dict with column definitions, key columns, and expression column.

    Side effects:
        None.
    """
    try:
        df = pd.read_excel(template_path, sheet_name=sheet_name, header=header_row)
        columns = [str(col) for col in df.columns]

        key_columns = [
            col
            for col in ["Branch Path", "Variable", "Scenario", "Region"]
            if col in columns
        ]
        id_columns = [
            col
            for col in ["BranchID", "VariableID", "ScenarioID", "RegionID"]
            if col in columns
        ]
        expression_column = "Expression" if "Expression" in columns else None
        level_columns = [col for col in columns if col.startswith("Level ")]
        year_columns = [col for col in columns if col.isdigit()]
        fixed_columns = [
            col for col in ["Scale", "Units", "Per..."] if col in columns
        ]

        column_roles = {
            "ids": id_columns,
            "keys": key_columns,
            "expression": expression_column,
            "levels": level_columns,
            "years": year_columns,
            "fixed": fixed_columns,
        }

        sample_rows = df.head(5).fillna("").to_dict(orient="records")

        schema = {
            "example_template_path": str(template_path),
            "sheet_name": sheet_name,
            "header_row": header_row,
            "expression_column": expression_column,
            "key_columns": key_columns,
            "id_columns": id_columns,
            "level_columns": level_columns,
            "year_columns": year_columns,
            "fixed_columns": fixed_columns,
            "all_columns": columns,
            "column_roles": column_roles,
            "sample_rows": sample_rows,
            "notes": {
                "expression_column": (
                    "Column to populate with LEAP expression strings "
                    "(e.g., constants, Interp(year,value,...) or Data(year,value,...))."
                ),
                "key_columns": (
                    "Columns that uniquely identify a LEAP variable row; used to match rows "
                    "when filling expressions."
                ),
                "id_columns": (
                    "Optional LEAP IDs; keep but do not modify when filling expressions "
                    "(they anchor rows to LEAP objects)."
                ),
                "fixed_columns": (
                    "Units/Scale/Per... are usually preserved from the template so expressions "
                    "use consistent metadata."
                ),
                "expression_format": (
                    "Expressions can be constants, Interp(year,value,...) for linear "
                    "interpolation, or Data(year,value,...) for explicit yearly series."
                ),
                "column_formatting": (
                    "LEAP export sheets include IDs, paths, variables, scenarios, regions, "
                    "metadata, and hierarchical Level columns to mirror the branch tree."
                ),
                "level_columns": (
                    "Level columns split the Branch Path into hierarchy segments for import "
                    "compatibility and sorting."
                ),
                "year_columns": (
                    "Year columns (if present) represent explicit values; Expression may "
                    "override or complement them depending on import mode."
                ),
            },
        }
        return schema
    except Exception as exc:
        print(f"Failed to build schema from {template_path}: {exc}")
        raise


def validate_schema(schema):
    """Validate that schema contains essential columns.

    Inputs:
        schema: schema dict from build_template_schema

    Outputs:
        bool indicating if schema looks usable.

    Side effects:
        Prints warnings.
    """
    try:
        required = ["Branch Path", "Variable", "Scenario", "Region"]
        missing = [col for col in required if col not in schema.get("all_columns", [])]
        if missing:
            print(f"[WARN] Missing required columns: {missing}")
            return False
        if not schema.get("expression_column"):
            print("[WARN] Expression column not found.")
            return False
        if not schema.get("key_columns"):
            print("[WARN] No key columns inferred.")
            return False
        return True
    except Exception as exc:
        print(f"Failed to validate schema: {exc}")
        raise


def write_schema(schema, output_path):
    """Write schema dict to a JSON file.

    Inputs:
        schema: schema dict
        output_path: Path to write

    Outputs:
        None

    Side effects:
        Writes file to disk.
    """
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(schema, indent=2))
        print(f"Wrote schema to {output_path}")
    except Exception as exc:
        print(f"Failed to write schema to {output_path}: {exc}")
        raise


def preview_schema(schema):
    """Print a quick schema preview."""
    try:
        print("\n=== LEAP template schema preview ===")
        print(f"Template: {schema.get('template_path')}")
        print(f"Sheet: {schema.get('sheet_name')} (header_row={schema.get('header_row')})")
        print(f"Expression column: {schema.get('expression_column')}")
        print(f"Key columns: {schema.get('key_columns')}")
        print(f"ID columns: {schema.get('id_columns')}")
        print(f"Level columns: {schema.get('level_columns')}")
        print(f"Fixed columns: {schema.get('fixed_columns')}")
        sample_rows = schema.get("sample_rows", [])
        if sample_rows:
            print("Sample rows:")
            for row in sample_rows:
                print(row)
    except Exception as exc:
        print(f"Failed to preview schema: {exc}")
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
RUN_SCHEMA_PREVIEW = True
WRITE_SCHEMA_FILE = True
#%%

#%%
######### RUN #########
if RUN_SCHEMA_PREVIEW:
    ensure_repo_root()
    schema = build_template_schema(
        REFERENCE_TEMPLATE_PATH,
        DEFAULT_SHEET_NAME,
        DEFAULT_HEADER_ROW,
    )
    preview_schema(schema)
    if validate_schema(schema) and WRITE_SCHEMA_FILE:
        write_schema(schema, SCHEMA_OUTPUT_PATH)
#%%
