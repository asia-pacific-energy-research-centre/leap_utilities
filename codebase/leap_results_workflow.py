#%%
"""
Workflow to export LEAP Results tables programmatically via the COM API.

Run order (per user preference):
Connect → set area → ensure calc (NeedsCalculation or force) → optionally activate
favorite → set axes/context → export via CSV (or fetch values directly later).

The UI-export path mirrors what you see in Results view: LEAP renders the table,
`ExportResultsCSV` writes it to disk, and pandas can convert it to Excel.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Optional
from datetime import datetime

import pandas as pd

# Make repo importable even when run from a notebook elsewhere.
REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.functions.leap_results_functions import (
    activate_favorite,
    connect_leap,
    ensure_calculated,
    ensure_parent_dirs,
    export_results_csv,
    list_dimensions,
    select_area,
    set_axes,
    set_context,
)


# Stable constants (unlikely to change often)
DEFAULT_OUTPUT_DIR = Path("outputs/leap_results")
TEMPLATE_PATHS: list[Path] = [
    Path("data/leap results tables/transport_results_20_USA_Target.xlsx"),
    Path("data/leap results tables/transport_results_20_USA_Reference.xlsx"),
    Path("data/leap results tables/industry_results_20_USA_Target.xlsx"),
    Path("data/leap results tables/industry_results_20_USA_Reference.xlsx"),
    Path("data/leap results tables/demand_others_results_20_USA_Target.xlsx"),
    Path("data/leap results tables/demand_others_results_20_USA_Reference.xlsx"),
    Path("data/leap results tables/demand_others_results_20_USA_Target.xlsx"),
    Path("data/leap results tables/buildings_results_20_USA_Target.xlsx"),
    Path("data/leap results tables/buildings_results_20_USA_Reference.xlsx")
]
COMBINED_XLSX_PATH = DEFAULT_OUTPUT_DIR / "leap_results_combined.xlsx"


# Frequently changed settings (edit these for your run)
FORCE_RECALC = False
WRITE_EXCEL = True

# Table specs: add one entry per table you want to create/extract.
# Paths default to outputs/leap_results/<name>.csv|xlsx
TABLE_SPECS = {
    "default": {
        "area": None,  # e.g., "US Transport Study"
        "scenario": None,
        "region": None,
        "year": None,
        "unit": None,
        "branch": None,  # full path, e.g., "Demand\\Transport"
        "variable": None,  # e.g., "Energy Demand"
        "favorite": None,  # e.g., "Results#Transport Energy"
        "x_axis": None,  # e.g., "Years"
        "legend": None,  # e.g., "Scenarios"
        "csv_path": DEFAULT_OUTPUT_DIR / "leap_results_default.csv",
        "sheet": "default",  # sheet name to use in combined workbook
    },
    # Add more tables here with their own context/outputs.
}

# Optional scaling to adjust LEAP default units to desired display units.
# Per-sheet overrides:
# UNIT_SCALES = {"Passenger road": {"target_unit": "Petajoules", "scale_factor": 1e-15}}
UNIT_SCALES: dict[str, dict] = {}
# Per-variable defaults (variable name as shown in the sheet, e.g., "Final Energy Demand")
UNIT_SCALES_BY_VARIABLE: dict[str, dict] = {
    "Final Energy Demand": {"target_unit": "Petajoules", "scale_factor": 1e-6},
    # Add more: "Capacity": {...}, "Imports": {...}
}

# Template-driven extraction settings
USE_TEMPLATE = True
ARCHIVE_TEMPLATE = True
ARCHIVE_DIR = Path("outputs/leap_results/archive")


def ensure_repo_root() -> None:
    """Make sure we run from repository root for relative paths."""
    cwd = Path.cwd()
    if cwd != REPO_ROOT:
        os.chdir(REPO_ROOT)


def write_excel_from_csv(csv_path: Path, xlsx_path: Path, sheet_name: str = "Results") -> Path:
    """Load CSV and save to Excel for parity with manual exports."""
    df = pd.read_csv(csv_path)
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return xlsx_path


def parse_template_sheet(sheet: pd.DataFrame) -> dict:
    """Extract metadata (variable, scenario, region, branch, unit, legend, axis items) from a template sheet."""
    meta = {}
    meta["variable"] = str(sheet.iloc[0, 0]).strip()
    # Scenario and Region are in row 2 (index 1) as "Scenario: X, Region: Y"
    scenario_region = str(sheet.iloc[1, 0])
    for part in scenario_region.split(","):
        if "Scenario:" in part:
            meta["scenario"] = part.split(":", 1)[1].strip()
        if "Region:" in part:
            meta["region"] = part.split(":", 1)[1].strip()
    # Branch row
    branch_line = str(sheet.iloc[2, 0])
    meta["branch"] = branch_line.split(":", 1)[1].strip() if ":" in branch_line else branch_line.strip()
    # Units row
    units_line = str(sheet.iloc[3, 0])
    meta["units"] = units_line.split(":", 1)[1].strip() if ":" in units_line else units_line.strip()
    # Legend label at A6 (index 5, col 0)
    meta["legend_label"] = str(sheet.iloc[5, 0]).strip()
    # Axis (X) items are row 6 (index 5) columns 1+
    axis_items = []
    for val in sheet.iloc[5, 1:]:
        if pd.isna(val):
            continue
        axis_items.append(val)
    meta["x_items"] = axis_items
    # Legend members are col 0 from row 7 (index 6) downward until blank
    legend_members = []
    for val in sheet.iloc[6:, 0]:
        if pd.isna(val) or str(val).strip() == "":
            break
        legend_members.append(str(val).strip())
    meta["legend_members"] = legend_members
    return meta


def build_fresh_table(app, meta: dict, scale_spec: dict | None = None) -> pd.DataFrame:
    """Fetch values for the given template metadata using LEAP direct ValueRS calls."""
    # Set context
    set_context(
        app,
        scenario=meta.get("scenario"),
        region=meta.get("region"),
        # Do not set ActiveUnit/ActiveVariable here; we will use defaults to avoid unit/variable name mismatches.
        branch_path=meta.get("branch"),
    )
    branch_obj = app.Branches.Item(meta["branch"])
    try:
        variable_obj = branch_obj.Variables.Item(meta["variable"])
    except Exception as exc:  # noqa: BLE001
        # Fall back to the first variable to allow processing to continue; caller can inspect logs.
        variable_obj = branch_obj.Variables.Item(1)
        print(f"Warning: variable '{meta['variable']}' not found on branch '{meta['branch']}' ({exc}); using first variable instead.")

    legend_label = meta["legend_label"]
    x_items = meta["x_items"]
    legend_members = meta["legend_members"]

    # Build data rows: first row is header
    header = [legend_label] + x_items
    rows = [header]
    for member in legend_members:
        row = [member]
        filter_str = f"{legend_label}={member}" if legend_label else ""
        for year in x_items:
            try:
                # Use default units by passing empty string to avoid unit-name mismatches (e.g., "Petajoules" vs LEAP unit code).
                val = variable_obj.ValueRS(meta["region"], meta["scenario"], int(year), "", filter_str)
            except Exception:
                val = float("nan")
            row.append(val)
        rows.append(row)
    data_df = pd.DataFrame(rows)

    # Apply optional scaling to numeric data cells only.
    # Do NOT scale the header row (years), or year labels become values like 2.022e-12.
    scale_factor = scale_spec.get("scale_factor") if scale_spec else None
    target_unit = scale_spec.get("target_unit") if scale_spec else None
    if scale_factor is not None:
        value_cols = list(data_df.columns[1:])  # first column is legend labels
        for col in value_cols:
            data_df.loc[1:, col] = pd.to_numeric(data_df.loc[1:, col], errors="coerce") * float(scale_factor)

    # Prepend the metadata rows to match template structure
    meta_rows = [
        [meta["variable"]],
        [f"Scenario: {meta.get('scenario','')}, Region: {meta.get('region','')}"],
        [f"Branch: {meta.get('branch','')}"],
        [f"Units: {target_unit or meta.get('units','')}"],
        [""],
    ]
    final_df = pd.DataFrame(meta_rows + data_df.values.tolist())
    return final_df


def archive_file(path: Path, archive_dir: Path) -> Optional[Path]:
    """Copy an existing file to an archive directory with a date-stamped suffix."""
    if not path.exists():
        return None
    archive_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    archived = archive_dir / f"{path.stem}_{stamp}{path.suffix}"
    archived.write_bytes(path.read_bytes())
    return archived


def run_results_export() -> dict:
    """Perform UI exports for each table spec and return a log dict."""
    ensure_repo_root()
    app = connect_leap()

    results = []
    dfs_for_combined = []
    for name, spec in TABLE_SPECS.items():
        csv_path = Path(spec.get("csv_path", DEFAULT_OUTPUT_DIR / f"{name}.csv"))
        ensure_parent_dirs([csv_path, COMBINED_XLSX_PATH])

        select_area(app, spec.get("area"))
        ensure_calculated(app, force=FORCE_RECALC)

        fave_status = activate_favorite(app, spec.get("favorite"))

        set_axes(app, x_axis=spec.get("x_axis"), legend=spec.get("legend"))
        set_context(
            app,
            scenario=spec.get("scenario"),
            region=spec.get("region"),
            year=spec.get("year"),
            unit=spec.get("unit"),
            branch_path=spec.get("branch"),
            variable_name=spec.get("variable"),
        )

        dims = list_dimensions(app)
        csv_path = export_results_csv(app, csv_path)
        entry = {
            "table": name,
            "csv": str(csv_path),
            "favorite_status": fave_status,
            "dimensions": dims,
            "sheet": spec.get("sheet", name),
        }

        if WRITE_EXCEL:
            df = pd.read_csv(csv_path)
            dfs_for_combined.append((spec.get("sheet", name), df))

        results.append(entry)

    combined_path = None
    if WRITE_EXCEL and dfs_for_combined:
        COMBINED_XLSX_PATH.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(COMBINED_XLSX_PATH, engine="openpyxl") as writer:
            for sheet_name, df in dfs_for_combined:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        combined_path = str(COMBINED_XLSX_PATH)

    return {"tables": results, "combined_xlsx": combined_path}


def run_template_fill() -> dict:
    """Read one or more LEAP Results template workbooks and refill each sheet from LEAP."""
    if not TEMPLATE_PATHS:
        raise ValueError("TEMPLATE_PATHS must be set when USE_TEMPLATE is True.")

    paths = TEMPLATE_PATHS

    print("Template refill: starting", flush=True)
    ensure_repo_root()
    app = connect_leap()
    print("Connected to LEAP", flush=True)
    default_area = next(iter(TABLE_SPECS.values()), {}).get("area") if TABLE_SPECS else None
    select_area(app, default_area)
    ensure_calculated(app, force=FORCE_RECALC)
    print("LEAP calculation check done", flush=True)

    outputs = []

    for tpl in paths:
        print(f"Starting template {tpl}", flush=True)
        archive_path = None
        if ARCHIVE_TEMPLATE:
            archive_path = archive_file(tpl, ARCHIVE_DIR)

        xl = pd.ExcelFile(tpl)
        output_path = DEFAULT_OUTPUT_DIR / tpl.name
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            sheet_logs = []
            for sheet_name in xl.sheet_names:
                print(f"Processing {tpl} / {sheet_name}", flush=True)
                sheet_df = xl.parse(sheet_name, header=None)
                meta = parse_template_sheet(sheet_df)
                scale_spec = UNIT_SCALES.get(sheet_name) or UNIT_SCALES_BY_VARIABLE.get(meta.get("variable"))
                fresh_df = build_fresh_table(app, meta, scale_spec=scale_spec)
                fresh_df.to_excel(writer, sheet_name=sheet_name, header=False, index=False)
                # Summarize X items (years) as a range for compact logging
                x_items = [v for v in meta["x_items"] if not pd.isna(v)]
                year_values = [int(v) for v in x_items if isinstance(v, (int, float))]
                x_summary = None
                if year_values:
                    x_summary = f"{min(year_values)}-{max(year_values)}"
                if x_items and str(x_items[-1]).lower() == "total":
                    x_summary = f"{x_summary} + Total" if x_summary else "Total"
                sheet_logs.append(
                    {
                        "sheet": sheet_name,
                        "legend": meta["legend_label"],
                        "x_range": x_summary,
                        "scale_used": scale_spec,
                    }
                )

        outputs.append(
            {
                "template": str(tpl),
                "archived_copy": str(archive_path) if archive_path else None,
                "output": str(output_path),
                "sheets": sheet_logs,
            }
        )

    return {"templates_processed": outputs}


#%% Bottom run block (edit toggles above, then run this cell)
if __name__ == "__main__":
    try:
        if USE_TEMPLATE:
            run_log = run_template_fill()
            print("Template refill complete:")
        else:
            run_log = run_results_export()
            print("LEAP Results export complete:")
        for key, val in run_log.items():
            print(f"  {key}: {val}")
    except Exception as exc:  # noqa: BLE001
        print(f"Export failed: {exc}")

#%%
