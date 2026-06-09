from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))
from codebase.functions.analysis_input_write_dispatcher import ensure_api_write_allowed
from codebase.functions.analysis_input_write_dispatcher import (
    ensure_analysis_view_api_read_allowed,
)
from codebase.functions.leap_api_guard import ensure_leap_api_allowed

#%%
# PART 0: OPTIONAL EXTRACTION FROM LEAP INTO transformation_results_*.xlsx

LEAP_RESULTS_TEMPLATE_PATHS = [
    REPO_ROOT / "data" / "leap results tables" / "transformation_results_20_USA_Reference.xlsx",
    REPO_ROOT / "data" / "leap results tables" / "transformation_results_20_USA_Target.xlsx",
]
LEAP_RESULTS_FORCE_RECALC = False


def run_transformation_results_extraction(
    template_paths: list[Path | str] | None = None,
    *,
    force_recalc: bool = LEAP_RESULTS_FORCE_RECALC,
) -> dict:
    """
    Refill one or more transformation results template workbooks from LEAP.

    This calls codebase.leap_results_workflow.run_template_fill() with a narrowed
    TEMPLATE_PATHS list so you can inspect the workbook generation independently.
    """
    from codebase import leap_results_workflow as lrw

    selected_paths = [_resolve(path) for path in (template_paths or LEAP_RESULTS_TEMPLATE_PATHS)]
    original_paths = list(lrw.TEMPLATE_PATHS)
    original_force_recalc = bool(lrw.FORCE_RECALC)
    try:
        lrw.TEMPLATE_PATHS = selected_paths
        lrw.FORCE_RECALC = bool(force_recalc)
        print("\n" + "#" * 79)
        print("# RUNNING LEAP RESULTS TEMPLATE FILL FOR TRANSFORMATION WORKBOOKS")
        print("#" * 79)
        print(f"[INFO] TEMPLATE_PATHS = {[str(path) for path in selected_paths]}")
        print(f"[INFO] FORCE_RECALC = {lrw.FORCE_RECALC}")
        return lrw.run_template_fill()
    finally:
        lrw.TEMPLATE_PATHS = original_paths
        lrw.FORCE_RECALC = original_force_recalc


#%%
# PART 0A: RUN LEAP EXTRACTION INTO transformation_results_*.xlsx
#
# Run this cell if you want to regenerate the transformation results workbook(s)
# from LEAP before inspecting them.

# extraction_log = run_transformation_results_extraction(
#     LEAP_RESULTS_TEMPLATE_PATHS,
#     force_recalc=LEAP_RESULTS_FORCE_RECALC,
# )
# extraction_log


#%%
# SHARED HELPERS

def _resolve(path: Path | str) -> Path:
    raw = str(path).replace("\\", "/")
    candidate = Path(raw)
    return candidate if candidate.is_absolute() else (REPO_ROOT / candidate)


################################
#%%
# LEAP CONNECTION HELPERS

def connect_to_leap_probe():
    ensure_analysis_view_api_read_allowed(
        "transformation_leap_probe_workflow.connect_to_leap_probe"
    )
    ensure_leap_api_allowed("transformation_leap_probe_workflow.connect_to_leap_probe")
    from win32com.client import Dispatch, GetActiveObject

    try:
        app = GetActiveObject("LEAP.LEAPApplication")
        print("[INFO] Connected to existing LEAP instance")
    except Exception:
        app = Dispatch("LEAP.LEAPApplication")
        print("[INFO] Created new LEAP instance")
    print(f"[INFO] Active area: {app.ActiveArea}")
    return app


def _show_analysis_view(app) -> None:
    for method_name in ("ShowAnalysisView", "ShowAnalysisViewTree", "ShowAnalysisViewBranches"):
        method = getattr(app, method_name, None)
        if callable(method):
            try:
                method()
                print(f"[INFO] View reset via {method_name}()")
                break
            except Exception as exc:
                print(f"[WARN] View reset failed via {method_name}(): {exc}")


def _load_leap_sheet(export_path: Path) -> pd.DataFrame:
    df = pd.read_excel(export_path, sheet_name="LEAP", header=2)
    required = {"Branch Path", "Variable", "Scenario", "Region"}
    missing = required.difference(df.columns)
    if missing:
        raise ValueError(f"Missing required columns in {export_path}: {sorted(missing)}")
    return df


def _head(text, n: int = 120) -> str:
    value = str(text)
    return value[:n] + ("..." if len(value) > n else "")


#%%
# PART 1: CONFIG FOR RESULTS WORKBOOK INSPECTION

RESULTS_WORKBOOK_PATH = (
    REPO_ROOT / "data" / "leap results tables" / "transformation_results_20_USA_Reference.xlsx"
)
RESULTS_SHEET_NAME = "blast furnace"
RESULTS_PREVIEW_ROWS = 20

def inspect_transformation_results_workbook(
    results_path: Path | str,
    *,
    preview_rows: int = 12,
) -> dict[str, pd.DataFrame]:
    """
    Inspect a transformation results workbook produced upstream.

    Notebook usage:
        inspect_transformation_results_workbook(
            r"data\\leap results tables\\transformation_results_20_USA_Reference.xlsx"
        )
    """
    results_path = _resolve(results_path)
    workbook = pd.ExcelFile(results_path)
    print(f"[INFO] Inspecting results workbook: {results_path}")
    print(f"[INFO] Sheets: {workbook.sheet_names}")

    previews: dict[str, pd.DataFrame] = {}
    for sheet_name in workbook.sheet_names:
        print("\n" + "#" * 79)
        print(f"# RESULTS WORKBOOK SHEET: {sheet_name}")
        print("#" * 79)
        raw = pd.read_excel(results_path, sheet_name=sheet_name, header=None)
        preview = raw.head(int(preview_rows)).copy()
        previews[sheet_name] = preview
        print(preview.to_string(index=False, header=False))
    return previews


def inspect_transformation_results_sheet(
    results_path: Path | str,
    *,
    sheet_name: str,
    preview_rows: int = 20,
) -> pd.DataFrame:
    """
    Inspect one sheet from a transformation results workbook.

    Notebook usage:
        inspect_transformation_results_sheet(
            r"data\\leap results tables\\transformation_results_20_USA_Reference.xlsx",
            sheet_name="blast furnace"
        )
    """
    results_path = _resolve(results_path)
    raw = pd.read_excel(results_path, sheet_name=sheet_name, header=None)
    preview = raw.head(int(preview_rows)).copy()
    print("\n" + "#" * 79)
    print(f"# RESULTS WORKBOOK SINGLE SHEET PREVIEW: {sheet_name}")
    print("#" * 79)
    print(preview.to_string(index=False, header=False))
    return raw


#%%
# PART 1A: RUN FULL RESULTS WORKBOOK PREVIEW
#
# Run this cell to print the top rows of every sheet in the transformation
# results workbook.

# inspect_transformation_results_workbook(
#     RESULTS_WORKBOOK_PATH,
#     preview_rows=RESULTS_PREVIEW_ROWS,
# )


#%%
# PART 1B: RUN SINGLE RESULTS SHEET PREVIEW
#
# Run this cell to inspect one sheet only.

# inspect_transformation_results_sheet(
#     RESULTS_WORKBOOK_PATH,
#     sheet_name=RESULTS_SHEET_NAME,
#     preview_rows=RESULTS_PREVIEW_ROWS,
# )


#%%
# PART 2: CONFIG FOR LEAP IMPORT WORKBOOK INSPECTION

LEAP_IMPORT_WORKBOOK_PATH = (
    REPO_ROOT
    / "outputs"
    / "results_supply_link"
    / "leap_exports"
    / "transformation_leap_imports_20_USA_Reference.xlsx"
)
LEAP_IMPORT_SCENARIO = "Reference"
LEAP_IMPORT_REGION_NAME = "United States"
LEAP_IMPORT_BRANCH_FILTER = "blast furnaces"
LEAP_IMPORT_VARIABLE_FILTER = None
LEAP_IMPORT_LIMIT = 20

def inspect_transformation_leap_workbook(
    export_path: Path | str,
    *,
    scenario: str = "Reference",
    region: str = "United States",
    branch_filter: str | None = None,
    variable_filter: str | None = None,
    limit: int | None = 20,
) -> pd.DataFrame:
    """
    Inspect the rows that will be written from the transformation LEAP import workbook.

    Notebook usage:
        inspect_transformation_leap_workbook(
            r"outputs\\results_supply_link\\leap_exports\\transformation_leap_imports_20_USA_Reference.xlsx",
            branch_filter="blast furnaces"
        )
    """
    export_path = _resolve(export_path)
    df = _load_leap_sheet(export_path)
    df = df[df["Scenario"].astype(str).str.strip() == str(scenario).strip()].copy()
    df = df[df["Region"].astype(str).str.strip() == str(region).strip()].copy()
    df = df[df["Branch Path"].astype(str).str.startswith("Transformation\\")].copy()
    if branch_filter:
        token = str(branch_filter).strip().lower()
        df = df[df["Branch Path"].astype(str).str.lower().str.contains(token, na=False)].copy()
    if variable_filter:
        token = str(variable_filter).strip().lower()
        df = df[df["Variable"].astype(str).str.lower() == token].copy()
    df = df.reset_index(drop=True)
    if limit is not None:
        df = df.head(int(limit)).copy()
    print("\n" + "#" * 79)
    print("# LEAP IMPORT WORKBOOK ROWS")
    print("#" * 79)
    print(df.to_string(index=False))
    return df


#%%
# PART 2B: LIST TRANSFORMATION SECTORS IN THE LEAP IMPORT WORKBOOK

def list_transformation_sectors(
    export_path: Path | str,
    *,
    scenario: str = LEAP_IMPORT_SCENARIO,
    region: str = LEAP_IMPORT_REGION_NAME,
) -> list[str]:
    export_path = _resolve(export_path)
    df = _load_leap_sheet(export_path)
    df = df[df["Scenario"].astype(str).str.strip() == str(scenario).strip()].copy()
    df = df[df["Region"].astype(str).str.strip() == str(region).strip()].copy()
    df = df[df["Branch Path"].astype(str).str.startswith("Transformation\\")].copy()
    sectors = (
        df["Branch Path"]
        .astype(str)
        .str.extract(r"^Transformation\\([^\\]+)", expand=False)
        .dropna()
        .astype(str)
        .drop_duplicates()
        .sort_values()
        .tolist()
    )
    print("\n" + "#" * 79)
    print("# TRANSFORMATION SECTORS")
    print("#" * 79)
    for sector in sectors:
        print(sector)
    return sectors


#%%
# PART 2C: LIST TRANSFORMATION SECTORS
#
# Run this cell to get the sector names for sector-by-sector probing.

list_transformation_sectors(
    LEAP_IMPORT_WORKBOOK_PATH,
    scenario=LEAP_IMPORT_SCENARIO,
    region=LEAP_IMPORT_REGION_NAME,
)


#%%
# PART 2A: SHOW LEAP IMPORT WORKBOOK ROWS
#
# Run this cell to inspect exactly which rows are being fed into LEAP.

# inspect_transformation_leap_workbook(
#     LEAP_IMPORT_WORKBOOK_PATH,
#     scenario=LEAP_IMPORT_SCENARIO,
#     region=LEAP_IMPORT_REGION_NAME,
#     branch_filter=LEAP_IMPORT_BRANCH_FILTER,
#     variable_filter=LEAP_IMPORT_VARIABLE_FILTER,
#     limit=LEAP_IMPORT_LIMIT,
# )


#%%
# PART 3: CONFIG FOR DIRECT LEAP PROBE

PROBE_SCENARIO = LEAP_IMPORT_SCENARIO
PROBE_REGION = LEAP_IMPORT_REGION_NAME
PROBE_BRANCH_FILTER = LEAP_IMPORT_BRANCH_FILTER
PROBE_VARIABLE_FILTER = LEAP_IMPORT_VARIABLE_FILTER
PROBE_LIMIT = LEAP_IMPORT_LIMIT
PROBE_WRITE_EXPRESSION = False
PROBE_SECTOR_NAME = "Blast furnaces"

def probe_transformation_workbook(
    export_path: Path | str,
    *,
    scenario: str = "Reference",
    region: str = "United States",
    branch_filter: str | None = None,
    variable_filter: str | None = None,
    limit: int | None = 20,
    write_expression: bool = False,
) -> list[dict[str, object]]:
    export_path = _resolve(export_path)
    df = _load_leap_sheet(export_path)
    df = df[df["Scenario"].astype(str).str.strip() == str(scenario).strip()].copy()
    df = df[df["Region"].astype(str).str.strip() == str(region).strip()].copy()
    df = df[df["Branch Path"].astype(str).str.startswith("Transformation\\")].copy()
    if branch_filter:
        token = str(branch_filter).strip().lower()
        df = df[df["Branch Path"].astype(str).str.lower().str.contains(token, na=False)].copy()
    if variable_filter:
        token = str(variable_filter).strip().lower()
        df = df[df["Variable"].astype(str).str.lower() == token].copy()
    df = df.reset_index(drop=True)
    if limit is not None:
        df = df.head(int(limit)).copy()

    app = connect_to_leap_probe()
    app.ActiveScenario = scenario
    app.ActiveRegion = region
    _show_analysis_view(app)
    try:
        app.ActiveBranch = "Transformation"
    except Exception as exc:
        print(f"[WARN] Failed to set ActiveBranch to Transformation: {exc}")

    results: list[dict[str, object]] = []
    print(f"[INFO] Probing {len(df)} transformation row(s) from {export_path.name}")
    for idx, row in df.iterrows():
        branch_path = str(row["Branch Path"]).strip()
        variable_name = str(row["Variable"]).strip()
        expression = row["Expression"] if "Expression" in row.index else None
        print(f"\n[{idx + 1}/{len(df)}] {branch_path} / {variable_name}")
        outcome: dict[str, object] = {
            "branch_path": branch_path,
            "variable": variable_name,
            "exists": False,
            "branch_ok": False,
            "variable_ok": False,
            "write_ok": None,
            "error": "",
        }
        try:
            exists = bool(app.Branches.Exists(branch_path))
            outcome["exists"] = exists
            print(f"  exists: {exists}")
            if not exists:
                results.append(outcome)
                continue
            branch = app.Branch(branch_path)
            outcome["branch_ok"] = True
            print(f"  branch full: {_head(branch.FullName)}")
            variable = branch.Variable(variable_name)
            if variable is None:
                print("  variable: None")
                results.append(outcome)
                continue
            outcome["variable_ok"] = True
            try:
                print(f"  current expr: {_head(variable.Expression)}")
            except Exception as exc:
                print(f"  current expr read failed: {exc}")
            if write_expression:
                try:
                    ensure_api_write_allowed(
                        "transformation_leap_probe_workflow.probe_transformation_workbook"
                    )
                    variable.Expression = expression
                    outcome["write_ok"] = True
                    print("  write: ok")
                except Exception as exc:
                    outcome["write_ok"] = False
                    outcome["error"] = str(exc)
                    print(f"  write failed: {exc}")
                    results.append(outcome)
                    break
            results.append(outcome)
        except Exception as exc:
            outcome["error"] = str(exc)
            print(f"  probe failed: {exc}")
            results.append(outcome)
            break
    return results


#%%
# PART 3A: PROBE ONE TRANSFORMATION SECTOR

def probe_transformation_sector(
    export_path: Path | str,
    *,
    sector_name: str,
    scenario: str = PROBE_SCENARIO,
    region: str = PROBE_REGION,
    variable_filter: str | None = None,
    limit: int | None = None,
    write_expression: bool = False,
) -> list[dict[str, object]]:
    export_path = _resolve(export_path)
    df = _load_leap_sheet(export_path)
    df = df[df["Scenario"].astype(str).str.strip() == str(scenario).strip()].copy()
    df = df[df["Region"].astype(str).str.strip() == str(region).strip()].copy()
    df = df[df["Branch Path"].astype(str).str.startswith(f"Transformation\\{sector_name}\\")].copy()
    if variable_filter:
        token = str(variable_filter).strip().lower()
        df = df[df["Variable"].astype(str).str.lower() == token].copy()
    if limit is not None:
        df = df.head(int(limit)).copy()

    tmp_path = export_path
    # Reuse the existing probe logic by temporarily filtering on exact sector rows.
    # This keeps the LEAP write path identical to the row-by-row probe above.
    app = connect_to_leap_probe()
    app.ActiveScenario = scenario
    app.ActiveRegion = region
    _show_analysis_view(app)
    try:
        app.ActiveBranch = "Transformation"
    except Exception as exc:
        print(f"[WARN] Failed to set ActiveBranch to Transformation: {exc}")

    results: list[dict[str, object]] = []
    print(
        f"[INFO] Probing {len(df)} transformation row(s) for sector '{sector_name}' "
        f"from {tmp_path.name}"
    )
    for idx, row in df.reset_index(drop=True).iterrows():
        branch_path = str(row["Branch Path"]).strip()
        variable_name = str(row["Variable"]).strip()
        expression = row["Expression"] if "Expression" in row.index else None
        print(f"\n[{idx + 1}/{len(df)}] {branch_path} / {variable_name}")
        outcome: dict[str, object] = {
            "sector": sector_name,
            "branch_path": branch_path,
            "variable": variable_name,
            "exists": False,
            "branch_ok": False,
            "variable_ok": False,
            "write_ok": None,
            "error": "",
        }
        try:
            exists = bool(app.Branches.Exists(branch_path))
            outcome["exists"] = exists
            print(f"  exists: {exists}")
            if not exists:
                results.append(outcome)
                continue
            branch = app.Branch(branch_path)
            outcome["branch_ok"] = True
            print(f"  branch full: {_head(branch.FullName)}")
            variable = branch.Variable(variable_name)
            if variable is None:
                print("  variable: None")
                results.append(outcome)
                continue
            outcome["variable_ok"] = True
            try:
                print(f"  current expr: {_head(variable.Expression)}")
            except Exception as exc:
                print(f"  current expr read failed: {exc}")
            if write_expression:
                try:
                    ensure_api_write_allowed(
                        "transformation_leap_probe_workflow.probe_transformation_sector"
                    )
                    variable.Expression = expression
                    outcome["write_ok"] = True
                    print("  write: ok")
                except Exception as exc:
                    outcome["write_ok"] = False
                    outcome["error"] = str(exc)
                    print("  write: no")
                    results.append(outcome)
                    break
            results.append(outcome)
        except Exception as exc:
            outcome["error"] = str(exc)
            print(f"  probe failed: {exc}")
            results.append(outcome)
            break
    return results


#%%
# PART 3B: PROBE ALL TRANSFORMATION SECTORS

def probe_all_transformation_sectors(
    export_path: Path | str,
    *,
    scenario: str = PROBE_SCENARIO,
    region: str = PROBE_REGION,
    variable_filter: str | None = None,
    limit_per_sector: int | None = None,
    write_expression: bool = True,
) -> pd.DataFrame:
    sectors = list_transformation_sectors(
        export_path,
        scenario=scenario,
        region=region,
    )
    summary_rows: list[dict[str, object]] = []
    for sector in sectors:
        print("\n" + "=" * 79)
        print(f"SECTOR PROBE: {sector}")
        print("=" * 79)
        sector_results = probe_transformation_sector(
            export_path,
            sector_name=sector,
            scenario=scenario,
            region=region,
            variable_filter=variable_filter,
            limit=limit_per_sector,
            write_expression=write_expression,
        )
        attempted = [row for row in sector_results if row.get("variable_ok")]
        failed = [row for row in sector_results if row.get("write_ok") is False]
        summary_rows.append(
            {
                "sector": sector,
                "rows_seen": len(sector_results),
                "rows_attempted": len(attempted),
                "rows_failed": len(failed),
                "first_failed_branch": failed[0]["branch_path"] if failed else "",
                "first_failed_variable": failed[0]["variable"] if failed else "",
                "status": "failed" if failed else "ok",
            }
        )
    summary = pd.DataFrame(summary_rows).sort_values(["status", "sector"]).reset_index(drop=True)
    print("\n" + "#" * 79)
    print("# ALL-SECTOR PROBE SUMMARY")
    print("#" * 79)
    if not summary.empty:
        print(summary.to_string(index=False))
    return summary


#%%
# PART 3C: RUN ONE-SECTOR PROBE FROM PROBE_SECTOR_NAME
#
# Uses the current value of PROBE_SECTOR_NAME.
probe_transformation_sector(
    LEAP_IMPORT_WORKBOOK_PATH,
    sector_name=PROBE_SECTOR_NAME,
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    variable_filter=PROBE_VARIABLE_FILTER,
    limit=PROBE_LIMIT,
    write_expression=True,
)


#%%
# PART 3D: BLAST FURNACES

probe_transformation_sector(
    LEAP_IMPORT_WORKBOOK_PATH,
    sector_name="Blast furnaces",
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    variable_filter=PROBE_VARIABLE_FILTER,
    limit=PROBE_LIMIT,
    write_expression=True,
)


#%%
# PART 3E: COKE OVENS

probe_transformation_sector(
    LEAP_IMPORT_WORKBOOK_PATH,
    sector_name="Coke ovens",
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    variable_filter=PROBE_VARIABLE_FILTER,
    limit=PROBE_LIMIT,
    write_expression=True,
)


#%%
# PART 3F: GAS WORKS PLANTS

probe_transformation_sector(
    LEAP_IMPORT_WORKBOOK_PATH,
    sector_name="Gas works plants",
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    variable_filter=PROBE_VARIABLE_FILTER,
    limit=PROBE_LIMIT,
    write_expression=True,
)


#%%
# PART 3G: HYDROGEN TRANSFORMATION

probe_transformation_sector(
    LEAP_IMPORT_WORKBOOK_PATH,
    sector_name="Hydrogen transformation",
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    variable_filter=PROBE_VARIABLE_FILTER,
    limit=PROBE_LIMIT,
    write_expression=True,
)


#%%
# PART 3H: NG LIQUEFACTION

probe_transformation_sector(
    LEAP_IMPORT_WORKBOOK_PATH,
    sector_name="NG Liquefaction",
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    variable_filter=PROBE_VARIABLE_FILTER,
    limit=PROBE_LIMIT,
    write_expression=True,
)


#%%
# PART 3I: NATURAL GAS BLENDING PLANTS

probe_transformation_sector(
    LEAP_IMPORT_WORKBOOK_PATH,
    sector_name="Natural gas blending plants",
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    variable_filter=PROBE_VARIABLE_FILTER,
    limit=PROBE_LIMIT,
    write_expression=True,
)


#%%
# PART 3J: NON SPECIFIED TRANSFORMATION

probe_transformation_sector(
    LEAP_IMPORT_WORKBOOK_PATH,
    sector_name="Non specified transformation",
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    variable_filter=PROBE_VARIABLE_FILTER,
    limit=PROBE_LIMIT,
    write_expression=True,
)


#%%
# PART 3D: RUN ALL-SECTOR PROBE
#
# Use this to identify which transformation sectors work and which should be
# rebuilt/replaced in the problematic LEAP area.

probe_all_transformation_sectors(
    LEAP_IMPORT_WORKBOOK_PATH,
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    variable_filter=PROBE_VARIABLE_FILTER,
    limit_per_sector=None,
    write_expression=True,
)


#%%
# PART 3A: CONNECT TO LEAP ONLY
#
# Run this cell to verify LEAP connection and current area without probing rows.

_probe_app = connect_to_leap_probe()


#%%
# PART 3B: PROBE TRANSFORMATION ROWS WITHOUT WRITING
#
# Run this cell to test branch lookup and variable access only.

probe_transformation_workbook(
    LEAP_IMPORT_WORKBOOK_PATH,
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    branch_filter=PROBE_BRANCH_FILTER,
    variable_filter=PROBE_VARIABLE_FILTER,
    limit=PROBE_LIMIT,
    write_expression=False,
)


#%%
# PART 3C: PROBE TRANSFORMATION ROWS WITH WRITES ENABLED
#
# Run this cell only when you want to reproduce the LEAP write failure.

probe_transformation_workbook(
    LEAP_IMPORT_WORKBOOK_PATH,
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    branch_filter=PROBE_BRANCH_FILTER,
    variable_filter=PROBE_VARIABLE_FILTER,
    limit=PROBE_LIMIT,
    write_expression=True,
)
#%%
# Blast furnaces
# Hydrogen transformation 
# NG Liquefaction 
# Natural gas blending plants 
# Non specified transformation
# Gas works plants
# Coke ovens


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
