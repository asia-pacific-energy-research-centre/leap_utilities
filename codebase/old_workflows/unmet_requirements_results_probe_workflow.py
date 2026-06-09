#%%
"""Notebook-style probe for extracting LEAP Results 'Unmet Requirements' on Resources branches.

This file is intentionally step-by-step and inspection-friendly:
1) connect to LEAP Results API,
2) inventory resource fuel branches and variable availability,
3) try multiple extraction methods,
4) save per-method outputs + comparison table.
"""
from __future__ import annotations

import sys
from pathlib import Path
from typing import Iterable

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase import leap_results_workflow as lrw
from codebase.functions.leap_results_functions import connect_leap, export_results_csv, set_axes, set_context


#%%
# Functions

def _resolve(path: Path | str) -> Path:
    """Resolve relative paths against repo root for notebook-safe execution."""
    raw = str(path).replace("\\", "/")
    candidate = Path(raw)
    return candidate if candidate.is_absolute() else (REPO_ROOT / candidate)


def _safe_float(value: object) -> float | None:
    numeric = pd.to_numeric(value, errors="coerce")
    if pd.isna(numeric):
        return None
    return float(numeric)


def _safe_branch(app, path: str):
    """Return LEAP branch object or None without raising."""
    branch_path = str(path or "").strip()
    if not branch_path:
        return None
    try:
        if not app.Branches.Exists(branch_path):
            return None
        return app.Branches.Item(branch_path)
    except Exception:
        return None


def _list_child_branches(parent_branch) -> list[tuple[str, str]]:
    """List child branches as (name, full_path)."""
    rows: list[tuple[str, str]] = []
    if parent_branch is None:
        return rows
    try:
        children = parent_branch.Children
        count = int(children.Count)
    except Exception:
        return rows
    for idx in range(1, count + 1):
        try:
            child = children.Item(idx)
        except Exception:
            continue
        try:
            name = str(child.Name).strip()
        except Exception:
            name = ""
        try:
            full_name = str(child.FullName).strip()
        except Exception:
            full_name = ""
        if not name and full_name and "\\" in full_name:
            name = full_name.rsplit("\\", 1)[-1].strip()
        if name:
            rows.append((name, full_name or name))
    return rows


def _available_variable_names(branch_obj) -> list[str]:
    """Return variable names available on one LEAP branch."""
    names: list[str] = []
    if branch_obj is None:
        return names
    try:
        variables = branch_obj.Variables
        count = int(getattr(variables, "Count", 0))
    except Exception:
        return names
    for idx in range(1, count + 1):
        try:
            name = str(getattr(variables.Item(idx), "Name", "") or "").strip()
        except Exception:
            name = ""
        if name:
            names.append(name)
    return names


def connect_results_api(*, visible: bool = True):
    """Connect to LEAP app and switch to Results table view."""
    try:
        app = connect_leap(visible=visible, reuse_running=True)
    except Exception as exc:
        raise RuntimeError(
            "Failed to connect to LEAP COM API. "
            "Run this on Windows with LEAP open and pywin32 installed. "
            f"Details: {exc}"
        ) from exc

    try:
        app.ShowResultsViewTable()
    except Exception as exc:
        print(f"[WARN] Could not force Results table view: {exc}")

    try:
        active_area = str(app.ActiveArea)
    except Exception:
        active_area = ""
    try:
        active_scenario = str(app.ActiveScenario)
    except Exception:
        active_scenario = ""
    try:
        active_region = str(app.ActiveRegion)
    except Exception:
        active_region = ""

    print("[INFO] Connected to LEAP")
    print(f"[INFO] Active area: {active_area}")
    print(f"[INFO] Active scenario: {active_scenario}")
    print(f"[INFO] Active region: {active_region}")
    return app


def inventory_resource_unmet_variable(
    app,
    *,
    roots: Iterable[str],
    variable_name: str,
) -> pd.DataFrame:
    """Inventory resource fuel branches and whether variable exists on each branch."""
    rows: list[dict[str, object]] = []
    for root in [str(item).strip() for item in roots if str(item).strip()]:
        root_branch = _safe_branch(app, root)
        if root_branch is None:
            rows.append(
                {
                    "root": root,
                    "fuel_name": "",
                    "branch_path": root,
                    "branch_exists": False,
                    "has_variable": False,
                    "available_variables": "",
                }
            )
            continue

        for fuel_name, fuel_full in _list_child_branches(root_branch):
            fuel_path = fuel_full or f"{root}\\{fuel_name}"
            fuel_branch = _safe_branch(app, fuel_path)
            var_names = _available_variable_names(fuel_branch)
            has_var = any(
                str(name).strip().lower() == str(variable_name).strip().lower()
                for name in var_names
            )
            rows.append(
                {
                    "root": root,
                    "fuel_name": fuel_name,
                    "branch_path": fuel_path,
                    "branch_exists": fuel_branch is not None,
                    "has_variable": bool(has_var),
                    "available_variables": " | ".join(var_names),
                }
            )

    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values(["root", "fuel_name"]).reset_index(drop=True)
    print(
        "[INFO] Resource inventory rows="
        f"{len(out)} | variable='{variable_name}' present on "
        f"{int(out['has_variable'].sum()) if not out.empty else 0} branches"
    )
    return out


def probe_unmet_via_branch_valuesrs(
    app,
    *,
    inventory_df: pd.DataFrame,
    scenario: str,
    region: str,
    years: Iterable[int],
    variable_name: str,
) -> pd.DataFrame:
    """Method A: direct ValueRS on each fuel branch variable."""
    rows: list[dict[str, object]] = []
    target_years = [int(year) for year in years]

    for _, item in inventory_df.iterrows():
        branch_path = str(item.get("branch_path") or "").strip()
        fuel_name = str(item.get("fuel_name") or "").strip()
        has_var = bool(item.get("has_variable"))
        if not branch_path:
            continue

        if not has_var:
            for year in target_years:
                rows.append(
                    {
                        "method": "branch_valuesrs",
                        "root": str(item.get("root") or ""),
                        "fuel_name": fuel_name,
                        "branch_path": branch_path,
                        "year": int(year),
                        "value": None,
                        "error": "variable_not_on_branch",
                        "filter_used": "",
                    }
                )
            continue

        try:
            set_axes(app, x_axis="Years", legend=None)
            set_context(
                app,
                scenario=scenario,
                region=region,
                branch_path=branch_path,
            )
            variable_obj = lrw._resolve_branch_variable(
                app,
                branch_path,
                variable_name,
                allow_substitution=False,
            )
        except Exception as exc:
            for year in target_years:
                rows.append(
                    {
                        "method": "branch_valuesrs",
                        "root": str(item.get("root") or ""),
                        "fuel_name": fuel_name,
                        "branch_path": branch_path,
                        "year": int(year),
                        "value": None,
                        "error": f"resolve_failed: {exc}",
                        "filter_used": "",
                    }
                )
            continue

        for year in target_years:
            value = None
            error = ""
            try:
                queried = variable_obj.ValueRS(region, scenario, int(year), "", "")
                value = _safe_float(queried)
            except Exception as exc_valuesrs:
                try:
                    queried = variable_obj.Value(int(year))
                    value = _safe_float(queried)
                except Exception as exc_value:
                    error = f"ValueRS failed: {exc_valuesrs}; Value failed: {exc_value}"
            rows.append(
                {
                    "method": "branch_valuesrs",
                    "root": str(item.get("root") or ""),
                    "fuel_name": fuel_name,
                    "branch_path": branch_path,
                    "year": int(year),
                    "value": value,
                    "error": error,
                    "filter_used": "",
                }
            )

    out = pd.DataFrame(rows)
    print(
        "[INFO] Method A (branch ValueRS) rows="
        f"{len(out)} | non-null values={int(out['value'].notna().sum()) if not out.empty else 0}"
    )
    return out


def probe_unmet_via_root_filter_valuesrs(
    app,
    *,
    inventory_df: pd.DataFrame,
    roots: Iterable[str],
    scenario: str,
    region: str,
    years: Iterable[int],
    variable_name: str,
    max_fuels_per_root: int | None = None,
) -> pd.DataFrame:
    """Method B: ValueRS from resource root with fuel filter strings."""
    rows: list[dict[str, object]] = []
    target_years = [int(year) for year in years]

    for root in [str(item).strip() for item in roots if str(item).strip()]:
        try:
            set_axes(app, x_axis="Years", legend="Fuel")
            set_context(app, scenario=scenario, region=region, branch_path=root)
            variable_obj = lrw._resolve_branch_variable(
                app,
                root,
                variable_name,
                allow_substitution=False,
            )
        except Exception as exc:
            rows.append(
                {
                    "method": "root_filter_valuesrs",
                    "root": root,
                    "fuel_name": "",
                    "branch_path": root,
                    "year": None,
                    "value": None,
                    "error": f"root_resolve_failed: {exc}",
                    "filter_used": "",
                }
            )
            continue

        dim_name, members = lrw.discover_legend_members_from_api(
            app,
            "Fuel",
            preferred_dimension_names=["Fuel", "Output Fuel", "Feedstock Fuel"],
        )
        if not members:
            fallback_members = (
                inventory_df[
                    inventory_df["root"].astype(str).str.strip().str.lower()
                    == root.strip().lower()
                ]
                .get("fuel_name", pd.Series(dtype=str))
                .dropna()
                .astype(str)
                .str.strip()
            )
            members = [item for item in fallback_members.tolist() if item]
        if max_fuels_per_root is not None and max_fuels_per_root > 0:
            members = members[: int(max_fuels_per_root)]

        for fuel_name in members:
            best_filter_name = ""
            for year in target_years:
                value = None
                error = ""
                for dim_candidate in [dim_name, "Fuel", "Output Fuel", "Feedstock Fuel"]:
                    dim_token = str(dim_candidate or "").strip()
                    if not dim_token:
                        continue
                    filter_str = f"{dim_token}={fuel_name}"
                    try:
                        queried = variable_obj.ValueRS(region, scenario, int(year), "", filter_str)
                        numeric = _safe_float(queried)
                        if numeric is None:
                            continue
                        value = numeric
                        best_filter_name = dim_token
                        error = ""
                        break
                    except Exception:
                        continue
                if value is None and not error:
                    error = "all_filter_variants_failed"
                rows.append(
                    {
                        "method": "root_filter_valuesrs",
                        "root": root,
                        "fuel_name": str(fuel_name),
                        "branch_path": f"{root}\\{fuel_name}",
                        "year": int(year),
                        "value": value,
                        "error": error,
                        "filter_used": best_filter_name,
                    }
                )

    out = pd.DataFrame(rows)
    print(
        "[INFO] Method B (root+filter ValueRS) rows="
        f"{len(out)} | non-null values={int(out['value'].notna().sum()) if not out.empty else 0}"
    )
    return out


def probe_unmet_via_resultvalue(
    app,
    *,
    inventory_df: pd.DataFrame,
    scenario: str,
    region: str,
    years: Iterable[int],
    variable_name: str,
) -> pd.DataFrame:
    """Method C: variable-object .Value(year) probes (no deprecated ResultValue API)."""
    rows: list[dict[str, object]] = []
    target_years = [int(year) for year in years]

    # Ensure context first; variable-object Value follows active scenario/region context.
    try:
        set_context(app, scenario=scenario, region=region)
    except Exception as exc:
        print(f"[WARN] Failed setting scenario/region before method C probes: {exc}")

    for _, item in inventory_df.iterrows():
        branch_path = str(item.get("branch_path") or "").strip()
        if not branch_path:
            continue
        fuel_name = str(item.get("fuel_name") or "").strip()
        variable_obj = None
        resolve_error = ""
        try:
            set_context(app, scenario=scenario, region=region, branch_path=branch_path)
            variable_obj = lrw._resolve_branch_variable(
                app,
                branch_path,
                variable_name,
                allow_substitution=False,
            )
        except Exception as exc:
            resolve_error = f"resolve_failed: {exc}"

        for year in target_years:
            value = None
            error = ""
            if variable_obj is None:
                error = resolve_error or "variable_resolution_failed"
            else:
                try:
                    queried = variable_obj.Value(int(year))
                    value = _safe_float(queried)
                    if value is None:
                        error = "value_returned_non_numeric"
                except Exception as exc:
                    error = f"variable_value_failed: {exc}"
            rows.append(
                {
                    "method": "variable_value",
                    "root": str(item.get("root") or ""),
                    "fuel_name": fuel_name,
                    "branch_path": branch_path,
                    "year": int(year),
                    "value": value,
                    "error": error,
                    "filter_used": "",
                }
            )

    out = pd.DataFrame(rows)
    print(
        "[INFO] Method C (Variable.Value) rows="
        f"{len(out)} | non-null values={int(out['value'].notna().sum()) if not out.empty else 0}"
    )
    return out


def probe_unmet_via_exported_csv(
    app,
    *,
    roots: Iterable[str],
    scenario: str,
    region: str,
    variable_name: str,
    output_dir: Path | str,
) -> tuple[pd.DataFrame, list[Path]]:
    """Method D: export current Results table to CSV and parse fuel/year values."""
    rows: list[dict[str, object]] = []
    csv_paths: list[Path] = []
    out_dir = _resolve(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    for root in [str(item).strip() for item in roots if str(item).strip()]:
        token = root.replace("\\", "_").replace("/", "_").replace(" ", "_")
        csv_path = out_dir / f"unmet_requirements_{token}_{scenario}.csv"

        try:
            set_axes(app, x_axis="Years", legend="Fuel")
            set_context(app, scenario=scenario, region=region, branch_path=root)
            _ = lrw._resolve_branch_variable(
                app,
                root,
                variable_name,
                allow_substitution=False,
            )
            # Do not set app.ActiveVariable here. In some LEAP contexts this can
            # trigger blocking modal errors ("Invalid variable"). We keep this
            # method modal-safe by exporting the current Results table for the
            # configured branch/scenario/region context only.
            try:
                app.ShowResultsViewTable()
            except Exception:
                pass
            export_results_csv(app, csv_path)
            csv_paths.append(csv_path)
        except Exception as exc:
            rows.append(
                {
                    "method": "export_results_csv",
                    "root": root,
                    "fuel_name": "",
                    "branch_path": root,
                    "year": None,
                    "value": None,
                    "error": f"export_failed: {exc}",
                    "filter_used": str(csv_path),
                }
            )
            continue

        try:
            parsed = lrw.parse_exported_results_csv(csv_path)
        except Exception as exc:
            rows.append(
                {
                    "method": "export_results_csv",
                    "root": root,
                    "fuel_name": "",
                    "branch_path": root,
                    "year": None,
                    "value": None,
                    "error": f"parse_failed: {exc}",
                    "filter_used": str(csv_path),
                }
            )
            continue

        if parsed.empty or len(parsed.columns) < 2:
            rows.append(
                {
                    "method": "export_results_csv",
                    "root": root,
                    "fuel_name": "",
                    "branch_path": root,
                    "year": None,
                    "value": None,
                    "error": "parsed_table_empty",
                    "filter_used": str(csv_path),
                }
            )
            continue

        first_col = parsed.columns[0]
        year_columns: list[tuple[str, int]] = []
        for col in parsed.columns[1:]:
            text = str(col or "").strip()
            if text.lower() == "total":
                continue
            try:
                year = int(float(text))
            except Exception:
                continue
            year_columns.append((str(col), int(year)))

        if not year_columns:
            rows.append(
                {
                    "method": "export_results_csv",
                    "root": root,
                    "fuel_name": "",
                    "branch_path": root,
                    "year": None,
                    "value": None,
                    "error": "no_year_columns_in_csv",
                    "filter_used": str(csv_path),
                }
            )
            continue

        for _, item in parsed.iterrows():
            fuel_name = str(item.get(first_col) or "").strip()
            if not fuel_name or fuel_name.lower() == "total":
                continue
            for col_name, year in year_columns:
                numeric = _safe_float(item.get(col_name))
                rows.append(
                    {
                        "method": "export_results_csv",
                        "root": root,
                        "fuel_name": fuel_name,
                        "branch_path": f"{root}\\{fuel_name}",
                        "year": int(year),
                        "value": numeric,
                        "error": "",
                        "filter_used": str(csv_path),
                    }
                )

    out = pd.DataFrame(rows)
    print(
        "[INFO] Method D (ExportResultsCSV parse) rows="
        f"{len(out)} | non-null values={int(out['value'].notna().sum()) if not out.empty else 0}"
    )
    return out, csv_paths


def save_probe_table(df: pd.DataFrame, path: Path | str) -> Path:
    """Save a probe dataframe to CSV and return absolute path."""
    out_path = _resolve(path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(out_path, index=False)
    print(f"[INFO] Saved: {out_path}")
    return out_path


def build_method_summary(method_tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Build one-line diagnostics per method for quick comparison."""
    rows: list[dict[str, object]] = []
    for method_name, df in method_tables.items():
        if df is None or df.empty:
            rows.append(
                {
                    "method": method_name,
                    "rows": 0,
                    "non_null_values": 0,
                    "rows_with_error": 0,
                    "distinct_fuels": 0,
                    "distinct_years": 0,
                }
            )
            continue
        values = pd.to_numeric(df.get("value"), errors="coerce")
        errors = df.get("error", pd.Series(index=df.index, dtype=str)).astype(str).str.strip()
        fuels = df.get("fuel_name", pd.Series(index=df.index, dtype=str)).astype(str).str.strip()
        years = pd.to_numeric(df.get("year"), errors="coerce")
        rows.append(
            {
                "method": method_name,
                "rows": int(len(df)),
                "non_null_values": int(values.notna().sum()),
                "rows_with_error": int((errors != "").sum()),
                "distinct_fuels": int(fuels[fuels != ""].nunique()),
                "distinct_years": int(years.dropna().astype(int).nunique()),
            }
        )
    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values("method").reset_index(drop=True)
    return out


#%%
# Runtime configuration (edit these before running cells)

PROBE_SCENARIO = "Reference"
PROBE_REGION = "United States"
PROBE_VARIABLE = "Unmet Requirements"
PROBE_ROOTS = ("Resources\\Primary", "Resources\\Secondary")
PROBE_YEARS = [2022, 2023, 2024]#, 2030, 2040, 2050, 2060]
PROBE_MAX_FUELS_PER_ROOT = None

PROBE_OUTPUT_DIR = REPO_ROOT / "outputs" / "leap_results" / "unmet_requirements_probe"
PROBE_SUPPORT_DIR = PROBE_OUTPUT_DIR / "supporting_files"
OUTPUT_INVENTORY_CSV = PROBE_OUTPUT_DIR / "resource_branch_inventory.csv"
OUTPUT_METHOD_A_CSV = PROBE_SUPPORT_DIR / "method_a_branch_valuesrs.csv"
OUTPUT_METHOD_B_CSV = PROBE_SUPPORT_DIR / "method_b_root_filter_valuesrs.csv"
OUTPUT_METHOD_C_CSV = PROBE_SUPPORT_DIR / "method_c_variable_value.csv"
OUTPUT_METHOD_D_CSV = PROBE_SUPPORT_DIR / "method_d_export_results_csv.csv"
OUTPUT_METHOD_SUMMARY_CSV = PROBE_OUTPUT_DIR / "method_summary.csv"


#%%
# Step 1: Connect to LEAP Results API

APP = connect_results_api(visible=True)


#%%
# Step 2: Inventory resource fuel branches and variable presence

inventory_df = inventory_resource_unmet_variable(
    APP,
    roots=PROBE_ROOTS,
    variable_name=PROBE_VARIABLE,
)
save_probe_table(inventory_df, OUTPUT_INVENTORY_CSV)
print(inventory_df.head(20).to_string(index=False))


#%%
# Step 3: Method A - direct branch ValueRS reads

method_a_df = probe_unmet_via_branch_valuesrs(
    APP,
    inventory_df=inventory_df,
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    years=PROBE_YEARS,
    variable_name=PROBE_VARIABLE,
)
save_probe_table(method_a_df, OUTPUT_METHOD_A_CSV)
print(method_a_df.head(20).to_string(index=False))


#%%
# Step 4: Method B - root branch + fuel filter ValueRS reads

method_b_df = probe_unmet_via_root_filter_valuesrs(
    APP,
    inventory_df=inventory_df,
    roots=PROBE_ROOTS,
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    years=PROBE_YEARS,
    variable_name=PROBE_VARIABLE,
    max_fuels_per_root=PROBE_MAX_FUELS_PER_ROOT,
)
save_probe_table(method_b_df, OUTPUT_METHOD_B_CSV)
print(method_b_df.head(20).to_string(index=False))


#%%
# Step 5: Method C - Variable.Value probes

method_c_df = probe_unmet_via_resultvalue(
    APP,
    inventory_df=inventory_df,
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    years=PROBE_YEARS,
    variable_name=PROBE_VARIABLE,
)
save_probe_table(method_c_df, OUTPUT_METHOD_C_CSV)
print(method_c_df.head(20).to_string(index=False))


#%%
# Step 6: Method D - ExportResultsCSV parsing

method_d_df, exported_csv_paths = probe_unmet_via_exported_csv(
    APP,
    roots=PROBE_ROOTS,
    scenario=PROBE_SCENARIO,
    region=PROBE_REGION,
    variable_name=PROBE_VARIABLE,
    output_dir=PROBE_SUPPORT_DIR,
)
save_probe_table(method_d_df, OUTPUT_METHOD_D_CSV)
print(method_d_df.head(20).to_string(index=False))
print("[INFO] Exported CSV files:")
for item in exported_csv_paths:
    print(f"  - {item}")


#%%
# Step 7: Summary comparison across methods

method_summary_df = build_method_summary(
    {
        "method_a_branch_valuesrs": method_a_df,
        "method_b_root_filter_valuesrs": method_b_df,
        "method_c_variable_value": method_c_df,
        "method_d_export_results_csv": method_d_df,
    }
)
save_probe_table(method_summary_df, OUTPUT_METHOD_SUMMARY_CSV)
print(method_summary_df.to_string(index=False))

print("\n[INFO] Probe run complete.")
print(f"[INFO] Output directory: {_resolve(PROBE_OUTPUT_DIR)}")


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
