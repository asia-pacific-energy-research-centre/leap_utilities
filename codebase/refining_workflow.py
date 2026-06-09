#%%
"""
Build and optionally import LEAP refining branches from export data.

This workflow prepares the refining export workbook, optionally remaps
transformation fuels, and imports branch structure and time-series values into
LEAP. Some refining variables are intentionally skipped in code and should be
reviewed manually where needed.
"""

# Refining mapping workflow using code to create and fill branches from an export file.
import sys
from pathlib import Path
from typing import Sequence

import pandas as pd

from codebase.utilities.master_config import config_table_exists, read_config_table

# Allow repo root on sys.path so code imports resolve without install
REPO_ROOT = Path(__file__).resolve().parents[1]
if REPO_ROOT.exists() and str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.configuration.config import (
    BRANCH_DEMAND_CATEGORY,
    BRANCH_DEMAND_TECHNOLOGY,
)
from codebase.functions.leap_core import (
    fill_branches_from_export_file,
    create_branches_from_export_file,
    connect_to_leap,
)
from codebase.functions.analysis_input_write_dispatcher import (
    dispatch_analysis_input_write,
    get_analysis_input_write_mode,
)
from codebase.functions.leap_exports import list_scenarios as list_export_scenarios
from codebase.functions.leap_excel_io import (
    read_export_sheet,
    write_export_sheet,
)
from codebase.functions.transformation_fuel_remap import remap_transformation_export_fuels
from codebase.utilities.output_paths import STANDALONE_LEAP_EXPORTS_ROOT

WRITE_MODE = get_analysis_input_write_mode()
L = None
if WRITE_MODE == "api":
    L = connect_to_leap()

CREATE_BRANCHES_FROM_EXPORT_FILE = True

# Define parameters
leap_export_filename = "../data/refining model export.xlsx"
ECONOMY = "20_USA"
BASE_YEAR = 2022
SCENARIOS = ["Reference", "Target", "Current Accounts"]
SCENARIO = "Target"  # Used only when FILL_ALL_SCENARIOS=False.
SOURCE_SCENARIO_FOR_MISSING = {
    "Target": "Reference",
}
FILL_ALL_SCENARIOS = True
CREATE_BRANCHES_FOR_ALL_SCENARIOS = True
REGION = "United States"
sheet_name = "refining"

CURRENT_ACCOUNT_LABELS = {"current accounts", "current account"}


def _safe_filename_segment(value: str) -> str:
    return "".join(ch if ch.isalnum() or ch in {"_", "-"} else "_" for ch in str(value).strip())


def _normalize_scenarios(scenarios: Sequence[str] | None) -> list[str]:
    cleaned: list[str] = []
    seen: set[str] = set()
    for scenario_name in scenarios or []:
        scenario_text = str(scenario_name).strip()
        if not scenario_text:
            continue
        scenario_key = scenario_text.lower()
        if scenario_key in seen:
            continue
        seen.add(scenario_key)
        cleaned.append(scenario_text)
    return cleaned


def _resolve_fill_scenarios() -> list[str]:
    configured = [
        scenario
        for scenario in _normalize_scenarios(SCENARIOS)
        if scenario.lower() not in CURRENT_ACCOUNT_LABELS
    ]
    if not FILL_ALL_SCENARIOS:
        selected = str(SCENARIO).strip()
        return [selected] if selected else configured[:1]
    return configured


def _ensure_export_contains_scenarios(
    export_filename: str | Path,
    export_sheet_name: str,
    target_scenarios: Sequence[str],
    source_scenario_for_missing: dict[str, str] | None = None,
) -> None:
    """Ensure workbook contains each target scenario, copying fallback rows if needed."""
    source_map = {str(key).lower(): str(value) for key, value in (source_scenario_for_missing or {}).items()}
    target_list = _normalize_scenarios(target_scenarios)
    if not target_list:
        return

    header_rows, data, columns = read_export_sheet(export_filename, export_sheet_name)
    if "Scenario" not in data.columns:
        raise ValueError(
            f"Scenario column missing from '{export_filename}' (sheet '{export_sheet_name}')."
        )
    working = data.copy()
    working["Scenario"] = working["Scenario"].astype("string").fillna("").str.strip()
    working["_scenario_norm"] = working["Scenario"].str.lower()
    non_ca_rows = working[~working["_scenario_norm"].isin(CURRENT_ACCOUNT_LABELS)].copy()
    ca_rows = working[working["_scenario_norm"].isin(CURRENT_ACCOUNT_LABELS)].copy()
    if non_ca_rows.empty:
        raise ValueError(
            f"No non-'Current Accounts' rows found in '{export_filename}' (sheet '{export_sheet_name}')."
        )
    first_available_source_norm = str(non_ca_rows["_scenario_norm"].iloc[0]).strip()
    rebuilt_non_ca_rows: list[pd.DataFrame] = []
    for target_scenario in target_list:
        target_norm = target_scenario.lower()
        target_rows = non_ca_rows[non_ca_rows["_scenario_norm"] == target_norm].copy()
        if target_rows.empty:
            source_scenario = source_map.get(target_norm, target_scenario)
            source_norm = str(source_scenario).strip().lower()
            source_rows = non_ca_rows[non_ca_rows["_scenario_norm"] == source_norm].copy()
            if source_rows.empty:
                source_rows = non_ca_rows[
                    non_ca_rows["_scenario_norm"] == first_available_source_norm
                ].copy()
                source_scenario = str(source_rows["Scenario"].iloc[0])
            target_rows = source_rows
            print(
                f"[INFO] Scenario '{target_scenario}' missing in sheet '{export_sheet_name}'; "
                f"copying rows from '{source_scenario}'."
            )
        target_rows["Scenario"] = target_scenario
        if "ScenarioID" in target_rows.columns:
            target_rows["ScenarioID"] = pd.NA
        target_rows["_scenario_norm"] = target_norm
        rebuilt_non_ca_rows.append(target_rows)
    rebuilt_non_ca = (
        pd.concat(rebuilt_non_ca_rows, ignore_index=True)
        if rebuilt_non_ca_rows
        else non_ca_rows.iloc[0:0].copy()
    )
    combined = pd.concat([rebuilt_non_ca, ca_rows], ignore_index=True)
    combined = combined.drop(columns=["_scenario_norm"], errors="ignore")
    combined = combined.reindex(columns=columns)
    write_export_sheet(
        path=export_filename,
        sheet_name=export_sheet_name,
        header_rows=header_rows,
        columns=columns,
        data=combined,
    )


def _discover_fill_scenarios(
    export_filename: str,
    export_sheet_name: str,
    desired_scenarios: Sequence[str],
) -> list[str]:
    desired = _normalize_scenarios(desired_scenarios)
    if not desired:
        return []

    try:
        raw_scenarios = list_export_scenarios(Path(export_filename), sheet_name=export_sheet_name)
    except Exception as exc:
        print(
            f"[WARN] Failed to read scenarios from '{export_filename}' (sheet '{export_sheet_name}'): {exc}"
        )
        raw_scenarios = []

    available_by_key: dict[str, str] = {}
    for scenario_name in raw_scenarios:
        scenario_text = str(scenario_name).strip()
        if not scenario_text:
            continue
        scenario_key = scenario_text.lower()
        if scenario_key in CURRENT_ACCOUNT_LABELS:
            continue
        if scenario_key not in available_by_key:
            available_by_key[scenario_key] = scenario_text

    resolved = [available_by_key[name.lower()] for name in desired if name.lower() in available_by_key]
    missing = [name for name in desired if name.lower() not in available_by_key]
    if missing:
        print(f"[WARN] Export is missing configured scenarios: {missing}")
    return resolved


# Optional: remap refining fuels to ESTO product names before creating/filling branches.
REMAP_FUELS = True
MAPPING_CSV_PATH = "../config/refining_fuel_mapping.csv"
NINTH_TO_ESTO_PAIRS_PATH = "../config/ninth_pairs_to_esto_pairs.xlsx"
REMAP_OUTPUT_PATH = (
    STANDALONE_LEAP_EXPORTS_ROOT
    / f"refining_export_remapped_{_safe_filename_segment(ECONOMY)}_{_safe_filename_segment('all_scenarios' if FILL_ALL_SCENARIOS else SCENARIO)}.xlsx"
)
REMAP_REPORT_PATH = "../intermediate_data/refining_fuel_remap_report.csv"

configured_fill_scenarios = _resolve_fill_scenarios()

if REMAP_FUELS:
    report = remap_transformation_export_fuels(
        input_path=leap_export_filename,
        output_path=REMAP_OUTPUT_PATH,
        mapping_csv_path=MAPPING_CSV_PATH,
        pairs_path=NINTH_TO_ESTO_PAIRS_PATH,
        sheet_name=sheet_name,
        branch_root="Transformation\\Oil Refining",
        fuel_group_labels=("Output Fuels", "Feedstock Fuels", "Auxiliary Fuels"),
        clear_id_columns=True,
        report_path=REMAP_REPORT_PATH,
    )
    if report.get("unmapped_fuels"):
        print(f"[WARN] Unmapped fuels: {sorted(set(report['unmapped_fuels']))}")
    if report.get("ambiguous_mappings"):
        print("[WARN] Ambiguous mappings found; see report for details.")
    leap_export_filename = REMAP_OUTPUT_PATH

if configured_fill_scenarios:
    _ensure_export_contains_scenarios(
        export_filename=leap_export_filename,
        export_sheet_name=sheet_name,
        target_scenarios=configured_fill_scenarios,
        source_scenario_for_missing=SOURCE_SCENARIO_FOR_MISSING,
    )

#%%
if WRITE_MODE == "workbook" and (CREATE_BRANCHES_FROM_EXPORT_FILE or FILL_BRANCHES_FROM_EXPORT_FILE):
    dispatch_analysis_input_write(
        export_path=Path(leap_export_filename),
        sheet_name=sheet_name,
        scenario=configured_fill_scenarios[0] if configured_fill_scenarios else None,
        region=REGION,
        context_label="refining_workflow",
    )

#%%
if CREATE_BRANCHES_FROM_EXPORT_FILE and WRITE_MODE == "api":
    create_scenarios = _discover_fill_scenarios(
        leap_export_filename,
        sheet_name,
        configured_fill_scenarios,
    )
    if not create_scenarios:
        raise ValueError(
            "No configured scenarios available to create branches. Check SCENARIOS/SCENARIO settings and export workbook Scenario column."
        )
    scenario_filter = None if CREATE_BRANCHES_FOR_ALL_SCENARIOS else create_scenarios[0]
    if scenario_filter is None:
        print(f"[INFO] Creating refining branches for configured scenarios: {create_scenarios}")
    else:
        print(f"[INFO] Creating refining branches for scenario '{scenario_filter}'.")
    create_branches_from_export_file(
        L,
        leap_export_filename,
        sheet_name=sheet_name,
        branch_path_col="Branch Path",
        scenario=scenario_filter,
        region=REGION,
        branch_type_mapping=None,
        default_branch_type=(BRANCH_DEMAND_CATEGORY, BRANCH_DEMAND_CATEGORY, BRANCH_DEMAND_TECHNOLOGY),
        RAISE_ERROR_ON_FAILED_BRANCH_CREATION=True,
    )

#%%
FILL_BRANCHES_FROM_EXPORT_FILE = True
HANDLE_CURRENT_ACCOUNTS_TOO = True
SET_UNITS = True
SKIP_VARIABLES = {
    "Dispatchable",
    "Optimize",
    "Surplus Rule",
    "Shortfall Rule",
    "Priority Output",
    "Dispatch Rule",
    "First Simulation Year",
}
if FILL_BRANCHES_FROM_EXPORT_FILE and WRITE_MODE == "api":
    scenarios_to_fill = _discover_fill_scenarios(
        leap_export_filename,
        sheet_name,
        configured_fill_scenarios,
    )
    if not scenarios_to_fill:
        raise ValueError(
            "No scenarios available to fill. Set SCENARIO or check the export workbook Scenario column."
        )
    print(f"[INFO] Filling refining data for scenarios: {scenarios_to_fill}")
    for i, scenario_name in enumerate(scenarios_to_fill):
        include_current_accounts = HANDLE_CURRENT_ACCOUNTS_TOO and i == 0
        fill_branches_from_export_file(
            L,
            leap_export_filename,
            sheet_name=sheet_name,
            scenario=scenario_name,
            region=REGION,
            RAISE_ERROR_ON_FAILED_SET=True,
            SET_UNITS=SET_UNITS,
            HANDLE_CURRENT_ACCOUNTS_TOO=include_current_accounts,
            CHECK_STALE_CHILD_BRANCHES=True,
            PROMPT_DELETE_STALE_BRANCHES=True,
            SKIP_VARIABLES=SKIP_VARIABLES,
        )
#%%


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
