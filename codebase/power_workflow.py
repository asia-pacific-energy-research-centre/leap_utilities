#%%
"""
INACTIVE. WE GAVE UP ON THIS APPROACH FOR POWER.
Build and optionally import LEAP power-sector branches from export data.

This workflow prepares the power export workbook, validates/remaps fuel labels,
and fills generation-related LEAP branches for the configured scenarios. It also
handles hardcoded variable overrides and write-mode dispatch so power imports can
run through either the LEAP API or workbook workflow.
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path
from typing import Sequence

import pandas as pd

# Allow repo root on sys.path so code imports resolve without install.
REPO_ROOT = Path(__file__).resolve().parents[1]
if REPO_ROOT.exists() and str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.configuration.config import (
    BRANCH_DEMAND_CATEGORY,
    BRANCH_DEMAND_TECHNOLOGY,
)
from codebase.functions.leap_core import (
    connect_to_leap,
    create_branches_from_export_file,
    fill_branches_from_export_file,
    safe_branch_call,
    safe_set_variable,
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
from codebase.functions.leap_labels import clean_fuel_label_for_leap
from codebase.scrapbook.utilities import load_augmented_reference_tables
from codebase.utilities.workflow_common import archive_config_dir_once_per_day
from codebase.utilities.output_paths import STANDALONE_LEAP_EXPORTS_ROOT


CREATE_BRANCHES_FROM_EXPORT_FILE = True
FILL_BRANCHES_FROM_EXPORT_FILE = True
FILL_ALL_SCENARIOS = True
CREATE_BRANCHES_FOR_ALL_SCENARIOS = True
HANDLE_CURRENT_ACCOUNTS_TOO = True
SET_UNITS = True
CHECK_STALE_CHILD_BRANCHES = True
PROMPT_DELETE_STALE_BRANCHES = True
PROMPT_ON_SHARE_MISMATCH = False
NORMALIZE_SHARE_COLUMNS = False

LEAP_EXPORT_FILENAME = "../data/power export.xlsx"
SHEET_NAME = "Export"
ECONOMY = "20_USA"
BASE_YEAR = 2022
SCENARIOS = ["Reference", "Target", "Current Accounts"]
SCENARIO = "Target"  # Used only when FILL_ALL_SCENARIOS=False.
SOURCE_SCENARIO_FOR_MISSING = {
    "Reference": "Optimization",
    "Target": "Optimization",
}
REGION = "United States of America"
ESTO_DATA_PATH = "../data/00APEC_2024_low.csv"
SUBTOTAL_MAPPING_PATH = "../config/ESTO_subtotal_mapping.xlsx"
SYNTHETIC_RULES_PATH = "../config/synthetic_reference_rows.csv"
REFERENCE_CACHE_DIR = REPO_ROOT / "data/.cache/power_reference_tables"

SKIP_VARIABLES = {
    "Dispatchable",
    "Optimize",
    "Surplus Rule",
    "Shortfall Rule",
    "Priority Output",
    "Dispatch Rule",
    "First Simulation Year",
    "Usage Rule",
    "Use Addition Size",
    "Module Costs",
    # "Renewable Target",
    "Minimum Charge",
    "Full Load Hours",#Default 0
    "Minimum Share of Production",#cannot find in leap
    "Annual Storage Carryover",#Default No
    "Hourly Storage Carryover",#Default Yes
    "Seasonal Storage Carryover",#Default No
    "Interest Rate", #Default DiscountRate
}

CURRENT_ACCOUNT_LABELS = {"current accounts", "current account"}
FUEL_GROUP_LABELS = ("Output Fuels", "Feedstock Fuels", "Auxiliary Fuels")
EXCLUDED_BRANCH_LEAF_LABELS_FROM_API = {
    "Carbon Dioxide",
    "Carbon Monoxide",
    "Methane",
    "Nitrogen Oxides",
    "Nitrous Oxide",
    "Non Methane Volatile Organic Compounds",
    "Sulfur Dioxide",
}

# Explicitly configured entries that should be applied in LEAP after fill.
# Schema:
# - scenario (required)
# - branch_path (required)
# - variable (required)
# - expression (required)
# - units (optional)
# - reason (optional)
HARDCODED_VARIABLE_OVERRIDES: list[dict[str, str | None]] = []


def _safe_filename_segment(value: str) -> str:
    return "".join(ch if ch.isalnum() or ch in {"_", "-"} else "_" for ch in str(value).strip())


def _normalize_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _normalize_scenarios(scenarios: Sequence[str] | None) -> list[str]:
    cleaned: list[str] = []
    seen: set[str] = set()
    for scenario_name in scenarios or []:
        scenario_text = _normalize_text(scenario_name)
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
        selected = _normalize_text(SCENARIO)
        return [selected] if selected else configured[:1]
    return configured


def _resolve_repo_path(path_value: str | Path) -> Path:
    path = Path(path_value)
    if path.is_absolute():
        return path
    return (Path(__file__).resolve().parent / path).resolve()


def _prepare_working_export_file(source_path: str | Path, output_path: str | Path) -> Path:
    source = _resolve_repo_path(source_path)
    target = Path(output_path).resolve()
    if not source.exists():
        raise FileNotFoundError(f"Power export workbook not found: {source}")
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)
    return target


def _discover_fill_scenarios(
    export_filename: str | Path,
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
        scenario_text = _normalize_text(scenario_name)
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
    first_available_source_norm = _normalize_text(non_ca_rows["_scenario_norm"].iloc[0]).lower()
    rebuilt_non_ca_rows: list[pd.DataFrame] = []
    for target_scenario in target_list:
        target_norm = target_scenario.lower()
        target_rows = non_ca_rows[non_ca_rows["_scenario_norm"] == target_norm].copy()
        if target_rows.empty:
            source_scenario = source_map.get(target_norm, target_scenario)
            source_norm = _normalize_text(source_scenario).lower()
            source_rows = non_ca_rows[non_ca_rows["_scenario_norm"] == source_norm].copy()
            if source_rows.empty:
                source_rows = non_ca_rows[
                    non_ca_rows["_scenario_norm"] == first_available_source_norm
                ].copy()
                source_scenario = _normalize_text(source_rows["Scenario"].iloc[0])
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


def _extract_level_values(row: pd.Series, level_cols: list[str]) -> list[str]:
    values = [_normalize_text(row.get(col)) for col in level_cols]
    return [value for value in values if value and value.lower() not in {"nan", "none"}]


def _split_branch_path(branch_path: str) -> list[str]:
    return [part.strip() for part in _normalize_text(branch_path).split("\\") if part.strip()]


def _is_branch_excluded_from_api(branch_path: str) -> bool:
    parts = _split_branch_path(branch_path)
    if not parts:
        return False
    if "Feedstock Fuels" not in parts:
        return False
    leaf = _normalize_text(parts[-1])
    return leaf in EXCLUDED_BRANCH_LEAF_LABELS_FROM_API


def _build_api_filtered_export(
    source_export_path: str | Path,
    sheet_name: str,
    api_export_path: str | Path,
) -> tuple[Path, pd.DataFrame]:
    header_rows, data, columns = read_export_sheet(source_export_path, sheet_name)
    if "Branch Path" not in data.columns:
        raise ValueError(
            f"'Branch Path' column missing from '{source_export_path}' (sheet '{sheet_name}')."
        )
    working = data.copy()
    excluded_mask = working["Branch Path"].apply(_is_branch_excluded_from_api)
    excluded_rows = working.loc[excluded_mask].copy()
    filtered_rows = working.loc[~excluded_mask].copy()

    api_export_path_obj = Path(api_export_path).resolve()
    api_export_path_obj.parent.mkdir(parents=True, exist_ok=True)
    write_export_sheet(
        path=api_export_path_obj,
        sheet_name=sheet_name,
        header_rows=header_rows,
        columns=columns,
        data=filtered_rows.reindex(columns=columns),
    )

    if not excluded_rows.empty:
        excluded_labels = sorted(
            {
                _normalize_text(path).split("\\")[-1]
                for path in excluded_rows["Branch Path"].dropna().astype(str).tolist()
                if _normalize_text(path)
            }
        )
        print(
            "[INFO] Excluding "
            f"{len(excluded_rows)} row(s) from API create/fill for manual handling. "
            f"Excluded leaf labels: {excluded_labels}"
        )
    return api_export_path_obj, excluded_rows


def _extract_export_fuel_rows(data: pd.DataFrame, columns: Sequence[str]) -> pd.DataFrame:
    level_cols = [col for col in columns if str(col).startswith("Level ")]
    discovered: dict[tuple[str, str], str] = {}
    for _, row in data.iterrows():
        branch_path = _normalize_text(row.get("Branch Path"))
        if not branch_path:
            continue
        levels = _extract_level_values(row, level_cols)
        if not levels:
            levels = [part.strip() for part in branch_path.split("\\") if part.strip()]
        if not levels:
            continue
        last_index = len(levels) - 1
        for idx, token in enumerate(levels[:-1]):
            if token not in FUEL_GROUP_LABELS:
                continue
            fuel_idx = idx + 1
            if fuel_idx >= len(levels):
                continue
            fuel_label = _normalize_text(levels[fuel_idx])
            if not fuel_label:
                continue
            if fuel_idx != last_index:
                continue
            key = (fuel_label, token)
            if key not in discovered:
                discovered[key] = branch_path

    rows = []
    for (fuel_label, fuel_group), example_path in sorted(discovered.items()):
        rows.append(
            {
                "fuel_label": fuel_label,
                "fuel_group": fuel_group,
                "example_branch_path": example_path,
            }
        )
    return pd.DataFrame(rows)


def _build_cleaned_esto_products_set(esto_data_path: str | Path) -> set[str]:
    esto_path = _resolve_repo_path(esto_data_path)
    if not esto_path.exists():
        raise FileNotFoundError(f"ESTO data file not found: {esto_path}")
    archive_config_dir_once_per_day()
    df, _ = load_augmented_reference_tables(
        esto_path=esto_path,
        ninth_path=REPO_ROOT / "data/merged_file_energy_ALL_20251106.csv",
        subtotal_mapping_path=_resolve_repo_path(SUBTOTAL_MAPPING_PATH),
        synthetic_rules_path=_resolve_repo_path(SYNTHETIC_RULES_PATH),
        cache_dir=REFERENCE_CACHE_DIR,
        apply_esto_subtotal_map=True,
        filter_esto_subtotals_flag=True,
        filter_ninth_subtotals_flag=False,
    )
    df = df[["products"]].copy() if "products" in df.columns else pd.DataFrame(columns=["products"])
    cleaned = {
        clean_fuel_label_for_leap(_normalize_text(value)).strip().lower()
        for value in df["products"].dropna().tolist()
        if _normalize_text(value)
    }
    return {value for value in cleaned if value}


def _validate_export_fuels_against_esto(
    export_path: str | Path,
    sheet_name: str,
    esto_data_path: str | Path,
    report_path: str | Path | None = None,
) -> pd.DataFrame:
    _, data, columns = read_export_sheet(export_path, sheet_name)
    fuel_rows = _extract_export_fuel_rows(data, columns)
    esto_cleaned = _build_cleaned_esto_products_set(esto_data_path)

    if fuel_rows.empty:
        report_df = pd.DataFrame(
            columns=["fuel_label", "fuel_group", "matched_cleaned_esto", "status", "example_branch_path"]
        )
    else:
        report_df = fuel_rows.copy()
        report_df["matched_cleaned_esto"] = report_df["fuel_label"].apply(
            lambda label: clean_fuel_label_for_leap(_normalize_text(label)).strip().lower() in esto_cleaned
        )
        report_df["status"] = report_df["matched_cleaned_esto"].map(
            lambda matched: "matched" if bool(matched) else "unmapped"
        )
        report_df = report_df[
            ["fuel_label", "fuel_group", "matched_cleaned_esto", "status", "example_branch_path"]
        ]

    if report_path is not None:
        report_path_obj = Path(report_path)
        report_path_obj.parent.mkdir(parents=True, exist_ok=True)
        report_df.to_csv(report_path_obj, index=False)

    unmapped = sorted(report_df.loc[report_df["status"] == "unmapped", "fuel_label"].unique())
    if unmapped:
        print(f"[WARN] Power export includes fuels without cleaned ESTO matches: {unmapped}")
    else:
        print("[INFO] All power export fuels have cleaned ESTO matches.")
    return report_df


def _build_preconfigured_skip_audit_rows(
    export_data: pd.DataFrame,
    scenario_names: Sequence[str],
    skip_variables: set[str],
) -> list[dict[str, str]]:
    if export_data.empty or not skip_variables:
        return []
    scenario_keys = {name.strip().lower() for name in scenario_names if _normalize_text(name)}
    if not scenario_keys:
        return []

    rows: list[dict[str, str]] = []
    seen: set[tuple[str, str, str]] = set()
    for _, row in export_data.iterrows():
        scenario = _normalize_text(row.get("Scenario"))
        variable = _normalize_text(row.get("Variable"))
        branch_path = _normalize_text(row.get("Branch Path"))
        if not scenario or not variable or not branch_path:
            continue
        if scenario.strip().lower() not in scenario_keys:
            continue
        if variable.strip().lower() not in skip_variables:
            continue
        key = (scenario, branch_path, variable)
        if key in seen:
            continue
        seen.add(key)
        rows.append(
            {
                "scenario": scenario,
                "branch_path": branch_path,
                "variable": variable,
                "action": "preconfigured_skip",
                "reason": "Variable listed in SKIP_VARIABLES.",
            }
        )
    return rows


def _validate_hardcoded_overrides(
    overrides: Sequence[dict[str, str | None]],
) -> list[dict[str, str | None]]:
    required_fields = ("scenario", "branch_path", "variable", "expression")
    validated: list[dict[str, str | None]] = []
    for index, entry in enumerate(overrides):
        if not isinstance(entry, dict):
            raise ValueError(f"HARDCODED override at index {index} must be a dict.")
        normalized_entry: dict[str, str | None] = {}
        for field in required_fields:
            value = _normalize_text(entry.get(field))
            if not value:
                raise ValueError(
                    f"HARDCODED override at index {index} is missing required field '{field}'."
                )
            normalized_entry[field] = value
        units_value = _normalize_text(entry.get("units"))
        reason_value = _normalize_text(entry.get("reason"))
        normalized_entry["units"] = units_value or None
        normalized_entry["reason"] = reason_value or None
        validated.append(normalized_entry)
    return validated


def _activate_scenario(L, scenario_name: str) -> bool:
    if not scenario_name:
        return True
    target = scenario_name.strip().lower()
    try:
        scenarios = L.Scenarios
        count = int(scenarios.Count)
    except Exception as exc:
        print(f"[WARN] Could not list LEAP scenarios: {exc}")
        return False

    resolved_name = None
    if target in CURRENT_ACCOUNT_LABELS:
        for idx in range(1, count + 1):
            try:
                item = scenarios.Item(idx)
                if bool(item.IsCurrentAccounts):
                    resolved_name = str(item.Name)
                    break
            except Exception:
                continue
    else:
        for idx in range(1, count + 1):
            try:
                item = scenarios.Item(idx)
                name = _normalize_text(item.Name)
                if name.lower() == target:
                    resolved_name = name
                    break
            except Exception:
                continue
        if resolved_name is None:
            for idx in range(1, count + 1):
                try:
                    item = scenarios.Item(idx)
                    abbreviation = _normalize_text(item.Abbreviation)
                    if abbreviation.lower() == target:
                        resolved_name = _normalize_text(item.Name)
                        break
                except Exception:
                    continue
    if resolved_name is None:
        print(f"[WARN] Scenario '{scenario_name}' not found in LEAP.")
        return False
    try:
        L.ActiveScenario = resolved_name
        return True
    except Exception as exc:
        print(f"[WARN] Failed to activate LEAP scenario '{resolved_name}': {exc}")
        return False


def _activate_region(L, region_name: str) -> bool:
    if not region_name:
        return True
    target = region_name.strip().lower()
    try:
        regions = L.Regions
        if regions.Exists(region_name):
            L.ActiveRegion = region_name
            return True
        count = int(regions.Count)
    except Exception as exc:
        print(f"[WARN] Could not list LEAP regions: {exc}")
        return False

    resolved_name = None
    for idx in range(1, count + 1):
        try:
            item = regions.Item(idx)
            name = _normalize_text(item.Name)
            if name.lower() == target:
                resolved_name = name
                break
        except Exception:
            continue
    if resolved_name is None and count == 1:
        try:
            resolved_name = _normalize_text(regions.Item(1).Name)
            print(
                f"[WARN] Region '{region_name}' not found; using only LEAP region '{resolved_name}'."
            )
        except Exception:
            resolved_name = None
    if resolved_name is None:
        print(f"[WARN] Region '{region_name}' not found in LEAP.")
        return False
    try:
        L.ActiveRegion = resolved_name
        return True
    except Exception as exc:
        print(f"[WARN] Failed to activate LEAP region '{resolved_name}': {exc}")
        return False


def _apply_hardcoded_overrides_for_scenario(
    L,
    overrides: Sequence[dict[str, str | None]],
    scenario_name: str,
    region_name: str,
) -> list[dict[str, str]]:
    results: list[dict[str, str]] = []
    applicable = [
        entry
        for entry in overrides
        if _normalize_text(entry.get("scenario")).lower() in {scenario_name.lower(), "all", "*"}
    ]
    if not applicable:
        return results

    if not _activate_scenario(L, scenario_name):
        for entry in applicable:
            results.append(
                {
                    "scenario": scenario_name,
                    "branch_path": _normalize_text(entry.get("branch_path")),
                    "variable": _normalize_text(entry.get("variable")),
                    "expression": _normalize_text(entry.get("expression")),
                    "units": _normalize_text(entry.get("units")),
                    "reason": _normalize_text(entry.get("reason")),
                    "status": "failed",
                    "message": f"Scenario '{scenario_name}' could not be activated in LEAP.",
                }
            )
        return results
    if not _activate_region(L, region_name):
        for entry in applicable:
            results.append(
                {
                    "scenario": scenario_name,
                    "branch_path": _normalize_text(entry.get("branch_path")),
                    "variable": _normalize_text(entry.get("variable")),
                    "expression": _normalize_text(entry.get("expression")),
                    "units": _normalize_text(entry.get("units")),
                    "reason": _normalize_text(entry.get("reason")),
                    "status": "failed",
                    "message": f"Region '{region_name}' could not be activated in LEAP.",
                }
            )
        return results

    for entry in applicable:
        branch_path = _normalize_text(entry.get("branch_path"))
        variable_name = _normalize_text(entry.get("variable"))
        expression = _normalize_text(entry.get("expression"))
        units_value = _normalize_text(entry.get("units")) or None
        reason_value = _normalize_text(entry.get("reason"))

        branch = safe_branch_call(
            L,
            branch_path,
            AUTO_SET_MISSING_BRANCHES=False,
            THROW_ERROR_ON_MISSING=False,
        )
        if branch is None:
            results.append(
                {
                    "scenario": scenario_name,
                    "branch_path": branch_path,
                    "variable": variable_name,
                    "expression": expression,
                    "units": units_value or "",
                    "reason": reason_value,
                    "status": "skipped_missing_branch",
                    "message": "Branch not found in LEAP.",
                }
            )
            continue
        try:
            variable_obj = branch.Variable(variable_name)
        except Exception:
            variable_obj = None
        if variable_obj is None:
            results.append(
                {
                    "scenario": scenario_name,
                    "branch_path": branch_path,
                    "variable": variable_name,
                    "expression": expression,
                    "units": units_value or "",
                    "reason": reason_value,
                    "status": "skipped_missing_variable",
                    "message": "Variable not found on branch.",
                }
            )
            continue

        ok = safe_set_variable(
            L,
            branch,
            variable_name,
            expression,
            unit_name=units_value,
            context=branch_path,
        )
        status = "applied" if ok else "failed"
        message = "" if ok else "safe_set_variable returned False."
        results.append(
            {
                "scenario": scenario_name,
                "branch_path": branch_path,
                "variable": variable_name,
                "expression": expression,
                "units": units_value or "",
                "reason": reason_value,
                "status": status,
                "message": message,
            }
        )
    return results


def _hardcoded_results_to_audit_rows(results: Sequence[dict[str, str]]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for item in results:
        status = _normalize_text(item.get("status"))
        message = _normalize_text(item.get("message"))
        reason = _normalize_text(item.get("reason"))
        combined_reason = reason if reason else ""
        if message:
            combined_reason = f"{combined_reason} {message}".strip()
        rows.append(
            {
                "scenario": _normalize_text(item.get("scenario")),
                "branch_path": _normalize_text(item.get("branch_path")),
                "variable": _normalize_text(item.get("variable")),
                "action": f"hardcoded_{status}",
                "reason": combined_reason,
            }
        )
    return rows


def _write_audit_report(rows: Sequence[dict[str, str]], path: str | Path) -> None:
    columns = ["scenario", "branch_path", "variable", "action", "reason"]
    output = pd.DataFrame(rows, columns=columns)
    report_path = Path(path)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    output.to_csv(report_path, index=False)


def _write_hardcoded_report(rows: Sequence[dict[str, str]], path: str | Path) -> None:
    columns = [
        "scenario",
        "branch_path",
        "variable",
        "expression",
        "units",
        "reason",
        "status",
        "message",
    ]
    output = pd.DataFrame(rows, columns=columns)
    report_path = Path(path)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    output.to_csv(report_path, index=False)


def _excluded_rows_to_audit_rows(excluded_rows: pd.DataFrame) -> list[dict[str, str]]:
    if excluded_rows is None or excluded_rows.empty:
        return []
    rows: list[dict[str, str]] = []
    seen: set[tuple[str, str, str]] = set()
    for _, item in excluded_rows.iterrows():
        scenario = _normalize_text(item.get("Scenario"))
        branch_path = _normalize_text(item.get("Branch Path"))
        variable = _normalize_text(item.get("Variable"))
        if not branch_path:
            continue
        key = (scenario, branch_path, variable)
        if key in seen:
            continue
        seen.add(key)
        rows.append(
            {
                "scenario": scenario,
                "branch_path": branch_path,
                "variable": variable,
                "action": "excluded_branch_manual",
                "reason": (
                    "Excluded from API create/fill by EXCLUDED_BRANCH_LEAF_LABELS_FROM_API; "
                    "set manually in LEAP."
                ),
            }
        )
    return rows


def run_power_workflow() -> dict[str, object]:
    scenario_token = "all_scenarios" if FILL_ALL_SCENARIOS else SCENARIO
    prepared_export_path = (
        STANDALONE_LEAP_EXPORTS_ROOT
        / (
            f"power_export_prepared_{_safe_filename_segment(ECONOMY)}_"
            f"{_safe_filename_segment(scenario_token)}.xlsx"
        )
    )
    api_filtered_export_path = (
        STANDALONE_LEAP_EXPORTS_ROOT
        / (
            f"power_export_api_filtered_{_safe_filename_segment(ECONOMY)}_"
            f"{_safe_filename_segment(scenario_token)}.xlsx"
        )
    )
    fuel_report_path = REPO_ROOT / "intermediate_data" / "power_fuel_validation_report.csv"
    hardcoded_report_path = REPO_ROOT / "intermediate_data" / "power_hardcoded_values_report.csv"
    fill_audit_path = REPO_ROOT / "intermediate_data" / "power_fill_audit_report.csv"

    export_path = _prepare_working_export_file(
        source_path=LEAP_EXPORT_FILENAME,
        output_path=prepared_export_path,
    )

    configured_fill_scenarios = _resolve_fill_scenarios()
    if configured_fill_scenarios:
        _ensure_export_contains_scenarios(
            export_filename=export_path,
            export_sheet_name=SHEET_NAME,
            target_scenarios=configured_fill_scenarios,
            source_scenario_for_missing=SOURCE_SCENARIO_FOR_MISSING,
        )

    fuel_report_df = _validate_export_fuels_against_esto(
        export_path=export_path,
        sheet_name=SHEET_NAME,
        esto_data_path=ESTO_DATA_PATH,
        report_path=fuel_report_path,
    )
    api_export_path, excluded_rows = _build_api_filtered_export(
        source_export_path=export_path,
        sheet_name=SHEET_NAME,
        api_export_path=api_filtered_export_path,
    )

    validated_overrides = _validate_hardcoded_overrides(HARDCODED_VARIABLE_OVERRIDES)
    write_mode = get_analysis_input_write_mode()
    requires_leap = (
        CREATE_BRANCHES_FROM_EXPORT_FILE
        or FILL_BRANCHES_FROM_EXPORT_FILE
        or bool(validated_overrides)
    )
    L = None
    if requires_leap and write_mode == "workbook":
        dispatch_analysis_input_write(
            export_path=api_export_path,
            sheet_name=SHEET_NAME,
            scenario=configured_fill_scenarios[0] if configured_fill_scenarios else None,
            region=REGION,
            context_label="power_workflow.run_power_workflow",
        )
        if validated_overrides:
            print(
                "[WARN] HARDCODED_VARIABLE_OVERRIDES are not applied in workbook mode. "
                "Apply these values manually in LEAP after workbook import if needed."
            )
    elif requires_leap:
        L = connect_to_leap()
        if L is None:
            raise RuntimeError("Could not connect to LEAP.")

    if CREATE_BRANCHES_FROM_EXPORT_FILE and write_mode == "api":
        create_scenarios = _discover_fill_scenarios(
            api_export_path,
            SHEET_NAME,
            configured_fill_scenarios,
        )
        if not create_scenarios:
            raise ValueError(
                "No configured scenarios available to create branches. "
                "Check SCENARIOS/SCENARIO settings and export workbook Scenario column."
            )
        scenario_filter = None if CREATE_BRANCHES_FOR_ALL_SCENARIOS else create_scenarios[0]
        if scenario_filter is None:
            print(f"[INFO] Creating power branches for configured scenarios: {create_scenarios}")
        else:
            print(f"[INFO] Creating power branches for scenario '{scenario_filter}'.")
        create_branches_from_export_file(
            L,
            api_export_path,
            sheet_name=SHEET_NAME,
            branch_path_col="Branch Path",
            scenario=scenario_filter,
            region=REGION,
            branch_type_mapping=None,
            default_branch_type=(BRANCH_DEMAND_CATEGORY, BRANCH_DEMAND_CATEGORY, BRANCH_DEMAND_TECHNOLOGY),
            RAISE_ERROR_ON_FAILED_BRANCH_CREATION=True,
        )

    audit_rows: list[dict[str, str]] = []
    hardcoded_rows: list[dict[str, str]] = []

    if FILL_BRANCHES_FROM_EXPORT_FILE and write_mode == "api":
        _, export_data, _ = read_export_sheet(api_export_path, SHEET_NAME)
        scenarios_to_fill = _discover_fill_scenarios(
            api_export_path,
            SHEET_NAME,
            configured_fill_scenarios,
        )
        if not scenarios_to_fill:
            raise ValueError(
                "No scenarios available to fill. Set SCENARIO or check the export workbook Scenario column."
            )
        print(f"[INFO] Filling power data for scenarios: {scenarios_to_fill}")

        skip_variables = {value.strip().lower() for value in SKIP_VARIABLES if _normalize_text(value)}
        for i, scenario_name in enumerate(scenarios_to_fill):
            include_current_accounts = HANDLE_CURRENT_ACCOUNTS_TOO and i == 0

            scenario_names_for_skip = [scenario_name]
            if include_current_accounts:
                scenario_names_for_skip.append("Current Accounts")
            audit_rows.extend(
                _build_preconfigured_skip_audit_rows(
                    export_data=export_data,
                    scenario_names=scenario_names_for_skip,
                    skip_variables=skip_variables,
                )
            )

            fill_result = fill_branches_from_export_file(
                L,
                api_export_path,
                sheet_name=SHEET_NAME,
                scenario=scenario_name,
                region=REGION,
                RAISE_ERROR_ON_FAILED_SET=True,
                SET_UNITS=SET_UNITS,
                HANDLE_CURRENT_ACCOUNTS_TOO=include_current_accounts,
                CHECK_STALE_CHILD_BRANCHES=CHECK_STALE_CHILD_BRANCHES,
                PROMPT_DELETE_STALE_BRANCHES=PROMPT_DELETE_STALE_BRANCHES,
                SKIP_VARIABLES=SKIP_VARIABLES,
                PROMPT_ON_SHARE_MISMATCH=PROMPT_ON_SHARE_MISMATCH,
                NORMALIZE_SHARE_COLUMNS=NORMALIZE_SHARE_COLUMNS,
            )
            for branch_path, variable in fill_result.get("success", []):
                audit_rows.append(
                    {
                        "scenario": scenario_name,
                        "branch_path": _normalize_text(branch_path),
                        "variable": _normalize_text(variable),
                        "action": "fill_success",
                        "reason": "",
                    }
                )
            for branch_path, variable in fill_result.get("failed", []):
                audit_rows.append(
                    {
                        "scenario": scenario_name,
                        "branch_path": _normalize_text(branch_path),
                        "variable": _normalize_text(variable),
                        "action": "fill_failed",
                        "reason": "",
                    }
                )
            for branch_path, variable in fill_result.get("skipped", []):
                audit_rows.append(
                    {
                        "scenario": scenario_name,
                        "branch_path": _normalize_text(branch_path),
                        "variable": _normalize_text(variable),
                        "action": "fill_skipped",
                        "reason": "",
                    }
                )

            hardcoded_scenarios = [scenario_name]
            if include_current_accounts:
                hardcoded_scenarios.append("Current Accounts")
            for hardcoded_scenario in hardcoded_scenarios:
                scenario_results = _apply_hardcoded_overrides_for_scenario(
                    L=L,
                    overrides=validated_overrides,
                    scenario_name=hardcoded_scenario,
                    region_name=REGION,
                )
                hardcoded_rows.extend(scenario_results)
                audit_rows.extend(_hardcoded_results_to_audit_rows(scenario_results))

    audit_rows = _excluded_rows_to_audit_rows(excluded_rows) + audit_rows
    _write_hardcoded_report(hardcoded_rows, hardcoded_report_path)
    _write_audit_report(audit_rows, fill_audit_path)

    unmapped_fuels = sorted(
        fuel_report_df.loc[fuel_report_df["status"] == "unmapped", "fuel_label"].unique()
        if not fuel_report_df.empty
        else []
    )

    return {
        "prepared_export_path": str(export_path),
        "api_filtered_export_path": str(api_export_path),
        "fuel_validation_report_path": str(fuel_report_path),
        "hardcoded_report_path": str(hardcoded_report_path),
        "fill_audit_report_path": str(fill_audit_path),
        "unmapped_fuels": unmapped_fuels,
        "excluded_rows_for_manual_entry": int(len(excluded_rows)),
    }


if __name__ == "__main__":
    run_power_workflow()
#%%


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
