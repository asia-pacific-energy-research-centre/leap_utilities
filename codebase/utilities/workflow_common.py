from __future__ import annotations

import re
from pathlib import Path
from typing import Callable, Iterable, Sequence

import pandas as pd

from codebase.functions.leap_core import (
    connect_to_leap,
    create_branches_from_export_file,
    fill_branches_from_export_file,
)

AGGREGATE_ECONOMY_LABELS = {"00_APEC", "ALL_ECONOMIES", "ALL"}


def normalize_economies(economies: str | Iterable[str] | None) -> list[str]:
    """Return a normalized list of economy labels."""
    if economies is None:
        return []
    if isinstance(economies, str):
        text = economies.strip()
        return [text] if text else []
    return [str(value).strip() for value in economies if str(value).strip()]


def resolve_aggregate_economy(
    economies: str | Iterable[str] | None,
    aggregate_label: str | None = None,
    *,
    aggregate_labels: set[str] | None = None,
) -> tuple[bool, str, list[str]]:
    """Return (should_aggregate, aggregate_label, normalized_economies)."""
    normalized = normalize_economies(economies)
    labels = aggregate_labels or AGGREGATE_ECONOMY_LABELS
    if len(normalized) == 1 and normalized[0] in labels:
        return True, normalized[0], normalized
    resolved_label = aggregate_label or "ALL_ECONOMIES"
    return False, resolved_label, normalized

def format_filename_segment(value: str | None) -> str:
    """Return a file-safe string for economy or scenario labels."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    sanitized = re.sub(r"[^A-Za-z0-9_-]+", "_", text)
    return sanitized.strip("_") or text


def normalize_scenarios(scenarios: str | Iterable[str] | None) -> list[str]:
    """Return a list of scenario labels."""
    if scenarios is None:
        return []
    if isinstance(scenarios, str):
        return [scenarios]
    return list(scenarios)


def normalize_workflow_scenarios(
    scenarios: str | Iterable[str] | None,
    default_scenarios: Sequence[str],
) -> list[str]:
    """Return cleaned scenario names for export/import workflow operations."""
    if scenarios is None:
        scenario_values = list(default_scenarios)
    elif isinstance(scenarios, str):
        scenario_values = [scenarios]
    else:
        scenario_values = list(scenarios)
    cleaned: list[str] = []
    seen: set[str] = set()
    for value in scenario_values:
        scenario_name = str(value).strip()
        if not scenario_name or scenario_name in seen:
            continue
        seen.add(scenario_name)
        cleaned.append(scenario_name)
    return cleaned or list(default_scenarios)


def resolve_import_scenarios(
    scenario_list: Sequence[str],
    import_scenario: str | Sequence[str] | None,
    *,
    current_accounts_labels: set[str] | None = None,
) -> list[str]:
    """Return ordered scenario names to import, excluding current-accounts labels."""
    account_labels = current_accounts_labels or {"current accounts", "current account"}
    available_by_lower = {str(name).strip().lower(): str(name) for name in scenario_list}
    default_scenarios = [
        scenario
        for scenario in scenario_list
        if str(scenario).strip().lower() not in account_labels
    ]
    if import_scenario is None:
        if not default_scenarios:
            raise ValueError(
                f"No non-'Current Accounts' scenarios available for import in {list(scenario_list)}."
            )
        return list(default_scenarios)

    if isinstance(import_scenario, str):
        requested_values = [import_scenario]
    else:
        requested_values = list(import_scenario)

    resolved: list[str] = []
    for value in requested_values:
        scenario_name = str(value).strip()
        if not scenario_name:
            continue
        scenario_key = scenario_name.lower()
        if scenario_key in account_labels:
            continue
        if scenario_key not in available_by_lower:
            raise ValueError(
                f"Import scenario '{scenario_name}' is not in exported scenarios: {list(scenario_list)}"
            )
        matched = available_by_lower[scenario_key]
        if matched not in resolved:
            resolved.append(matched)
    if not resolved:
        if not default_scenarios:
            raise ValueError(
                f"No non-'Current Accounts' scenarios available for import in {list(scenario_list)}."
            )
        return list(default_scenarios)
    return resolved


def _format_scenario_segment(
    scenarios: Sequence[str],
    format_segment_fn: Callable[[str], str],
) -> str:
    tokens = [format_segment_fn(segment) for segment in scenarios if segment]
    sanitized = "_".join(token for token in tokens if token)
    return sanitized or "scenarios"


def format_export_filename(
    economy_label: str,
    scenarios: Sequence[str],
    template: str,
    format_segment_fn: Callable[[str], str],
    fallback_template: str | None = None,
) -> str:
    """Return a safe filename for export workbooks."""
    scenario_segment = _format_scenario_segment(scenarios, format_segment_fn)
    economy_segment = format_segment_fn(economy_label)
    try:
        return template.format(economy=economy_segment, scenario=scenario_segment)
    except Exception as exc:
        print(f"Failed to format export filename: {exc}")
        fallback = fallback_template or template
        try:
            return fallback.format(economy=economy_segment, scenario=scenario_segment)
        except Exception:
            return fallback


def build_workflow_export_filename(
    economy_label: str,
    scenarios: str | Iterable[str] | None,
    template: str,
    format_segment_fn: Callable[[str], str] = format_filename_segment,
    fallback_template: str | None = None,
) -> str:
    """Return a filename that includes economy and scenario(s)."""
    scenario_list = normalize_scenarios(scenarios)
    return format_export_filename(
        economy_label,
        scenario_list,
        template,
        format_segment_fn,
        fallback_template=fallback_template,
    )


def read_export_column_values(
    export_path: Path,
    sheet_name: str,
    column: str,
) -> list[str]:
    """Return unique values in a column while preserving order."""
    for header in (2, 0):
        try:
            df = pd.read_excel(
                export_path, sheet_name=sheet_name, header=header, usecols=[column]
            )
        except Exception:
            continue
        if column not in df.columns:
            continue
        seen: list[str] = []
        for value in df[column].dropna().astype(str):
            if value not in seen:
                seen.append(value)
        if seen:
            return seen
    return []


def list_export_scenarios(export_path: Path, sheet_name: str) -> list[str]:
    """Return the Scenario column values in declaration order."""
    return read_export_column_values(export_path, sheet_name, "Scenario")


def validate_export_region(export_path: Path, sheet_name: str, region: str) -> None:
    """Ensure the workbook contains the requested region."""
    regions = read_export_column_values(export_path, sheet_name, "Region")
    if not regions:
        print(f"Warning: 'Region' column missing from {export_path.name}; skipping region check.")
        return
    if region not in regions:
        raise ValueError(
            f"Requested region '{region}' not present in {export_path.name}; available: {regions}"
        )


def find_latest_export_workbook(
    directory: Path | str,
    prefix: str,
    filename: str | None = None,
) -> Path:
    """Locate a workbook by explicit name or latest matching prefix."""
    directory_path = Path(directory)
    if filename:
        candidate = directory_path / filename
        if candidate.exists():
            return candidate
        raise FileNotFoundError(f"Specified export missing: {candidate}")
    matches = sorted(directory_path.glob(f"{prefix}*.xlsx"))
    if not matches:
        raise FileNotFoundError(f"No exports detected in {directory_path}")
    return matches[-1]


def import_workbook_to_leap(
    export_path: Path,
    sheet_name: str,
    scenario: str | None,
    region: str | None,
    create_branches: bool = True,
    fill_branches: bool = True,
    include_current_accounts: bool = True,
    default_branch_type: tuple | None = None,
    branch_type_mapping: dict | None = None,
    branch_root: str | None = None,
    branch_path_col: str | None = None,
    raise_on_missing_branch: bool = False,
) -> Path:
    """Connect to LEAP, validate the workbook, and fill branches."""
    available = list_export_scenarios(export_path, sheet_name)
    scenario_choice = scenario or (available[0] if available else None)
    if scenario_choice and scenario_choice not in available:
        raise ValueError(
            f"Scenario '{scenario_choice}' not found in {export_path.name}; options {available}"
        )
    if region:
        validate_export_region(export_path, sheet_name, region)

    leap_conn = connect_to_leap()
    if leap_conn is None:
        raise RuntimeError("Unable to connect to LEAP.")
    if create_branches:
        create_kwargs = {
            "sheet_name": sheet_name,
            "branch_root": branch_root,
            "branch_type_mapping": branch_type_mapping,
            "default_branch_type": default_branch_type,
            "RAISE_ERROR_ON_FAILED_BRANCH_CREATION": raise_on_missing_branch,
        }
        if branch_path_col is not None:
            create_kwargs["branch_path_col"] = branch_path_col
        create_branches_from_export_file(
            leap_conn,
            export_path,
            **create_kwargs,
        )
    if fill_branches:
        fill_branches_from_export_file(
            leap_conn,
            export_path,
            sheet_name=sheet_name,
            scenario=scenario_choice,
            region=region,
            HANDLE_CURRENT_ACCOUNTS_TOO=include_current_accounts,
        )
    return export_path
