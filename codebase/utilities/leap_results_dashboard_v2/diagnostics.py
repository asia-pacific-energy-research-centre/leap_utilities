from __future__ import annotations

from pathlib import Path

import pandas as pd

from codebase import leap_results_dashboard_workflow as v1_workflow
from codebase.utilities.leap_results_dashboard_utils import basic_checks
from codebase.utilities.workflow_outputs import build_workflow_output_layout


def write_diagnostics(
    *,
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    out_dir: Path,
    base_year: int,
    diagnostic_probe_year: int,
    top_diagnostic_rows: int,
) -> dict[str, str | None]:
    artifacts: dict[str, str | None] = {
        "gap_diagnostics": None,
        "mapping_rundown_by_sheet": None,
        "mapping_rundown_details": None,
        "comparison_issue_summary": None,
        "comparison_issue_cause_summary": None,
    }
    layout = build_workflow_output_layout(out_dir)
    generated = v1_workflow._write_gap_and_mapping_diagnostics(
        comparison_long=comparison_long,
        mapping_status=mapping_status,
        diagnostics_dir=layout.diagnostics,
        mapping_dir=layout.mapping,
        base_year=base_year,
    )
    artifacts.update(generated)
    issue_path = v1_workflow._write_issue_summary(
        comparison_long=comparison_long,
        mapping_status=mapping_status,
        diagnostics_dir=layout.diagnostics,
        base_year=base_year,
        probe_year=diagnostic_probe_year,
        top_n=top_diagnostic_rows,
    )
    artifacts["comparison_issue_summary"] = issue_path
    cause_path = out_dir / "comparison_issue_cause_summary.csv"
    if cause_path.exists():
        artifacts["comparison_issue_cause_summary"] = str(cause_path)
    return artifacts


def run_basic_checks(
    sheet_map: pd.DataFrame,
    fuel_aliases: dict[str, dict[str, str]],
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
) -> dict[str, object]:
    return basic_checks(sheet_map, fuel_aliases, comparison_long, mapping_status)
