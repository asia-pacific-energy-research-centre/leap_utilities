#%%
"""
Load LEAP balance exports into dataframes and build balance comparisons.

This workflow extracts REF/TGT LEAP balance workbooks into normalized long-form
dataframes, maps LEAP balance rows to ESTO products/flows, and compares them to
ESTO base-year and 9th projection series. It also writes coverage, unit,
mapping, ledger, and dashboard artifacts for auditing the comparison.
"""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path
from typing import Sequence

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.utilities.leap_results_dashboard_balance import (  # noqa: E402
    attach_chart_groups_to_dashboard_exposure,
    build_simple_leap_balance_table,
    build_simple_leap_ninth_balance_table,
    build_simple_ninth_balance_table,
    build_balance_comparison,
    load_balance_leap_long,
    render_balance_dashboards,
    write_balance_missing_mapping_candidates,
    write_dashboard_comparator_pair_coverage,
    write_ninth_mapping_data_coverage,
    write_runtime_missing_pair_summary,
)
from codebase.utilities.leap_results_dashboard_utils import _prepare_render_long  # noqa: E402
from codebase.utilities.leap_results_dashboard_v2.comparison_engine import (  # noqa: E402
    build_chart_line_mapping_ledger,
    build_total_component_ledger,
)
from codebase.utilities.leap_results_dashboard_v2.diagnostics import (  # noqa: E402
    run_basic_checks,
    write_diagnostics,
)
from codebase.utilities.leap_results_dashboard_v2.output_writer import write_core_outputs  # noqa: E402
from codebase.utilities.leap_balance_export_resolver import resolve_balance_export_workbook  # noqa: E402
from codebase.utilities.workflow_common import (  # noqa: E402
    WorkflowTimer,
    archive_config_dir_once_per_day,
)
from codebase.utilities.workflow_outputs import build_workflow_output_layout, write_output_manifest  # noqa: E402


#%%
def _resolve(path: Path | str) -> Path:
    raw = str(path).replace("\\", "/")
    drive_match = re.match(r"^([a-zA-Z]):/(.*)$", raw)
    if drive_match:
        drive = drive_match.group(1).lower()
        rest = drive_match.group(2)
        if os.name == "nt":
            return Path(f"{drive.upper()}:/{rest}")
        return Path(f"/mnt/{drive}/{rest}")
    candidate = Path(raw)
    return candidate if candidate.is_absolute() else (REPO_ROOT / candidate)


#%%
def _resolve_existing(path: Path | str, *fallbacks: Path | str) -> Path:
    candidates = [_resolve(path), *(_resolve(fallback) for fallback in fallbacks)]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


#%%
BALANCE_EXPORT_ECONOMY = "20_USA"
REF_BALANCE_EXPORT_DATE_ID: str | None = None
TGT_BALANCE_EXPORT_DATE_ID: str | None = None
REF_WORKBOOK_PATH = resolve_balance_export_workbook(
    economy=BALANCE_EXPORT_ECONOMY,
    scenario="REF",
    date_id=REF_BALANCE_EXPORT_DATE_ID,
)
TGT_WORKBOOK_PATH = resolve_balance_export_workbook(
    economy=BALANCE_EXPORT_ECONOMY,
    scenario="TGT",
    date_id=TGT_BALANCE_EXPORT_DATE_ID,
)

STRUCTURE_CONFIG_PATH = _resolve_existing(
    "config/leap_results_balance_dashboard_structure.json",
    "config/archive/leap_results_balance_dashboard_structure.json",
)
KNOWN_ISSUES_CONFIG_PATH = _resolve("config/leap_results_balance_known_issues.json")
CHART_NAVIGATION_GUIDE_PATH = _resolve("config/leap_comparison_dashboard_template.json")

OUTPUT_DIR = _resolve("outputs/dashboards/leap_results_dashboard_balance/USA")

LEAP_TO_ESTO_MAPPING = (_resolve("config/leap_mappings.xlsx"), "leap_combined_esto")
NINTH_TO_ESTO_MAPPING = (_resolve("config/ninth_pairs_to_esto_pairs.xlsx"), "ninth_pairs_to_esto_pairs")
CODEBOOK_PATH = _resolve("config/sector_fuel_codes_to_names.xlsx")
SHEET_MAP_PATH = _resolve_existing(
    "config/leap_results_sheet_map.csv",
    "config/archive/leap_results_sheet_map.csv",
)
BACKUP_MAPPINGS_PATH = _resolve("config/backup_leap_mappings.xlsx")
EXPLICIT_MAPPINGS_PATH = _resolve("config/leap_results_explicit_mappings.csv")
EXPLICIT_REASSIGNMENTS_PATH = _resolve("config/leap_results_explicit_reassignments.csv")
SYNTHETIC_REFERENCE_ROWS_PATH = _resolve("config/synthetic_reference_rows.csv")

BASE_TABLE_PATH = _resolve("data/00APEC_2025_low_with_subtotals.csv")
PROJECTION_TABLE_PATH = _resolve("data/merged_file_energy_ALL_20251106.csv")

BASE_YEAR = 2022
MAX_OUTPUT_YEAR = 2060
PROJECTION_YEARS: Sequence[int] = tuple(range(BASE_YEAR + 1, MAX_OUTPUT_YEAR + 1))
SCENARIO_MAP = {"Reference": "reference", "Target": "target"}
BASE_ECONOMY = "20USA"
PROJECTION_ECONOMY = "20_USA"

CHART_BACKEND = "plotly"
HIDE_LEAP_ONLY_CHARTS = False
RUN_ESTO_AXIS_WORKFLOW_AFTER_BALANCE = True
ENABLE_WORKFLOW_TIMING = True
WRITE_WORKFLOW_TIMING_CSV = True
WORKFLOW_TIMING_FILENAME = "workflow_stage_timings.csv"
FAIL_ON_UNMAPPED_BALANCE_ROWS = os.getenv("FAIL_ON_UNMAPPED_BALANCE_ROWS", "1").strip().lower() in {
    "1",
    "true",
    "yes",
    "y",
}


#%%
def _load_json(path: Path) -> dict[str, object]:
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def _mapping_workbook(mapping_ref: tuple[Path, str]) -> Path:
    return mapping_ref[0]


def _raise_if_unmapped_balance_rows(runtime_issues: pd.DataFrame, runtime_issues_path: str | None) -> None:
    if runtime_issues.empty:
        return
    if not FAIL_ON_UNMAPPED_BALANCE_ROWS:
        return
    reason_counts = (
        runtime_issues.groupby("reason", dropna=False)
        .size()
        .reset_index(name="row_count")
        .sort_values(["row_count", "reason"], ascending=[False, True])
    )
    counts_text = ", ".join(
        f"{row.reason}: {int(row.row_count)}" for row in reason_counts.itertuples(index=False)
    )
    raise RuntimeError(
        "Unmapped LEAP balance rows remain after writing dashboard outputs. "
        f"See {runtime_issues_path}. Counts: {counts_text}"
    )


def _write_consolidated_issues(
    *,
    runtime_dir: Path,
    diagnostics_dir: Path,
    known_issues: dict[str, object],
    runtime_issues: pd.DataFrame,
    override_report: pd.DataFrame,
) -> dict[str, str | None]:
    runtime_dir.mkdir(parents=True, exist_ok=True)
    diagnostics_dir.mkdir(parents=True, exist_ok=True)
    runtime_path = runtime_dir / "balance_runtime_issues.csv"
    summary_path = diagnostics_dir / "balance_issues_summary.json"
    override_path = runtime_dir / "balance_override_application_report.csv"

    if runtime_issues.empty:
        runtime_issues = pd.DataFrame(
            columns=[
                "reason",
                "details",
                "scenario",
                "year",
                "source_sheet",
                "leap_flow",
                "leap_flow_name",
                "leap_product",
                "leap_product_name",
                "esto_flow",
                "esto_product",
                "value_petajoule",
            ]
        )
    runtime_issues.to_csv(runtime_path, index=False)
    missing_pair_summary_path = runtime_dir / "balance_runtime_missing_pair_summary.xlsx"
    write_runtime_missing_pair_summary(
        runtime_issues=runtime_issues,
        output_path=missing_pair_summary_path,
    )

    if override_report.empty:
        override_report = pd.DataFrame(columns=["override_index", "applied_rows", "match", "set"])
    override_report.to_csv(override_path, index=False)

    issue_counts = (
        runtime_issues.groupby("reason", dropna=False)
        .size()
        .reset_index(name="row_count")
        .sort_values(["row_count", "reason"], ascending=[False, True])
    )
    summary = {
        "known_issues_config": known_issues,
        "runtime_issue_count": int(len(runtime_issues)),
        "runtime_issue_counts": issue_counts.to_dict("records"),
        "override_report_row_count": int(len(override_report)),
    }
    summary_path.write_text(json.dumps(summary, ensure_ascii=True, indent=2), encoding="utf-8")

    return {
        "runtime_issues_csv": str(runtime_path),
        "runtime_missing_pair_summary_xlsx": str(missing_pair_summary_path),
        "issues_summary_json": str(summary_path),
        "override_report_csv": str(override_path),
    }


def _write_mapping_fix_candidates(*, mapping_dir: Path, runtime_issues: pd.DataFrame) -> str:
    mapping_dir.mkdir(parents=True, exist_ok=True)
    path = mapping_dir / "balance_mapping_fix_candidates.csv"
    if runtime_issues.empty:
        pd.DataFrame(
            columns=[
                "leap_flow_name",
                "leap_product_name",
                "leap_flow",
                "leap_product",
                "esto_flow",
                "esto_product",
                "missing_leap_code",
                "missing_esto_code",
                "rows",
                "value_pj",
            ]
        ).to_csv(path, index=False)
        return str(path)

    issues = runtime_issues.copy()
    for col in [
        "leap_flow_name",
        "leap_product_name",
        "leap_flow",
        "leap_product",
        "esto_flow",
        "esto_product",
    ]:
        if col not in issues.columns:
            issues[col] = ""
        issues[col] = issues[col].fillna("").astype(str).str.strip()

    issues["value_petajoule"] = pd.to_numeric(issues.get("value_petajoule"), errors="coerce")
    issues["missing_leap_code"] = issues["leap_flow"].eq("") | issues["leap_product"].eq("")
    issues["missing_esto_code"] = issues["esto_flow"].eq("") | issues["esto_product"].eq("")

    summary = (
        issues.groupby(
            [
                "leap_flow_name",
                "leap_product_name",
                "leap_flow",
                "leap_product",
                "esto_flow",
                "esto_product",
                "missing_leap_code",
                "missing_esto_code",
            ],
            dropna=False,
        )
        .agg(rows=("reason", "size"), value_pj=("value_petajoule", "sum"))
        .reset_index()
        .sort_values(["rows", "value_pj"], ascending=[False, False], kind="mergesort")
    )
    summary.to_csv(path, index=False)
    return str(path)


def _run_esto_axis_workflow_after_balance() -> dict[str, object]:
    from codebase.leap_results_dashboard_balance_estoaxis_workflow import run_workflow as run_esto_axis_workflow

    try:
        result = run_esto_axis_workflow()
        return {f"estoaxis_{key}": value for key, value in result.items()}
    except RuntimeError as exc:
        message = str(exc)
        if "Unmapped LEAP balance rows remain after writing dashboard outputs." not in message:
            raise
        return {
            "estoaxis_runtime_error": message,
            "estoaxis_completed_with_unmapped_rows": True,
        }


def _capture_unmapped_balance_row_error(
    runtime_issues: pd.DataFrame,
    runtime_issues_path: str | None,
    *,
    prefix: str = "",
) -> dict[str, object]:
    try:
        _raise_if_unmapped_balance_rows(runtime_issues, runtime_issues_path)
        return {}
    except RuntimeError as exc:
        key_prefix = f"{prefix}_" if prefix else ""
        return {
            f"{key_prefix}runtime_error": str(exc),
            f"{key_prefix}completed_with_unmapped_rows": True,
        }


def run_workflow() -> dict[str, object]:
    timer = WorkflowTimer("leap_results_dashboard_balance", enabled=ENABLE_WORKFLOW_TIMING)
    archive_config_dir_once_per_day()
    out_dir = _resolve(OUTPUT_DIR)
    layout = build_workflow_output_layout(out_dir)
    timing_path = layout.runtime / WORKFLOW_TIMING_FILENAME

    structure_config = _load_json(STRUCTURE_CONFIG_PATH)
    known_issues = _load_json(KNOWN_ISSUES_CONFIG_PATH)
    timer.lap("setup")

    ingestion = load_balance_leap_long(
        ref_workbook_path=REF_WORKBOOK_PATH,
        tgt_workbook_path=TGT_WORKBOOK_PATH,
        template_sheet="EBal|2060",
        mapping_pairs_path=_mapping_workbook(LEAP_TO_ESTO_MAPPING),
        codebook_path=CODEBOOK_PATH,
        structure_config=structure_config,
        known_issues=known_issues,
        projection_economy=PROJECTION_ECONOMY,
        explicit_pair_mappings_only=True,
    )
    timer.lap("extract and map LEAP balance workbooks")

    comparison = build_balance_comparison(
        leap_long=ingestion["leap_long"],
        mapping_status=ingestion["mapping_status"],
        base_year=BASE_YEAR,
        projection_years=tuple(PROJECTION_YEARS),
        base_economy=BASE_ECONOMY,
        projection_economy=PROJECTION_ECONOMY,
        scenario_map=SCENARIO_MAP,
        sheet_map_path=SHEET_MAP_PATH,
        backup_mappings_path=BACKUP_MAPPINGS_PATH,
        codebook_path=CODEBOOK_PATH,
        canonical_pairs_path=NINTH_TO_ESTO_MAPPING,
        explicit_mappings_path=EXPLICIT_MAPPINGS_PATH,
        explicit_reassignments_path=EXPLICIT_REASSIGNMENTS_PATH,
        synthetic_reference_rows_path=SYNTHETIC_REFERENCE_ROWS_PATH,
        esto_table_path=BASE_TABLE_PATH,
        projection_table_path=PROJECTION_TABLE_PATH,
        known_issues=known_issues,
    )
    timer.lap("build balance comparison")

    comparison_long = comparison["comparison_long"].copy()
    comparison_wide = comparison["comparison_wide"].copy()
    mapping_status = comparison["mapping_status"].copy()
    leap_long = ingestion["leap_long"].copy()

    comparison_long = comparison_long[pd.to_numeric(comparison_long["year"], errors="coerce").le(MAX_OUTPUT_YEAR)].copy()
    comparison_wide = comparison_wide[pd.to_numeric(comparison_wide["year"], errors="coerce").le(MAX_OUTPUT_YEAR)].copy()
    leap_long = leap_long[pd.to_numeric(leap_long["year"], errors="coerce").le(MAX_OUTPUT_YEAR)].copy()

    core_paths = write_core_outputs(
        out_dir=layout.root,
        supporting_dir=layout.supporting,
        comparison_long=comparison_long,
        comparison_wide=comparison_wide,
        mapping_status=mapping_status,
        leap_long=leap_long,
    )
    simple_leap_balance = build_simple_leap_balance_table(leap_long)
    simple_leap_ninth_balance = build_simple_leap_ninth_balance_table(leap_long)
    simple_ninth_balance = build_simple_ninth_balance_table(
        comparison_long=comparison_long,
        mapping_status=mapping_status,
    )
    simple_leap_balance_path = layout.root / "simple_leap_balance_mapped.csv"
    simple_leap_ninth_balance_path = layout.root / "simple_leap_ninth_balance_mapped.csv"
    simple_ninth_balance_path = layout.root / "simple_ninth_balance_mapped.csv"
    simple_leap_balance.to_csv(simple_leap_balance_path, index=False)
    simple_leap_ninth_balance.to_csv(simple_leap_ninth_balance_path, index=False)
    simple_ninth_balance.to_csv(simple_ninth_balance_path, index=False)
    timer.lap("write core and simple balance outputs")

    diagnostics_paths = write_diagnostics(
        comparison_long=comparison_long,
        mapping_status=mapping_status,
        out_dir=layout.root,
        base_year=BASE_YEAR,
        diagnostic_probe_year=min(2030, MAX_OUTPUT_YEAR),
        top_diagnostic_rows=40,
    )
    timer.lap("write diagnostics")
    chart_input = _prepare_render_long(comparison_long)
    chart_line_mapping_ledger = build_chart_line_mapping_ledger(chart_input, mapping_status)
    chart_total_component_ledger = build_total_component_ledger(chart_input, mapping_status)

    dashboard_paths = render_balance_dashboards(
        comparison_long=comparison_long,
        mapping_status=mapping_status,
        structure_config=structure_config,
        output_dir=layout.root,
        chart_backend=CHART_BACKEND,
        hide_leap_only_charts=HIDE_LEAP_ONLY_CHARTS,
        chart_navigation_guide_path=CHART_NAVIGATION_GUIDE_PATH,
    )
    timer.lap("render dashboards")
    chart_line_mapping_ledger = attach_chart_groups_to_dashboard_exposure(
        chart_line_mapping_ledger,
        dashboard_paths.get("chart_group_exposure"),
        dashboard_paths.get("all_chart_groups"),
    )
    chart_group_exposure_df = (
        pd.read_csv(dashboard_paths["chart_group_exposure"], dtype=str).fillna("")
        if dashboard_paths.get("chart_group_exposure")
        else pd.DataFrame()
    )
    all_chart_groups_df = (
        pd.read_csv(dashboard_paths["all_chart_groups"], dtype=str).fillna("")
        if dashboard_paths.get("all_chart_groups")
        else pd.DataFrame()
    )

    chart_line_mapping_path = layout.ledgers / "chart_line_mapping_ledger.csv"
    chart_total_component_path = layout.ledgers / "chart_total_component_ledger.csv"
    chart_line_mapping_ledger.to_csv(chart_line_mapping_path, index=False)
    chart_total_component_ledger.to_csv(chart_total_component_path, index=False)
    timer.lap("write chart ledgers")

    comparator_pair_coverage_xlsx = write_dashboard_comparator_pair_coverage(
        mapping_status=mapping_status,
        dashboard_exposure=chart_line_mapping_ledger,
        chart_group_exposure=chart_group_exposure_df,
        all_chart_groups=all_chart_groups_df,
        base_df=comparison.get("base_df", pd.DataFrame()),
        ninth_df=comparison.get("ninth_df", pd.DataFrame()),
        output_path=layout.coverage / "dashboard_comparator_pair_coverage.xlsx",
        base_economy=BASE_ECONOMY,
        projection_economy=PROJECTION_ECONOMY,
        base_year=BASE_YEAR,
        projection_years=tuple(PROJECTION_YEARS),
        scenarios=tuple(SCENARIO_MAP.values()),
        runtime_issues=ingestion["issues"],
        chart_navigation_guide_path=CHART_NAVIGATION_GUIDE_PATH,
        mapping_workbook_path=_mapping_workbook(LEAP_TO_ESTO_MAPPING),
        mapping_sheet_name=LEAP_TO_ESTO_MAPPING[1],
    )

    issues_paths = _write_consolidated_issues(
        runtime_dir=layout.runtime,
        diagnostics_dir=layout.diagnostics,
        known_issues=known_issues,
        runtime_issues=ingestion["issues"],
        override_report=ingestion["override_report"],
    )
    mapping_fix_candidates_csv = _write_mapping_fix_candidates(
        mapping_dir=layout.mapping,
        runtime_issues=ingestion["issues"],
    )
    missing_mapping_candidates_xlsx = write_balance_missing_mapping_candidates(
        runtime_issues=ingestion["issues"],
        output_path=layout.mapping / "balance_missing_mapping_candidates.xlsx",
        mapping_workbook_path=_mapping_workbook(LEAP_TO_ESTO_MAPPING),
    )
    mapping_inputs = comparison.get("mapping_inputs") or {}
    leap_combined_ninth_mapping = pd.read_excel(
        _mapping_workbook(LEAP_TO_ESTO_MAPPING),
        sheet_name="leap_combined_ninth",
        dtype=str,
    ).fillna("")
    ninth_mapping_data_coverage_xlsx = write_ninth_mapping_data_coverage(
        ninth_df=comparison.get("ninth_df", pd.DataFrame()),
        ninth_mapping_pairs=leap_combined_ninth_mapping,
        output_path=layout.coverage / "ninth_mapping_data_coverage.xlsx",
        projection_economy=PROJECTION_ECONOMY,
        scenarios=tuple(SCENARIO_MAP.values()),
        years=tuple(PROJECTION_YEARS),
    )
    timer.lap("write mapping and coverage checks")

    coverage_path = layout.coverage / "balance_coverage.csv"
    unit_diag_path = layout.coverage / "balance_unit_diagnostics.csv"
    matching_diag_path = layout.coverage / "balance_matching_diagnostics.csv"
    extraction_summary_path = layout.coverage / "balance_extraction_summary.json"

    ingestion["coverage"].to_csv(coverage_path, index=False)
    ingestion["unit_diagnostics"].to_csv(unit_diag_path, index=False)
    ingestion.get("matching_diagnostics", pd.DataFrame()).to_csv(matching_diag_path, index=False)
    extraction_summary_path.write_text(
        json.dumps(ingestion["extraction_summary"], ensure_ascii=True, indent=2),
        encoding="utf-8",
    )

    checks = run_basic_checks(
        mapping_inputs.get("sheet_map", pd.DataFrame()),
        mapping_inputs.get("fuel_aliases", {}),
        comparison_long,
        mapping_status,
    )

    manifest = write_output_manifest(
        out_dir=layout.root,
        primary_outputs={
            "comparison_long": str(layout.root / "comparison_long.csv"),
            "comparison_wide": str(layout.root / "comparison_wide.csv"),
            "mapping_status": str(layout.root / "mapping_status.xlsx"),
            "leap_long": str(layout.root / "leap_long.csv"),
            "simple_leap_balance_mapped": str(simple_leap_balance_path),
            "simple_leap_ninth_balance_mapped": str(simple_leap_ninth_balance_path),
            "simple_ninth_balance_mapped": str(simple_ninth_balance_path),
            "dashboards_dir": str(layout.dashboards),
            "charts_dir": str(layout.charts),
        },
        supporting_outputs={
            "gap_diagnostics": diagnostics_paths.get("gap_diagnostics"),
            "mapping_rundown_by_sheet": diagnostics_paths.get("mapping_rundown_by_sheet"),
            "mapping_rundown_details": diagnostics_paths.get("mapping_rundown_details"),
            "comparison_issue_summary": diagnostics_paths.get("comparison_issue_summary"),
            "comparison_issue_cause_summary": diagnostics_paths.get("comparison_issue_cause_summary"),
            "issues_summary_json": issues_paths.get("issues_summary_json"),
            "dashboard_comparator_pair_coverage_xlsx": comparator_pair_coverage_xlsx,
            "runtime_issues_csv": issues_paths.get("runtime_issues_csv"),
            "runtime_missing_pair_summary_xlsx": issues_paths.get("runtime_missing_pair_summary_xlsx"),
            "issues_summary_json": issues_paths.get("issues_summary_json"),
            "override_report_csv": issues_paths.get("override_report_csv"),
            "mapping_fix_candidates_csv": mapping_fix_candidates_csv,
            "missing_mapping_candidates_xlsx": missing_mapping_candidates_xlsx,
            "ninth_mapping_data_coverage_xlsx": ninth_mapping_data_coverage_xlsx,
            "balance_coverage_csv": str(coverage_path),
            "balance_unit_diagnostics_csv": str(unit_diag_path),
            "balance_matching_diagnostics_csv": str(matching_diag_path),
            "balance_extraction_summary_json": str(extraction_summary_path),
            "chart_line_mapping_ledger": str(chart_line_mapping_path),
            "chart_total_component_ledger": str(chart_total_component_path),
            "workflow_stage_timings_csv": str(timing_path),
        },
        primary_output_descriptions={
            "comparison_long": "Main balance comparison table across LEAP, ESTO, and 9th sources.",
            "comparison_wide": "Wide balance comparison table for quick source-to-source checks.",
            "mapping_status": "Balance mapping workbook showing how each LEAP balance row was mapped.",
            "leap_long": "Extracted LEAP balance rows in normalized long form.",
            "simple_leap_balance_mapped": "Compact LEAP balance table for direct inspection.",
            "simple_leap_ninth_balance_mapped": "Compact LEAP table aligned to the 9th-side balance structure.",
            "simple_ninth_balance_mapped": "Compact 9th balance table aligned to the LEAP balance view.",
            "dashboards_dir": "Rendered balance dashboard HTML pages.",
            "charts_dir": "Balance chart files used by the dashboards.",
        },
        supporting_output_descriptions={
            "gap_diagnostics": "Largest balance gaps between LEAP and comparator sources.",
            "mapping_rundown_by_sheet": "Sheet-level summary of balance mapping completeness.",
            "mapping_rundown_details": "Detailed balance mapping audit workbook.",
            "comparison_issue_summary": "Prioritized comparison issues with gap metrics and hints.",
            "comparison_issue_cause_summary": "Frequency summary of comparison issue categories.",
            "issues_summary_json": "JSON summary of runtime balance issues and override counts.",
            "dashboard_comparator_pair_coverage_xlsx": "Coverage audit for comparator pairs actually exposed in dashboards.",
            "runtime_issues_csv": "Runtime balance rows that could not be mapped cleanly.",
            "runtime_missing_pair_summary_xlsx": "Grouped summary of missing mapping pairs seen at runtime.",
            "override_report_csv": "Report of which manual overrides were applied.",
            "mapping_fix_candidates_csv": "Grouped suggestions for mapping rows that need new config entries.",
            "missing_mapping_candidates_xlsx": "Workbook of candidate mapping additions based on runtime misses.",
            "ninth_mapping_data_coverage_xlsx": "Coverage check for 9th mapping pairs against available 9th data.",
            "balance_coverage_csv": "Coverage summary from LEAP balance extraction.",
            "balance_unit_diagnostics_csv": "Unit normalization checks from LEAP balance extraction.",
            "balance_matching_diagnostics_csv": "Row-level detail-mode and allocation diagnostics from LEAP balance extraction.",
            "balance_extraction_summary_json": "Summary metadata from the balance extraction stage.",
            "chart_line_mapping_ledger": "Per-chart-line ledger linking visible chart rows to mapping decisions.",
            "chart_total_component_ledger": "Ledger showing how visible total lines were constructed.",
            "workflow_stage_timings_csv": "Runtime duration by broad workflow stage.",
        },
        notes=[
            "Primary balance comparison files stay at the top level.",
            "Runtime evidence, ledgers, coverage, and mapping checks are grouped under supporting_files/.",
        ],
    )
    timer.lap("write manifest")

    result = {
        **core_paths,
        "gap_diagnostics": diagnostics_paths.get("gap_diagnostics"),
        "dashboard_comparator_pair_coverage_xlsx": comparator_pair_coverage_xlsx,
        "mapping_rundown_by_sheet": diagnostics_paths.get("mapping_rundown_by_sheet"),
        "mapping_rundown_details": diagnostics_paths.get("mapping_rundown_details"),
        "comparison_issue_summary": diagnostics_paths.get("comparison_issue_summary"),
        "comparison_issue_cause_summary": diagnostics_paths.get("comparison_issue_cause_summary"),
        "simple_leap_balance_mapped": str(simple_leap_balance_path),
        "simple_leap_ninth_balance_mapped": str(simple_leap_ninth_balance_path),
        "simple_ninth_balance_mapped": str(simple_ninth_balance_path),
        "chart_line_mapping_ledger": str(chart_line_mapping_path),
        "chart_total_component_ledger": str(chart_total_component_path),
        "dashboard_index": dashboard_paths.get("dashboard_index"),
        "charts_written": dashboard_paths.get("charts_written"),
        "empty_pages_csv": dashboard_paths.get("empty_pages_csv"),
        "chart_group_exposure": dashboard_paths.get("chart_group_exposure"),
        "all_chart_groups": dashboard_paths.get("all_chart_groups"),
        "chart_navigation_hierarchy": dashboard_paths.get("chart_navigation_hierarchy"),
        "chart_navigation_hierarchy_flat": dashboard_paths.get("chart_navigation_hierarchy_flat"),
        "graph_fuel_coverage_csv": dashboard_paths.get("graph_fuel_coverage_csv"),
        "runtime_issues_csv": issues_paths.get("runtime_issues_csv"),
        "issues_summary_json": issues_paths.get("issues_summary_json"),
        "override_report_csv": issues_paths.get("override_report_csv"),
        "mapping_fix_candidates_csv": mapping_fix_candidates_csv,
        "missing_mapping_candidates_xlsx": missing_mapping_candidates_xlsx,
        "ninth_mapping_data_coverage_xlsx": ninth_mapping_data_coverage_xlsx,
        "balance_coverage_csv": str(coverage_path),
        "balance_unit_diagnostics_csv": str(unit_diag_path),
        "balance_matching_diagnostics_csv": str(matching_diag_path),
        "balance_extraction_summary_json": str(extraction_summary_path),
        "diagnostics": checks,
        "output_manifest": str(manifest),
        "workflow_stage_timings_csv": str(timing_path),
    }
    if RUN_ESTO_AXIS_WORKFLOW_AFTER_BALANCE:
        result.update(_run_esto_axis_workflow_after_balance())
        timer.lap("run ESTO-axis workflow after balance")
    result.update(
        _capture_unmapped_balance_row_error(
            ingestion["issues"],
            issues_paths.get("runtime_issues_csv"),
        )
    )
    timer.finish()
    if WRITE_WORKFLOW_TIMING_CSV:
        timer.write_csv(timing_path)
    return result


#%%
# Notebook run cell.
RUN_WORKFLOW = True
WORKFLOW_RESULT: dict[str, object] | None = None
if RUN_WORKFLOW:
    WORKFLOW_RESULT = run_workflow()
    print("[OK] Balance dashboard workflow complete.")
    for key, value in WORKFLOW_RESULT.items():
        print(f"- {key}: {value}")
        
#%%
