#%%
"""
Convert LEAP and 9th balance data into ESTO-style long balance tables.

This workflow writes dashboard-independent CSVs keyed by scenario, year, ESTO
flow, and ESTO product. Supporting diagnostics are written under
supporting_files.
"""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path
from typing import Any, Sequence

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.utilities.leap_results_dashboard_balance import (  # noqa: E402
    build_balance_comparison_esto_axis,
    build_esto_axis_structure_from_dashboard_template,
    build_ninth_balance_esto_long_table,
    build_simple_ninth_balance_table,
    convert_leap_balances_to_esto_long_table,
)
from codebase.utilities.leap_balance_export_resolver import resolve_balance_export_workbook  # noqa: E402
from codebase.utilities.workflow_common import archive_config_dir_once_per_day  # noqa: E402
from codebase.utilities.workflow_outputs import build_workflow_output_layout, write_output_manifest  # noqa: E402
from codebase.utilities.output_paths import BALANCE_TABLES_ROOT  # noqa: E402


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


def _load_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


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

KNOWN_ISSUES_CONFIG_PATH = _resolve("config/leap_results_balance_known_issues.json")
CHART_NAVIGATION_GUIDE_PATH = _resolve("config/leap_comparison_dashboard_template.json")

LEAP_TO_ESTO_MAPPING = (_resolve("config/leap_mappings.xlsx"), "leap_combined_esto")
NINTH_TO_ESTO_MAPPING = (_resolve("config/ninth_pairs_to_esto_pairs.xlsx"), "ninth_pairs_to_esto_pairs")
CODEBOOK_PATH = _resolve("config/sector_fuel_codes_to_names.xlsx")
SHEET_MAP_PATH = _resolve("config/leap_results_sheet_map.csv")
BACKUP_MAPPINGS_PATH = _resolve("config/backup_leap_mappings.xlsx")
EXPLICIT_MAPPINGS_PATH = _resolve("config/leap_results_explicit_mappings.csv")
EXPLICIT_REASSIGNMENTS_PATH = _resolve("config/leap_results_explicit_reassignments.csv")
SYNTHETIC_REFERENCE_ROWS_PATH = _resolve("config/synthetic_reference_rows.csv")

BASE_TABLE_PATH = _resolve("data/00APEC_2025_low_with_subtotals.csv")
PROJECTION_TABLE_PATH = _resolve("data/merged_file_energy_ALL_20251106.csv")

OUTPUT_DIR = BALANCE_TABLES_ROOT / "leap_balance_to_esto_long" / "USA"
BASE_YEAR = 2022
PROJECTION_ECONOMY = "20_USA"
BASE_ECONOMY = "20USA"
MAX_OUTPUT_YEAR = 2060
PROJECTION_YEARS: Sequence[int] = tuple(range(BASE_YEAR + 1, MAX_OUTPUT_YEAR + 1))
SCENARIO_MAP = {"Reference": "reference", "Target": "target"}


#%%
def _mapping_workbook(mapping_ref: tuple[Path, str]) -> Path:
    return mapping_ref[0]


def run_workflow() -> dict[str, object]:
    archive_config_dir_once_per_day()
    out_dir = _resolve(OUTPUT_DIR)
    layout = build_workflow_output_layout(out_dir)

    structure_config = build_esto_axis_structure_from_dashboard_template(CHART_NAVIGATION_GUIDE_PATH)
    known_issues = _load_json(KNOWN_ISSUES_CONFIG_PATH)

    conversion = convert_leap_balances_to_esto_long_table(
        ref_workbook_path=REF_WORKBOOK_PATH,
        tgt_workbook_path=TGT_WORKBOOK_PATH,
        template_sheet="EBal|2060",
        mapping_pairs_path=_mapping_workbook(LEAP_TO_ESTO_MAPPING),
        codebook_path=CODEBOOK_PATH,
        structure_config=structure_config,
        known_issues=known_issues,
        projection_economy=PROJECTION_ECONOMY,
        max_output_year=MAX_OUTPUT_YEAR,
        explicit_pair_mappings_only=True,
    )

    comparison = build_balance_comparison_esto_axis(
        leap_long=conversion["leap_long"],
        mapping_status=conversion["mapping_status"],
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
        chart_navigation_guide_path=CHART_NAVIGATION_GUIDE_PATH,
        known_issues=known_issues,
    )
    comparison_long = comparison["comparison_long"].copy()
    comparison_long = comparison_long[comparison_long["year"].le(MAX_OUTPUT_YEAR)].copy()
    mapping_status = comparison["mapping_status"].copy()
    ninth_balance = build_simple_ninth_balance_table(
        comparison_long=comparison_long,
        mapping_status=mapping_status,
    )
    ninth_esto_long = build_ninth_balance_esto_long_table(ninth_balance)

    esto_long_path = layout.root / "leap_balance_esto_long.csv"
    ninth_esto_long_path = layout.root / "ninth_balance_esto_long.csv"
    leap_long_path = layout.analysis / "leap_balance_mapped_detail_long.csv"
    mapping_status_path = layout.mapping / "leap_balance_mapping_status.csv"
    comparison_long_path = layout.analysis / "esto_axis_comparison_long.csv"
    ninth_balance_path = layout.analysis / "ninth_balance_esto_long_semantic_columns.csv"
    runtime_issues_path = layout.runtime / "leap_balance_runtime_issues.csv"
    override_report_path = layout.runtime / "leap_balance_override_application_report.csv"
    auto_sheet_path = layout.runtime / "auto_sheet_rows.csv"
    coverage_path = layout.checks / "leap_balance_coverage.csv"
    unit_diag_path = layout.checks / "leap_balance_unit_diagnostics.csv"
    extraction_summary_path = layout.runtime / "leap_balance_extraction_summary.json"
    resolved_structure_path = layout.runtime / "resolved_structure_config.json"

    conversion["esto_long"].to_csv(esto_long_path, index=False)
    ninth_esto_long.to_csv(ninth_esto_long_path, index=False)
    conversion["leap_long"].to_csv(leap_long_path, index=False)
    mapping_status.to_csv(mapping_status_path, index=False)
    comparison_long.to_csv(comparison_long_path, index=False)
    ninth_balance.to_csv(ninth_balance_path, index=False)
    conversion["issues"].to_csv(runtime_issues_path, index=False)
    conversion["override_report"].to_csv(override_report_path, index=False)
    conversion["auto_sheet_rows"].to_csv(auto_sheet_path, index=False)
    conversion["coverage"].to_csv(coverage_path, index=False)
    conversion["unit_diagnostics"].to_csv(unit_diag_path, index=False)
    extraction_summary_path.write_text(
        json.dumps(conversion["extraction_summary"], ensure_ascii=True, indent=2),
        encoding="utf-8",
    )
    resolved_structure_path.write_text(
        json.dumps(conversion["resolved_structure"], ensure_ascii=True, indent=2),
        encoding="utf-8",
    )

    primary_outputs = {
        "leap_balance_esto_long": str(esto_long_path),
        "ninth_balance_esto_long": str(ninth_esto_long_path),
    }
    supporting_outputs = {
        "leap_balance_mapped_detail_long": str(leap_long_path),
        "leap_balance_mapping_status": str(mapping_status_path),
        "esto_axis_comparison_long": str(comparison_long_path),
        "ninth_balance_esto_long_semantic_columns": str(ninth_balance_path),
        "runtime_issues_csv": str(runtime_issues_path),
        "override_report_csv": str(override_report_path),
        "auto_sheet_rows_csv": str(auto_sheet_path),
        "balance_coverage_csv": str(coverage_path),
        "balance_unit_diagnostics_csv": str(unit_diag_path),
        "balance_extraction_summary_json": str(extraction_summary_path),
        "resolved_structure_json": str(resolved_structure_path),
    }
    manifest_path = write_output_manifest(
        out_dir=layout.root,
        primary_outputs=primary_outputs,
        supporting_outputs=supporting_outputs,
        primary_output_descriptions={
            "leap_balance_esto_long": "Primary LEAP balance table aligned to ESTO flow/product columns.",
            "ninth_balance_esto_long": "Primary 9th projection table aligned to the same ESTO flow/product columns.",
        },
        supporting_output_descriptions={
            "leap_balance_mapped_detail_long": "Detailed LEAP long table before collapsing to the primary ESTO-style output.",
            "leap_balance_mapping_status": "Row-level mapping status for LEAP balance sector and product mappings.",
            "esto_axis_comparison_long": "Combined LEAP and 9th ESTO-axis comparison rows used to build derived balance tables.",
            "ninth_balance_esto_long_semantic_columns": "9th ESTO-axis rows in semantic balance-column form.",
            "runtime_issues_csv": "Unmapped or problematic LEAP balance rows recorded during extraction.",
            "override_report_csv": "Applied override rows from the balance structure and known-issues logic.",
            "auto_sheet_rows_csv": "Rows created by the workflow when structure rules auto-generated balance-sheet assignments.",
            "balance_coverage_csv": "Coverage summary for extracted balance rows and mappings.",
            "balance_unit_diagnostics_csv": "Unit-normalization diagnostics from balance extraction.",
            "balance_extraction_summary_json": "Summary counts from the extraction run.",
            "resolved_structure_json": "Resolved structure config after defaults and overrides were applied.",
        },
        notes=[
            "Primary tables stay at the workflow root.",
            "Supporting diagnostics, runtime artifacts, and detailed intermediates live under supporting_files/.",
        ],
    )
    result = {**primary_outputs, **supporting_outputs, "output_manifest_json": str(manifest_path)}
    return result


#%%
WORKFLOW_RESULT: dict[str, object] | None = None
if __name__ == "__main__":
    WORKFLOW_RESULT = run_workflow()
    print("[OK] LEAP balance to ESTO long workflow complete.")
    for key, value in WORKFLOW_RESULT.items():
        print(f"- {key}: {value}")

#%%
