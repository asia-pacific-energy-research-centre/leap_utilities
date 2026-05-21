#%%
"""
LEAP Results dashboard workflow (utilities repo)

This script builds the original LEAP results comparison dashboard and supporting
mapping artifacts for LEAP, 9th Outlook, and ESTO series.
It is the V1 dashboard path, kept for comparison with the V2 dashboard workflow
and for workflows that still rely on its output layout.

Notebook-first script that:
- reads exported LEAP result workbooks (transport/industry/demand_others)
- maps each sheet to 9th/ESTO sectors using config/leap_results_sheet_map.csv
- aligns fuels via canonical config/ninth_pairs_to_esto_pairs.xlsx
  + config/sector_fuel_codes_to_names.xlsx (optional overrides in config/backup_leap_mappings.xlsx)
- compares LEAP series to base-year ESTO (2022) and 9th projections
- generates charts and dashboards styled like leap_transport
"""
from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Sequence

import pandas as pd
from openpyxl.comments import Comment

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.mappings.canonical_mapping import (
    DEFAULT_BACKUP_LEAP_MAPPINGS,
    DEFAULT_CODEBOOK,
    DEFAULT_NINTH_TO_ESTO,
    DEFAULT_SHEET_MAP,
    build_sector_to_esto_flow_lookup,
    load_canonical_pairs,
    load_fuel_aliases,
    load_sheet_map,
)
from codebase.utilities.leap_results_dashboard_utils import (
    DEFAULT_EXPLICIT_LEAP_MAPPINGS,
    DEFAULT_EXPLICIT_LEAP_REASSIGNMENTS,
    apply_explicit_sector_reassignments,
    basic_checks,
    build_charts,
    build_comparisons,
    build_dashboards,
    ensure_repo_root,
    load_explicit_sector_reassignments,
    load_leap_workbook,
    load_explicit_sector_fuel_mappings,
)
from codebase.scrapbook.utilities import load_augmented_reference_tables
from codebase.utilities.workflow_common import archive_config_dir_once_per_day
from codebase.utilities.workflow_outputs import build_workflow_output_layout, write_output_manifest


# -----------------------------------------------------------------------------
# Notebook-editable toggles (keep simple and explicit)
# -----------------------------------------------------------------------------
LEAP_RESULTS_DIR = REPO_ROOT / "outputs/leap_results"
# Economy/scenario filters for workbook discovery
ECONOMY_TOKEN = "USA"  # substring to match in filenames
SCENARIOS = ("Reference", "Target")
SHEET_MAP_PATH = DEFAULT_SHEET_MAP
BACKUP_MAPPINGS_PATH = DEFAULT_BACKUP_LEAP_MAPPINGS
EXPLICIT_MAPPINGS_PATH = DEFAULT_EXPLICIT_LEAP_MAPPINGS
EXPLICIT_REASSIGNMENTS_PATH = DEFAULT_EXPLICIT_LEAP_REASSIGNMENTS
CODEBOOK_PATH = DEFAULT_CODEBOOK
NINTH_TO_ESTO_PATH = DEFAULT_NINTH_TO_ESTO
BASE_TABLE_PATH = REPO_ROOT / "data/00APEC_2025_low_with_subtotals.csv"
PROJECTION_TABLE_PATH = REPO_ROOT / "data/merged_file_energy_ALL_20251106.csv"
OUTPUT_DIR = REPO_ROOT / "outputs/dashboards/leap_results_dashboard/USA"
MAPPING_VIEWS_DIR = REPO_ROOT / "config/computer_generated_config/leap_mapping_views/USA"
REFERENCE_CACHE_DIR = REPO_ROOT / "data/.cache/leap_results_dashboard_reference_tables"
BASE_YEAR = 2022  # from 00APEC_2025_low_with_subtotals.csv
PROJECTION_YEARS: Sequence[int] = tuple(range(2023, 2071))
SCENARIO_MAP = {"reference": "reference", "target": "target"}
BASE_ECONOMY = "20USA"
PROJECTION_ECONOMY = "20_USA"
CHART_BACKEND = "plotly"  # "plotly" or "static"
USE_ESTO_AGG_ONLY = False  # include both ESTO base-year and 9th projection on charts
SIBLING_COMPARATOR_MODE = "aggregate_to_parent"  # "none", "allocate_by_leap_share", or "aggregate_to_parent"
INCLUDE_SIBLING_PARENT_TOTALS = True  # add promoted parent-category chart rows (for example, "Road")
GENERATE_CHARTS = False  # keep False while debugging number differences
GENERATE_DASHBOARDS = False  # keep False while debugging number differences
HIDE_LEAP_ONLY_CHARTS = True  # suppress charts when a sheet/fuel has no comparator series
DIAGNOSTIC_PROBE_YEAR = 2030
TOP_DIAGNOSTIC_ROWS = 40
DROP_ALL_ZERO_BASE_ROWS = True
DROP_ALL_ZERO_PROJECTION_ROWS = False


# -----------------------------------------------------------------------------
# Core workflow
# -----------------------------------------------------------------------------
def _resolve(path: Path | str) -> Path:
    """Resolve a path relative to repo root if not absolute."""
    p = Path(path)
    return p if p.is_absolute() else (REPO_ROOT / p)


def _drop_all_zero_year_rows(df: pd.DataFrame, *, dataset_name: str) -> pd.DataFrame:
    """
    Drop rows where every year column is zero (treating NaN as zero).

    This avoids keeping fully empty time-series rows that can block parent-level
    rollups or force misleading zero-valued comparator lines.
    """
    if df.empty:
        return df

    year_cols = [c for c in df.columns if str(c).strip().isdigit() and len(str(c).strip()) == 4]
    if not year_cols:
        print(f"[INFO] {dataset_name}: no year columns found for all-zero filtering.")
        return df

    year_values = df[year_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    non_zero_mask = year_values.ne(0).any(axis=1)
    dropped_rows = int((~non_zero_mask).sum())
    kept = df.loc[non_zero_mask].copy()
    if dropped_rows:
        print(
            f"[INFO] {dataset_name}: dropped {dropped_rows} all-zero rows "
            f"(kept {len(kept)} of {len(df)})."
        )
    else:
        print(f"[INFO] {dataset_name}: no all-zero rows dropped.")
    return kept


_HEADER_NOTES = {
    "mapping_source": {
        "canonical": "Matched directly from config/ninth_pairs_to_esto_pairs.xlsx (main sheet).",
        "codebook_fallback": "Matched from config/sector_fuel_codes_to_names.xlsx using the ESTO_LEAP_names or code_to_name sheet.",
        "explicit": "Matched from config/leap_results_explicit_mappings.csv using an exact sheet/fuel(/sector) override.",
        "override": "Matched from config/backup_leap_mappings.xlsx (manual override).",
    },
    "flow_source": {
        "canonical": "ESTO flow came from config/ninth_pairs_to_esto_pairs.xlsx (main sheet).",
        "explicit": "ESTO flow came from config/leap_results_explicit_mappings.csv using an exact override.",
        "sector_fallback": "ESTO flow came from config/sector_fuel_codes_to_names.xlsx, sheet code_to_name.",
        "sheet_override": "ESTO flow came from config/leap_results_sheet_map.csv, column esto_flow_override.",
        "override": "ESTO flow came from config/backup_leap_mappings.xlsx (manual override).",
    },
    "fuel_source": {
        "canonical": "9th fuel code came from config/ninth_pairs_to_esto_pairs.xlsx (main sheet).",
        "explicit": "9th fuel code came from config/leap_results_explicit_mappings.csv using an exact override.",
        "inferred": "9th fuel code was inferred by the workflow from the ESTO flow and product after the initial lookup.",
        "override": "9th fuel code came from config/backup_leap_mappings.xlsx (manual override).",
    },
    "mapped": {
        "true": "Row is complete for both ESTO base-year and 9th projection comparisons: esto_flow, esto_product, and ninth_fuel_code are all present.",
        "false": "Row is incomplete for at least one comparison requirement; see the completeness flags.",
    },
    "has_any_mapping": {
        "true": "At least one of ninth_fuel_code, esto_flow, or esto_product was filled.",
        "false": "All of ninth_fuel_code, esto_flow, and esto_product are blank.",
    },
    "base_mapping_complete": {
        "true": "Both esto_flow and esto_product are present, so ESTO base-year comparison is well-specified.",
        "false": "At least one of esto_flow or esto_product is missing, so ESTO base-year comparison is incomplete.",
    },
    "projection_mapping_complete": {
        "true": "ninth_fuel_code is present, so the workflow can pull a fuel-specific 9th projection series.",
        "false": "ninth_fuel_code is missing, so projection comparison is incomplete.",
    },
    "partially_mapped": {
        "true": "Some mapping fields are present, but the row is not complete for both ESTO base-year and 9th projection comparisons.",
        "false": "Either the row is fully mapped or nothing was mapped.",
    },
    "missing_ninth_fuel": {
        "true": "ninth_fuel_code is blank.",
        "false": "ninth_fuel_code is present.",
    },
    "missing_esto_flow": {
        "true": "esto_flow is blank.",
        "false": "esto_flow is present.",
    },
    "missing_esto_product": {
        "true": "esto_product is blank.",
        "false": "esto_product is present.",
    },
    "has_mapping_note": {
        "true": "mapping_note contains an extra note for this row.",
        "false": "mapping_note is blank.",
    },
}


def _build_header_note(df: pd.DataFrame, column_name: str) -> str:
    """Build an Excel header comment for a known coded/flag column."""
    value_notes = _HEADER_NOTES.get(column_name)
    if not value_notes or column_name not in df.columns:
        return ""

    values = (
        df[column_name]
        .dropna()
        .astype(str)
        .str.strip()
        .str.lower()
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    )
    if not values:
        return ""

    lines = [f"{column_name} values in this file:"]
    for value in sorted(values):
        lines.append(f"{value}: {value_notes.get(value, 'Used by the workflow; see mapping logic for details.')}")
    return "\n".join(lines)


def _write_workbook_with_header_comments(df: pd.DataFrame, path: Path, *, sheet_name: str) -> None:
    """Write a DataFrame to Excel and add header comments for known coded columns."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.book[sheet_name]
        for col_idx, column_name in enumerate(df.columns, start=1):
            note = _build_header_note(df, column_name)
            if note:
                worksheet.cell(row=1, column=col_idx).comment = Comment(note, "Codex")


def _discover_workbooks(root: Path, economy_token: str, scenarios: Sequence[str]) -> list[Path]:
    """
    Find LEAP result workbooks in root matching economy token and scenario labels.
    """
    root = _resolve(root)
    if not root.exists():
        raise FileNotFoundError(f"LEAP results directory not found: {root}")

    economy_token = economy_token.lower()
    scen_tokens = [s.lower() for s in scenarios]
    candidates = sorted(root.glob("*.xls*"))

    matched: list[Path] = []
    for path in candidates:
        name = path.name.lower()
        if economy_token not in name:
            continue
        if any(s in name for s in scen_tokens):
            matched.append(path)
    if not matched:
        raise FileNotFoundError(
            f"No LEAP workbooks found in {root} matching economy '{economy_token}' and scenarios {scenarios}."
        )
    return matched


def _write_issue_summary(
    *,
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    diagnostics_dir: Path,
    base_year: int,
    probe_year: int,
    top_n: int,
) -> str | None:
    """
    Build a compact sortable audit table for the largest LEAP/comparator gaps.
    This is meant for quick iteration when charts are disabled.
    """
    if comparison_long.empty:
        return None

    comp = comparison_long.copy()
    comp["value"] = pd.to_numeric(comp["value"], errors="coerce")
    comp["year"] = pd.to_numeric(comp["year"], errors="coerce").astype("Int64")
    comp["source_compare"] = comp["source"].replace(
        {
            "base_estimated": "base",
            "base_mixed": "base",
            "projection_estimated": "projection",
            "projection_mixed": "projection",
        }
    )
    comp = comp[comp["source_compare"].isin(["leap", "base", "projection"])].copy()
    if comp.empty:
        return None

    wide = (
        comp.pivot_table(
            index=["sheet", "fuel_label", "scenario", "year"],
            columns="source_compare",
            values="value",
            aggfunc="first",
        )
        .reset_index()
    )
    for col in ["leap", "base", "projection"]:
        if col not in wide.columns:
            wide[col] = pd.NA

    base_rows = wide[wide["year"] == base_year].copy()
    if not base_rows.empty:
        base_rows["base_abs_gap"] = (
            pd.to_numeric(base_rows["leap"], errors="coerce") - pd.to_numeric(base_rows["base"], errors="coerce")
        ).abs()
        base_rows["base_gap_ratio"] = base_rows["base_abs_gap"] / pd.to_numeric(
            base_rows["leap"], errors="coerce"
        ).abs().where(lambda s: s > 1e-9)
    else:
        base_rows = pd.DataFrame(columns=["sheet", "fuel_label", "scenario", "base_abs_gap", "base_gap_ratio"])

    probe_rows = wide[wide["year"] == probe_year].copy()
    if not probe_rows.empty:
        probe_rows["projection_abs_gap"] = (
            pd.to_numeric(probe_rows["leap"], errors="coerce") - pd.to_numeric(probe_rows["projection"], errors="coerce")
        ).abs()
        probe_rows["projection_gap_ratio"] = probe_rows["projection_abs_gap"] / pd.to_numeric(
            probe_rows["leap"], errors="coerce"
        ).abs().where(lambda s: s > 1e-9)
    else:
        probe_rows = pd.DataFrame(
            columns=["sheet", "fuel_label", "scenario", "projection_abs_gap", "projection_gap_ratio"]
        )

    summary = base_rows[
        ["sheet", "fuel_label", "scenario", "base_abs_gap", "base_gap_ratio"]
    ].merge(
        probe_rows[
            ["sheet", "fuel_label", "scenario", "projection_abs_gap", "projection_gap_ratio"]
        ],
        on=["sheet", "fuel_label", "scenario"],
        how="outer",
    )

    status = mapping_status.copy()
    if not status.empty:
        keep_cols = [
            "sheet",
            "fuel_label",
            "ninth_fuel_code",
            "mapping_source",
            "flow_source",
            "fuel_source",
            "sector_match_method",
            "mapping_note",
        ]
        for col in ["comparator_scope", "uses_parent_flow", "allow_parent_estimate"]:
            if col in status.columns:
                keep_cols.append(col)
        status = status[keep_cols].drop_duplicates()
        summary = summary.merge(status, on=["sheet", "fuel_label"], how="left")

    def _as_bool(value: object) -> bool:
        if pd.isna(value):
            return False
        if isinstance(value, bool):
            return value
        text = str(value).strip().lower()
        return text in {"true", "1", "yes"}

    def _classify_issue(row: pd.Series) -> str:
        mapping_note = str(row.get("mapping_note") or "").strip().lower()
        mapping_source = str(row.get("mapping_source") or "").strip().lower()
        comparator_scope = str(row.get("comparator_scope") or "").strip().lower()
        fuel_label = str(row.get("fuel_label") or "").strip().lower()
        uses_parent_flow = _as_bool(row.get("uses_parent_flow"))
        allow_parent_estimate = _as_bool(row.get("allow_parent_estimate"))

        if comparator_scope == "parent" or (uses_parent_flow and allow_parent_estimate):
            return "parent_comparator_only"
        if "ambiguous canonical matches" in mapping_note:
            return "ambiguous_canonical_mapping"
        if mapping_source in {"canonical_aggregated", "category_sector"}:
            return "aggregated_mapping"
        if "_x_" in str(row.get("ninth_fuel_code") or ""):
            return "aggregate_fuel_mapping"
        if fuel_label in {"others", "other sources"}:
            return "catch_all_fuel_bucket"
        return "direct_mapping_numeric_gap"

    def _agent_hint(row: pd.Series) -> str:
        cause = str(row.get("issue_cause") or "").strip()
        if cause == "parent_comparator_only":
            return "Check whether this child should compare only at the promoted parent sheet."
        if cause == "ambiguous_canonical_mapping":
            return "Review canonical pairs and explicit mappings; the same sector+fuel resolves to multiple ESTO targets."
        if cause == "aggregated_mapping":
            return "This row aggregates multiple targets; verify the aggregation is intended and not masking a missing child mapping."
        if cause == "aggregate_fuel_mapping":
            return "The 9th fuel is an aggregate bucket; compare against sibling fuels and the shared aggregate fuel audit."
        if cause == "catch_all_fuel_bucket":
            return "Catch-all fuel labels often hide reassignment problems; compare against specific fuels in the same sheet."
        return "Mapping looks direct; inspect source data values, signs, and scenario-specific series."

    for col in ["base_abs_gap", "base_gap_ratio", "projection_abs_gap", "projection_gap_ratio"]:
        summary[col] = pd.to_numeric(summary[col], errors="coerce")
    summary["worst_gap_ratio"] = summary[["base_gap_ratio", "projection_gap_ratio"]].max(axis=1, skipna=True)
    summary["worst_abs_gap"] = summary[["base_abs_gap", "projection_abs_gap"]].max(axis=1, skipna=True)
    if "ninth_fuel_code" not in summary.columns:
        summary["ninth_fuel_code"] = pd.NA
    summary["issue_cause"] = summary.apply(_classify_issue, axis=1)
    summary["agent_debug_hint"] = summary.apply(_agent_hint, axis=1)

    ordered = summary.sort_values(
        ["worst_gap_ratio", "worst_abs_gap", "sheet", "fuel_label", "scenario"],
        ascending=[False, False, True, True, True],
        na_position="last",
    ).reset_index(drop=True)

    diagnostics_dir.mkdir(parents=True, exist_ok=True)
    out_path = diagnostics_dir / "comparison_issue_summary.csv"
    ordered.to_csv(out_path, index=False)

    preview = ordered.head(max(0, int(top_n)))
    if not preview.empty:
        print(
            "[INFO] Top comparison issues "
            f"(base year {base_year}, projection probe {probe_year})"
        )
        display_cols = [
            "sheet",
            "fuel_label",
            "scenario",
            "worst_gap_ratio",
            "worst_abs_gap",
            "issue_cause",
            "mapping_source",
            "comparator_scope",
            "uses_parent_flow",
            "allow_parent_estimate",
        ]
        display_cols = [col for col in display_cols if col in preview.columns]
        print(preview[display_cols].to_string(index=False))

    cause_summary = (
        ordered.groupby("issue_cause", dropna=False)
        .agg(
            rows=("issue_cause", "size"),
            max_gap_ratio=("worst_gap_ratio", "max"),
            max_abs_gap=("worst_abs_gap", "max"),
        )
        .reset_index()
        .sort_values(["rows", "max_gap_ratio", "max_abs_gap"], ascending=[False, False, False], na_position="last")
    )
    if not cause_summary.empty:
        cause_path = diagnostics_dir / "comparison_issue_cause_summary.csv"
        cause_summary.to_csv(cause_path, index=False)
        print("[INFO] Issue causes by frequency")
        print(cause_summary.to_string(index=False))

    return str(out_path)


def _write_gap_and_mapping_diagnostics(
    *,
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    diagnostics_dir: Path,
    mapping_dir: Path,
    base_year: int,
    projection_probe_year: int = 2030,
) -> dict[str, str]:
    """
    Write compact diagnostics files for large comparison gaps and mapping completeness.
    Returns a dict of artifact paths.
    """
    artifacts: dict[str, str] = {}
    diagnostics_dir.mkdir(parents=True, exist_ok=True)
    mapping_dir.mkdir(parents=True, exist_ok=True)

    # ---- Gap diagnostics (LEAP vs base/projection) ----
    gap_path = diagnostics_dir / "comparison_gap_diagnostics.csv"
    if not comparison_long.empty:
        comp = comparison_long.copy()
        comp["value"] = pd.to_numeric(comp["value"], errors="coerce")
        wide = (
            comp.pivot_table(
                index=["sheet", "fuel_label", "scenario", "year"],
                columns="source",
                values="value",
                aggfunc="first",
            )
            .reset_index()
        )
        for col in ["leap", "base", "projection", "esto_aggregated"]:
            if col not in wide.columns:
                wide[col] = pd.NA

        # Base-year diagnostics
        base_rows = wide[wide["year"] == base_year].copy()
        base_rows["gap_base"] = pd.to_numeric(base_rows["leap"], errors="coerce") - pd.to_numeric(base_rows["base"], errors="coerce")
        base_rows["abs_gap_base"] = base_rows["gap_base"].abs()
        base_rows["ratio_base_to_leap"] = pd.to_numeric(base_rows["base"], errors="coerce") / pd.to_numeric(base_rows["leap"], errors="coerce")

        # Probe-year diagnostics (default 2030)
        proj_rows = wide[wide["year"] == projection_probe_year].copy()
        proj_rows["gap_projection"] = pd.to_numeric(proj_rows["leap"], errors="coerce") - pd.to_numeric(proj_rows["projection"], errors="coerce")
        proj_rows["abs_gap_projection"] = proj_rows["gap_projection"].abs()
        proj_rows["ratio_projection_to_leap"] = pd.to_numeric(proj_rows["projection"], errors="coerce") / pd.to_numeric(proj_rows["leap"], errors="coerce")

        key_cols = ["sheet", "fuel_label", "scenario"]
        merged = base_rows[key_cols + ["leap", "base", "gap_base", "abs_gap_base", "ratio_base_to_leap"]].rename(
            columns={"leap": f"leap_{base_year}", "base": f"base_{base_year}"}
        ).merge(
            proj_rows[key_cols + ["leap", "projection", "gap_projection", "abs_gap_projection", "ratio_projection_to_leap"]].rename(
                columns={"leap": f"leap_{projection_probe_year}", "projection": f"projection_{projection_probe_year}"}
            ),
            on=key_cols,
            how="outer",
        )

        # Coverage diagnostics across projection horizon.
        proj_horizon = wide[wide["year"] > base_year].copy()
        proj_cov = (
            proj_horizon.assign(
                leap_present=pd.to_numeric(proj_horizon["leap"], errors="coerce").notna(),
                projection_present=pd.to_numeric(proj_horizon["projection"], errors="coerce").notna(),
            )
            .groupby(key_cols, as_index=False)[["leap_present", "projection_present"]]
            .sum()
            .rename(columns={"leap_present": "leap_year_points", "projection_present": "projection_year_points"})
        )
        merged = merged.merge(proj_cov, on=key_cols, how="left")
        merged["projection_missing_year_points"] = (
            merged["leap_year_points"].fillna(0) - merged["projection_year_points"].fillna(0)
        ).clip(lower=0)

        # Helpful classification tags for quick triage.
        merged["diagnostic_flag"] = ""
        merged.loc[merged["projection_missing_year_points"] > 0, "diagnostic_flag"] = "missing_projection_points"
        merged.loc[
            (merged["diagnostic_flag"] == "") & (pd.to_numeric(merged["gap_projection"], errors="coerce").fillna(0) < 0),
            "diagnostic_flag",
        ] = "projection_above_leap"
        merged.loc[
            (merged["diagnostic_flag"] == "")
            & (pd.to_numeric(merged["leap_2030"], errors="coerce") * pd.to_numeric(merged["projection_2030"], errors="coerce") < 0),
            "diagnostic_flag",
        ] = "sign_mismatch_projection"
        merged.loc[
            (merged["diagnostic_flag"] == "") & (pd.to_numeric(merged["abs_gap_base"], errors="coerce") > 0),
            "diagnostic_flag",
        ] = "base_gap_present"

        merged.sort_values(
            by=["projection_missing_year_points", "abs_gap_projection", "abs_gap_base"],
            ascending=[False, False, False],
            inplace=True,
        )
        merged.to_csv(gap_path, index=False)
        artifacts["gap_diagnostics"] = str(gap_path)

    # ---- Mapping rundown ----
    if not mapping_status.empty:
        status = mapping_status.copy()
        for col in [
            "ninth_fuel_code",
            "esto_flow",
            "esto_product",
            "mapping_note",
            "has_any_mapping",
            "base_mapping_complete",
            "projection_mapping_complete",
            "partially_mapped",
            "mapped",
        ]:
            if col not in status.columns:
                status[col] = ""
            if col in {"has_any_mapping", "base_mapping_complete", "projection_mapping_complete", "partially_mapped", "mapped"}:
                status[col] = status[col].fillna(False).astype(bool)
            else:
                status[col] = status[col].fillna("").astype(str).str.strip()

        status["missing_ninth_fuel"] = status["ninth_fuel_code"] == ""
        status["missing_esto_flow"] = status["esto_flow"] == ""
        status["missing_esto_product"] = status["esto_product"] == ""
        status["has_mapping_note"] = status["mapping_note"] != ""

        detail_path = mapping_dir / "mapping_rundown_details.xlsx"
        legacy_detail_csv_path = mapping_dir / "mapping_rundown_details.csv"
        try:
            _write_workbook_with_header_comments(status, detail_path, sheet_name="mapping_rundown_details")
        except PermissionError:
            print(
                "[WARN] Could not write mapping_rundown_details workbook because it is in use. "
                f"Close it and rerun if you need it refreshed: {detail_path}"
            )
        if legacy_detail_csv_path.exists():
            try:
                legacy_detail_csv_path.unlink()
                print(f"[INFO] Removed legacy mapping_rundown_details CSV: {legacy_detail_csv_path}")
            except PermissionError:
                print(
                    "[WARN] Could not remove legacy mapping_rundown_details CSV because it is in use. "
                    f"Close it and remove manually: {legacy_detail_csv_path}"
                )
        artifacts["mapping_rundown_details"] = str(detail_path)

        by_sheet = (
            status.groupby("sheet", as_index=False)
            .agg(
                rows=("sheet", "size"),
                fully_mapped=("mapped", "sum"),
                partially_mapped=("partially_mapped", "sum"),
                any_mapping=("has_any_mapping", "sum"),
                base_mapping_complete=("base_mapping_complete", "sum"),
                projection_mapping_complete=("projection_mapping_complete", "sum"),
                missing_ninth_fuel=("missing_ninth_fuel", "sum"),
                missing_esto_flow=("missing_esto_flow", "sum"),
                missing_esto_product=("missing_esto_product", "sum"),
                with_mapping_notes=("has_mapping_note", "sum"),
            )
            .sort_values(
                by=["fully_mapped", "partially_mapped", "missing_ninth_fuel", "missing_esto_flow", "missing_esto_product", "rows"],
                ascending=[True, False, False, False, False, False],
            )
        )
        by_sheet_path = mapping_dir / "mapping_rundown_by_sheet.csv"
        by_sheet.to_csv(by_sheet_path, index=False)
        artifacts["mapping_rundown_by_sheet"] = str(by_sheet_path)

        aggregate_fuel_audit = status.copy()
        aggregate_fuel_audit["ninth_fuel_code"] = aggregate_fuel_audit["ninth_fuel_code"].astype(str).str.strip()
        aggregate_fuel_audit = aggregate_fuel_audit[
            aggregate_fuel_audit["ninth_fuel_code"].str.contains("_x_", regex=False)
        ].copy()
        if not aggregate_fuel_audit.empty:
            aggregate_fuel_audit = (
                aggregate_fuel_audit.groupby(["sheet", "ninth_fuel_code"], as_index=False)
                .agg(
                    label_count=("fuel_label", "nunique"),
                    fuel_labels=(
                        "fuel_label",
                        lambda s: " | ".join(sorted({str(v).strip() for v in s if str(v).strip()})),
                    ),
                )
                .sort_values(["label_count", "sheet", "ninth_fuel_code"], ascending=[False, True, True])
            )
            aggregate_fuel_audit_path = mapping_dir / "shared_aggregate_fuel_mapping_audit.csv"
            aggregate_fuel_audit.to_csv(aggregate_fuel_audit_path, index=False)
            artifacts["shared_aggregate_fuel_mapping_audit"] = str(aggregate_fuel_audit_path)

        ambiguous_fallback_audit = status.copy()
        ambiguous_fallback_audit["mapping_note"] = ambiguous_fallback_audit["mapping_note"].astype(str)
        ambiguous_fallback_audit = ambiguous_fallback_audit[
            ambiguous_fallback_audit["mapping_note"].str.contains(
                "ambiguous canonical matches for sector+fuel",
                case=False,
                regex=False,
            )
        ].copy()
        if not ambiguous_fallback_audit.empty:
            ambiguous_fallback_audit = ambiguous_fallback_audit[
                [
                    "sheet",
                    "fuel_label",
                    "ninth_fuel_code",
                    "esto_flow",
                    "esto_product",
                    "mapping_source",
                    "flow_source",
                    "fuel_source",
                    "sector_match_method",
                    "mapping_note",
                ]
            ].sort_values(["sheet", "fuel_label"])
            ambiguous_fallback_audit_path = mapping_dir / "ambiguous_canonical_fallback_audit.csv"
            ambiguous_fallback_audit.to_csv(ambiguous_fallback_audit_path, index=False)
            artifacts["ambiguous_canonical_fallback_audit"] = str(ambiguous_fallback_audit_path)

    return artifacts


def run_workflow() -> dict[str, object]:
    ensure_repo_root()
    out_dir = _resolve(OUTPUT_DIR)
    layout = build_workflow_output_layout(out_dir)
    print(f"[INFO] Output dir: {out_dir}")

    leap_workbooks = _discover_workbooks(LEAP_RESULTS_DIR, ECONOMY_TOKEN, SCENARIOS)
    print(f"[INFO] Using {len(leap_workbooks)} LEAP workbook(s):")
    for wb in leap_workbooks:
        print(f"  - {wb}")

    sheet_map = load_sheet_map(_resolve(SHEET_MAP_PATH))
    fuel_aliases = load_fuel_aliases(
        _resolve(BACKUP_MAPPINGS_PATH) if BACKUP_MAPPINGS_PATH else None,
        _resolve(CODEBOOK_PATH),
    )
    explicit_mappings = load_explicit_sector_fuel_mappings(_resolve(EXPLICIT_MAPPINGS_PATH))
    explicit_reassignments = load_explicit_sector_reassignments(_resolve(EXPLICIT_REASSIGNMENTS_PATH))
    sector_flow_mapping = build_sector_to_esto_flow_lookup(_resolve(CODEBOOK_PATH))
    ninth_pairs, canonical_conflicts = load_canonical_pairs(_resolve(NINTH_TO_ESTO_PATH), strict=False)
    if not canonical_conflicts.empty:
        mapping_views_dir = _resolve(MAPPING_VIEWS_DIR)
        mapping_views_dir.mkdir(parents=True, exist_ok=True)
        conflict_path = mapping_views_dir / "mapping_conflicts.csv"
        canonical_conflicts.to_csv(conflict_path, index=False)
    archive_config_dir_once_per_day()
    base_df, ninth_df = load_augmented_reference_tables(
        esto_path=_resolve(BASE_TABLE_PATH),
        ninth_path=_resolve(PROJECTION_TABLE_PATH),
        explicit_reassignments_path=_resolve(EXPLICIT_REASSIGNMENTS_PATH),
        explicit_mappings_path=_resolve(EXPLICIT_MAPPINGS_PATH),
        canonical_pairs_path=_resolve(NINTH_TO_ESTO_PATH),
        synthetic_rules_path=REPO_ROOT / "config/synthetic_reference_rows.csv",
        cache_dir=REFERENCE_CACHE_DIR,
        apply_esto_subtotal_map=False,
        filter_esto_subtotals_flag=False,
        filter_ninth_subtotals_flag=False,
    )
    reassignment_status = pd.DataFrame()
    if DROP_ALL_ZERO_BASE_ROWS:
        base_df = _drop_all_zero_year_rows(base_df, dataset_name="base_df (ESTO)")
    if DROP_ALL_ZERO_PROJECTION_ROWS:
        ninth_df = _drop_all_zero_year_rows(ninth_df, dataset_name="ninth_df (9th)")
    if not reassignment_status.empty:
        reassignment_path = layout.mapping / "explicit_reassignment_status.csv"
        reassignment_status.to_csv(reassignment_path, index=False)
        target_summary = (
            reassignment_status.assign(
                matched_rows=pd.to_numeric(reassignment_status["matched_rows"], errors="coerce").fillna(0),
                target_esto_flow=reassignment_status["target_esto_flow"].fillna("").astype(str).str.strip(),
                target_esto_product=reassignment_status["target_esto_product"].fillna("").astype(str).str.strip(),
            )
            .groupby(["target_esto_flow", "target_esto_product"], dropna=False, as_index=False)["matched_rows"]
            .sum()
            .sort_values(["matched_rows", "target_esto_flow", "target_esto_product"], ascending=[False, True, True])
        )
        target_bits = [
            f"{row.target_esto_flow} | {row.target_esto_product}: {int(row.matched_rows)}"
            for row in target_summary.itertuples(index=False)
        ]
        print(
            "[INFO] Applied explicit sector reassignments by target "
            f"({'; '.join(target_bits)}). Details: {reassignment_path}"
        )

    # Load LEAP long data from all workbooks
    leap_frames = [load_leap_workbook(wb, sheet_map=sheet_map) for wb in leap_workbooks]
    leap_long = pd.concat(leap_frames, ignore_index=True) if leap_frames else pd.DataFrame()
    if leap_long.empty:
        raise RuntimeError("No LEAP data loaded; check workbook paths.")

    comparison_long, comparison_wide, mapping_status = build_comparisons(
        leap_long,
        sheet_map=sheet_map,
        fuel_mapping=fuel_aliases,
        sector_flow_mapping=sector_flow_mapping,
        ninth_pairs=ninth_pairs,
        base_df=base_df,
        ninth_df=ninth_df,
        explicit_mappings=explicit_mappings,
        base_year=BASE_YEAR,
        base_economy=BASE_ECONOMY,
        projection_economy=PROJECTION_ECONOMY,
        projection_years=PROJECTION_YEARS,
        scenario_map=SCENARIO_MAP,
        use_esto_agg_only=USE_ESTO_AGG_ONLY,
        sibling_comparator_mode=SIBLING_COMPARATOR_MODE,
        include_sibling_parent_totals=INCLUDE_SIBLING_PARENT_TOTALS,
    )

    # Write outputs
    comparison_long_path = layout.root / "comparison_long.csv"
    comparison_wide_path = layout.root / "comparison_wide.csv"
    mapping_status_path = layout.root / "mapping_status.xlsx"
    legacy_mapping_status_csv_path = layout.root / "mapping_status.csv"
    leap_long_path = layout.root / "leap_long.csv"

    comparison_long.to_csv(comparison_long_path, index=False)
    comparison_wide.to_csv(comparison_wide_path, index=False)
    try:
        _write_workbook_with_header_comments(mapping_status, mapping_status_path, sheet_name="mapping_status")
        print(f"[INFO] Wrote mapping_status: {mapping_status_path}")
    except PermissionError:
        print(
            "[WARN] Could not write mapping_status workbook because it is in use. "
            f"Close it and rerun if you need it refreshed: {mapping_status_path}"
        )
    if legacy_mapping_status_csv_path.exists():
        try:
            legacy_mapping_status_csv_path.unlink()
            print(f"[INFO] Removed legacy mapping_status CSV: {legacy_mapping_status_csv_path}")
        except PermissionError:
            print(
                "[WARN] Could not remove legacy mapping_status CSV because it is in use. "
                f"Close it and remove manually: {legacy_mapping_status_csv_path}"
            )
    leap_long.to_csv(leap_long_path, index=False)
    print(f"[INFO] Wrote comparison_long: {comparison_long_path}")
    diagnostics_artifacts: dict[str, str] = {}
    try:
        diagnostics_artifacts = _write_gap_and_mapping_diagnostics(
            comparison_long=comparison_long,
            mapping_status=mapping_status,
            diagnostics_dir=layout.diagnostics,
            mapping_dir=layout.mapping,
            base_year=BASE_YEAR,
        )
        for key, path in diagnostics_artifacts.items():
            print(f"[INFO] Wrote {key}: {path}")
    except PermissionError as exc:
        print(
            "[WARN] Could not refresh one or more diagnostics files because a file is in use. "
            f"Close it and rerun if needed: {exc}"
        )
    issue_summary_path = _write_issue_summary(
        comparison_long=comparison_long,
        mapping_status=mapping_status,
        diagnostics_dir=layout.diagnostics,
        base_year=BASE_YEAR,
        probe_year=DIAGNOSTIC_PROBE_YEAR,
        top_n=TOP_DIAGNOSTIC_ROWS,
    )
    if issue_summary_path:
        print(f"[INFO] Wrote comparison_issue_summary: {issue_summary_path}")
    issue_cause_summary_path = layout.diagnostics / "comparison_issue_cause_summary.csv"
    if not issue_cause_summary_path.exists():
        issue_cause_summary_path = None

    def _remove_stale_unmapped_file(path: Path) -> None:
        if not path.exists():
            return
        try:
            path.unlink()
            print(f"[INFO] Removed stale unmapped file: {path}")
        except PermissionError:
            print(
                "[WARN] Could not remove stale unmapped file because it is in use. "
                f"Close it and remove manually: {path}"
            )

    # Fail fast on unmapped fuels that carry data; warn on empty unmapped rows.
    unmapped_path = layout.mapping / "unmapped_fuels_with_data.csv"
    unmapped = mapping_status[~mapping_status["mapped"]]
    if not unmapped.empty:
        merged = unmapped.merge(
            leap_long,
            left_on=["sheet", "fuel_label"],
            right_on=["sheet_name", "fuel_label"],
            how="left",
            suffixes=("", "_leap"),
        )
        has_numbers = pd.to_numeric(merged["leap_value"], errors="coerce").notna()
        with_data = merged[has_numbers]
        if not with_data.empty:
            sample = with_data[["sheet", "fuel_label", "sector_code_9th"]].drop_duplicates().head(12)
            sample_table = sample.rename(
                columns={
                    "sheet": "Sheet",
                    "fuel_label": "Fuel label",
                    "sector_code_9th": "9th sector code",
                }
            ).to_string(index=False)
            cols_to_save = [
                "sheet",
                "fuel_label",
                "sector_code_9th",
                "leap_value",
                "year",
                "scenario",
                "region",
            ]
            with_data[cols_to_save].to_csv(unmapped_path, index=False)
            raise RuntimeError(
                "Unmapped fuels with data detected. "
                f"Total rows: {len(with_data)}. "
                f"See details: {unmapped_path} and {mapping_status_path}. "
                "Check canonical mappings in config/sector_fuel_codes_to_names.xlsx "
                "and config/ninth_pairs_to_esto_pairs.xlsx; "
                "use config/backup_leap_mappings.xlsx only for explicit overrides. "
                "Also verify sector sheet mapping in config/leap_results_sheet_map.csv. "
                "Examples (first 12):\n"
                f"{sample_table}"
            )
        else:
            uniq = unmapped[["sheet", "fuel_label"]].drop_duplicates()
            print(
                f"[WARN] {len(uniq)} unmapped fuel(s) with no LEAP values. "
                f"Review {mapping_status_path} for details."
            )
            _remove_stale_unmapped_file(unmapped_path)
    else:
        _remove_stale_unmapped_file(unmapped_path)

    charts_dir = layout.charts
    written_charts: list[Path] = []
    dashboard_index: Path | None = None
    if GENERATE_CHARTS:
        written_charts = build_charts(
            comparison_long,
            charts_dir=charts_dir,
            backend=CHART_BACKEND,
            hide_leap_only_charts=HIDE_LEAP_ONLY_CHARTS,
        )
    else:
        print("[INFO] Skipped chart generation (GENERATE_CHARTS=False).")
    if GENERATE_DASHBOARDS:
        if not GENERATE_CHARTS:
            print("[WARN] GENERATE_DASHBOARDS=True but GENERATE_CHARTS=False; dashboards may reference stale charts.")
        dashboard_index = build_dashboards(
            output_dir=layout.root,
            comparison_long=comparison_long,
            charts_dir=charts_dir,
            mapping_status=mapping_status,
        )
    else:
        print("[INFO] Skipped dashboard generation (GENERATE_DASHBOARDS=False).")

    checks = basic_checks(sheet_map, fuel_aliases, comparison_long, mapping_status)

    manifest = write_output_manifest(
        out_dir=layout.root,
        primary_outputs={
            "comparison_long": str(comparison_long_path),
            "comparison_wide": str(comparison_wide_path),
            "mapping_status": str(mapping_status_path),
            "leap_long": str(leap_long_path),
            "dashboards_dir": str(layout.dashboards),
            "charts_dir": str(layout.charts),
        },
        supporting_outputs={
            "gap_diagnostics": diagnostics_artifacts.get("gap_diagnostics"),
            "mapping_rundown_by_sheet": diagnostics_artifacts.get("mapping_rundown_by_sheet"),
            "mapping_rundown_details": diagnostics_artifacts.get("mapping_rundown_details"),
            "comparison_issue_summary": issue_summary_path,
            "comparison_issue_cause_summary": str(layout.diagnostics / "comparison_issue_cause_summary.csv")
            if (layout.diagnostics / "comparison_issue_cause_summary.csv").exists()
            else None,
            "unmapped_fuels_with_data": str(unmapped_path) if unmapped_path.exists() else None,
        },
        primary_output_descriptions={
            "comparison_long": "Main long-form comparison table used for downstream analysis and charting.",
            "comparison_wide": "Wide-form comparison table with one column per source for quick spot checks.",
            "mapping_status": "Per-sheet and per-fuel mapping status workbook with mapping flags and sources.",
            "leap_long": "Normalized LEAP long table before comparator joins.",
            "dashboards_dir": "Rendered dashboard HTML pages.",
            "charts_dir": "Individual chart files used by the dashboards.",
        },
        supporting_output_descriptions={
            "gap_diagnostics": "Largest LEAP versus comparator gaps at the configured probe years.",
            "mapping_rundown_by_sheet": "Sheet-level summary of mapping completeness.",
            "mapping_rundown_details": "Detailed row-level mapping audit workbook.",
            "comparison_issue_summary": "Prioritized comparison issues with gap metrics and debug hints.",
            "comparison_issue_cause_summary": "Frequency summary of comparison issue categories.",
            "unmapped_fuels_with_data": "Unmapped fuel rows that still carry LEAP data and need mapping fixes.",
        },
        notes=[
            "Primary comparison outputs are written at the top level.",
            "Supporting diagnostics and mapping evidence are grouped under supporting_files/.",
        ],
    )

    return {
        "comparison_long": str(comparison_long_path),
        "comparison_wide": str(comparison_wide_path),
        "mapping_status": str(mapping_status_path),
        "leap_long": str(leap_long_path),
        "gap_diagnostics": diagnostics_artifacts.get("gap_diagnostics"),
        "mapping_rundown_by_sheet": diagnostics_artifacts.get("mapping_rundown_by_sheet"),
        "mapping_rundown_details": diagnostics_artifacts.get("mapping_rundown_details"),
        "comparison_issue_summary": issue_summary_path,
        "comparison_issue_cause_summary": str(issue_cause_summary_path) if issue_cause_summary_path else None,
        "charts_written": len(written_charts),
        "dashboard_index": str(dashboard_index) if dashboard_index else None,
        "diagnostics": checks,
        "output_manifest": str(manifest),
    }


# -----------------------------------------------------------------------------
# Bottom run block (ready for notebooks)
# -----------------------------------------------------------------------------
if __name__ == "__main__":  # pragma: no cover
    try:
        result = run_workflow()
        print("[OK] LEAP Results dashboard workflow complete.")
        for k, v in result.items():
            print(f"- {k}: {v}")
    except Exception as exc:  # noqa: BLE001
        print(f"[ERROR] Workflow failed: {exc}")

#%%


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
