#%%
"""
LEAP Results dashboard workflow (utilities repo)

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
    basic_checks,
    build_charts,
    build_comparisons,
    build_dashboards,
    ensure_repo_root,
    load_leap_workbook,
)


# -----------------------------------------------------------------------------
# Notebook-editable toggles (keep simple and explicit)
# -----------------------------------------------------------------------------
LEAP_RESULTS_DIR = REPO_ROOT / "outputs/leap_results"
# Economy/scenario filters for workbook discovery
ECONOMY_TOKEN = "USA"  # substring to match in filenames
SCENARIOS = ("Reference", "Target")
SHEET_MAP_PATH = DEFAULT_SHEET_MAP
BACKUP_MAPPINGS_PATH = DEFAULT_BACKUP_LEAP_MAPPINGS
CODEBOOK_PATH = DEFAULT_CODEBOOK
NINTH_TO_ESTO_PATH = DEFAULT_NINTH_TO_ESTO
BASE_TABLE_PATH = REPO_ROOT / "data/00APEC_2025_low_with_subtotals.csv"
PROJECTION_TABLE_PATH = REPO_ROOT / "data/merged_file_energy_ALL_20251106.csv"
OUTPUT_DIR = REPO_ROOT / "outputs/leap_results_dashboard/USA"
MAPPING_VIEWS_DIR = REPO_ROOT / "config/computer_generated_config/leap_mapping_views/USA"
BASE_YEAR = 2022  # from 00APEC_2025_low_with_subtotals.csv
PROJECTION_YEARS: Sequence[int] = tuple(range(2023, 2071))
SCENARIO_MAP = {"reference": "reference", "target": "target"}
BASE_ECONOMY = "20USA"
PROJECTION_ECONOMY = "20_USA"
CHART_BACKEND = "plotly"  # "plotly" or "static"
USE_ESTO_AGG_ONLY = False  # include both ESTO base-year and 9th projection on charts
SIBLING_COMPARATOR_MODE = "allocate_by_leap_share"  # "none" or "allocate_by_leap_share"
INCLUDE_SIBLING_PARENT_TOTALS = True  # add promoted parent-category chart rows (for example, "Road")


# -----------------------------------------------------------------------------
# Core workflow
# -----------------------------------------------------------------------------
def _resolve(path: Path | str) -> Path:
    """Resolve a path relative to repo root if not absolute."""
    p = Path(path)
    return p if p.is_absolute() else (REPO_ROOT / p)


_HEADER_NOTES = {
    "mapping_source": {
        "canonical": "Matched directly from config/ninth_pairs_to_esto_pairs.xlsx (main sheet).",
        "codebook_fallback": "Matched from config/sector_fuel_codes_to_names.xlsx using the ESTO_LEAP_names or code_to_name sheet.",
        "override": "Matched from config/backup_leap_mappings.xlsx (manual override).",
    },
    "flow_source": {
        "canonical": "ESTO flow came from config/ninth_pairs_to_esto_pairs.xlsx (main sheet).",
        "sector_fallback": "ESTO flow came from config/sector_fuel_codes_to_names.xlsx, sheet code_to_name.",
        "sheet_override": "ESTO flow came from config/leap_results_sheet_map.csv, column esto_flow_override.",
        "override": "ESTO flow came from config/backup_leap_mappings.xlsx (manual override).",
    },
    "fuel_source": {
        "canonical": "9th fuel code came from config/ninth_pairs_to_esto_pairs.xlsx (main sheet).",
        "inferred": "9th fuel code was inferred by the workflow from the ESTO flow and product after the initial lookup.",
        "override": "9th fuel code came from config/backup_leap_mappings.xlsx (manual override).",
    },
    "mapped": {
        "true": "At least one of ninth_fuel_code, esto_flow, or esto_product was filled.",
        "false": "All of ninth_fuel_code, esto_flow, and esto_product are blank.",
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


def _write_gap_and_mapping_diagnostics(
    *,
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    out_dir: Path,
    base_year: int,
    projection_probe_year: int = 2030,
) -> dict[str, str]:
    """
    Write compact diagnostics files for large comparison gaps and mapping completeness.
    Returns a dict of artifact paths.
    """
    artifacts: dict[str, str] = {}

    # ---- Gap diagnostics (LEAP vs base/projection) ----
    gap_path = out_dir / "comparison_gap_diagnostics.csv"
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
        for col in ["ninth_fuel_code", "esto_flow", "esto_product", "mapping_note"]:
            if col not in status.columns:
                status[col] = ""
            status[col] = status[col].fillna("").astype(str).str.strip()

        status["missing_ninth_fuel"] = status["ninth_fuel_code"] == ""
        status["missing_esto_flow"] = status["esto_flow"] == ""
        status["missing_esto_product"] = status["esto_product"] == ""
        status["has_mapping_note"] = status["mapping_note"] != ""

        detail_path = out_dir / "mapping_rundown_details.xlsx"
        legacy_detail_csv_path = out_dir / "mapping_rundown_details.csv"
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
                missing_ninth_fuel=("missing_ninth_fuel", "sum"),
                missing_esto_flow=("missing_esto_flow", "sum"),
                missing_esto_product=("missing_esto_product", "sum"),
                with_mapping_notes=("has_mapping_note", "sum"),
            )
            .sort_values(
                by=["missing_ninth_fuel", "missing_esto_flow", "missing_esto_product", "rows"],
                ascending=[False, False, False, False],
            )
        )
        by_sheet_path = out_dir / "mapping_rundown_by_sheet.csv"
        by_sheet.to_csv(by_sheet_path, index=False)
        artifacts["mapping_rundown_by_sheet"] = str(by_sheet_path)

    return artifacts


def run_workflow() -> dict[str, object]:
    ensure_repo_root()
    out_dir = _resolve(OUTPUT_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)
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
    sector_flow_mapping = build_sector_to_esto_flow_lookup(_resolve(CODEBOOK_PATH))
    ninth_pairs, canonical_conflicts = load_canonical_pairs(_resolve(NINTH_TO_ESTO_PATH), strict=False)
    if not canonical_conflicts.empty:
        mapping_views_dir = _resolve(MAPPING_VIEWS_DIR)
        mapping_views_dir.mkdir(parents=True, exist_ok=True)
        conflict_path = mapping_views_dir / "mapping_conflicts.csv"
        canonical_conflicts.to_csv(conflict_path, index=False)
        print(
            "[WARN] Canonical key conflicts found in ninth_pairs_to_esto_pairs.xlsx "
            f"({len(canonical_conflicts)} row(s)); using deterministic first match per key. "
            f"Details: {conflict_path}"
        )
    base_df = pd.read_csv(_resolve(BASE_TABLE_PATH))
    ninth_df = pd.read_csv(_resolve(PROJECTION_TABLE_PATH))

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
    comparison_long_path = out_dir / "comparison_long.csv"
    comparison_wide_path = out_dir / "comparison_wide.csv"
    mapping_status_path = out_dir / "mapping_status.xlsx"
    legacy_mapping_status_csv_path = out_dir / "mapping_status.csv"
    leap_long_path = out_dir / "leap_long.csv"

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
            out_dir=out_dir,
            base_year=BASE_YEAR,
        )
        for key, path in diagnostics_artifacts.items():
            print(f"[INFO] Wrote {key}: {path}")
    except PermissionError as exc:
        print(
            "[WARN] Could not refresh one or more diagnostics files because a file is in use. "
            f"Close it and rerun if needed: {exc}"
        )

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
    unmapped_path = out_dir / "unmapped_fuels_with_data.csv"
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

    charts_dir = out_dir / "charts"
    written_charts = build_charts(comparison_long, charts_dir=charts_dir, backend=CHART_BACKEND)
    dashboard_index = build_dashboards(
        output_dir=out_dir,
        comparison_long=comparison_long,
        charts_dir=charts_dir,
        mapping_status=mapping_status,
    )

    checks = basic_checks(sheet_map, fuel_aliases, comparison_long, mapping_status)

    return {
        "comparison_long": str(comparison_long_path),
        "comparison_wide": str(comparison_wide_path),
        "mapping_status": str(mapping_status_path),
        "leap_long": str(leap_long_path),
        "gap_diagnostics": diagnostics_artifacts.get("gap_diagnostics"),
        "mapping_rundown_by_sheet": diagnostics_artifacts.get("mapping_rundown_by_sheet"),
        "mapping_rundown_details": diagnostics_artifacts.get("mapping_rundown_details"),
        "charts_written": len(written_charts),
        "dashboard_index": str(dashboard_index) if dashboard_index else None,
        "diagnostics": checks,
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
