#%%
"""
Build the V2 LEAP results dashboard with mapping diagnostics and ledgers.

This workflow loads exported LEAP results, reference ESTO/9th data, mapping
configuration, and explicit reassignment rules into one comparison pipeline.
It writes the long comparison table plus mapping status, chart ledgers, gap
diagnostics, and dashboard HTML outputs used to audit the V2 dashboard.
"""

from __future__ import annotations

import sys
import gzip
import shutil
import os
import time
import importlib
import json
from pathlib import Path
from typing import Sequence

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.mappings.canonical_mapping import (  # noqa: E402
    DEFAULT_BACKUP_LEAP_MAPPINGS,
    DEFAULT_CODEBOOK,
    DEFAULT_NINTH_TO_ESTO,
    DEFAULT_SHEET_MAP,
)
from codebase.utilities import leap_results_dashboard_utils as _dashboard_utils  # noqa: E402
_dashboard_utils = importlib.reload(_dashboard_utils)

DEFAULT_EXPLICIT_LEAP_MAPPINGS = _dashboard_utils.DEFAULT_EXPLICIT_LEAP_MAPPINGS
DEFAULT_EXPLICIT_LEAP_REASSIGNMENTS = _dashboard_utils.DEFAULT_EXPLICIT_LEAP_REASSIGNMENTS
build_charts = _dashboard_utils.build_charts
build_dashboards = _dashboard_utils.build_dashboards
ensure_repo_root = _dashboard_utils.ensure_repo_root
from codebase.utilities.leap_results_dashboard_v2.comparison_engine import (  # noqa: E402
    build_chart_line_mapping_ledger,
    build_total_rows_for_charts,
    build_total_component_ledger,
    build_comparisons_v2,
    collapse_sheets_by_sector_name_for_charts,
    add_parent_totals_from_dashboard_overrides,
    _drop_parent_child_comparator_overlaps,
)
from codebase.utilities.leap_results_dashboard_v2.config_loader import (  # noqa: E402
    load_mapping_inputs,
    write_canonical_conflicts,
)
from codebase.utilities.leap_results_dashboard_v2.diagnostics import (  # noqa: E402
    run_basic_checks,
    write_diagnostics,
)
from codebase.utilities.leap_results_dashboard_v2.atomic_engine import (  # noqa: E402
    build_atomic_outputs,
    build_shadow_delta_reports,
)
from codebase.utilities.leap_results_dashboard_v2.leap_loader import (  # noqa: E402
    discover_workbooks,
    load_leap_long,
)
from codebase.utilities.leap_results_dashboard_v2.models import AtomicSettings, DashboardV2Settings  # noqa: E402
from codebase.utilities.leap_results_dashboard_v2.output_writer import write_core_outputs  # noqa: E402
from codebase.utilities.ninth_to_esto_mapping_coverage import (  # noqa: E402
    run_mapping_coverage_check,
)
from codebase.utilities.leap_results_dashboard_v2.pathing import resolve_path  # noqa: E402
from codebase.utilities.leap_results_dashboard_v2.reference_loader import load_reference_tables  # noqa: E402
from codebase.utilities.leap_results_dashboard_v2.shadow_compare import compare_outputs  # noqa: E402
from codebase.utilities.leap_results_dashboard_v2.derived_transformation_metrics import (  # noqa: E402
    build_derived_transformation_comparison_rows,
    build_derived_transformation_leap_long,
    write_derived_transformation_artifacts,
)
from codebase.utilities.workflow_common import archive_config_dir_once_per_day  # noqa: E402
from codebase.utilities.workflow_outputs import build_workflow_output_layout, write_output_manifest  # noqa: E402

################################
LEAP_RESULTS_DIR = REPO_ROOT / "data/leap results tables"
ECONOMY_TOKEN = "USA"
SCENARIOS = ("Reference", "Target")
SHEET_MAP_PATH = DEFAULT_SHEET_MAP
BACKUP_MAPPINGS_PATH = DEFAULT_BACKUP_LEAP_MAPPINGS
EXPLICIT_MAPPINGS_PATH = DEFAULT_EXPLICIT_LEAP_MAPPINGS
EXPLICIT_REASSIGNMENTS_PATH = DEFAULT_EXPLICIT_LEAP_REASSIGNMENTS
SYNTHETIC_REFERENCE_ROWS_PATH = REPO_ROOT / "config/synthetic_reference_rows.csv"
CODEBOOK_PATH = DEFAULT_CODEBOOK
NINTH_TO_ESTO_PATH = DEFAULT_NINTH_TO_ESTO
BASE_TABLE_PATH = REPO_ROOT / "data/00APEC_2025_low_with_subtotals.csv"
PROJECTION_TABLE_PATH = REPO_ROOT / "data/merged_file_energy_ALL_20251106.csv"
OUTPUT_DIR = REPO_ROOT / "outputs/dashboards/leap_results_dashboard_v2/USA"
V1_OUTPUT_DIR = REPO_ROOT / "outputs/dashboards/leap_results_dashboard/USA"
MAPPING_VIEWS_DIR = REPO_ROOT / "config/computer_generated_config/leap_mapping_views/USA"
BASE_YEAR = 2022
MAX_OUTPUT_YEAR = 2060
PROJECTION_YEARS: Sequence[int] = tuple(range(2023, MAX_OUTPUT_YEAR + 1))
SCENARIO_MAP = {"reference": "reference", "target": "target"}
BASE_ECONOMY = "20USA"
PROJECTION_ECONOMY = "20_USA"
CHART_BACKEND = "plotly"
USE_ESTO_AGG_ONLY = False
SIBLING_COMPARATOR_MODE = "aggregate_to_parent"
INCLUDE_SIBLING_PARENT_TOTALS = True
GENERATE_CHARTS = True
GENERATE_DASHBOARDS = True
HIDE_LEAP_ONLY_CHARTS = False
SHOW_PARENT_SCOPED_CHILD_CONTEXT_COMPARATORS = False
DIAGNOSTIC_PROBE_YEAR = 2030
TOP_DIAGNOSTIC_ROWS = 40
DROP_ALL_ZERO_BASE_ROWS = True
DROP_ALL_ZERO_PROJECTION_ROWS = False
RUN_SHADOW_COMPARE = False
CHART_SNAPSHOT_RETENTION_DAYS = 60
FILTER_TO_SHEET_MAP = False
HIDE_SECTORS_FROM_DASHBOARD = {"Other sector",}
ALLOW_SHEETS_NOT_IN_MAP = {
    "Buildings",
    "Residential",
    "Commercial and public services",
}

################################
def _env_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return str(raw).strip().lower() in {"1", "true", "yes", "y", "on"}


def _coerce_bool_series(values: pd.Series, default: bool = True) -> pd.Series:
    text = values.fillna(default).astype(str).str.strip().str.lower()
    true_tokens = {"1", "true", "yes", "y", "t"}
    false_tokens = {"0", "false", "no", "n", "f"}
    out = pd.Series(default, index=values.index, dtype="bool")
    out.loc[text.isin(true_tokens)] = True
    out.loc[text.isin(false_tokens)] = False
    return out


def _rebuild_comparison_wide(comparison_long: pd.DataFrame) -> pd.DataFrame:
    if comparison_long.empty:
        return comparison_long.copy()
    wide = (
        comparison_long.pivot_table(
            index=["economy", "scenario", "sheet", "fuel_label", "year"],
            columns="source",
            values="value",
            aggfunc="first",
        )
        .reset_index()
    )
    if hasattr(wide.columns, "name"):
        wide.columns.name = None
    return wide


def _chart_output_relative_path(sheet: object, measure: object, fuel: object, *, backend: str) -> str:
    measure_text = str(measure or "").strip()
    sheet_text = str(sheet or "").replace("\\", "_")
    sheet_key = f"{sheet_text}__{measure_text}" if measure_text else sheet_text
    sheet_slug = _dashboard_utils._safe_token(sheet_key)
    fuel_slug = _dashboard_utils._safe_token(str(fuel))
    suffix = ".html" if str(backend).strip().lower() == "plotly" else ".png"
    return f"charts/{sheet_slug}__{fuel_slug}{suffix}"


def _sector_category_sheets(sheet_map: pd.DataFrame, mapping_status: pd.DataFrame) -> set[str]:
    hidden: set[str] = set()
    if not sheet_map.empty and {"sheet_name", "category_type"}.issubset(sheet_map.columns):
        sm = sheet_map.copy()
        sm["sheet_name"] = sm["sheet_name"].fillna("").astype(str).str.strip()
        sm["category_type"] = sm["category_type"].fillna("").astype(str).str.strip().str.lower()
        hidden.update(sm.loc[sm["category_type"].eq("sector"), "sheet_name"].tolist())
    if not mapping_status.empty and "sheet" in mapping_status.columns:
        ms = mapping_status.copy()
        mapping_source = (
            ms["mapping_source"].fillna("").astype(str).str.strip().str.lower()
            if "mapping_source" in ms.columns
            else pd.Series("", index=ms.index, dtype="object")
        )
        mapping_note = (
            ms["mapping_note"].fillna("").astype(str)
            if "mapping_note" in ms.columns
            else pd.Series("", index=ms.index, dtype="object")
        )
        category_mask = mapping_source.eq("category_sector") | mapping_note.str.contains(
            "category labels treated as sectors",
            case=False,
            regex=False,
        )
        hidden.update(ms.loc[category_mask, "sheet"].dropna().astype(str).str.strip().tolist())
    return {sheet for sheet in hidden if sheet}


def _display_dashboard_path(path: Sequence[str]) -> list[str]:
    top_label_overrides = {
        "Industry sector": "Industry",
        "Transport sector": "Transport",
        "Other sector": "Other",
    }
    if not path:
        return []
    display_path = [str(part).strip() for part in path if str(part).strip()]
    if display_path:
        display_path[0] = top_label_overrides.get(display_path[0], display_path[0])
    return display_path


def _write_chart_navigation_hierarchy(
    *,
    chart_input: pd.DataFrame,
    mapping_status: pd.DataFrame,
    sheet_map: pd.DataFrame,
    out_dir: Path,
    charts_dir: Path,
    backend: str,
    hide_leap_only_charts: bool,
) -> tuple[Path, Path]:
    """
    Write the dashboard-rendered chart hierarchy as data.

    The JSON is intentionally built from chart_input and build_dashboards'
    sheet-path resolver so it tracks the same hierarchy used by the HTML pages.
    """
    render_long = _dashboard_utils._prepare_render_long(chart_input)
    sheet_paths = build_dashboards(
        output_dir=out_dir,
        comparison_long=render_long,
        charts_dir=charts_dir,
        mapping_status=mapping_status,
        return_sheet_paths=True,
    )
    if not isinstance(sheet_paths, dict):
        sheet_paths = {}

    hidden_sheets = _sector_category_sheets(sheet_map, mapping_status)
    rows: list[dict[str, str]] = []
    hierarchy: dict[str, object] = {}

    for (sheet, measure, fuel), sub in render_long.groupby(["sheet", "measure", "fuel_label"], dropna=False):
        sheet_text = str(sheet).strip()
        if not sheet_text or sheet_text in hidden_sheets:
            continue
        values = pd.to_numeric(sub["value"], errors="coerce").fillna(0.0)
        if not values.ne(0).any():
            continue
        force_show_chart = (
            bool(sub["force_show_chart"].fillna(False).astype(bool).any())
            if "force_show_chart" in sub.columns
            else False
        )
        if hide_leap_only_charts:
            non_leap_sources = {
                str(source).strip()
                for source in sub["source"].dropna().astype(str)
                if str(source).strip() and str(source).strip() != "leap"
            }
            if not non_leap_sources and not force_show_chart:
                continue

        path = [str(part).strip() for part in sheet_paths.get(sheet_text, [sheet_text]) if str(part).strip()]
        if not path:
            path = [sheet_text]
        path = _display_dashboard_path(path)
        measure_text = str(measure or "").strip()
        fuel_text = str(fuel or "").strip()
        chart_file = _chart_output_relative_path(sheet_text, measure_text, fuel_text, backend=backend)

        node = hierarchy
        for part in path:
            node = node.setdefault(part, {})  # type: ignore[assignment]
        fuels = node.setdefault("fuels", {})  # type: ignore[union-attr]
        chart_record = {
            "sheet": sheet_text,
            "measure": measure_text,
            "chart_file": chart_file,
        }
        fuels.setdefault(fuel_text, []).append(chart_record)  # type: ignore[union-attr]

        rows.append(
            {
                "dashboard_path": " > ".join(path),
                "sheet": sheet_text,
                "measure": measure_text,
                "fuel": fuel_text,
                "chart_file": chart_file,
            }
        )

    def _sort_nested(value: object) -> object:
        if isinstance(value, dict):
            return {key: _sort_nested(value[key]) for key in sorted(value)}
        if isinstance(value, list):
            return sorted(value, key=lambda item: tuple(str(item.get(col, "")) for col in ["measure", "sheet", "chart_file"]))
        return value

    json_path = out_dir / "chart_navigation_hierarchy.json"
    csv_path = out_dir / "chart_navigation_hierarchy.csv"
    with json_path.open("w", encoding="utf-8") as handle:
        json.dump(_sort_nested(hierarchy), handle, indent=2, ensure_ascii=True)
        handle.write("\n")
    pd.DataFrame(rows).sort_values(["dashboard_path", "fuel", "measure", "sheet"]).to_csv(csv_path, index=False)
    return json_path, csv_path


def _filter_to_sheet_map(
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    sheet_map: pd.DataFrame,
    *,
    out_dir: Path,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if comparison_long.empty or sheet_map.empty:
        return comparison_long, mapping_status
    sm = sheet_map.copy()
    for col in ["sheet_name", "active"]:
        if col not in sm.columns:
            sm[col] = ""
        sm[col] = sm[col].fillna("").astype(str).str.strip()
    if "active" in sm.columns:
        sm = sm[sm["active"].str.lower().isin({"true", "1", "yes", "y"})].copy()
    allowed = {str(name).strip() for name in sm["sheet_name"].tolist() if str(name).strip()}
    allowed |= {str(name).strip() for name in ALLOW_SHEETS_NOT_IN_MAP if str(name).strip()}
    if not allowed:
        return comparison_long, mapping_status
    comp = comparison_long.copy()
    comp["sheet"] = comp["sheet"].fillna("").astype(str).str.strip()
    skipped = sorted(set(comp["sheet"].unique()) - allowed)
    if skipped:
        print(f"[WARN] Skipping {len(skipped)} sheets not in sheet map.", flush=True)
        ms = mapping_status.copy()
        if not ms.empty and "sheet" in ms.columns:
            ms["sheet"] = ms["sheet"].fillna("").astype(str).str.strip()
        summary_rows = []
        for sheet in skipped:
            summary_rows.append(
                {
                    "sheet": sheet,
                    "comparison_rows": int((comp["sheet"] == sheet).sum()),
                    "mapping_rows": int((ms["sheet"] == sheet).sum()) if not ms.empty and "sheet" in ms.columns else 0,
                }
            )
        skip_df = pd.DataFrame(summary_rows).sort_values("sheet")
        try:
            skip_df.to_csv(out_dir / "skipped_sheets_not_in_map.csv", index=False)
        except PermissionError as exc:
            print(f"[WARN] Skipping write to locked skipped-sheets output: {exc}", flush=True)
    comp = comp[comp["sheet"].isin(allowed)].copy()
    if not mapping_status.empty and "sheet" in mapping_status.columns:
        mapping_status = mapping_status[mapping_status["sheet"].fillna("").astype(str).str.strip().isin(allowed)].copy()
    return comp, mapping_status


def _synthetic_row_mask(frame: pd.DataFrame, column: str) -> pd.Series:
    if column not in frame.columns:
        return pd.Series(False, index=frame.index)
    return (
        frame[column]
        .astype(str)
        .str.strip()
        .str.lower()
        .isin({"true", "1", "yes", "y", "t"})
    )


def _build_transformation_input_addback_pairs(sheet_map: pd.DataFrame) -> list[tuple[str, str]]:
    if sheet_map.empty or "sheet_name" not in sheet_map.columns:
        return []
    names = {
        str(name).strip()
        for name in sheet_map["sheet_name"].fillna("").astype(str).tolist()
        if str(name).strip()
    }
    pairs: list[tuple[str, str]] = []
    for loss_sheet in sorted(name for name in names if name.endswith("_loss_own_use_total")):
        prefix = loss_sheet[: -len("_loss_own_use_total")]
        input_sheet = f"{prefix}_inputs"
        if input_sheet in names:
            pairs.append((input_sheet, loss_sheet))
    return pairs


def _add_loss_own_use_to_transformation_inputs_in_comparison(
    comparison_long: pd.DataFrame,
    *,
    sheet_map: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if comparison_long.empty:
        return comparison_long, pd.DataFrame()
    pairs = _build_transformation_input_addback_pairs(sheet_map)
    if not pairs:
        return comparison_long, pd.DataFrame()
    out = comparison_long.copy()
    out["value"] = pd.to_numeric(out["value"], errors="coerce")
    adjustments: list[pd.DataFrame] = []
    for input_sheet, loss_sheet in pairs:
        input_mask = out["sheet"].astype(str).eq(input_sheet)
        if not input_mask.any():
            continue
        loss_rows = out[out["sheet"].astype(str).eq(loss_sheet)].copy()
        if loss_rows.empty:
            continue
        group_cols = ["economy", "scenario", "source", "year", "fuel_label"]
        loss_add = (
            loss_rows.groupby(group_cols, dropna=False)["value"]
            .sum(min_count=1)
            .reset_index()
            .rename(columns={"value": "_loss_own_use_addback"})
        )
        target = out.loc[input_mask].copy()
        target["_row_index"] = target.index
        target = target.merge(loss_add, on=group_cols, how="left")
        target["_loss_own_use_addback"] = pd.to_numeric(target["_loss_own_use_addback"], errors="coerce").fillna(0.0)
        out.loc[target["_row_index"], "value"] = (
            pd.to_numeric(target["value"], errors="coerce") + target["_loss_own_use_addback"]
        ).to_numpy()
        adj = target[group_cols + ["_loss_own_use_addback"]].copy()
        adj["input_sheet"] = input_sheet
        adj["loss_sheet"] = loss_sheet
        adjustments.append(adj.rename(columns={"_loss_own_use_addback": "addback_value"}))
    if not adjustments:
        return out, pd.DataFrame()
    adjustment_report = pd.concat(adjustments, ignore_index=True).sort_values(
        ["input_sheet", "scenario", "source", "year", "fuel_label"],
        kind="mergesort",
    )
    return out, adjustment_report


def _add_loss_own_use_to_transformation_inputs_in_leap_long(
    leap_long: pd.DataFrame,
    *,
    sheet_map: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if leap_long.empty:
        return leap_long, pd.DataFrame()
    pairs = _build_transformation_input_addback_pairs(sheet_map)
    if not pairs:
        return leap_long, pd.DataFrame()
    out = leap_long.copy()
    out["leap_value"] = pd.to_numeric(out["leap_value"], errors="coerce")
    adjustments: list[pd.DataFrame] = []
    for input_sheet, loss_sheet in pairs:
        input_mask = out["sheet_name"].astype(str).eq(input_sheet)
        if not input_mask.any():
            continue
        loss_rows = out[out["sheet_name"].astype(str).eq(loss_sheet)].copy()
        if loss_rows.empty:
            continue
        group_cols = ["economy", "scenario", "region", "year", "fuel_label"]
        loss_add = (
            loss_rows.groupby(group_cols, dropna=False)["leap_value"]
            .sum(min_count=1)
            .reset_index()
            .rename(columns={"leap_value": "_loss_own_use_addback"})
        )
        target = out.loc[input_mask].copy()
        target["_row_index"] = target.index
        target = target.merge(loss_add, on=group_cols, how="left")
        target["_loss_own_use_addback"] = pd.to_numeric(target["_loss_own_use_addback"], errors="coerce").fillna(0.0)
        out.loc[target["_row_index"], "leap_value"] = (
            pd.to_numeric(target["leap_value"], errors="coerce") + target["_loss_own_use_addback"]
        ).to_numpy()
        adj = target[group_cols + ["_loss_own_use_addback"]].copy()
        adj["input_sheet"] = input_sheet
        adj["loss_sheet"] = loss_sheet
        adjustments.append(adj.rename(columns={"_loss_own_use_addback": "addback_value"}))
    if not adjustments:
        return out, pd.DataFrame()
    adjustment_report = pd.concat(adjustments, ignore_index=True).sort_values(
        ["input_sheet", "scenario", "year", "fuel_label"],
        kind="mergesort",
    )
    return out, adjustment_report


def _drop_loss_own_use_sheets(
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    leap_long: pd.DataFrame | None = None,
    *,
    keep_loss_sheets: set[str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame | None]:
    keep = {str(name).strip() for name in (keep_loss_sheets or set()) if str(name).strip()}
    loss_mask = comparison_long["sheet"].astype(str).str.endswith("_loss_own_use_total")
    if keep:
        keep_mask = comparison_long["sheet"].astype(str).isin(keep)
        loss_mask = loss_mask & ~keep_mask
    comp = comparison_long.copy()
    comp = comp[~loss_mask].copy()
    ms = mapping_status.copy()
    if not ms.empty and "sheet" in ms.columns:
        ms_loss_mask = ms["sheet"].astype(str).str.endswith("_loss_own_use_total")
        if keep:
            ms_loss_mask = ms_loss_mask & ~ms["sheet"].astype(str).isin(keep)
        ms = ms[~ms_loss_mask].copy()
    leap_out = leap_long
    if leap_out is not None and not leap_out.empty and "sheet_name" in leap_out.columns:
        leap_loss_mask = leap_out["sheet_name"].astype(str).str.endswith("_loss_own_use_total")
        if keep:
            leap_loss_mask = leap_loss_mask & ~leap_out["sheet_name"].astype(str).isin(keep)
        leap_out = leap_out[~leap_loss_mask].copy()
    return comp, ms, leap_out


def _retitle_combined_transformation_input_measures(
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    leap_long: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame | None]:
    def _apply(frame: pd.DataFrame, sheet_col: str) -> pd.DataFrame:
        if frame.empty or sheet_col not in frame.columns or "measure" not in frame.columns:
            return frame
        out = frame.copy()
        mask = out[sheet_col].astype(str).str.endswith("_inputs")
        out.loc[mask, "measure"] = TRANSFORMATION_COMBINED_INPUT_MEASURE_LABEL
        return out

    comp = _apply(comparison_long, "sheet")
    ms = _apply(mapping_status, "sheet")
    leap_out = _apply(leap_long, "sheet_name") if leap_long is not None else leap_long
    return comp, ms, leap_out


def _retitle_trans_dist_loss_measure(
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    leap_long: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame | None]:
    def _apply(frame: pd.DataFrame, sheet_col: str) -> pd.DataFrame:
        if frame.empty or sheet_col not in frame.columns or "measure" not in frame.columns:
            return frame
        out = frame.copy()
        mask = out[sheet_col].astype(str).eq("trans_dist_loss_own_use_total")
        out.loc[mask, "measure"] = TRANS_DIST_LOSS_MEASURE_LABEL
        return out

    comp = _apply(comparison_long, "sheet")
    ms = _apply(mapping_status, "sheet")
    leap_out = _apply(leap_long, "sheet_name") if leap_long is not None else leap_long
    return comp, ms, leap_out


PRODUCTION_FAST_MODE = _env_bool("PRODUCTION_FAST_MODE", True)
ATOMIC_ENABLED = _env_bool("ATOMIC_ENABLED", not PRODUCTION_FAST_MODE)
ATOMIC_ROLLOUT_MODE = os.getenv("ATOMIC_ROLLOUT_MODE", "atomic")
ATOMIC_MANY_TO_MANY_POLICY = os.getenv("ATOMIC_MANY_TO_MANY_POLICY", "error")
ATOMIC_WRITE_SHADOW_OUTPUTS = _env_bool("ATOMIC_WRITE_SHADOW_OUTPUTS", False)
WRITE_DIAGNOSTIC_ARTIFACTS = _env_bool("WRITE_DIAGNOSTIC_ARTIFACTS", not PRODUCTION_FAST_MODE)
WRITE_CHART_SNAPSHOTS = _env_bool("WRITE_CHART_SNAPSHOTS", not PRODUCTION_FAST_MODE)
WRITE_CHART_LEDGERS = _env_bool("WRITE_CHART_LEDGERS", not PRODUCTION_FAST_MODE)
REFRESH_LEAP_RESULTS_BEFORE_DASHBOARD = _env_bool("REFRESH_LEAP_RESULTS_BEFORE_DASHBOARD", False)
RUN_MAPPING_COVERAGE_CHECK = _env_bool("RUN_MAPPING_COVERAGE_CHECK", not PRODUCTION_FAST_MODE)
DERIVED_TRANSFORMATION_METHOD = os.getenv("DERIVED_TRANSFORMATION_METHOD", "aux_direct").strip().lower()
TRANSFORMATION_DISPLAY_MODE = os.getenv("TRANSFORMATION_DISPLAY_MODE", "combined").strip().lower()
TRANSFORMATION_COMBINED_INPUT_MEASURE_LABEL = "Transformation inputs + losses & own use (PJ)"
COMBINED_MODE_KEEP_LOSS_SHEETS = {"trans_dist_loss_own_use_total"}
TRANS_DIST_LOSS_MEASURE_LABEL = "Transmission and distribution losses (PJ)"

################################

V2_SETTINGS = DashboardV2Settings(
    mapping_graph_mode="common_level_only",
    mapping_precedence="explicit_canonical_fallback",
    ambiguous_policy="aggregate",
    leaf_hole_policy="fail_fast",
)
ATOMIC_SETTINGS = AtomicSettings(
    enabled=ATOMIC_ENABLED,
    rollout_mode=ATOMIC_ROLLOUT_MODE,
    many_to_many_policy=ATOMIC_MANY_TO_MANY_POLICY,
    write_shadow_outputs=ATOMIC_WRITE_SHADOW_OUTPUTS,
)
################################

def _refresh_leap_results_inputs() -> dict[str, object]:
    """
    Refill LEAP template workbooks used directly by the dashboard.

    `leap_results_workflow.py` overwrites the source templates under
    `data/leap results tables`, and this dashboard now reads those workbooks
    directly.
    """
    import codebase.leap_results_workflow as _leap_results_workflow

    _leap_results_workflow = importlib.reload(_leap_results_workflow)
    TEMPLATE_PATHS = _leap_results_workflow.TEMPLATE_PATHS
    run_template_fill = _leap_results_workflow.run_template_fill

    refresh_log = run_template_fill()
    refresh_log["synced_workbooks"] = [
        str((REPO_ROOT / template_path).resolve())
        for template_path in TEMPLATE_PATHS
        if (REPO_ROOT / template_path).resolve().exists()
    ]
    return refresh_log


def _build_chart_input_for_rendering(
    *,
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    sheet_map: pd.DataFrame,
    child_context_comparison_long: pd.DataFrame | None = None,
) -> pd.DataFrame:
    comparison_long = comparison_long[
        ~comparison_long["sheet"].map(_dashboard_utils._suppress_feedstock_output_chart)
    ].copy()
    mapping_status = mapping_status[
        ~mapping_status["sheet"].map(_dashboard_utils._suppress_feedstock_output_chart)
    ].copy()
    if "measure" not in comparison_long.columns:
        comparison_long["measure"] = ""
    comparison_long["measure"] = comparison_long["measure"].fillna("").astype(str)
    if "measure" not in mapping_status.columns:
        mapping_status["measure"] = ""
    mapping_status["measure"] = mapping_status["measure"].fillna("").astype(str)

    mapping_meta_cols = ["sheet", "measure", "fuel_label"]
    for col in ["mapped", "partially_mapped", "has_any_mapping"]:
        if col not in mapping_status.columns:
            mapping_status[col] = False
        mapping_status[col] = mapping_status[col].fillna(False).astype(bool)
        mapping_meta_cols.append(col)
    for col in [
        "sector_code_9th",
        "ninth_fuel_code",
        "projection_parent_fallback",
        "projection_parent_sector_code",
        "comparator_scope",
    ]:
        if col not in mapping_status.columns:
            mapping_status[col] = ""
        mapping_meta_cols.append(col)

    comp = comparison_long.merge(
        mapping_status[mapping_meta_cols].drop_duplicates(
            subset=["sheet", "measure", "fuel_label"],
            keep="first",
        ),
        on=["sheet", "measure", "fuel_label"],
        how="left",
    )
    for col in ["mapped", "partially_mapped", "has_any_mapping"]:
        comp[col] = comp[col].fillna(False).astype(bool)
    if "projection_parent_fallback" in comp.columns:
        comp["projection_parent_fallback"] = comp["projection_parent_fallback"].fillna(False).astype(bool)
    for col in ["sector_code_9th", "ninth_fuel_code", "projection_parent_sector_code", "comparator_scope"]:
        if col in comp.columns:
            comp[col] = comp[col].fillna("").astype(str).str.strip()
    comp["used_in_total"] = True

    if child_context_comparison_long is not None and not child_context_comparison_long.empty:
        parent_scope = mapping_status.copy()
        if "comparator_scope" not in parent_scope.columns:
            parent_scope["comparator_scope"] = ""
        parent_scope["comparator_scope"] = parent_scope["comparator_scope"].fillna("").astype(str).str.strip().str.lower()
        parent_scope_keys = parent_scope[parent_scope["comparator_scope"].eq("parent")][
            ["sheet", "measure", "fuel_label"]
        ].drop_duplicates()
        if not parent_scope_keys.empty:
            ctx = child_context_comparison_long.copy()
            if "measure" not in ctx.columns:
                ctx["measure"] = ""
            ctx["measure"] = ctx["measure"].fillna("").astype(str)
            for col in ["sheet", "fuel_label", "scenario", "source"]:
                if col not in ctx.columns:
                    ctx[col] = ""
                ctx[col] = ctx[col].fillna("").astype(str)
            comparator_sources = {
                "base",
                "base_estimated",
                "base_mixed",
                "projection",
                "projection_estimated",
                "projection_mixed",
            }
            ctx = ctx[ctx["source"].isin(comparator_sources)].copy()
            ctx = ctx.merge(parent_scope_keys, on=["sheet", "measure", "fuel_label"], how="inner")
            if not ctx.empty:
                for col in mapping_meta_cols:
                    if col not in ctx.columns:
                        ctx[col] = ""
                ctx = ctx.merge(
                    mapping_status[mapping_meta_cols].drop_duplicates(
                        subset=["sheet", "measure", "fuel_label"],
                        keep="first",
                    ),
                    on=["sheet", "measure", "fuel_label"],
                    how="left",
                    suffixes=("", "_map"),
                )
                for col in mapping_meta_cols:
                    map_col = f"{col}_map"
                    if map_col in ctx.columns:
                        if col in {"mapped", "partially_mapped", "has_any_mapping"}:
                            ctx[col] = ctx[map_col].combine_first(ctx[col]).fillna(False).astype(bool)
                        else:
                            ctx[col] = (
                                ctx[map_col].combine_first(ctx[col]).fillna("").astype(str).str.strip()
                            )
                        ctx = ctx.drop(columns=[map_col], errors="ignore")
                for col in ["mapped", "partially_mapped", "has_any_mapping"]:
                    if col not in ctx.columns:
                        ctx[col] = False
                    ctx[col] = ctx[col].fillna(False).astype(bool)
                if "projection_parent_fallback" in ctx.columns:
                    ctx["projection_parent_fallback"] = ctx["projection_parent_fallback"].fillna(False).astype(bool)
                for col in ["sector_code_9th", "ninth_fuel_code", "projection_parent_sector_code", "comparator_scope"]:
                    if col not in ctx.columns:
                        ctx[col] = ""
                    ctx[col] = ctx[col].fillna("").astype(str).str.strip()
                ctx["used_in_total"] = False
                ctx["context_only_not_used_in_total"] = True
                keep_key = ["sheet", "measure", "fuel_label", "scenario", "source", "year"]
                existing_keys = comp[keep_key].drop_duplicates()
                ctx = ctx.merge(existing_keys.assign(_exists=True), on=keep_key, how="left")
                ctx = ctx[~ctx["_exists"].fillna(False).astype(bool)].drop(columns=["_exists"], errors="ignore")
                if not ctx.empty:
                    comp = pd.concat([comp, ctx], ignore_index=True, sort=False)
    comp["force_show_chart"] = (~comp["mapped"]) | comp["partially_mapped"] | (~comp["has_any_mapping"])

    if HIDE_SECTORS_FROM_DASHBOARD:
        hide = {str(name).strip() for name in HIDE_SECTORS_FROM_DASHBOARD if str(name).strip()}
        if hide:
            comp = comp[~comp["sheet"].isin(hide)].copy()
            if not mapping_status.empty and "sheet" in mapping_status.columns:
                mapping_status = mapping_status[
                    ~mapping_status["sheet"].fillna("").astype(str).str.strip().isin(hide)
                ].copy()

    chart_input = collapse_sheets_by_sector_name_for_charts(
        comparison_long=comp,
        sheet_map=sheet_map,
    )
    chart_input = _drop_parent_child_comparator_overlaps(chart_input, mapping_status)
    chart_input["value"] = pd.to_numeric(chart_input["value"], errors="coerce").abs()
    key_cols = ["sheet", "measure", "fuel_label", "scenario"]
    availability = (
        chart_input.groupby(key_cols, as_index=False)
        .apply(
            lambda g: pd.Series(
                {
                    "has_leap": g.loc[
                        g["source"].isin(["leap"]),
                        "value",
                    ].notna().any(),
                    "has_base": g.loc[
                        g["source"].isin(["base", "base_estimated", "base_mixed"]),
                        "value",
                    ].notna().any(),
                    "has_projection": g.loc[
                        g["source"].isin(["projection", "projection_estimated", "projection_mixed"]),
                        "value",
                    ].notna().any(),
                    "force_show_chart": (
                        bool(g["force_show_chart"].fillna(False).astype(bool).any())
                        if "force_show_chart" in g.columns
                        else False
                    ),
                }
            ),
            include_groups=False,
        )
        .reset_index()
    )
    availability["keep_display"] = (
        availability["has_leap"] | availability["has_base"] | availability["has_projection"] | availability["force_show_chart"]
    )
    display_keys = availability[availability["keep_display"]][key_cols]
    if display_keys.empty:
        return chart_input.iloc[0:0].copy()
    display_rows = chart_input.merge(display_keys, on=key_cols, how="inner")
    total_rows = build_total_rows_for_charts(display_rows, mapping_status)
    # Preserve direct total-only series (for example derived transformation
    # loss/own-use sheets) when no non-total components exist to rebuild them.
    display_non_total = display_rows[display_rows["fuel_label"].astype(str) != "Total"].copy()
    display_total = display_rows[display_rows["fuel_label"].astype(str) == "Total"].copy()
    total_only_rows = display_total.iloc[0:0].copy()
    if not display_total.empty:
        rollup_key_cols = ["economy", "sheet", "measure", "scenario", "source", "year"]
        non_total_keys = display_non_total[rollup_key_cols].drop_duplicates()
        total_with_children = (
            display_total[rollup_key_cols]
            .drop_duplicates()
            .merge(non_total_keys.assign(_has_children=True), on=rollup_key_cols, how="left")
        )
        total_only_keys = total_with_children.loc[
            ~total_with_children["_has_children"].fillna(False).astype(bool),
            rollup_key_cols,
        ].drop_duplicates()
        if not total_only_keys.empty:
            total_only_rows = display_total.merge(total_only_keys, on=rollup_key_cols, how="inner")
            comparator_sources = {"base", "base_estimated", "base_mixed", "projection", "projection_estimated", "projection_mixed"}
            if not total_only_rows.empty:
                non_total_any = (
                    display_non_total[["economy", "sheet", "measure", "scenario", "year"]]
                    .drop_duplicates()
                    .assign(_has_any_non_total=True)
                )
                comparator_non_total = display_non_total[display_non_total["source"].isin(comparator_sources)].copy()
                if "used_in_total" in comparator_non_total.columns:
                    comparator_non_total["used_in_total"] = _coerce_bool_series(
                        comparator_non_total["used_in_total"], default=True
                    )
                    comparator_non_total = comparator_non_total[comparator_non_total["used_in_total"]].copy()
                comparator_non_total = (
                    comparator_non_total[rollup_key_cols]
                    .drop_duplicates()
                    .assign(_has_included_non_total_for_source=True)
                )
                total_only_rows = total_only_rows.merge(
                    non_total_any,
                    on=["economy", "sheet", "measure", "scenario", "year"],
                    how="left",
                )
                total_only_rows = total_only_rows.merge(
                    comparator_non_total,
                    on=rollup_key_cols,
                    how="left",
                )
                drop_total_only_comparator = (
                    total_only_rows["source"].isin(comparator_sources)
                    & total_only_rows["_has_any_non_total"].fillna(False).astype(bool)
                    & ~total_only_rows["_has_included_non_total_for_source"].fillna(False).astype(bool)
                )
                total_only_rows = total_only_rows[~drop_total_only_comparator].drop(
                    columns=["_has_any_non_total", "_has_included_non_total_for_source"],
                    errors="ignore",
                )

    chart_input = pd.concat(
        [
            display_non_total,
            total_rows,
            total_only_rows,
        ],
        ignore_index=True,
        sort=False,
    )
    chart_input = chart_input.drop_duplicates(
        subset=["economy", "sheet", "measure", "fuel_label", "scenario", "source", "year"],
        keep="first",
    )
    chart_input = _dedupe_blank_measure_chart_rows(chart_input)
    chart_input = _dashboard_utils._collapse_base_family_rows_for_display(chart_input)
    chart_input = _dashboard_utils._collapse_projection_family_rows_for_display(chart_input)
    chart_input["value"] = pd.to_numeric(chart_input["value"], errors="coerce").abs()
    if "used_in_total" not in chart_input.columns:
        chart_input["used_in_total"] = True
    chart_input["used_in_total"] = _coerce_bool_series(chart_input["used_in_total"], default=True)
    return chart_input


def _dedupe_blank_measure_chart_rows(chart_input: pd.DataFrame) -> pd.DataFrame:
    """
    Remove exact duplicates where only the measure label differs by blankness.

    Prefer the explicit measure label over the blank one so one-to-one alias
    sheets do not render duplicate charts like the bunker totals.
    """
    if chart_input.empty or "measure" not in chart_input.columns:
        return chart_input

    out = chart_input.copy()
    out["measure"] = out["measure"].fillna("").astype(str)
    out["value"] = pd.to_numeric(out["value"], errors="coerce")

    # Pass 1: when the same line key exists in both blank/non-blank measures,
    # keep only non-blank rows to avoid parallel blank-measure chart series.
    line_key = ["economy", "sheet", "fuel_label", "scenario", "source", "year"]
    measure_presence = (
        out.assign(_is_blank_measure=out["measure"].str.strip().eq(""))
        .groupby(line_key, as_index=False)["_is_blank_measure"]
        .agg(has_blank="any")
    )
    measure_presence["has_non_blank"] = ~measure_presence["has_blank"]
    non_blank_presence = (
        out.assign(_is_non_blank_measure=out["measure"].str.strip().ne(""))
        .groupby(line_key, as_index=False)["_is_non_blank_measure"]
        .agg(has_non_blank="any")
    )
    measure_presence = measure_presence.drop(columns=["has_non_blank"]).merge(
        non_blank_presence,
        on=line_key,
        how="left",
    )
    out = out.merge(measure_presence, on=line_key, how="left")
    drop_blank = (
        out["has_blank"].fillna(False).astype(bool)
        & out["has_non_blank"].fillna(False).astype(bool)
        & out["measure"].str.strip().eq("")
    )
    out = out.loc[~drop_blank].drop(columns=["has_blank", "has_non_blank"], errors="ignore")

    # Pass 2: if a blank/non-blank duplicate still exists with identical value,
    # keep the non-blank row.
    key_cols = ["economy", "sheet", "fuel_label", "scenario", "source", "year", "value"]
    duplicate_mask = out.duplicated(subset=key_cols, keep=False)
    if duplicate_mask.any():
        candidates = out.loc[duplicate_mask].copy()
        candidates["_measure_rank"] = candidates["measure"].str.strip().ne("").astype(int)
        preferred = (
            candidates.sort_values(key_cols + ["_measure_rank"], ascending=[True] * len(key_cols) + [False])
            .drop_duplicates(subset=key_cols, keep="first")
            .drop(columns="_measure_rank")
        )
        non_candidates = out.loc[~duplicate_mask].copy()
        out = pd.concat([non_candidates, preferred], ignore_index=True, sort=False)

    return out.sort_values(
        ["economy", "sheet", "measure", "fuel_label", "scenario", "source", "year"],
        kind="mergesort",
    ).reset_index(drop=True)


def _truncate_to_max_year(df: pd.DataFrame, *, max_year: int = MAX_OUTPUT_YEAR) -> pd.DataFrame:
    if df.empty or "year" not in df.columns:
        return df

    out = df.copy()
    numeric_year = pd.to_numeric(out["year"], errors="coerce")
    return out.loc[numeric_year.isna() | (numeric_year <= max_year)].copy()


def _sheet_rows_equivalent(
    df: pd.DataFrame,
    *,
    left_sheet: str,
    right_sheet: str,
) -> bool:
    """
    Return True when two sheet labels carry the exact same comparison series.

    We only collapse alias labels when this strict equivalence check passes.
    """
    compare_cols = [
        "economy",
        "fuel_label",
        "scenario",
        "source",
        "year",
        "value",
    ]
    available = [col for col in compare_cols if col in df.columns]
    if not available:
        return False

    left = df[df["sheet"].astype(str) == str(left_sheet)].copy()
    right = df[df["sheet"].astype(str) == str(right_sheet)].copy()
    if left.empty or right.empty:
        return False

    def _measure_norm(series: pd.Series) -> pd.Series:
        out = series.fillna("").astype(str).str.strip().str.lower()
        out = out.str.replace("petajoules", "pj", regex=False)
        out = out.str.replace(r"\s+", " ", regex=True)
        return out

    for frame in (left, right):
        for col in available:
            if col == "value":
                frame[col] = pd.to_numeric(frame[col], errors="coerce").round(12)
            else:
                frame[col] = frame[col].fillna("").astype(str).str.strip()

    if "measure" in df.columns:
        left_measure = _measure_norm(left.get("measure", pd.Series(dtype="object")))
        right_measure = _measure_norm(right.get("measure", pd.Series(dtype="object")))
        left_non_blank = {m for m in left_measure.tolist() if m}
        right_non_blank = {m for m in right_measure.tolist() if m}
        if len(left_non_blank) > 1 or len(right_non_blank) > 1:
            return False
        if left_non_blank and right_non_blank and left_non_blank != right_non_blank:
            return False

    left_norm = (
        left[available]
        .drop_duplicates()
        .sort_values(available, kind="mergesort")
        .reset_index(drop=True)
    )
    right_norm = (
        right[available]
        .drop_duplicates()
        .sort_values(available, kind="mergesort")
        .reset_index(drop=True)
    )
    return left_norm.equals(right_norm)


def _collapse_sheet_alias_rows(
    comparison_long: pd.DataFrame,
    sheet_map: pd.DataFrame,
) -> pd.DataFrame:
    """
    Collapse duplicated sheet aliases (sector-name vs final-category labels).

    We map `sector_name -> final_category_name` only when both labels are
    present and their full series are identical, so this removes true naming
    duplicates (for example bunker alias pairs) without collapsing distinct data.
    """
    if comparison_long.empty or sheet_map.empty or "sheet" not in comparison_long.columns:
        return comparison_long

    sm = sheet_map.copy()
    for col in ["sheet_name", "sector_code_9th", "sector_name", "final_category_name", "active"]:
        if col not in sm.columns:
            sm[col] = ""
        sm[col] = sm[col].fillna("").astype(str).str.strip()
    sm = sm[sm["sheet_name"].ne("")].copy()
    if sm.empty:
        return comparison_long
    sm = sm[sm["active"].str.lower().isin({"true", "1", "yes", "y"})].copy()
    if sm.empty:
        return comparison_long

    sm["canonical_name"] = sm["final_category_name"].where(
        sm["final_category_name"].ne(""),
        sm["sheet_name"],
    )
    sm["alias_name"] = sm["sector_name"]
    sm = sm[
        sm["canonical_name"].ne("")
        & sm["alias_name"].ne("")
        & sm["alias_name"].ne(sm["canonical_name"])
    ].copy()
    if sm.empty:
        return comparison_long

    alias_candidates: dict[str, set[str]] = {}
    for _, row in sm.iterrows():
        alias_candidates.setdefault(str(row["alias_name"]), set()).add(str(row["canonical_name"]))
    alias_map = {
        alias: next(iter(canonicals))
        for alias, canonicals in alias_candidates.items()
        if len(canonicals) == 1
    }
    if not alias_map:
        return comparison_long

    out = comparison_long.copy()
    out["sheet"] = out["sheet"].fillna("").astype(str).str.strip()
    present = set(out["sheet"].tolist())
    effective_map = {
        alias: canonical
        for alias, canonical in alias_map.items()
        if alias in present and canonical in present
    }
    if not effective_map:
        return comparison_long

    safe_map = {
        alias: canonical
        for alias, canonical in effective_map.items()
        if _sheet_rows_equivalent(out, left_sheet=alias, right_sheet=canonical)
    }
    if not safe_map:
        return comparison_long

    out["sheet"] = out["sheet"].replace(safe_map)
    dedupe_cols = [c for c in ["economy", "sheet", "measure", "fuel_label", "scenario", "source", "year", "value"] if c in out.columns]
    if dedupe_cols:
        out = out.drop_duplicates(subset=dedupe_cols, keep="first")
    out = _dedupe_blank_measure_chart_rows(out)
    return out


def _assert_not_used_comparators_excluded_from_totals(chart_input: pd.DataFrame) -> None:
    if chart_input.empty:
        return
    required = {"sheet", "measure", "fuel_label", "scenario", "source", "year", "used_in_total"}
    if not required.issubset(set(chart_input.columns)):
        return
    comparator_sources = {"base", "base_estimated", "base_mixed", "projection", "projection_estimated", "projection_mixed"}
    frame = chart_input.copy()
    frame["measure"] = frame["measure"].fillna("").astype(str)
    frame["fuel_label"] = frame["fuel_label"].fillna("").astype(str)
    frame["source"] = frame["source"].fillna("").astype(str)
    frame["scenario"] = frame["scenario"].fillna("").astype(str)
    frame["year"] = pd.to_numeric(frame["year"], errors="coerce").astype("Int64")
    frame["used_in_total"] = _coerce_bool_series(frame["used_in_total"], default=True)

    non_total_comp = frame[
        frame["source"].isin(comparator_sources) & frame["fuel_label"].ne("Total")
    ].copy()
    if non_total_comp.empty:
        return
    scope_key = ["sheet", "measure", "scenario", "year"]
    scope = (
        non_total_comp.groupby(scope_key, as_index=False)
        .agg(
            has_non_total=("fuel_label", "size"),
            has_included=("used_in_total", "any"),
        )
    )
    fully_excluded = scope[(scope["has_non_total"] > 0) & (~scope["has_included"])][scope_key]
    if fully_excluded.empty:
        return

    total_comp = frame[
        frame["source"].isin(comparator_sources) & frame["fuel_label"].eq("Total")
    ][scope_key + ["source"]].drop_duplicates()
    leaks = fully_excluded.merge(total_comp, on=scope_key, how="inner")
    if leaks.empty:
        return
    sample = leaks.head(12).to_dict("records")
    raise RuntimeError(
        "Invalid chart input: comparator lines marked not used in total leaked into comparator Total rows. "
        f"Total leaks: {len(leaks)}. Examples: {sample}"
    )


def _replace_directory_contents(staged_dir: Path, live_dir: Path) -> None:
    """Publish a staged directory over the live one after a successful render."""
    if not staged_dir.exists():
        print(f"[WARN] Skipping publish; staged directory does not exist: {staged_dir}", flush=True)
        return
    backup_dir = live_dir.with_name(f"{live_dir.name}__old")
    retry_delays = (0.0, 0.2, 0.5, 1.0, 2.0)
    last_permission_error: PermissionError | None = None

    for attempt, delay_s in enumerate(retry_delays, start=1):
        if delay_s > 0:
            time.sleep(delay_s)
        if backup_dir.exists():
            shutil.rmtree(backup_dir, ignore_errors=True)
        try:
            if live_dir.exists():
                live_dir.replace(backup_dir)
            staged_dir.replace(live_dir)
        except PermissionError as exc:
            last_permission_error = exc
            if not live_dir.exists() and backup_dir.exists():
                try:
                    backup_dir.replace(live_dir)
                except Exception:
                    pass
            if attempt < len(retry_delays):
                print(
                    f"[WARN] Publish retry {attempt}/{len(retry_delays)} for {live_dir.name} "
                    f"after permission error: {exc}",
                    flush=True,
                )
                continue
            break
        except Exception:
            if live_dir.exists():
                shutil.rmtree(live_dir, ignore_errors=True)
            if backup_dir.exists():
                backup_dir.replace(live_dir)
            raise
        else:
            if backup_dir.exists():
                shutil.rmtree(backup_dir, ignore_errors=True)
            return

    if last_permission_error is not None:
        raise PermissionError(
            f"Failed to publish staged directory '{staged_dir}' to '{live_dir}'. "
            "A process is likely holding a lock (browser preview, Explorer, antivirus, sync client)."
        ) from last_permission_error


def run_workflow() -> dict[str, object]:
    ensure_repo_root()
    out_dir = resolve_path(OUTPUT_DIR)
    layout = build_workflow_output_layout(out_dir)
    coverage_result: dict[str, object] | None = None
    if RUN_MAPPING_COVERAGE_CHECK:
        coverage_result = run_mapping_coverage_check(
            mapping_path=resolve_path(NINTH_TO_ESTO_PATH),
            esto_data_path=resolve_path(BASE_TABLE_PATH),
            ninth_data_path=resolve_path(PROJECTION_TABLE_PATH),
            output_dir=layout.coverage / "mapping_coverage",
            base_year=BASE_YEAR,
            projection_years=tuple(PROJECTION_YEARS),
            scenario="reference",
        )
        coverage_summary = coverage_result["summary"]
        print(
            "Mapping coverage check: "
            f"missing ESTO pairs={coverage_summary['missing_esto_pairs']}, "
            f"missing 9th pairs={coverage_summary['missing_ninth_pairs']}",
            flush=True,
        )
    else:
        print("[INFO] Skipping mapping coverage check in fast mode", flush=True)
    leap_refresh_log: dict[str, object] | None = None

    if REFRESH_LEAP_RESULTS_BEFORE_DASHBOARD:
        leap_refresh_log = _refresh_leap_results_inputs()

    mapping_inputs = load_mapping_inputs(
        sheet_map_path=resolve_path(SHEET_MAP_PATH),
        backup_mappings_path=resolve_path(BACKUP_MAPPINGS_PATH),
        codebook_path=resolve_path(CODEBOOK_PATH),
        canonical_pairs_path=resolve_path(NINTH_TO_ESTO_PATH),
        explicit_mappings_path=resolve_path(EXPLICIT_MAPPINGS_PATH),
        explicit_reassignments_path=resolve_path(EXPLICIT_REASSIGNMENTS_PATH),
    )
    print("[INFO] Loaded mapping inputs", flush=True)
    write_canonical_conflicts(mapping_inputs["canonical_conflicts"], resolve_path(MAPPING_VIEWS_DIR))

    workbooks = discover_workbooks(resolve_path(LEAP_RESULTS_DIR), ECONOMY_TOKEN, tuple(SCENARIOS))
    leap_long = load_leap_long(
        workbooks,
        mapping_inputs["sheet_map"],
        additional_long_paths=None,
    )
    if leap_long.empty:
        raise RuntimeError("No LEAP data loaded; check workbook paths.")
    print("[INFO] Loaded LEAP workbooks", flush=True)
    derived_leap_long, derived_leap_audit, derived_summary = build_derived_transformation_leap_long(
        leap_long,
        sheet_map=mapping_inputs["sheet_map"],
        data_dir=resolve_path(LEAP_RESULTS_DIR),
        economy_token=ECONOMY_TOKEN,
        derivation_mode=DERIVED_TRANSFORMATION_METHOD,
    )
    leap_long_for_output = pd.concat([leap_long, derived_leap_long], ignore_index=True, sort=False)
    leap_input_addback_report = pd.DataFrame()
    if TRANSFORMATION_DISPLAY_MODE != "separate":
        leap_long_for_output, leap_input_addback_report = _add_loss_own_use_to_transformation_inputs_in_leap_long(
            leap_long_for_output,
            sheet_map=mapping_inputs["sheet_map"],
        )

    archive_config_dir_once_per_day()
    esto_df, ninth_df, reassignment_status, synthetic_reference_status = load_reference_tables(
        esto_table_path=resolve_path(BASE_TABLE_PATH),
        projection_table_path=resolve_path(PROJECTION_TABLE_PATH),
        explicit_reassignments=mapping_inputs["explicit_reassignments"],
        explicit_mappings=mapping_inputs["explicit_mappings"],
        canonical_pairs=mapping_inputs["canonical_pairs"],
        synthetic_reference_rows_path=resolve_path(SYNTHETIC_REFERENCE_ROWS_PATH),
        drop_all_zero_base_rows=DROP_ALL_ZERO_BASE_ROWS,
        drop_all_zero_projection_rows=DROP_ALL_ZERO_PROJECTION_ROWS,
    )
    if not reassignment_status.empty:
        reassignment_status.to_csv(layout.mapping / "explicit_reassignment_status.csv", index=False)
    if not synthetic_reference_status.empty:
        synthetic_reference_status.to_csv(layout.mapping / "synthetic_reference_row_status.csv", index=False)
    synthetic_esto_mask = _synthetic_row_mask(esto_df, "_synthetic_esto_row")
    synthetic_ninth_mask = _synthetic_row_mask(ninth_df, "_synthetic_ninth_row")
    synthetic_esto_rows = esto_df.loc[synthetic_esto_mask].copy()
    synthetic_ninth_rows = ninth_df.loc[synthetic_ninth_mask].copy()
    if not synthetic_esto_rows.empty or not synthetic_ninth_rows.empty:
        synthetic_workbook = layout.mapping / "synthetic_reference_rows.xlsx"
        drop_cols = ["_synthetic_esto_row", "_synthetic_ninth_row", "_synthetic_rule_name"]
        with pd.ExcelWriter(synthetic_workbook) as writer:
            synthetic_esto_rows.drop(columns=drop_cols, errors="ignore").to_excel(
                writer,
                sheet_name="ESTO Synthetic Rows",
                index=False,
            )
            synthetic_ninth_rows.drop(columns=drop_cols, errors="ignore").to_excel(
                writer,
                sheet_name="9th Synthetic Rows",
                index=False,
            )
    print("[INFO] Loaded reference tables", flush=True)

    print("[INFO] Building comparisons...", flush=True)
    comparison_long, comparison_wide, mapping_status = build_comparisons_v2(
        leap_long=leap_long,
        sheet_map=mapping_inputs["sheet_map"],
        fuel_mapping=mapping_inputs["fuel_aliases"],
        sector_flow_mapping=mapping_inputs["sector_flow_mapping"],
        ninth_pairs=mapping_inputs["canonical_pairs"],
        base_df=esto_df,
        ninth_df=ninth_df,
        explicit_mappings=mapping_inputs["explicit_mappings"],
        base_year=BASE_YEAR,
        base_economy=BASE_ECONOMY,
        projection_economy=PROJECTION_ECONOMY,
        projection_years=tuple(PROJECTION_YEARS),
        scenario_map=SCENARIO_MAP,
        use_esto_agg_only=USE_ESTO_AGG_ONLY,
        sibling_comparator_mode=SIBLING_COMPARATOR_MODE,
        include_sibling_parent_totals=INCLUDE_SIBLING_PARENT_TOTALS,
        settings=V2_SETTINGS,
    )
    print("[INFO] Comparison build complete", flush=True)
    derived_comparison_long, derived_mapping_status, derived_comparison_audit = build_derived_transformation_comparison_rows(
        comparison_long,
        mapping_status,
        sheet_map=mapping_inputs["sheet_map"],
        data_dir=resolve_path(LEAP_RESULTS_DIR),
        economy_token=ECONOMY_TOKEN,
        base_df=esto_df,
        ninth_df=ninth_df,
        base_year=BASE_YEAR,
        projection_years=tuple(PROJECTION_YEARS),
        derivation_mode=DERIVED_TRANSFORMATION_METHOD,
        derived_leap_long=derived_leap_long,
    )
    if not derived_comparison_long.empty:
        comparison_long = pd.concat([comparison_long, derived_comparison_long], ignore_index=True, sort=False)
    if not derived_mapping_status.empty:
        mapping_status = pd.concat([mapping_status, derived_mapping_status], ignore_index=True, sort=False)
    comparison_input_addback_report = pd.DataFrame()
    if TRANSFORMATION_DISPLAY_MODE != "separate":
        comparison_long, comparison_input_addback_report = _add_loss_own_use_to_transformation_inputs_in_comparison(
            comparison_long,
            sheet_map=mapping_inputs["sheet_map"],
        )
    if not comparison_input_addback_report.empty:
        try:
            comparison_input_addback_report.to_csv(
                layout.derived / "transformation_input_addback_report.csv",
                index=False,
            )
        except PermissionError as exc:
            print(f"[WARN] Skipping write to locked input-addback report: {exc}", flush=True)
    if not leap_input_addback_report.empty:
        try:
            leap_input_addback_report.to_csv(
                layout.derived / "transformation_input_addback_leap_report.csv",
                index=False,
            )
        except PermissionError as exc:
            print(f"[WARN] Skipping write to locked LEAP input-addback report: {exc}", flush=True)
    if TRANSFORMATION_DISPLAY_MODE != "separate":
        comparison_long, mapping_status, leap_long_for_output = _drop_loss_own_use_sheets(
            comparison_long,
            mapping_status,
            leap_long_for_output,
            keep_loss_sheets=COMBINED_MODE_KEEP_LOSS_SHEETS,
        )
        comparison_long, mapping_status, leap_long_for_output = _retitle_combined_transformation_input_measures(
            comparison_long,
            mapping_status,
            leap_long_for_output,
        )
        comparison_long, mapping_status, leap_long_for_output = _retitle_trans_dist_loss_measure(
            comparison_long,
            mapping_status,
            leap_long_for_output,
        )
    comparison_long, mapping_status = add_parent_totals_from_dashboard_overrides(
        comparison_long,
        mapping_status,
        sheet_map=mapping_inputs["sheet_map"],
    )
    try:
        sm = mapping_inputs["sheet_map"].copy()
        for col in ["sheet_name", "active"]:
            if col not in sm.columns:
                sm[col] = ""
            sm[col] = sm[col].fillna("").astype(str).str.strip()
        if "active" in sm.columns:
            sm = sm[sm["active"].str.lower().isin({"true", "1", "yes", "y"})].copy()
        allowed = {str(name).strip() for name in sm["sheet_name"].tolist() if str(name).strip()}
        comp_sheets = set(comparison_long["sheet"].fillna("").astype(str).str.strip().tolist())
        missing = sorted(sheet for sheet in comp_sheets if sheet and sheet not in allowed)
        pd.DataFrame({"sheet": missing}).to_csv(layout.mapping / "sheets_not_in_map_before_filter.csv", index=False)
        if "mapping_source" in mapping_status.columns:
            derived = mapping_status[
                mapping_status["mapping_source"].fillna("").astype(str).str.strip().eq("derived_parent_total")
            ]["sheet"].dropna().astype(str)
            derived_missing = sorted(sheet for sheet in set(derived) if sheet not in allowed)
            pd.DataFrame({"sheet": derived_missing}).to_csv(
                layout.mapping / "derived_parent_sheets_not_in_map_before_filter.csv", index=False
            )
    except PermissionError as exc:
        print(f"[WARN] Skipping write to locked sheet-map report: {exc}", flush=True)
    if FILTER_TO_SHEET_MAP:
        comparison_long, mapping_status = _filter_to_sheet_map(
            comparison_long,
            mapping_status,
            mapping_inputs["sheet_map"],
            out_dir=layout.mapping,
        )
    comparison_long = _collapse_sheet_alias_rows(comparison_long, mapping_inputs["sheet_map"])
    comparison_wide = _rebuild_comparison_wide(comparison_long)
    try:
        derived_artifacts = write_derived_transformation_artifacts(
            data_dir=resolve_path(LEAP_RESULTS_DIR),
            out_dir=layout.derived,
            economy_token=ECONOMY_TOKEN,
            derived_leap_long=derived_leap_long,
            leap_audit=derived_leap_audit,
            comparison_audit=derived_comparison_audit,
            summary=derived_summary,
        )
    except PermissionError as exc:
        print(
            f"[WARN] Skipping write to locked derived-transformation artifact path: {exc}",
            flush=True,
        )
        derived_artifacts = {}
    leap_long = _truncate_to_max_year(leap_long_for_output)
    comparison_long = _truncate_to_max_year(comparison_long)
    comparison_wide = _truncate_to_max_year(comparison_wide)

    atomic_outputs: dict[str, pd.DataFrame] = {}
    atomic_many_to_many_errors_path: str | None = None
    atomic_shadow_delta_series_path: str | None = None
    atomic_shadow_delta_totals_path: str | None = None
    atomic_shadow_delta_summary_path: str | None = None
    atomic_chart_input = pd.DataFrame()
    if ATOMIC_SETTINGS.enabled:
        print("[INFO] Building atomic outputs...", flush=True)
        atomic_outputs = build_atomic_outputs(
            comparison_long=comparison_long,
            mapping_status=mapping_status,
            sheet_map=mapping_inputs["sheet_map"],
            canonical_pairs=mapping_inputs["canonical_pairs"],
            base_df=esto_df,
            ninth_df=ninth_df,
            leap_long=leap_long,
            base_economy=BASE_ECONOMY,
            projection_economy=PROJECTION_ECONOMY,
            settings=ATOMIC_SETTINGS,
        )
        print("[INFO] Atomic output build complete", flush=True)
        atomic_derived_comparison, _, _ = build_derived_transformation_comparison_rows(
            atomic_outputs.get("atomic_comparison_long", pd.DataFrame()),
            mapping_status,
            sheet_map=mapping_inputs["sheet_map"],
            data_dir=resolve_path(LEAP_RESULTS_DIR),
            economy_token=ECONOMY_TOKEN,
            base_df=esto_df,
            ninth_df=ninth_df,
            base_year=BASE_YEAR,
            projection_years=tuple(PROJECTION_YEARS),
            derivation_mode=DERIVED_TRANSFORMATION_METHOD,
            derived_leap_long=derived_leap_long,
        )
        if not atomic_derived_comparison.empty and "atomic_comparison_long" in atomic_outputs:
            atomic_outputs["atomic_comparison_long"] = pd.concat(
                [atomic_outputs["atomic_comparison_long"], atomic_derived_comparison],
                ignore_index=True,
                sort=False,
            )
            if TRANSFORMATION_DISPLAY_MODE != "separate":
                atomic_outputs["atomic_comparison_long"], _, _ = _drop_loss_own_use_sheets(
                    atomic_outputs["atomic_comparison_long"],
                    mapping_status,
                    None,
                    keep_loss_sheets=COMBINED_MODE_KEEP_LOSS_SHEETS,
                )
                atomic_outputs["atomic_comparison_long"], _, _ = _retitle_combined_transformation_input_measures(
                    atomic_outputs["atomic_comparison_long"],
                    mapping_status,
                    None,
                )
                atomic_outputs["atomic_comparison_long"], _, _ = _retitle_trans_dist_loss_measure(
                    atomic_outputs["atomic_comparison_long"],
                    mapping_status,
                    None,
                )
            atomic_outputs["atomic_comparison_long"] = _collapse_sheet_alias_rows(
                atomic_outputs["atomic_comparison_long"],
                mapping_inputs["sheet_map"],
            )
            atomic_outputs["atomic_comparison_wide"] = _rebuild_comparison_wide(atomic_outputs["atomic_comparison_long"])
        many_to_many_errors = atomic_outputs.get("atomic_many_to_many_errors", pd.DataFrame())
        if not many_to_many_errors.empty:
            mm_path = layout.atomic / "atomic_many_to_many_errors.csv"
            many_to_many_errors.to_csv(mm_path, index=False)
            atomic_many_to_many_errors_path = str(mm_path)
            if str(ATOMIC_SETTINGS.many_to_many_policy).strip().lower() == "error":
                sample = many_to_many_errors.head(10).to_dict("records")
                raise RuntimeError(
                    "Atomic validation failed: unresolved many-to-many mapping components detected. "
                    f"Total components: {len(many_to_many_errors)}. Examples: {sample}"
                )
        atomic_chart_input = _build_chart_input_for_rendering(
            comparison_long=atomic_outputs.get("atomic_comparison_long", pd.DataFrame()),
            mapping_status=mapping_status,
            sheet_map=mapping_inputs["sheet_map"],
        )
        atomic_chart_input = _truncate_to_max_year(atomic_chart_input)
        for key in ("atomic_comparison_long", "atomic_comparison_wide"):
            if key in atomic_outputs:
                atomic_outputs[key] = _truncate_to_max_year(atomic_outputs[key])

    print("[INFO] Writing core outputs...", flush=True)
    output_paths = write_core_outputs(
        out_dir=out_dir,
        supporting_dir=layout.supporting,
        comparison_long=comparison_long,
        comparison_wide=comparison_wide,
        mapping_status=mapping_status,
        leap_long=leap_long,
        atomic_comparison_long=(
            atomic_outputs.get("atomic_comparison_long")
            if (ATOMIC_SETTINGS.enabled and ATOMIC_SETTINGS.write_shadow_outputs)
            else None
        ),
        atomic_comparison_wide=(
            atomic_outputs.get("atomic_comparison_wide")
            if (ATOMIC_SETTINGS.enabled and ATOMIC_SETTINGS.write_shadow_outputs)
            else None
        ),
        atomic_mapping_edges=(
            atomic_outputs.get("atomic_mapping_edges")
            if (ATOMIC_SETTINGS.enabled and ATOMIC_SETTINGS.write_shadow_outputs)
            else None
        ),
        atomic_validation_report=(
            atomic_outputs.get("atomic_validation_report")
            if (ATOMIC_SETTINGS.enabled and ATOMIC_SETTINGS.write_shadow_outputs)
            else None
        ),
    )
    print("[INFO] Core outputs written", flush=True)

    diagnostics_artifacts: dict[str, str] = {}
    if WRITE_DIAGNOSTIC_ARTIFACTS:
        print("[INFO] Writing diagnostics...", flush=True)
        diagnostics_artifacts = write_diagnostics(
            comparison_long=comparison_long,
            mapping_status=mapping_status,
            out_dir=layout.root,
            base_year=BASE_YEAR,
            diagnostic_probe_year=DIAGNOSTIC_PROBE_YEAR,
            top_diagnostic_rows=TOP_DIAGNOSTIC_ROWS,
        )
        print("[INFO] Diagnostics complete", flush=True)

    charts_dir = layout.charts
    dashboards_dir = layout.dashboards
    rollout_mode = str(ATOMIC_SETTINGS.rollout_mode).strip().lower()
    is_atomic_rollout = ATOMIC_SETTINGS.enabled and rollout_mode == "atomic"
    needs_legacy_chart_input = (
        (not is_atomic_rollout)
        or (ATOMIC_SETTINGS.enabled and ATOMIC_SETTINGS.write_shadow_outputs)
    )
    legacy_chart_input = pd.DataFrame()
    child_context_comparison_long = pd.DataFrame()
    if needs_legacy_chart_input:
        if SHOW_PARENT_SCOPED_CHILD_CONTEXT_COMPARATORS and str(SIBLING_COMPARATOR_MODE).strip().lower() == "aggregate_to_parent":
            print("[INFO] Building child-level context comparators (excluded from totals)...", flush=True)
            child_context_comparison_long, _, _ = build_comparisons_v2(
                leap_long=leap_long,
                sheet_map=mapping_inputs["sheet_map"],
                fuel_mapping=mapping_inputs["fuel_aliases"],
                sector_flow_mapping=mapping_inputs["sector_flow_mapping"],
                ninth_pairs=mapping_inputs["canonical_pairs"],
                base_df=esto_df,
                ninth_df=ninth_df,
                explicit_mappings=mapping_inputs["explicit_mappings"],
                base_year=BASE_YEAR,
                base_economy=BASE_ECONOMY,
                projection_economy=PROJECTION_ECONOMY,
                projection_years=tuple(PROJECTION_YEARS),
                scenario_map=SCENARIO_MAP,
                use_esto_agg_only=USE_ESTO_AGG_ONLY,
                sibling_comparator_mode="allocate_by_leap_share",
                include_sibling_parent_totals=INCLUDE_SIBLING_PARENT_TOTALS,
                settings=V2_SETTINGS,
            )
            child_context_comparison_long = _collapse_sheet_alias_rows(child_context_comparison_long, mapping_inputs["sheet_map"])
            child_context_comparison_long = _truncate_to_max_year(child_context_comparison_long)
            print("[INFO] Child-level context comparators ready", flush=True)
        legacy_chart_input = _build_chart_input_for_rendering(
            comparison_long=comparison_long,
            mapping_status=mapping_status,
            sheet_map=mapping_inputs["sheet_map"],
            child_context_comparison_long=child_context_comparison_long,
        )
        legacy_chart_input = _truncate_to_max_year(legacy_chart_input)

    if is_atomic_rollout and not atomic_chart_input.empty:
        chart_input = atomic_chart_input.copy()
    else:
        if legacy_chart_input.empty:
            legacy_chart_input = _build_chart_input_for_rendering(
                comparison_long=comparison_long,
                mapping_status=mapping_status,
                sheet_map=mapping_inputs["sheet_map"],
                child_context_comparison_long=child_context_comparison_long,
            )
            legacy_chart_input = _truncate_to_max_year(legacy_chart_input)
        chart_input = legacy_chart_input.copy()

    if (
        ATOMIC_SETTINGS.enabled
        and ATOMIC_SETTINGS.write_shadow_outputs
        and not atomic_chart_input.empty
        and not legacy_chart_input.empty
    ):
        delta_frames = build_shadow_delta_reports(
            legacy_chart_input=legacy_chart_input,
            atomic_chart_input=atomic_chart_input,
        )
        series_path = layout.atomic / "atomic_shadow_delta_series.csv"
        totals_path = layout.atomic / "atomic_shadow_delta_totals.csv"
        summary_path = layout.atomic / "atomic_shadow_delta_summary.csv"
        delta_frames["atomic_shadow_delta_series"].to_csv(series_path, index=False)
        delta_frames["atomic_shadow_delta_totals"].to_csv(totals_path, index=False)
        delta_frames["atomic_shadow_delta_summary"].to_csv(summary_path, index=False)
        atomic_shadow_delta_series_path = str(series_path)
        atomic_shadow_delta_totals_path = str(totals_path)
        atomic_shadow_delta_summary_path = str(summary_path)

    latest_snapshot_path: Path | None = None
    daily_archive_path: Path | None = None
    _assert_not_used_comparators_excluded_from_totals(chart_input)
    if WRITE_CHART_SNAPSHOTS:
        # Persist exact chart-rendered series and keep a compressed daily archive.
        snapshot_cols = ["economy", "sheet", "measure", "fuel_label", "scenario", "source", "year", "value"]
        chart_snapshot = chart_input.copy()
        for col in snapshot_cols:
            if col not in chart_snapshot.columns:
                chart_snapshot[col] = pd.NA
        chart_snapshot = chart_snapshot[snapshot_cols].sort_values(snapshot_cols[:-1]).reset_index(drop=True)
        snapshot_dir = layout.snapshots / "chart_series_snapshots"
        archive_dir = snapshot_dir / "daily_archives"
        snapshot_dir.mkdir(parents=True, exist_ok=True)
        archive_dir.mkdir(parents=True, exist_ok=True)
        latest_snapshot_path = snapshot_dir / "chart_series_latest.csv"
        chart_snapshot.to_csv(latest_snapshot_path, index=False)

        # One compressed archive per day to reduce space while preserving history.
        now_utc = pd.Timestamp.now(tz="UTC")
        today_stamp = now_utc.strftime("%Y%m%d")
        daily_archive_path = archive_dir / f"chart_series_{today_stamp}.csv.gz"
        with latest_snapshot_path.open("rb") as src, gzip.open(daily_archive_path, "wb") as dst:
            shutil.copyfileobj(src, dst)

        # Prune old archives by retention window.
        cutoff_date = (now_utc - pd.Timedelta(days=CHART_SNAPSHOT_RETENTION_DAYS)).date()
        for old in archive_dir.glob("chart_series_*.csv.gz"):
            stem = old.stem  # chart_series_YYYYMMDD.csv
            token = stem.replace("chart_series_", "").replace(".csv", "")
            try:
                day = pd.to_datetime(token, format="%Y%m%d", errors="raise")
            except Exception:
                continue
            if day.date() < cutoff_date:
                try:
                    old.unlink()
                except OSError:
                    pass

    chart_line_mapping_path: Path | None = None
    total_component_path: Path | None = None
    chart_hierarchy_json_path: Path | None = None
    chart_hierarchy_csv_path: Path | None = None
    if WRITE_CHART_LEDGERS:
        chart_line_mapping_ledger = build_chart_line_mapping_ledger(chart_input, mapping_status)
        chart_line_mapping_path = layout.ledgers / "chart_line_mapping_ledger.csv"
        chart_line_mapping_ledger.to_csv(chart_line_mapping_path, index=False)
        total_component_ledger = build_total_component_ledger(chart_input, mapping_status)
        total_component_path = layout.ledgers / "chart_total_component_ledger.csv"
        total_component_ledger.to_csv(total_component_path, index=False)
    chart_hierarchy_json_path, chart_hierarchy_csv_path = _write_chart_navigation_hierarchy(
        chart_input=chart_input,
        mapping_status=mapping_status,
        sheet_map=mapping_inputs["sheet_map"],
        out_dir=layout.navigation,
        charts_dir=layout.charts,
        backend=CHART_BACKEND,
        hide_leap_only_charts=HIDE_LEAP_ONLY_CHARTS,
    )
    written_charts: list[Path] = []
    dashboard_index: Path | None = None
    render_stage_root = layout.supporting / "_render_stage"
    stage_output_dir = render_stage_root / "site"
    stage_charts_dir = stage_output_dir / "charts"
    stage_dashboards_dir = stage_output_dir / "dashboards"
    if render_stage_root.exists():
        shutil.rmtree(render_stage_root, ignore_errors=True)
    if GENERATE_CHARTS or GENERATE_DASHBOARDS:
        render_stage_root.mkdir(parents=True, exist_ok=True)
    if GENERATE_CHARTS:
        print("[INFO] Rendering charts to staging...", flush=True)
        written_charts = build_charts(
            chart_input,
            charts_dir=stage_charts_dir,
            backend=CHART_BACKEND,
            hide_leap_only_charts=HIDE_LEAP_ONLY_CHARTS,
        )
        print(f"[INFO] Chart staging complete: {len(written_charts)} charts", flush=True)
    if GENERATE_DASHBOARDS:
        dashboard_charts_dir = stage_charts_dir if stage_charts_dir.exists() else layout.charts
        if not dashboard_charts_dir.exists():
            raise RuntimeError("Dashboard render requested but no chart directory is available.")
        print("[INFO] Rendering dashboards to staging...", flush=True)
        dashboard_index = build_dashboards(
            output_dir=stage_output_dir,
            comparison_long=chart_input,
            charts_dir=dashboard_charts_dir,
            mapping_status=mapping_status,
        )
        print("[INFO] Dashboard staging complete", flush=True)
    if GENERATE_CHARTS:
        _replace_directory_contents(stage_charts_dir, charts_dir)
        written_charts = sorted(charts_dir.glob("*.html"))
        print("[INFO] Published staged charts", flush=True)
    if GENERATE_DASHBOARDS:
        _replace_directory_contents(stage_dashboards_dir, dashboards_dir)
        dashboard_index = dashboards_dir / "index.html"
        print("[INFO] Published staged dashboards", flush=True)
    if render_stage_root.exists():
        shutil.rmtree(render_stage_root, ignore_errors=True)

    checks = run_basic_checks(
        mapping_inputs["sheet_map"],
        mapping_inputs["fuel_aliases"],
        comparison_long,
        mapping_status,
    )

    shadow_compare_path: str | None = None
    if RUN_SHADOW_COMPARE:
        shadow_path = compare_outputs(
            v1_output_dir=resolve_path(V1_OUTPUT_DIR),
            v2_output_dir=layout.root,
            out_path=layout.shadow_compare / "shadow_compare_summary.csv",
        )
        shadow_compare_path = str(shadow_path)

    manifest = write_output_manifest(
        out_dir=layout.root,
        primary_outputs={
            "comparison_long": output_paths.get("comparison_long"),
            "comparison_wide": output_paths.get("comparison_wide"),
            "mapping_status": output_paths.get("mapping_status"),
            "leap_long": output_paths.get("leap_long"),
            "dashboards_dir": str(layout.dashboards),
            "charts_dir": str(layout.charts),
        },
        supporting_outputs={
            "mapping_coverage_dir": str(layout.coverage / "mapping_coverage"),
            "gap_diagnostics": diagnostics_artifacts.get("gap_diagnostics"),
            "mapping_rundown_by_sheet": diagnostics_artifacts.get("mapping_rundown_by_sheet"),
            "mapping_rundown_details": diagnostics_artifacts.get("mapping_rundown_details"),
            "comparison_issue_summary": diagnostics_artifacts.get("comparison_issue_summary"),
            "comparison_issue_cause_summary": diagnostics_artifacts.get("comparison_issue_cause_summary"),
            "chart_series_latest": str(latest_snapshot_path) if latest_snapshot_path else None,
            "chart_series_daily_archive": str(daily_archive_path) if daily_archive_path else None,
            "chart_line_mapping_ledger": str(chart_line_mapping_path) if chart_line_mapping_path else None,
            "chart_total_component_ledger": str(total_component_path) if total_component_path else None,
            "chart_navigation_hierarchy": str(chart_hierarchy_json_path) if chart_hierarchy_json_path else None,
            "chart_navigation_hierarchy_flat": str(chart_hierarchy_csv_path) if chart_hierarchy_csv_path else None,
            "atomic_many_to_many_errors": atomic_many_to_many_errors_path,
            "atomic_shadow_delta_series": atomic_shadow_delta_series_path,
            "atomic_shadow_delta_totals": atomic_shadow_delta_totals_path,
            "atomic_shadow_delta_summary": atomic_shadow_delta_summary_path,
            "shadow_compare_summary": shadow_compare_path,
            "derived_transformation_leap_long": derived_artifacts.get("derived_transformation_leap_long"),
            "derived_transformation_leap_audit": derived_artifacts.get("derived_transformation_leap_audit"),
            "derived_transformation_comparison_audit": derived_artifacts.get("derived_transformation_comparison_audit"),
            "derived_transformation_metric_assessment": derived_artifacts.get("derived_transformation_metric_assessment"),
            "synthetic_reference_row_status": (
                str(layout.mapping / "synthetic_reference_row_status.csv")
                if (layout.mapping / "synthetic_reference_row_status.csv").exists()
                else None
            ),
        },
        primary_output_descriptions={
            "comparison_long": "Main V2 long-form comparison table used for dashboard rendering and audits.",
            "comparison_wide": "Wide-form V2 comparison table with one column per source.",
            "mapping_status": "Per-sheet and per-fuel mapping workbook for the V2 comparison.",
            "leap_long": "Normalized LEAP long table after V2 preprocessing.",
            "dashboards_dir": "Rendered V2 dashboard HTML pages.",
            "charts_dir": "Individual chart files used by the V2 dashboards.",
        },
        supporting_output_descriptions={
            "mapping_coverage_dir": "Coverage-check outputs for ninth-to-ESTO mapping completeness.",
            "gap_diagnostics": "Largest LEAP versus comparator gaps at the configured probe years.",
            "mapping_rundown_by_sheet": "Sheet-level summary of V2 mapping completeness.",
            "mapping_rundown_details": "Detailed V2 mapping audit workbook.",
            "comparison_issue_summary": "Prioritized comparison issues with gap metrics and hints.",
            "comparison_issue_cause_summary": "Frequency summary of comparison issue categories.",
            "chart_series_latest": "Latest exact chart-series snapshot used for archive and regression checks.",
            "chart_series_daily_archive": "Compressed daily archive of the chart-series snapshot.",
            "chart_line_mapping_ledger": "Per-chart-line ledger linking visible chart rows to mapping decisions.",
            "chart_total_component_ledger": "Ledger showing how visible total lines were constructed.",
            "chart_navigation_hierarchy": "JSON export of the rendered dashboard hierarchy.",
            "chart_navigation_hierarchy_flat": "Flat CSV version of the rendered dashboard hierarchy.",
            "atomic_many_to_many_errors": "Atomic comparison rows that still violate the many-to-many policy.",
            "atomic_shadow_delta_series": "Series-level differences between legacy and atomic chart inputs.",
            "atomic_shadow_delta_totals": "Total-line differences between legacy and atomic chart inputs.",
            "atomic_shadow_delta_summary": "Summary of legacy versus atomic chart-input deltas.",
            "shadow_compare_summary": "High-level V1 versus V2 output comparison summary.",
            "derived_transformation_leap_long": "Derived transformation metric series written for downstream use.",
            "derived_transformation_leap_audit": "Audit table for derived LEAP-side transformation metrics.",
            "derived_transformation_comparison_audit": "Audit table for derived comparator-side transformation metrics.",
            "derived_transformation_metric_assessment": "Assessment summary for derived transformation metrics.",
            "synthetic_reference_row_status": "Status table for synthetic reference rows injected into the workflow.",
        },
        notes=[
            "Primary outputs remain at the workflow root for quick inspection.",
            "Supporting evidence is split by use under supporting_files/.",
        ],
    )

    return {
        **output_paths,
        "mapping_coverage": coverage_result,
        "gap_diagnostics": diagnostics_artifacts.get("gap_diagnostics"),
        "mapping_rundown_by_sheet": diagnostics_artifacts.get("mapping_rundown_by_sheet"),
        "mapping_rundown_details": diagnostics_artifacts.get("mapping_rundown_details"),
        "comparison_issue_summary": diagnostics_artifacts.get("comparison_issue_summary"),
        "comparison_issue_cause_summary": diagnostics_artifacts.get("comparison_issue_cause_summary"),
        "charts_written": len(written_charts),
        "dashboard_index": str(dashboard_index) if dashboard_index else None,
        "chart_series_latest": str(latest_snapshot_path) if latest_snapshot_path else None,
        "chart_series_daily_archive": str(daily_archive_path) if daily_archive_path else None,
        "chart_line_mapping_ledger": str(chart_line_mapping_path) if chart_line_mapping_path else None,
        "chart_total_component_ledger": str(total_component_path) if total_component_path else None,
        "chart_navigation_hierarchy": str(chart_hierarchy_json_path) if chart_hierarchy_json_path else None,
        "chart_navigation_hierarchy_flat": str(chart_hierarchy_csv_path) if chart_hierarchy_csv_path else None,
        "atomic_many_to_many_errors": atomic_many_to_many_errors_path,
        "atomic_shadow_delta_series": atomic_shadow_delta_series_path,
        "atomic_shadow_delta_totals": atomic_shadow_delta_totals_path,
        "atomic_shadow_delta_summary": atomic_shadow_delta_summary_path,
        "shadow_compare_summary": shadow_compare_path,
        "leap_refresh_log": leap_refresh_log,
        "diagnostics": checks,
        "derived_transformation_leap_long": derived_artifacts.get("derived_transformation_leap_long"),
        "derived_transformation_leap_audit": derived_artifacts.get("derived_transformation_leap_audit"),
        "derived_transformation_comparison_audit": derived_artifacts.get("derived_transformation_comparison_audit"),
        "derived_transformation_metric_assessment": derived_artifacts.get("derived_transformation_metric_assessment"),
        "synthetic_reference_row_status": (
            str(layout.mapping / "synthetic_reference_row_status.csv")
            if (layout.mapping / "synthetic_reference_row_status.csv").exists()
            else None
        ),
        "output_manifest": str(manifest),
    }


if __name__ == "__main__":  # pragma: no cover
    result = run_workflow()
    print("[OK] LEAP Results dashboard V2 workflow complete.", flush=True)
    for k, v in result.items():
        print(f"- {k}: {v}", flush=True)
#%%


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
#%%
