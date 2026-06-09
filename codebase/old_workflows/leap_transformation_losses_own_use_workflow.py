#%%
"""Legacy power-only transformation losses / own-use extraction.

This workflow is no longer part of the official `leap_results_dashboard_v2_workflow.py`
path. The official dashboard now uses generalized derived transformation-gap metrics
generated from the shared sheet-map-driven component sheets instead.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from config import leap_transformation_losses_own_use_config as cfg  # noqa: E402
from codebase.utilities.leap_transformation_losses_own_use import (  # noqa: E402
    build_dashboard_leap_long,
    build_inspection_charts,
    build_ninth_mapped_charts,
    build_ninth_merge_exploration,
    build_qa_flags,
    empty_dashboard_leap_long,
    empty_dimension_discovery,
    empty_normalized_long,
    empty_qa_flags,
    empty_raw_result_pulls,
    extract_transformation_losses_own_use,
    normalize_raw_results,
    scale_raw_energy_values,
    write_losses_own_use_outputs,
)


def _resolve(path_value: Path | str) -> Path:
    text = str(path_value).replace("\\", "/")
    path = Path(text)
    return path if path.is_absolute() else (REPO_ROOT / path).resolve()


def ensure_repo_root() -> None:
    if Path.cwd() != REPO_ROOT:
        os.chdir(REPO_ROOT)


def _read_csv_if_exists(path: Path, empty_frame: pd.DataFrame) -> pd.DataFrame:
    if not path.exists():
        return empty_frame
    try:
        return pd.read_csv(path)
    except Exception:
        return empty_frame


def _load_existing_outputs(output_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    raw_df = _read_csv_if_exists(output_dir / "raw_result_pulls.csv", empty_raw_result_pulls())
    normalized_df = _read_csv_if_exists(output_dir / "normalized_long.csv", empty_normalized_long())
    dashboard_df = _read_csv_if_exists(output_dir / "dashboard_leap_long.csv", empty_dashboard_leap_long())
    qa_df = _read_csv_if_exists(output_dir / "qa_flags.csv", empty_qa_flags())
    discovery_df = _read_csv_if_exists(output_dir / "dimension_discovery.csv", empty_dimension_discovery())
    return raw_df, normalized_df, dashboard_df, qa_df, discovery_df


def _write_merge_exploration(
    *,
    output_dir: Path,
    normalized_df: pd.DataFrame,
    economy: str,
) -> Path:
    projection_df = pd.read_csv(_resolve(cfg.PROJECTION_TABLE_PATH), low_memory=False)
    merge_df = build_ninth_merge_exploration(
        normalized_df,
        projection_df=projection_df,
        projection_economy=f"20_{economy}" if "_" not in economy else economy,
        scenario="reference",
        merge_candidates=cfg.NINTH_MERGE_CANDIDATES,
    )
    out_path = output_dir / "ninth_merge_exploration.csv"
    merge_df.to_csv(out_path, index=False)
    return out_path


def run_workflow(
    *,
    economy: str = "USA",
    dry_run: bool = False,
    use_api: bool = True,
    rebuild_from_existing: bool = True,
) -> dict[str, object]:
    ensure_repo_root()
    output_dir = _resolve(cfg.OUTPUT_ROOT) / economy
    regions = tuple(cfg.REGIONS) if cfg.REGIONS else ("United States",)

    if dry_run:
        raw_df = empty_raw_result_pulls()
        discovery_df = empty_dimension_discovery()
        normalized_df = empty_normalized_long()
        dashboard_df = empty_dashboard_leap_long()
        qa_df = empty_qa_flags()
    elif not use_api:
        raw_df, normalized_df, dashboard_df, qa_df, discovery_df = _load_existing_outputs(output_dir)
        raw_df = scale_raw_energy_values(raw_df)
        if rebuild_from_existing:
            if not raw_df.empty:
                normalized_df = normalize_raw_results(raw_df)
            if not normalized_df.empty:
                dashboard_df = build_dashboard_leap_long(
                    normalized_df,
                    economy=f"20_{economy}" if "_" not in economy else economy,
                    dashboard_sheet_definitions=cfg.DASHBOARD_SHEET_DEFINITIONS,
                    unit=cfg.UNIT,
                )
            if not normalized_df.empty:
                qa_df = build_qa_flags(normalized_df)
    else:
        raw_df, discovery_df = extract_transformation_losses_own_use(
            scenarios=cfg.SCENARIOS,
            regions=regions,
            years=cfg.YEAR_RANGE,
            modules=cfg.MODULES,
            result_dimension_hints=cfg.RESULT_DIMENSION_HINTS,
            result_member_hints=cfg.RESULT_MEMBER_HINTS,
            manual_result_member_overrides=cfg.MANUAL_RESULT_MEMBER_OVERRIDES,
            unit=cfg.UNIT,
        )
        normalized_df = normalize_raw_results(raw_df)
        dashboard_df = build_dashboard_leap_long(
            normalized_df,
            economy=f"20_{economy}" if "_" not in economy else economy,
            dashboard_sheet_definitions=cfg.DASHBOARD_SHEET_DEFINITIONS,
            unit=cfg.UNIT,
        )
        qa_df = build_qa_flags(normalized_df)

    output_paths = write_losses_own_use_outputs(
        output_dir=output_dir,
        raw_df=raw_df,
        normalized_df=normalized_df,
        dashboard_df=dashboard_df,
        qa_df=qa_df,
        discovery_df=discovery_df,
    )
    chart_outputs = build_inspection_charts(normalized_df, output_dir=output_dir)
    merge_exploration_path = _write_merge_exploration(
        output_dir=output_dir,
        normalized_df=normalized_df,
        economy=economy,
    )
    mapped_chart_outputs = build_ninth_mapped_charts(
        pd.read_csv(merge_exploration_path),
        output_dir=output_dir,
    )
    return {
        "economy": economy,
        "dry_run": dry_run,
        "use_api": use_api,
        "rebuild_from_existing": rebuild_from_existing,
        "rows": {
            "raw_result_pulls": int(len(raw_df)),
            "normalized_long": int(len(normalized_df)),
            "dashboard_leap_long": int(len(dashboard_df)),
            "qa_flags": int(len(qa_df)),
            "dimension_discovery": int(len(discovery_df)),
        },
        "outputs": output_paths,
        "charts": chart_outputs,
        "mapped_charts": mapped_chart_outputs,
        "merge_exploration": str(merge_exploration_path),
    }

#%%
# Simple notebook-focused configuration block.
NOTEBOOK_ECONOMY = "USA"
NOTEBOOK_DRY_RUN = False
NOTEBOOK_USE_API = False
NOTEBOOK_REBUILD_FROM_EXISTING = True


def run_with_notebook_config() -> dict[str, object]:
    """Run the workflow using the editable notebook constants above."""
    return run_workflow(
        economy=NOTEBOOK_ECONOMY,
        dry_run=NOTEBOOK_DRY_RUN,
        use_api=NOTEBOOK_USE_API,
        rebuild_from_existing=NOTEBOOK_REBUILD_FROM_EXISTING,
    )


if __name__ == "__main__":
    result = run_with_notebook_config()
    print(result)
#%%


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
