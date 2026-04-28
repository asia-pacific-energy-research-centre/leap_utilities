from __future__ import annotations

from pathlib import Path

import pandas as pd

from codebase import leap_results_dashboard_workflow as v1_workflow


def write_core_outputs(
    *,
    out_dir: Path,
    comparison_long: pd.DataFrame,
    comparison_wide: pd.DataFrame,
    mapping_status: pd.DataFrame,
    leap_long: pd.DataFrame,
    supporting_dir: Path | None = None,
    atomic_comparison_long: pd.DataFrame | None = None,
    atomic_comparison_wide: pd.DataFrame | None = None,
    atomic_mapping_edges: pd.DataFrame | None = None,
    atomic_validation_report: pd.DataFrame | None = None,
) -> dict[str, str]:
    out_dir.mkdir(parents=True, exist_ok=True)
    comparison_long_path = out_dir / "comparison_long.csv"
    comparison_wide_path = out_dir / "comparison_wide.csv"
    mapping_status_path = out_dir / "mapping_status.xlsx"
    leap_long_path = out_dir / "leap_long.csv"

    comparison_long.to_csv(comparison_long_path, index=False)
    comparison_wide.to_csv(comparison_wide_path, index=False)
    v1_workflow._write_workbook_with_header_comments(mapping_status, mapping_status_path, sheet_name="mapping_status")
    leap_long.to_csv(leap_long_path, index=False)

    atomic_dir = (supporting_dir / "atomic") if supporting_dir is not None else out_dir
    if atomic_dir is not None:
        atomic_dir.mkdir(parents=True, exist_ok=True)

    out = {
        "comparison_long": str(comparison_long_path),
        "comparison_wide": str(comparison_wide_path),
        "mapping_status": str(mapping_status_path),
        "leap_long": str(leap_long_path),
    }
    if atomic_comparison_long is not None:
        p = atomic_dir / "atomic_comparison_long.csv"
        atomic_comparison_long.to_csv(p, index=False)
        out["atomic_comparison_long"] = str(p)
    if atomic_comparison_wide is not None:
        p = atomic_dir / "atomic_comparison_wide.csv"
        atomic_comparison_wide.to_csv(p, index=False)
        out["atomic_comparison_wide"] = str(p)
    if atomic_mapping_edges is not None:
        p = atomic_dir / "atomic_mapping_edges.csv"
        atomic_mapping_edges.to_csv(p, index=False)
        out["atomic_mapping_edges"] = str(p)
    if atomic_validation_report is not None:
        p = atomic_dir / "atomic_validation_report.csv"
        atomic_validation_report.to_csv(p, index=False)
        out["atomic_validation_report"] = str(p)
    return out
