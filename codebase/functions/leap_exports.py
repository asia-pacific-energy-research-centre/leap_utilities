from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd

from codebase.utilities import workflow_common
from codebase.functions.leap_excel_io import finalise_export_df, save_export_files


DEFAULT_SHEET_NAME = "LEAP"


def build_workbook_filename(
    economy_label: str,
    scenarios: str | Iterable[str] | None,
    template: str,
    fallback_template: str | None = None,
) -> str:
    """Return a safe export workbook filename."""
    return workflow_common.build_workflow_export_filename(
        economy_label=economy_label,
        scenarios=scenarios,
        template=template,
        format_segment_fn=workflow_common.format_filename_segment,
        fallback_template=fallback_template,
    )


def build_export_dataframe(
    log_df: pd.DataFrame,
    scenario: str,
    region: str,
    base_year: int,
    final_year: int,
) -> pd.DataFrame | None:
    """Convert log-style rows to LEAP import dataframe layout."""
    return finalise_export_df(
        log_df=log_df,
        scenario=scenario,
        region=region,
        base_year=base_year,
        final_year=final_year,
    )


def save_workbook(
    leap_export_df: pd.DataFrame,
    output_path: Path | str,
    *,
    base_year: int,
    final_year: int,
    model_name: str,
    viewing_df: pd.DataFrame | None = None,
) -> Path:
    """Write LEAP + FOR_VIEWING sheets using the project-standard layout."""
    export_path = Path(output_path)
    save_export_files(
        leap_export_df=leap_export_df,
        export_df_for_viewing=viewing_df if viewing_df is not None else leap_export_df,
        leap_export_filename=export_path,
        base_year=base_year,
        final_year=final_year,
        model_name=model_name,
    )
    return export_path


def build_and_save_workbook(
    log_df: pd.DataFrame,
    output_path: Path | str,
    *,
    scenario: str,
    region: str,
    base_year: int,
    final_year: int,
    model_name: str,
    viewing_df: pd.DataFrame | None = None,
) -> Path | None:
    """Build the LEAP dataframe from log rows and save it to disk."""
    leap_export_df = build_export_dataframe(
        log_df=log_df,
        scenario=scenario,
        region=region,
        base_year=base_year,
        final_year=final_year,
    )
    if leap_export_df is None or leap_export_df.empty:
        return None
    return save_workbook(
        leap_export_df=leap_export_df,
        output_path=output_path,
        base_year=base_year,
        final_year=final_year,
        model_name=model_name,
        viewing_df=viewing_df,
    )


def list_scenarios(export_path: Path, sheet_name: str = DEFAULT_SHEET_NAME) -> list[str]:
    """Return scenario labels in declaration order from a workbook."""
    return workflow_common.list_export_scenarios(export_path, sheet_name)


def validate_region(export_path: Path, region: str, sheet_name: str = DEFAULT_SHEET_NAME) -> None:
    """Raise if region is absent from the export workbook."""
    workflow_common.validate_export_region(export_path, sheet_name, region)


def find_workbook(
    directory: Path | str,
    prefix: str,
    filename: str | None = None,
) -> Path:
    """Locate an explicit workbook name or the latest file matching a prefix."""
    return workflow_common.find_latest_export_workbook(
        directory=directory,
        prefix=prefix,
        filename=filename,
    )
