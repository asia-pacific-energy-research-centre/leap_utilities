#%%
"""
Build 9th end-node sector/fuel subtotal mappings from the merged projection file.

The output maps each most-specific 9th sector/fuel pair to the row-level
``subtotal_results`` flag used in projected years. It keeps full hierarchy
columns so subtotal rows can be audited back to the source table.
"""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path
from typing import Iterable

import pandas as pd

from codebase.utilities.workflow_outputs import build_workflow_output_layout, write_output_manifest
from codebase.utilities.output_paths import MAPPINGS_ROOT


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


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


DATA_PATH = _resolve(r"C:\Users\Work\github\leap_utilities\data\merged_file_energy_ALL_20251106.csv")
OUTPUT_DIR = MAPPINGS_ROOT / "ninth_end_node_subtotal_mapping"

PROJECTED_START_YEAR = 2023
SECTOR_COLS = ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]
FUEL_COLS = ["fuels", "subfuels"]
META_COLS = ["scenarios", "economy", "subtotal_layout", "subtotal_results"]


def _year_columns(columns: Iterable[str]) -> list[str]:
    return [str(col) for col in columns if str(col).isdigit()]


def _projected_year_columns(columns: Iterable[str]) -> list[str]:
    return [col for col in _year_columns(columns) if int(col) >= PROJECTED_START_YEAR]


def _clean_code(value: object) -> str:
    if pd.isna(value):
        return "x"
    cleaned = str(value).strip()
    return cleaned if cleaned else "x"


def _coerce_subtotal_flag(value: object) -> bool:
    if pd.isna(value):
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    return str(value).strip().lower() in {"true", "1", "yes", "y"}


def _most_specific(row: pd.Series, columns: list[str]) -> tuple[str, str]:
    for col in reversed(columns):
        value = _clean_code(row.get(col))
        if value != "x":
            return value, col
    return "x", columns[0]


def _add_end_nodes(df: pd.DataFrame) -> pd.DataFrame:
    working = df.copy()
    sector_nodes = working.apply(lambda row: _most_specific(row, SECTOR_COLS), axis=1)
    fuel_nodes = working.apply(lambda row: _most_specific(row, FUEL_COLS), axis=1)

    working["ninth_end_node_sector"] = [value for value, _ in sector_nodes]
    working["ninth_end_node_sector_level"] = [level for _, level in sector_nodes]
    working["ninth_end_node_fuel"] = [value for value, _ in fuel_nodes]
    working["ninth_end_node_fuel_level"] = [level for _, level in fuel_nodes]
    return working


def _has_projected_data(df: pd.DataFrame, projected_years: list[str]) -> pd.Series:
    values = df[projected_years].apply(pd.to_numeric, errors="coerce").fillna(0)
    return values.abs().sum(axis=1) != 0


def _join_values(values: pd.Series) -> str:
    unique = sorted({str(value) for value in values.dropna() if str(value).strip()})
    return " | ".join(unique)


def _join_limited(values: pd.Series, limit: int = 60) -> str:
    unique = sorted({str(value) for value in values.dropna() if str(value).strip()})
    shown = unique[:limit]
    suffix = f" | ... ({len(unique) - limit} more)" if len(unique) > limit else ""
    return " | ".join(shown) + suffix


def _status(flags: pd.Series) -> str:
    unique = {bool(value) for value in flags.dropna()}
    if unique == {True}:
        return "subtotal"
    if unique == {False}:
        return "non_subtotal"
    return "mixed"


def _build_mapping(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    group_cols = ["ninth_end_node_sector", "ninth_end_node_fuel"]
    detailed_group_cols = [
        *SECTOR_COLS,
        *FUEL_COLS,
        "ninth_end_node_sector",
        "ninth_end_node_sector_level",
        "ninth_end_node_fuel",
        "ninth_end_node_fuel_level",
    ]

    detailed = (
        df.groupby(detailed_group_cols, dropna=False)
        .agg(
            subtotal_status=("subtotal_results", _status),
            subtotal_results_values=("subtotal_results", lambda s: _join_values(s.astype(bool))),
            row_count=("subtotal_results", "size"),
            economy_count=("economy", "nunique"),
            economies=("economy", _join_limited),
            scenarios=("scenarios", _join_limited),
            has_projected_data=("has_projected_data_2023_plus", "max"),
        )
        .reset_index()
        .sort_values(
            [
                "ninth_end_node_sector",
                "ninth_end_node_fuel",
                "sectors",
                "sub1sectors",
                "sub2sectors",
                "sub3sectors",
                "sub4sectors",
                "fuels",
                "subfuels",
            ]
        )
    )

    pair_summary = (
        df.groupby(group_cols, dropna=False)
        .agg(
            subtotal_status=("subtotal_results", _status),
            subtotal_results_values=("subtotal_results", lambda s: _join_values(s.astype(bool))),
            row_count=("subtotal_results", "size"),
            economy_count=("economy", "nunique"),
            economies=("economy", _join_limited),
            scenarios=("scenarios", _join_limited),
            sector_levels=("ninth_end_node_sector_level", _join_values),
            fuel_levels=("ninth_end_node_fuel_level", _join_values),
            has_projected_data=("has_projected_data_2023_plus", "max"),
        )
        .reset_index()
        .sort_values(["ninth_end_node_sector", "ninth_end_node_fuel"])
    )
    full_path_counts = detailed.groupby(group_cols, dropna=False).size().reset_index(name="full_path_count")
    pair_summary = pair_summary.merge(full_path_counts, on=group_cols, how="left")

    economy_scenario = (
        df.groupby([*group_cols, "economy", "scenarios"], dropna=False)
        .agg(
            subtotal_status=("subtotal_results", _status),
            subtotal_results_values=("subtotal_results", lambda s: _join_values(s.astype(bool))),
            row_count=("subtotal_results", "size"),
            sector_levels=("ninth_end_node_sector_level", _join_values),
            fuel_levels=("ninth_end_node_fuel_level", _join_values),
            has_projected_data=("has_projected_data_2023_plus", "max"),
        )
        .reset_index()
        .sort_values(["ninth_end_node_sector", "ninth_end_node_fuel", "economy", "scenarios"])
    )

    conflicts = pair_summary[pair_summary["subtotal_status"].eq("mixed")].copy()
    return pair_summary, economy_scenario, detailed, conflicts


def run_workflow() -> dict[str, object]:
    if not DATA_PATH.exists():
        raise FileNotFoundError(DATA_PATH)

    header = pd.read_csv(DATA_PATH, nrows=0)
    projected_years = _projected_year_columns(header.columns)
    if not projected_years:
        raise ValueError(f"No year columns >= {PROJECTED_START_YEAR} found in {DATA_PATH}")

    keep_cols = [col for col in META_COLS + SECTOR_COLS + FUEL_COLS if col in header.columns]
    keep_cols += projected_years
    missing = [col for col in META_COLS + SECTOR_COLS + FUEL_COLS if col not in header.columns]
    if missing:
        raise ValueError(f"Missing required source columns: {missing}")

    raw = pd.read_csv(DATA_PATH, usecols=keep_cols, low_memory=False)
    for col in META_COLS + SECTOR_COLS + FUEL_COLS:
        if col in {"subtotal_layout", "subtotal_results"}:
            continue
        raw[col] = raw[col].map(_clean_code)
    raw["subtotal_results"] = raw["subtotal_results"].map(_coerce_subtotal_flag)
    raw["has_projected_data_2023_plus"] = _has_projected_data(raw, projected_years)

    # Keep the mapping scoped to the projected horizon while avoiding rows that
    # are present structurally but have no values in any projected year.
    projected = raw[raw["has_projected_data_2023_plus"]].copy()
    projected = _add_end_nodes(projected)

    pair_summary, economy_scenario, detailed, conflicts = _build_mapping(projected)

    layout = build_workflow_output_layout(OUTPUT_DIR)
    pair_summary_path = layout.root / "ninth_end_node_sector_fuel_subtotal_mapping.csv"
    economy_scenario_path = (
        layout.analysis / "ninth_end_node_sector_fuel_subtotal_mapping_by_economy_scenario.csv"
    )
    detailed_path = layout.analysis / "ninth_end_node_sector_fuel_subtotal_mapping_detailed.csv"
    conflicts_path = layout.checks / "ninth_end_node_sector_fuel_subtotal_conflicts.csv"
    workbook_path = layout.root / "ninth_end_node_sector_fuel_subtotal_mapping.xlsx"
    summary_path = layout.runtime / "ninth_end_node_sector_fuel_subtotal_mapping_summary.json"

    pair_summary.to_csv(pair_summary_path, index=False)
    economy_scenario.to_csv(economy_scenario_path, index=False)
    detailed.to_csv(detailed_path, index=False)
    conflicts.to_csv(conflicts_path, index=False)
    with pd.ExcelWriter(workbook_path) as writer:
        pair_summary.to_excel(writer, sheet_name="end_node_pair_summary", index=False)
        economy_scenario.to_excel(writer, sheet_name="by_economy_scenario", index=False)
        detailed.to_excel(writer, sheet_name="full_path_detail", index=False)
        conflicts.to_excel(writer, sheet_name="mixed_status_conflicts", index=False)

    summary = {
        "source": str(DATA_PATH.relative_to(REPO_ROOT)),
        "projected_start_year": PROJECTED_START_YEAR,
        "projected_year_count": len(projected_years),
        "source_rows": int(len(raw)),
        "projected_nonzero_rows": int(len(projected)),
        "end_node_pair_count": int(len(pair_summary)),
        "full_path_count": int(len(detailed)),
        "economy_scenario_pair_count": int(len(economy_scenario)),
        "mixed_status_pair_count": int(len(conflicts)),
        "mixed_status_economy_scenario_pair_count": int(
            economy_scenario["subtotal_status"].eq("mixed").sum()
        ),
        "subtotal_status_counts": pair_summary["subtotal_status"].value_counts().to_dict(),
        "subtotal_status_counts_by_economy_scenario": economy_scenario["subtotal_status"]
        .value_counts()
        .to_dict(),
        "outputs": {
            "pair_summary_csv": str(pair_summary_path.relative_to(REPO_ROOT)),
            "economy_scenario_csv": str(economy_scenario_path.relative_to(REPO_ROOT)),
            "detailed_csv": str(detailed_path.relative_to(REPO_ROOT)),
            "conflicts_csv": str(conflicts_path.relative_to(REPO_ROOT)),
            "workbook": str(workbook_path.relative_to(REPO_ROOT)),
        },
    }
    summary_path.write_text(json.dumps(summary, ensure_ascii=True, indent=2), encoding="utf-8")
    manifest_path = write_output_manifest(
        out_dir=layout.root,
        primary_outputs={
            "pair_summary_csv": str(pair_summary_path),
            "workbook": str(workbook_path),
        },
        supporting_outputs={
            "economy_scenario_csv": str(economy_scenario_path),
            "detailed_csv": str(detailed_path),
            "conflicts_csv": str(conflicts_path),
            "summary_json": str(summary_path),
        },
        primary_output_descriptions={
            "pair_summary_csv": "Primary subtotal-status mapping by 9th end-node sector and fuel.",
            "workbook": "Workbook version of the subtotal mapping with audit sheets included.",
        },
        supporting_output_descriptions={
            "economy_scenario_csv": "Economy and scenario breakout of subtotal status by end-node pair.",
            "detailed_csv": "Full path-level breakdown behind each end-node subtotal status.",
            "conflicts_csv": "Pairs whose subtotal status differs across source rows.",
            "summary_json": "Run summary and output inventory for the subtotal mapping workflow.",
        },
        notes=[
            "Primary pair-summary outputs stay at the workflow root.",
            "Per-economy, detailed, and conflict diagnostics are grouped under supporting_files/.",
        ],
    )
    print(json.dumps(summary, ensure_ascii=True, indent=2))
    return {**summary, "output_manifest_json": str(manifest_path.relative_to(REPO_ROOT))}


if __name__ == "__main__":
    run_workflow()
