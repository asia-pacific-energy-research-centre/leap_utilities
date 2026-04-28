from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd

_DIRECTORY_PREVIEW_LIMIT = 25


@dataclass(frozen=True)
class WorkflowOutputLayout:
    root: Path
    supporting: Path
    mapping: Path
    checks: Path
    runtime: Path
    charting: Path
    analysis: Path
    charts: Path
    dashboards: Path

    @property
    def diagnostics(self) -> Path:
        return self.checks

    @property
    def coverage(self) -> Path:
        return self.checks

    @property
    def ledgers(self) -> Path:
        return self.charting

    @property
    def navigation(self) -> Path:
        return self.charting

    @property
    def snapshots(self) -> Path:
        return self.charting

    @property
    def atomic(self) -> Path:
        return self.analysis

    @property
    def derived(self) -> Path:
        return self.analysis

    @property
    def shadow_compare(self) -> Path:
        return self.analysis


_COLUMN_GLOSSARY = {
    "economy": "Economy code for the series or mapping row.",
    "scenario": "Scenario label used for the row.",
    "sheet": "Dashboard sheet or comparison group name.",
    "sheet_name": "Original LEAP sheet name.",
    "fuel_label": "Display fuel/category label used in outputs and charts.",
    "source": "Series source, such as LEAP, base, or projection.",
    "year": "Year for the observation.",
    "value": "Numeric value for the row after workflow transformations.",
    "leap_value": "Raw or prepared LEAP numeric value.",
    "leap_units": "Units reported by the LEAP source row.",
    "measure": "Measure label used to distinguish related chart lines.",
    "mapping_source": "How the sector/fuel mapping was resolved.",
    "flow_source": "How the ESTO flow mapping was resolved.",
    "fuel_source": "How the 9th fuel mapping was resolved.",
    "mapped": "True when the row has the required mappings for comparison.",
    "partially_mapped": "True when only part of the mapping information was resolved.",
    "has_any_mapping": "True when any mapping fields were filled.",
    "esto_flow": "ESTO flow code or label used for comparison.",
    "esto_product": "ESTO product code or label used for comparison.",
    "ninth_fuel_code": "9th fuel code used for projection comparison.",
    "sector_code_9th": "9th sector code tied to the row.",
    "leap_flow": "Mapped LEAP flow code from the balance workflow.",
    "leap_product": "Mapped LEAP product code from the balance workflow.",
    "leap_flow_name": "Human-readable LEAP flow name.",
    "leap_product_name": "Human-readable LEAP product name.",
    "rows": "Count of source rows contributing to the summary row.",
    "value_pj": "Summed value in petajoules.",
    "row_count": "Number of rows in the grouped result.",
    "reason": "Issue or classification reason for the row.",
    "details": "Extra detail about the issue or mapping decision.",
    "region": "Region label carried through from LEAP extraction.",
    "dashboard_path": "Rendered dashboard page path for the chart.",
    "chart_file": "Relative chart file written by the workflow.",
    "issue_cause": "High-level classification for a comparison problem.",
    "agent_debug_hint": "Short debugging hint for the issue row.",
}

_SUPPORTING_FOLDER_GUIDE = {
    "mapping": "How rows were matched, mapped, reassigned, or grouped.",
    "checks": "Coverage checks, duplicates, gaps, and other QA outputs.",
    "runtime": "Problems or exception-style outputs produced during the run.",
    "charting": "Chart ledgers, hierarchy files, and chart-support artifacts.",
    "analysis": "Deeper method-comparison, derived, atomic, or shadow-analysis outputs.",
}


def build_workflow_output_layout(out_dir: Path) -> WorkflowOutputLayout:
    root = Path(out_dir)
    supporting = root / "supporting_files"
    layout = WorkflowOutputLayout(
        root=root,
        supporting=supporting,
        mapping=supporting / "mapping",
        checks=supporting / "checks",
        runtime=supporting / "runtime",
        charting=supporting / "charting",
        analysis=supporting / "analysis",
        charts=root / "charts",
        dashboards=root / "dashboards",
    )
    for path in [
        layout.root,
        layout.supporting,
        layout.mapping,
        layout.checks,
        layout.runtime,
        layout.charting,
        layout.analysis,
        layout.charts,
        layout.dashboards,
    ]:
        path.mkdir(parents=True, exist_ok=True)
    return layout


def _column_descriptions(columns: list[str]) -> list[dict[str, str | None]]:
    return [
        {
            "name": column,
            "description": _COLUMN_GLOSSARY.get(column),
        }
        for column in columns
    ]


def _read_tabular_columns(path: Path) -> list[str] | None:
    suffix = path.suffix.lower()
    try:
        if suffix == ".csv":
            return pd.read_csv(path, nrows=0).columns.astype(str).tolist()
        if suffix in {".xlsx", ".xls"}:
            workbook = pd.ExcelFile(path)
            if not workbook.sheet_names:
                return []
            first_sheet = workbook.sheet_names[0]
            return pd.read_excel(path, sheet_name=first_sheet, nrows=0).columns.astype(str).tolist()
    except Exception:
        return None
    return None


def _describe_output(alias: str, target: str, description: str | None) -> dict[str, object]:
    path = Path(target)
    entry: dict[str, object] = {
        "path": target,
        "description": description,
        "exists": path.exists(),
        "type": "directory" if path.is_dir() else "file",
    }
    if path.is_file():
        entry["suffix"] = path.suffix.lower()
        columns = _read_tabular_columns(path)
        if columns is not None:
            entry["columns"] = _column_descriptions(columns)
    elif path.is_dir():
        try:
            items = sorted(child.name for child in path.iterdir())
            entry["item_count"] = len(items)
            entry["contains_preview"] = items[:_DIRECTORY_PREVIEW_LIMIT]
        except OSError:
            entry["item_count"] = None
            entry["contains_preview"] = []
    entry["alias"] = alias
    return entry


def write_output_manifest(
    *,
    out_dir: Path,
    primary_outputs: dict[str, str | None],
    supporting_outputs: dict[str, str | None],
    primary_output_descriptions: dict[str, str] | None = None,
    supporting_output_descriptions: dict[str, str] | None = None,
    notes: Iterable[str] | None = None,
) -> Path:
    manifest_path = Path(out_dir) / "output_manifest.json"
    primary_output_descriptions = primary_output_descriptions or {}
    supporting_output_descriptions = supporting_output_descriptions or {}
    supporting_root = Path(out_dir) / "supporting_files"
    payload = {
        "primary_outputs": {
            key: _describe_output(key, value, primary_output_descriptions.get(key))
            for key, value in primary_outputs.items()
            if value
        },
        "supporting_outputs": {
            key: _describe_output(key, value, supporting_output_descriptions.get(key))
            for key, value in supporting_outputs.items()
            if value
        },
        "supporting_folders": {
            name: _describe_output(
                name,
                str(supporting_root / name),
                description,
            )
            for name, description in _SUPPORTING_FOLDER_GUIDE.items()
            if (supporting_root / name).exists()
        },
        "notes": list(notes or ()),
    }
    manifest_path.write_text(json.dumps(payload, ensure_ascii=True, indent=2), encoding="utf-8")
    return manifest_path
