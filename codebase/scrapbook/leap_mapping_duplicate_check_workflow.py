#%%
"""
Check explicit LEAP balance mapping sheets for duplicate mapping risks.

This writes a workbook under outputs/mappings/mapping_checks with:
- exact duplicate active source/target mappings
- source/fuel keys that have both active mappings and remove_row=True rows
"""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path

import pandas as pd

from codebase.utilities.workflow_outputs import write_output_manifest

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


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


MAPPING_WORKBOOK_PATH = _resolve("config/leap_mappings.xlsx")
OUTPUT_DIR = _resolve("outputs/mappings/mapping_checks")
OUTPUT_WORKBOOK_PATH = OUTPUT_DIR / "leap_mapping_duplicate_check.xlsx"

MAPPING_SHEETS = {
    "leap_combined_esto": ["esto_flow", "esto_product"],
    "leap_combined_ninth": ["ninth_sector", "ninth_fuel"],
}


#%%
def _truthy(value: object) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "yes", "y", "on", "t"}


def _clean(value: object) -> str:
    text = str(value or "").strip()
    return "" if text.lower() in {"nan", "none"} else text


def _clean_frame(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    return frame[columns].map(_clean)


def _sheet_duplicate_diagnostics(
    frame: pd.DataFrame,
    *,
    target_columns: list[str],
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, int]]:
    work = frame.copy().fillna("")
    source_columns = ["leap_sector_name_full_path", "raw_leap_fuel_name"]
    required_columns = [*source_columns, *target_columns, "remove_row", "duplicate_to_remove"]
    for column in required_columns:
        if column not in work.columns:
            work[column] = ""

    active = ~work["remove_row"].map(_truthy) & ~work["duplicate_to_remove"].map(_truthy)
    valid = active.copy()
    for column in [*source_columns, *target_columns]:
        valid &= work[column].map(_clean).ne("")

    exact_columns = [*source_columns, *target_columns]
    exact_duplicates = work.loc[valid & work.loc[valid].duplicated(subset=exact_columns, keep=False)].copy()
    exact_duplicates.insert(0, "mapping_row_number", exact_duplicates.index + 2)

    removed_keys = set(
        map(
            tuple,
            _clean_frame(work.loc[work["remove_row"].map(_truthy)], source_columns).to_numpy(),
        )
    )
    active_keys = set(map(tuple, _clean_frame(work.loc[valid], source_columns).to_numpy()))
    conflicting_removed_keys = removed_keys & active_keys
    remove_conflicts = work.loc[
        work.apply(
            lambda row: (_clean(row[source_columns[0]]), _clean(row[source_columns[1]])) in conflicting_removed_keys,
            axis=1,
        )
    ].copy()
    remove_conflicts.insert(0, "mapping_row_number", remove_conflicts.index + 2)

    summary = {
        "exact_duplicate_active_rows": int(len(exact_duplicates)),
        "exact_duplicate_groups": int(exact_duplicates[exact_columns].drop_duplicates().shape[0])
        if not exact_duplicates.empty
        else 0,
        "active_remove_row_conflict_rows": int(len(remove_conflicts)),
        "active_remove_row_conflict_groups": int(len(conflicting_removed_keys)),
    }
    return exact_duplicates, remove_conflicts, summary


def write_mapping_duplicate_check(
    *,
    mapping_workbook_path: Path = MAPPING_WORKBOOK_PATH,
    output_workbook_path: Path = OUTPUT_WORKBOOK_PATH,
) -> dict[str, object]:
    summary_rows: list[dict[str, object]] = []
    output_sheets: list[tuple[str, pd.DataFrame]] = []

    for sheet_name, target_columns in MAPPING_SHEETS.items():
        frame = pd.read_excel(mapping_workbook_path, sheet_name=sheet_name, dtype=object)
        exact_duplicates, remove_conflicts, summary = _sheet_duplicate_diagnostics(
            frame,
            target_columns=target_columns,
        )
        summary_rows.append({"sheet": sheet_name, **summary})
        output_sheets.append((f"{sheet_name}_exact_duplicates"[:31], exact_duplicates))
        output_sheets.append((f"{sheet_name}_remove_conflicts"[:31], remove_conflicts))

    summary = pd.DataFrame(summary_rows)
    output_workbook_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_workbook_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        for sheet_name, frame in output_sheets:
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
    manifest_path = write_output_manifest(
        out_dir=output_workbook_path.parent,
        primary_outputs={"output_workbook": str(output_workbook_path)},
        supporting_outputs={},
        primary_output_descriptions={
            "output_workbook": "Workbook containing duplicate and remove-row conflict checks for LEAP balance mappings.",
        },
        notes=[
            "This workflow produces a single primary workbook.",
        ],
    )

    return {
        "output_workbook": str(output_workbook_path),
        "summary": summary,
        "output_manifest_json": str(manifest_path),
    }


#%%
RUN_WORKFLOW = True
WORKFLOW_RESULT: dict[str, object] | None = None
if RUN_WORKFLOW:
    WORKFLOW_RESULT = write_mapping_duplicate_check()
    print("[OK] LEAP mapping duplicate check complete.")
    for key, value in WORKFLOW_RESULT.items():
        if isinstance(value, pd.DataFrame):
            print(f"{key}:")
            print(value.to_string(index=False))
        else:
            print(f"{key}: {value}")
