#%%
"""
Drop exact duplicate mapping rows from leap_combined_esto and leap_combined_ninth.

The workflow is intentionally conservative:
- it only removes extra rows from exact duplicate active groups
- it raises if any duplicate group has differing non-key columns
- it backs up the workbook before writing
"""

from __future__ import annotations

import os
import re
import shutil
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

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
DUPLICATE_MAPPINGS_CSV_PATH = OUTPUT_DIR / "leap_mapping_duplicate_mappings.csv"
ARCHIVE_DIR = MAPPING_WORKBOOK_PATH.parent / "archive"

SHEETS: dict[str, list[str]] = {
    "leap_combined_esto": ["esto_flow", "esto_product"],
    "leap_combined_ninth": ["ninth_sector", "ninth_fuel"],
}

SOURCE_COLUMNS = ["leap_sector_name_full_path", "raw_leap_fuel_name"]


#%%
def _clean(value: object) -> str:
    text = str(value or "").strip()
    return "" if text.lower() in {"", "nan", "none"} else text


def _truthy(value: object) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "t", "yes", "y", "on"}


def _assert_not_open(path: Path) -> None:
    try:
        with open(path, "r+b"):
            pass
    except PermissionError:
        raise PermissionError(
            f"{path.name} is open in another application (e.g. Excel). Close it and re-run the workflow."
        ) from None


def _backup_workbook(path: Path) -> Path:
    ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
    backup_path = ARCHIVE_DIR / f"{path.stem}.before_drop_duplicate_rows_{pd.Timestamp.now():%Y%m%d_%H%M%S}{path.suffix}"
    shutil.copy2(path, backup_path)
    return backup_path


def _active_mask(frame: pd.DataFrame) -> pd.Series:
    remove_mask = frame.get("remove_row", False)
    duplicate_mask = frame.get("duplicate_to_remove", False)
    remove_mask = pd.Series(remove_mask, index=frame.index).map(_truthy)
    duplicate_mask = pd.Series(duplicate_mask, index=frame.index).map(_truthy)
    return ~(remove_mask | duplicate_mask)


def _normalize_frame(frame: pd.DataFrame) -> pd.DataFrame:
    out = frame.copy().fillna("")
    for column in out.columns:
        out[column] = out[column].map(_clean)
    return out


def _duplicate_groups(frame: pd.DataFrame, *, sheet_name: str, target_columns: list[str]) -> pd.DataFrame:
    work = frame.copy().fillna("")
    required_cols = [*SOURCE_COLUMNS, *target_columns, "remove_row", "duplicate_to_remove"]
    for column in required_cols:
        if column not in work.columns:
            work[column] = ""
        work[column] = work[column].fillna("").astype(str).str.strip()

    active = work.loc[_active_mask(work)].copy()
    valid = active[SOURCE_COLUMNS + target_columns].apply(lambda col: col.map(_clean).ne("")).all(axis=1)
    active = active.loc[valid].copy()
    if active.empty:
        return pd.DataFrame(
            columns=[
                "sheet_name",
                "mapping_row_number",
                "duplicate_group_size",
                *SOURCE_COLUMNS,
                *target_columns,
            ]
        )

    duplicate_mask = active.duplicated(subset=[*SOURCE_COLUMNS, *target_columns], keep=False)
    duplicates = active.loc[duplicate_mask].copy()
    if duplicates.empty:
        return pd.DataFrame(
            columns=[
                "sheet_name",
                "mapping_row_number",
                "duplicate_group_size",
                *SOURCE_COLUMNS,
                *target_columns,
            ]
        )

    duplicates.insert(0, "mapping_row_number", duplicates.index + 2)
    duplicates.insert(0, "sheet_name", sheet_name)
    duplicates["duplicate_group_size"] = duplicates.groupby([*SOURCE_COLUMNS, *target_columns])[SOURCE_COLUMNS[0]].transform("size")
    return duplicates[
        [
            "sheet_name",
            "mapping_row_number",
            "duplicate_group_size",
            *SOURCE_COLUMNS,
            *target_columns,
        ]
    ].reset_index(drop=True)


def _duplicate_group_signature(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return frame.copy()
    columns = ["sheet_name", *SOURCE_COLUMNS]
    if "esto_flow" in frame.columns:
        columns.extend(["esto_flow", "esto_product"])
    if "ninth_sector" in frame.columns:
        columns.extend(["ninth_sector", "ninth_fuel"])
    columns.append("duplicate_group_size")
    signature = frame.loc[:, columns].copy().fillna("").astype(str)
    return signature.sort_values(columns).reset_index(drop=True)


def _raise_on_differences(frame: pd.DataFrame, *, sheet_name: str, target_columns: list[str]) -> None:
    work = frame.copy().fillna("")
    key_columns = [*SOURCE_COLUMNS, *target_columns]
    if work.empty:
        return

    for column in [*SOURCE_COLUMNS, *target_columns, "remove_row", "duplicate_to_remove"]:
        if column not in work.columns:
            work[column] = ""
        work[column] = work[column].fillna("").astype(str).str.strip()

    active = work.loc[_active_mask(work)].copy()
    valid = active[SOURCE_COLUMNS + target_columns].apply(lambda col: col.map(_clean).ne("")).all(axis=1)
    active = active.loc[valid].copy()
    if active.empty:
        return

    non_key_columns = [column for column in active.columns if column not in key_columns]
    if not non_key_columns:
        return

    mismatches: list[str] = []
    for _, group in active.groupby(key_columns, sort=False):
        if len(group) <= 1:
            continue
        normalized = _normalize_frame(group[non_key_columns])
        unique_rows = normalized.drop_duplicates()
        if len(unique_rows) > 1:
            row_numbers = (group.index + 2).tolist()
            changed_cols = [
                column
                for column in non_key_columns
                if normalized[column].nunique(dropna=False) > 1
            ]
            mismatches.append(
                f"{sheet_name} | rows {row_numbers} | differing columns: {', '.join(changed_cols)}"
            )

    if mismatches:
        preview = "\n".join(mismatches[:20])
        raise ValueError(
            "Duplicate groups contain differing non-key rows. "
            "No workbook changes were made.\n"
            f"{preview}"
        )


def _delete_rows_from_sheet(workbook_path: Path, *, sheet_name: str, row_numbers: list[int]) -> None:
    if not row_numbers:
        return
    workbook = load_workbook(workbook_path)
    worksheet = workbook[sheet_name]
    for row_number in sorted(row_numbers, reverse=True):
        worksheet.delete_rows(row_number, 1)
    workbook.save(workbook_path)


def _refresh_duplicate_report(workbook_path: Path, output_csv_path: Path) -> pd.DataFrame:
    summary_frames: list[pd.DataFrame] = []
    for sheet_name, target_columns in SHEETS.items():
        frame = pd.read_excel(workbook_path, sheet_name=sheet_name, dtype=object)
        summary_frames.append(
            _duplicate_groups(
                frame,
                sheet_name=sheet_name,
                target_columns=target_columns,
            )
        )
    duplicates = pd.concat(summary_frames, ignore_index=True) if summary_frames else pd.DataFrame()
    output_csv_path.parent.mkdir(parents=True, exist_ok=True)
    duplicates.to_csv(output_csv_path, index=False)
    return duplicates


def drop_exact_duplicate_rows(
    *,
    mapping_workbook_path: Path = MAPPING_WORKBOOK_PATH,
    duplicate_mappings_csv_path: Path = DUPLICATE_MAPPINGS_CSV_PATH,
) -> dict[str, object]:
    if not mapping_workbook_path.exists():
        raise FileNotFoundError(f"Missing mapping workbook: {mapping_workbook_path}")
    _assert_not_open(mapping_workbook_path)

    backup_path = _backup_workbook(mapping_workbook_path)

    current_duplicate_groups: dict[str, pd.DataFrame] = {}
    rows_to_delete: dict[str, list[int]] = {}

    for sheet_name, target_columns in SHEETS.items():
        frame = pd.read_excel(mapping_workbook_path, sheet_name=sheet_name, dtype=object).fillna("")
        duplicates = _duplicate_groups(frame, sheet_name=sheet_name, target_columns=target_columns)
        current_duplicate_groups[sheet_name] = duplicates
        _raise_on_differences(frame, sheet_name=sheet_name, target_columns=target_columns)

        if duplicates.empty:
            rows_to_delete[sheet_name] = []
            continue

        delete_rows: list[int] = []
        for _, group in duplicates.groupby([*SOURCE_COLUMNS, *target_columns], sort=False):
            row_numbers = sorted(int(row_number) for row_number in group["mapping_row_number"].tolist())
            delete_rows.extend(row_numbers[1:])
        rows_to_delete[sheet_name] = sorted(delete_rows)

    report_rows = []
    for frame in current_duplicate_groups.values():
        if not frame.empty:
            report_rows.append(frame)
    if report_rows:
        duplicate_report_before = pd.concat(report_rows, ignore_index=True)
    else:
        duplicate_report_before = pd.DataFrame()

    if duplicate_mappings_csv_path.exists():
        previous_report = pd.read_csv(duplicate_mappings_csv_path, dtype=object).fillna("")
    else:
        previous_report = pd.DataFrame()

    if not previous_report.empty and not duplicate_report_before.empty:
        previous_signature = _duplicate_group_signature(previous_report)
        current_signature = _duplicate_group_signature(duplicate_report_before)
        if not previous_signature.equals(current_signature):
            raise ValueError(
                "The provided duplicate_mappings_csv does not match the current workbook duplicate set. "
                "No workbook changes were made."
            )

    for sheet_name, delete_rows in rows_to_delete.items():
        _delete_rows_from_sheet(mapping_workbook_path, sheet_name=sheet_name, row_numbers=delete_rows)

    refreshed_report = _refresh_duplicate_report(mapping_workbook_path, duplicate_mappings_csv_path)

    return {
        "mapping_workbook": str(mapping_workbook_path),
        "backup_workbook": str(backup_path),
        "deleted_rows_esto": len(rows_to_delete.get("leap_combined_esto", [])),
        "deleted_rows_ninth": len(rows_to_delete.get("leap_combined_ninth", [])),
        "duplicate_mappings_before": int(len(duplicate_report_before)),
        "duplicate_mappings_after": int(len(refreshed_report)),
        "duplicate_mappings_csv": str(duplicate_mappings_csv_path),
    }


#%%
RUN_WORKFLOW = True
WORKFLOW_RESULT: dict[str, object] | None = None
if RUN_WORKFLOW:
    WORKFLOW_RESULT = drop_exact_duplicate_rows()
    print("[OK] Exact duplicate rows removed from LEAP mapping workbook.")
    for key, value in WORKFLOW_RESULT.items():
        print(f"{key}: {value}")
