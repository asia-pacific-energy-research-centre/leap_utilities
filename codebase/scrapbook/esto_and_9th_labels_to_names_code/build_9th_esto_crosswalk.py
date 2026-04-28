#%%
# Build a 9th -> ESTO crosswalk using code prefixes and name matches.
# Output is a candidate mapping for review (many-to-one and one-to-many allowed).
#%%
from __future__ import annotations

from pathlib import Path
import re

import pandas as pd

from codebase.utilities.master_config import read_config_table


REPO_ROOT = Path(__file__).resolve().parents[3]
WORKBOOK_PATH = REPO_ROOT / "config" / "sector_fuel_codes_to_names.xlsx"
OUTPUT_PATH = REPO_ROOT / "outputs" / "ninth_to_esto_crosswalk_candidates.csv"
REPORT_PATH = REPO_ROOT / "outputs" / "ninth_to_esto_crosswalk_report.txt"

NINTH_COLUMNS = [
    "sectors",
    "sub1sectors",
    "sub2sectors",
    "sub3sectors",
    "sub4sectors",
    "fuels",
    "subfuels",
]
ESTO_COLUMNS = ["flows", "products"]

NINTH_CODE_PATTERN = re.compile(r"^\d+(?:_\d+)*")
ESTO_CODE_PATTERN = re.compile(r"^\d+(?:\.\d+)*")


def _read_sheet(workbook_path: Path, sheet_name: str) -> pd.DataFrame:
    return read_config_table(workbook_path, sheet_name=sheet_name, dtype=str).fillna("")


def _extract_ninth_code(label: str) -> str:
    match = NINTH_CODE_PATTERN.match(label or "")
    return match.group(0) if match else ""


def _extract_esto_code(label: str) -> str:
    match = ESTO_CODE_PATTERN.match(label or "")
    return match.group(0) if match else ""


def _build_label_column_map(df: pd.DataFrame, columns: list[str]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for column in columns:
        if column not in df.columns:
            continue
        for raw_label in df[column].dropna().astype(str):
            label = raw_label.strip()
            if not label:
                continue
            mapping.setdefault(label, column)
    return mapping


def _build_name_maps(code_to_name_df: pd.DataFrame) -> tuple[dict[str, str], dict[str, str]]:
    ninth_map: dict[str, str] = {}
    esto_map: dict[str, str] = {}
    for _, row in code_to_name_df.iterrows():
        name = str(row.get("name") or "").strip()
        if not name:
            continue
        ninth_label = str(row.get("9th_label") or "").strip()
        if ninth_label:
            ninth_map.setdefault(ninth_label, name)
        esto_label = str(row.get("esto_label") or "").strip()
        if esto_label:
            esto_map.setdefault(esto_label, name)
    return ninth_map, esto_map


def _invert_name_map(name_map: dict[str, str]) -> dict[str, list[str]]:
    inv: dict[str, list[str]] = {}
    for label, name in name_map.items():
        inv.setdefault(name, []).append(label)
    return inv


def _build_esto_code_index(esto_labels: list[str]) -> dict[str, list[str]]:
    index: dict[str, list[str]] = {}
    for label in esto_labels:
        code = _extract_esto_code(label)
        if not code:
            continue
        index.setdefault(code, []).append(label)
    return index


def _find_code_matches(
    ninth_label: str,
    esto_code_index: dict[str, list[str]],
    esto_label_map: dict[str, str],
    allowed_esto_columns: set[str] | None,
) -> list[tuple[str, str, str]]:
    code = _extract_ninth_code(ninth_label)
    if not code:
        return []
    dotted = code.replace("_", ".")
    segments = dotted.split(".")
    for i in range(len(segments), 0, -1):
        candidate = ".".join(segments[:i])
        if candidate in esto_code_index:
            if i == len(segments):
                method = "direct_code_match"
            else:
                method = f"parent_code_match_{len(segments) - i}_levels_up"
            labels = esto_code_index[candidate]
            if allowed_esto_columns:
                labels = [
                    label
                    for label in labels
                    if esto_label_map.get(label, "") in allowed_esto_columns
                ]
            if labels:
                return [(label, method, candidate) for label in labels]
    return []


def build_crosswalk(workbook_path: Path) -> tuple[pd.DataFrame, str]:
    ninth_df = _read_sheet(workbook_path, "9th")
    esto_df = _read_sheet(workbook_path, "ESTO")
    code_to_name_df = _read_sheet(workbook_path, "code_to_name")

    ninth_label_map = _build_label_column_map(ninth_df, NINTH_COLUMNS)
    esto_label_map = _build_label_column_map(esto_df, ESTO_COLUMNS)
    ninth_name_map, esto_name_map = _build_name_maps(code_to_name_df)
    esto_name_to_labels = _invert_name_map(esto_name_map)

    ninth_labels = sorted(ninth_label_map)
    esto_labels = sorted(esto_label_map)
    esto_code_index = _build_esto_code_index(esto_labels)

    rows: list[dict] = []
    unmapped: list[str] = []

    for ninth_label in ninth_labels:
        ninth_name = ninth_name_map.get(ninth_label, "")
        ninth_column = ninth_label_map.get(ninth_label, "")
        if ninth_column in {"fuels", "subfuels"}:
            allowed_esto_columns = {"products"}
        elif ninth_column in {"sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"}:
            allowed_esto_columns = {"flows"}
        else:
            allowed_esto_columns = None

        matches = _find_code_matches(
            ninth_label,
            esto_code_index,
            esto_label_map,
            allowed_esto_columns,
        )
        seen_labels = {m[0] for m in matches}

        if ninth_name:
            for esto_label in esto_name_to_labels.get(ninth_name, []):
                if allowed_esto_columns and esto_label_map.get(esto_label, "") not in allowed_esto_columns:
                    continue
                if esto_label in seen_labels:
                    continue
                matches.append((esto_label, "name_match", _extract_esto_code(esto_label)))
                seen_labels.add(esto_label)

        if not matches:
            unmapped.append(ninth_label)
            continue

        for esto_label, method, code_match in matches:
            rows.append(
                {
                    "9th_label": ninth_label,
                    "9th_column": ninth_column,
                    "9th_code": _extract_ninth_code(ninth_label),
                    "9th_name": ninth_name,
                    "esto_label": esto_label,
                    "esto_column": esto_label_map.get(esto_label, ""),
                    "esto_code": _extract_esto_code(esto_label),
                    "esto_name": esto_name_map.get(esto_label, ""),
                    "match_method": method,
                    "match_code": code_match,
                }
            )

    crosswalk = pd.DataFrame(rows)
    if not crosswalk.empty:
        crosswalk = crosswalk.sort_values(
            ["9th_label", "match_method", "esto_label"], ascending=[True, True, True]
        )

    report_lines = []
    report_lines.append(f"9th labels: {len(ninth_labels)}")
    report_lines.append(f"ESTO labels: {len(esto_labels)}")
    report_lines.append(f"Mapped 9th labels: {crosswalk['9th_label'].nunique() if not crosswalk.empty else 0}")
    report_lines.append(f"Unmapped 9th labels: {len(unmapped)}")
    if unmapped:
        report_lines.append("Unmapped 9th labels:")
        report_lines.extend(f"  - {label}" for label in unmapped)

    if not crosswalk.empty:
        multi = (
            crosswalk.groupby("9th_label")["esto_label"]
            .nunique()
            .sort_values(ascending=False)
        )
        ambiguous = multi[multi > 1]
        report_lines.append(f"Ambiguous 9th labels (multiple ESTO matches): {len(ambiguous)}")
        if len(ambiguous):
            report_lines.append("Top ambiguous examples:")
            for label in ambiguous.head(25).index:
                esto_targets = (
                    crosswalk[crosswalk["9th_label"] == label]["esto_label"]
                    .dropna()
                    .unique()
                    .tolist()
                )
                report_lines.append(f"  - {label} -> {', '.join(esto_targets)}")

    return crosswalk, "\n".join(report_lines)


def main() -> None:
    crosswalk, report = build_crosswalk(WORKBOOK_PATH)
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    crosswalk.to_csv(OUTPUT_PATH, index=False)
    REPORT_PATH.write_text(report, encoding="utf-8")
    print(f"Wrote mapping candidates to: {OUTPUT_PATH}")
    print(f"Wrote report to: {REPORT_PATH}")


if __name__ == "__main__":
    main()
