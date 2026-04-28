from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from codebase.utilities.master_config import read_config_table
from codebase.functions.leap_excel_io import read_export_sheet, write_export_sheet
from codebase.functions.leap_labels import clean_fuel_label_for_leap


@dataclass(frozen=True)
class MappingRow:
    source_fuel: str
    ninth_fuel: str
    esto_product_override: str
    notes: str


def _normalize_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _load_mapping(mapping_csv_path: Path) -> dict[str, MappingRow]:
    df = read_config_table(mapping_csv_path).fillna("")
    rows: dict[str, MappingRow] = {}
    for _, row in df.iterrows():
        source_fuel = _normalize_text(row.get("source_fuel"))
        if not source_fuel:
            continue
        ninth_fuel = _normalize_text(row.get("ninth_fuel"))
        override = _normalize_text(row.get("esto_product_override"))
        notes = _normalize_text(row.get("notes"))
        rows[source_fuel] = MappingRow(
            source_fuel=source_fuel,
            ninth_fuel=ninth_fuel,
            esto_product_override=override,
            notes=notes,
        )
    return rows


def _load_pairs(pairs_path: Path) -> pd.DataFrame:
    df = read_config_table(pairs_path)
    df = df.fillna("")
    df["9th_fuel"] = df["9th_fuel"].astype(str).str.strip()
    df["esto_product"] = df["esto_product"].astype(str).str.strip()
    return df


def _resolve_esto_product(
    pairs_df: pd.DataFrame,
    ninth_fuel: str,
    override: str,
) -> tuple[str | None, list[str]]:
    if override:
        return override, []
    subset = pairs_df[pairs_df["9th_fuel"] == ninth_fuel]
    products = [p for p in subset["esto_product"].tolist() if p]
    unique_products = sorted(set(products))
    if len(unique_products) == 1:
        return unique_products[0], []
    return None, unique_products


def _clean_levels(row: pd.Series, level_cols: list[str]) -> list[str]:
    return [_normalize_text(row.get(col)) for col in level_cols]


def _find_fuel_group(levels: list[str], fuel_group_labels: tuple[str, ...]) -> tuple[int, str] | None:
    for idx, value in enumerate(levels):
        if value in fuel_group_labels:
            return idx, value
    return None


def _update_branch_path(row: pd.Series, level_cols: list[str], fuel_idx: int, new_fuel: str) -> pd.Series:
    levels = _clean_levels(row, level_cols)
    if fuel_idx >= len(levels):
        return row
    levels[fuel_idx] = clean_fuel_label_for_leap(new_fuel)
    branch_path = "\\".join([val for val in levels if val])
    updated = row.copy()
    updated["Branch Path"] = branch_path
    for col, value in zip(level_cols, levels):
        updated[col] = value if value else pd.NA
    return updated


def _is_fuel_leaf(
    row: pd.Series,
    level_cols: list[str],
    branch_root: str,
    fuel_group_labels: tuple[str, ...],
) -> tuple[bool, str, int]:
    branch_path = _normalize_text(row.get("Branch Path"))
    if not branch_path or not branch_path.startswith(branch_root):
        return False, "", -1
    levels = _clean_levels(row, level_cols)
    group_info = _find_fuel_group(levels, fuel_group_labels)
    if group_info is None:
        return False, "", -1
    group_idx, _group_name = group_info
    fuel_idx = group_idx + 1
    if fuel_idx >= len(levels):
        return False, "", -1
    fuel = levels[fuel_idx]
    if not fuel:
        return False, "", -1
    last_non_empty = max((idx for idx, val in enumerate(levels) if val), default=-1)
    if last_non_empty != fuel_idx:
        return False, "", -1
    return True, fuel, fuel_idx


def remap_transformation_export_fuels(
    input_path: str | Path,
    output_path: str | Path,
    mapping_csv_path: str | Path,
    pairs_path: str | Path,
    sheet_name: str = "Export",
    branch_root: str = "Transformation\\Oil Refining",
    fuel_group_labels: tuple[str, ...] = ("Output Fuels", "Feedstock Fuels", "Auxiliary Fuels"),
    clear_id_columns: bool = True,
    report_path: str | Path | None = None,
) -> dict[str, object]:
    input_path = Path(input_path)
    output_path = Path(output_path)
    mapping_csv_path = Path(mapping_csv_path)
    pairs_path = Path(pairs_path)

    header_rows, df, columns = read_export_sheet(input_path, sheet_name)
    mapping_by_fuel = _load_mapping(mapping_csv_path)
    pairs_df = _load_pairs(pairs_path)

    level_cols = [col for col in columns if str(col).startswith("Level ")]
    id_cols = [col for col in ["BranchID", "VariableID", "ScenarioID", "RegionID"] if col in columns]

    report = {
        "unmapped_fuels": [],
        "ambiguous_mappings": [],
        "resolved_mappings": [],
    }

    new_rows = []
    for _, row in df.iterrows():
        is_fuel, fuel, fuel_idx = _is_fuel_leaf(
            row,
            level_cols,
            branch_root=branch_root,
            fuel_group_labels=fuel_group_labels,
        )
        if not is_fuel:
            new_rows.append(row)
            continue

        mapping_row = mapping_by_fuel.get(fuel)
        if not mapping_row:
            report["unmapped_fuels"].append(fuel)
            new_rows.append(row)
            continue

        esto_product, candidates = _resolve_esto_product(
            pairs_df,
            mapping_row.ninth_fuel,
            mapping_row.esto_product_override,
        )
        if not esto_product:
            report["ambiguous_mappings"].append(
                {
                    "source_fuel": fuel,
                    "ninth_fuel": mapping_row.ninth_fuel,
                    "candidates": candidates,
                }
            )
            new_rows.append(row)
            continue

        updated = _update_branch_path(row, level_cols, fuel_idx, esto_product)
        if clear_id_columns:
            for col in id_cols:
                updated[col] = pd.NA
        report["resolved_mappings"].append(
            {
                "source_fuel": fuel,
                "ninth_fuel": mapping_row.ninth_fuel,
                "esto_product": esto_product,
            }
        )
        new_rows.append(updated)

    mapped_df = pd.DataFrame([row.tolist() for row in new_rows], columns=columns)

    write_export_sheet(
        path=output_path,
        sheet_name=sheet_name,
        header_rows=header_rows,
        columns=columns,
        data=mapped_df,
    )

    if report_path:
        report_path = Path(report_path)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        rows = []
        for fuel in sorted(set(report["unmapped_fuels"])):
            rows.append({"issue_type": "unmapped_fuel", "detail": fuel})
        for item in report["ambiguous_mappings"]:
            rows.append(
                {
                    "issue_type": "ambiguous_mapping",
                    "detail": item.get("source_fuel"),
                    "ninth_fuel": item.get("ninth_fuel"),
                    "candidates": ";".join(item.get("candidates") or []),
                }
            )
        if rows:
            pd.DataFrame(rows).to_csv(report_path, index=False)

    return report


__all__ = ["remap_transformation_export_fuels"]
