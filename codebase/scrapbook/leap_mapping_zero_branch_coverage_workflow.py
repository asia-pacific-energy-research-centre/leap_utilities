#%%
"""
Prototype: find LEAP branch paths that are missing from balance mappings.

This checks the full LEAP model export, including branches with zero outputs,
against the explicit LEAP balance mapping sheets. It writes missing candidate
rows in the same column shape as `leap_combined_esto` and `leap_combined_ninth`
so they can be filled and pasted back into `config/leap_mappings.xlsx`.
"""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path
from typing import Iterable

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.utilities.leap_results_dashboard_balance import (  # noqa: E402
    BALANCE_ESTO_MAPPING_COLUMNS,
    BALANCE_NINTH_MAPPING_COLUMNS,
)
from codebase.utilities.workflow_outputs import write_output_manifest  # noqa: E402


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


FULL_MODEL_EXPORT_PATH = _resolve(r"C:\Users\Work\github\leap_utilities\data\full model export.xlsx")
MAPPING_WORKBOOK_PATH = _resolve("config/leap_mappings.xlsx")
OUTPUT_DIR = _resolve("outputs/mappings/leap_mapping_zero_branch_coverage")
OUTPUT_WORKBOOK_PATH = OUTPUT_DIR / "missing_zero_branch_mapping_candidates.xlsx"

EXPORT_SHEET_NAME = "Export"
MAPPING_ESTO_SHEET = "leap_combined_esto"
MAPPING_NINTH_SHEET = "leap_combined_ninth"

DEMAND_ROOT = "Demand"
TRANSFORMATION_ROOT = "Transformation"
RESOURCES_ROOT = "Resources"

DEMAND_ENERGY_VARIABLES = {
    "final energy intensity",
    "total energy",
}
TRANSFORMATION_FUEL_MARKERS = {
    "Output Fuels",
    "Feedstock Fuels",
    "Auxiliary Fuels",
}
SUPPLY_BALANCE_SECTORS = [
    "Production",
    "Imports",
    "Exports",
    "Total Primary Supply",
]
IGNORED_FUEL_LIKE_LEAVES = {
    "carbon dioxide",
    "carbon monoxide",
    "methane",
    "non methane volatile organic compounds",
    "nitrogen oxides",
    "nitrous oxide",
    "sulfur dioxide",
}


#%%
def _clean(value: object) -> str:
    text = str(value or "").strip()
    return "" if text.lower() in {"", "nan", "none"} else text


def _norm(value: object) -> str:
    return " ".join(_clean(value).lower().split())


def _split_branch_path(path: object) -> list[str]:
    return [part.strip() for part in _clean(path).replace("/", "\\").split("\\") if part.strip()]


def _join_path(parts: Iterable[str]) -> str:
    return "/".join(_clean(part) for part in parts if _clean(part))


def _truthy(value: object) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "yes", "y", "on", "t"}


def _read_leap_export(path: Path, sheet_name: str = EXPORT_SHEET_NAME) -> pd.DataFrame:
    preview = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=20)
    header_row = None
    for idx, row in preview.iterrows():
        values = {_clean(value) for value in row.tolist()}
        if "Branch Path" in values and "Variable" in values:
            header_row = int(idx)
            break
    if header_row is None:
        raise ValueError(f"Could not find a header row containing Branch Path and Variable in {path}.")
    df = pd.read_excel(path, sheet_name=sheet_name, header=header_row, dtype=str).fillna("")
    df = df[df.get("Branch Path", "").astype(str).str.strip().ne("")].copy()
    return df


def _mapping_key_set(mapping: pd.DataFrame) -> tuple[set[tuple[str, str]], set[tuple[str, str]]]:
    work = mapping.copy()
    for col in ["leap_sector_name_full_path", "raw_leap_fuel_name", "remove_row"]:
        if col not in work.columns:
            work[col] = ""
    work["key"] = list(
        zip(
            work["leap_sector_name_full_path"].map(_norm),
            work["raw_leap_fuel_name"].map(_norm),
        )
    )
    all_keys = {key for key in work["key"] if key[0] and key[1]}
    active = work[~work["remove_row"].map(_truthy)].copy()
    active_keys = {key for key in active["key"] if key[0] and key[1]}
    return all_keys, active_keys


def _known_mapping_fuels(mapping_sheets: dict[str, pd.DataFrame]) -> set[str]:
    """Return normalized LEAP fuel labels that are known to the explicit mapping workbook."""
    fuels: set[str] = set()
    for sheet_name, column in [
        (MAPPING_ESTO_SHEET, "raw_leap_fuel_name"),
        (MAPPING_NINTH_SHEET, "raw_leap_fuel_name"),
        ("fuel_product_final_proposed", "leap_fuel_name"),
        ("fuel_ninth_final_proposed", "leap_fuel_name"),
    ]:
        sheet = mapping_sheets.get(sheet_name)
        if sheet is None or column not in sheet.columns:
            continue
        fuels.update(_norm(value) for value in sheet[column].dropna().tolist() if _norm(value))
    return fuels


def _empty_candidate_frame() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "model_area",
            "branch_path",
            "inference_rule",
            "leap_sector_name_original",
            "leap_sector_name_full_path",
            "raw_leap_fuel_name",
        ]
    )


def _infer_demand_candidates(
    export_df: pd.DataFrame,
    *,
    known_fuels: set[str],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    work = export_df.copy()
    work["variable_norm"] = work["Variable"].map(_norm)
    work = work[work["variable_norm"].isin(DEMAND_ENERGY_VARIABLES)].copy()
    rows: list[dict[str, object]] = []
    excluded_rows: list[dict[str, object]] = []
    for branch_path in sorted(work["Branch Path"].drop_duplicates().astype(str)):
        parts = _split_branch_path(branch_path)
        if len(parts) < 3 or parts[0] != DEMAND_ROOT:
            continue
        leaf = parts[-1]
        if _norm(leaf) in IGNORED_FUEL_LIKE_LEAVES:
            continue
        sector_parts = parts[1:-1]
        if not sector_parts:
            continue
        if _norm(leaf) not in known_fuels:
            excluded_rows.append(
                {
                    "model_area": "demand",
                    "branch_path": branch_path,
                    "inference_rule": (
                        "Excluded demand leaf because it is not a known LEAP fuel label in "
                        "leap_combined_* or fuel_*_final_proposed sheets"
                    ),
                    "leap_sector_name_original": sector_parts[-1],
                    "leap_sector_name_full_path": _join_path(sector_parts),
                    "raw_leap_fuel_name": leaf,
                }
            )
            continue
        rows.append(
            {
                "model_area": "demand",
                "branch_path": branch_path,
                "inference_rule": (
                    "Demand energy-variable leaf with known fuel label: "
                    "sector=parent path after Demand, fuel=leaf branch"
                ),
                "leap_sector_name_original": sector_parts[-1],
                "leap_sector_name_full_path": _join_path(sector_parts),
                "raw_leap_fuel_name": leaf,
            }
        )
    candidates = pd.DataFrame(rows) if rows else _empty_candidate_frame()
    excluded = pd.DataFrame(excluded_rows) if excluded_rows else _empty_candidate_frame()
    return candidates, excluded


def _infer_transformation_candidates(export_df: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    for branch_path in sorted(export_df["Branch Path"].drop_duplicates().astype(str)):
        parts = _split_branch_path(branch_path)
        if len(parts) < 4 or parts[0] != TRANSFORMATION_ROOT:
            continue
        marker_index = next((idx for idx, part in enumerate(parts) if part in TRANSFORMATION_FUEL_MARKERS), None)
        if marker_index is None or marker_index >= len(parts) - 1:
            continue
        fuel = parts[marker_index + 1]
        if _norm(fuel) in IGNORED_FUEL_LIKE_LEAVES:
            continue
        plant = parts[1]
        if parts[marker_index] == "Output Fuels":
            sector_parts = [plant]
            rule = "Transformation output fuel branch: sector=plant, fuel=output fuel"
        else:
            process = parts[marker_index - 1] if marker_index >= 1 else ""
            sector_parts = [plant] + ([process] if process and process != plant else [])
            rule = f"Transformation {parts[marker_index]} branch: sector=plant/process, fuel=fuel child"
        rows.append(
            {
                "model_area": "transformation",
                "branch_path": branch_path,
                "inference_rule": rule,
                "leap_sector_name_original": sector_parts[-1],
                "leap_sector_name_full_path": _join_path(sector_parts),
                "raw_leap_fuel_name": fuel,
            }
        )
    return pd.DataFrame(rows)


def _infer_supply_candidates(export_df: pd.DataFrame) -> pd.DataFrame:
    fuels: set[str] = set()
    source_paths: dict[str, list[str]] = {}
    for branch_path in sorted(export_df["Branch Path"].drop_duplicates().astype(str)):
        parts = _split_branch_path(branch_path)
        if len(parts) != 3 or parts[0] != RESOURCES_ROOT:
            continue
        fuel = parts[-1]
        if _norm(fuel) in IGNORED_FUEL_LIKE_LEAVES:
            continue
        fuels.add(fuel)
        source_paths.setdefault(fuel, []).append(branch_path)

    rows: list[dict[str, object]] = []
    for fuel in sorted(fuels, key=str.lower):
        for sector in SUPPLY_BALANCE_SECTORS:
            rows.append(
                {
                    "model_area": "supply",
                    "branch_path": "|".join(source_paths.get(fuel, [])),
                    "inference_rule": (
                        "Resources fuel branch: candidate generated for core supply balance rows "
                        "(Production, Imports, Exports, Total Primary Supply)"
                    ),
                    "leap_sector_name_original": sector,
                    "leap_sector_name_full_path": sector,
                    "raw_leap_fuel_name": fuel,
                }
            )
    return pd.DataFrame(rows)


def _candidate_template(columns: list[str], candidates: pd.DataFrame, *, target: str) -> pd.DataFrame:
    out = pd.DataFrame(columns=columns)
    if candidates.empty:
        return out
    out["leap_sector_name_original"] = candidates["leap_sector_name_original"]
    out["leap_sector_name_full_path"] = candidates["leap_sector_name_full_path"]
    out["raw_leap_fuel_name"] = candidates["raw_leap_fuel_name"]
    out["value"] = 0.0
    out["pair_mapping_cardinality"] = ""
    out["leap_is_subtotal"] = False
    out["subtotal_mismatch_is_ok"] = False
    out["subtotal_alignment"] = ""
    out["many_to_many_is_ok"] = False
    out["remove_row"] = False
    out["remove_row_reason"] = ""
    present_any_col = f"{target}_mapping_present_any"
    present_active_col = f"{target}_mapping_present_active"
    if present_any_col in candidates.columns and present_active_col in candidates.columns:
        removed_only = candidates[present_any_col].fillna(False).astype(bool) & ~candidates[
            present_active_col
        ].fillna(False).astype(bool)
        sheet_name = "leap_combined_esto" if target == "esto" else "leap_combined_ninth"
        out.loc[removed_only.to_numpy(), "remove_row"] = True
        out.loc[removed_only.to_numpy(), "remove_row_reason"] = (
            f"this row exists in the {sheet_name} mapping but has its remove_row set to true "
            "so it is not available"
        )
    if target == "esto":
        out["esto_flow"] = ""
        out["esto_product"] = ""
        out["esto_pair_is_subtotal"] = False
        out["esto_pair_abs_sum"] = ""
    else:
        out["ninth_sector"] = ""
        out["ninth_fuel"] = ""
        out["ninth_pair_is_subtotal"] = False
        out["ninth_pair_abs_sum"] = ""
    return out[columns]


def build_zero_branch_mapping_coverage(
    *,
    export_path: Path = FULL_MODEL_EXPORT_PATH,
    mapping_workbook_path: Path = MAPPING_WORKBOOK_PATH,
    output_workbook_path: Path = OUTPUT_WORKBOOK_PATH,
) -> dict[str, object]:
    export_df = _read_leap_export(export_path)
    mapping_sheets = pd.read_excel(mapping_workbook_path, sheet_name=None, dtype=str)
    mapping_sheets = {name: frame.fillna("") for name, frame in mapping_sheets.items()}
    esto_mapping = mapping_sheets[MAPPING_ESTO_SHEET]
    ninth_mapping = mapping_sheets[MAPPING_NINTH_SHEET]
    esto_all_keys, esto_active_keys = _mapping_key_set(esto_mapping)
    ninth_all_keys, ninth_active_keys = _mapping_key_set(ninth_mapping)
    known_fuels = _known_mapping_fuels(mapping_sheets)
    demand_candidates, excluded_demand_leaf_candidates = _infer_demand_candidates(
        export_df,
        known_fuels=known_fuels,
    )

    candidates = pd.concat(
        [
            demand_candidates,
            _infer_transformation_candidates(export_df),
            _infer_supply_candidates(export_df),
        ],
        ignore_index=True,
        sort=False,
    )
    if candidates.empty:
        candidates = pd.DataFrame(
            columns=[
                "model_area",
                "branch_path",
                "inference_rule",
                "leap_sector_name_original",
                "leap_sector_name_full_path",
                "raw_leap_fuel_name",
            ]
        )
    candidates = candidates.drop_duplicates(
        subset=["leap_sector_name_full_path", "raw_leap_fuel_name", "model_area"],
        keep="first",
    ).reset_index(drop=True)
    candidates["mapping_key"] = list(
        zip(
            candidates["leap_sector_name_full_path"].map(_norm),
            candidates["raw_leap_fuel_name"].map(_norm),
        )
    )
    candidates["esto_mapping_present_any"] = candidates["mapping_key"].isin(esto_all_keys)
    candidates["esto_mapping_present_active"] = candidates["mapping_key"].isin(esto_active_keys)
    candidates["ninth_mapping_present_any"] = candidates["mapping_key"].isin(ninth_all_keys)
    candidates["ninth_mapping_present_active"] = candidates["mapping_key"].isin(ninth_active_keys)
    candidates = candidates.sort_values(
        ["model_area", "leap_sector_name_full_path", "raw_leap_fuel_name"],
        kind="mergesort",
    ).reset_index(drop=True)

    missing_esto = candidates[~candidates["esto_mapping_present_active"]].copy()
    missing_ninth = candidates[~candidates["ninth_mapping_present_active"]].copy()
    esto_missing_paths = _candidate_template(BALANCE_ESTO_MAPPING_COLUMNS, missing_esto, target="esto")
    ninth_missing_paths = _candidate_template(BALANCE_NINTH_MAPPING_COLUMNS, missing_ninth, target="ninth")

    summary = pd.DataFrame(
        [
            {"metric": "export_rows", "value": int(len(export_df))},
            {"metric": "known_mapping_fuels", "value": int(len(known_fuels))},
            {"metric": "excluded_demand_leaf_candidates", "value": int(len(excluded_demand_leaf_candidates))},
            {"metric": "candidate_branch_pairs", "value": int(len(candidates))},
            {"metric": "missing_active_esto_pairs", "value": int(len(missing_esto))},
            {"metric": "missing_active_ninth_pairs", "value": int(len(missing_ninth))},
            {
                "metric": "present_only_as_removed_esto_pairs",
                "value": int((candidates["esto_mapping_present_any"] & ~candidates["esto_mapping_present_active"]).sum()),
            },
            {
                "metric": "present_only_as_removed_ninth_pairs",
                "value": int((candidates["ninth_mapping_present_any"] & ~candidates["ninth_mapping_present_active"]).sum()),
            },
        ]
    )

    output_workbook_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_workbook_path) as writer:
        esto_missing_paths.to_excel(writer, sheet_name="esto_missing_paths", index=False)
        ninth_missing_paths.to_excel(writer, sheet_name="ninth_missing_paths", index=False)
        candidates.drop(columns=["mapping_key"]).to_excel(writer, sheet_name="all_branch_candidates", index=False)
        excluded_demand_leaf_candidates.to_excel(
            writer,
            sheet_name="excluded_demand_leaf_candidates",
            index=False,
        )
        summary.to_excel(writer, sheet_name="summary", index=False)
    manifest_path = write_output_manifest(
        out_dir=output_workbook_path.parent,
        primary_outputs={"output_workbook": str(output_workbook_path)},
        supporting_outputs={},
        primary_output_descriptions={
            "output_workbook": "Workbook of zero-branch balance-mapping candidates and audit sheets.",
        },
        notes=[
            "This workflow produces a single primary workbook.",
        ],
    )

    return {
        "output_workbook": str(output_workbook_path),
        "summary": summary,
        "candidate_branch_pairs": int(len(candidates)),
        "excluded_demand_leaf_candidates": int(len(excluded_demand_leaf_candidates)),
        "missing_active_esto_pairs": int(len(missing_esto)),
        "missing_active_ninth_pairs": int(len(missing_ninth)),
        "output_manifest_json": str(manifest_path),
    }


#%%
RUN_WORKFLOW = True
WORKFLOW_RESULT: dict[str, object] | None = None
if RUN_WORKFLOW:
    WORKFLOW_RESULT = build_zero_branch_mapping_coverage()
    print("[OK] Zero-branch mapping coverage prototype complete.")
    for key, value in WORKFLOW_RESULT.items():
        if isinstance(value, pd.DataFrame):
            print(f"{key}:")
            print(value.to_string(index=False))
        else:
            print(f"{key}: {value}")
