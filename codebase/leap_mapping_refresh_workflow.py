#%%

"""
Refresh maintenance columns in config/leap_mappings.xlsx.

This workflow recomputes the lightweight audit columns used to maintain
`leap_combined_esto` and `leap_combined_ninth` without rerunning the full
dashboard process.
"""

from __future__ import annotations

import os
import re
import shutil
import sys
from pathlib import Path
from typing import Sequence

import pandas as pd

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
ESTO_TABLE_PATH = _resolve("data/00APEC_2025_low_with_subtotals.csv")
NINTH_TABLE_PATH = _resolve("data/merged_file_energy_ALL_20251106.csv")

ESTO_SHEET = "leap_combined_esto"
NINTH_SHEET = "leap_combined_ninth"

BASE_YEAR = 2022
PROJECTION_YEARS: Sequence[int] = tuple(range(2023, 2061))
PROJECTION_SCENARIOS: Sequence[str] = ("reference", "target")


#%%
def _clean(value: object) -> str:
    text = str(value or "").strip()
    return "" if text.lower() in {"", "nan", "none"} else text


def _norm_text(value: object) -> str:
    return " ".join(_clean(value).lower().split())


def _truthy(value: object) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "t", "yes", "y", "on"}


def _path_key(path: object) -> str:
    parts = [part.strip() for part in str(path or "").split("/") if part.strip()]
    return "/".join(_norm_text(part) for part in parts)


def _mapping_cardinality(source_target_count: int, target_source_count: int) -> str:
    if source_target_count <= 0 or target_source_count <= 0:
        return ""
    if source_target_count == 1 and target_source_count == 1:
        return "one_to_one"
    if source_target_count > 1 and target_source_count == 1:
        return "one_to_many"
    if source_target_count == 1 and target_source_count > 1:
        return "many_to_one"
    return "many_to_many"


def _subtotal_alignment(leap_is_subtotal: bool, target_is_subtotal: bool) -> str:
    if leap_is_subtotal and target_is_subtotal:
        return "aligned_subtotal"
    if (not leap_is_subtotal) and (not target_is_subtotal):
        return "aligned_non_subtotal"
    return "mismatch"


def _active_mask(frame: pd.DataFrame) -> pd.Series:
    remove_mask = frame.get("remove_row", False)
    duplicate_mask = frame.get("duplicate_to_remove", False)
    remove_mask = pd.Series(remove_mask, index=frame.index).map(_truthy)
    duplicate_mask = pd.Series(duplicate_mask, index=frame.index).map(_truthy)
    return ~(remove_mask | duplicate_mask)


def _drop_unnamed_columns(frame: pd.DataFrame) -> pd.DataFrame:
    keep_cols = [col for col in frame.columns if not str(col).startswith("Unnamed:")]
    return frame.loc[:, keep_cols].copy()


def _drop_columns_if_present(frame: pd.DataFrame, columns: Sequence[str]) -> pd.DataFrame:
    drop_cols = [col for col in columns if col in frame.columns]
    if not drop_cols:
        return frame.copy()
    return frame.drop(columns=drop_cols).copy()


def _reorder_columns(frame: pd.DataFrame, preferred_columns: Sequence[str]) -> pd.DataFrame:
    ordered = [col for col in preferred_columns if col in frame.columns]
    trailing = [col for col in frame.columns if col not in ordered]
    return frame.loc[:, ordered + trailing].copy()


def _compute_leap_subtotals(frame: pd.DataFrame) -> pd.DataFrame:
    out = frame.copy()
    for col in ["leap_sector_name_full_path", "raw_leap_fuel_name"]:
        if col not in out.columns:
            out[col] = ""
        out[col] = out[col].fillna("").astype(str).str.strip()

    active = _active_mask(out)
    active_paths = {
        _clean(value)
        for value in out.loc[active, "leap_sector_name_full_path"].tolist()
        if _clean(value)
    }

    def leap_sector_is_subtotal(path: object) -> bool:
        text = _clean(path)
        key = _path_key(text)
        if not key:
            return False
        if key.startswith("total "):
            return True
        prefix = f"{text}/"
        return any(other != text and other.startswith(prefix) for other in active_paths)

    def leap_fuel_is_subtotal(fuel: object) -> bool:
        key = _norm_text(fuel)
        return key == "total" or key.startswith("total ")

    leap_sector_is_subtotal = out["leap_sector_name_full_path"].map(leap_sector_is_subtotal)
    leap_fuel_is_subtotal = out["raw_leap_fuel_name"].map(leap_fuel_is_subtotal)
    out["leap_is_subtotal"] = leap_sector_is_subtotal.fillna(False).astype(bool) | leap_fuel_is_subtotal.fillna(False).astype(bool)
    return out


def _compute_pair_cardinality(frame: pd.DataFrame, target_sector_col: str, target_fuel_col: str) -> pd.DataFrame:
    """Compute cardinality of (leap_sector, leap_fuel) <-> (target_sector, target_fuel) pairs."""
    out = frame.copy()
    source_cols = ["leap_sector_name_full_path", "raw_leap_fuel_name"]
    target_cols = [target_sector_col, target_fuel_col]
    all_cols = source_cols + [c for c in target_cols if c not in source_cols]
    for col in all_cols:
        if col not in out.columns:
            out[col] = ""
        out[col] = out[col].fillna("").astype(str).str.strip()
    active = _active_mask(out)
    valid = (
        active
        & out["leap_sector_name_full_path"].ne("")
        & out["raw_leap_fuel_name"].ne("")
        & out[target_sector_col].ne("")
        & out[target_fuel_col].ne("")
    )
    pair_frame = out.loc[valid, source_cols + [target_sector_col, target_fuel_col]].copy()
    pair_frame["_source_key"] = pair_frame["leap_sector_name_full_path"] + "|||" + pair_frame["raw_leap_fuel_name"]
    pair_frame["_target_key"] = pair_frame[target_sector_col] + "|||" + pair_frame[target_fuel_col]
    pairs = pair_frame[["_source_key", "_target_key"]].drop_duplicates()
    source_count = pairs.groupby("_source_key")["_target_key"].nunique()
    target_count = pairs.groupby("_target_key")["_source_key"].nunique()
    out["_source_key"] = out["leap_sector_name_full_path"].fillna("").astype(str).str.strip() + "|||" + out["raw_leap_fuel_name"].fillna("").astype(str).str.strip()
    out["_target_key"] = out[target_sector_col].fillna("").astype(str).str.strip() + "|||" + out[target_fuel_col].fillna("").astype(str).str.strip()
    out["pair_mapping_cardinality"] = ""
    valid_rows = out["leap_sector_name_full_path"].ne("") & out["raw_leap_fuel_name"].ne("") & out[target_sector_col].ne("") & out[target_fuel_col].ne("")
    out.loc[valid_rows, "pair_mapping_cardinality"] = out.loc[valid_rows].apply(
        lambda row: _mapping_cardinality(
            int(source_count.get(row["_source_key"], 0)),
            int(target_count.get(row["_target_key"], 0)),
        ),
        axis=1,
    )
    out = out.drop(columns=["_source_key", "_target_key"])
    return out


def _load_esto_lookup() -> pd.DataFrame:
    base_df = pd.read_csv(ESTO_TABLE_PATH)
    work = base_df.copy()
    if "is_subtotal" not in work.columns:
        work["is_subtotal"] = False
    for col in ["economy", "flows", "products", str(BASE_YEAR), "is_subtotal"]:
        if col not in work.columns:
            work[col] = ""
    work["esto_flow"] = work["flows"].fillna("").astype(str).str.strip()
    work["esto_product"] = work["products"].fillna("").astype(str).str.strip()
    work["value"] = pd.to_numeric(work[str(BASE_YEAR)], errors="coerce").fillna(0.0)
    work["is_subtotal"] = work["is_subtotal"].fillna(False).map(_truthy)
    work = work[work["esto_flow"].ne("") & work["esto_product"].ne("")].copy()
    grouped = (
        work.groupby(["esto_flow", "esto_product"], as_index=False)
        .agg(
            pair_value_sum=("value", "sum"),
            esto_pair_is_subtotal=("is_subtotal", "max"),
        )
        .reset_index(drop=True)
    )
    grouped["esto_pair_abs_sum"] = grouped["pair_value_sum"].abs()
    return grouped


def _load_ninth_lookup() -> pd.DataFrame:
    ninth_df = pd.read_csv(NINTH_TABLE_PATH)
    work = ninth_df.copy()
    for col in [
        "economy",
        "scenarios",
        "sectors",
        "sub1sectors",
        "sub2sectors",
        "sub3sectors",
        "sub4sectors",
        "fuels",
        "subfuels",
        "subtotal_layout",
        "subtotal_results",
    ]:
        if col not in work.columns:
            work[col] = ""
    for col in ["subtotal_layout", "subtotal_results"]:
        work[col] = work[col].fillna(False).map(_truthy)
    scenario_set = {str(value).strip().lower() for value in PROJECTION_SCENARIOS}
    work = work[work["scenarios"].fillna("").astype(str).str.strip().str.lower().isin(scenario_set)].copy()
    year_cols = [str(year) for year in PROJECTION_YEARS if str(year) in work.columns]
    if not year_cols or work.empty:
        return pd.DataFrame(
            columns=[
                "ninth_sector",
                "ninth_fuel",
                "ninth_pair_is_subtotal",
                "ninth_pair_abs_sum",
            ]
        )
    values = work[year_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    work["ninth_sector"] = work.apply(
        lambda row: next(
            (
                _clean(row.get(col, ""))
                for col in ["sub4sectors", "sub3sectors", "sub2sectors", "sub1sectors", "sectors"]
                if _clean(row.get(col, ""))
            ),
            "",
        ),
        axis=1,
    )
    work["ninth_fuel"] = work.apply(
        lambda row: next(
            (
                _clean(row.get(col, ""))
                for col in ["subfuels", "fuels"]
                if _clean(row.get(col, ""))
            ),
            "",
        ),
        axis=1,
    )
    work["value_abs_sum_row"] = values.abs().sum(axis=1)
    work = work[work["ninth_sector"].ne("") & work["ninth_fuel"].ne("")].copy()
    grouped = (
        work.groupby(["ninth_sector", "ninth_fuel"], as_index=False)
        .agg(
            subtotal_layout=("subtotal_layout", "max"),
            subtotal_results=("subtotal_results", "max"),
            ninth_pair_abs_sum=("value_abs_sum_row", "sum"),
        )
        .reset_index(drop=True)
    )
    grouped["ninth_pair_is_subtotal"] = (
        grouped["subtotal_layout"].fillna(False).astype(bool)
        | grouped["subtotal_results"].fillna(False).astype(bool)
    )
    return grouped


def _refresh_esto_sheet(frame: pd.DataFrame, esto_lookup: pd.DataFrame) -> pd.DataFrame:
    out = _drop_unnamed_columns(frame)
    out = _drop_columns_if_present(
        out,
        [
            "esto_pair_is_subtotal",
            "esto_pair_is_subtotal_x",
            "esto_pair_is_subtotal_y",
            "esto_pair_abs_sum",
            "esto_pair_abs_sum_x",
            "esto_pair_abs_sum_y",
            "leap_sector_is_subtotal_computed",
            "leap_fuel_is_subtotal_computed",
        ],
    )
    out = _compute_leap_subtotals(out)
    out = _compute_pair_cardinality(out, "esto_flow", "esto_product")
    lookup = esto_lookup.copy()
    for col in ["esto_flow", "esto_product"]:
        if col not in out.columns:
            out[col] = ""
        out[col] = out[col].fillna("").astype(str).str.strip()
    out = out.merge(
        lookup[["esto_flow", "esto_product", "esto_pair_is_subtotal", "esto_pair_abs_sum"]],
        on=["esto_flow", "esto_product"],
        how="left",
    )
    if "esto_pair_is_subtotal" not in out.columns:
        out["esto_pair_is_subtotal"] = False
    out["esto_pair_is_subtotal"] = out["esto_pair_is_subtotal"].fillna(False).astype(bool)
    if "esto_pair_abs_sum" not in out.columns:
        out["esto_pair_abs_sum"] = 0.0
    out["esto_pair_abs_sum"] = pd.to_numeric(out["esto_pair_abs_sum"], errors="coerce").fillna(0.0)
    total_mask = out["esto_product"].fillna("").astype(str).str.strip().str.lower().eq("19 total")
    out.loc[total_mask, "esto_pair_is_subtotal"] = True
    out["subtotal_alignment"] = out.apply(
        lambda row: _subtotal_alignment(bool(row.get("leap_is_subtotal", False)), bool(row.get("esto_pair_is_subtotal", False))),
        axis=1,
    )
    return _reorder_columns(
        out,
        [
            "leap_sector_name_original",
            "leap_sector_name_full_path",
            "raw_leap_fuel_name",
            "value",
            "esto_flow",
            "esto_product",
            "pair_mapping_cardinality",
            "leap_is_subtotal",
            "esto_pair_is_subtotal",
            "subtotal_mismatch_is_ok",
            "subtotal_alignment",
            "esto_pair_abs_sum",
            "many_to_many_is_ok",
            "remove_row",
            "remove_row_reason",
        ],
    )


def _refresh_ninth_sheet(frame: pd.DataFrame, ninth_lookup: pd.DataFrame) -> pd.DataFrame:
    out = _drop_unnamed_columns(frame)
    out = _drop_columns_if_present(
        out,
        [
            "ninth_pair_is_subtotal",
            "ninth_pair_is_subtotal_x",
            "ninth_pair_is_subtotal_y",
            "ninth_pair_abs_sum",
            "ninth_pair_abs_sum_x",
            "ninth_pair_abs_sum_y",
            "leap_sector_is_subtotal_computed",
            "leap_fuel_is_subtotal_computed",
        ],
    )
    out = _compute_leap_subtotals(out)
    out = _compute_pair_cardinality(out, "ninth_sector", "ninth_fuel")
    lookup = ninth_lookup.copy()
    for col in ["ninth_sector", "ninth_fuel"]:
        if col not in out.columns:
            out[col] = ""
        out[col] = out[col].fillna("").astype(str).str.strip()
    out = out.merge(
        lookup[["ninth_sector", "ninth_fuel", "ninth_pair_is_subtotal", "ninth_pair_abs_sum"]],
        on=["ninth_sector", "ninth_fuel"],
        how="left",
    )
    if "ninth_pair_is_subtotal" not in out.columns:
        out["ninth_pair_is_subtotal"] = False
    out["ninth_pair_is_subtotal"] = out["ninth_pair_is_subtotal"].fillna(False).astype(bool)
    if "ninth_pair_abs_sum" not in out.columns:
        out["ninth_pair_abs_sum"] = 0.0
    out["ninth_pair_abs_sum"] = pd.to_numeric(out["ninth_pair_abs_sum"], errors="coerce").fillna(0.0)
    total_mask = out["ninth_fuel"].fillna("").astype(str).str.strip().str.lower().eq("19_total")
    out.loc[total_mask, "ninth_pair_is_subtotal"] = True
    out["subtotal_alignment"] = out.apply(
        lambda row: _subtotal_alignment(bool(row.get("leap_is_subtotal", False)), bool(row.get("ninth_pair_is_subtotal", False))),
        axis=1,
    )
    return _reorder_columns(
        out,
        [
            "leap_sector_name_original",
            "leap_sector_name_full_path",
            "raw_leap_fuel_name",
            "value",
            "ninth_sector",
            "ninth_fuel",
            "pair_mapping_cardinality",
            "leap_is_subtotal",
            "ninth_pair_is_subtotal",
            "subtotal_mismatch_is_ok",
            "subtotal_alignment",
            "ninth_pair_abs_sum",
            "many_to_many_is_ok",
            "remove_row",
            "remove_row_reason",
        ],
    )


def _active_pairs(frame: pd.DataFrame, col_a: str, col_b: str) -> set[tuple[str, str]]:
    """Return the set of (col_a, col_b) pairs in active (non-removed) rows."""
    active = frame[_active_mask(frame)].copy()
    a = active[col_a].fillna("").astype(str).str.strip() if col_a in active.columns else pd.Series("", index=active.index)
    b = active[col_b].fillna("").astype(str).str.strip() if col_b in active.columns else pd.Series("", index=active.index)
    return {(av, bv) for av, bv in zip(a, b) if av and bv}


def _build_coverage_gaps(
    esto_sheet: pd.DataFrame,
    ninth_sheet: pd.DataFrame,
    esto_lookup: pd.DataFrame,
    ninth_lookup: pd.DataFrame,
) -> pd.DataFrame:
    """
    Return a DataFrame of all coverage gaps — pairs with abs values > 0 that are
    missing from the active mapping rows.

    Columns: gap_type, key_col_1, key_col_2, pair_1, pair_2, abs_sum
      gap_type values:
        "esto_missing"   – esto data pair not in any active esto mapping row
        "ninth_missing"  – 9th data pair not in any active ninth mapping row
        "leap_unmapped_esto"  – LEAP pair with value > 0 but no esto target in esto sheet
        "leap_unmapped_ninth" – LEAP pair with value > 0 but no ninth target in ninth sheet
    """
    records: list[dict] = []

    # --- 1. ESTO data pairs missing from the esto mapping ---
    esto_data_pairs = set(
        zip(
            esto_lookup.loc[esto_lookup["esto_pair_abs_sum"] > 0, "esto_flow"].astype(str).str.strip(),
            esto_lookup.loc[esto_lookup["esto_pair_abs_sum"] > 0, "esto_product"].astype(str).str.strip(),
        )
    )
    esto_mapped_pairs = _active_pairs(esto_sheet, "esto_flow", "esto_product")
    for flow, product in sorted(esto_data_pairs - esto_mapped_pairs):
        abs_sum = float(
            esto_lookup.loc[
                esto_lookup["esto_flow"].astype(str).str.strip().eq(flow)
                & esto_lookup["esto_product"].astype(str).str.strip().eq(product),
                "esto_pair_abs_sum",
            ].sum()
        )
        records.append({"gap_type": "esto_missing", "key_col_1": "esto_flow", "key_col_2": "esto_product", "pair_1": flow, "pair_2": product, "abs_sum": abs_sum})

    # --- 2. 9th data pairs missing from the ninth mapping ---
    ninth_data_pairs = set(
        zip(
            ninth_lookup.loc[ninth_lookup["ninth_pair_abs_sum"] > 0, "ninth_sector"].astype(str).str.strip(),
            ninth_lookup.loc[ninth_lookup["ninth_pair_abs_sum"] > 0, "ninth_fuel"].astype(str).str.strip(),
        )
    )
    ninth_mapped_pairs = _active_pairs(ninth_sheet, "ninth_sector", "ninth_fuel")
    for sector, fuel in sorted(ninth_data_pairs - ninth_mapped_pairs):
        abs_sum = float(
            ninth_lookup.loc[
                ninth_lookup["ninth_sector"].astype(str).str.strip().eq(sector)
                & ninth_lookup["ninth_fuel"].astype(str).str.strip().eq(fuel),
                "ninth_pair_abs_sum",
            ].sum()
        )
        records.append({"gap_type": "ninth_missing", "key_col_1": "ninth_sector", "key_col_2": "ninth_fuel", "pair_1": sector, "pair_2": fuel, "abs_sum": abs_sum})

    # --- 3. LEAP pairs with abs(value) > 0 that are unmapped ---
    for sheet, gap_type, target_a, target_b in [
        (esto_sheet, "leap_unmapped_esto", "esto_flow", "esto_product"),
        (ninth_sheet, "leap_unmapped_ninth", "ninth_sector", "ninth_fuel"),
    ]:
        active = sheet[_active_mask(sheet)].copy()
        if "value" not in active.columns:
            continue
        values = pd.to_numeric(active["value"], errors="coerce").fillna(0.0).abs()
        leap_sector = active["leap_sector_name_full_path"].fillna("").astype(str).str.strip() if "leap_sector_name_full_path" in active.columns else pd.Series("", index=active.index)
        leap_fuel = active["raw_leap_fuel_name"].fillna("").astype(str).str.strip() if "raw_leap_fuel_name" in active.columns else pd.Series("", index=active.index)
        ta = active[target_a].fillna("").astype(str).str.strip() if target_a in active.columns else pd.Series("", index=active.index)
        tb = active[target_b].fillna("").astype(str).str.strip() if target_b in active.columns else pd.Series("", index=active.index)
        unmapped_mask = (values > 0) & (ta.eq("") | tb.eq(""))
        for sector, fuel in sorted(set(zip(leap_sector[unmapped_mask], leap_fuel[unmapped_mask]))):
            abs_sum = float(values[unmapped_mask & leap_sector.eq(sector) & leap_fuel.eq(fuel)].sum())
            records.append({"gap_type": gap_type, "key_col_1": "leap_sector_name_full_path", "key_col_2": "raw_leap_fuel_name", "pair_1": sector, "pair_2": fuel, "abs_sum": abs_sum})

    return pd.DataFrame(records, columns=["gap_type", "key_col_1", "key_col_2", "pair_1", "pair_2", "abs_sum"])


COVERAGE_GAPS_PATH = MAPPING_WORKBOOK_PATH.parent / "mapping_coverage_gaps.csv"


def _report_coverage_gaps(gaps: pd.DataFrame, *, error_on_gaps: bool) -> None:
    """Write gaps CSV and either raise or warn depending on *error_on_gaps*."""
    if gaps.empty:
        if COVERAGE_GAPS_PATH.exists():
            COVERAGE_GAPS_PATH.unlink()
        return

    gaps.to_csv(COVERAGE_GAPS_PATH, index=False)

    summary_lines: list[str] = []
    for gap_type, group in gaps.groupby("gap_type"):
        summary_lines.append(f"  {gap_type}: {len(group)} pair(s)")
    summary = (
        f"{len(gaps)} coverage gap(s) found in leap_mappings.xlsx "
        f"(written to {COVERAGE_GAPS_PATH.name}):\n" + "\n".join(summary_lines)
    )

    if error_on_gaps:
        raise ValueError(summary)
    else:
        import warnings
        warnings.warn(summary, stacklevel=3)


ARCHIVE_DIR = MAPPING_WORKBOOK_PATH.parent / "archive"


def _backup_workbook(path: Path) -> Path:
    ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
    backup_path = ARCHIVE_DIR / f"{path.stem}.before_refresh_mapping_maintenance_columns_{pd.Timestamp.now():%Y%m%d_%H%M%S}{path.suffix}"
    shutil.copy2(path, backup_path)
    return backup_path


def _assert_not_open(path: Path) -> None:
    """Raise a clear error if the file is locked (open in Excel or another process)."""
    try:
        with open(path, "r+b"):
            pass
    except PermissionError:
        raise PermissionError(
            f"{path.name} is open in another application (e.g. Excel). "
            "Close it and re-run the workflow."
        ) from None


def run_workflow(*, error_on_gaps: bool = True) -> dict[str, object]:
    """
    Refresh mapping maintenance columns.

    Parameters
    ----------
    error_on_gaps:
        If True (default), raise a ValueError when coverage gaps are found.
        If False, emit a warning and continue; gaps are still written to
        config/mapping_coverage_gaps.csv.
    """
    if not MAPPING_WORKBOOK_PATH.exists():
        raise FileNotFoundError(f"Missing mapping workbook: {MAPPING_WORKBOOK_PATH}")
    _assert_not_open(MAPPING_WORKBOOK_PATH)
    backup_path = _backup_workbook(MAPPING_WORKBOOK_PATH)

    esto_sheet = pd.read_excel(MAPPING_WORKBOOK_PATH, sheet_name=ESTO_SHEET, dtype=object).fillna("")
    ninth_sheet = pd.read_excel(MAPPING_WORKBOOK_PATH, sheet_name=NINTH_SHEET, dtype=object).fillna("")

    esto_lookup = _load_esto_lookup()
    ninth_lookup = _load_ninth_lookup()

    refreshed_esto = _refresh_esto_sheet(esto_sheet, esto_lookup)
    refreshed_ninth = _refresh_ninth_sheet(ninth_sheet, ninth_lookup)

    gaps = _build_coverage_gaps(esto_sheet, ninth_sheet, esto_lookup, ninth_lookup)
    _report_coverage_gaps(gaps, error_on_gaps=error_on_gaps)

    with pd.ExcelWriter(MAPPING_WORKBOOK_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        refreshed_esto.to_excel(writer, sheet_name=ESTO_SHEET, index=False)
        refreshed_ninth.to_excel(writer, sheet_name=NINTH_SHEET, index=False)

    return {
        "mapping_workbook": str(MAPPING_WORKBOOK_PATH),
        "backup_workbook": str(backup_path),
        "coverage_gaps_csv": str(COVERAGE_GAPS_PATH) if not gaps.empty else None,
        "coverage_gaps_count": int(len(gaps)),
        "leap_combined_esto_rows": int(len(refreshed_esto)),
        "leap_combined_ninth_rows": int(len(refreshed_ninth)),
    }


#%%
RUN_WORKFLOW = True
# Set to False to emit a warning instead of raising an error when coverage gaps are found.
ERROR_ON_GAPS = False

WORKFLOW_RESULT: dict[str, object] | None = None
if RUN_WORKFLOW:
    WORKFLOW_RESULT = run_workflow(error_on_gaps=ERROR_ON_GAPS)
    print("[OK] Mapping maintenance columns refreshed.")
    for key, value in WORKFLOW_RESULT.items():
        print(f"- {key}: {value}")
#%%