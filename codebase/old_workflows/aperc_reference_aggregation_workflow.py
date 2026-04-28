#%%
"""
APERC aggregated reference prototype workflow.

Notebook-first script that builds:
- total final energy demand aggregated reference series (7th + 8th + 9th)
- fuel-level aggregated reference series for comparable fuels
- diagnostics and assumptions
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from codebase.utilities.output_paths import REFERENCE_ANALYSIS_ROOT


# -----------------------------------------------------------------------------
# Notebook-editable toggles
# -----------------------------------------------------------------------------
REPO_ROOT = Path(__file__).resolve().parents[1]
APERC_WORKBOOK_PATH = REPO_ROOT / "data" / "Data for comparison  - APERC outlooks .xlsx"
USE_NINTH_FUEL_SOURCE = True
NINTH_FUEL_SOURCE_PATH = APERC_WORKBOOK_PATH
NINTH_FUEL_SOURCE_SHEET = "APEC 9th fuels"
NINTH_FUEL_SOURCE_SECTOR = "FED"
OUTPUT_DIR = REFERENCE_ANALYSIS_ROOT / "aperc_reference_aggregation_prototype"
CHART_BACKEND = "plotly"  # "plotly" preferred, fallback to "none"
WRITE_CHARTS = True
INTERPOLATE_7TH_ANNUAL = True
WRITE_DASHBOARD = True
EXTEND_7TH_8TH_TO_2060 = True
EXTENSION_START_YEAR = 2051
EXTENSION_END_YEAR = 2060
EXTENSION_LOOKBACK_YEARS = 12
USE_WEIGHTED_EDITION_AVERAGE = True
LIMIT_AGGREGATED_SERIES_START = True
AGGREGATED_SERIES_START_EDITION = "9th"
AGGREGATED_SERIES_START_YEAR = 2022
REBASE_AGGREGATED_SERIES_TO_START_EDITION = True
EDITION_WEIGHTS = {
    "7th": 33.0,
    "8th": 66.0,
    "9th": 99.0,
    "9th_optional": 99.0,
}


# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
SHEET_7TH = "APEC 7th (in PJ)"
SHEET_8TH = "APEC 8th"
SHEET_9TH = "APEC 9th"

TOTAL_METRIC = "total_final_energy_demand"
FUEL_METRIC = "fuel_final_energy_demand"
TPES_METRIC = "total_primary_energy_supply"
BASE_YEAR_MARKERS = [2016, 2018, 2022]

TARGET_FUELS = [
    "Coal",
    "Oil",
    "Gas",
    "Electricity",
    "Heat",
    "Hydrogen",
    "Other",
    "Renewables_total",
]


@dataclass
class ParseResult:
    total: pd.DataFrame
    fuel: pd.DataFrame
    notes: list[dict[str, Any]]
    anchors: dict[str, Any]


def _clean_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _norm(value: Any) -> str:
    return _clean_text(value).lower()


def _is_year_label(value: Any) -> bool:
    if pd.isna(value):
        return False
    try:
        year = int(float(value))
    except (TypeError, ValueError):
        return False
    return 1800 <= year <= 2200


def _year_columns_from_frame(df: pd.DataFrame) -> list[str]:
    out: list[str] = []
    for col in df.columns:
        text = _clean_text(col)
        if text and text.replace(".", "", 1).isdigit():
            year = int(float(text))
            if 1800 <= year <= 2200:
                out.append(col)
    return out


def _assert_no_duplicates(df: pd.DataFrame, keys: list[str], name: str) -> None:
    dupes = df.duplicated(subset=keys, keep=False)
    if dupes.any():
        sample = df.loc[dupes, keys].head(10).to_dict(orient="records")
        raise RuntimeError(f"{name} has duplicate keys on {keys}. Sample: {sample}")


def _edition_has_numeric_year(df: pd.DataFrame, edition: str, year: int) -> bool:
    if df.empty or "edition" not in df.columns:
        return False
    subset = df[df["edition"] == edition].copy()
    if subset.empty:
        return False
    subset = subset[pd.to_numeric(subset["value"], errors="coerce").notna()].copy()
    if subset.empty:
        return False
    return bool((subset["year"].astype(int) == int(year)).any())


def _aggregate_numeric_rows(
    rows: pd.DataFrame,
    *,
    edition_weights: dict[str, float] | None = None,
) -> Any:
    if rows.empty:
        return pd.NA
    vals = pd.to_numeric(rows["value"], errors="coerce")
    valid = rows.loc[vals.notna()].copy()
    if valid.empty:
        return pd.NA
    valid["value_num"] = pd.to_numeric(valid["value"], errors="coerce")
    if not edition_weights:
        return valid["value_num"].mean()
    def _resolve_weight(edition: Any) -> float:
        key = str(edition)
        if key in edition_weights:
            return float(edition_weights[key])
        if key == "9th_optional" and "9th" in edition_weights:
            return float(edition_weights["9th"])
        if key == "9th" and "9th_optional" in edition_weights:
            return float(edition_weights["9th_optional"])
        return 0.0

    valid["weight"] = valid["edition"].map(_resolve_weight)
    weighted = valid[valid["weight"] > 0].copy()
    if weighted.empty:
        return valid["value_num"].mean()
    weight_sum = float(weighted["weight"].sum())
    if abs(weight_sum) <= 1e-12:
        return valid["value_num"].mean()
    return float((weighted["value_num"] * weighted["weight"]).sum() / weight_sum)


def _edition_value_column_name(edition: str) -> str:
    return f"value_{edition}_pj"


def _rebase_aggregated_series_to_edition(
    df: pd.DataFrame,
    *,
    start_year: int | None,
    start_edition: str | None,
    group_cols: list[str] | None = None,
) -> pd.DataFrame:
    if df.empty or start_year is None or not start_edition:
        return df
    ref_col = _edition_value_column_name(start_edition)
    if ref_col not in df.columns:
        return df

    out = df.copy()
    out["aggregated_value_pj"] = pd.to_numeric(out["aggregated_value_pj"], errors="coerce")
    out[ref_col] = pd.to_numeric(out[ref_col], errors="coerce")

    keys = [col for col in (group_cols or []) if col in out.columns]
    if keys:
        group_iter = out.groupby(keys, dropna=False).groups.values()
    else:
        group_iter = [out.index]

    for idx in group_iter:
        grp = out.loc[idx].copy()
        anchor = grp[grp["year"].astype(int) == int(start_year)].copy()
        if anchor.empty:
            continue
        anchor_agg = pd.to_numeric(anchor["aggregated_value_pj"], errors="coerce")
        anchor_ref = pd.to_numeric(anchor[ref_col], errors="coerce")
        if anchor_agg.isna().all() or anchor_ref.isna().all():
            continue
        delta = float(anchor_ref.iloc[0] - anchor_agg.iloc[0])
        apply_mask = (out.index.isin(idx)) & (out["year"].astype(int) >= int(start_year)) & out["aggregated_value_pj"].notna()
        out.loc[apply_mask, "aggregated_value_pj"] = out.loc[apply_mask, "aggregated_value_pj"] + delta

    return out


def extract_7th_bau_pj_total_and_fuels(workbook_path: Path) -> ParseResult:
    notes: list[dict[str, Any]] = []
    anchors: dict[str, Any] = {}
    df = pd.read_excel(workbook_path, sheet_name=SHEET_7TH, header=None)

    label_col = 1
    # Header year row can appear at row 1 in some files, but row 4 is more stable.
    candidate_header_rows = [1, 4]
    scenario_row = 0

    bau_cols = [c for c in range(df.shape[1]) if _norm(df.iat[scenario_row, c]) == "bau"]
    if len(bau_cols) < 2:
        raise RuntimeError("Could not locate first BAU block boundaries in APEC 7th (in PJ).")
    bau_start_col = bau_cols[0]
    bau_end_col = bau_cols[1]
    anchors["bau_block_start_col"] = int(bau_start_col)
    anchors["bau_block_end_col"] = int(bau_end_col)

    # Pick header row with most year labels in BAU block.
    header_candidates: dict[int, dict[int, int]] = {}
    for hr in candidate_header_rows:
        year_to_col: dict[int, int] = {}
        for c in range(bau_start_col, bau_end_col + 1):
            if not _is_year_label(df.iat[hr, c]):
                continue
            year = int(float(df.iat[hr, c]))
            year_to_col[year] = c
        header_candidates[hr] = year_to_col
    header_row = max(header_candidates, key=lambda r: len(header_candidates[r]))
    year_to_mtoe_col = header_candidates[header_row]

    year_to_pj_col: dict[int, int] = {}
    for year, c in year_to_mtoe_col.items():
        if c + 1 < df.shape[1]:
            year_to_pj_col[year] = c + 1

    if not year_to_mtoe_col:
        raise RuntimeError("No year columns found for 7th BAU block.")
    anchors["header_row_7th"] = int(header_row)
    anchors["years_7th"] = sorted(year_to_mtoe_col.keys())

    label_series = df[label_col].astype(str).str.strip()
    total_label = "Final energy demand by fuel (Mtoe)"
    total_rows = df.index[label_series == total_label].tolist()
    if not total_rows:
        raise RuntimeError("Could not locate 'Final energy demand by fuel (Mtoe)' in APEC 7th (in PJ).")
    total_row = total_rows[0]
    anchors["total_row_7th"] = int(total_row)

    # Validate PJ side when present using Mtoe->PJ conversion (~41.868 PJ per Mtoe).
    # If cached PJ values are missing (common after workbook re-save), fall back to Mtoe*41.868.
    ratios: list[float] = []
    for year, m_col in year_to_mtoe_col.items():
        p_col = year_to_pj_col.get(year)
        if p_col is None:
            continue
        m_val = pd.to_numeric(df.iat[total_row, m_col], errors="coerce")
        p_val = pd.to_numeric(df.iat[total_row, p_col], errors="coerce")
        if pd.notna(m_val) and pd.notna(p_val) and abs(float(m_val)) > 1e-12:
            ratios.append(float(p_val) / float(m_val))
    if ratios:
        median_ratio = float(pd.Series(ratios).median())
        anchors["mtoe_to_pj_median_ratio"] = median_ratio
        if not (35.0 <= median_ratio <= 48.0):
            raise RuntimeError(
                "7th PJ column validation failed. "
                f"Expected Mtoe->PJ ratio near 41.868, got median {median_ratio:.3f}."
            )
        use_fallback_conversion = False
    else:
        anchors["mtoe_to_pj_median_ratio"] = "unavailable_cached_values"
        use_fallback_conversion = True
        notes.append(
            {
                "issue_type": "fallback_conversion",
                "edition": "7th",
                "fuel": "",
                "year": "",
                "detail": "PJ cached cells unavailable; computed PJ from Mtoe using factor 41.868.",
                "handling_rule": "pj = mtoe * 41.868",
            }
        )

    total_rows_out: list[dict[str, Any]] = []
    for year, m_col in sorted(year_to_mtoe_col.items()):
        p_col = year_to_pj_col.get(year)
        if use_fallback_conversion or p_col is None:
            m_val = pd.to_numeric(df.iat[total_row, m_col], errors="coerce")
            value = float(m_val) * 41.868 if pd.notna(m_val) else pd.NA
        else:
            value = pd.to_numeric(df.iat[total_row, p_col], errors="coerce")
        total_rows_out.append(
            {
                "edition": "7th",
                "scenario": "BAU",
                "metric": TOTAL_METRIC,
                "fuel": "ALL",
                "year": year,
                "value": value,
                "units": "PJ",
                "source_sheet": SHEET_7TH,
                "source_label": total_label,
            }
        )

    # Top fuel block only.
    fuel_rows_out: list[dict[str, Any]] = []
    for r in range(total_row + 1, df.shape[0]):
        fuel_label = _clean_text(df.iat[r, label_col])
        if not fuel_label:
            break
        if "(Mtoe)" in fuel_label or "final energy demand by sector" in _norm(fuel_label):
            break
        for year, m_col in sorted(year_to_mtoe_col.items()):
            p_col = year_to_pj_col.get(year)
            if use_fallback_conversion or p_col is None:
                m_val = pd.to_numeric(df.iat[r, m_col], errors="coerce")
                value = float(m_val) * 41.868 if pd.notna(m_val) else pd.NA
            else:
                value = pd.to_numeric(df.iat[r, p_col], errors="coerce")
            fuel_rows_out.append(
                {
                    "edition": "7th",
                    "scenario": "BAU",
                    "metric": FUEL_METRIC,
                    "fuel_raw": fuel_label,
                    "year": year,
                    "value": value,
                    "units": "PJ",
                    "source_sheet": SHEET_7TH,
                    "source_label": total_label,
                }
            )

    # Explicitly record known exclusion.
    notes.append(
        {
            "issue_type": "excluded_fuel",
            "edition": "7th",
            "fuel": "Cooling",
            "year": "",
            "detail": "Excluded from comparable fuel buckets.",
            "handling_rule": "excluded_not_comparable",
        }
    )

    return ParseResult(
        total=pd.DataFrame(total_rows_out),
        fuel=pd.DataFrame(fuel_rows_out),
        notes=notes,
        anchors=anchors,
    )


def extract_8th_reference_total_and_fuels(workbook_path: Path) -> ParseResult:
    notes: list[dict[str, Any]] = []
    anchors: dict[str, Any] = {}
    df = pd.read_excel(workbook_path, sheet_name=SHEET_8TH, header=0)
    label_col = df.columns[0]

    # Total: use first "Final energy demand (PJ)" row from Reference scenario summary.
    total_rows = df.index[df[label_col].astype(str).str.strip() == "Final energy demand (PJ)"].tolist()
    if not total_rows:
        raise RuntimeError("Could not locate 'Final energy demand (PJ)' in APEC 8th.")
    total_row = total_rows[0]
    anchors["total_row_8th"] = int(total_row)

    demand_ref_rows = df.index[df[label_col].astype(str).str.strip() == "Demand - Reference"].tolist()
    if not demand_ref_rows:
        raise RuntimeError("Could not locate 'Demand - Reference' in APEC 8th.")
    demand_ref_row = demand_ref_rows[0]
    anchors["demand_reference_row_8th"] = int(demand_ref_row)

    fuel_header_rows = df.index[
        (df.index > demand_ref_row) & (df[label_col].astype(str).str.strip() == "Final consumption by fuel (PJ)")
    ].tolist()
    if not fuel_header_rows:
        raise RuntimeError("Could not locate 'Final consumption by fuel (PJ)' under Demand - Reference in APEC 8th.")
    fuel_header_row = fuel_header_rows[0]
    anchors["fuel_header_row_8th"] = int(fuel_header_row)

    # Year labels may be in DataFrame columns (original) or on in-sheet header row (row 4 after some re-saves).
    year_positions: list[tuple[int, int]] = []
    year_cols = _year_columns_from_frame(df)
    if year_cols:
        col_to_pos = {c: idx for idx, c in enumerate(df.columns)}
        for col in year_cols:
            year_positions.append((int(float(_clean_text(col))), col_to_pos[col]))
        anchors["year_source_8th"] = "columns"
    else:
        header_row = 4
        for c in range(df.shape[1]):
            val = df.iat[header_row, c]
            if _is_year_label(val):
                year_positions.append((int(float(val)), c))
        if not year_positions:
            raise RuntimeError("No numeric year columns found in APEC 8th.")
        anchors["year_source_8th"] = f"row_{header_row}"

    total_rows_out: list[dict[str, Any]] = []
    for year, c in year_positions:
        total_rows_out.append(
            {
                "edition": "8th",
                "scenario": "Reference",
                "metric": TOTAL_METRIC,
                "fuel": "ALL",
                "year": year,
                "value": pd.to_numeric(df.iat[total_row, c], errors="coerce"),
                "units": "PJ",
                "source_sheet": SHEET_8TH,
                "source_label": "Final energy demand (PJ)",
            }
        )

    fuel_rows_out: list[dict[str, Any]] = []
    for r in range(fuel_header_row + 1, df.shape[0]):
        fuel_label = _clean_text(df.at[r, label_col])
        if not fuel_label:
            continue
        if fuel_label == "Industry (PJ)":
            break
        for year, c in year_positions:
            fuel_rows_out.append(
                {
                    "edition": "8th",
                    "scenario": "Reference",
                    "metric": FUEL_METRIC,
                    "fuel_raw": fuel_label,
                    "year": year,
                    "value": pd.to_numeric(df.iat[r, c], errors="coerce"),
                    "units": "PJ",
                    "source_sheet": SHEET_8TH,
                    "source_label": "Final consumption by fuel (PJ)",
                }
            )

    return ParseResult(
        total=pd.DataFrame(total_rows_out),
        fuel=pd.DataFrame(fuel_rows_out),
        notes=notes,
        anchors=anchors,
    )


def extract_9th_reference_total(workbook_path: Path) -> ParseResult:
    notes: list[dict[str, Any]] = []
    anchors: dict[str, Any] = {}
    raw = pd.read_excel(workbook_path, sheet_name=SHEET_9TH, header=None)

    header_row = 5
    table = raw.iloc[header_row + 1 :].copy()
    table.columns = raw.iloc[header_row].tolist()
    if "Sector" not in table.columns or "Fuel" not in table.columns:
        raise RuntimeError("APEC 9th table header with Sector/Fuel was not found.")

    note_rows = table.index[table["Sector"].astype(str).str.strip() == "PJ, TFEC is TFC minus non energy demand"].tolist()
    if not note_rows:
        raise RuntimeError("Could not locate TFEC note row in APEC 9th.")
    note_row = note_rows[0]
    anchors["tfec_note_row_9th"] = int(note_row)

    ref_rows = table.index[
        (table.index > note_row) & (table["Sector"].astype(str).str.strip() == "Reference")
    ].tolist()
    if not ref_rows:
        raise RuntimeError("Could not locate Reference block after TFEC note in APEC 9th.")
    ref_row = ref_rows[0]
    anchors["reference_row_9th"] = int(ref_row)

    next_target = table.index[
        (table.index > ref_row) & (table["Sector"].astype(str).str.strip() == "Target")
    ].tolist()
    end_row = next_target[0] if next_target else int(table.index.max()) + 1

    block = table[(table.index > ref_row) & (table.index < end_row)].copy()
    tfc_rows = block[
        (block["Sector"].astype(str).str.strip() == "TFC") & (block["Fuel"].astype(str).str.strip() == "TFC")
    ]
    if tfc_rows.empty:
        raise RuntimeError("Could not locate TFC/TFC row in APEC 9th Reference block.")
    tfc_row_idx = tfc_rows.index[0]
    anchors["tfc_total_row_9th"] = int(tfc_row_idx)

    year_cols = [c for c in block.columns if _is_year_label(c)]
    if not year_cols:
        raise RuntimeError("No numeric year columns found in APEC 9th Reference TFC block.")

    total_rows_out: list[dict[str, Any]] = []
    for col in year_cols:
        year = int(float(col))
        total_rows_out.append(
            {
                "edition": "9th",
                "scenario": "Reference",
                "metric": TOTAL_METRIC,
                "fuel": "ALL",
                "year": year,
                "value": pd.to_numeric(block.at[tfc_row_idx, col], errors="coerce"),
                "units": "PJ",
                "source_sheet": SHEET_9TH,
                "source_label": "TFC/TFC in Reference TFEC block",
            }
        )

    notes.append(
        {
            "issue_type": "fuel_gap",
            "edition": "9th",
            "fuel": "",
            "year": "",
            "detail": "Workbook lacks full comparable final-demand-by-fuel table.",
            "handling_rule": "use_optional_external_fuel_source_if_provided",
        }
    )

    return ParseResult(
        total=pd.DataFrame(total_rows_out),
        fuel=pd.DataFrame(columns=["edition", "scenario", "metric", "fuel_raw", "year", "value", "units"]),
        notes=notes,
        anchors=anchors,
    )


def extract_7th_bau_tpes_total(workbook_path: Path) -> pd.DataFrame:
    df = pd.read_excel(workbook_path, sheet_name=SHEET_7TH, header=None)
    label_col = 1
    scenario_row = 0

    bau_cols = [c for c in range(df.shape[1]) if _norm(df.iat[scenario_row, c]) == "bau"]
    if len(bau_cols) < 2:
        raise RuntimeError("Could not locate BAU block in APEC 7th (in PJ) for TPES extraction.")
    bau_start_col, bau_end_col = bau_cols[0], bau_cols[1]

    year_to_mtoe: dict[int, int] = {}
    for hr in (1, 4):
        tmp: dict[int, int] = {}
        for c in range(bau_start_col, bau_end_col + 1):
            if _is_year_label(df.iat[hr, c]):
                tmp[int(float(df.iat[hr, c]))] = c
        if len(tmp) > len(year_to_mtoe):
            year_to_mtoe = tmp
    if not year_to_mtoe:
        raise RuntimeError("No year columns found for APEC 7th TPES extraction.")
    year_to_pj = {y: c + 1 for y, c in year_to_mtoe.items() if c + 1 < df.shape[1]}

    rows = df.index[df[label_col].astype(str).str.strip() == "Total primary energy supply (Mtoe)"].tolist()
    if not rows:
        raise RuntimeError("Could not locate 'Total primary energy supply (Mtoe)' in APEC 7th (in PJ).")
    row = rows[0]

    # Prefer PJ cached values; fallback to conversion when cache is absent.
    use_fallback = True
    for y, c in year_to_pj.items():
        if pd.notna(pd.to_numeric(df.iat[row, c], errors="coerce")):
            use_fallback = False
            break

    out: list[dict[str, Any]] = []
    for year, m_col in sorted(year_to_mtoe.items()):
        p_col = year_to_pj.get(year)
        if (p_col is None) or use_fallback:
            m_val = pd.to_numeric(df.iat[row, m_col], errors="coerce")
            val = float(m_val) * 41.868 if pd.notna(m_val) else pd.NA
        else:
            val = pd.to_numeric(df.iat[row, p_col], errors="coerce")
        out.append(
            {
                "edition": "7th",
                "scenario": "BAU",
                "metric": TPES_METRIC,
                "fuel": "ALL",
                "year": int(year),
                "value": val,
                "units": "PJ",
                "source_sheet": SHEET_7TH,
                "source_label": "Total primary energy supply (Mtoe)",
            }
        )
    return pd.DataFrame(out)


def extract_8th_reference_tpes_total(workbook_path: Path) -> pd.DataFrame:
    df = pd.read_excel(workbook_path, sheet_name=SHEET_8TH, header=0)
    label_col = df.columns[0]
    rows = df.index[df[label_col].astype(str).str.strip() == "Total primary energy supply (PJ)"].tolist()
    if not rows:
        raise RuntimeError("Could not locate 'Total primary energy supply (PJ)' in APEC 8th.")
    row = rows[0]

    year_positions: list[tuple[int, int]] = []
    year_cols = _year_columns_from_frame(df)
    if year_cols:
        col_to_pos = {c: idx for idx, c in enumerate(df.columns)}
        year_positions = [(int(float(_clean_text(c))), col_to_pos[c]) for c in year_cols]
    else:
        for c in range(df.shape[1]):
            if _is_year_label(df.iat[4, c]):
                year_positions.append((int(float(df.iat[4, c])), c))
    if not year_positions:
        raise RuntimeError("No year columns found for APEC 8th TPES extraction.")

    out: list[dict[str, Any]] = []
    for year, c in year_positions:
        out.append(
            {
                "edition": "8th",
                "scenario": "Reference",
                "metric": TPES_METRIC,
                "fuel": "ALL",
                "year": int(year),
                "value": pd.to_numeric(df.iat[row, c], errors="coerce"),
                "units": "PJ",
                "source_sheet": SHEET_8TH,
                "source_label": "Total primary energy supply (PJ)",
            }
        )
    return pd.DataFrame(out)


def extract_9th_reference_tpes_total(workbook_path: Path) -> pd.DataFrame:
    raw = pd.read_excel(workbook_path, sheet_name=SHEET_9TH, header=None)
    header_row = 5
    table = raw.iloc[header_row + 1 :].copy()
    table.columns = raw.iloc[header_row].tolist()
    if "Sector" not in table.columns:
        raise RuntimeError("APEC 9th Sector column not found for TPES extraction.")

    # TPES rows are in the leading block under Sector=TPES_including_bunkers.
    tpes = table[table["Sector"].astype(str).str.strip() == "TPES_including_bunkers"].copy()
    if tpes.empty:
        raise RuntimeError("No TPES_including_bunkers rows found in APEC 9th.")
    year_cols = [c for c in tpes.columns if _is_year_label(c)]
    if not year_cols:
        raise RuntimeError("No year columns found in APEC 9th TPES rows.")

    summed = tpes[year_cols].apply(pd.to_numeric, errors="coerce").sum(axis=0, min_count=1)
    out = [
        {
            "edition": "9th",
            "scenario": "Reference",
            "metric": TPES_METRIC,
            "fuel": "ALL",
            "year": int(float(col)),
            "value": float(summed[col]) if pd.notna(summed[col]) else pd.NA,
            "units": "PJ",
            "source_sheet": SHEET_9TH,
            "source_label": "Sum of TPES_including_bunkers fuels",
        }
        for col in year_cols
    ]
    return pd.DataFrame(out)


def extract_7th_bau_tpes_fuel(workbook_path: Path) -> pd.DataFrame:
    df = pd.read_excel(workbook_path, sheet_name=SHEET_7TH, header=None)
    label_col = 1
    scenario_row = 0

    bau_cols = [c for c in range(df.shape[1]) if _norm(df.iat[scenario_row, c]) == "bau"]
    if len(bau_cols) < 2:
        raise RuntimeError("Could not locate BAU block in APEC 7th (in PJ) for TPES fuel extraction.")
    bau_start_col, bau_end_col = bau_cols[0], bau_cols[1]

    year_to_mtoe: dict[int, int] = {}
    for hr in (1, 4):
        tmp: dict[int, int] = {}
        for c in range(bau_start_col, bau_end_col + 1):
            if _is_year_label(df.iat[hr, c]):
                tmp[int(float(df.iat[hr, c]))] = c
        if len(tmp) > len(year_to_mtoe):
            year_to_mtoe = tmp
    year_to_pj = {y: c + 1 for y, c in year_to_mtoe.items() if c + 1 < df.shape[1]}
    if not year_to_mtoe:
        raise RuntimeError("No year columns found for APEC 7th TPES fuel extraction.")

    labels = df[label_col].astype(str).str.strip()
    candidate_rows = df.index[labels == "Total primary energy supply (Mtoe)"].tolist()
    total_rows = [r for r in candidate_rows if (r + 1 < df.shape[0]) and (_norm(df.iat[r + 1, label_col]) == "coal")]
    if not total_rows:
        raise RuntimeError("Could not locate TPES total row in APEC 7th (in PJ).")
    total_row = total_rows[0]

    out: list[dict[str, Any]] = []
    allowed = {"coal", "oil", "gas", "nuclear", "hydro", "non-hydro renewables", "other", "electricity", "hydrogen", "heat"}
    for r in range(total_row + 1, df.shape[0]):
        fuel_label = _clean_text(df.iat[r, label_col])
        if not fuel_label:
            break
        if (_norm(fuel_label) not in allowed) or fuel_label in {"Transformation", "Demand"} or "(Mtoe)" in fuel_label:
            break
        for year, m_col in sorted(year_to_mtoe.items()):
            p_col = year_to_pj.get(year)
            if p_col is not None and pd.notna(pd.to_numeric(df.iat[r, p_col], errors="coerce")):
                val = pd.to_numeric(df.iat[r, p_col], errors="coerce")
            else:
                m_val = pd.to_numeric(df.iat[r, m_col], errors="coerce")
                val = float(m_val) * 41.868 if pd.notna(m_val) else pd.NA
            out.append(
                {
                    "edition": "7th",
                    "scenario": "BAU",
                    "metric": TPES_METRIC,
                    "fuel_raw": fuel_label,
                    "year": int(year),
                    "value": val,
                    "units": "PJ",
                    "source_sheet": SHEET_7TH,
                    "source_label": "Total primary energy supply (Mtoe) by fuel",
                }
            )
    return pd.DataFrame(out)


def extract_8th_reference_tpes_fuel(workbook_path: Path) -> pd.DataFrame:
    df = pd.read_excel(workbook_path, sheet_name=SHEET_8TH, header=0)
    label_col = df.columns[0]
    labels = df[label_col].astype(str).str.strip()
    candidate_rows = df.index[labels == "Total primary energy supply (PJ)"].tolist()
    total_rows = [r for r in candidate_rows if (r + 1 < df.shape[0]) and (_norm(df.at[r + 1, label_col]) == "coal")]
    if not total_rows:
        raise RuntimeError("Could not locate TPES row in APEC 8th for fuel extraction.")
    total_row = total_rows[0]

    year_positions: list[tuple[int, int]] = []
    year_cols = _year_columns_from_frame(df)
    if year_cols:
        col_to_pos = {c: idx for idx, c in enumerate(df.columns)}
        year_positions = [(int(float(_clean_text(c))), col_to_pos[c]) for c in year_cols]
    else:
        for c in range(df.shape[1]):
            if _is_year_label(df.iat[4, c]):
                year_positions.append((int(float(df.iat[4, c])), c))
    if not year_positions:
        raise RuntimeError("No year columns found for APEC 8th TPES fuel extraction.")

    out: list[dict[str, Any]] = []
    allowed = {"coal", "oil", "gas", "nuclear", "hydro", "other renewables", "electricity", "hydrogen", "other"}
    for r in range(total_row + 1, df.shape[0]):
        fuel_label = _clean_text(df.at[r, label_col])
        if not fuel_label:
            continue
        if fuel_label == "Transformation - Reference" or fuel_label.endswith("scenario"):
            break
        if fuel_label == "Series":
            continue
        if _norm(fuel_label) not in allowed:
            break
        for year, c in year_positions:
            out.append(
                {
                    "edition": "8th",
                    "scenario": "Reference",
                    "metric": TPES_METRIC,
                    "fuel_raw": fuel_label,
                    "year": int(year),
                    "value": pd.to_numeric(df.iat[r, c], errors="coerce"),
                    "units": "PJ",
                    "source_sheet": SHEET_8TH,
                    "source_label": "Total primary energy supply (PJ) by fuel",
                }
            )
    return pd.DataFrame(out)


def extract_optional_9th_fuel_source(
    path: Path,
    sheet_name: str | None = None,
    source_sector: str = "FED",
) -> ParseResult:
    notes: list[dict[str, Any]] = []
    anchors: dict[str, Any] = {"optional_source_path": str(path)}
    if not path.exists():
        raise FileNotFoundError(f"Optional 9th fuel source not found: {path}")

    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    elif path.suffix.lower() in {".xlsx", ".xls"}:
        df = pd.read_excel(path, sheet_name=sheet_name)
    else:
        raise ValueError(f"Unsupported optional 9th fuel source extension: {path.suffix}")

    if df.empty:
        raise RuntimeError("Optional 9th fuel source is empty.")

    col_map = {_norm(c): c for c in df.columns}
    fuel_col = None
    for candidate in ("fuel", "fuels", "fuel_label", "series", "product"):
        if candidate in col_map:
            fuel_col = col_map[candidate]
            break
    if fuel_col is None:
        raise RuntimeError("Optional 9th fuel source requires a fuel label column (e.g., 'fuel').")

    scenario_col = col_map.get("scenario") or col_map.get("scenarios")
    sector_col = col_map.get("sector") or col_map.get("sectors")
    unit_col = col_map.get("unit") or col_map.get("units")
    year_cols = [c for c in df.columns if _is_year_label(c)]
    if not year_cols:
        raise RuntimeError("Optional 9th fuel source has no numeric year columns.")

    working = df.copy()
    if scenario_col is not None:
        ref_mask = working[scenario_col].astype(str).str.contains("ref", case=False, na=False)
        working = working[ref_mask].copy()
        if working.empty:
            raise RuntimeError("Optional 9th fuel source has a scenario column but no reference rows.")
        anchors["scenario_filter"] = str(scenario_col)

    if sector_col is not None:
        sector_value = str(source_sector).strip().upper()
        fed_mask = working[sector_col].astype(str).str.strip().str.upper().eq(sector_value)
        working = working[fed_mask].copy()
        if working.empty:
            raise RuntimeError(
                "Optional 9th fuel source has a sector column but no rows where "
                f"{sector_col} == {sector_value}."
            )
        anchors["sector_filter"] = f"{sector_col}={sector_value}"

    if unit_col is not None:
        non_pj = ~working[unit_col].astype(str).str.contains("pj|petajoule", case=False, na=False)
        if non_pj.any():
            bad = sorted(working.loc[non_pj, unit_col].astype(str).dropna().unique().tolist())
            raise RuntimeError(f"Optional 9th fuel source must be in PJ. Found non-PJ units: {bad}")
        anchors["unit_column"] = str(unit_col)

    # Map supplied fuel labels to required buckets.
    fuel_map = {
        "coal": "Coal",
        "oil": "Oil",
        "gas": "Gas",
        "electricity": "Electricity",
        "heat": "Heat",
        "hydrogen": "Hydrogen",
        "other": "Other",
        "renewables_total": "Renewables_total",
        "biomass": "Biomass",
        "other renewables": "Other renewables",
    }

    long_rows: list[dict[str, Any]] = []
    for _, row in working.iterrows():
        raw_label = _clean_text(row[fuel_col])
        mapped = fuel_map.get(_norm(raw_label))
        if mapped is None:
            notes.append(
                {
                    "issue_type": "unmatched_optional_9th_fuel",
                    "edition": "9th_optional",
                    "fuel": raw_label,
                    "year": "",
                    "detail": "Fuel label not in harmonized mapping.",
                    "handling_rule": "excluded_not_comparable",
                }
            )
            continue
        for col in year_cols:
            long_rows.append(
                {
                    "edition": "9th_optional",
                    "scenario": "Reference",
                    "metric": FUEL_METRIC,
                    "fuel_raw": mapped,
                    "year": int(float(col)),
                    "value": pd.to_numeric(row[col], errors="coerce"),
                    "units": "PJ",
                    "source_sheet": str(path.name),
                    "source_label": "optional_9th_fuel_source",
                }
            )

    fuel_df = pd.DataFrame(long_rows)
    if fuel_df.empty:
        raise RuntimeError("Optional 9th fuel source produced no mappable fuel rows.")

    # Build renewables_total if needed from Biomass + Other renewables.
    if "Renewables_total" not in set(fuel_df["fuel_raw"].unique()):
        needed = {"Biomass", "Other renewables"}
        if needed.issubset(set(fuel_df["fuel_raw"].unique())):
            tmp = fuel_df[fuel_df["fuel_raw"].isin(needed)].copy()
            ren = tmp.groupby(["edition", "scenario", "metric", "year", "units", "source_sheet", "source_label"], as_index=False)[
                "value"
            ].sum(min_count=1)
            ren["fuel_raw"] = "Renewables_total"
            fuel_df = pd.concat([fuel_df, ren], ignore_index=True)
        else:
            missing = sorted(list(needed - set(fuel_df["fuel_raw"].unique())))
            raise RuntimeError(
                "Optional 9th fuel source must provide either 'Renewables_total' "
                f"or both 'Biomass' and 'Other renewables'. Missing: {missing}"
            )

    required = {"Coal", "Oil", "Gas", "Electricity", "Heat", "Hydrogen", "Other", "Renewables_total"}
    available = set(fuel_df["fuel_raw"].unique())
    missing_required = sorted(list(required - available))
    if missing_required:
        raise RuntimeError(f"Optional 9th fuel source missing required fuels: {missing_required}")

    return ParseResult(
        total=pd.DataFrame(columns=["edition", "scenario", "metric", "fuel", "year", "value", "units"]),
        fuel=fuel_df,
        notes=notes,
        anchors=anchors,
    )


def _harmonize_fuel_rows(
    fuel_7: pd.DataFrame,
    fuel_8: pd.DataFrame,
    fuel_9_optional: pd.DataFrame | None,
) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    notes: list[dict[str, Any]] = []

    map_7 = {
        "coal": "Coal",
        "oil": "Oil",
        "gas": "Gas",
        "renewables": "Renewables_total",
        "electricity": "Electricity",
        "heat": "Heat",
        "hydrogen": "Hydrogen",
        "other": "Other",
        "cooling": None,
    }
    map_8 = {
        "coal": "Coal",
        "oil": "Oil",
        "gas": "Gas",
        "biomass": "Renewables_total",
        "other renewables": "Renewables_total",
        "hydrogen": "Hydrogen",
        "electricity": "Electricity",
        "heat": "Heat",
        "other": "Other",
    }

    def _apply(df: pd.DataFrame, mapping: dict[str, str | None], edition: str) -> pd.DataFrame:
        if df.empty:
            return df.copy()
        out = df.copy()
        out["fuel_bucket"] = out["fuel_raw"].map(lambda x: mapping.get(_norm(x)))

        excluded = out[out["fuel_bucket"].isna()]["fuel_raw"].dropna().unique().tolist()
        for fuel in sorted(excluded):
            notes.append(
                {
                    "issue_type": "excluded_fuel",
                    "edition": edition,
                    "fuel": str(fuel),
                    "year": "",
                    "detail": "Fuel not included in comparable buckets.",
                    "handling_rule": "excluded_not_comparable",
                }
            )
        out = out[out["fuel_bucket"].notna()].copy()
        grp_cols = ["edition", "scenario", "metric", "fuel_bucket", "year", "units", "source_sheet", "source_label"]
        out = out.groupby(grp_cols, as_index=False)["value"].sum(min_count=1)
        return out

    fuel7_h = _apply(fuel_7, map_7, "7th")
    fuel8_h = _apply(fuel_8, map_8, "8th")

    frames = [fuel7_h, fuel8_h]
    if fuel_9_optional is not None and not fuel_9_optional.empty:
        f9 = fuel_9_optional.copy()
        f9["fuel_bucket"] = f9["fuel_raw"]
        grp_cols = ["edition", "scenario", "metric", "fuel_bucket", "year", "units", "source_sheet", "source_label"]
        f9 = f9.groupby(grp_cols, as_index=False)["value"].sum(min_count=1)
        frames.append(f9)

    all_fuels = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return all_fuels, notes


def _harmonize_tpes_fuel_rows(
    fuel_7: pd.DataFrame,
    fuel_8: pd.DataFrame,
    fuel_9_optional: pd.DataFrame | None,
) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    notes: list[dict[str, Any]] = []
    map_7 = {
        "coal": "Coal",
        "oil": "Oil",
        "gas": "Gas",
        "electricity": "Electricity",
        "heat": "Heat",
        "hydrogen": "Hydrogen",
        "hydro": "Renewables_total",
        "non-hydro renewables": "Renewables_total",
        "nuclear": "Other",
        "other": "Other",
    }
    map_8 = {
        "coal": "Coal",
        "oil": "Oil",
        "gas": "Gas",
        "electricity": "Electricity",
        "heat": "Heat",
        "hydrogen": "Hydrogen",
        "hydro": "Renewables_total",
        "other renewables": "Renewables_total",
        "nuclear": "Other",
        "other": "Other",
    }

    def _apply(df: pd.DataFrame, mapping: dict[str, str], edition: str) -> pd.DataFrame:
        if df.empty:
            return df.copy()
        out = df.copy()
        out["metric"] = TPES_METRIC
        out["fuel_bucket"] = out["fuel_raw"].map(lambda x: mapping.get(_norm(x)))
        excluded = out[out["fuel_bucket"].isna()]["fuel_raw"].dropna().unique().tolist()
        for fuel in sorted(excluded):
            notes.append(
                {
                    "issue_type": "excluded_tpes_fuel",
                    "edition": edition,
                    "fuel": str(fuel),
                    "year": "",
                    "detail": "TPES fuel not included in target buckets.",
                    "handling_rule": "excluded_not_comparable",
                }
            )
        out = out[out["fuel_bucket"].notna()].copy()
        grp_cols = ["edition", "scenario", "metric", "fuel_bucket", "year", "units", "source_sheet", "source_label"]
        out = out.groupby(grp_cols, as_index=False)["value"].sum(min_count=1)
        return out

    f7 = _apply(fuel_7, map_7, "7th")
    f8 = _apply(fuel_8, map_8, "8th")
    frames = [f7, f8]
    if fuel_9_optional is not None and not fuel_9_optional.empty:
        f9 = fuel_9_optional.copy()
        f9["metric"] = TPES_METRIC
        f9["fuel_bucket"] = f9["fuel_raw"]
        grp_cols = ["edition", "scenario", "metric", "fuel_bucket", "year", "units", "source_sheet", "source_label"]
        f9 = f9.groupby(grp_cols, as_index=False)["value"].sum(min_count=1)
        frames.append(f9)
    return (pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()), notes


def _build_aggregated_total(
    total_long: pd.DataFrame,
    *,
    edition_weights: dict[str, float] | None = None,
    aggregation_start_year: int | None = None,
) -> pd.DataFrame:
    if total_long.empty:
        return pd.DataFrame(columns=["year", "aggregated_value_pj", "contributors_count", "contributors_list"])
    rows: list[dict[str, Any]] = []
    for year in sorted(total_long["year"].dropna().astype(int).unique().tolist()):
        year_rows = total_long[total_long["year"] == year].copy()
        year_rows = year_rows[pd.to_numeric(year_rows["value"], errors="coerce").notna()].copy()
        contributors = sorted(year_rows["edition"].astype(str).unique().tolist())
        if aggregation_start_year is not None and int(year) < int(aggregation_start_year):
            value = pd.NA
            contributors_count = 0
            contributors_list = ""
        else:
            value = _aggregate_numeric_rows(year_rows, edition_weights=edition_weights) if contributors else pd.NA
            contributors_count = int(len(contributors))
            contributors_list = "|".join(contributors)
        row = {
            "year": int(year),
            "aggregated_value_pj": value,
            "contributors_count": contributors_count,
            "contributors_list": contributors_list,
        }
        for edition in ("7th", "8th", "9th"):
            subset = year_rows[year_rows["edition"] == edition]["value"]
            row[f"value_{edition}_pj"] = float(subset.iloc[0]) if len(subset) else pd.NA
            row[f"has_{edition}"] = bool(len(subset))
        rows.append(row)
    return pd.DataFrame(rows)


def _build_aggregated_fuel(
    fuel_long: pd.DataFrame,
    *,
    edition_weights: dict[str, float] | None = None,
    aggregation_start_year: int | None = None,
) -> pd.DataFrame:
    if fuel_long.empty:
        return pd.DataFrame(
            columns=["fuel", "year", "aggregated_value_pj", "contributors_count", "contributors_list", "value_7th_pj", "value_8th_pj", "value_9th_optional_pj"]
        )
    rows: list[dict[str, Any]] = []
    for fuel in TARGET_FUELS:
        f = fuel_long[fuel_long["fuel_bucket"] == fuel].copy()
        if f.empty:
            continue
        for year in sorted(f["year"].dropna().astype(int).unique().tolist()):
            fy = f[f["year"] == year].copy()
            fy = fy[pd.to_numeric(fy["value"], errors="coerce").notna()].copy()
            contributors = sorted(fy["edition"].astype(str).unique().tolist())
            if aggregation_start_year is not None and int(year) < int(aggregation_start_year):
                val = pd.NA
                contributors_count = 0
                contributors_list = ""
            else:
                val = _aggregate_numeric_rows(fy, edition_weights=edition_weights) if contributors else pd.NA
                contributors_count = int(len(contributors))
                contributors_list = "|".join(contributors)
            row = {
                "fuel": fuel,
                "year": int(year),
                "aggregated_value_pj": val,
                "contributors_count": contributors_count,
                "contributors_list": contributors_list,
            }
            for edition in ("7th", "8th", "9th_optional"):
                subset = fy[fy["edition"] == edition]["value"]
                row[f"value_{edition}_pj"] = float(subset.iloc[0]) if len(subset) else pd.NA
            rows.append(row)
    return pd.DataFrame(rows)


def _extend_7th_8th_series(
    df: pd.DataFrame,
    group_cols: list[str],
    start_year: int = EXTENSION_START_YEAR,
    end_year: int = EXTENSION_END_YEAR,
    lookback_years: int = EXTENSION_LOOKBACK_YEARS,
) -> pd.DataFrame:
    """Extend 7th/8th annually using recent weighted trend up to start_year-1."""
    if df.empty:
        return df
    out = df.copy()
    if "is_extended_projection" not in out.columns:
        out["is_extended_projection"] = False

    rebuilt: list[pd.DataFrame] = []
    for (edition, *grp_vals), grp in out.groupby(["edition"] + group_cols, dropna=False):
        grp = grp.sort_values("year").copy()
        if edition not in {"7th", "8th"}:
            rebuilt.append(grp)
            continue
        history = grp[(grp["year"] <= start_year - 1) & pd.to_numeric(grp["value"], errors="coerce").notna()].copy()
        if history.empty:
            rebuilt.append(grp)
            continue
        y_last = int(history["year"].max())
        if y_last < start_year - 1:
            rebuilt.append(grp)
            continue
        lookback = max(3, int(lookback_years))
        recent_start = max(int(history["year"].min()), y_last - lookback + 1)
        recent = history[history["year"] >= recent_start].copy().sort_values("year")
        v_last = float(recent.loc[recent["year"] == y_last, "value"].iloc[0])

        # Weighted mean of recent year-over-year growth rates (recent years heavier).
        rates: list[float] = []
        rate_weights: list[float] = []
        for i in range(1, len(recent)):
            y_prev = int(recent.iloc[i - 1]["year"])
            y_cur = int(recent.iloc[i]["year"])
            v_prev = float(recent.iloc[i - 1]["value"])
            v_cur = float(recent.iloc[i]["value"])
            if y_cur != y_prev + 1 or abs(v_prev) <= 1e-12:
                continue
            rates.append((v_cur / v_prev) - 1.0)
            rate_weights.append(float(i))

        if rates:
            growth = float((pd.Series(rates) * pd.Series(rate_weights)).sum() / pd.Series(rate_weights).sum())
            proj = {y: v_last * ((1.0 + growth) ** (y - y_last)) for y in range(start_year, end_year + 1)}
        else:
            # Fallback: weighted linear slope over recent window.
            x = recent["year"].astype(float).values
            yv = pd.to_numeric(recent["value"], errors="coerce").astype(float).values
            w = pd.Series(range(1, len(recent) + 1), dtype=float).values
            x_bar = float((x * w).sum() / w.sum())
            y_bar = float((yv * w).sum() / w.sum())
            denom = float(((w * (x - x_bar) ** 2)).sum())
            slope = float(((w * (x - x_bar) * (yv - y_bar)).sum()) / denom) if denom > 1e-12 else 0.0
            proj = {y: v_last + slope * (y - y_last) for y in range(start_year, end_year + 1)}

        # Preserve existing values if present.
        existing_years = set(grp["year"].astype(int).tolist())
        new_rows = []
        for y in range(start_year, end_year + 1):
            if y in existing_years:
                continue
            base = grp.iloc[-1].copy()
            base["year"] = y
            base["value"] = float(proj[y])
            base["is_extended_projection"] = True
            new_rows.append(base)
        if new_rows:
            grp = pd.concat([grp, pd.DataFrame(new_rows)], ignore_index=True).sort_values("year")
        rebuilt.append(grp)

    combined = pd.concat(rebuilt, ignore_index=True) if rebuilt else out
    return combined


def _interpolate_7th_annual(
    total_long: pd.DataFrame,
    fuel_long: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Interpolate 7th series to annual points between existing years (no extrapolation)."""
    if total_long.empty and fuel_long.empty:
        return total_long, fuel_long

    total_out = total_long.copy()
    fuel_out = fuel_long.copy()
    if "is_extended_projection" not in total_out.columns:
        total_out["is_extended_projection"] = False
    if "is_extended_projection" not in fuel_out.columns:
        fuel_out["is_extended_projection"] = False

    # Total: one series (fuel=ALL) for 7th.
    t7 = total_out[total_out["edition"] == "7th"].copy()
    if not t7.empty:
        t_other = total_out[total_out["edition"] != "7th"].copy()
        t7 = t7.sort_values("year")
        y0, y1 = int(t7["year"].min()), int(t7["year"].max())
        years = pd.DataFrame({"year": list(range(y0, y1 + 1))})
        t7_full = years.merge(t7.drop(columns=[]), on="year", how="left")
        t7_full["value"] = pd.to_numeric(t7_full["value"], errors="coerce").interpolate(
            method="linear", limit_area="inside"
        )
        for col in ["edition", "scenario", "metric", "fuel", "units", "source_sheet", "source_label"]:
            if col in t7.columns:
                t7_full[col] = t7[col].iloc[0]
        total_out = pd.concat([t_other, t7_full[total_out.columns]], ignore_index=True)

    # Fuel: interpolate each 7th fuel independently.
    f7 = fuel_out[fuel_out["edition"] == "7th"].copy()
    if not f7.empty:
        f_other = fuel_out[fuel_out["edition"] != "7th"].copy()
        rebuilt: list[pd.DataFrame] = []
        for fuel_label, grp in f7.groupby("fuel_raw", dropna=False):
            grp = grp.sort_values("year")
            y0, y1 = int(grp["year"].min()), int(grp["year"].max())
            years = pd.DataFrame({"year": list(range(y0, y1 + 1))})
            full = years.merge(grp, on="year", how="left")
            full["value"] = pd.to_numeric(full["value"], errors="coerce").interpolate(
                method="linear", limit_area="inside"
            )
            for col in ["edition", "scenario", "metric", "fuel_raw", "units", "source_sheet", "source_label"]:
                if col in grp.columns:
                    full[col] = grp[col].iloc[0]
            rebuilt.append(full[fuel_out.columns])
        fuel_out = pd.concat([f_other] + rebuilt, ignore_index=True)

    return total_out, fuel_out


def _interpolate_7th_annual_series(
    df: pd.DataFrame,
    group_cols: list[str],
) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    if "is_extended_projection" not in out.columns:
        out["is_extended_projection"] = False
    other = out[out["edition"] != "7th"].copy()
    seven = out[out["edition"] == "7th"].copy()
    if seven.empty:
        return out

    rebuilt: list[pd.DataFrame] = [other]
    for _, grp in seven.groupby(group_cols, dropna=False):
        grp = grp.sort_values("year").copy()
        y0, y1 = int(grp["year"].min()), int(grp["year"].max())
        years = pd.DataFrame({"year": list(range(y0, y1 + 1))})
        full = years.merge(grp, on="year", how="left")
        full["value"] = pd.to_numeric(full["value"], errors="coerce").interpolate(method="linear", limit_area="inside")
        for col in ["edition"] + group_cols + ["units", "source_sheet", "source_label"]:
            if col in grp.columns:
                full[col] = grp[col].iloc[0]
        full["is_extended_projection"] = False
        rebuilt.append(full[out.columns])
    return pd.concat(rebuilt, ignore_index=True)


def _write_assumptions_md(
    path: Path,
    used_optional_9th: bool,
    use_weighted_average: bool,
    aggregation_start_year: int | None,
    aggregation_start_edition: str | None,
    rebase_to_start_edition: bool,
) -> None:
    if aggregation_start_year is None or not aggregation_start_edition:
        start_line = "- Aggregated series start as soon as data are available."
    else:
        start_line = (
            f"- Aggregated series are blank before {int(aggregation_start_year)}, "
            f"using the configured start year from the {aggregation_start_edition} series."
        )
    if rebase_to_start_edition and aggregation_start_year is not None and aggregation_start_edition:
        rebase_line = (
            f"- From {int(aggregation_start_year)} onward, aggregated series are shifted to match the "
            f"{aggregation_start_edition} value at the start year while preserving the aggregated trend."
        )
    else:
        rebase_line = "- Aggregated series are not vertically rebased to a source edition."
    text = f"""# APERC Aggregated Reference Prototype - Assumptions

- 7th uses BAU as reference-equivalent.
- 7th alternating Mtoe/PJ structure is parsed by selecting adjacent PJ columns.
- 8th uses Demand - Reference and Final consumption by fuel (PJ).
- 9th workbook contributes total via Reference TFC/TFC in the TFEC note block.
- 9th workbook does not include a fully comparable fuel table for all buckets.
- Optional 9th fuel source used: {"Yes" if used_optional_9th else "No"}.
- Fuel harmonization bucket: Renewables_total
  - 7th: Renewables
  - 8th: Biomass + Other renewables
- 9th optional: Renewables_total directly, or Biomass + Other renewables
- {start_line[2:]}
- {rebase_line[2:]}
- Aggregation method: {"weighted average using 33/66/99 for 7th/8th/9th" if use_weighted_average else "simple average across available editions"}.
- 7th data are interpolated to annual values before aggregation.
- 7th/8th series are extended from 2051 to 2060 using each series historical trend.
- Missing years are not imputed in output tables; charts connect gaps visually.
"""
    path.write_text(text, encoding="utf-8")


def _build_dashboard(
    output_dir: Path,
    total_long: pd.DataFrame,
    tpes_long: pd.DataFrame,
    total_agg: pd.DataFrame,
    tpes_agg: pd.DataFrame,
    fuel_h: pd.DataFrame,
    fuel_agg: pd.DataFrame,
    tpes_fuel_h: pd.DataFrame,
    tpes_fuel_agg: pd.DataFrame,
) -> Path | None:
    dashboards_dir = output_dir / "dashboards"
    dashboards_dir.mkdir(parents=True, exist_ok=True)
    dashboard_path = dashboards_dir / "index.html"

    try:
        import plotly.graph_objects as go
    except Exception:
        # Fallback: lightweight HTML notice
        dashboard_path.write_text(
            "<html><body><h2>Dashboard unavailable</h2><p>Install plotly to enable dropdown dashboard.</p></body></html>",
            encoding="utf-8",
        )
        return dashboard_path

    fig = go.Figure()
    groups: list[str] = (
        ["FED Total", "TPES Total"]
        + [f"FED {fuel}" for fuel in TARGET_FUELS]
        + [f"TPES {fuel}" for fuel in TARGET_FUELS]
    )
    trace_group: list[str] = []

    # FED total group traces
    t = total_long.copy()
    for edition in ("7th", "8th", "9th"):
        s = t[t["edition"] == edition]
        if s.empty:
            continue
        s_obs = s[~s["is_extended_projection"].fillna(False)]
        s_ext = s[s["is_extended_projection"].fillna(False)]
        if not s_obs.empty:
            fig.add_trace(
                go.Scatter(
                    x=s_obs["year"],
                    y=s_obs["value"],
                    mode="lines+markers",
                    name=f"{edition} source",
                    connectgaps=True,
                    visible=True,
                )
            )
            trace_group.append("FED Total")
        if not s_ext.empty:
            fig.add_trace(
                go.Scatter(
                    x=s_ext["year"],
                    y=s_ext["value"],
                    mode="lines+markers",
                    name=f"{edition} projected extension",
                    line={"dash": "dot"},
                    connectgaps=True,
                    visible=True,
                )
            )
            trace_group.append("FED Total")
    fig.add_trace(
        go.Scatter(
            x=total_agg["year"],
            y=total_agg["aggregated_value_pj"],
            mode="lines+markers",
            name="Aggregated reference",
            line={"width": 4},
            connectgaps=True,
            visible=True,
        )
    )
    trace_group.append("FED Total")

    # TPES total group traces
    for edition in ("7th", "8th", "9th"):
        s = tpes_long[tpes_long["edition"] == edition]
        if s.empty:
            continue
        s_obs = s[~s["is_extended_projection"].fillna(False)]
        s_ext = s[s["is_extended_projection"].fillna(False)]
        if not s_obs.empty:
            fig.add_trace(
                go.Scatter(
                    x=s_obs["year"],
                    y=s_obs["value"],
                    mode="lines+markers",
                    name=f"{edition} source",
                    connectgaps=True,
                    visible=False,
                )
            )
            trace_group.append("TPES Total")
        if not s_ext.empty:
            fig.add_trace(
                go.Scatter(
                    x=s_ext["year"],
                    y=s_ext["value"],
                    mode="lines+markers",
                    name=f"{edition} projected extension",
                    line={"dash": "dot"},
                    connectgaps=True,
                    visible=False,
                )
            )
            trace_group.append("TPES Total")
    fig.add_trace(
        go.Scatter(
            x=tpes_agg["year"],
            y=tpes_agg["aggregated_value_pj"],
            mode="lines+markers",
            name="Aggregated reference",
            line={"width": 4},
            connectgaps=True,
            visible=False,
        )
    )
    trace_group.append("TPES Total")

    # FED fuel group traces (hidden by default)
    for fuel in TARGET_FUELS:
        sf_agg = fuel_agg[fuel_agg["fuel"] == fuel].copy()
        for edition in ("7th", "8th", "9th_optional"):
            vals = sf_agg[f"value_{edition}_pj"] if f"value_{edition}_pj" in sf_agg.columns else pd.Series(dtype=float)
            fig.add_trace(
                go.Scatter(
                    x=sf_agg["year"],
                    y=vals,
                    mode="lines+markers",
                    name=f"{edition} source",
                    connectgaps=True,
                    visible=False,
                )
            )
            trace_group.append(f"FED {fuel}")
        fig.add_trace(
            go.Scatter(
                x=sf_agg["year"],
                y=sf_agg["aggregated_value_pj"],
                mode="lines+markers",
                name="Aggregated reference",
                line={"width": 4},
                connectgaps=True,
                visible=False,
            )
        )
        trace_group.append(f"FED {fuel}")

    # TPES fuel group traces (hidden by default; currently sourced from optional 9th fuel rows).
    for fuel in TARGET_FUELS:
        sf_agg = tpes_fuel_agg[tpes_fuel_agg["fuel"] == fuel].copy()
        for edition in ("7th", "8th", "9th_optional"):
            col = f"value_{edition}_pj"
            vals = sf_agg[col] if col in sf_agg.columns else pd.Series(dtype=float)
            fig.add_trace(
                go.Scatter(
                    x=sf_agg["year"],
                    y=vals,
                    mode="lines+markers",
                    name=f"{edition} source",
                    connectgaps=True,
                    visible=False,
                )
            )
            trace_group.append(f"TPES {fuel}")
        fig.add_trace(
            go.Scatter(
                x=sf_agg["year"],
                y=sf_agg["aggregated_value_pj"] if "aggregated_value_pj" in sf_agg.columns else pd.Series(dtype=float),
                mode="lines+markers",
                name="Aggregated reference",
                line={"width": 4},
                connectgaps=True,
                visible=False,
            )
        )
        trace_group.append(f"TPES {fuel}")

    buttons = []
    for g in groups:
        visible = [tg == g for tg in trace_group]
        buttons.append(
            dict(
                label=g,
                method="update",
                args=[
                    {"visible": visible},
                    {"title": f"APERC Aggregated Reference - {g}", "xaxis": {"title": "Year"}, "yaxis": {"title": "PJ"}},
                ],
            )
        )

    fig.update_layout(
        title="APERC Aggregated Reference - FED Total",
        xaxis_title="Year",
        yaxis_title="PJ",
        updatemenus=[dict(buttons=buttons, direction="down", showactive=True, x=1.02, y=1.0)],
        legend={"orientation": "h", "y": -0.2},
        shapes=[
            {
                "type": "line",
                "x0": by,
                "x1": by,
                "xref": "x",
                "y0": 0,
                "y1": 1,
                "yref": "paper",
                "line": {"color": "#777", "dash": "dot", "width": 1},
            }
            for by in BASE_YEAR_MARKERS
        ],
    )

    fig.write_html(dashboard_path, include_plotlyjs="cdn")
    return dashboard_path


def _build_charts(
    total_long: pd.DataFrame,
    tpes_long: pd.DataFrame,
    total_agg: pd.DataFrame,
    tpes_agg: pd.DataFrame,
    fuel_agg: pd.DataFrame,
    charts_dir: Path,
    backend: str = "plotly",
) -> list[Path]:
    charts_dir.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []
    if backend == "plotly":
        try:
            import plotly.graph_objects as go

            # Total chart
            fig_total = go.Figure()
            for edition in ("7th", "8th", "9th"):
                s = total_long[total_long["edition"] == edition].copy()
                if s.empty:
                    continue
                s_obs = s[~s["is_extended_projection"].fillna(False)]
                s_ext = s[s["is_extended_projection"].fillna(False)]
                if not s_obs.empty:
                    fig_total.add_trace(
                        go.Scatter(
                            x=s_obs["year"],
                            y=s_obs["value"],
                            mode="lines+markers",
                            name=f"{edition} source",
                            connectgaps=True,
                        )
                    )
                if not s_ext.empty:
                    fig_total.add_trace(
                        go.Scatter(
                            x=s_ext["year"],
                            y=s_ext["value"],
                            mode="lines+markers",
                            name=f"{edition} projected extension",
                            line={"dash": "dot"},
                            connectgaps=True,
                        )
                    )
            fig_total.add_trace(
                go.Scatter(
                    x=total_agg["year"],
                    y=total_agg["aggregated_value_pj"],
                    mode="lines+markers",
                    name="Aggregated reference",
                    line={"width": 4},
                    connectgaps=True,
                )
            )
            fig_total.update_layout(
                title="Aggregated Total Final Energy Demand (PJ)",
                xaxis_title="Year",
                yaxis_title="PJ",
                shapes=[
                    {
                        "type": "line",
                        "x0": by,
                        "x1": by,
                        "xref": "x",
                        "y0": 0,
                        "y1": 1,
                        "yref": "paper",
                        "line": {"color": "#777", "dash": "dot", "width": 1},
                    }
                    for by in BASE_YEAR_MARKERS
                ],
            )
            total_path = charts_dir / "total_final_energy_demand.html"
            fig_total.write_html(total_path, include_plotlyjs="cdn")
            written.append(total_path)

            # TPES chart
            fig_tpes = go.Figure()
            for edition in ("7th", "8th", "9th"):
                s = tpes_long[tpes_long["edition"] == edition].copy()
                if s.empty:
                    continue
                s_obs = s[~s["is_extended_projection"].fillna(False)]
                s_ext = s[s["is_extended_projection"].fillna(False)]
                if not s_obs.empty:
                    fig_tpes.add_trace(
                        go.Scatter(
                            x=s_obs["year"],
                            y=s_obs["value"],
                            mode="lines+markers",
                            name=f"{edition} source",
                            connectgaps=True,
                        )
                    )
                if not s_ext.empty:
                    fig_tpes.add_trace(
                        go.Scatter(
                            x=s_ext["year"],
                            y=s_ext["value"],
                            mode="lines+markers",
                            name=f"{edition} projected extension",
                            line={"dash": "dot"},
                            connectgaps=True,
                        )
                    )
            fig_tpes.add_trace(
                go.Scatter(
                    x=tpes_agg["year"],
                    y=tpes_agg["aggregated_value_pj"],
                    mode="lines+markers",
                    name="Aggregated reference",
                    line={"width": 4},
                    connectgaps=True,
                )
            )
            fig_tpes.update_layout(
                title="Aggregated Total Primary Energy Supply (PJ)",
                xaxis_title="Year",
                yaxis_title="PJ",
                shapes=[
                    {
                        "type": "line",
                        "x0": by,
                        "x1": by,
                        "xref": "x",
                        "y0": 0,
                        "y1": 1,
                        "yref": "paper",
                        "line": {"color": "#777", "dash": "dot", "width": 1},
                    }
                    for by in BASE_YEAR_MARKERS
                ],
            )
            tpes_path = charts_dir / "total_primary_energy_supply.html"
            fig_tpes.write_html(tpes_path, include_plotlyjs="cdn")
            written.append(tpes_path)

            # Fuel charts
            for fuel in TARGET_FUELS:
                sf = fuel_agg[fuel_agg["fuel"] == fuel].copy()
                if sf.empty:
                    continue
                fig = go.Figure()
                for edition in ("7th", "8th", "9th_optional"):
                    col = f"value_{edition}_pj"
                    if col not in sf.columns:
                        continue
                    fig.add_trace(
                        go.Scatter(
                            x=sf["year"],
                            y=sf[col],
                            mode="lines+markers",
                            name=f"{edition} source",
                            connectgaps=True,
                        )
                    )
                fig.add_trace(
                    go.Scatter(
                        x=sf["year"],
                        y=sf["aggregated_value_pj"],
                        mode="lines+markers",
                        name="Aggregated reference",
                        line={"width": 4},
                        connectgaps=True,
                    )
                )
                fig.update_layout(
                    title=f"Aggregated Fuel Series: {fuel} (PJ)",
                    xaxis_title="Year",
                    yaxis_title="PJ",
                    shapes=[
                        {
                            "type": "line",
                            "x0": by,
                            "x1": by,
                            "xref": "x",
                            "y0": 0,
                            "y1": 1,
                            "yref": "paper",
                            "line": {"color": "#777", "dash": "dot", "width": 1},
                        }
                        for by in BASE_YEAR_MARKERS
                    ],
                )
                out = charts_dir / f"fuel_{fuel.lower()}.html"
                fig.write_html(out, include_plotlyjs="cdn")
                written.append(out)
            return written
        except Exception:
            # Fall through to matplotlib PNGs.
            pass

    try:
        import matplotlib.pyplot as plt
    except Exception:
        return written

    # Fallback charting: interpolate only for plotting continuity.
    def _plot_with_display_interp(ax, x: pd.Series, y: pd.Series, label: str, linewidth: float = 1.8) -> None:
        ys = pd.to_numeric(y, errors="coerce")
        ys_interp = ys.interpolate(method="linear", limit_direction="both")
        ax.plot(x, ys_interp, marker="o", linewidth=linewidth, label=label)

    fig, ax = plt.subplots(figsize=(11, 6))
    for edition in ("7th", "8th", "9th"):
        s = total_long[total_long["edition"] == edition].copy()
        if s.empty:
            continue
        s_obs = s[~s["is_extended_projection"].fillna(False)]
        s_ext = s[s["is_extended_projection"].fillna(False)]
        if not s_obs.empty:
            _plot_with_display_interp(ax, s_obs["year"], s_obs["value"], f"{edition} source")
        if not s_ext.empty:
            ys = pd.to_numeric(s_ext["value"], errors="coerce").interpolate(method="linear", limit_direction="both")
            ax.plot(s_ext["year"], ys, marker="o", linestyle=":", linewidth=1.8, label=f"{edition} projected extension")
    _plot_with_display_interp(ax, total_agg["year"], total_agg["aggregated_value_pj"], "Aggregated reference", linewidth=3.0)
    for by in BASE_YEAR_MARKERS:
        ax.axvline(by, linestyle=":", linewidth=1.0, color="#666")
    ax.set_title("Aggregated Total Final Energy Demand (PJ)")
    ax.set_xlabel("Year")
    ax.set_ylabel("PJ")
    ax.grid(True, alpha=0.3)
    ax.legend()
    total_png = charts_dir / "total_final_energy_demand.png"
    fig.tight_layout()
    fig.savefig(total_png, dpi=160)
    plt.close(fig)
    written.append(total_png)

    # TPES fallback chart
    fig, ax = plt.subplots(figsize=(11, 6))
    for edition in ("7th", "8th", "9th"):
        s = tpes_long[tpes_long["edition"] == edition].copy()
        if s.empty:
            continue
        s_obs = s[~s["is_extended_projection"].fillna(False)]
        s_ext = s[s["is_extended_projection"].fillna(False)]
        if not s_obs.empty:
            _plot_with_display_interp(ax, s_obs["year"], s_obs["value"], f"{edition} source")
        if not s_ext.empty:
            ys = pd.to_numeric(s_ext["value"], errors="coerce").interpolate(method="linear", limit_direction="both")
            ax.plot(s_ext["year"], ys, marker="o", linestyle=":", linewidth=1.8, label=f"{edition} projected extension")
    _plot_with_display_interp(ax, tpes_agg["year"], tpes_agg["aggregated_value_pj"], "Aggregated reference", linewidth=3.0)
    for by in BASE_YEAR_MARKERS:
        ax.axvline(by, linestyle=":", linewidth=1.0, color="#666")
    ax.set_title("Aggregated Total Primary Energy Supply (PJ)")
    ax.set_xlabel("Year")
    ax.set_ylabel("PJ")
    ax.grid(True, alpha=0.3)
    ax.legend()
    tpes_png = charts_dir / "total_primary_energy_supply.png"
    fig.tight_layout()
    fig.savefig(tpes_png, dpi=160)
    plt.close(fig)
    written.append(tpes_png)

    for fuel in TARGET_FUELS:
        sf = fuel_agg[fuel_agg["fuel"] == fuel].copy()
        if sf.empty:
            continue
        fig, ax = plt.subplots(figsize=(11, 6))
        for edition in ("7th", "8th", "9th_optional"):
            col = f"value_{edition}_pj"
            if col in sf.columns:
                _plot_with_display_interp(ax, sf["year"], sf[col], f"{edition} source")
        _plot_with_display_interp(ax, sf["year"], sf["aggregated_value_pj"], "Aggregated reference", linewidth=3.0)
        for by in BASE_YEAR_MARKERS:
            ax.axvline(by, linestyle=":", linewidth=1.0, color="#666")
        ax.set_title(f"Aggregated Fuel Series: {fuel} (PJ)")
        ax.set_xlabel("Year")
        ax.set_ylabel("PJ")
        ax.grid(True, alpha=0.3)
        ax.legend()
        out_png = charts_dir / f"fuel_{fuel.lower()}.png"
        fig.tight_layout()
        fig.savefig(out_png, dpi=160)
        plt.close(fig)
        written.append(out_png)

    return written


def run_workflow(config: dict[str, Any] | None = None) -> dict[str, str]:
    cfg = {
        "APERC_WORKBOOK_PATH": APERC_WORKBOOK_PATH,
        "NINTH_FUEL_SOURCE_PATH": NINTH_FUEL_SOURCE_PATH,
        "NINTH_FUEL_SOURCE_SHEET": NINTH_FUEL_SOURCE_SHEET,
        "NINTH_FUEL_SOURCE_SECTOR": NINTH_FUEL_SOURCE_SECTOR,
        "USE_NINTH_FUEL_SOURCE": USE_NINTH_FUEL_SOURCE,
        "OUTPUT_DIR": OUTPUT_DIR,
        "CHART_BACKEND": CHART_BACKEND,
        "WRITE_CHARTS": WRITE_CHARTS,
        "INTERPOLATE_7TH_ANNUAL": INTERPOLATE_7TH_ANNUAL,
        "WRITE_DASHBOARD": WRITE_DASHBOARD,
        "EXTEND_7TH_8TH_TO_2060": EXTEND_7TH_8TH_TO_2060,
        "EXTENSION_START_YEAR": EXTENSION_START_YEAR,
        "EXTENSION_END_YEAR": EXTENSION_END_YEAR,
        "EXTENSION_LOOKBACK_YEARS": EXTENSION_LOOKBACK_YEARS,
        "USE_WEIGHTED_EDITION_AVERAGE": USE_WEIGHTED_EDITION_AVERAGE,
        "LIMIT_AGGREGATED_SERIES_START": LIMIT_AGGREGATED_SERIES_START,
        "AGGREGATED_SERIES_START_EDITION": AGGREGATED_SERIES_START_EDITION,
        "AGGREGATED_SERIES_START_YEAR": AGGREGATED_SERIES_START_YEAR,
        "REBASE_AGGREGATED_SERIES_TO_START_EDITION": REBASE_AGGREGATED_SERIES_TO_START_EDITION,
        "EDITION_WEIGHTS": dict(EDITION_WEIGHTS),
    }
    if config:
        cfg.update(config)

    workbook_path = Path(cfg["APERC_WORKBOOK_PATH"])
    output_dir = Path(cfg["OUTPUT_DIR"])
    output_dir.mkdir(parents=True, exist_ok=True)
    supporting_dir = output_dir / "supporting_files"
    analysis_dir = supporting_dir / "analysis"
    checks_dir = supporting_dir / "checks"
    supporting_dir.mkdir(parents=True, exist_ok=True)
    analysis_dir.mkdir(parents=True, exist_ok=True)
    checks_dir.mkdir(parents=True, exist_ok=True)
    charts_dir = output_dir / "charts"
    p7 = extract_7th_bau_pj_total_and_fuels(workbook_path)
    p8 = extract_8th_reference_total_and_fuels(workbook_path)
    p9 = extract_9th_reference_total(workbook_path)
    t7 = extract_7th_bau_tpes_total(workbook_path)
    t8 = extract_8th_reference_tpes_total(workbook_path)
    t9 = extract_9th_reference_tpes_total(workbook_path)
    t7_fuel = extract_7th_bau_tpes_fuel(workbook_path)
    t8_fuel = extract_8th_reference_tpes_fuel(workbook_path)

    optional_9th: ParseResult | None = None
    optional_9th_tpes: ParseResult | None = None
    if bool(cfg["USE_NINTH_FUEL_SOURCE"]):
        fuel_path_raw = cfg.get("NINTH_FUEL_SOURCE_PATH")
        if fuel_path_raw is None:
            raise RuntimeError("USE_NINTH_FUEL_SOURCE=True but NINTH_FUEL_SOURCE_PATH is not set.")
        optional_9th = extract_optional_9th_fuel_source(
            Path(fuel_path_raw),
            sheet_name=cfg.get("NINTH_FUEL_SOURCE_SHEET"),
            source_sector=str(cfg.get("NINTH_FUEL_SOURCE_SECTOR", "FED")),
        )
        # Optional TPES fuel rows from same source/sheet when sector column is present.
        try:
            optional_9th_tpes = extract_optional_9th_fuel_source(
                Path(fuel_path_raw),
                sheet_name=cfg.get("NINTH_FUEL_SOURCE_SHEET"),
                source_sector="TPES",
            )
        except Exception:
            optional_9th_tpes = None

    total_long = pd.concat([p7.total, p8.total, p9.total], ignore_index=True)
    total_long["is_extended_projection"] = False
    fuel_long_raw = pd.concat(
        [p7.fuel, p8.fuel, optional_9th.fuel if optional_9th else pd.DataFrame(columns=p7.fuel.columns)],
        ignore_index=True,
    )
    fuel_long_raw["is_extended_projection"] = False
    tpes_long = pd.concat([t7, t8, t9], ignore_index=True)
    tpes_long["is_extended_projection"] = False

    if bool(cfg["INTERPOLATE_7TH_ANNUAL"]):
        total_long, fuel_long_raw = _interpolate_7th_annual(total_long=total_long, fuel_long=fuel_long_raw)
        tpes_long = _interpolate_7th_annual_series(
            tpes_long,
            group_cols=["scenario", "metric", "fuel"],
        )
        t7_fuel = _interpolate_7th_annual_series(
            t7_fuel.assign(is_extended_projection=False),
            group_cols=["scenario", "metric", "fuel_raw"],
        ).drop(columns=["is_extended_projection"], errors="ignore")

    if bool(cfg["EXTEND_7TH_8TH_TO_2060"]):
        total_long = _extend_7th_8th_series(
            total_long,
            group_cols=["scenario", "metric", "fuel", "units", "source_sheet", "source_label"],
            start_year=int(cfg["EXTENSION_START_YEAR"]),
            end_year=int(cfg["EXTENSION_END_YEAR"]),
            lookback_years=int(cfg["EXTENSION_LOOKBACK_YEARS"]),
        )
        fuel_long_raw = _extend_7th_8th_series(
            fuel_long_raw,
            group_cols=["scenario", "metric", "fuel_raw", "units", "source_sheet", "source_label"],
            start_year=int(cfg["EXTENSION_START_YEAR"]),
            end_year=int(cfg["EXTENSION_END_YEAR"]),
            lookback_years=int(cfg["EXTENSION_LOOKBACK_YEARS"]),
        )
        tpes_long = _extend_7th_8th_series(
            tpes_long,
            group_cols=["scenario", "metric", "fuel", "units", "source_sheet", "source_label"],
            start_year=int(cfg["EXTENSION_START_YEAR"]),
            end_year=int(cfg["EXTENSION_END_YEAR"]),
            lookback_years=int(cfg["EXTENSION_LOOKBACK_YEARS"]),
        )

    fuel_h, fuel_notes = _harmonize_fuel_rows(
        fuel_7=fuel_long_raw[fuel_long_raw["edition"] == "7th"].copy(),
        fuel_8=fuel_long_raw[fuel_long_raw["edition"] == "8th"].copy(),
        fuel_9_optional=optional_9th.fuel if optional_9th else None,
    )
    tpes_fuel_h = pd.DataFrame()
    tpes_fuel_agg = pd.DataFrame(
        columns=[
            "fuel",
            "year",
            "aggregated_value_pj",
            "contributors_count",
            "contributors_list",
            "value_7th_pj",
            "value_8th_pj",
            "value_9th_optional_pj",
        ]
    )
    tpes_fuel_h, tpes_fuel_notes = _harmonize_tpes_fuel_rows(
        fuel_7=t7_fuel,
        fuel_8=t8_fuel,
        fuel_9_optional=optional_9th_tpes.fuel if optional_9th_tpes is not None else None,
    )
    if bool(cfg["EXTEND_7TH_8TH_TO_2060"]) and not tpes_fuel_h.empty:
        tpes_fuel_h = _extend_7th_8th_series(
            tpes_fuel_h,
            group_cols=["scenario", "metric", "fuel_bucket", "units", "source_sheet", "source_label"],
            start_year=int(cfg["EXTENSION_START_YEAR"]),
            end_year=int(cfg["EXTENSION_END_YEAR"]),
            lookback_years=int(cfg["EXTENSION_LOOKBACK_YEARS"]),
        )
    aggregation_start_year: int | None = None
    aggregation_start_edition: str | None = None
    if bool(cfg["LIMIT_AGGREGATED_SERIES_START"]):
        aggregation_start_edition = str(cfg["AGGREGATED_SERIES_START_EDITION"])
        aggregation_start_year = int(cfg["AGGREGATED_SERIES_START_YEAR"])
        if not _edition_has_numeric_year(total_long, aggregation_start_edition, aggregation_start_year):
            raise RuntimeError(
                "Configured aggregated series start year was not found in the selected edition. "
                f"Edition={aggregation_start_edition}, year={aggregation_start_year}."
            )

    if not tpes_fuel_h.empty:
        tpes_fuel_agg = _build_aggregated_fuel(
            tpes_fuel_h,
            edition_weights=cfg["EDITION_WEIGHTS"] if bool(cfg["USE_WEIGHTED_EDITION_AVERAGE"]) else None,
            aggregation_start_year=aggregation_start_year,
        )

    source_series_long = pd.concat(
        [
            total_long.assign(fuel_raw="ALL", fuel_bucket="ALL"),
            tpes_long.assign(fuel_raw="ALL", fuel_bucket="ALL"),
            fuel_h.rename(columns={"fuel_bucket": "fuel_bucket"}),
            tpes_fuel_h.rename(columns={"fuel_bucket": "fuel_bucket"}) if not tpes_fuel_h.empty else pd.DataFrame(),
        ],
        ignore_index=True,
        sort=False,
    )
    source_series_long["series_type"] = source_series_long["metric"]

    aggregation_weights = cfg["EDITION_WEIGHTS"] if bool(cfg["USE_WEIGHTED_EDITION_AVERAGE"]) else None
    total_agg = _build_aggregated_total(
        total_long,
        edition_weights=aggregation_weights,
        aggregation_start_year=aggregation_start_year,
    )
    tpes_agg = _build_aggregated_total(
        tpes_long,
        edition_weights=aggregation_weights,
        aggregation_start_year=aggregation_start_year,
    )
    fuel_agg = _build_aggregated_fuel(
        fuel_h,
        edition_weights=aggregation_weights,
        aggregation_start_year=aggregation_start_year,
    )
    if bool(cfg["REBASE_AGGREGATED_SERIES_TO_START_EDITION"]):
        total_agg = _rebase_aggregated_series_to_edition(
            total_agg,
            start_year=aggregation_start_year,
            start_edition=aggregation_start_edition,
        )
        tpes_agg = _rebase_aggregated_series_to_edition(
            tpes_agg,
            start_year=aggregation_start_year,
            start_edition=aggregation_start_edition,
        )
        fuel_agg = _rebase_aggregated_series_to_edition(
            fuel_agg,
            start_year=aggregation_start_year,
            start_edition=aggregation_start_edition,
            group_cols=["fuel"],
        )
        tpes_fuel_agg = _rebase_aggregated_series_to_edition(
            tpes_fuel_agg,
            start_year=aggregation_start_year,
            start_edition=aggregation_start_edition,
            group_cols=["fuel"],
        )

    # Validation checks
    if set(total_long["scenario"].unique()) != {"BAU", "Reference"}:
        raise RuntimeError("Scenario check failed for total series (expected BAU + Reference).")
    _assert_no_duplicates(total_agg, ["year"], "aggregated_total_final_energy_demand")
    _assert_no_duplicates(tpes_agg, ["year"], "aggregated_total_primary_energy_supply")
    _assert_no_duplicates(fuel_agg, ["fuel", "year"], "aggregated_fuel_partial")

    notes = (
        p7.notes
        + p8.notes
        + p9.notes
        + fuel_notes
        + tpes_fuel_notes
        + (optional_9th.notes if optional_9th else [])
        + (optional_9th_tpes.notes if optional_9th_tpes else [])
    )
    notes_df = pd.DataFrame(
        notes,
        columns=["issue_type", "edition", "fuel", "year", "detail", "handling_rule"],
    )
    if not notes_df.empty:
        notes_df = notes_df.drop_duplicates(ignore_index=True)

    # Write outputs
    source_long_path = analysis_dir / "source_series_long.csv"
    total_agg_path = output_dir / "aggregated_total_final_energy_demand.csv"
    tpes_agg_path = output_dir / "aggregated_total_primary_energy_supply.csv"
    fuel_agg_path = output_dir / "aggregated_fuel_partial.csv"
    tpes_fuel_agg_path = output_dir / "aggregated_tpes_fuel_partial.csv"
    notes_path = checks_dir / "data_quality_notes.csv"
    assumptions_path = supporting_dir / "README_assumptions.md"

    source_series_long.to_csv(source_long_path, index=False)
    total_agg.to_csv(total_agg_path, index=False)
    tpes_agg.to_csv(tpes_agg_path, index=False)
    fuel_agg.to_csv(fuel_agg_path, index=False)
    tpes_fuel_agg.to_csv(tpes_fuel_agg_path, index=False)
    notes_df.to_csv(notes_path, index=False)
    _write_assumptions_md(
        assumptions_path,
        used_optional_9th=optional_9th is not None,
        use_weighted_average=bool(cfg["USE_WEIGHTED_EDITION_AVERAGE"]),
        aggregation_start_year=aggregation_start_year,
        aggregation_start_edition=aggregation_start_edition,
        rebase_to_start_edition=bool(cfg["REBASE_AGGREGATED_SERIES_TO_START_EDITION"]),
    )

    written_charts: list[Path] = []
    if bool(cfg["WRITE_CHARTS"]):
        written_charts = _build_charts(
            total_long=total_long,
            tpes_long=tpes_long,
            total_agg=total_agg,
            tpes_agg=tpes_agg,
            fuel_agg=fuel_agg,
            charts_dir=charts_dir,
            backend=str(cfg["CHART_BACKEND"]),
        )

    dashboard_index: Path | None = None
    if bool(cfg["WRITE_DASHBOARD"]):
        dashboard_index = _build_dashboard(
            output_dir=output_dir,
            total_long=total_long,
            tpes_long=tpes_long,
            total_agg=total_agg,
            tpes_agg=tpes_agg,
            fuel_h=fuel_h,
            fuel_agg=fuel_agg,
            tpes_fuel_h=tpes_fuel_h,
            tpes_fuel_agg=tpes_fuel_agg,
        )

    return {
        "source_series_long": str(source_long_path),
        "aggregated_total_final_energy_demand": str(total_agg_path),
        "aggregated_total_primary_energy_supply": str(tpes_agg_path),
        "aggregated_fuel_partial": str(fuel_agg_path),
        "aggregated_tpes_fuel_partial": str(tpes_fuel_agg_path),
        "data_quality_notes": str(notes_path),
        "assumptions": str(assumptions_path),
        "charts_dir": str(charts_dir),
        "dashboard_index": str(dashboard_index) if dashboard_index else "",
    }


if __name__ == "__main__":  # pragma: no cover
    artifacts = run_workflow()
    print("[OK] APERC reference aggregation workflow complete.")
    for key, value in artifacts.items():
        print(f"- {key}: {value}")
#%%


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
