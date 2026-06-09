#%%
"""
Generate input and output fuel rankings for ESTO transformation sectors.

Writes:
  outputs/transformation_input_fuel_rankings.xlsx   (negative values = feedstock inputs)
  outputs/transformation_output_fuel_rankings.xlsx  (positive values = produced outputs)

Data sources:
  data/00APEC_2025_low_with_subtotals.csv  (primary, up to 2023)
  data/00APEC_2024_low_with_subtotals.csv  (secondary / cross-check, up to 2022)

LNG (09.06.02) and hydrogen (09.13) are 9th-edition sectors and are excluded.
"""
from __future__ import annotations

import os
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if Path.cwd() != REPO_ROOT:
    os.chdir(REPO_ROOT)
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
DATA_2025 = REPO_ROOT / "data" / "00APEC_2025_low_with_subtotals.csv"
DATA_2024 = REPO_ROOT / "data" / "00APEC_2024_low_with_subtotals.csv"
OUTPUT_INPUTS = REPO_ROOT / "outputs" / "transformation_input_fuel_rankings.xlsx"
OUTPUT_OUTPUTS = REPO_ROOT / "outputs" / "transformation_output_fuel_rankings.xlsx"

def _safe_output_path(path: Path) -> Path:
    """If the target is locked, write to a _new suffixed copy instead."""
    if not path.exists():
        return path
    try:
        import io
        path.open("r+b").close()
        return path
    except PermissionError:
        stem = path.stem + "_new"
        alt = path.with_stem(stem)
        print(f"  [WARN] {path.name} is locked — writing to {alt.name} instead")
        return alt

# Year columns to display. These are taken from the 2025 file; 2022 is also
# cross-checked against the 2024 file. Projection years are no longer
# included because the _with_subtotals files are historical snapshots only.
DISPLAY_YEARS_2025 = [2022, 2023]
DISPLAY_YEARS_2024 = [2022]

# ---------------------------------------------------------------------------
# Sector config  (ESTO transformation flow codes from MAJOR_SECTOR_CONFIG,
# excluding LNG/hydrogen which are 9th-edition datasets)
# ---------------------------------------------------------------------------
SECTOR_CONFIG = [
    {
        "key": "gas_works",
        "label": "Gas works plants",
        "flow_codes": ["09.06.01 Gas works plants"],
    },
    {
        "key": "gas_blending",
        "label": "Natural gas blending plants",
        "flow_codes": ["09.06.03 Natural gas blending plants"],
    },
    {
        "key": "gas_to_liquids_plants",
        "label": "Gas to liquids plants",
        "flow_codes": ["09.06.04 Gas-to-liquids plants"],
    },
    {
        "key": "oil_refineries",
        "label": "Oil refineries",
        "flow_codes": ["09.07 Oil refineries"],
    },
    {
        "key": "coal_coke_ovens",
        "label": "Coke ovens",
        "flow_codes": ["09.08.01 Coke ovens"],
    },
    {
        "key": "coal_blast_furnaces",
        "label": "Blast furnaces",
        "flow_codes": ["09.08.02 Blast furnaces"],
    },
    {
        "key": "coal_patent_fuel_plants",
        "label": "Patent fuel plants",
        "flow_codes": ["09.08.03 Patent fuel plants"],
    },
    {
        "key": "coal_bkb_pb_plants",
        "label": "BKB PB plants",
        "flow_codes": ["09.08.04 BKB/PB plants"],
    },
    {
        "key": "coal_liquefaction",
        "label": "Liquefaction (coal to oil)",
        "flow_codes": ["09.08.05 Liquefaction (coal to oil)"],
    },
    {
        "key": "coal_mines",
        "label": "Coal mines",
        "flow_codes": ["09.08.06 Coal mines"],
    },
    {
        "key": "electric_boilers",
        "label": "Electric boilers",
        "flow_codes": ["09.04 Electric boilers"],
    },
    {
        "key": "chemical_heat_for_electricity_production",
        "label": "Chemical heat for electricity p",
        "flow_codes": ["09.05 Chemical heat for electricity production"],
    },
    {
        "key": "petrochemical_industry",
        "label": "Petrochemical industry",
        "flow_codes": ["09.09 Petrochemical industry"],
    },
    {
        "key": "biofuels_processing",
        "label": "Biofuels processing",
        "flow_codes": ["09.10 Biofuels processing"],
    },
    {
        "key": "charcoal_processing",
        "label": "Charcoal processing",
        "flow_codes": ["09.11 Charcoal processing"],
    },
    {
        "key": "nonspecified_transformation",
        "label": "Non-specified transformation",
        "flow_codes": ["09.12 Non-specified transformation"],
    },
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalize_economy(value: str) -> str:
    text = str(value or "").strip().upper()
    if len(text) >= 4 and text[:2].isdigit() and "_" not in text:
        return f"{text[:2]}_{text[2:]}"
    return text


def load_esto(path: Path) -> tuple[pd.DataFrame, list[int]]:
    """Load an ESTO CSV, drop subtotal rows, normalize economies, append 00_APEC total."""
    df = pd.read_csv(path, low_memory=False)
    year_cols = sorted(int(c) for c in df.columns if str(c).isdigit())
    df.columns = [int(c) if str(c).isdigit() else c for c in df.columns]
    if "is_subtotal" in df.columns:
        df = df[df["is_subtotal"] == False].copy()
    df["economy"] = df["economy"].apply(_normalize_economy)
    df = _append_apec_total(df, year_cols)
    return df, year_cols


def _append_apec_total(df: pd.DataFrame, year_cols: list[int]) -> pd.DataFrame:
    """Sum all individual economies into a synthetic 00_APEC row set."""
    if "00_APEC" in df["economy"].values:
        return df
    meta_cols = [c for c in df.columns if c not in year_cols and c != "economy"]
    apec = df.groupby(meta_cols, dropna=False)[year_cols].sum().reset_index()
    apec["economy"] = "00_APEC"
    return pd.concat([df, apec], ignore_index=True)


def _sector_rows(df: pd.DataFrame, flows: list[str]) -> pd.DataFrame:
    if "flows" not in df.columns:
        return df.iloc[0:0]
    return df[df["flows"].isin(flows)].copy()


def _economy_order(df25: pd.DataFrame) -> list[str]:
    """Return economies sorted with 00_APEC first."""
    econs = sorted(e for e in df25["economy"].dropna().unique() if e != "00_APEC")
    return ["00_APEC"] + econs


def build_sector_frame(
    df25: pd.DataFrame,
    years25: list[int],
    df24: pd.DataFrame,
    years24: list[int],
    sector_cfg: dict,
    economies: list[str],
    mode: str,  # "inputs" or "outputs"
) -> pd.DataFrame:
    """Return ranked fuel rows for one sector × all economies.

    Fuel eligibility: a product is included if it has a nonzero value in the
    final year of EITHER source file (2023 from the 2025 file, 2022 from the
    2024 file).  All year columns are unified — no _2024file suffix.  For
    years present in both files the 2025 value is used; the 2024 value is the
    fallback when the 2025 value is zero.
    """
    flows = sector_cfg["flow_codes"]
    rows25 = _sector_rows(df25, flows)
    rows24 = _sector_rows(df24, flows)

    display25 = [y for y in DISPLAY_YEARS_2025 if y in years25]
    display24 = [y for y in DISPLAY_YEARS_2024 if y in years24]
    # Unified display years — no duplicates, sorted ascending
    all_display_years = sorted(set(display25) | set(display24))

    final_year_25 = max(display25) if display25 else None  # 2023
    final_year_24 = max(display24) if display24 else None  # 2022

    value_col = "input" if mode == "inputs" else "output"
    share_col = f"share_of_{value_col}_percent"
    rank_col = f"rank_basis_{value_col}_total_pj"

    records: list[dict] = []

    for economy in economies:
        econ_25 = rows25[rows25["economy"] == economy]
        econ_24 = rows24[rows24["economy"] == economy]

        # --- Eligible fuels: nonzero in final year of either file ---
        fuels_25: set[str] = set()
        if final_year_25 is not None and not econ_25.empty and final_year_25 in econ_25.columns:
            s = econ_25.groupby("products")[final_year_25].sum()
            fuels_25 = set(s[s.abs() > 0].index)

        fuels_24: set[str] = set()
        if final_year_24 is not None and not econ_24.empty and final_year_24 in econ_24.columns:
            s = econ_24.groupby("products")[final_year_24].sum()
            fuels_24 = set(s[s.abs() > 0].index)

        eligible = fuels_25 | fuels_24
        if not eligible:
            continue

        # --- Net totals for ranking: prefer 2025, fall back to 2024 ---
        nt25: dict[str, float] = {}
        if not econ_25.empty and display25:
            nt25 = econ_25.groupby("products")[display25].sum().sum(axis=1).to_dict()

        nt24: dict[str, float] = {}
        if not econ_24.empty and display24:
            nt24 = econ_24.groupby("products")[display24].sum().sum(axis=1).to_dict()

        net_totals: dict[str, float] = {
            prod: (nt25[prod] if nt25.get(prod, 0.0) != 0.0 else nt24.get(prod, 0.0))
            for prod in eligible
        }

        if mode == "inputs":
            qualified = {p: v for p, v in net_totals.items() if v < 0.0}
        else:
            qualified = {p: v for p, v in net_totals.items() if v > 0.0}

        if not qualified:
            continue

        total_abs = sum(abs(v) for v in qualified.values())
        ranked = sorted(qualified.items(), key=lambda kv: abs(kv[1]), reverse=True)

        for rank_idx, (product, net_total) in enumerate(ranked, start=1):
            prod_25 = econ_25[econ_25["products"] == product] if not econ_25.empty else pd.DataFrame()
            prod_24 = econ_24[econ_24["products"] == product] if not econ_24.empty else pd.DataFrame()

            # Nonzero year count: prefer 2025 history, fall back to 2024
            if not prod_25.empty:
                all_yr_vals = prod_25[years25].sum()
                nonzero_yrs = [y for y in years25 if abs(float(all_yr_vals.get(y, 0.0))) > 0.0]
            elif not prod_24.empty:
                all_yr_vals = prod_24[years24].sum()
                nonzero_yrs = [y for y in years24 if abs(float(all_yr_vals.get(y, 0.0))) > 0.0]
            else:
                nonzero_yrs = []

            row: dict = {
                "sector_key": sector_cfg["key"],
                "sector_title": sector_cfg["label"],
                "flow_codes": "; ".join(flows),
                "economy": economy,
                "rank": rank_idx,
                "fuel_code": product,
                "fuel_name": product,
                rank_col: abs(net_total),
                "net_total_pj": net_total,
                share_col: abs(net_total) / total_abs * 100 if total_abs else 0.0,
                "nonzero_year_count": len(nonzero_yrs),
                "first_nonzero_year": float(min(nonzero_yrs)) if nonzero_yrs else None,
                "last_nonzero_year": float(max(nonzero_yrs)) if nonzero_yrs else None,
            }

            # Unified year columns: 2025 value preferred; fall back to 2024
            for y in all_display_years:
                v25 = float(prod_25[y].sum()) if not prod_25.empty and y in prod_25.columns else 0.0
                v24 = float(prod_24[y].sum()) if not prod_24.empty and y in prod_24.columns else 0.0
                row[f"{value_col}_{y}_pj"] = v25 if v25 != 0.0 else v24

            records.append(row)

    df = pd.DataFrame(records)
    if not df.empty:
        df = df.sort_values(["economy", "rank", "fuel_name"], kind="mergesort").reset_index(drop=True)
    return df


def _empty_row(
    sector_cfg: dict,
    economy: str,
    display25: list[int],
    display24: list[int],
    mode: str,
) -> dict:
    value_col = "input" if mode == "inputs" else "output"
    share_col = f"share_of_{value_col}_percent"
    rank_col = f"rank_basis_{value_col}_total_pj"
    placeholder = "(no negative inputs)" if mode == "inputs" else "(no positive outputs)"
    row: dict = {
        "sector_key": sector_cfg["key"],
        "sector_title": sector_cfg["label"],
        "flow_codes": "; ".join(sector_cfg["flow_codes"]),
        "economy": economy,
        "rank": None,
        "fuel_code": placeholder,
        "fuel_name": placeholder,
        rank_col: 0.0,
        "net_total_pj": 0.0,
        share_col: 0.0,
        "nonzero_year_count": 0,
        "first_nonzero_year": None,
        "last_nonzero_year": None,
    }
    for y in display25:
        row[f"{value_col}_{y}_pj"] = 0.0
    for y in display24:
        row[f"{value_col}_{y}_pj_2024file"] = 0.0
    return row


def _sheet_name(label: str) -> str:
    return label[:31]


def build_summary(sector_frames: dict[str, pd.DataFrame], mode: str) -> pd.DataFrame:
    value_col = "input" if mode == "inputs" else "output"
    rank_col = f"rank_basis_{value_col}_total_pj"
    placeholder = "(no negative inputs)" if mode == "inputs" else "(no positive outputs)"
    label = "inputs" if mode == "inputs" else "outputs"

    rows: list[dict] = []
    for cfg in SECTOR_CONFIG:
        key = cfg["key"]
        frame = sector_frames.get(key, pd.DataFrame())
        if frame.empty:
            continue
        active = frame[frame["rank"] >= 1].copy() if "rank" in frame.columns else pd.DataFrame()
        economies_with_values = int(active["economy"].nunique()) if not active.empty else 0
        rank_1 = active[active["rank"] == 1]["fuel_code"].dropna().tolist() if not active.empty else []
        rank_1_clean = [f for f in rank_1 if f != placeholder]
        unique_rank_1 = len(set(rank_1_clean))
        if rank_1_clean:
            most_common_fuel = Counter(rank_1_clean).most_common(1)[0][0]
            most_common_econs = sorted(
                active.loc[
                    (active["rank"] == 1) & (active["fuel_code"] == most_common_fuel),
                    "economy",
                ].tolist()
            )
        else:
            most_common_fuel = ""
            most_common_econs = []

        rows.append(
            {
                "sector_key": key,
                "sector_title": cfg["label"],
                "flow_codes": "; ".join(cfg["flow_codes"]),
                f"economies_with_{label}": economies_with_values,
                f"unique_rank_1_fuels": unique_rank_1,
                f"most_common_rank_1_fuel": most_common_fuel,
                f"most_common_rank_1_economies": "; ".join(most_common_econs),
                "sheet": _sheet_name(cfg["label"]),
            }
        )
    return pd.DataFrame(rows)


def build_notes(mode: str) -> pd.DataFrame:
    value_col = "input" if mode == "inputs" else "output"
    sign_desc = "Negative net totals are treated as input fuels; rank by absolute value." if mode == "inputs" \
        else "Positive net totals are treated as output fuels; rank by value."
    return pd.DataFrame(
        [
            ("Purpose", f"Rank ESTO transformation {value_col} fuels by economy and sector."),
            ("Ranking basis", sign_desc),
            ("Ranking years", f"Sum across {DISPLAY_YEARS_2025} from the 2025 file."),
            ("Primary source", DATA_2025.name),
            ("Secondary source (cross-check)", DATA_2024.name),
            ("Subtotals excluded", "Yes — is_subtotal == True rows removed from both files"),
            (
                "Included sectors",
                "ESTO flow-based transformation configs. "
                "LNG (09.06.02) and hydrogen (09.13) are 9th-edition sectors and are excluded.",
            ),
            ("Created", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ],
        columns=["item", "note"],
    )


def write_rankings(
    df25: pd.DataFrame,
    years25: list[int],
    df24: pd.DataFrame,
    years24: list[int],
    economies: list[str],
    mode: str,
    output_path: Path,
) -> None:
    sector_frames: dict[str, pd.DataFrame] = {}
    for cfg in SECTOR_CONFIG:
        print(f"  [{mode}] {cfg['label']} ...")
        frame = build_sector_frame(df25, years25, df24, years24, cfg, economies, mode)
        sector_frames[cfg["key"]] = frame

    summary_df = build_summary(sector_frames, mode)
    notes_df = build_notes(mode)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        notes_df.to_excel(writer, sheet_name="Notes", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        for cfg in SECTOR_CONFIG:
            frame = sector_frames.get(cfg["key"], pd.DataFrame())
            if frame.empty:
                continue
            frame.to_excel(writer, sheet_name=_sheet_name(cfg["label"]), index=False)

    print(f"  Written: {output_path}")


def main() -> None:
    print(f"Loading {DATA_2025.name} ...")
    df25, years25 = load_esto(DATA_2025)
    print(f"Loading {DATA_2024.name} ...")
    df24, years24 = load_esto(DATA_2024)

    economies = _economy_order(df25)
    print(f"Economies ({len(economies)}): {economies}")

    print("\nBuilding input fuel rankings ...")
    write_rankings(df25, years25, df24, years24, economies, "inputs", _safe_output_path(OUTPUT_INPUTS))

    print("\nBuilding output fuel rankings ...")
    write_rankings(df25, years25, df24, years24, economies, "outputs", _safe_output_path(OUTPUT_OUTPUTS))

    print("\nDone.")


if __name__ == "__main__":
    main()
