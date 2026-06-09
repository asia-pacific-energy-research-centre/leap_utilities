#%%
"""
Generate fuel rankings for losses and own-use sectors from ESTO data.

Produces outputs/loss_own_use_fuel_rankings.xlsx, structured like
transformation_input_fuel_rankings.xlsx but covering the 10.01.x / 10.02
own-use and loss flow codes for every economy.

Data sources:
  - data/00APEC_2025_low_with_subtotals.csv  (primary, up to 2023)
  - data/00APEC_2024_low_with_subtotals.csv  (secondary cross-check, up to 2022)
"""
from __future__ import annotations

import os
import sys
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
OUTPUT_PATH = REPO_ROOT / "outputs" / "loss_own_use_fuel_rankings.xlsx"

# Years to display as individual columns in the output sheets.
# Both files share 2022; only the 2025 file has 2023.
DISPLAY_YEARS_2025 = [2022, 2023]
DISPLAY_YEARS_2024 = [2022]

# Ranking is based on abs total across DISPLAY_YEARS_2025 from the 2025 file.

# ---------------------------------------------------------------------------
# Sector config: map sector key/label to the 10.01.x / 10.02 flow codes.
# Entries without target flows (e.g. CCS) are skipped automatically.
# ---------------------------------------------------------------------------
SECTOR_CONFIG = [
    {
        "key": "electricity_chp_heat_plants",
        "label": "Electricity, CHP and heat plants",
        "loss_flows": ["10.01.01 Electricity, CHP and heat plants"],
    },
    {
        "key": "gas_works_plants",
        "label": "Gas works plants",
        "loss_flows": ["10.01.02 Gas works plants"],
    },
    {
        "key": "liquefaction_regasification_plants",
        "label": "Liquefaction/regasification plants",
        "loss_flows": ["10.01.03 Liquefaction/regasification plants"],
    },
    {
        "key": "gas_to_liquids_plants",
        "label": "Gas-to-liquids plants",
        "loss_flows": ["10.01.04 Gas-to-liquids plants"],
    },
    {
        "key": "coke_ovens",
        "label": "Coke ovens",
        "loss_flows": ["10.01.05 Coke ovens"],
    },
    {
        "key": "coal_mines",
        "label": "Coal mines",
        "loss_flows": ["10.01.06 Coal mines"],
    },
    {
        "key": "blast_furnaces",
        "label": "Blast furnaces",
        "loss_flows": ["10.01.07 Blast furnaces"],
    },
    {
        "key": "patent_fuel_plants",
        "label": "Patent fuel plants",
        "loss_flows": ["10.01.08 Patent fuel plants"],
    },
    {
        "key": "bkb_pb_plants",
        "label": "BKB/PB plants",
        "loss_flows": ["10.01.09 BKB/PB plants"],
    },
    {
        "key": "liquefaction_coal_to_oil",
        "label": "Liquefaction plants (Coal to Oil)",
        "loss_flows": ["10.01.10 Liquefaction plants (Coal to Oil)"],
    },
    {
        "key": "oil_refineries",
        "label": "Oil refineries",
        "loss_flows": ["10.01.11 Oil refineries"],
    },
    {
        "key": "oil_and_gas_extraction",
        "label": "Oil and gas extraction",
        "loss_flows": ["10.01.12 Oil and gas extraction"],
    },
    {
        "key": "pump_storage_plants",
        "label": "Pump storage plants",
        "loss_flows": ["10.01.13 Pump storage plants"],
    },
    {
        "key": "nuclear_industry",
        "label": "Nuclear industry",
        "loss_flows": ["10.01.14 Nuclear industry"],
    },
    {
        "key": "charcoal_production_plants",
        "label": "Charcoal production plants",
        "loss_flows": ["10.01.15 Charcoal production plants"],
    },
    {
        "key": "gasification_plants_biogases",
        "label": "Gasification plants for biogases",
        "loss_flows": ["10.01.16 Gasification plants for biogases"],
    },
    {
        "key": "nonspecified_own_uses",
        "label": "Non-specified own uses",
        "loss_flows": ["10.01.17 Non-specified own uses"],
    },
    {
        "key": "transmission_distribution_losses",
        "label": "Transmission and distribution losses",
        "loss_flows": ["10.02 Transmission and distribution losses"],
    },
]


def _normalize_economy(value: str) -> str:
    """Convert 01AUS -> 01_AUS and 02BD -> 02_BD style codes."""
    text = str(value or "").strip().upper()
    if len(text) >= 4 and text[:2].isdigit() and "_" not in text:
        return f"{text[:2]}_{text[2:]}"
    return text


def load_esto(path: Path) -> tuple[pd.DataFrame, list[int]]:
    """Load an ESTO CSV, filter out subtotal rows, append 00_APEC total, return (df, year_cols)."""
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


def build_sector_frame(
    df25: pd.DataFrame,
    years25: list[int],
    df24: pd.DataFrame,
    years24: list[int],
    sector_cfg: dict,
    economies: list[str],
) -> pd.DataFrame:
    """Return ranked fuel rows for one sector, covering all economies.

    Fuel eligibility: nonzero in the final year of EITHER source file (2023
    from the 2025 file, 2022 from the 2024 file).  All year columns are
    unified — no _2024file suffix.  For years present in both files the 2025
    value is used; 2024 is the fallback when 2025 is zero.
    """
    flows = sector_cfg["loss_flows"]
    rows25 = _sector_rows(df25, flows)
    rows24 = _sector_rows(df24, flows)

    display25 = [y for y in DISPLAY_YEARS_2025 if y in years25]
    display24 = [y for y in DISPLAY_YEARS_2024 if y in years24]
    all_display_years = sorted(set(display25) | set(display24))

    final_year_25 = max(display25) if display25 else None  # 2023
    final_year_24 = max(display24) if display24 else None  # 2022

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

        qualified = {p: v for p, v in net_totals.items() if abs(v) > 0.0}
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
                "loss_flow_codes": "; ".join(flows),
                "economy": economy,
                "rank": rank_idx,
                "fuel_code": product,
                "fuel_name": product,
                "rank_basis_total_pj": abs(net_total),
                "net_total_pj": net_total,
                "share_of_total_percent": abs(net_total) / total_abs * 100 if total_abs else 0.0,
                "nonzero_year_count": len(nonzero_yrs),
                "first_nonzero_year": float(min(nonzero_yrs)) if nonzero_yrs else None,
                "last_nonzero_year": float(max(nonzero_yrs)) if nonzero_yrs else None,
            }

            # Unified year columns: 2025 value preferred; fall back to 2024
            for y in all_display_years:
                v25 = float(prod_25[y].sum()) if not prod_25.empty and y in prod_25.columns else 0.0
                v24 = float(prod_24[y].sum()) if not prod_24.empty and y in prod_24.columns else 0.0
                row[f"value_{y}_pj"] = v25 if v25 != 0.0 else v24

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
) -> dict:
    row: dict = {
        "sector_key": sector_cfg["key"],
        "sector_title": sector_cfg["label"],
        "loss_flow_codes": "; ".join(sector_cfg["loss_flows"]),
        "economy": economy,
        "rank": None,
        "fuel_code": "(no fuels with nonzero values)",
        "fuel_name": "(no fuels with nonzero values)",
        "rank_basis_total_pj": 0.0,
        "net_total_pj": 0.0,
        "share_of_total_percent": 0.0,
        "nonzero_year_count": 0,
        "first_nonzero_year": None,
        "last_nonzero_year": None,
    }
    for y in display25:
        row[f"value_{y}_pj_2025file"] = 0.0
    for y in display24:
        row[f"value_{y}_pj_2024file"] = 0.0
    return row


def _sheet_name(label: str) -> str:
    """Truncate to Excel 31-char sheet name limit."""
    safe = label.replace("/", " ").replace("\\", " ")
    return safe[:31]


def build_summary(sector_frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Build a cross-sector summary of total abs values per economy."""
    rows: list[dict] = []
    for cfg in SECTOR_CONFIG:
        key = cfg["key"]
        label = cfg["label"]
        df = sector_frames.get(key, pd.DataFrame())
        if df.empty:
            continue
        for economy, grp in df.groupby("economy"):
            total_abs = grp["rank_basis_total_pj"].sum()
            fuel_count = grp[grp["rank"] >= 1].shape[0] if "rank" in grp.columns else 0
            rows.append(
                {
                    "sector_key": key,
                    "sector_title": label,
                    "economy": economy,
                    "nonzero_fuel_count": fuel_count,
                    "total_abs_pj": total_abs,
                }
            )
    return pd.DataFrame(rows)


def main() -> None:
    print(f"Loading {DATA_2025.name} ...")
    df25, years25 = load_esto(DATA_2025)
    print(f"Loading {DATA_2024.name} ...")
    df24, years24 = load_esto(DATA_2024)

    _econs = sorted(e for e in df25["economy"].dropna().unique() if e != "00_APEC")
    economies = ["00_APEC"] + _econs
    print(f"Found {len(economies)} economies: {economies}")

    sector_frames: dict[str, pd.DataFrame] = {}
    all_rows: list[pd.DataFrame] = []

    for cfg in SECTOR_CONFIG:
        print(f"  Processing: {cfg['label']} ...")
        frame = build_sector_frame(df25, years25, df24, years24, cfg, economies)
        sector_frames[cfg["key"]] = frame
        all_rows.append(frame)

    summary_df = build_summary(sector_frames)

    notes_df = pd.DataFrame(
        [
            ("Purpose", "Rank fuels in each loss/own-use sector by economy using ESTO data."),
            ("Primary source", DATA_2025.name),
            ("Secondary source", DATA_2024.name),
            ("Ranking basis", "Abs sum of 2025-file values across " + str(DISPLAY_YEARS_2025)),
            ("Display year columns (2025 file)", str(DISPLAY_YEARS_2025)),
            ("Display year columns (2024 file)", str(DISPLAY_YEARS_2024)),
            ("Subtotals excluded", "Yes — is_subtotal == True rows removed"),
            ("Created", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ],
        columns=["item", "note"],
    )

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    out = OUTPUT_PATH
    try:
        out.open("r+b").close() if out.exists() else None
    except PermissionError:
        out = out.with_stem(out.stem + "_new")
        print(f"  [WARN] output locked — writing to {out.name}")
    print(f"\nWriting {out} ...")
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        notes_df.to_excel(writer, sheet_name="Notes", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        for cfg in SECTOR_CONFIG:
            key = cfg["key"]
            frame = sector_frames.get(key, pd.DataFrame())
            if frame.empty:
                continue
            sheet = _sheet_name(cfg["label"])
            frame.to_excel(writer, sheet_name=sheet, index=False)

    print(f"Done — {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
