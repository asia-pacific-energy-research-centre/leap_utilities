#%%
"""
GDP-intensity comparison for minor demand sectors (experimental).

Tests whether using GDP as the Activity Level driver produces results close to
the 9th projection (TGT) for Agriculture, Fishing, and Non-specified others.

Approach
--------
- Activity Level = GDP (indexed: base year = 1.0 after normalising)
- Final Energy Intensity = esto_base_year_energy  (since indexed GDP base = 1)
- Projected energy[year] = GDP_index[year] * base_year_energy
- TGT = allocated 9th projection (current minor_demand_workflow output)

GDP file: data/9th_macro_data.csv
  Columns: Economy, Scenario, Date, Population, Gdp, Gdp_per_capita
  Long format — one row per economy / scenario / year.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
CURRENT_DIR = Path.cwd()
if CURRENT_DIR != REPO_ROOT:
    os.chdir(REPO_ROOT)
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.functions.ninth_projection_mapping import normalize_economy_key
from codebase.minor_demand_workflow import (
    EXPORT_BASE_YEAR,
    MINOR_DEMAND_FLOW_CONFIG,
    NINTH_SCENARIO,
    _projection_years,
    build_allocated_projection_table,
    build_base_year_activity_value,
    filter_mapping_for_minor_demand,
    load_esto_data,
    load_ninth_data,
    load_mapping,
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

GDP_DATA_PATH = "data/9th_macro_data.csv"

# GDP scenario to use from the macro data ("Reference" or "Target").
GDP_SCENARIO = "Reference"

# Economies to run. USA + 5 others covering a range of APEC members.
COMPARISON_ECONOMIES = [
    "20_USA",   # United States
    "05_PRC",   # China
    "08_JPN",   # Japan
    "01_AUS",   # Australia
    "09_ROK",   # Korea
    "11_MEX",   # Mexico
]

# Which years to show in the printed summary table.
DISPLAY_YEARS = [2022, 2025, 2030, 2035, 2040, 2045, 2050, 2060]

# Where to save the full comparison CSV.
OUTPUT_CSV_PATH = "data/temp/gdp_intensity_comparison.csv"


# ---------------------------------------------------------------------------
# GDP loading  (long-format: Economy / Scenario / Date / Gdp)
# ---------------------------------------------------------------------------

def load_gdp_wide(
    path: str,
    scenario: str = GDP_SCENARIO,
) -> pd.DataFrame:
    """
    Load the macro CSV and return a wide DataFrame indexed by economy_key,
    with int year columns and Gdp values.
    """
    df = pd.read_csv(path)
    df.columns = [str(c).strip() for c in df.columns]

    required = {"Economy", "Scenario", "Date", "Gdp"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"GDP file missing columns: {missing}. Found: {list(df.columns)}")

    df = df[df["Scenario"] == scenario].copy()
    if df.empty:
        raise ValueError(
            f"No rows for scenario '{scenario}' in '{path}'. "
            f"Available: {pd.read_csv(path)['Scenario'].unique().tolist()}"
        )

    df["economy_key"] = df["Economy"].apply(normalize_economy_key)
    df["year"] = pd.to_numeric(df["Date"], errors="coerce").astype("Int64")
    df["Gdp"] = pd.to_numeric(df["Gdp"], errors="coerce")
    df = df.dropna(subset=["year", "Gdp"])

    wide = df.pivot_table(index="economy_key", columns="year", values="Gdp", aggfunc="first")
    wide.columns = [int(c) for c in wide.columns]
    return wide


def get_gdp_series(wide: pd.DataFrame, economy_key: str) -> pd.Series:
    """Return a year-indexed GDP Series for one economy."""
    if economy_key not in wide.index:
        raise ValueError(
            f"Economy '{economy_key}' not found in GDP data. "
            f"Available: {list(wide.index)}"
        )
    return wide.loc[economy_key].dropna().astype(float)


def build_gdp_index(gdp_series: pd.Series, base_year: int) -> pd.Series:
    """Return GDP indexed to base_year = 1.0."""
    base = gdp_series.get(base_year)
    if base is None or base == 0:
        raise ValueError(f"GDP base-year ({base_year}) value is zero or missing for this economy.")
    return gdp_series / base


# ---------------------------------------------------------------------------
# TGT: 9th allocated projection
# ---------------------------------------------------------------------------

def build_ninth_tgt(
    esto_data: pd.DataFrame,
    ninth_data: pd.DataFrame,
    mapping: pd.DataFrame,
    economy_key: str,
    projection_years: list[int],
) -> dict[str, dict[int, float]]:
    """Return 9th-allocated projection energy (PJ) by flow: {esto_flow: {year: PJ}}."""
    allocated = build_allocated_projection_table(
        ninth_data=ninth_data,
        esto_data=esto_data,
        projection_years=projection_years,
        scenario=NINTH_SCENARIO,
        mapping=mapping,
    )
    result: dict[str, dict[int, float]] = {}
    for cfg in MINOR_DEMAND_FLOW_CONFIG:
        flow = cfg["esto_flow"]
        if allocated is None or allocated.empty:
            result[flow] = {year: 0.0 for year in projection_years}
            continue
        year_cols = [y for y in projection_years if y in allocated.columns]
        subset = allocated[
            (allocated["economy_key"] == economy_key) & (allocated["esto_flow"] == flow)
        ]
        if subset.empty:
            result[flow] = {year: 0.0 for year in projection_years}
        else:
            totals = subset[year_cols].sum().to_dict()
            result[flow] = {int(y): float(v) for y, v in totals.items()}
    return result


# ---------------------------------------------------------------------------
# GDP-intensity energy projection
# ---------------------------------------------------------------------------

def project_energy_from_gdp_index(
    base_year_energy: float,
    gdp_index: pd.Series,
    projection_years: list[int],
) -> dict[int, float]:
    """energy[year] = base_year_energy * gdp_index[year]."""
    result = {}
    for year in projection_years:
        idx = gdp_index.get(year)
        result[year] = base_year_energy * float(idx) if idx is not None and not pd.isna(idx) else float("nan")
    return result


# ---------------------------------------------------------------------------
# Per-economy comparison table
# ---------------------------------------------------------------------------

def build_economy_comparison(
    economy_key: str,
    esto_data: pd.DataFrame,
    ninth_data: pd.DataFrame,
    mapping: pd.DataFrame,
    gdp_wide: pd.DataFrame,
    projection_years: list[int],
    base_year: int = EXPORT_BASE_YEAR,
) -> pd.DataFrame:
    gdp_series = get_gdp_series(gdp_wide, economy_key)
    gdp_index = build_gdp_index(gdp_series, base_year)
    ninth_tgt = build_ninth_tgt(esto_data, ninth_data, mapping, economy_key, projection_years)

    rows = []
    for cfg in MINOR_DEMAND_FLOW_CONFIG:
        flow = cfg["esto_flow"]
        label = cfg["sector_label"]
        base_energy = build_base_year_activity_value(esto_data, flow, economy_key, base_year)
        gdp_proj = project_energy_from_gdp_index(base_energy, gdp_index, projection_years)
        ninth_proj = ninth_tgt.get(flow, {})

        for year in projection_years:
            gdp_val = gdp_proj.get(year, float("nan"))
            ninth_val = ninth_proj.get(year, 0.0)
            diff = gdp_val - ninth_val if not pd.isna(gdp_val) else float("nan")
            diff_pct = (diff / ninth_val * 100) if ninth_val != 0 and not pd.isna(diff) else float("nan")
            rows.append({
                "economy_key": economy_key,
                "flow": flow,
                "sector_label": label,
                "year": int(year),
                "gdp_energy_pj": round(gdp_val, 4) if not pd.isna(gdp_val) else float("nan"),
                "ninth_energy_pj": round(ninth_val, 4),
                "diff_pj": round(diff, 4) if not pd.isna(diff) else float("nan"),
                "diff_pct": round(diff_pct, 2) if not pd.isna(diff_pct) else float("nan"),
            })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Display
# ---------------------------------------------------------------------------

def print_economy_summary(df: pd.DataFrame, display_years: list[int]) -> None:
    """Print a compact block for one economy."""
    economy_key = df["economy_key"].iloc[0]
    subset = df[df["year"].isin(display_years)].copy()

    year_header = "  ".join(f"{y:>8}" for y in display_years)
    print(f"\n{'Flow':<28}  {year_header}")
    print("-" * (30 + 10 * len(display_years)))

    for cfg in MINOR_DEMAND_FLOW_CONFIG:
        label = cfg["sector_label"]
        flow_data = subset[subset["sector_label"] == label].set_index("year")

        def fmt(col, year, fmt_str="{:>8.2f}"):
            if year not in flow_data.index:
                return f"{'n/a':>8}"
            v = flow_data.loc[year, col]
            return fmt_str.format(v) if not pd.isna(v) else f"{'n/a':>8}"

        gdp_row  = "  ".join(fmt("gdp_energy_pj", y)  for y in display_years)
        ninth_row = "  ".join(fmt("ninth_energy_pj", y) for y in display_years)
        pct_row  = "  ".join(fmt("diff_pct", y, "{:>7.1f}%") for y in display_years)

        print(f"\n  {label}")
        print(f"    {'GDP-derived (PJ)':<22}  {gdp_row}")
        print(f"    {'9th TGT (PJ)':<22}  {ninth_row}")
        print(f"    {'diff %':<22}  {pct_row}")


def print_comparison_report(
    all_dfs: dict[str, pd.DataFrame],
    display_years: list[int] = DISPLAY_YEARS,
) -> None:
    print("\n" + "=" * 72)
    print("GDP-intensity vs 9th projection — minor demand sectors")
    print(f"Base year: {EXPORT_BASE_YEAR}  |  GDP scenario: {GDP_SCENARIO}")
    print(f"Intensity locked at base year; projected = GDP_index * base_energy")
    print("=" * 72)

    for economy_key, df in all_dfs.items():
        if df.empty:
            continue
        print(f"\n{'-' * 72}")
        print(f"  Economy: {economy_key}")
        print_economy_summary(df, display_years)

    print("\n" + "=" * 72)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def run_gdp_intensity_comparison(
    economies: list[str] = COMPARISON_ECONOMIES,
    gdp_data_path: str = GDP_DATA_PATH,
    display_years: list[int] = DISPLAY_YEARS,
    save_csv: bool = True,
) -> pd.DataFrame:
    projection_years = _projection_years()

    print("Loading shared datasets...")
    esto_data = load_esto_data()
    ninth_data = load_ninth_data()
    mapping = load_mapping()
    mapping = filter_mapping_for_minor_demand(mapping, esto_data, MINOR_DEMAND_FLOW_CONFIG)

    print(f"Loading GDP data from '{gdp_data_path}'...")
    gdp_wide = load_gdp_wide(gdp_data_path, scenario=GDP_SCENARIO)

    all_dfs: dict[str, pd.DataFrame] = {}
    for economy in economies:
        economy_key = normalize_economy_key(economy)
        print(f"\nBuilding comparison for {economy_key}...")
        try:
            df = build_economy_comparison(
                economy_key=economy_key,
                esto_data=esto_data,
                ninth_data=ninth_data,
                mapping=mapping,
                gdp_wide=gdp_wide,
                projection_years=projection_years,
            )
            all_dfs[economy_key] = df
        except Exception as exc:
            print(f"  [ERROR] {economy_key}: {exc}")

    print_comparison_report(all_dfs, display_years=display_years)

    combined = pd.concat(all_dfs.values(), ignore_index=True) if all_dfs else pd.DataFrame()
    if save_csv and not combined.empty:
        Path(OUTPUT_CSV_PATH).parent.mkdir(parents=True, exist_ok=True)
        combined.to_csv(OUTPUT_CSV_PATH, index=False)
        print(f"\n[INFO] Full comparison saved to: {OUTPUT_CSV_PATH}")

    return combined


#%%
if __name__ == "__main__":
    result_df = run_gdp_intensity_comparison()
