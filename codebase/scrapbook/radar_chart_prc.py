#%%
"""
Build final energy structure radar charts from the merged APERC-style CSV.

Outputs
-------
1. One PNG per economy-scenario pair
2. One CSV with the underlying fuel-share table used for plotting

Notes
-----
- Final energy is taken as:
    14_industry_sector
    15_transport_sector
    16_other_sector

- Fuel buckets are aggregated to:
    Coal
    Oil
    Nature gas
    Electricity
    Other
"""

import re
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


#%%
CATEGORY_ORDER = ["Coal", "Oil", "Nature gas", "Electricity", "Other"]

FINAL_ENERGY_SECTORS = [
    "14_industry_sector",
    "15_transport_sector",
    "16_other_sector",
]

FUEL_BUCKET_MAP = {
    "01_coal": "Coal",
    "02_coal_products": "Coal",
    "03_peat": "Coal",
    "04_peat_products": "Coal",
    "05_oil_shale_and_oil_sands": "Coal",
    "06_crude_oil_and_ngl": "Oil",
    "07_petroleum_products": "Oil",
    "08_gas": "Nature gas",
    "17_electricity": "Electricity",
    "17_x_green_electricity": "Electricity",
}

DEFAULT_LINE_COLORS = [
    "#ed7d31",  # orange
    "#4472c4",  # blue
    "#a5a5a5",  # grey
    "#00b050",  # green
    "#7030a0",  # purple
]


#%%
def load_energy_file(file_path: str) -> pd.DataFrame:
    """Load the merged CSV file."""
    try:
        df = pd.read_csv(file_path)
        print(f"Loaded file: {file_path}")
        print(f"Shape: {df.shape}")
        return df
    except Exception as exc:
        print(f"Failed to load file: {file_path}")
        raise exc


def get_year_columns(df: pd.DataFrame) -> list[str]:
    """Return columns that are 4-digit years."""
    year_cols = []

    for col in df.columns:
        col_str = str(col)
        if col_str.isdigit() and len(col_str) == 4:
            year_cols.append(col_str)

    year_cols = sorted(year_cols, key=int)
    return year_cols


def map_fuel_to_bucket(fuel: str) -> str:
    """
    Map detailed fuel labels into the 5 chart buckets.

    Everything not explicitly mapped becomes 'Other'.
    """
    fuel = str(fuel)
    return FUEL_BUCKET_MAP.get(fuel, "Other")


def sanitize_filename(text: str) -> str:
    """Make a safe filename."""
    text = str(text).strip()
    text = re.sub(r"[^\w\-\.]+", "_", text)
    text = re.sub(r"_+", "_", text)
    return text.strip("_")


def polar_to_cartesian(radius: np.ndarray, angles_rad: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    """Convert radius-angle coordinates to x-y coordinates."""
    x = radius * np.cos(angles_rad)
    y = radius * np.sin(angles_rad)
    return x, y


def close_polygon(values: np.ndarray) -> np.ndarray:
    """Close a radar polygon by repeating the first value."""
    return np.append(values, values[0])


def prepare_final_energy_rows(df: pd.DataFrame, selected_years: list[int]) -> pd.DataFrame:
    """
    Keep only the final energy sectors and selected years,
    then map fuels into plotting buckets.
    """
    year_cols = [str(year) for year in selected_years if str(year) in df.columns]

    if not year_cols:
        raise ValueError("None of the selected years were found in the file.")

    required_cols = ["economy", "scenarios", "sectors", "fuels", *year_cols]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {missing_cols}")

    out_df = df.loc[df["sectors"].isin(FINAL_ENERGY_SECTORS), required_cols].copy()

    for col in year_cols:
        out_df[col] = pd.to_numeric(out_df[col], errors="coerce").fillna(0.0)

    out_df["fuel_bucket"] = out_df["fuels"].map(map_fuel_to_bucket)

    print(f"Rows kept for final energy sectors: {len(out_df):,}")
    return out_df


def build_share_table_for_combo(
    combo_df: pd.DataFrame,
    selected_years: list[int],
) -> pd.DataFrame:
    """
    Build a share table for one economy-scenario pair.

    Returns
    -------
    pd.DataFrame
        index   = CATEGORY_ORDER
        columns = selected year strings
        values  = percentage shares
    """
    year_cols = [str(year) for year in selected_years if str(year) in combo_df.columns]

    grouped = (
        combo_df.groupby("fuel_bucket", as_index=True)[year_cols]
        .sum()
        .reindex(CATEGORY_ORDER)
        .fillna(0.0)
    )

    totals = grouped.sum(axis=0)
    share_df = grouped.div(totals.replace(0, np.nan), axis=1) * 100
    share_df = share_df.fillna(0.0)

    return share_df

#%%
def add_apec_aggregate_economy(
    df: pd.DataFrame,
    aggregate_economy_code: str = "00_apec",
    economy_col: str = "economy",
    scenario_col: str = "scenarios",
) -> pd.DataFrame:
    """
    Create an aggregate economy equal to the sum of all economies by scenario
    and every other non-year structural column, then append it back onto the
    dataframe.

    What it does
    ------------
    - Detects year columns automatically using get_year_columns(df)
    - Excludes any existing aggregate row matching aggregate_economy_code
      so the function is safe to rerun
    - Groups by all non-year columns except the economy column
    - Sums all year values
    - Sets economy = aggregate_economy_code
    - Appends the aggregate rows back to the original dataframe

    This is designed to work on the merged energy file where labels are spread
    across columns such as scenarios, sectors, fuels, subfuels, etc.
    """
    year_cols = get_year_columns(df)

    if not year_cols:
        raise ValueError("No year columns found, so the APEC aggregate cannot be built.")

    if economy_col not in df.columns:
        raise ValueError(f"Missing required economy column: {economy_col}")

    if scenario_col not in df.columns:
        raise ValueError(f"Missing required scenario column: {scenario_col}")

    # Remove any pre-existing aggregate rows so rerunning does not double count
    base_df = df.loc[df[economy_col] != aggregate_economy_code].copy()

    # All structural columns except economy become grouping columns
    group_cols = [col for col in base_df.columns if col not in year_cols + [economy_col]]

    # Make sure year columns are numeric before summing
    for col in year_cols:
        base_df[col] = pd.to_numeric(base_df[col], errors="coerce").fillna(0.0)

    apec_df = (
        base_df.groupby(group_cols, dropna=False, as_index=False)[year_cols]
        .sum()
    )

    apec_df[economy_col] = aggregate_economy_code

    # Match original column order
    apec_df = apec_df[df.columns.tolist()].copy()

    out_df = pd.concat([base_df, apec_df], ignore_index=True)

    print(f"Added aggregate economy: {aggregate_economy_code}")
    print(f"Original rows (without existing aggregate): {len(base_df):,}")
    print(f"Aggregate rows added: {len(apec_df):,}")
    print(f"Final rows: {len(out_df):,}")

    return out_df


#%%
def build_all_share_tables(
    df: pd.DataFrame,
    selected_years: list[int],
) -> tuple[dict[tuple[str, str], pd.DataFrame], pd.DataFrame]:
    """
    Build one share table per economy-scenario pair and also return
    a long-format combined table for QA/checking.
    """
    combos = (
        df[["economy", "scenarios"]]
        .drop_duplicates()
        .sort_values(["economy", "scenarios"])
        .itertuples(index=False, name=None)
    )

    share_tables = {}
    share_table_long_parts = []

    for economy, scenario in combos:
        combo_df = df.loc[
            (df["economy"] == economy) & (df["scenarios"] == scenario)
        ].copy()

        share_df = build_share_table_for_combo(combo_df, selected_years)
        share_tables[(economy, scenario)] = share_df

        temp = (
            share_df.reset_index()
            .rename(columns={"index": "fuel_bucket"})
            .melt(
                id_vars="fuel_bucket",
                var_name="year",
                value_name="share_percent",
            )
        )
        temp["economy"] = economy
        temp["scenario"] = scenario
        share_table_long_parts.append(temp)

    share_table_long = pd.concat(share_table_long_parts, ignore_index=True)

    share_table_long = share_table_long[
        ["economy", "scenario", "year", "fuel_bucket", "share_percent"]
    ].copy()

    return share_tables, share_table_long


def compute_global_radial_max(
    share_tables: dict[tuple[str, str], pd.DataFrame],
    minimum_max: float = 70.0,
) -> float:
    """
    Compute a single radial max for all charts so they are comparable.

    Rounded up to the nearest 10.
    """
    max_value = minimum_max

    for share_df in share_tables.values():
        combo_max = float(share_df.to_numpy().max())
        if combo_max > max_value:
            max_value = combo_max

    radial_max = float(np.ceil(max_value / 10.0) * 10.0)
    return radial_max


def plot_polygon_radar(
    share_df: pd.DataFrame,
    selected_years: list[int],
    radial_max: float,
    title: str,
) -> tuple[plt.Figure, plt.Axes]:
    """
    Plot a polygon-style radar chart similar to the example image.
    """
    labels = CATEGORY_ORDER
    n = len(labels)

    angles_deg = np.linspace(90, 90 - 360, n, endpoint=False)
    angles_rad = np.deg2rad(angles_deg)
    closed_angles = np.append(angles_rad, angles_rad[0])

    fig, ax = plt.subplots(figsize=(7.6, 5.4))
    ax.set_aspect("equal")
    ax.axis("off")

    grid_levels = np.arange(10, radial_max + 0.1, 10)

    for level in grid_levels:
        radius = np.repeat(level / radial_max, n)
        xg, yg = polar_to_cartesian(close_polygon(radius), closed_angles)
        ax.plot(xg, yg, color="#cfcfcf", linewidth=1)

    for angle in angles_rad:
        xs, ys = polar_to_cartesian(np.array([0.0, 1.0]), np.array([angle, angle]))
        ax.plot(xs, ys, color="#d9d9d9", linewidth=1)

    for label, angle in zip(labels, angles_rad):
        x, y = polar_to_cartesian(np.array([1.12]), np.array([angle]))
        ax.text(
            x[0],
            y[0],
            label,
            ha="center",
            va="center",
            fontsize=12,
            fontfamily="serif",
        )

    tick_angle = np.deg2rad(126)
    for level in grid_levels:
        x, y = polar_to_cartesian(np.array([level / radial_max]), np.array([tick_angle]))
        ax.text(
            x[0],
            y[0],
            f"{int(level)}%",
            ha="center",
            va="center",
            fontsize=10,
            color="#4d4d4d",
            fontfamily="serif",
        )

    available_years = [str(year) for year in selected_years if str(year) in share_df.columns]

    for i, year in enumerate(available_years):
        values = share_df[year].reindex(labels).to_numpy(dtype=float)
        radius = values / radial_max
        x, y = polar_to_cartesian(close_polygon(radius), closed_angles)

        ax.plot(
            x,
            y,
            linewidth=2.3,
            color=DEFAULT_LINE_COLORS[i % len(DEFAULT_LINE_COLORS)],
            label=year,
        )

    ax.legend(
        loc="center left",
        bbox_to_anchor=(1.02, 0.5),
        frameon=False,
        fontsize=10,
    )

    ax.set_title(
        title,
        fontsize=16,
        fontweight="bold",
        color="#a04a3c",
        fontfamily="serif",
        pad=18,
    )

    return fig, ax


def save_all_radar_charts(
    share_tables: dict[tuple[str, str], pd.DataFrame],
    selected_years: list[int],
    output_dir: str,
    radial_max: float,
) -> None:
    """Save one radar chart per economy-scenario pair."""
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    saved_count = 0

    for (economy, scenario), share_df in share_tables.items():
        title = f"{economy} - {scenario}"
        fig, ax = plot_polygon_radar(
            share_df=share_df,
            selected_years=selected_years,
            radial_max=radial_max,
            title=title,
        )

        file_name = f"{sanitize_filename(economy)}__{sanitize_filename(scenario)}__final_energy_structure_radar.png"
        save_path = output_path / file_name

        fig.savefig(save_path, dpi=300, bbox_inches="tight")
        plt.close(fig)

        saved_count += 1

    print(f"Saved {saved_count} radar charts to: {output_path}")


def save_share_table(share_table_long: pd.DataFrame, output_csv_path: str) -> None:
    """Save the long-format share table used for plotting."""
    out_path = Path(output_csv_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    share_table_long.to_csv(out_path, index=False)
    print(f"Saved share table to: {out_path}")


#%%
INPUT_FILE = r"C:\Users\Work\github\leap_utilities\data\merged_file_energy_ALL_20251106 - for chatgpt.csv"

MAJOR_YEARS = [2023, 2035, 2050, 2060]

OUTPUT_DIR = r"C:\Users\Work\github\leap_utilities\outputs\final_energy_structure_radar_all"
OUTPUT_SHARE_TABLE = r"C:\Users\Work\github\leap_utilities\outputs\final_energy_structure_radar_all\final_energy_structure_shares_long.csv"

MINIMUM_RADIAL_MAX = 70.0
OVERRIDE_RADIAL_MAX = 100.0
USE_OVERRIDE_RADIAL_MAX = True

PRINT_SAMPLE_TABLE = True
SAMPLE_ECONOMY = "01_AUS"
SAMPLE_SCENARIO = "reference"


#%%#%%
try:
    raw_df = load_energy_file(INPUT_FILE)

    raw_df = add_apec_aggregate_economy(
        raw_df,
        aggregate_economy_code="00_apec",
    )

    available_years = get_year_columns(raw_df)
    print(f"Available years: {available_years[:5]} ... {available_years[-5:]}")
    print(f"Requested major years: {MAJOR_YEARS}")

    prepared_df = prepare_final_energy_rows(raw_df, MAJOR_YEARS)

    share_tables, share_table_long = build_all_share_tables(
        prepared_df,
        selected_years=MAJOR_YEARS,
    )

    computed_radial_max = compute_global_radial_max(
        share_tables=share_tables,
        minimum_max=MINIMUM_RADIAL_MAX,
    )

    if USE_OVERRIDE_RADIAL_MAX:
        radial_max_to_use = OVERRIDE_RADIAL_MAX
    else:
        radial_max_to_use = computed_radial_max

    print(f"Computed radial max: {computed_radial_max}")
    print(f"Using radial max: {radial_max_to_use}")

    if PRINT_SAMPLE_TABLE and (SAMPLE_ECONOMY, SAMPLE_SCENARIO) in share_tables:
        print()
        print(f"Sample share table: {SAMPLE_ECONOMY} - {SAMPLE_SCENARIO}")
        print(share_tables[(SAMPLE_ECONOMY, SAMPLE_SCENARIO)].round(2))

    save_all_radar_charts(
        share_tables=share_tables,
        selected_years=MAJOR_YEARS,
        output_dir=OUTPUT_DIR,
        radial_max=radial_max_to_use,
    )

    save_share_table(
        share_table_long=share_table_long,
        output_csv_path=OUTPUT_SHARE_TABLE,
    )

except Exception as exc:
    print(f"Something went wrong: {exc}")
    raise

#%%

#%%