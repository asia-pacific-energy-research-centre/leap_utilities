#%%
"""
Build aggregated demand by LEAP fuel from ESTO base year and ninth projection data.

Combines demand from all relevant sectors into a single branch per fuel:
  Demand\\All demand aggregated\\{fuel_name}

Base year (2022): ESTO sectors including own-use, T&D losses, and main demand.
Projection years (2023+): ninth dataset filtered to subtotal_results=False and
specific sector/sub1/sub2 hierarchy.

All values are converted to positive (abs). Electricity is excluded from
Transmission & Distribution losses rows.

Standalone use:
    python -m codebase.aggregated_demand_workflow

Integration: import build_aggregated_demand() or build_aggregated_demand_as_dummy()
and pass results to results_supply_link_workflow.py when USE_AGGREGATED_DEMAND_AS_DUMMY
is enabled.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.configuration import workflow_config as workflow_cfg
from codebase.utilities.output_paths import STANDALONE_LEAP_EXPORTS_ROOT

# ── Data sources ──────────────────────────────────────────────────────────────
DATA_DIR = REPO_ROOT / "data"
CONFIG_DIR = REPO_ROOT / "config"
PROJECTION_DATA_PATH = DATA_DIR / "merged_file_energy_ALL_20250814_pre_trump.csv"
FUEL_MAPPINGS_PATH = CONFIG_DIR / "leap_mappings.xlsx"
FUEL_NINTH_SHEET = "fuel_ninth_final_proposed"

# ── Year settings ─────────────────────────────────────────────────────────────
BASE_YEAR = 2022
PROJECTION_START_YEAR = 2023
PROJECTION_END_YEAR = 2060

# ── LEAP branch / export settings ─────────────────────────────────────────────
DEMAND_BRANCH_ROOT = r"Demand\All demand aggregated"
VARIABLE_NAME = "Total Energy"
UNITS = "Petajoule"

# ── Intensity / activity mode ─────────────────────────────────────────────────
# When True, branches are written as Activity Level (=1) + Final Energy Intensity
# instead of a single Total Energy row.  LEAP computes total energy = intensity × activity,
# so with activity=1 the intensity value equals total energy.
USE_INTENSITY_ACTIVITY_MODE = True
INTENSITY_VARIABLE_NAME = "Final Energy Intensity"
ACTIVITY_VARIABLE_NAME = "Activity Level"
ACTIVITY_UNITS = "Unspecified Unit"
LEAP_SCENARIOS = ["Current Accounts", "Reference", "Target"]
# Maps LEAP scenario names to the 'scenarios' column values in the merged CSV
SCENARIO_CSV_MAP: dict[str, str] = {
    "Current Accounts": "reference",
    "Reference": "reference",
    "Target": "target",
}
DEFAULT_EXPORT_FILENAME_TEMPLATE = "aggregated_demand_{economy}_{scenario}.xlsx"
DEFAULT_EXPORT_REGION = getattr(workflow_cfg, "GLOBAL_REGION", "United States")

# Economy codes that mean "aggregate all member economies rather than filtering"
_AGGREGATE_ECONOMY_SENTINELS: frozenset[str] = frozenset({
    "00_apec", "00apec", "all_economies", "all",
})
# ── ESTO base-year demand sector filters ──────────────────────────────────────
# Own-use sub2sectors to include from 10_01_own_use
ESTO_OWN_USE_SUB2: frozenset[str] = frozenset({
    "10_01_01_electricity_chp_and_heat_plants",
    "10_01_03_liquefaction_regasification_plants",
    "10_01_06_coal_mines",
    "10_01_11_oil_refineries",
    "10_01_12_oil_and_gas_extraction",
    "10_01_13_pump_storage_plants",
})
# Demand sectors with no sub-sector restriction for base year
ESTO_OTHER_DEMAND_SECTORS: frozenset[str] = frozenset({
    "04_international_marine_bunkers",
    "05_international_aviation_bunkers",
    "14_industry_sector",
    "15_transport_sector",
    "16_other_sector",
    "17_nonenergy_use",
})

# ── Ninth projection demand sector filters ────────────────────────────────────
# All three levels must be satisfied simultaneously (not one at a time)
NINTH_SECTORS: frozenset[str] = frozenset({
    "10_losses_and_own_use",
    "14_industry_sector",
    "16_other_sector",
    "17_nonenergy_use",
    "15_transport_sector",
    "04_international_marine_bunkers",
    "05_international_aviation_bunkers",
})
NINTH_SUB1_SECTORS: frozenset[str] = frozenset({
    "x",
    "10_01_own_use",
    "10_02_transmission_and_distribution_losses",
    "14_01_mining_and_quarrying",
    "14_02_construction",
    "14_03_manufacturing",
    "16_01_buildings",
    "15_03_rail",
    "15_04_domestic_navigation",
    "16_02_agriculture_and_fishing",
    "16_05_nonspecified_others",
    "15_01_domestic_air_transport",
    "15_02_road",
    "15_05_pipeline_transport",
    "15_06_nonspecified_transport",
})
NINTH_SUB2_SECTORS: frozenset[str] = frozenset({
    "x",
    "10_01_01_electricity_chp_and_heat_plants",
    "10_01_03_liquefaction_regasification_plants",
    "10_01_11_oil_refineries",
    "10_01_12_oil_and_gas_extraction",
    "10_01_13_pump_storage_plants",
    "14_03_01_iron_and_steel",
    "14_03_02_chemical_incl_petrochemical",
    "14_03_03_non_ferrous_metals",
    "14_03_04_nonmetallic_mineral_products",
    "14_03_05_transportation_equipment",
    "14_03_06_machinery",
    "14_03_07_food_beverages_and_tobacco",
    "14_03_08_pulp_paper_and_printing",
    "14_03_09_wood_and_wood_products",
    "14_03_10_textiles_and_leather",
    "14_03_11_nonspecified_industry",
    "16_01_01_commercial_and_public_services",
    "16_01_02_residential",
    "16_01_03_ai_training",
    "16_01_04_traditional_data_centres",
    "16_02_03_agriculture",
    "16_02_04_fishing",
    "15_01_01_passenger",
    "15_01_02_freight",
    "15_02_01_passenger",
    "15_02_02_freight",
    "15_03_01_passenger",
    "15_03_02_freight",
    "15_04_01_passenger",
    "15_04_02_freight",
})

# Fuel to exclude from T&D losses (10_02)
TD_LOSSES_SUB1 = "10_02_transmission_and_distribution_losses"
TD_LOSSES_EXCLUDE_FUEL = "17_electricity"

# ── Demand zeroing mode ───────────────────────────────────────────────────────
# Variables that must NOT be zeroed because they are share/ratio variables that
# must remain coherent across sibling branches (LEAP enforces that shares sum to
# 100 and will error if they don't).
DEMAND_SHARE_VARIABLES: frozenset[str] = frozenset({
    "Device Share",
    "Sales Share",
    "Stock Share",
})

# Branch path prefix for the aggregated demand branches written by this workflow.
# These are excluded from zeroing so the aggregated demand values are preserved.
DEMAND_AGGREGATED_BRANCH_PREFIX = "Demand\\All demand aggregated"

# Branch path prefix for the Other loss and own use proxy branches.
# When the proxy workflow is running in the same pass these should neither be
# zeroed (they're being set by the proxy) nor included in the aggregated demand
# total (to avoid double-counting with the proxy output).
DEMAND_OTHER_LOSS_OWN_USE_BRANCH_PREFIX = "Demand\\Other loss and own use"

# ESTO sectors whose flows go to Demand\Other loss and own use in LEAP.
# When exclude_own_use_td_losses=True these are dropped from the aggregated sum.
OWN_USE_SECTORS: frozenset[str] = frozenset({"10_01_own_use"})
TD_LOSSES_SECTORS: frozenset[str] = frozenset({"10_02_transmission_and_distribution_losses"})

# Default source for zeroing: the full model export in data/
FULL_MODEL_EXPORT_PATH = DATA_DIR / "full model export.xlsx"
FULL_MODEL_EXPORT_SHEET = "Export"


# ── Helpers ───────────────────────────────────────────────────────────────────

def load_fuel_mapping(
    path: Path = FUEL_MAPPINGS_PATH,
    sheet: str = FUEL_NINTH_SHEET,
) -> dict[str, str]:
    """Return {ninth_fuel_code -> leap_fuel_name} from leap_mappings.xlsx."""
    df = pd.read_excel(path, sheet_name=sheet)
    df["ninth_fuel"] = df["ninth_fuel"].astype(str).str.strip()
    df["leap_fuel_name"] = df["leap_fuel_name"].astype(str).str.strip()
    df = df.drop_duplicates(subset=["ninth_fuel"], keep="first")
    return dict(zip(df["ninth_fuel"], df["leap_fuel_name"]))


def _resolve_fuel_code(fuels: pd.Series, subfuels: pd.Series) -> pd.Series:
    """Use the deepest non-'x' fuel code: subfuel if set, otherwise parent fuel."""
    sub = subfuels.astype(str).str.strip()
    parent = fuels.astype(str).str.strip()
    use_sub = sub.str.lower().ne("x") & sub.ne("") & sub.ne("nan")
    return sub.where(use_sub, parent)




def _is_aggregate_economy(economy: str | None) -> bool:
    return not economy or str(economy).strip().lower() in _AGGREGATE_ECONOMY_SENTINELS


def _load_demand_csv(
    path: Path = PROJECTION_DATA_PATH,
    economy: str | None = None,
    final_year: int = PROJECTION_END_YEAR,
) -> pd.DataFrame:
    """Load the merged energy CSV, keeping only columns needed for demand extraction.

    When economy is an aggregate sentinel (00_APEC, ALL_ECONOMIES, etc.) all
    member economies are loaded and summed later by the caller.
    """
    stable_cols = [
        "economy", "scenarios", "sectors", "sub1sectors", "sub2sectors",
        "fuels", "subfuels", "subtotal_results",
    ]
    year_cols = [str(y) for y in range(BASE_YEAR, final_year + 1)]
    header = pd.read_csv(path, nrows=0)
    use_cols = [c for c in [*stable_cols, *year_cols] if c in header.columns]
    df = pd.read_csv(path, usecols=use_cols, low_memory=False)
    for col in ["economy", "scenarios", "sectors", "sub1sectors", "sub2sectors", "fuels", "subfuels"]:
        df[col] = df[col].astype(str).str.strip()
    df["subtotal_results"] = (
        df["subtotal_results"].astype(str).str.strip().str.lower().isin({"true", "1", "yes"})
    )
    if economy and not _is_aggregate_economy(economy):
        economy_key = str(economy).strip()
        df = df[df["economy"] == economy_key].copy()
        if df.empty:
            raise ValueError(
                f"Economy {economy_key!r} not found in {path.name}. "
                f"Use an aggregate sentinel (e.g. '00_APEC') to sum all economies, "
                f"or check the economy code."
            )
    return df


# ── Core extraction ───────────────────────────────────────────────────────────

def _extract_base_year(
    df: pd.DataFrame,
    exclude_own_use_td_losses: bool = False,
) -> pd.DataFrame:
    """
    Filter to ESTO base-year (2022) demand rows. Returns long DataFrame:
    columns: economy, fuel_code, year, value.

    Sectors included:
      - 10_01 own use: sub2sectors in ESTO_OWN_USE_SUB2
      - 10_02 T&D losses: all fuels except electricity
      - 04, 05, 14, 15, 16, 17: all non-subtotal leaf rows

    When exclude_own_use_td_losses=True, the own-use and T&D losses rows are
    omitted so they are not double-counted with the other_loss_own_use proxy.
    """
    not_subtotal = ~df["subtotal_results"]

    own_use_mask = (
        (df["sectors"] == "10_losses_and_own_use") &
        (df["sub1sectors"] == "10_01_own_use") &
        (df["sub2sectors"].isin(ESTO_OWN_USE_SUB2))
    )
    td_mask = (
        (df["sectors"] == "10_losses_and_own_use") &
        (df["sub1sectors"] == TD_LOSSES_SUB1)
    )
    other_mask = df["sectors"].isin(ESTO_OTHER_DEMAND_SECTORS)

    if exclude_own_use_td_losses:
        combined_mask = not_subtotal & other_mask
    else:
        combined_mask = not_subtotal & (own_use_mask | td_mask | other_mask)

    filtered = df[combined_mask].copy()

    # Remove electricity from T&D losses (only relevant when not excluding)
    if not exclude_own_use_td_losses:
        td_elec = (filtered["sub1sectors"] == TD_LOSSES_SUB1) & (filtered["fuels"] == TD_LOSSES_EXCLUDE_FUEL)
        filtered = filtered[~td_elec].copy()

    # Use a single scenario for base year (historical values are scenario-invariant)
    filtered = filtered[filtered["scenarios"] == "reference"].copy()

    base_col = str(BASE_YEAR)
    if base_col not in filtered.columns:
        raise KeyError(f"Base year column '{BASE_YEAR}' not found in data.")

    filtered["fuel_code"] = _resolve_fuel_code(filtered["fuels"], filtered["subfuels"])
    filtered["year"] = BASE_YEAR
    filtered["value"] = pd.to_numeric(filtered[base_col], errors="coerce").abs().fillna(0.0)

    return filtered[["economy", "fuel_code", "year", "value"]].copy()


def _extract_projection_years(
    df: pd.DataFrame,
    csv_scenario: str,
    final_year: int = PROJECTION_END_YEAR,
    exclude_own_use_td_losses: bool = False,
) -> pd.DataFrame:
    """
    Filter to ninth projection rows (>=2023, subtotal_results=False).
    Applies sector + sub1sector + sub2sector filter simultaneously.
    Returns long DataFrame: economy, fuel_code, year, value.

    When exclude_own_use_td_losses=True, rows with sub1sector in OWN_USE_SECTORS
    or TD_LOSSES_SECTORS are omitted.
    """
    mask = (
        ~df["subtotal_results"] &
        df["sectors"].isin(NINTH_SECTORS) &
        df["sub1sectors"].isin(NINTH_SUB1_SECTORS) &
        df["sub2sectors"].isin(NINTH_SUB2_SECTORS) &
        (df["scenarios"] == csv_scenario)
    )
    if exclude_own_use_td_losses:
        mask = mask & ~df["sub1sectors"].isin(OWN_USE_SECTORS | TD_LOSSES_SECTORS)
    filtered = df[mask].copy()

    # Remove electricity from T&D losses (only relevant when not excluding)
    if not exclude_own_use_td_losses:
        td_elec = (filtered["sub1sectors"] == TD_LOSSES_SUB1) & (filtered["fuels"] == TD_LOSSES_EXCLUDE_FUEL)
        filtered = filtered[~td_elec].copy()

    year_cols = [
        str(y) for y in range(PROJECTION_START_YEAR, final_year + 1)
        if str(y) in filtered.columns
    ]
    if not year_cols:
        return pd.DataFrame(columns=["economy", "fuel_code", "year", "value"])

    filtered["fuel_code"] = _resolve_fuel_code(filtered["fuels"], filtered["subfuels"])
    long = filtered[["economy", "fuel_code", *year_cols]].melt(
        id_vars=["economy", "fuel_code"],
        value_vars=year_cols,
        var_name="year",
        value_name="value",
    )
    long["year"] = pd.to_numeric(long["year"], errors="coerce").astype("Int64")
    long["value"] = pd.to_numeric(long["value"], errors="coerce").abs().fillna(0.0)
    return long[["economy", "fuel_code", "year", "value"]].copy()


# ── Public API ─────────────────────────────────────────────────────────────────

def build_aggregated_demand(
    economy: str,
    scenario: str = "Reference",
    base_year: int = BASE_YEAR,
    final_year: int = PROJECTION_END_YEAR,
    data_path: Path = PROJECTION_DATA_PATH,
    fuel_mappings_path: Path = FUEL_MAPPINGS_PATH,
    exclude_own_use_td_losses: bool = False,
) -> pd.DataFrame:
    """
    Build aggregated demand by LEAP fuel for one economy and scenario.

    Returns DataFrame with columns:
        economy, scenario, leap_fuel_name, year, value  (value in PJ, positive)

    Fuel codes not found in fuel_ninth_final_proposed are dropped with a warning.

    When exclude_own_use_td_losses=True, own-use (10_01) and T&D losses (10_02)
    sectors are excluded from the sum — use this when the other_loss_own_use proxy
    is running in the same pass to avoid double-counting those amounts.
    """
    fuel_map = load_fuel_mapping(fuel_mappings_path, FUEL_NINTH_SHEET)
    df = _load_demand_csv(data_path, economy=economy, final_year=final_year)

    # For aggregate sentinels, collapse all member economies into one label
    economy_label = str(economy).strip()
    if _is_aggregate_economy(economy):
        df = df.copy()
        df["economy"] = economy_label

    # Base year (use 'reference' CSV scenario regardless of requested scenario)
    base_rows = _extract_base_year(df, exclude_own_use_td_losses=exclude_own_use_td_losses)
    base_agg = (
        base_rows
        .groupby(["economy", "fuel_code", "year"], as_index=False)["value"]
        .sum(min_count=1)
    )
    base_agg["value"] = base_agg["value"].fillna(0.0)

    if scenario == "Current Accounts":
        combined = base_agg.copy()
    else:
        csv_scen = SCENARIO_CSV_MAP.get(scenario, "reference").lower()
        proj_rows = _extract_projection_years(
            df,
            csv_scenario=csv_scen,
            final_year=final_year,
            exclude_own_use_td_losses=exclude_own_use_td_losses,
        )
        proj_agg = (
            proj_rows
            .groupby(["economy", "fuel_code", "year"], as_index=False)["value"]
            .sum(min_count=1)
        )
        proj_agg["value"] = proj_agg["value"].fillna(0.0)
        combined = pd.concat([base_agg, proj_agg], ignore_index=True)

    combined["year"] = combined["year"].astype(int)
    combined = combined[(combined["year"] >= base_year) & (combined["year"] <= final_year)].copy()

    # Map fuel codes → LEAP fuel names
    combined["leap_fuel_name"] = combined["fuel_code"].map(fuel_map)
    unmapped = combined.loc[combined["leap_fuel_name"].isna(), "fuel_code"].unique()
    if len(unmapped):
        print(
            f"[WARN] {len(unmapped)} fuel codes have no mapping in {FUEL_NINTH_SHEET},"
            f" dropped: {sorted(unmapped)[:15]}"
        )
    combined = combined[combined["leap_fuel_name"].notna()].copy()

    # Aggregate many-to-one fuel mappings
    result = combined.groupby(
        ["economy", "leap_fuel_name", "year"], as_index=False
    )["value"].sum(min_count=1)
    result["value"] = result["value"].fillna(0.0)
    result["scenario"] = scenario

    return (
        result[["economy", "scenario", "leap_fuel_name", "year", "value"]]
        .sort_values(["economy", "scenario", "leap_fuel_name", "year"])
        .reset_index(drop=True)
    )


def build_aggregated_demand_all_scenarios(
    economy: str,
    scenarios: list[str] = LEAP_SCENARIOS,
    base_year: int = BASE_YEAR,
    final_year: int = PROJECTION_END_YEAR,
    data_path: Path = PROJECTION_DATA_PATH,
    fuel_mappings_path: Path = FUEL_MAPPINGS_PATH,
    exclude_own_use_td_losses: bool = False,
) -> pd.DataFrame:
    """Build aggregated demand for all LEAP scenarios and return combined DataFrame."""
    parts = [
        build_aggregated_demand(
            economy=economy,
            scenario=s,
            base_year=base_year,
            final_year=final_year,
            data_path=data_path,
            fuel_mappings_path=fuel_mappings_path,
            exclude_own_use_td_losses=exclude_own_use_td_losses,
        )
        for s in scenarios
    ]
    return pd.concat(parts, ignore_index=True)


def build_aggregated_demand_as_dummy(
    economy: str,
    scenarios: list[str] | None = None,
    base_year: int = BASE_YEAR,
    final_year: int = PROJECTION_END_YEAR,
    data_path: Path = PROJECTION_DATA_PATH,
    fuel_mappings_path: Path = FUEL_MAPPINGS_PATH,
) -> pd.DataFrame:
    """
    Return aggregated demand data in the format expected by load_results_demand_table
    in results_supply_link_workflow.py for use as dummy demand.

    Returns DataFrame with columns:
        economy, scenario, esto_product, year, demand_value, demand_source

    Fuel names are mapped back to esto_product codes via fuel_product_final_proposed.
    Rows where no esto_product mapping exists are dropped.
    """
    use_scenarios = scenarios if scenarios is not None else LEAP_SCENARIOS
    long = build_aggregated_demand_all_scenarios(
        economy=economy,
        scenarios=use_scenarios,
        base_year=base_year,
        final_year=final_year,
        data_path=data_path,
        fuel_mappings_path=fuel_mappings_path,
    )
    if long.empty:
        return pd.DataFrame(
            columns=["economy", "scenario", "esto_product", "year", "demand_value", "demand_source"]
        )

    # Load reverse mapping: leap_fuel_name → esto_product
    fuel_prod = pd.read_excel(fuel_mappings_path, sheet_name="fuel_product_final_proposed")
    fuel_prod["leap_fuel_name"] = fuel_prod["leap_fuel_name"].astype(str).str.strip()
    fuel_prod["esto_product"] = fuel_prod["esto_product"].astype(str).str.strip()
    # When multiple esto_products map to one leap_fuel_name, keep first (many-to-one is OK here)
    prod_map = (
        fuel_prod.drop_duplicates(subset=["leap_fuel_name"], keep="first")
        .set_index("leap_fuel_name")["esto_product"]
        .to_dict()
    )

    long["esto_product"] = long["leap_fuel_name"].map(prod_map)
    unmapped = long.loc[long["esto_product"].isna(), "leap_fuel_name"].unique()
    if len(unmapped):
        print(
            f"[WARN] {len(unmapped)} LEAP fuel names have no esto_product mapping, dropped:"
            f" {sorted(unmapped)[:10]}"
        )
    long = long[long["esto_product"].notna()].copy()

    result = long.groupby(
        ["economy", "scenario", "esto_product", "year"], as_index=False
    )["value"].sum(min_count=1)
    result["value"] = result["value"].fillna(0.0)
    result["demand_source"] = "aggregated_demand_projection"

    return (
        result.rename(columns={"value": "demand_value"})
        [["economy", "scenario", "esto_product", "year", "demand_value", "demand_source"]]
        .reset_index(drop=True)
    )


# ── Aggregated demand LEAP workbook ──────────────────────────────────────────

def _build_id_lookups(
    id_lookup_path: Path | str,
) -> tuple[dict[str, int], dict[str, int], dict[str, int]]:
    """Return (branch_to_id, variable_to_id, scenario_to_id) dicts from a LEAP full export."""
    raw = pd.read_excel(Path(id_lookup_path), header=2)
    branch_to_id = (
        raw[["Branch Path", "BranchID"]].dropna(subset=["Branch Path"])
        .drop_duplicates(subset=["Branch Path"])
        .set_index("Branch Path")["BranchID"]
        .apply(lambda x: int(x) if pd.notna(x) else -1)
        .to_dict()
    )
    variable_to_id = (
        raw[["Variable", "VariableID"]].dropna(subset=["Variable"])
        .drop_duplicates(subset=["Variable"])
        .set_index("Variable")["VariableID"]
        .apply(lambda x: int(x) if pd.notna(x) else -1)
        .to_dict()
    )
    scenario_to_id = (
        raw[["Scenario", "ScenarioID"]].dropna(subset=["Scenario"])
        .drop_duplicates(subset=["Scenario"])
        .set_index("Scenario")["ScenarioID"]
        .apply(lambda x: int(x) if pd.notna(x) else -1)
        .to_dict()
    )
    return branch_to_id, variable_to_id, scenario_to_id


def save_aggregated_demand_as_leap_workbook(
    economy: str,
    output_path: Path,
    scenarios: list[str] | None = None,
    region: str = DEFAULT_EXPORT_REGION,
    base_year: int = BASE_YEAR,
    final_year: int = PROJECTION_END_YEAR,
    data_path: Path = PROJECTION_DATA_PATH,
    fuel_mappings_path: Path = FUEL_MAPPINGS_PATH,
    model_name: str = "",
    exclude_own_use_td_losses: bool = False,
    id_lookup_path: Path | str | None = None,
) -> Path | None:
    """
    Build aggregated demand and save as a LEAP-importable workbook (LEAP + FOR_VIEWING sheets).

    Writes Demand\\All demand aggregated\\{fuel_name} rows with Variable=Total Energy
    and Expression as a scalar (Current Accounts) or Data(...) series (other scenarios).
    Returns the output path, or None if there was nothing to write.

    When exclude_own_use_td_losses=True, own-use and T&D losses sectors are excluded
    from the demand sum so the aggregated total does not double-count amounts that the
    other_loss_own_use proxy handles separately in Demand\\Other loss and own use.

    When id_lookup_path is provided, BranchID/VariableID/ScenarioID columns are merged
    from that file (a LEAP full export with header=2). RegionID is always set to 1.
    """
    use_scenarios = scenarios if scenarios is not None else list(LEAP_SCENARIOS)
    demand = build_aggregated_demand_all_scenarios(
        economy=economy,
        scenarios=use_scenarios,
        base_year=base_year,
        final_year=final_year,
        data_path=data_path,
        fuel_mappings_path=fuel_mappings_path,
        exclude_own_use_td_losses=exclude_own_use_td_losses,
    )
    if demand.empty:
        print("[INFO] save_aggregated_demand_as_leap_workbook: no demand data — workbook not written.")
        return None

    rows = []
    for (fuel_name, scenario), grp in demand.groupby(["leap_fuel_name", "scenario"], sort=True):
        grp = grp.sort_values("year")
        year_val = list(zip(grp["year"].astype(int), grp["value"].astype(float)))
        if scenario == "Current Accounts":
            base_vals = [(yr, v) for yr, v in year_val if yr == base_year]
            expr = f"{base_vals[0][1]:.6g}" if base_vals else "0"
        else:
            tokens: list[str] = []
            for yr, val in year_val:
                tokens.append(str(yr))
                tokens.append(f"{val:.6g}")
            expr = "Data(" + ", ".join(tokens) + ")"
        branch = f"{DEMAND_BRANCH_ROOT}\\{fuel_name}"
        if USE_INTENSITY_ACTIVITY_MODE:
            rows.append({
                "Branch Path": branch,
                "Variable": ACTIVITY_VARIABLE_NAME,
                "Scenario": scenario,
                "Region": region,
                "Scale": "",
                "Units": ACTIVITY_UNITS,
                "Per...": "",
                "Expression": expr,
            })
            rows.append({
                "Branch Path": branch,
                "Variable": INTENSITY_VARIABLE_NAME,
                "Scenario": scenario,
                "Region": region,
                "Scale": "",
                "Units": UNITS,
                "Per...": ACTIVITY_UNITS,
                "Expression": "1",
            })
        else:
            rows.append({
                "Branch Path": branch,
                "Variable": VARIABLE_NAME,
                "Scenario": scenario,
                "Region": region,
                "Scale": "",
                "Units": UNITS,
                "Per...": "",
                "Expression": expr,
            })

    if not rows:
        print("[INFO] save_aggregated_demand_as_leap_workbook: no rows after grouping — workbook not written.")
        return None

    export_df = pd.DataFrame(rows)

    id_lookup_resolved = Path(id_lookup_path) if id_lookup_path is not None else None
    if id_lookup_resolved is not None and id_lookup_resolved.exists():
        branch_to_id, variable_to_id, scenario_to_id = _build_id_lookups(id_lookup_resolved)
        export_df.insert(0, "BranchID", export_df["Branch Path"].map(
            lambda x: branch_to_id.get(str(x).strip(), -1)))
        export_df.insert(1, "VariableID", export_df["Variable"].map(
            lambda x: variable_to_id.get(str(x).strip(), -1)))
        export_df.insert(2, "ScenarioID", export_df["Scenario"].map(
            lambda x: scenario_to_id.get(str(x).strip(), -1)))
        export_df.insert(3, "RegionID", 1)
        matched = int((export_df["BranchID"] != -1).sum())
        print(f"[INFO] Merged IDs: {matched}/{len(export_df)} rows matched BranchID.")
    elif id_lookup_resolved is not None:
        print(f"[WARN] id_lookup_path not found, skipping ID merge: {id_lookup_resolved}")

    cols = list(export_df.columns)

    preamble_row = {col: "" for col in cols}
    preamble_row["Branch Path"] = "Area:"
    preamble_row["Variable"] = model_name or ""
    preamble_row["Scenario"] = "Ver:"
    preamble_row["Region"] = "2"
    empty_row = {col: pd.NA for col in cols}
    header_row_data = {col: col for col in cols}

    full_df = pd.concat(
        [
            pd.DataFrame([preamble_row]),
            pd.DataFrame([empty_row]),
            pd.DataFrame([header_row_data]),
            export_df,
        ],
        ignore_index=True,
    )

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        full_df.to_excel(writer, sheet_name="LEAP", index=False, header=False)
        full_df.to_excel(writer, sheet_name="FOR_VIEWING", index=False, header=False)

    print(f"[INFO] Saved {len(rows)} aggregated demand rows to {output_path}")
    return output_path


# ── Demand zeroing export ─────────────────────────────────────────────────────

def build_demand_zeroing_rows(
    source_path: Path = FULL_MODEL_EXPORT_PATH,
    sheet_name: str = FULL_MODEL_EXPORT_SHEET,
    scenarios: list[str] | None = None,
    region: str = DEFAULT_EXPORT_REGION,
    exclude_branch_prefixes: list[str] | None = None,
) -> pd.DataFrame:
    """
    Build LEAP import rows to zero out all non-share demand branches.

    Reads Demand branch rows from source_path (typically data/full model export.xlsx),
    excluding:
      - Demand\\All demand aggregated\\... branches (where aggregated demand is written)
      - Share variables listed in DEMAND_SHARE_VARIABLES (Device Share, Sales Share,
        Stock Share) which must remain coherent across siblings
      - Any branch prefixes listed in exclude_branch_prefixes (e.g.
        DEMAND_OTHER_LOSS_OWN_USE_BRANCH_PREFIX when the proxy is running in the
        same pass)

    Returns a DataFrame with LEAP import columns: Branch Path, Variable, Scenario,
    Region, Scale, Units, Per..., Expression. Expression is "0" for all rows.
    """
    _LEAP_EXPORT_COLS = [
        "Branch Path", "Variable", "Scenario", "Region",
        "Scale", "Units", "Per...", "Expression",
    ]
    empty = pd.DataFrame(columns=_LEAP_EXPORT_COLS)

    path = Path(source_path)
    if not path.exists():
        print(f"[WARN] Demand zeroing source not found: {path}")
        return empty

    try:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    except Exception as exc:
        print(f"[WARN] Failed reading {path} for demand zeroing: {exc}")
        return empty

    header_row = None
    for idx in range(len(raw.index)):
        row_vals = {
            str(v).strip().lower()
            for v in raw.iloc[idx].tolist()
            if str(v or "").strip()
        }
        if "branch path" in row_vals and "variable" in row_vals:
            header_row = idx
            break
    if header_row is None:
        print(f"[WARN] Could not find LEAP header row in {path}")
        return empty

    df = raw.iloc[header_row + 1:].copy()
    df.columns = raw.iloc[header_row].tolist()
    df = df.dropna(how="all").reset_index(drop=True)

    for col in ["Branch Path", "Variable", "Scenario", "Region"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    mask = (
        df["Branch Path"].str.startswith("Demand\\")
        & ~df["Branch Path"].str.startswith(DEMAND_AGGREGATED_BRANCH_PREFIX)
        & ~df["Variable"].isin(DEMAND_SHARE_VARIABLES)
    )
    if exclude_branch_prefixes:
        for prefix in exclude_branch_prefixes:
            mask = mask & ~df["Branch Path"].str.startswith(prefix)
    df = df[mask].copy()

    if scenarios:
        df = df[df["Scenario"].isin(scenarios)].copy()

    if df.empty:
        print("[INFO] No demand zeroing rows found after filtering.")
        return empty

    df = df.drop_duplicates(subset=["Branch Path", "Variable", "Scenario"], keep="first")

    result = df[["Branch Path", "Variable", "Scenario"]].copy()
    result["Region"] = region
    result["Scale"] = df["Scale"].fillna("") if "Scale" in df.columns else ""
    result["Units"] = df["Units"].fillna("") if "Units" in df.columns else ""
    result["Per..."] = df["Per..."].fillna("") if "Per..." in df.columns else ""
    result["Expression"] = "0"

    return result[_LEAP_EXPORT_COLS].reset_index(drop=True)


def save_demand_zeroing_workbook(
    output_path: Path,
    source_path: Path = FULL_MODEL_EXPORT_PATH,
    sheet_name: str = FULL_MODEL_EXPORT_SHEET,
    scenarios: list[str] | None = None,
    region: str = DEFAULT_EXPORT_REGION,
    model_name: str = "",
    exclude_branch_prefixes: list[str] | None = None,
) -> Path | None:
    """
    Save a LEAP-importable workbook that sets all non-share demand branches to 0.

    The workbook has LEAP and FOR_VIEWING sheets in the format expected by
    _merge_workbook_sheets and fill_branches_from_export_file. Rows cover every
    (Branch Path, Variable, Scenario) combination found in source_path under
    Demand\\..., except aggregated-demand branches, share variables, and any
    prefixes listed in exclude_branch_prefixes.
    """
    rows = build_demand_zeroing_rows(
        source_path=source_path,
        sheet_name=sheet_name,
        scenarios=scenarios,
        region=region,
        exclude_branch_prefixes=exclude_branch_prefixes,
    )
    if rows.empty:
        print("[INFO] No demand zeroing rows — workbook not written.")
        return None

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    cols = list(rows.columns)
    preamble_row = {col: "" for col in cols}
    preamble_row["Branch Path"] = "Area:"
    preamble_row["Variable"] = model_name or ""
    preamble_row["Scenario"] = "Ver:"
    preamble_row["Region"] = "2"
    empty_row = {col: pd.NA for col in cols}
    header_row_data = {col: col for col in cols}

    full_df = pd.concat(
        [
            pd.DataFrame([preamble_row]),
            pd.DataFrame([empty_row]),
            pd.DataFrame([header_row_data]),
            rows,
        ],
        ignore_index=True,
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        full_df.to_excel(writer, sheet_name="LEAP", index=False, header=False)
        full_df.to_excel(writer, sheet_name="FOR_VIEWING", index=False, header=False)

    print(f"[INFO] Saved {len(rows)} demand zeroing rows to {output_path}")
    return output_path


# ── LEAP export writer ────────────────────────────────────────────────────────

def _data_expression(year_val_pairs: list[tuple[int, float]]) -> str:
    """Build LEAP Data(year, val, ...) expression string."""
    tokens = []
    for yr, val in year_val_pairs:
        tokens.append(str(yr))
        tokens.append(f"{val:.6g}")
    return "Data(" + ", ".join(tokens) + ")"


def save_to_leap_export(
    demand_df: pd.DataFrame,
    output_path: Path,
    region: str = DEFAULT_EXPORT_REGION,
    branch_root: str = DEMAND_BRANCH_ROOT,
) -> None:
    """
    Write aggregated demand to a LEAP-importable Excel workbook.

    demand_df must have columns: economy, scenario, leap_fuel_name, year, value.
    Produces one row per (fuel, scenario) in the LEAP export format.
    """
    if demand_df is None or demand_df.empty:
        print("[WARN] save_to_leap_export called with empty DataFrame — nothing written.")
        return

    rows = []
    for (fuel_name, scenario), grp in demand_df.groupby(
        ["leap_fuel_name", "scenario"], sort=True
    ):
        grp = grp.sort_values("year")
        year_val = list(zip(grp["year"].astype(int), grp["value"].astype(float)))

        if scenario == "Current Accounts":
            base_vals = [(yr, v) for yr, v in year_val if yr == BASE_YEAR]
            expr = f"{base_vals[0][1]:.6g}" if base_vals else "0"
        else:
            expr = _data_expression(year_val)

        branch = f"{branch_root}\\{fuel_name}"
        if USE_INTENSITY_ACTIVITY_MODE:
            rows.append({
                "Branch Path": branch,
                "Variable": ACTIVITY_VARIABLE_NAME,
                "Scenario": scenario,
                "Region": region,
                "Scale": "",
                "Units": ACTIVITY_UNITS,
                "Per...": "",
                "Expression": expr,
            })
            rows.append({
                "Branch Path": branch,
                "Variable": INTENSITY_VARIABLE_NAME,
                "Scenario": scenario,
                "Region": region,
                "Scale": "",
                "Units": UNITS,
                "Per...": ACTIVITY_UNITS,
                "Expression": "1",
            })
        else:
            rows.append({
                "Branch Path": branch,
                "Variable": VARIABLE_NAME,
                "Scenario": scenario,
                "Region": region,
                "Scale": "",
                "Units": UNITS,
                "Per...": "",
                "Expression": expr,
            })

    export_df = pd.DataFrame(rows)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Write two metadata header rows, then the column-name row, then data
    col_names = list(export_df.columns)
    meta = pd.DataFrame(
        [
            ["Area:", "", "", "", "", "", "", ""],
            ["Ver:", 2, "", "", "", "", "", ""],
        ],
        columns=col_names,
    )
    header_row = pd.DataFrame([col_names], columns=col_names)
    full_output = pd.concat([meta, header_row, export_df], ignore_index=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        full_output.to_excel(writer, sheet_name="Export", index=False, header=False)

    print(f"[INFO] Saved {len(export_df)} rows to {output_path}")


# ── Standalone entry point ────────────────────────────────────────────────────

def main(
    economy: str | None = None,
    scenarios: list[str] | None = None,
    final_year: int = PROJECTION_END_YEAR,
    output_dir: Path | None = None,
) -> None:
    """
    Run the aggregated demand workflow for one economy and save to Excel.

    If economy is None, uses workflow_config.GLOBAL_ECONOMIES[0].
    If scenarios is None, uses LEAP_SCENARIOS (Current Accounts, Reference, Target).
    """
    if economy is None:
        economies = list(getattr(workflow_cfg, "GLOBAL_ECONOMIES", ["20_USA"]))
        economy = economies[0] if economies else "20_USA"

    use_scenarios = scenarios if scenarios is not None else list(LEAP_SCENARIOS)
    out_dir = Path(output_dir) if output_dir else STANDALONE_LEAP_EXPORTS_ROOT

    if _is_aggregate_economy(economy):
        print(f"[INFO] Economy {economy!r} is an aggregate sentinel — summing all member economies.")
    print(f"[INFO] Building aggregated demand for economy={economy!r}")
    print(f"[INFO] Scenarios: {use_scenarios}, years: {BASE_YEAR}–{final_year}")

    demand = build_aggregated_demand_all_scenarios(
        economy=economy,
        scenarios=use_scenarios,
        base_year=BASE_YEAR,
        final_year=final_year,
        data_path=PROJECTION_DATA_PATH,
        fuel_mappings_path=FUEL_MAPPINGS_PATH,
    )

    fuels_found = sorted(demand["leap_fuel_name"].unique())
    print(f"[INFO] {len(fuels_found)} fuels after mapping: {fuels_found}")

    scenario_token = "_".join(
        "".join(c for c in s if c.isalnum()) for s in use_scenarios
    )
    econ_token = "".join(c for c in economy if c.isalnum() or c == "_")
    filename = f"aggregated_demand_{econ_token}_{scenario_token}.xlsx"
    output_path = out_dir / filename

    save_to_leap_export(demand, output_path=output_path)
    print(f"[INFO] Done.")


if __name__ == "__main__":
    main()
#%%