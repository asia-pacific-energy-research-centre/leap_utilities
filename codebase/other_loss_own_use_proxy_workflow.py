#%%
"""
Build proxy activity and intensity inputs for other loss / own-use demand branches.

The workflow creates LEAP demand-branch first estimates for balance-table
losses and own-use flows that do not fit cleanly in transformation modules.
For each configured process it:
- builds one proxy activity series, for example coal production for coal mines;
- pulls the target loss/own-use series by fuel;
- calculates fuel intensity as abs(target energy) / proxy activity;
- writes audit CSV files and a LEAP import workbook.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path
from collections.abc import Mapping, Sequence

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if Path.cwd() != REPO_ROOT:
    os.chdir(REPO_ROOT)
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.configuration import workflow_config as workflow_cfg
from codebase.configuration.config import BASE_YEAR
from codebase.functions.analysis_input_write_dispatcher import dispatch_analysis_input_write
from codebase.functions.leap_core import (
    create_branches_from_export_file,
    fill_branches_from_export_file,
    is_leap_api_available,
    sanitize_leap_name,
)
from codebase.functions.leap_excel_io import finalise_export_df, save_export_files
from codebase.functions.leap_labels import clean_fuel_label_for_leap
from codebase.functions.leap_expressions import build_data_expression_from_row
from codebase.scrapbook.utilities import apply_matt_subtotal_mapping, filter_matt_subtotals
from codebase.utilities import fuel_catalog_preflight, workflow_common
from codebase.utilities.leap_balance_export_resolver import (
    build_leap_balance_activity_series,
    load_leap_balance_activity_table as _load_shared_leap_balance_activity_table,
    normalize_balance_label,
    resolve_balance_export_workbook,
)


# --- Stable paths ---
ESTO_DATA_PATH = Path("data/00APEC_2024_low.csv")
NINTH_DATA_PATH = Path("data/merged_file_energy_ALL_20251106.csv")
ESTO_SUBTOTAL_MAPPING_PATH = Path("config/ESTO_subtotal_mapping.xlsx")
LEAP_MAPPINGS_PATH = Path("config/leap_mappings.xlsx")
OUTPUT_FUEL_VALIDATION_ESTO_PATHS = [
    Path("data/00APEC_2025_low_with_subtotals.csv"),
    Path("data/00APEC_2024_low_with_subtotals.csv"),
]


# --- LEAP export settings ---
ACTIVITY_VARIABLE = "Activity Level"
INTENSITY_VARIABLE = "Final Energy Intensity"
EXPORT_MODEL_NAME = "Other loss and own use proxy import"

# LEAP region written into the export workbook.
EXPORT_REGION = workflow_cfg.GLOBAL_REGION

# Scenarios written to the workbook. Common values are "Target", "Reference",
# and "Current Accounts". Keep "Current Accounts" if you want base-year
# expressions for Current Accounts.
EXPORT_SCENARIOS = list(workflow_cfg.GLOBAL_SCENARIOS)

EXPORT_BASE_YEAR = BASE_YEAR
EXPORT_FINAL_YEAR = 2060 if workflow_cfg.GLOBAL_FINAL_YEAR is None else int(workflow_cfg.GLOBAL_FINAL_YEAR)

# 9th Edition scenario used for first-run projections.
NINTH_SCENARIO = "target"

# Supporting CSV output location and main LEAP import workbook filename pattern.
OUTPUT_ROOT = Path(workflow_cfg.GLOBAL_EXPORT_OUTPUT_DIR) / "supporting_files" / "other_loss_own_use_proxy"
EXPORT_FILENAME_TEMPLATE = str(
    Path(workflow_cfg.GLOBAL_EXPORT_OUTPUT_DIR) / "other_loss_own_use_proxy_{economy}_{scenario}.xlsx"
)

# Existing LEAP export workbook used to attach BranchID, VariableID,
# ScenarioID, and RegionID to generated rows. Every generated row must match
# this file on Branch Path + Variable + Scenario + Region, otherwise the
# workflow raises because that means the branch/variable/scenario is unexpected.
EXPORT_KEY_WORKBOOK_PATH = Path("data/full model export.xlsx")
EXPORT_KEY_WORKBOOK_SHEET = "Export"

# Demand branch root. Generated paths are:
# Demand\Other loss and own use\<process>\<fuel>
DEMAND_ROOT_PARTS = ["Demand", "Other loss and own use"]

# Activity source mode:
# - "esto_ninth": first-run method. Activity uses ESTO base years and 9th
#   Edition projection years.
# - "leap_balance": second-run method. Activity uses a LEAP balance export
#   workbook, such as data/leap balances exports/20_USA/...xlsx.
PROXY_ACTIVITY_SOURCE_MODE = "esto_ninth"

# LEAP balance workbook controls used when PROXY_ACTIVITY_SOURCE_MODE is
# "leap_balance".
# - Set LEAP_BALANCE_WORKBOOK_PATH to a specific .xlsx path to use exactly
#   that file.
# - Leave it as None to auto-resolve from the selected economy,
#   LEAP_BALANCE_SCENARIO, and LEAP_BALANCE_DATE_ID.
LEAP_BALANCE_EXPORTS_ROOT = Path("data/leap balances exports")

# Scenario/date filters for auto-resolving a LEAP balance workbook when the
# explicit workbook path is None. Scenario examples: "Target", "Reference".
# Date ID examples depend on export filenames; None means use the resolver's
# latest/best match.
LEAP_BALANCE_SCENARIO = "Target"
LEAP_BALANCE_DATE_ID = None
LEAP_BALANCE_WORKBOOK_PATH = None

# Output fuel scope:
# - "economy": normal input-data mode. Keep only fuels that are non-zero for
#   the selected economy in the final year of both ESTO validation snapshots.
# - "all_economies": model-structure mode. Keep fuels that are non-zero in at
#   least one economy in both ESTO validation snapshots.
OUTPUT_FUEL_VALIDATION_SCOPE = "economy"

DEFAULT_MEASURE_UNITS = {
    ACTIVITY_VARIABLE: {"units": "Unspecified Unit", "scale": "", "per": ""},
    INTENSITY_VARIABLE: {"units": "Petajoule", "scale": "", "per": ""},
}


# Explicit LEAP balance fuel sets. These are intentionally separate from mapping
# workbooks so each proxy can state exactly which LEAP fuels drive its activity.
LEAP_BALANCE_FUEL_SETS = {
    "electricity_heat_output": [
        "Electricity",
        "Heat",
    ],
    "natural_gas_lng_output": [
        "Natural gas",
        "LNG",
    ],
    "gas_works_gas_output": [
        "Gas works gas",
    ],
    "gas_to_liquids_output": [
        "LPG",
        "Kerosene",
        "Naphtha",
        "Gas and diesel oil",
        "Fuel oil",
        "Refinery gas not liquefied",
        "Other products",
    ],
    "coal_primary_and_products": [
        "Coking coal",
        "Other bituminous coal",
        "Sub bituminous coal",
        "Anthracite",
        "Lignite",
        "Coal nonspecified",
        "Coke oven coke",
        "Gas coke",
        "Coke oven gas",
        "Blast furnace gas",
        "Other recovered gases",
        "Patent fuel",
        "Coal tar",
        "BKB and PB",
    ],
    "coke_oven_output": [
        "Coke oven coke",
        "Coke oven gas",
        "Coal tar",
    ],
    "blast_furnace_output": [
        "Blast furnace gas",
    ],
    "patent_fuel_output": [
        "Patent fuel",
    ],
    "bkb_pb_output": [
        "BKB and PB",
    ],
    "coal_to_oil_output": [
        "LPG",
        "Naphtha",
        "Gas and diesel oil",
        "Fuel oil",
        "Other products",
    ],
    "oil_refinery_output": [
        "LPG",
        "Naphtha",
        "Gas and diesel oil",
        "Fuel oil",
        "Kerosene",
        "Kerosene type jet fuel",
        "Gasoline type jet fuel",
        "Motor gasoline",
        "Aviation gasoline",
        "Bitumen",
        "Petroleum coke",
        "Lubricants",
        "Paraffin waxes",
        "White spirit SBP",
        "Refinery gas not liquefied",
        "Ethane",
        "Other products",
        "PetProd nonspecified",
    ],
    "oil_and_gas_primary_production": [
        "Crude oil",
        "Natural gas",
        "Natural gas liquids",
        "Other hydrocarbons",
    ],
    "electricity_output": [
        "Electricity",
    ],
    "nuclear_primary_production": [
        "Nuclear",
    ],
    "charcoal_output": [
        "Charcoal",
    ],
    "biogas_output": [
        "Biogas",
    ],
    "nonspecified_transformation_output": [
        "Natural gas liquids",
        "Other hydrocarbons",
        "LPG",
        "Ethane",
    ],
    "all_production_ex_total": [
        "Electricity",
        "Natural gas",
        "Kerosene",
        "LPG",
        "Crude oil",
        "Charcoal",
        "Hydro",
        "Biogasoline",
        "Wind",
        "Geothermal",
        "Nuclear",
        "Heat",
        "Hydrogen",
        "Biogas",
        "Naphtha",
        "Peat",
        "Bitumen",
        "Petroleum coke",
        "Lubricants",
        "Refinery feedstocks",
        "Biomass",
        "Bagasse",
        "LNG",
        "Ammonia",
        "Fuel oil",
        "Aviation gasoline",
        "Biodiesel",
        "Patent fuel",
        "Other bituminous coal",
        "Coking coal",
        "Coke oven gas",
        "Coal tar",
        "Blast furnace gas",
        "Coke oven coke",
        "Lignite",
        "Anthracite",
        "Additives and oxygenates",
        "BKB and PB",
        "Bio jet kerosene",
        "Black liqour",
        "Coal nonspecified",
        "Ethane",
        "Fuelwood and woodwaste",
        "Gas coke",
        "Gas and diesel oil",
        "Gasoline type jet fuel",
        "Industrial waste",
        "Kerosene type jet fuel",
        "Motor gasoline",
        "Municipal solid waste renewable",
        "Natural gas liquids",
        "Other biomass",
        "Other hydrocarbons",
        "Other liquid biofuels",
        "Other products",
        "Other recovered gases",
        "Other sources",
        "Paraffin waxes",
        "PetProd nonspecified",
        "Refinery gas not liquefied",
        "Solar nonspecified",
        "Tide wave ocean",
        "White spirit SBP",
        "of which Photovoltaics",
        "Municipal solid waste non renewable",
        "Sub bituminous coal",
    ],
    "production_ex_electricity": [
        "Natural gas",
        "Kerosene",
        "LPG",
        "Crude oil",
        "Charcoal",
        "Hydro",
        "Biogasoline",
        "Wind",
        "Geothermal",
        "Nuclear",
        "Heat",
        "Hydrogen",
        "Biogas",
        "Naphtha",
        "Peat",
        "Bitumen",
        "Petroleum coke",
        "Lubricants",
        "Refinery feedstocks",
        "Biomass",
        "Bagasse",
        "LNG",
        "Ammonia",
        "Fuel oil",
        "Aviation gasoline",
        "Biodiesel",
        "Patent fuel",
        "Other bituminous coal",
        "Coking coal",
        "Coke oven gas",
        "Coal tar",
        "Blast furnace gas",
        "Coke oven coke",
        "Lignite",
        "Anthracite",
        "Additives and oxygenates",
        "BKB and PB",
        "Bio jet kerosene",
        "Black liqour",
        "Coal nonspecified",
        "Ethane",
        "Fuelwood and woodwaste",
        "Gas coke",
        "Gas and diesel oil",
        "Gasoline type jet fuel",
        "Industrial waste",
        "Kerosene type jet fuel",
        "Motor gasoline",
        "Municipal solid waste renewable",
        "Natural gas liquids",
        "Other biomass",
        "Other hydrocarbons",
        "Other liquid biofuels",
        "Other products",
        "Other recovered gases",
        "Other sources",
        "Paraffin waxes",
        "PetProd nonspecified",
        "Refinery gas not liquefied",
        "Solar nonspecified",
        "Tide wave ocean",
        "White spirit SBP",
        "of which Photovoltaics",
        "Municipal solid waste non renewable",
        "Sub bituminous coal",
    ],
}

# Fallback activity definitions for lightweight LEAP balance exports.
# Some reduced-detail exports do not include process-specific rows such as
# "Pumped hydro". For those cases we provide a deterministic fallback so
# second-stage proxy activity remains usable.
LEAP_BALANCE_ACTIVITY_FALLBACKS = {
    "pump_storage_plants": {
        "balance_rows": ["Electricity Generation"],
        "fuel_set": "electricity_output",
        "value_mode": "positive_only",
    },
}


def make_proxy_config(
    *,
    process_key: str,
    process_label: str,
    esto_target_flows: Sequence[str],
    ninth_target_sectors: Sequence[str],
    activity_label: str = "TODO: define proxy activity",
    enabled: bool = False,
    leap_process_label: str | None = None,
    esto_activity_flows: Sequence[str] | None = None,
    esto_activity_product_prefixes: Sequence[str] | None = None,
    esto_activity_exact_products: Sequence[str] | None = None,
    esto_activity_exclude_products: Sequence[str] | None = None,
    esto_target_exclude_products: Sequence[str] | None = None,
    ninth_activity_sectors: Sequence[str] | None = None,
    ninth_activity_fuels: Sequence[str] | None = None,
    ninth_activity_subfuels: Sequence[str] | None = None,
    ninth_activity_exclude_fuels: Sequence[str] | None = None,
    ninth_activity_exclude_subfuels: Sequence[str] | None = None,
    ninth_target_exclude_fuels: Sequence[str] | None = None,
    ninth_target_exclude_subfuels: Sequence[str] | None = None,
    leap_balance_rows: Sequence[str] | None = None,
    leap_balance_fuel_set: str = "",
    activity_value_mode: str = "signed_sum",
    notes: str = "",
) -> dict[str, object]:
    """Return one editable proxy config entry."""
    return {
        "enabled": bool(enabled),
        "process_key": process_key,
        "process_label": process_label,
        "leap_process_label": leap_process_label or process_label,
        "activity_label": activity_label,
        "activity_sources": {
            "esto": {
                "flows": list(esto_activity_flows or []),
                "product_prefixes": list(esto_activity_product_prefixes or []),
                "include_exact_products": list(esto_activity_exact_products or []),
                "exclude_products": list(esto_activity_exclude_products or []),
                "value_mode": activity_value_mode,
            },
            "ninth": {
                "sector_codes": list(ninth_activity_sectors or []),
                "fuels": list(ninth_activity_fuels or []),
                "subfuels": list(ninth_activity_subfuels or []),
                "exclude_fuels": list(ninth_activity_exclude_fuels or []),
                "exclude_subfuels": list(ninth_activity_exclude_subfuels or []),
                "value_mode": activity_value_mode,
            },
            "leap_balance": {
                "balance_rows": list(leap_balance_rows or []),
                "fuel_set": leap_balance_fuel_set,
                "value_mode": activity_value_mode,
            },
        },
        "target_sources": {
            "esto": {
                "flows": list(esto_target_flows),
                "exclude_products": list(esto_target_exclude_products or []),
            },
            "ninth": {
                "sector_codes": list(ninth_target_sectors),
                "exclude_fuels": list(ninth_target_exclude_fuels or []),
                "exclude_subfuels": list(ninth_target_exclude_subfuels or []),
            },
        },
        "notes": notes,
    }


# Extracted own-use/loss children from the 9th and ESTO source tables. Keep
# disabled scaffold entries in this list so proxies can be filled in one by one
# without changing the builder functions.
PROXY_CONFIG = [
    make_proxy_config(
        enabled=True,#coal mines own-use/losses cannot be easily handled by auxiliary branch in leap since they are a supply related flow rather than a transformation flow, so keeping enabled for now
        process_key="coal_mines",
        process_label="Coal mines",
        activity_label="Primary coal and coal-products production",
        esto_activity_flows=["01 Production"],
        esto_activity_product_prefixes=["01.", "02."],
        ninth_activity_sectors=["01_production"],
        ninth_activity_fuels=["01_coal", "02_coal_products"],
        ninth_activity_subfuels=[
            "01_01_coking_coal",
            "01_05_lignite",
            "01_coal_unallocated",
            "01_x_thermal_coal",
            "02_coal_products",
        ],
        leap_balance_rows=["Production"],
        leap_balance_fuel_set="coal_primary_and_products",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.06 Coal mines"],
        ninth_target_sectors=["10_01_06_coal_mines"],
        notes=(
            "Proxy activity is the sum of primary production of subfuels under "
            "01 Coal and 02 Coal products. Fuel intensity is abs(10.01.06 Coal "
            "mines own-use/loss fuel) divided by that activity."
        ),
    ),
    make_proxy_config(
        enabled=True,#cannot be easily handled by auxiliary branch in leap since it includes both electricity and heat output, so keeping enabled for now
        process_key="electricity_chp_and_heat_plants",
        process_label="Electricity, CHP and heat plants",
        activity_label="Electricity and heat output from electricity, CHP, and heat plants",
        esto_activity_flows=[
            "09.01.01 Electricity plants",
            "09.01.02 CHP plants",
            "09.01.03 Heat plants",
            "09.02.01 Electricity plants",
            "09.02.02 CHP plants",
            "09.02.03 Heat plants",
        ],
        esto_activity_exact_products=["17 Electricity", "18 Heat"],
        ninth_activity_sectors=[
            "09_01_electricity_plants",
            "09_02_chp_plants",
            "09_x_heat_plants",
        ],
        ninth_activity_fuels=["17_electricity", "18_heat"],
        ninth_activity_subfuels=["17_electricity", "18_heat"],
        leap_balance_rows=["Electricity Generation", "CHP plants", "Heat plants"],
        leap_balance_fuel_set="electricity_heat_output",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.01 Electricity, CHP and heat plants"],
        ninth_target_sectors=["10_01_01_electricity_chp_and_heat_plants"],
        notes="Starter proxy: total electricity plus heat output from main/autoproducer electricity, CHP, and heat plants.",
    ),
    make_proxy_config(
        enabled=False,#can be handled by auxiliary branch in leap so keeping disabled for now
        process_key="gas_works_plants",
        process_label="Gas works plants",
        activity_label="Gas works gas output",
        esto_activity_flows=["09.06.01 Gas works plants"],
        ninth_activity_sectors=["09_06_01_gas_works_plants"],
        leap_balance_rows=["Gas works plants"],
        leap_balance_fuel_set="gas_works_gas_output",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.02 Gas works plants"],
        ninth_target_sectors=["10_01_02_gas_works_plants"],
        notes="Starter proxy: positive gas works gas output. If a LEAP balance export has no Gas works gas column, LEAP-balance mode returns zero for this proxy instead of failing.",
    ),
    make_proxy_config(
        enabled=True,#cannot be easily handled by auxiliary branch in leap since it includes both liquefaction and regasification, so keeping enabled
        process_key="liquefaction_regasification_plants",
        process_label="Liquefaction/regasification plants",
        activity_label="Natural gas and LNG throughput/output",
        esto_activity_flows=["09.06.02 Liquefaction/regasification plants"],
        esto_activity_exact_products=["08.01 Natural gas", "08.02 LNG"],
        ninth_activity_sectors=["09_06_02_liquefaction_regasification_plants"],
        ninth_activity_fuels=["08_gas"],
        ninth_activity_subfuels=["08_01_natural_gas", "08_02_lng"],
        leap_balance_rows=["NG Liquefaction", "LNG regasification"],
        leap_balance_fuel_set="natural_gas_lng_output",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.03 Liquefaction/regasification plants"],
        ninth_target_sectors=["10_01_03_liquefaction_regasification_plants"],
        notes="Starter proxy: positive output of natural gas and LNG from liquefaction/regasification.",
    ),
    make_proxy_config(
        enabled=False,#can be handled by auxiliary branch in leap so keeping disabled for now
        process_key="gas_to_liquids_plants",
        process_label="Gas-to-liquids plants",
        activity_label="Gas-to-liquids petroleum-product output",
        esto_activity_flows=["09.06.04 Gas-to-liquids plants"],
        esto_activity_product_prefixes=["07."],
        ninth_activity_sectors=["09_06_04_gastoliquids_plants"],
        ninth_activity_fuels=["07_petroleum_products"],
        ninth_activity_subfuels=[
            "07_06_kerosene",
            "07_07_gas_diesel_oil",
            "07_09_lpg",
            "07_10_refinery_gas_not_liquefied",
            "07_x_other_petroleum_products",
        ],
        leap_balance_rows=["Gas to liquids plants"],
        leap_balance_fuel_set="gas_to_liquids_output",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.04 Gas-to-liquids plants"],
        ninth_target_sectors=["10_01_04_gas_to_liquids_plants"],
        notes="Starter proxy: positive petroleum-product output from gas-to-liquids plants.",
    ),
    make_proxy_config(
        enabled=False,#can be handled by auxiliary branch in leap so keeping disabled for now
        process_key="coke_ovens",
        process_label="Coke ovens",
        activity_label="Coke oven coal-product output",
        esto_activity_flows=["09.08.01 Coke ovens"],
        ninth_activity_sectors=["09_08_01_coke_ovens"],
        leap_balance_rows=["Coke ovens"],
        leap_balance_fuel_set="coke_oven_output",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.05 Coke ovens"],
        ninth_target_sectors=["10_01_05_coke_ovens"],
        notes="Starter proxy: positive coke oven coke, coke oven gas, and coal tar output. If this detailed 9th proxy is zero while ESTO activity exists, the workflow falls back to broader parent 9th activity.",
    ),
    make_proxy_config(
        enabled=False,#can be handled by auxiliary branch in leap so keeping disabled for now
        process_key="blast_furnaces",
        process_label="Blast furnaces",
        activity_label="Blast furnace gas output",
        esto_activity_flows=["09.08.02 Blast furnaces"],
        ninth_activity_sectors=["09_08_02_blast_furnaces"],
        leap_balance_rows=["Blast furnaces"],
        leap_balance_fuel_set="blast_furnace_output",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.07 Blast furnaces"],
        ninth_target_sectors=["10_01_07_blast_furnaces"],
        notes="Starter proxy: positive blast furnace gas output. If this detailed 9th proxy is zero while ESTO activity exists, the workflow falls back to broader parent 9th activity.",
    ),
    make_proxy_config(
        enabled=False,#can be handled by auxiliary branch in leap so keeping disabled for now
        process_key="patent_fuel_plants",
        process_label="Patent fuel plants",
        activity_label="Patent fuel output",
        esto_activity_flows=["09.08.03 Patent fuel plants"],
        ninth_activity_sectors=["09_08_03_patent_fuel_plants"],
        leap_balance_rows=["Patent fuel plants"],
        leap_balance_fuel_set="patent_fuel_output",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.08 Patent fuel plants"],
        ninth_target_sectors=["10_01_08_patent_fuel_plants"],
        notes="Starter proxy: positive patent fuel output. If this detailed 9th proxy is zero while ESTO activity exists, the workflow falls back to broader parent 9th activity.",
    ),
    make_proxy_config(
        enabled=False,#can be handled by auxiliary branch in leap so keeping disabled for now, but enabling here to have the data in the export workbook for now since it is an esto own-use/loss flow in the 2024 dataset
        process_key="bkb_pb_plants",
        process_label="BKB/PB plants",
        activity_label="BKB/PB output",
        esto_activity_flows=["09.08.04 BKB/PB plants"],
        ninth_activity_sectors=["09_08_04_bkb_pb_plants"],
        leap_balance_rows=["BKB and PB plants"],
        leap_balance_fuel_set="bkb_pb_output",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.09 BKB/PB plants"],
        ninth_target_sectors=["10_01_09_bkb_pb_plants"],
        notes="Starter proxy: positive BKB/PB output. If this detailed 9th proxy is zero while ESTO activity exists, the workflow falls back to broader parent 9th activity.",
    ),
    make_proxy_config(
        enabled=False,#can be handled by auxiliary branch in leap so keeping disabled for now
        process_key="liquefaction_plants_coal_to_oil",
        process_label="Liquefaction plants (Coal to Oil)",
        activity_label="Coal-to-oil petroleum-product output",
        esto_activity_flows=["09.08.05 Liquefaction (coal to oil)"],
        esto_activity_product_prefixes=["07."],
        ninth_activity_sectors=["09_08_05_liquefaction_coal_to_oil"],
        ninth_activity_fuels=["07_petroleum_products"],
        ninth_activity_subfuels=["07_09_lpg", "07_03_naphtha", "07_07_gas_diesel_oil", "07_08_fuel_oil", "07_x_other_petroleum_products"],
        leap_balance_rows=["Liquefaction coal to oil"],
        leap_balance_fuel_set="coal_to_oil_output",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.10 Liquefaction plants (Coal to Oil)"],
        ninth_target_sectors=["10_01_10_liquefaction_plants_coal_to_oil"],
        notes="Starter proxy: positive petroleum-product output from coal-to-oil liquefaction. If this detailed 9th proxy is zero while ESTO activity exists, the workflow falls back to broader parent 9th activity.",
    ),
    make_proxy_config(
        enabled=False,#Can be handled by auxiliary branch in leap so keeping disabled for now
        process_key="oil_refineries",
        process_label="Oil Refining",
        activity_label="Oil refinery petroleum-product output",
        esto_activity_flows=["09.07 Oil refineries"],
        esto_activity_product_prefixes=["07."],
        ninth_activity_sectors=["09_07_oil_refineries"],
        ninth_activity_fuels=["07_petroleum_products"],
        ninth_activity_subfuels=[
            "07_01_motor_gasoline",
            "07_02_aviation_gasoline",
            "07_03_naphtha",
            "07_06_kerosene",
            "07_07_gas_diesel_oil",
            "07_08_fuel_oil",
            "07_09_lpg",
            "07_10_refinery_gas_not_liquefied",
            "07_11_ethane",
            "07_16_petroleum_coke",
            "07_petroleum_products_unallocated",
            "07_x_jet_fuel",
            "07_x_other_petroleum_products",
        ],
        leap_balance_rows=["Oil Refining"],
        leap_balance_fuel_set="oil_refinery_output",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.11 Oil refineries"],
        ninth_target_sectors=["10_01_11_oil_refineries"],
        notes="Starter proxy: positive petroleum-product output from oil refineries.",
    ),
    make_proxy_config(
        enabled=True,#cannot be easily handled by auxiliary branch in leap since it includes both production and processing own-use/losses, so keeping enabled
        process_key="oil_and_gas_extraction",
        process_label="Oil and gas extraction",
        activity_label="Primary oil and gas production",
        esto_activity_flows=["01 Production"],
        esto_activity_exact_products=[
            "06.01 Crude oil",
            "06.02 Natural gas liquids",
            "06.05 Other hydrocarbons",
            "08.01 Natural gas",
            "08.02 LNG",
            "08.03 Gas works gas",
            "08.99 Gas nonspecified",
        ],
        ninth_activity_sectors=["01_production"],
        ninth_activity_fuels=["06_crude_oil_and_ngl", "08_gas"],
        ninth_activity_subfuels=[
            "06_01_crude_oil",
            "06_02_natural_gas_liquids",
            "06_crude_oil_and_ngl_unallocated",
            "06_x_other_hydrocarbons",
            "08_01_natural_gas",
            "08_02_lng",
            "08_03_gas_works_gas",
            "08_gas_unallocated",
        ],
        leap_balance_rows=["Production"],
        leap_balance_fuel_set="oil_and_gas_primary_production",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.12 Oil and gas extraction"],
        ninth_target_sectors=["10_01_12_oil_and_gas_extraction"],
        notes="Starter proxy: primary production of crude oil/NGL/other hydrocarbons and gas products.",
    ),
    make_proxy_config(
        enabled=True,#cannot be easily handled by auxiliary branch in leap since it includes both pumping and generation, so keeping enabled
        process_key="pump_storage_plants",
        process_label="Pump storage plants",
        activity_label="Pump storage electricity output",
        ninth_activity_sectors=["09_01_05_03_pump"],
        ninth_activity_fuels=["17_electricity"],
        ninth_activity_subfuels=["17_electricity"],
        leap_balance_rows=["Pumped hydro"],
        leap_balance_fuel_set="electricity_output",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.13 Pump storage plants"],
        ninth_target_sectors=["10_01_13_pump_storage_plants"],
        notes="Starter proxy: 9th pump-storage electricity output. ESTO and simple LEAP balance exports do not isolate pump storage, so LEAP balance mode falls back to total electricity production unless refined later.",
    ),
    make_proxy_config(
        enabled=True,#cannot be easily handled by auxiliary branch in leap since it includes both pumping and generation, so keeping enabled
        process_key="nuclear_industry",
        process_label="Nuclear industry",
        activity_label="Primary nuclear production",
        esto_activity_flows=["01 Production"],
        esto_activity_exact_products=["09 Nuclear"],
        ninth_activity_sectors=["01_production"],
        ninth_activity_fuels=["09_nuclear"],
        ninth_activity_subfuels=["09_nuclear"],
        leap_balance_rows=["Production"],
        leap_balance_fuel_set="nuclear_primary_production",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.14 Nuclear industry"],
        ninth_target_sectors=["10_01_14_nuclear_industry"],
        notes="Starter proxy: primary nuclear production.",
    ),
    make_proxy_config(
        enabled=False,#can be handled by auxiliary branch in leap so keeping disabled for now
        process_key="charcoal_production_plants",
        process_label="Charcoal production plants",
        activity_label="Charcoal output",
        esto_activity_flows=["09.11 Charcoal processing"],
        ninth_activity_sectors=["09_11_charcoal_processing"],
        leap_balance_rows=["Charcoal processing"],
        leap_balance_fuel_set="charcoal_output",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.15 Charcoal production plants"],
        ninth_target_sectors=["10_01_15_charcoal_production_plants"],
        notes="Starter proxy: positive charcoal output from charcoal processing.",
    ),
    make_proxy_config(
        enabled=True,#not a transformation module in LEAP, so cannot be handled by auxiliary branch. 
        process_key="gasification_plants_for_biogases",
        process_label="Gasification plants for biogases",
        activity_label="Biogas production",
        esto_activity_flows=["01 Production"],
        esto_activity_exact_products=["16.01 Biogas"],
        ninth_activity_sectors=["01_production"],
        ninth_activity_fuels=["16_others"],
        ninth_activity_subfuels=["16_01_biogas"],
        leap_balance_rows=["Production"],
        leap_balance_fuel_set="biogas_output",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.16 Gasification plants for biogases"],
        ninth_target_sectors=["10_01_16_gasification_plants_for_biogases"],
        notes="Starter proxy: biogas production. There is no direct ESTO 09 gasification activity row, so this uses primary biogas production.",
    ),
    make_proxy_config(
        enabled=False,#can be handled by auxiliary branch in leap so keeping disabled for now
        process_key="nonspecified_own_uses",
        process_label="Non-specified own uses",
        activity_label="Non-specified transformation positive output",
        esto_activity_flows=["09.12 Non-specified transformation"],
        ninth_activity_sectors=["09_12_nonspecified_transformation"],
        leap_balance_rows=["Non specified transformation"],
        leap_balance_fuel_set="nonspecified_transformation_output",
        activity_value_mode="positive_only",
        esto_target_flows=["10.01.17 Non-specified own uses"],
        ninth_target_sectors=["10_01_17_nonspecified_own_uses"],
        notes="Starter proxy: positive output from non-specified transformation. Fuel set is based on nonzero ESTO/9th activity fuels across all economies.",
    ),
    make_proxy_config(
        enabled=True,#besides electricity own-use/losses, transmission and distribution losses for other fuels cannot be easily isolated in LEAP balances or auxiliary branches, so keeping enabled for now
        process_key="transmission_and_distribution_losses",
        process_label="Transmission and distribution losses",
        leap_process_label="Transmission and distribution loss",
        activity_label="Total production excluding electricity",
        esto_activity_flows=["01 Production"],
        esto_activity_exclude_products=["17 Electricity"],
        ninth_activity_sectors=["01_production"],
        ninth_activity_exclude_fuels=["17_electricity"],
        ninth_activity_exclude_subfuels=["17_electricity"],
        leap_balance_rows=["Production"],
        leap_balance_fuel_set="production_ex_electricity",
        activity_value_mode="positive_only",
        esto_target_flows=["10.02 Transmission and distribution losses"],
        esto_target_exclude_products=["17 Electricity"],
        ninth_target_sectors=["10_02_transmission_and_distribution_losses"],
        ninth_target_exclude_fuels=["17_electricity"],
        ninth_target_exclude_subfuels=["17_electricity"],
        notes="Starter proxy: total positive production excluding electricity, so LEAP-balance activity is not driven by produced electricity.",
    ),
    make_proxy_config(
        enabled=False,#CCS own-use/losses are not clearly isolated in either ESTO or 9th, so keeping disabled for now
        process_key="ccs",
        process_label="CCS",
        esto_target_flows=[],
        ninth_target_sectors=["10_01_18_ccs"],
        notes="9th has 10_01_18_ccs, but ESTO 00APEC_2024_low.csv has no matching 10.01.18 flow.",
    ),
]


def _resolve(path_value: Path | str) -> Path:
    text = str(path_value).replace("\\", "/")
    path = Path(text)
    return path if path.is_absolute() else (REPO_ROOT / path).resolve()


def _year_columns(df: pd.DataFrame) -> list[int]:
    years = []
    for col in df.columns:
        if str(col).isdigit():
            years.append(int(col))
    return sorted(set(years))


def _normalize_year_columns(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns={col: int(col) for col in df.columns if str(col).isdigit()})


def _normalize_economy(value: object) -> str:
    text = str(value or "").strip().upper()
    if len(text) >= 5 and text[:2].isdigit() and text[2] != "_":
        return f"{text[:2]}_{text[2:]}"
    return text


def _compact_economy(value: object) -> str:
    return _normalize_economy(value).replace("_", "")


def _clean_ninth_name(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    if text == "x":
        return text
    parts = text.split("_")
    while parts and (parts[0].isdigit() or parts[0] == "x"):
        parts.pop(0)
    label = " ".join(part for part in parts if part).strip()
    return _sentence_case_fuel_label(label) if label else text


def _sentence_case_fuel_label(value: object) -> str:
    """Apply the usual LEAP fuel-label style without title-casing every word."""
    text = " ".join(str(value or "").strip().split())
    if not text:
        return ""
    protected = {
        "BKB",
        "BKB/PB",
        "CCS",
        "CHP",
        "LNG",
        "LPG",
        "NGL",
        "PB",
    }

    def _format_token(token: str, *, is_first: bool) -> str:
        if not token:
            return token
        if "/" in token:
            return "/".join(_format_token(part, is_first=is_first and idx == 0) for idx, part in enumerate(token.split("/")))
        if token.upper() in protected:
            return token.upper()
        lower = token.lower()
        return lower[:1].upper() + lower[1:] if is_first else lower

    return " ".join(_format_token(token, is_first=idx == 0) for idx, token in enumerate(text.split(" ")))


def _clean_product_for_branch(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    if "_" in text and " " not in text:
        return _clean_ninth_name(text)
    return _sentence_case_fuel_label(clean_fuel_label_for_leap(text))


def _format_fuel_branch_label(
    value: object,
    *,
    source_name: str = "",
    fuel_mapping_lookup: Mapping[str, Mapping[str, str]] | None = None,
) -> str:
    """Return the LEAP branch fuel label, using mappings before fallback cleanup."""
    text = str(value or "").strip()
    if not text:
        return ""
    mapped = ""
    if fuel_mapping_lookup and source_name:
        mapped = fuel_mapping_lookup.get(source_name, {}).get(_normalize_source_token(text), "")
    if _normalize_source_token(text) in {"17 electricity", "17_electricity"} and _normalize_balance_label(mapped) == "green electricity":
        mapped = "Electricity"
    label = mapped or _clean_product_for_branch(text)
    return sanitize_leap_name(_sentence_case_fuel_label(label))


def _source_fuel_label(value: object) -> str:
    """Normalize ESTO/9th fuel labels to comparable LEAP-style labels."""
    text = str(value or "").strip()
    if not text or text == "x":
        return ""
    if "_" in text and " " not in text:
        return sanitize_leap_name(_clean_ninth_name(text))
    return sanitize_leap_name(clean_fuel_label_for_leap(text))


def _normalize_source_token(value: object) -> str:
    return str(value or "").strip().lower()


def _normalize_balance_label(value: object) -> str:
    return normalize_balance_label(value)


def _series_from_group(group: pd.DataFrame, year_cols: Sequence[int]) -> dict[int, float]:
    if group.empty:
        return {int(year): 0.0 for year in year_cols}
    values = group[list(year_cols)].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum()
    return {int(year): float(values.get(year, 0.0)) for year in year_cols}


def _apply_value_mode(df: pd.DataFrame, year_cols: Sequence[int], value_mode: str) -> pd.DataFrame:
    """Apply source-side sign handling before summing activity values."""
    if df.empty or not year_cols:
        return df
    mode = str(value_mode or "signed_sum").strip().lower()
    out = df.copy()
    year_list = list(year_cols)
    values = out[year_list].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    if mode in {"signed", "signed_sum", ""}:
        out[year_list] = values
        return out
    if mode in {"positive", "positive_only", "outputs"}:
        out[year_list] = values.where(values > 0.0, 0.0)
        return out
    if mode in {"negative_abs", "input_abs", "inputs_abs"}:
        out[year_list] = values.where(values < 0.0, 0.0).abs()
        return out
    if mode in {"absolute", "abs"}:
        out[year_list] = values.abs()
        return out
    raise ValueError(f"Invalid activity value_mode={value_mode!r}.")


def _has_prefix(value: object, prefixes: Sequence[str]) -> bool:
    text = str(value or "").strip()
    return any(text.startswith(prefix) for prefix in prefixes)


def _matches_exact_values(value: object, allowed_values: Sequence[str]) -> bool:
    if not allowed_values:
        return False
    text = str(value or "").strip().lower()
    allowed = {str(item or "").strip().lower() for item in allowed_values if str(item or "").strip()}
    return text in allowed


def _drop_ninth_parent_fuel_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Drop parent fuel rows where child subfuel rows exist in the same group."""
    if df.empty or "subfuels" not in df.columns or "fuels" not in df.columns:
        return df
    group_cols = [
        col
        for col in [
            "scenarios",
            "economy",
            "sectors",
            "sub1sectors",
            "sub2sectors",
            "sub3sectors",
            "sub4sectors",
            "fuels",
        ]
        if col in df.columns
    ]
    if not group_cols:
        return df
    out_parts = []
    for _, group in df.groupby(group_cols, dropna=False):
        subfuels = group["subfuels"].fillna("").astype(str).str.strip()
        has_child = (subfuels != "x").any()
        if has_child:
            group = group[subfuels != "x"].copy()
        out_parts.append(group)
    if not out_parts:
        return df.iloc[0:0].copy()
    return pd.concat(out_parts, ignore_index=True)


def _drop_ninth_subtotals(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "subtotal_results" in out.columns:
        out = out[out["subtotal_results"] == False].copy()
    total_codes = {"19_total", "20_total_renewables", "21_modern_renewables"}
    if "fuels" in out.columns:
        out = out[~out["fuels"].astype(str).isin(total_codes)].copy()
    if "subfuels" in out.columns:
        out = out[~out["subfuels"].astype(str).isin(total_codes)].copy()
    return out


def _select_ninth_sector(df: pd.DataFrame, sector_codes: Sequence[str]) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    hierarchy_cols = ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]
    mask = pd.Series(False, index=df.index)
    wanted = {str(code).strip().lower() for code in sector_codes}
    for col in hierarchy_cols:
        if col in df.columns:
            mask = mask | df[col].fillna("").astype(str).str.strip().str.lower().isin(wanted)
    return df[mask].copy()


def _next_parent_ninth_sector_codes(
    ninth_data: pd.DataFrame,
    sector_codes: Sequence[str],
) -> list[str]:
    """Return immediate parent sector codes for configured 9th hierarchy codes."""
    if ninth_data.empty or not sector_codes:
        return []
    hierarchy_cols = [
        col
        for col in ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]
        if col in ninth_data.columns
    ]
    wanted = {str(code).strip().lower() for code in sector_codes if str(code).strip()}
    if not hierarchy_cols or not wanted:
        return []
    parents: list[str] = []
    seen: set[str] = set()
    for _, row in ninth_data.iterrows():
        values = [str(row.get(col) or "").strip() for col in hierarchy_cols]
        lowered = [value.lower() for value in values]
        matching_indexes = [idx for idx, value in enumerate(lowered) if value in wanted]
        if not matching_indexes:
            continue
        match_idx = max(matching_indexes)
        for parent_idx in range(match_idx - 1, -1, -1):
            parent = values[parent_idx]
            parent_key = parent.lower()
            if not parent or parent_key == "x" or parent_key in wanted:
                continue
            if parent_key not in seen:
                seen.add(parent_key)
                parents.append(parent)
            break
    return parents


def load_esto_data(path: Path | str = ESTO_DATA_PATH) -> pd.DataFrame:
    df = pd.read_csv(_resolve(path), low_memory=False)
    df = _normalize_year_columns(df)
    df["economy_key"] = df["economy"].apply(_normalize_economy)
    try:
        labeled = apply_matt_subtotal_mapping(df, _resolve(ESTO_SUBTOTAL_MAPPING_PATH))
        df = filter_matt_subtotals(labeled)
    except Exception as exc:
        print(f"[WARN] ESTO subtotal filtering skipped: {exc}")
    total_products = {"19 Total", "20 Total Renewables", "21 Modern renewables"}
    if "products" in df.columns:
        df = df[~df["products"].astype(str).isin(total_products)].copy()
    return df


def load_ninth_data(path: Path | str = NINTH_DATA_PATH) -> pd.DataFrame:
    df = pd.read_csv(_resolve(path), low_memory=False)
    df = _normalize_year_columns(df)
    if "economy" in df.columns:
        df["economy_key"] = df["economy"].apply(_normalize_economy)
    if "scenarios" in df.columns:
        df = df[df["scenarios"].astype(str).str.lower() == NINTH_SCENARIO].copy()
    return _drop_ninth_subtotals(df)


def load_output_fuel_validation_esto_tables(
    paths: Sequence[Path | str] = OUTPUT_FUEL_VALIDATION_ESTO_PATHS,
) -> list[tuple[str, pd.DataFrame]]:
    """Load ESTO snapshot tables used to validate output branch fuel coverage."""
    tables: list[tuple[str, pd.DataFrame]] = []
    for path_value in paths:
        path = _resolve(path_value)
        df = pd.read_csv(path, low_memory=False)
        df = _normalize_year_columns(df)
        tables.append((path.name, df))
    return tables


def load_fuel_mapping_lookup(mapping_path: Path | str = LEAP_MAPPINGS_PATH) -> dict[str, dict[str, str]]:
    """Load ESTO/9th fuel -> LEAP fuel lookups from config/leap_mappings.xlsx."""
    path = _resolve(mapping_path)
    lookups = {"esto": {}, "ninth": {}}
    if not path.exists():
        return lookups
    try:
        esto = pd.read_excel(path, sheet_name="fuel_product_final_proposed", dtype=str).fillna("")
    except Exception:
        esto = pd.DataFrame()
    for _, row in esto.iterrows():
        source = str(row.get("esto_product", "") or "").strip()
        leap = str(row.get("leap_fuel_name", "") or "").strip()
        if source and leap:
            lookups["esto"][_normalize_source_token(source)] = sanitize_leap_name(leap)
    try:
        ninth = pd.read_excel(path, sheet_name="fuel_ninth_final_proposed", dtype=str).fillna("")
    except Exception:
        ninth = pd.DataFrame()
    for _, row in ninth.iterrows():
        source = str(row.get("ninth_fuel", "") or "").strip()
        leap = str(row.get("leap_fuel_name", "") or "").strip()
        if source and leap:
            lookups["ninth"][_normalize_source_token(source)] = sanitize_leap_name(leap)
    return lookups


def resolve_leap_balance_workbook_path(
    *,
    economy: str,
    scenario: str = LEAP_BALANCE_SCENARIO,
    date_id: str | None = LEAP_BALANCE_DATE_ID,
    workbook_path: Path | str | None = LEAP_BALANCE_WORKBOOK_PATH,
) -> Path:
    """Resolve the LEAP balance workbook used for second-run proxy activity."""
    if workbook_path:
        return _resolve(workbook_path)
    return resolve_balance_export_workbook(
        economy=_normalize_economy(economy),
        scenario=scenario,
        date_id=date_id,
        exports_root=_resolve(LEAP_BALANCE_EXPORTS_ROOT),
    )


def load_leap_balance_activity_table(
    workbook_path: Path | str,
    *,
    balance_rows: Sequence[str],
    fuels: Sequence[str],
) -> pd.DataFrame:
    """Return long LEAP balance values for selected row labels and fuel columns."""
    return _load_shared_leap_balance_activity_table(
        _resolve(workbook_path),
        balance_rows=balance_rows,
        fuels=fuels,
    )


def build_leap_balance_proxy_activity_series(
    *,
    leap_balance_activity: pd.DataFrame,
    config: Mapping[str, object],
    base_year: int,
    final_year: int,
) -> dict[int, float]:
    """Sum configured LEAP balance fuels/rows into one proxy activity series."""
    leap_cfg = config["activity_sources"].get("leap_balance", {})
    fuel_set_name = str(leap_cfg.get("fuel_set", "")).strip()
    fuels = LEAP_BALANCE_FUEL_SETS.get(fuel_set_name, [])
    series = build_leap_balance_activity_series(
        leap_balance_activity,
        balance_rows=leap_cfg.get("balance_rows", []),
        fuels=fuels,
        value_mode=str(leap_cfg.get("value_mode", "signed_sum") or "signed_sum"),
        base_year=base_year,
        final_year=final_year,
    )
    if any(abs(float(value)) > 0.0 for value in series.values()):
        return series

    process_key = str(config.get("process_key", "")).strip()
    fallback_cfg = LEAP_BALANCE_ACTIVITY_FALLBACKS.get(process_key)
    if not isinstance(fallback_cfg, dict):
        return series

    fallback_rows = [str(item) for item in fallback_cfg.get("balance_rows", []) if str(item).strip()]
    fallback_fuel_set = str(fallback_cfg.get("fuel_set", "")).strip()
    fallback_fuels = LEAP_BALANCE_FUEL_SETS.get(fallback_fuel_set, [])
    if not fallback_rows or not fallback_fuels:
        return series

    fallback_series = build_leap_balance_activity_series(
        leap_balance_activity,
        balance_rows=fallback_rows,
        fuels=fallback_fuels,
        value_mode=str(fallback_cfg.get("value_mode", leap_cfg.get("value_mode", "signed_sum")) or "signed_sum"),
        base_year=base_year,
        final_year=final_year,
    )
    if any(abs(float(value)) > 0.0 for value in fallback_series.values()):
        print(
            "[INFO] LEAP-balance proxy activity fallback applied for "
            f"{process_key}: rows={fallback_rows}, fuel_set={fallback_fuel_set}."
        )
        return fallback_series

    return series


def build_proxy_activity_series(
    *,
    esto_data: pd.DataFrame,
    ninth_data: pd.DataFrame,
    economy: str,
    config: Mapping[str, object],
    base_year: int,
    final_year: int,
) -> dict[int, float]:
    esto_activity = build_esto_proxy_activity_series(
        esto_data=esto_data,
        economy=economy,
        config=config,
        base_year=base_year,
    )
    ninth_activity = build_ninth_proxy_activity_series(
        ninth_data=ninth_data,
        economy=economy,
        config=config,
        base_year=base_year,
        final_year=final_year,
    )
    esto_cfg = config["activity_sources"]["esto"]
    has_esto_activity_definition = bool(esto_cfg.get("flows", []))
    has_esto_activity = any(abs(float(value)) > 0.0 for value in esto_activity.values())
    if has_esto_activity_definition and not has_esto_activity:
        ninth_activity = {year: 0.0 for year in ninth_activity}
    activity = dict(esto_activity)
    activity.update(ninth_activity)
    wanted_years = sorted(set(esto_activity) | set(ninth_activity) | {int(base_year)})
    return {int(year): activity.get(int(year), 0.0) for year in wanted_years}


def build_esto_proxy_activity_series(
    *,
    esto_data: pd.DataFrame,
    economy: str,
    config: Mapping[str, object],
    base_year: int,
) -> dict[int, float]:
    economy_key = _normalize_economy(economy)
    compact_economy = _compact_economy(economy)
    activity_sources = config["activity_sources"]

    esto_cfg = activity_sources["esto"]
    esto_years = [year for year in _year_columns(esto_data) if int(year) <= int(base_year)]
    esto_subset = esto_data[
        (esto_data["economy"].astype(str).str.upper().isin([compact_economy, economy_key.replace("_", "")]))
        | (esto_data["economy_key"] == economy_key)
    ].copy()
    esto_subset = esto_subset[esto_subset["flows"].isin(esto_cfg.get("flows", []))].copy()
    product_prefixes = list(esto_cfg.get("product_prefixes", []))
    exact_products = set(esto_cfg.get("include_exact_products", []))
    if product_prefixes or exact_products:
        product_mask = (
            esto_subset["products"]
            .apply(lambda value: _has_prefix(value, product_prefixes) or str(value) in exact_products)
            .fillna(False)
            .astype(bool)
        )
        esto_subset = esto_subset[product_mask].copy()
    exclude_products = {str(value).strip().lower() for value in esto_cfg.get("exclude_products", [])}
    if exclude_products:
        esto_subset = esto_subset[
            ~esto_subset["products"].fillna("").astype(str).str.strip().str.lower().isin(exclude_products)
        ].copy()
    esto_subset = _apply_value_mode(esto_subset, esto_years, str(esto_cfg.get("value_mode", "signed_sum")))
    return _series_from_group(esto_subset, esto_years)


def build_ninth_proxy_activity_series(
    *,
    ninth_data: pd.DataFrame,
    economy: str,
    config: Mapping[str, object],
    base_year: int,
    final_year: int,
) -> dict[int, float]:
    series, _ = build_ninth_proxy_activity_series_with_fallback(
        ninth_data=ninth_data,
        economy=economy,
        config=config,
        base_year=base_year,
        final_year=final_year,
    )
    return series


def _sum_ninth_proxy_activity_series(
    *,
    ninth_data: pd.DataFrame,
    economy_key: str,
    ninth_cfg: Mapping[str, object],
    sector_codes: Sequence[str],
    projection_years: Sequence[int],
    ignore_fuel_filters: bool = False,
) -> dict[int, float]:
    """Sum 9th activity for one candidate sector set."""
    ninth_subset = ninth_data[ninth_data["economy_key"] == economy_key].copy()
    ninth_subset = _select_ninth_sector(ninth_subset, sector_codes)
    fuel_values = [] if ignore_fuel_filters else list(ninth_cfg.get("fuels", []))
    subfuel_values = [] if ignore_fuel_filters else list(ninth_cfg.get("subfuels", []))
    if fuel_values or subfuel_values:
        ninth_subset = _drop_ninth_parent_fuel_rows(ninth_subset)
        fuel_mask = (
            ninth_subset["fuels"]
            .apply(lambda value: _matches_exact_values(value, fuel_values))
            .fillna(False)
            .astype(bool)
        )
        subfuel_mask = (
            ninth_subset["subfuels"]
            .apply(lambda value: _matches_exact_values(value, subfuel_values))
            .fillna(False)
            .astype(bool)
        )
        ninth_subset = ninth_subset[fuel_mask | subfuel_mask].copy()
    exclude_fuels = {str(value).strip().lower() for value in ninth_cfg.get("exclude_fuels", [])}
    exclude_subfuels = {str(value).strip().lower() for value in ninth_cfg.get("exclude_subfuels", [])}
    if exclude_fuels:
        ninth_subset = ninth_subset[
            ~ninth_subset["fuels"].fillna("").astype(str).str.strip().str.lower().isin(exclude_fuels)
        ].copy()
    if exclude_subfuels:
        ninth_subset = ninth_subset[
            ~ninth_subset["subfuels"].fillna("").astype(str).str.strip().str.lower().isin(exclude_subfuels)
        ].copy()
    ninth_subset = _apply_value_mode(ninth_subset, projection_years, str(ninth_cfg.get("value_mode", "signed_sum")))
    return _series_from_group(ninth_subset, projection_years)


def _series_has_nonzero_value(series: Mapping[int, float]) -> bool:
    return any(abs(float(value)) > 0.0 for value in series.values())


def build_ninth_proxy_activity_series_with_fallback(
    *,
    ninth_data: pd.DataFrame,
    economy: str,
    config: Mapping[str, object],
    base_year: int,
    final_year: int,
) -> tuple[dict[int, float], dict[str, object] | None]:
    """Build 9th activity, climbing to broader parent sectors if the configured proxy is zero."""
    economy_key = _normalize_economy(economy)
    activity_sources = config["activity_sources"]
    ninth_cfg = activity_sources["ninth"]
    projection_years = [year for year in range(int(base_year) + 1, int(final_year) + 1) if year in ninth_data.columns]
    original_sector_codes = [str(code) for code in ninth_cfg.get("sector_codes", []) if str(code).strip()]
    series = _sum_ninth_proxy_activity_series(
        ninth_data=ninth_data,
        economy_key=economy_key,
        ninth_cfg=ninth_cfg,
        sector_codes=original_sector_codes,
        projection_years=projection_years,
    )
    if _series_has_nonzero_value(series):
        return series, None

    tried = {str(code).strip().lower() for code in original_sector_codes if str(code).strip()}
    candidate_sector_codes = original_sector_codes
    fallback_level = 0
    while candidate_sector_codes:
        parent_sector_codes = _next_parent_ninth_sector_codes(ninth_data, candidate_sector_codes)
        parent_sector_codes = [
            code
            for code in parent_sector_codes
            if str(code).strip().lower() not in tried
        ]
        if not parent_sector_codes:
            break
        fallback_level += 1
        tried.update(str(code).strip().lower() for code in parent_sector_codes if str(code).strip())
        parent_series = _sum_ninth_proxy_activity_series(
            ninth_data=ninth_data,
            economy_key=economy_key,
            ninth_cfg=ninth_cfg,
            sector_codes=parent_sector_codes,
            projection_years=projection_years,
            ignore_fuel_filters=True,
        )
        if _series_has_nonzero_value(parent_series):
            return parent_series, {
                "economy": economy_key,
                "process_key": str(config.get("process_key", "")),
                "process_label": str(config.get("process_label", "")),
                "activity_label": str(config.get("activity_label", "")),
                "original_ninth_activity_sectors": "; ".join(original_sector_codes),
                "fallback_ninth_activity_sectors": "; ".join(parent_sector_codes),
                "fallback_level": fallback_level,
                "fallback_reason": "configured_9th_activity_all_zero",
                "fallback_uses_broad_parent_activity": True,
                "fallback_activity_total": sum(abs(float(value)) for value in parent_series.values()),
            }
        candidate_sector_codes = parent_sector_codes

    return series, None


def build_activity_source_gap_warnings(
    *,
    esto_data: pd.DataFrame,
    ninth_data: pd.DataFrame,
    economy: str,
    configs: Sequence[Mapping[str, object]],
    base_year: int,
    final_year: int,
) -> pd.DataFrame:
    """Identify proxies where ESTO activity exists but 9th projection activity is zero."""
    rows: list[dict[str, object]] = []
    for config in configs:
        if not bool(config.get("enabled", True)):
            continue
        esto_activity = build_esto_proxy_activity_series(
            esto_data=esto_data,
            economy=economy,
            config=config,
            base_year=base_year,
        )
        ninth_activity, fallback_info = build_ninth_proxy_activity_series_with_fallback(
            ninth_data=ninth_data,
            economy=economy,
            config=config,
            base_year=base_year,
            final_year=final_year,
        )
        esto_nonzero = {
            int(year): float(value)
            for year, value in esto_activity.items()
            if abs(float(value)) > 0.0
        }
        if not esto_nonzero:
            continue
        projection_years = [
            int(year)
            for year in range(int(base_year) + 1, int(final_year) + 1)
            if int(year) in ninth_activity
        ]
        if not projection_years:
            continue
        ninth_values = {year: float(ninth_activity.get(year, 0.0)) for year in projection_years}
        zero_years = [year for year, value in ninth_values.items() if abs(float(value)) <= 0.0]
        if len(zero_years) != len(projection_years):
            continue
        if fallback_info is not None:
            continue
        ninth_cfg = config["activity_sources"].get("ninth", {})
        esto_cfg = config["activity_sources"].get("esto", {})
        rows.append(
            {
                "economy": _normalize_economy(economy),
                "process_key": str(config.get("process_key", "")),
                "process_label": str(config.get("process_label", "")),
                "leap_process_label": str(config.get("leap_process_label", config.get("process_label", ""))),
                "activity_label": str(config.get("activity_label", "")),
                "warning_type": "esto_activity_nonzero_ninth_activity_all_zero",
                "esto_activity_years": "; ".join(str(year) for year in sorted(esto_nonzero)),
                "esto_activity_total": sum(abs(value) for value in esto_nonzero.values()),
                "ninth_projection_years": "; ".join(str(year) for year in projection_years),
                "ninth_activity_total": sum(abs(value) for value in ninth_values.values()),
                "esto_activity_flows": "; ".join(str(item) for item in esto_cfg.get("flows", [])),
                "ninth_activity_sectors": "; ".join(str(item) for item in ninth_cfg.get("sector_codes", [])),
                "ninth_activity_fuels": "; ".join(str(item) for item in ninth_cfg.get("fuels", [])),
                "ninth_activity_subfuels": "; ".join(str(item) for item in ninth_cfg.get("subfuels", [])),
                "notes": str(config.get("notes", "")),
            }
        )
    columns = [
        "economy",
        "process_key",
        "process_label",
        "leap_process_label",
        "activity_label",
        "warning_type",
        "esto_activity_years",
        "esto_activity_total",
        "ninth_projection_years",
        "ninth_activity_total",
        "esto_activity_flows",
        "ninth_activity_sectors",
        "ninth_activity_fuels",
        "ninth_activity_subfuels",
        "notes",
    ]
    if not rows:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(rows, columns=columns).sort_values(["process_label", "warning_type"], kind="mergesort")


def build_activity_source_fallback_report(
    *,
    esto_data: pd.DataFrame,
    ninth_data: pd.DataFrame,
    economy: str,
    configs: Sequence[Mapping[str, object]],
    base_year: int,
    final_year: int,
) -> pd.DataFrame:
    """Report 9th activity proxies that used a broader parent sector fallback."""
    rows: list[dict[str, object]] = []
    for config in configs:
        if not bool(config.get("enabled", True)):
            continue
        esto_activity = build_esto_proxy_activity_series(
            esto_data=esto_data,
            economy=economy,
            config=config,
            base_year=base_year,
        )
        has_esto_activity = any(abs(float(value)) > 0.0 for value in esto_activity.values())
        if not has_esto_activity:
            continue
        _, fallback_info = build_ninth_proxy_activity_series_with_fallback(
            ninth_data=ninth_data,
            economy=economy,
            config=config,
            base_year=base_year,
            final_year=final_year,
        )
        if fallback_info is not None:
            rows.append(fallback_info)
    columns = [
        "economy",
        "process_key",
        "process_label",
        "activity_label",
        "original_ninth_activity_sectors",
        "fallback_ninth_activity_sectors",
        "fallback_level",
        "fallback_reason",
        "fallback_uses_broad_parent_activity",
        "fallback_activity_total",
    ]
    if not rows:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(rows, columns=columns).sort_values(["process_label", "fallback_level"], kind="mergesort")


def build_activity_series_for_mode(
    *,
    esto_data: pd.DataFrame,
    ninth_data: pd.DataFrame,
    leap_balance_activity: pd.DataFrame | None,
    economy: str,
    config: Mapping[str, object],
    base_year: int,
    final_year: int,
    activity_source_mode: str,
) -> dict[int, float]:
    """Dispatch activity construction for first-run or LEAP-output second-run proxies."""
    mode = _normalize_activity_source_mode(activity_source_mode)
    if mode == "esto_ninth":
        return build_proxy_activity_series(
            esto_data=esto_data,
            ninth_data=ninth_data,
            economy=economy,
            config=config,
            base_year=base_year,
            final_year=final_year,
        )
    if mode == "leap_balance":
        if leap_balance_activity is None:
            raise ValueError("leap_balance_activity is required when activity_source_mode='leap_balance'.")
        return build_leap_balance_proxy_activity_series(
            leap_balance_activity=leap_balance_activity,
            config=config,
            base_year=base_year,
            final_year=final_year,
        )
    raise AssertionError(f"Unexpected normalized activity source mode: {mode}")


def _target_fuel_label_from_ninth(row: pd.Series) -> str:
    subfuel = str(row.get("subfuels", "") or "").strip()
    if subfuel and subfuel != "x":
        return subfuel
    return str(row.get("fuels", "") or "").strip()


def build_target_energy_long(
    *,
    esto_data: pd.DataFrame,
    ninth_data: pd.DataFrame,
    economy: str,
    config: Mapping[str, object],
    base_year: int,
    final_year: int,
    fuel_mapping_lookup: Mapping[str, Mapping[str, str]] | None = None,
) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    mapping_lookup = fuel_mapping_lookup or load_fuel_mapping_lookup()
    economy_key = _normalize_economy(economy)
    compact_economy = _compact_economy(economy)
    target_sources = config["target_sources"]

    esto_cfg = target_sources["esto"]
    esto_years = [year for year in _year_columns(esto_data) if int(year) <= int(base_year)]
    esto_subset = esto_data[
        (
            (esto_data["economy"].astype(str).str.upper().isin([compact_economy, economy_key.replace("_", "")]))
            | (esto_data["economy_key"] == economy_key)
        )
        & (esto_data["flows"].isin(esto_cfg.get("flows", [])))
    ].copy()
    esto_target_all_economies = esto_data[esto_data["flows"].isin(esto_cfg.get("flows", []))].copy()
    target_exclude_products = {str(value).strip().lower() for value in esto_cfg.get("exclude_products", [])}
    if target_exclude_products:
        esto_subset = esto_subset[
            ~esto_subset["products"].fillna("").astype(str).str.strip().str.lower().isin(target_exclude_products)
        ].copy()
        esto_target_all_economies = esto_target_all_economies[
            ~esto_target_all_economies["products"].fillna("").astype(str).str.strip().str.lower().isin(target_exclude_products)
        ].copy()
    allowed_esto_fuel_keys: set[str] = set()
    if not esto_target_all_economies.empty and esto_years:
        esto_values = esto_target_all_economies[list(esto_years)].apply(pd.to_numeric, errors="coerce").fillna(0.0).abs()
        nonzero_esto_subset = esto_target_all_economies[esto_values.sum(axis=1) > 0.0].copy()
        for product in nonzero_esto_subset["products"].dropna().unique():
            branch_label = _format_fuel_branch_label(
                product,
                source_name="esto",
                fuel_mapping_lookup=mapping_lookup,
            )
            if branch_label:
                allowed_esto_fuel_keys.add(_normalize_balance_label(branch_label))
    for product, group in esto_subset.groupby("products", dropna=False):
        fuel_branch_label = _format_fuel_branch_label(
            product,
            source_name="esto",
            fuel_mapping_lookup=mapping_lookup,
        )
        if _normalize_balance_label(fuel_branch_label) not in allowed_esto_fuel_keys:
            continue
        series = _series_from_group(group, esto_years)
        for year, value in series.items():
            rows.append(
                {
                    "source_dataset": "esto",
                    "economy": economy_key,
                    "process_key": config["process_key"],
                    "process_label": config["process_label"],
                    "leap_process_label": config.get("leap_process_label", config["process_label"]),
                    "fuel_label": str(product),
                    "fuel_branch_label": fuel_branch_label,
                    "year": int(year),
                    "target_energy": float(abs(value)),
                    "target_energy_signed": float(value),
                }
            )

    ninth_cfg = target_sources["ninth"]
    projection_years = [year for year in range(int(base_year) + 1, int(final_year) + 1) if year in ninth_data.columns]
    ninth_subset = ninth_data[ninth_data["economy_key"] == economy_key].copy()
    ninth_subset = _select_ninth_sector(ninth_subset, ninth_cfg.get("sector_codes", []))
    if not ninth_subset.empty:
        ninth_subset = ninth_subset.copy()
        ninth_subset = _drop_ninth_parent_fuel_rows(ninth_subset)
        exclude_fuels = {str(value).strip().lower() for value in ninth_cfg.get("exclude_fuels", [])}
        exclude_subfuels = {str(value).strip().lower() for value in ninth_cfg.get("exclude_subfuels", [])}
        if exclude_fuels:
            ninth_subset = ninth_subset[
                ~ninth_subset["fuels"].fillna("").astype(str).str.strip().str.lower().isin(exclude_fuels)
            ].copy()
        if exclude_subfuels:
            ninth_subset = ninth_subset[
                ~ninth_subset["subfuels"].fillna("").astype(str).str.strip().str.lower().isin(exclude_subfuels)
            ].copy()
        ninth_subset["fuel_label_for_grouping"] = ninth_subset.apply(_target_fuel_label_from_ninth, axis=1)
        ninth_subset["fuel_branch_label_for_grouping"] = ninth_subset["fuel_label_for_grouping"].apply(
            lambda value: _format_fuel_branch_label(
                value,
                source_name="ninth",
                fuel_mapping_lookup=mapping_lookup,
            )
        )
        if allowed_esto_fuel_keys:
            ninth_subset = ninth_subset[
                ninth_subset["fuel_branch_label_for_grouping"].map(_normalize_balance_label).isin(allowed_esto_fuel_keys)
            ].copy()
        else:
            ninth_subset = ninth_subset.iloc[0:0].copy()
        for fuel_branch_label, group in ninth_subset.groupby("fuel_branch_label_for_grouping", dropna=False):
            source_fuels = sorted({str(value) for value in group["fuel_label_for_grouping"].dropna().unique() if str(value).strip()})
            source_fuel_label = "; ".join(source_fuels) if source_fuels else str(fuel_branch_label)
            series = _series_from_group(group, projection_years)
            for year, value in series.items():
                rows.append(
                    {
                        "source_dataset": "ninth",
                        "economy": economy_key,
                        "process_key": config["process_key"],
                        "process_label": config["process_label"],
                        "leap_process_label": config.get("leap_process_label", config["process_label"]),
                        "fuel_label": source_fuel_label,
                        "fuel_branch_label": str(fuel_branch_label),
                        "year": int(year),
                        "target_energy": float(abs(value)),
                        "target_energy_signed": float(value),
                    }
                )
    return pd.DataFrame(rows)


def _nonzero_fuels_from_esto_activity(
    esto_data: pd.DataFrame,
    config: Mapping[str, object],
) -> dict[str, str]:
    esto_cfg = config.get("activity_sources", {}).get("esto", {})
    flows = list(esto_cfg.get("flows", []))
    if not flows or esto_data.empty:
        return {}
    year_cols = _year_columns(esto_data)
    subset = esto_data[esto_data["flows"].isin(flows)].copy()
    product_prefixes = list(esto_cfg.get("product_prefixes", []))
    exact_products = set(esto_cfg.get("include_exact_products", []))
    if product_prefixes or exact_products:
        product_mask = subset["products"].apply(
            lambda value: _has_prefix(value, product_prefixes) or str(value) in exact_products
        )
        subset = subset[product_mask].copy()
    exclude_products = {str(value).strip().lower() for value in esto_cfg.get("exclude_products", [])}
    if exclude_products:
        subset = subset[
            ~subset["products"].fillna("").astype(str).str.strip().str.lower().isin(exclude_products)
        ].copy()
    subset = _apply_value_mode(subset, year_cols, str(esto_cfg.get("value_mode", "signed_sum")))
    if subset.empty or not year_cols:
        return {}
    values = subset[year_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    nonzero = subset.loc[values.abs().gt(0.0).any(axis=1), "products"].dropna().astype(str)
    return {
        str(value).strip(): _source_fuel_label(value)
        for value in nonzero
        if _source_fuel_label(value)
    }


def _nonzero_fuels_from_ninth_activity(
    ninth_data: pd.DataFrame,
    config: Mapping[str, object],
) -> dict[str, str]:
    ninth_cfg = config.get("activity_sources", {}).get("ninth", {})
    sector_codes = list(ninth_cfg.get("sector_codes", []))
    if not sector_codes or ninth_data.empty:
        return {}
    year_cols = _year_columns(ninth_data)
    subset = _select_ninth_sector(ninth_data, sector_codes)
    fuel_values = list(ninth_cfg.get("fuels", []))
    subfuel_values = list(ninth_cfg.get("subfuels", []))
    if fuel_values or subfuel_values:
        subset = _drop_ninth_parent_fuel_rows(subset)
        fuel_mask = subset["fuels"].apply(lambda value: _matches_exact_values(value, fuel_values))
        subfuel_mask = subset["subfuels"].apply(lambda value: _matches_exact_values(value, subfuel_values))
        subset = subset[fuel_mask | subfuel_mask].copy()
    exclude_fuels = {str(value).strip().lower() for value in ninth_cfg.get("exclude_fuels", [])}
    exclude_subfuels = {str(value).strip().lower() for value in ninth_cfg.get("exclude_subfuels", [])}
    if exclude_fuels:
        subset = subset[
            ~subset["fuels"].fillna("").astype(str).str.strip().str.lower().isin(exclude_fuels)
        ].copy()
    if exclude_subfuels:
        subset = subset[
            ~subset["subfuels"].fillna("").astype(str).str.strip().str.lower().isin(exclude_subfuels)
        ].copy()
    subset = _apply_value_mode(subset, year_cols, str(ninth_cfg.get("value_mode", "signed_sum")))
    if subset.empty or not year_cols:
        return {}
    values = subset[year_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    subset = subset.loc[values.abs().gt(0.0).any(axis=1)].copy()
    if subset.empty:
        return {}
    labels = []
    for row in subset.itertuples(index=False):
        subfuel = str(getattr(row, "subfuels", "") or "").strip()
        fuel = str(getattr(row, "fuels", "") or "").strip()
        labels.append(subfuel if subfuel and subfuel != "x" else fuel)
    return {
        str(value).strip(): _source_fuel_label(value)
        for value in labels
        if _source_fuel_label(value)
    }


def _mapped_expected_fuels(
    raw_to_default_label: Mapping[str, str],
    *,
    source_name: str,
    fuel_mapping_lookup: Mapping[str, Mapping[str, str]],
) -> dict[str, dict[str, str]]:
    lookup = fuel_mapping_lookup.get(source_name, {})
    mapped: dict[str, dict[str, str]] = {}
    for raw_label, default_label in raw_to_default_label.items():
        mapped_label = lookup.get(_normalize_source_token(raw_label), default_label)
        if _normalize_balance_label(mapped_label) == "green electricity" and _normalize_balance_label(default_label) == "electricity":
            mapped_label = default_label
        mapped[raw_label] = {
            "source_raw_fuel": raw_label,
            "default_label": default_label,
            "mapped_leap_fuel": mapped_label,
            "mapping_status": "mapped_by_leap_mappings" if mapped_label != default_label else "fallback_label",
        }
    return mapped


def build_proxy_fuel_set_verification(
    *,
    esto_data: pd.DataFrame,
    ninth_data: pd.DataFrame,
    configs: Sequence[Mapping[str, object]],
    fuel_mapping_lookup: Mapping[str, Mapping[str, str]] | None = None,
) -> pd.DataFrame:
    """Compare configured LEAP fuel sets with nonzero ESTO/9th activity fuels."""
    rows: list[dict[str, object]] = []
    for config in configs:
        process_key = str(config.get("process_key", ""))
        process_label = str(config.get("process_label", ""))
        leap_cfg = config.get("activity_sources", {}).get("leap_balance", {})
        fuel_set_name = str(leap_cfg.get("fuel_set", "") or "").strip()
        configured_display = {
            _normalize_balance_label(sanitize_leap_name(str(fuel))): sanitize_leap_name(str(fuel))
            for fuel in LEAP_BALANCE_FUEL_SETS.get(fuel_set_name, [])
            if str(fuel).strip()
        }
        source_sets = {
            "esto": _mapped_expected_fuels(
                _nonzero_fuels_from_esto_activity(esto_data, config),
                source_name="esto",
                fuel_mapping_lookup=fuel_mapping_lookup or {},
            ),
            "ninth": _mapped_expected_fuels(
                _nonzero_fuels_from_ninth_activity(ninth_data, config),
                source_name="ninth",
                fuel_mapping_lookup=fuel_mapping_lookup or {},
            ),
        }
        for source_name, expected in source_sets.items():
            source_role = "source_of_truth" if source_name == "esto" else "supporting"
            expected_display = {
                _normalize_balance_label(info["mapped_leap_fuel"]): info
                for info in expected.values()
                if str(info.get("mapped_leap_fuel", "")).strip()
            }
            expected_keys = set(expected_display)
            configured_keys = set(configured_display)
            for fuel_key in sorted(expected_keys - configured_keys):
                status = (
                    "missing_from_leap_fuel_set"
                    if source_role == "source_of_truth"
                    else "supporting_missing_from_leap_fuel_set"
                )
                rows.append(
                    {
                        "process_key": process_key,
                        "process_label": process_label,
                        "enabled": bool(config.get("enabled", True)),
                        "source": source_name,
                        "source_role": source_role,
                        "fuel_set": fuel_set_name,
                        "fuel_label": expected_display[fuel_key]["mapped_leap_fuel"],
                        "source_raw_fuel": expected_display[fuel_key]["source_raw_fuel"],
                        "default_label": expected_display[fuel_key]["default_label"],
                        "mapping_status": expected_display[fuel_key]["mapping_status"],
                        "status": status,
                    }
                )
            for fuel_key in sorted(configured_keys - expected_keys):
                rows.append(
                    {
                        "process_key": process_key,
                        "process_label": process_label,
                        "enabled": bool(config.get("enabled", True)),
                        "source": source_name,
                        "source_role": source_role,
                        "fuel_set": fuel_set_name,
                        "fuel_label": configured_display[fuel_key],
                        "source_raw_fuel": "",
                        "default_label": "",
                        "mapping_status": "configured_only",
                        "status": "configured_not_seen_in_source_activity",
                    }
                )
            for fuel_key in sorted(expected_keys & configured_keys):
                rows.append(
                    {
                        "process_key": process_key,
                        "process_label": process_label,
                        "enabled": bool(config.get("enabled", True)),
                        "source": source_name,
                        "source_role": source_role,
                        "fuel_set": fuel_set_name,
                        "fuel_label": configured_display[fuel_key],
                        "source_raw_fuel": expected_display[fuel_key]["source_raw_fuel"],
                        "default_label": expected_display[fuel_key]["default_label"],
                        "mapping_status": expected_display[fuel_key]["mapping_status"],
                        "status": "matched",
                    }
                )
    columns = [
        "process_key",
        "process_label",
        "enabled",
        "source",
        "source_role",
        "fuel_set",
        "fuel_label",
        "source_raw_fuel",
        "default_label",
        "mapping_status",
        "status",
    ]
    return pd.DataFrame(rows, columns=columns)


def build_proxy_detail_table(
    *,
    esto_data: pd.DataFrame,
    ninth_data: pd.DataFrame,
    economy: str,
    configs: Sequence[Mapping[str, object]],
    leap_balance_activity: pd.DataFrame | None = None,
    activity_source_mode: str = PROXY_ACTIVITY_SOURCE_MODE,
    base_year: int = EXPORT_BASE_YEAR,
    final_year: int = EXPORT_FINAL_YEAR,
) -> pd.DataFrame:
    detail_frames = []
    fuel_mapping_lookup = load_fuel_mapping_lookup()
    for config in configs:
        if not bool(config.get("enabled", True)):
            continue
        activity = build_activity_series_for_mode(
            esto_data=esto_data,
            ninth_data=ninth_data,
            leap_balance_activity=leap_balance_activity,
            economy=economy,
            config=config,
            base_year=base_year,
            final_year=final_year,
            activity_source_mode=activity_source_mode,
        )
        target = build_target_energy_long(
            esto_data=esto_data,
            ninth_data=ninth_data,
            economy=economy,
            config=config,
            base_year=base_year,
            final_year=final_year,
            fuel_mapping_lookup=fuel_mapping_lookup,
        )
        if target.empty:
            continue
        target["proxy_activity"] = target["year"].map(activity).fillna(0.0)
        target["intensity"] = target.apply(
            lambda row: 0.0 if float(row["proxy_activity"]) == 0.0 else float(row["target_energy"]) / float(row["proxy_activity"]),
            axis=1,
        )
        target["activity_label"] = str(config.get("activity_label", ""))
        target["activity_source_mode"] = str(activity_source_mode)
        target["notes"] = str(config.get("notes", ""))
        detail_frames.append(target)
    if not detail_frames:
        return pd.DataFrame(
            columns=[
                "source_dataset",
                "economy",
                "process_key",
                "process_label",
                "leap_process_label",
                "fuel_label",
                "fuel_branch_label",
                "year",
                "proxy_activity",
                "target_energy",
                "target_energy_signed",
                "intensity",
                "activity_label",
                "activity_source_mode",
                "notes",
            ]
        )
    out = pd.concat(detail_frames, ignore_index=True)
    ordered = [
        "source_dataset",
        "economy",
        "process_key",
        "process_label",
        "leap_process_label",
        "fuel_label",
        "fuel_branch_label",
        "year",
        "proxy_activity",
        "target_energy",
        "target_energy_signed",
        "intensity",
        "activity_label",
        "activity_source_mode",
        "notes",
    ]
    return out[ordered].sort_values(["process_label", "fuel_branch_label", "year"], kind="mergesort")


def build_output_fuel_esto_validation(
    *,
    esto_data: pd.DataFrame,
    detail_df: pd.DataFrame,
    configs: Sequence[Mapping[str, object]],
    economy: str | None = None,
    base_year: int = EXPORT_BASE_YEAR,
    fuel_mapping_lookup: Mapping[str, Mapping[str, str]] | None = None,
    validation_esto_tables: Sequence[tuple[str, pd.DataFrame]] | None = None,
    output_fuel_scope: str = "all_economies",
) -> pd.DataFrame:
    """Check each output fuel branch is non-zero in each ESTO validation snapshot."""
    mapping_lookup = fuel_mapping_lookup or load_fuel_mapping_lookup()
    validation_tables = list(validation_esto_tables or [("esto_data", esto_data)])
    scope = _normalize_output_fuel_scope(output_fuel_scope)
    economy_key = _normalize_economy(economy) if economy else ""
    compact_economy = economy_key.replace("_", "")
    rows: list[dict[str, object]] = []
    if detail_df.empty:
        return pd.DataFrame(
            columns=[
                "process_key",
                "process_label",
                "fuel_branch_label",
                "target_esto_flows",
                "validation_files_checked",
                "validation_years_checked",
                "matched_validation_products",
                "missing_validation_files",
                "output_fuel_scope",
                "validation_economy",
                "status",
            ]
        )
    validation_products_by_flow: dict[str, dict[str, dict[str, object]]] = {}
    for table_name, table in validation_tables:
        year_cols = _year_columns(table)
        final_year = max(year_cols) if year_cols else None
        if final_year is None or table.empty:
            continue
        source = table.copy()
        if scope == "economy":
            if not economy_key:
                raise ValueError("economy is required when output_fuel_scope='economy'.")
            if "economy_key" in source.columns:
                source = source[source["economy_key"].apply(_normalize_economy).eq(economy_key)].copy()
            elif "economy" in source.columns:
                source = source[
                    source["economy"].astype(str).str.upper().isin({economy_key, compact_economy})
                    | source["economy"].apply(_normalize_economy).eq(economy_key)
                ].copy()
        if "is_subtotal" in source.columns:
            subtotal_text = source["is_subtotal"].fillna(False).astype(str).str.strip().str.lower()
            source = source[~subtotal_text.isin({"true", "1", "yes"})].copy()
        values = pd.to_numeric(source[final_year], errors="coerce").fillna(0.0).abs()
        source = source[values > 0.0].copy()
        for flow, flow_group in source.groupby("flows", dropna=False):
            flow_text = str(flow)
            flow_map = validation_products_by_flow.setdefault(flow_text, {})
            for product in flow_group["products"].dropna().unique():
                branch_label = _format_fuel_branch_label(
                    product,
                    source_name="esto",
                    fuel_mapping_lookup=mapping_lookup,
                )
                fuel_key = _normalize_balance_label(branch_label)
                if not fuel_key:
                    continue
                item = flow_map.setdefault(fuel_key, {"products_by_file": {}, "years_by_file": {}})
                item["products_by_file"].setdefault(table_name, set()).add(str(product))
                item["years_by_file"][table_name] = int(final_year)

    for config in configs:
        if not bool(config.get("enabled", True)):
            continue
        process_key = str(config.get("process_key", ""))
        process_label = str(config.get("process_label", ""))
        target_flows = list(config.get("target_sources", {}).get("esto", {}).get("flows", []))
        output_fuels = sorted(
            {
                str(value).strip()
                for value in detail_df.loc[
                    detail_df["process_key"].astype(str).eq(process_key),
                    "fuel_branch_label",
                ].dropna()
                if str(value).strip()
            }
        )
        if not output_fuels:
            continue

        for fuel_label in output_fuels:
            fuel_key = _normalize_balance_label(fuel_label)
            products_by_file: dict[str, set[str]] = {}
            years_by_file: dict[str, int] = {}
            for flow in target_flows:
                match_info = validation_products_by_flow.get(str(flow), {}).get(fuel_key, {})
                for file_name, products in match_info.get("products_by_file", {}).items():
                    products_by_file.setdefault(file_name, set()).update(products)
                for file_name, year in match_info.get("years_by_file", {}).items():
                    years_by_file[file_name] = int(year)
            expected_files = [name for name, _ in validation_tables]
            missing_files = [file_name for file_name in expected_files if not products_by_file.get(file_name)]
            matched_parts = [
                f"{file_name}: {', '.join(sorted(products))}"
                for file_name, products in sorted(products_by_file.items())
            ]
            year_parts = [
                f"{file_name}: {years_by_file[file_name]}"
                for file_name in expected_files
                if file_name in years_by_file
            ]
            rows.append(
                {
                    "process_key": process_key,
                    "process_label": process_label,
                    "fuel_branch_label": fuel_label,
                    "target_esto_flows": "; ".join(str(flow) for flow in target_flows),
                    "validation_files_checked": "; ".join(expected_files),
                    "validation_years_checked": "; ".join(year_parts),
                    "matched_validation_products": " | ".join(matched_parts),
                    "missing_validation_files": "; ".join(missing_files),
                    "output_fuel_scope": scope,
                    "validation_economy": economy_key if scope == "economy" else "all_economies",
                    "status": "matched_all_validation_files" if not missing_files else "missing_from_validation_file",
                }
            )
    columns = [
        "process_key",
        "process_label",
        "fuel_branch_label",
        "target_esto_flows",
        "validation_files_checked",
        "validation_years_checked",
        "matched_validation_products",
        "missing_validation_files",
        "output_fuel_scope",
        "validation_economy",
        "status",
    ]
    return pd.DataFrame(rows, columns=columns).sort_values(["process_label", "fuel_branch_label"], kind="mergesort")


def filter_detail_to_validated_output_fuels(
    detail_df: pd.DataFrame,
    validation_df: pd.DataFrame,
) -> pd.DataFrame:
    """Keep only process/fuel output rows that pass output fuel validation."""
    if detail_df.empty or validation_df.empty:
        return detail_df.copy()
    valid_pairs = validation_df[
        validation_df["status"].eq("matched_all_validation_files")
    ][["process_key", "fuel_branch_label"]].drop_duplicates()
    if valid_pairs.empty:
        return detail_df.iloc[0:0].copy()
    working = detail_df.copy()
    working["_process_key_norm"] = working["process_key"].astype(str)
    working["_fuel_branch_label_norm"] = working["fuel_branch_label"].astype(str)
    valid_pairs = valid_pairs.copy()
    valid_pairs["_process_key_norm"] = valid_pairs["process_key"].astype(str)
    valid_pairs["_fuel_branch_label_norm"] = valid_pairs["fuel_branch_label"].astype(str)
    filtered = working.merge(
        valid_pairs[["_process_key_norm", "_fuel_branch_label_norm"]],
        on=["_process_key_norm", "_fuel_branch_label_norm"],
        how="inner",
    )
    return filtered.drop(columns=["_process_key_norm", "_fuel_branch_label_norm"])


def build_branch_path(parts: Sequence[str]) -> str:
    return "\\".join(sanitize_leap_name(part) for part in parts if str(part or "").strip())


def build_year_rows(
    branch_path: str,
    variable: str,
    scenario: str,
    value_by_year: Mapping[int, float],
    units: str,
    scale: str,
    per_value: str,
) -> list[dict[str, object]]:
    rows = []
    for year, value in sorted(value_by_year.items()):
        rows.append(
            {
                "Branch_Path": branch_path,
                "Scenario": scenario,
                "Measure": variable,
                "Units": units,
                "Scale": scale,
                "Per...": per_value,
                "Date": int(year),
                "Value": float(value),
            }
        )
    return rows


def build_proxy_log_rows(
    detail_df: pd.DataFrame,
    *,
    scenario: str,
    measure_units: Mapping[str, Mapping[str, object]] | None = None,
    include_zero_target_fuel_branches: bool = True,
) -> list[dict[str, object]]:
    if detail_df.empty:
        return []
    units = {key: dict(value) for key, value in DEFAULT_MEASURE_UNITS.items()}
    if measure_units:
        for key, value in measure_units.items():
            units.setdefault(key, {}).update(dict(value))

    rows: list[dict[str, object]] = []
    child_process_labels: list[str] = []
    process_group_col = "leap_process_label" if "leap_process_label" in detail_df.columns else "process_label"
    for process_label, process_group in detail_df.groupby(process_group_col, dropna=False):
        child_process_labels.append(str(process_label))
        for fuel_label, fuel_group in process_group.groupby("fuel_branch_label", dropna=False):
            if not str(fuel_label or "").strip():
                continue
            if not include_zero_target_fuel_branches and not (fuel_group["target_energy"].abs() > 0).any():
                continue
            fuel_path = build_branch_path([*DEMAND_ROOT_PARTS, str(process_label), str(fuel_label)])
            activity_by_year = fuel_group.set_index("year")["proxy_activity"].to_dict()
            rows.extend(
                build_year_rows(
                    fuel_path,
                    ACTIVITY_VARIABLE,
                    scenario,
                    activity_by_year,
                    str(units[ACTIVITY_VARIABLE].get("units", "")),
                    str(units[ACTIVITY_VARIABLE].get("scale", "")),
                    str(units[ACTIVITY_VARIABLE].get("per", "")),
                )
            )
            intensity_by_year = fuel_group.set_index("year")["intensity"].to_dict()
            rows.extend(
                build_year_rows(
                    fuel_path,
                    INTENSITY_VARIABLE,
                    scenario,
                    intensity_by_year,
                    str(units[INTENSITY_VARIABLE].get("units", "")),
                    str(units[INTENSITY_VARIABLE].get("scale", "")),
                    str(units[INTENSITY_VARIABLE].get("per", "")),
                )
            )

    # Write root and process Activity Level rows with Units="No data" and all
    # values = 0. This explicitly clears any stale values in the LEAP template
    # export files without feeding a real activity into the hierarchy (which
    # would cause LEAP to cascade-multiply parent × child × leaf × intensity).
    all_years = sorted({int(y) for y in detail_df["year"].dropna().unique()}) if not detail_df.empty else []
    zero_by_year = {y: 0.0 for y in all_years}
    parent_paths = [build_branch_path(DEMAND_ROOT_PARTS)] + [
        build_branch_path([*DEMAND_ROOT_PARTS, label]) for label in child_process_labels
    ]
    parent_rows: list[dict[str, object]] = []
    for path in parent_paths:
        parent_rows.extend(
            build_year_rows(path, ACTIVITY_VARIABLE, scenario, zero_by_year, "No data", "", "")
        )
    return parent_rows + rows


def build_expression_export_df(export_df: pd.DataFrame, *, base_year: int) -> pd.DataFrame:
    year_cols = sorted([int(col) for col in export_df.columns if str(col).isdigit()])
    out = export_df.copy()
    out["Expression"] = out.apply(
        lambda row: build_data_expression_from_row(row, year_cols),
        axis=1,
    )
    current_mask = out["Scenario"].fillna("").astype(str).str.strip().str.lower().isin({"current accounts", "current account"})
    if current_mask.any():
        out.loc[current_mask, "Expression"] = out.loc[current_mask].apply(
            lambda row: f"Data({int(base_year)},{float(pd.to_numeric(row.get(base_year), errors='coerce') if pd.notna(pd.to_numeric(row.get(base_year), errors='coerce')) else 0.0)})",
            axis=1,
        )
    out = out.drop(columns=year_cols)
    base_cols = ["Branch Path", "Variable", "Scenario", "Region", "Scale", "Units", "Per...", "Expression"]
    level_cols = [col for col in out.columns if str(col).startswith("Level ")]
    return out[base_cols + level_cols]


def load_export_key_table(
    path: Path | str = EXPORT_KEY_WORKBOOK_PATH,
    *,
    sheet_name: str = EXPORT_KEY_WORKBOOK_SHEET,
) -> pd.DataFrame:
    """Load LEAP export rows that provide Branch/Variable/Scenario/Region IDs."""
    resolved = _resolve(path)
    if not resolved.exists():
        raise FileNotFoundError(f"Missing export key workbook: {resolved}")
    df = pd.read_excel(resolved, sheet_name=sheet_name, header=2)
    df = df.rename(columns={col: str(col).strip() for col in df.columns})
    required_cols = [
        "BranchID",
        "VariableID",
        "ScenarioID",
        "RegionID",
        "Branch Path",
        "Variable",
        "Scenario",
        "Region",
    ]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(
            f"Export key workbook {resolved} sheet {sheet_name!r} is missing required columns: {missing_cols}"
        )
    df = df[required_cols].copy()
    for col in ["Branch Path", "Variable", "Scenario", "Region"]:
        df[col] = df[col].fillna("").astype(str).str.strip()
    df = df[df["Branch Path"].ne("") & df["Variable"].ne("") & df["Scenario"].ne("") & df["Region"].ne("")].copy()

    # Scope key rows to the branch/variables managed by this workflow before
    # duplicate validation. Full-model exports can legitimately contain
    # duplicate key tuples in unrelated parts of the model; those should not
    # block Other loss/own-use import generation.
    managed_root = build_branch_path(DEMAND_ROOT_PARTS)
    managed_variables = {ACTIVITY_VARIABLE, INTENSITY_VARIABLE}
    df = df[
        df["Branch Path"].astype(str).str.startswith(managed_root)
        & df["Variable"].isin(managed_variables)
    ].copy()

    duplicate_mask = df.duplicated(["Branch Path", "Variable", "Scenario", "Region"], keep=False)
    if duplicate_mask.any():
        duplicates = df.loc[duplicate_mask, ["Branch Path", "Variable", "Scenario", "Region"]].head(20)
        raise ValueError(
            "Export key workbook has duplicate Branch Path + Variable + Scenario + Region rows. "
            f"First duplicates:\n{duplicates.to_string(index=False)}"
        )
    return df


def merge_export_ids(
    export_df: pd.DataFrame,
    *,
    export_key_table: pd.DataFrame,
) -> pd.DataFrame:
    """Attach LEAP ID columns to generated rows and fail on any missing match."""
    key_cols = ["Branch Path", "Variable", "Scenario", "Region"]
    id_cols = ["BranchID", "VariableID", "ScenarioID", "RegionID"]
    missing_export_cols = [col for col in key_cols if col not in export_df.columns]
    if missing_export_cols:
        raise ValueError(f"Generated export data is missing merge columns: {missing_export_cols}")
    missing_key_cols = [col for col in [*id_cols, *key_cols] if col not in export_key_table.columns]
    if missing_key_cols:
        raise ValueError(f"Export key table is missing required columns: {missing_key_cols}")
    generated = export_df.copy()
    for col in key_cols:
        generated[col] = generated[col].fillna("").astype(str).str.strip()
    key_table = export_key_table[[*key_cols, *id_cols]].copy()
    for col in key_cols:
        key_table[col] = key_table[col].fillna("").astype(str).str.strip()
    merged = generated.merge(
        key_table,
        how="left",
        on=key_cols,
        indicator=True,
    )
    missing = merged[merged["_merge"].ne("both")].copy()
    if not missing.empty:
        preview = missing[key_cols].head(30)
        raise ValueError(
            "Generated rows could not be matched to BranchID/VariableID/ScenarioID/RegionID "
            "from the export key workbook. These branch-variable-scenario-region rows are unexpected:\n"
            f"{preview.to_string(index=False)}"
        )
    merged = merged.drop(columns=["_merge"])
    for col in id_cols:
        merged[col] = pd.to_numeric(merged[col], errors="coerce").astype("Int64")
        if merged[col].isna().any():
            missing_ids = merged[merged[col].isna()][key_cols].head(30)
            raise ValueError(
                f"Generated rows matched the key workbook but have missing {col} values:\n"
                f"{missing_ids.to_string(index=False)}"
            )
    return merged[[*id_cols, *[col for col in merged.columns if col not in id_cols]]]


def _zero_data_expression_for_scenario(
    scenario: object,
    *,
    base_year: int = EXPORT_BASE_YEAR,
    final_year: int = EXPORT_FINAL_YEAR,
) -> str:
    scenario_text = str(scenario or "").strip().lower()
    if scenario_text in {"current accounts", "current account"}:
        return f"Data({int(base_year)},0.0)"
    parts: list[str] = []
    for year in range(int(base_year), int(final_year) + 1):
        parts.extend([str(year), "0.0"])
    return f"Data({', '.join(parts)})"


def add_zero_rows_for_unset_values(
    export_df: pd.DataFrame,
    *,
    export_key_table: pd.DataFrame,
    variables: Sequence[str] = (ACTIVITY_VARIABLE, INTENSITY_VARIABLE),
    demand_root_parts: Sequence[str] = DEMAND_ROOT_PARTS,
    base_year: int = EXPORT_BASE_YEAR,
    final_year: int = EXPORT_FINAL_YEAR,
) -> pd.DataFrame:
    """Add zero expressions for managed rows in the key workbook not otherwise set."""
    key_cols = ["Branch Path", "Variable", "Scenario", "Region"]
    if export_df.empty:
        return export_df.copy()
    out = export_df.copy()
    for col in key_cols:
        out[col] = out[col].fillna("").astype(str).str.strip()
    scenarios = sorted({str(value).strip() for value in out["Scenario"].dropna().unique() if str(value).strip()})
    regions = sorted({str(value).strip() for value in out["Region"].dropna().unique() if str(value).strip()})
    root_path = build_branch_path(demand_root_parts)
    key_table = export_key_table.copy()
    for col in key_cols:
        key_table[col] = key_table[col].fillna("").astype(str).str.strip()
    managed = key_table[
        key_table["Branch Path"].astype(str).str.startswith(root_path)
        & key_table["Variable"].isin(list(variables))
        & key_table["Scenario"].isin(scenarios)
        & key_table["Region"].isin(regions)
    ][key_cols].drop_duplicates()
    already_set = out[key_cols].drop_duplicates()
    missing = managed.merge(already_set, on=key_cols, how="left", indicator=True)
    missing = missing[missing["_merge"].eq("left_only")].drop(columns=["_merge"])
    if missing.empty:
        return out
    template_cols = [col for col in out.columns if col not in key_cols]
    zero_rows = []
    for row in missing.itertuples(index=False):
        item = {
            "Branch Path": row[0],
            "Variable": row[1],
            "Scenario": row[2],
            "Region": row[3],
        }
        for col in template_cols:
            item[col] = pd.NA
        item["Expression"] = _zero_data_expression_for_scenario(
            item["Scenario"],
            base_year=base_year,
            final_year=final_year,
        )
        zero_rows.append(item)
    combined = pd.concat([out, pd.DataFrame(zero_rows)], ignore_index=True, sort=False)
    level_cols = [col for col in combined.columns if str(col).startswith("Level ")]
    if level_cols:
        level_values = combined["Branch Path"].apply(lambda value: str(value).split("\\"))
        for idx, col in enumerate(level_cols):
            combined[col] = combined[col].fillna(level_values.apply(lambda parts: parts[idx] if idx < len(parts) else pd.NA))
    return combined[out.columns]


def _add_leap_header_rows(df: pd.DataFrame, *, model_name: str = EXPORT_MODEL_NAME) -> pd.DataFrame:
    """Return df with the three LEAP-style header rows used by import workbooks."""
    header_data = {col: "" for col in df.columns}
    first_label_col = "Branch Path" if "Branch Path" in df.columns else df.columns[0]
    second_label_col = "Variable" if "Variable" in df.columns else df.columns[min(1, len(df.columns) - 1)]
    third_label_col = "Scenario" if "Scenario" in df.columns else df.columns[min(2, len(df.columns) - 1)]
    fourth_label_col = "Region" if "Region" in df.columns else df.columns[min(3, len(df.columns) - 1)]
    header_data[first_label_col] = "Area:"
    header_data[second_label_col] = model_name
    header_data[third_label_col] = "Ver:"
    header_data[fourth_label_col] = "2"
    header_row_0 = pd.DataFrame([header_data])
    empty_row = pd.DataFrame([{col: pd.NA for col in df.columns}])
    header_row_2 = pd.DataFrame([df.columns], columns=df.columns)
    return pd.concat([header_row_0, empty_row, header_row_2, df], ignore_index=True)


def add_export_id_sheet(
    workbook_path: Path | str,
    export_df: pd.DataFrame,
    *,
    export_key_workbook_path: Path | str = EXPORT_KEY_WORKBOOK_PATH,
    export_key_sheet: str = EXPORT_KEY_WORKBOOK_SHEET,
    output_sheet_name: str = "LEAP_WITH_IDS",
    model_name: str = EXPORT_MODEL_NAME,
    keep_only_id_sheet: bool = False,
    include_zero_rows_for_unset_values: bool = True,
    base_year: int = EXPORT_BASE_YEAR,
    final_year: int = EXPORT_FINAL_YEAR,
) -> pd.DataFrame:
    """Add a workbook sheet containing generated rows merged to LEAP ID columns."""
    key_table = load_export_key_table(export_key_workbook_path, sheet_name=export_key_sheet)
    if include_zero_rows_for_unset_values:
        export_df = add_zero_rows_for_unset_values(
            export_df,
            export_key_table=key_table,
            base_year=base_year,
            final_year=final_year,
        )
    merged = merge_export_ids(export_df, export_key_table=key_table)
    sheet_df = _add_leap_header_rows(merged, model_name=model_name)
    mode = "w" if keep_only_id_sheet else "a"
    writer_kwargs = {"engine": "openpyxl", "mode": mode}
    if not keep_only_id_sheet:
        writer_kwargs["if_sheet_exists"] = "replace"
    with pd.ExcelWriter(_resolve(workbook_path), **writer_kwargs) as writer:
        sheet_df.to_excel(writer, sheet_name=output_sheet_name, index=False, header=False)
    return merged


def _write_csv_with_locked_file_fallback(df: pd.DataFrame, path: Path) -> Path:
    """Write CSV, falling back to *_new.csv when Windows locks an open file."""
    try:
        df.to_csv(path, index=False)
        return path
    except PermissionError:
        fallback = path.with_name(f"{path.stem}_new{path.suffix}")
        df.to_csv(fallback, index=False)
        print(f"[WARN] Could not overwrite locked file {path}; wrote {fallback} instead.")
        return fallback


def _normalize_scenarios(value: str | Sequence[str] | None) -> list[str]:
    if value is None:
        return list(EXPORT_SCENARIOS)
    if isinstance(value, str):
        return [value]
    return [str(item) for item in value]


def _normalize_activity_source_mode(value: str | None) -> str:
    mode = str(value or "").strip().lower()
    if mode in {"first", "first_run", "esto_ninth", "balance_tables"}:
        return "esto_ninth"
    if mode in {"second", "second_run", "leap", "leap_balance", "leap_outputs"}:
        return "leap_balance"
    raise ValueError(
        f"Invalid activity_source_mode={value!r}. "
        "Use 'esto_ninth' for first-run proxies or 'leap_balance' for LEAP-output proxies."
    )


def _normalize_output_fuel_scope(value: str | None) -> str:
    text = str(value or "economy").strip().lower()
    if text in {"economy", "selected_economy", "individual_economy", "current_economy"}:
        return "economy"
    if text in {"all", "all_economies", "global", "model_structure", "structure"}:
        return "all_economies"
    raise ValueError(
        f"Invalid output_fuel_scope={value!r}. Use 'economy' for economy-specific data "
        "or 'all_economies' for model-structure branch scaffolding."
    )


def _resolve_export_filename(economy: str, scenarios: Sequence[str], export_filename: str | None) -> str:
    template = export_filename or EXPORT_FILENAME_TEMPLATE
    return workflow_common.build_workflow_export_filename(
        economy,
        scenarios,
        template,
        workflow_common.format_filename_segment,
        fallback_template="other_loss_own_use_proxy.xlsx",
    )


def assemble_proxy_workbook(
    *,
    economy: str = "20_USA",
    scenarios: str | Sequence[str] | None = None,
    import_scenario: str | Sequence[str] | None = None,
    region: str | None = None,
    export_filename: str | None = None,
    include_leap_import: bool = False,
    configs: Sequence[Mapping[str, object]] = PROXY_CONFIG,
    activity_source_mode: str = PROXY_ACTIVITY_SOURCE_MODE,
    leap_balance_workbook_path: Path | str | None = LEAP_BALANCE_WORKBOOK_PATH,
    leap_balance_scenario: str = LEAP_BALANCE_SCENARIO,
    leap_balance_date_id: str | None = LEAP_BALANCE_DATE_ID,
    output_fuel_scope: str = OUTPUT_FUEL_VALIDATION_SCOPE,
    export_key_workbook_path: Path | str = EXPORT_KEY_WORKBOOK_PATH,
    export_key_sheet: str = EXPORT_KEY_WORKBOOK_SHEET,
    measure_units: Mapping[str, Mapping[str, object]] | None = None,
) -> Path:
    scenario_list = _normalize_scenarios(scenarios)
    output_scope = _normalize_output_fuel_scope(output_fuel_scope)
    region = region or EXPORT_REGION
    esto_data = load_esto_data()
    ninth_data = load_ninth_data()
    leap_balance_activity = pd.DataFrame()
    if str(activity_source_mode or "").strip().lower() in {"second", "second_run", "leap", "leap_balance", "leap_outputs"}:
        balance_path = resolve_leap_balance_workbook_path(
            economy=economy,
            scenario=leap_balance_scenario,
            date_id=leap_balance_date_id,
            workbook_path=leap_balance_workbook_path,
        )
        requested_rows: set[str] = set()
        requested_fuels: set[str] = set()
        for config in configs:
            if not bool(config.get("enabled", True)):
                continue
            leap_cfg = config.get("activity_sources", {}).get("leap_balance", {})
            requested_rows.update(str(row) for row in leap_cfg.get("balance_rows", []) if str(row).strip())
            fuel_set_name = str(leap_cfg.get("fuel_set", "")).strip()
            requested_fuels.update(LEAP_BALANCE_FUEL_SETS.get(fuel_set_name, []))
        leap_balance_activity = load_leap_balance_activity_table(
            balance_path,
            balance_rows=sorted(requested_rows),
            fuels=sorted(requested_fuels),
        )
    detail_df = build_proxy_detail_table(
        esto_data=esto_data,
        ninth_data=ninth_data,
        economy=economy,
        configs=configs,
        leap_balance_activity=leap_balance_activity,
        activity_source_mode=activity_source_mode,
        base_year=EXPORT_BASE_YEAR,
        final_year=EXPORT_FINAL_YEAR,
    )
    if detail_df.empty:
        raise ValueError("No proxy detail rows were generated; check config and source data.")

    output_dir = _resolve(OUTPUT_ROOT) / _normalize_economy(economy)
    output_dir.mkdir(parents=True, exist_ok=True)
    detail_path = output_dir / "proxy_activity_intensity_detail.csv"
    summary_path = output_dir / "proxy_activity_summary.csv"
    leap_activity_path = output_dir / "leap_balance_activity_source.csv"
    fuel_set_verification_path = output_dir / "proxy_fuel_set_verification.csv"
    output_fuel_validation_path = output_dir / "proxy_output_fuel_esto_validation.csv"
    output_fuel_candidate_validation_path = output_dir / "proxy_output_fuel_esto_validation_all_candidates.csv"
    activity_source_warnings_path = output_dir / "proxy_activity_source_warnings.csv"
    activity_source_fallback_path = output_dir / "proxy_activity_source_fallbacks.csv"
    if not leap_balance_activity.empty:
        leap_balance_activity.to_csv(leap_activity_path, index=False)
    fuel_mapping_lookup = load_fuel_mapping_lookup()
    validation_esto_tables = load_output_fuel_validation_esto_tables()
    output_fuel_candidate_validation = build_output_fuel_esto_validation(
        esto_data=esto_data,
        detail_df=detail_df,
        configs=configs,
        economy=economy,
        base_year=EXPORT_BASE_YEAR,
        fuel_mapping_lookup=fuel_mapping_lookup,
        validation_esto_tables=validation_esto_tables,
        output_fuel_scope=output_scope,
    )
    output_fuel_candidate_validation.to_csv(output_fuel_candidate_validation_path, index=False)
    detail_df = filter_detail_to_validated_output_fuels(detail_df, output_fuel_candidate_validation)
    if detail_df.empty:
        raise ValueError(
            "No proxy detail rows remained after output fuel validation. "
            f"See {output_fuel_candidate_validation_path}."
        )
    detail_df.to_csv(detail_path, index=False)
    output_fuel_validation = build_output_fuel_esto_validation(
        esto_data=esto_data,
        detail_df=detail_df,
        configs=configs,
        economy=economy,
        base_year=EXPORT_BASE_YEAR,
        fuel_mapping_lookup=fuel_mapping_lookup,
        validation_esto_tables=validation_esto_tables,
        output_fuel_scope=output_scope,
    )
    output_fuel_validation.to_csv(output_fuel_validation_path, index=False)
    invalid_output_fuels = output_fuel_validation[
        output_fuel_validation["status"].eq("missing_from_validation_file")
    ]
    if not invalid_output_fuels.empty:
        preview = invalid_output_fuels[
            ["process_key", "fuel_branch_label", "missing_validation_files"]
        ].head(20).to_dict("records")
        raise ValueError(
            "Output contains fuel branches that are not non-zero in every ESTO validation snapshot for the matching target flow. "
            f"See {output_fuel_validation_path}. First rows: {preview}"
        )
    activity_source_warnings = pd.DataFrame()
    activity_source_fallbacks = pd.DataFrame()
    if _normalize_activity_source_mode(activity_source_mode) == "esto_ninth":
        activity_source_fallbacks = build_activity_source_fallback_report(
            esto_data=esto_data,
            ninth_data=ninth_data,
            economy=economy,
            configs=configs,
            base_year=EXPORT_BASE_YEAR,
            final_year=EXPORT_FINAL_YEAR,
        )
        activity_source_warnings = build_activity_source_gap_warnings(
            esto_data=esto_data,
            ninth_data=ninth_data,
            economy=economy,
            configs=configs,
            base_year=EXPORT_BASE_YEAR,
            final_year=EXPORT_FINAL_YEAR,
        )
    _write_csv_with_locked_file_fallback(activity_source_fallbacks, activity_source_fallback_path)
    if not activity_source_fallbacks.empty:
        print(
            "\n[INFO] Used broader 9th parent-sector activity for other loss/own-use proxy gaps."
        )
        print(
            f"[INFO] Activity source fallbacks: {len(activity_source_fallbacks)} "
            f"(details saved to {activity_source_fallback_path})"
        )
        for _, row in activity_source_fallbacks.head(20).iterrows():
            print(
                "  - {process}: '{original}' -> '{fallback}'".format(
                    process=str(row.get("process_label") or "").strip(),
                    original=str(row.get("original_ninth_activity_sectors") or "").strip(),
                    fallback=str(row.get("fallback_ninth_activity_sectors") or "").strip(),
                )
            )
        if len(activity_source_fallbacks) > 20:
            print(f"  ... plus {len(activity_source_fallbacks) - 20} more activity source fallback(s)")
    _write_csv_with_locked_file_fallback(activity_source_warnings, activity_source_warnings_path)
    if not activity_source_warnings.empty:
        print(
            "\n[WARN] Found other loss/own-use proxy activity gaps where ESTO activity "
            "is non-zero but all 9th projection activity is zero."
        )
        print(
            f"[WARN] Activity source warnings: {len(activity_source_warnings)} "
            f"(details saved to {activity_source_warnings_path})"
        )
        for _, row in activity_source_warnings.head(20).iterrows():
            print(
                "  - {process}: ESTO total={esto_total:.3f}; 9th sectors='{sectors}'".format(
                    process=str(row.get("process_label") or "").strip(),
                    esto_total=float(row.get("esto_activity_total") or 0.0),
                    sectors=str(row.get("ninth_activity_sectors") or "").strip(),
                )
            )
        if len(activity_source_warnings) > 20:
            print(f"  ... plus {len(activity_source_warnings) - 20} more activity source warning(s)")
    fuel_set_verification = build_proxy_fuel_set_verification(
        esto_data=esto_data,
        ninth_data=ninth_data,
        configs=configs,
        fuel_mapping_lookup=fuel_mapping_lookup,
    )
    _write_csv_with_locked_file_fallback(fuel_set_verification, fuel_set_verification_path)
    (
        detail_df.groupby(["process_key", "process_label", "year"], as_index=False)["proxy_activity"]
        .max()
        .to_csv(summary_path, index=False)
    )

    base_rows = build_proxy_log_rows(
        detail_df,
        scenario=scenario_list[0],
        measure_units=measure_units,
    )
    rows = []
    for scenario_name in scenario_list:
        for row in base_rows:
            copied = row.copy()
            copied["Scenario"] = scenario_name
            rows.append(copied)
    log_df = pd.DataFrame(rows)
    import_scenarios = workflow_common.resolve_import_scenarios(scenario_list, import_scenario)
    scenario_to_import = import_scenarios[0]
    export_df = finalise_export_df(
        log_df,
        scenario=scenario_to_import,
        region=region,
        base_year=EXPORT_BASE_YEAR,
        final_year=EXPORT_FINAL_YEAR,
    )
    if export_df is None or export_df.empty:
        raise ValueError("Export dataframe is empty.")
    expression_df = build_expression_export_df(export_df, base_year=EXPORT_BASE_YEAR)
    output_path = Path(_resolve_export_filename(_normalize_economy(economy), scenario_list, export_filename))
    save_export_files(
        leap_export_df=expression_df,
        export_df_for_viewing=export_df,
        leap_export_filename=output_path,
        base_year=EXPORT_BASE_YEAR,
        final_year=EXPORT_FINAL_YEAR,
        model_name=EXPORT_MODEL_NAME,
    )
    add_export_id_sheet(
        output_path,
        expression_df,
        export_key_workbook_path=export_key_workbook_path,
        export_key_sheet=export_key_sheet,
        output_sheet_name="LEAP_WITH_IDS",
        model_name=EXPORT_MODEL_NAME,
        include_zero_rows_for_unset_values=True,
        base_year=EXPORT_BASE_YEAR,
        final_year=EXPORT_FINAL_YEAR,

    )

    if include_leap_import:
        dispatch_result = dispatch_analysis_input_write(
            export_path=output_path,
            sheet_name="LEAP",
            scenario=scenario_to_import,
            region=region,
            context_label="other_loss_own_use_proxy.include_leap_import",
        )
        if dispatch_result.get("mode") != "workbook":
            if not is_leap_api_available():
                print("[INFO] LEAP API unavailable; skipping LEAP import.")
                return output_path
            from codebase.functions.leap_core import connect_to_leap

            L = connect_to_leap()
            fuel_catalog_preflight.run_fuel_catalog_preflight(
                export_path=output_path,
                sheet_name="LEAP",
                scenario=scenario_to_import,
                context="other_loss_own_use_proxy.include_leap_import",
                leap_app=L,
            )
            create_branches_from_export_file(
                L,
                output_path,
                sheet_name="LEAP",
                scenario=None,
                region=region,
                default_branch_type=None,
                RAISE_ERROR_ON_FAILED_BRANCH_CREATION=True,
            )
            for index, scenario_name in enumerate(import_scenarios):
                fill_branches_from_export_file(
                    L,
                    output_path,
                    sheet_name="LEAP",
                    scenario=scenario_name,
                    region=region,
                    RAISE_ERROR_ON_FAILED_SET=True,
                    HANDLE_CURRENT_ACCOUNTS_TOO=index == 0,
                    RUN_FUEL_CATALOG_PREFLIGHT=False,
                )
    return output_path


#%%
# Notebook-focused settings.
NOTEBOOK_ECONOMY = "20_USA"
NOTEBOOK_SCENARIOS = EXPORT_SCENARIOS
NOTEBOOK_IMPORT_SCENARIOS = [
    scenario.lower()
    for scenario in NOTEBOOK_SCENARIOS
    if scenario.lower() not in {"current accounts", "current account"}
]
NOTEBOOK_INCLUDE_LEAP_IMPORT = False
NOTEBOOK_ACTIVITY_SOURCE_MODE = PROXY_ACTIVITY_SOURCE_MODE
NOTEBOOK_LEAP_BALANCE_WORKBOOK_PATH = LEAP_BALANCE_WORKBOOK_PATH
NOTEBOOK_LEAP_BALANCE_SCENARIO = LEAP_BALANCE_SCENARIO
NOTEBOOK_LEAP_BALANCE_DATE_ID = LEAP_BALANCE_DATE_ID
NOTEBOOK_OUTPUT_FUEL_SCOPE = OUTPUT_FUEL_VALIDATION_SCOPE


def run_with_notebook_config() -> Path:
    return assemble_proxy_workbook(
        economy=NOTEBOOK_ECONOMY,
        scenarios=NOTEBOOK_SCENARIOS,
        import_scenario=NOTEBOOK_IMPORT_SCENARIOS,
        include_leap_import=NOTEBOOK_INCLUDE_LEAP_IMPORT,
        activity_source_mode=NOTEBOOK_ACTIVITY_SOURCE_MODE,
        leap_balance_workbook_path=NOTEBOOK_LEAP_BALANCE_WORKBOOK_PATH,
        leap_balance_scenario=NOTEBOOK_LEAP_BALANCE_SCENARIO,
        leap_balance_date_id=NOTEBOOK_LEAP_BALANCE_DATE_ID,
        output_fuel_scope=NOTEBOOK_OUTPUT_FUEL_SCOPE,
    )


if __name__ == "__main__":
    run_with_notebook_config()
#%%


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
#%%
