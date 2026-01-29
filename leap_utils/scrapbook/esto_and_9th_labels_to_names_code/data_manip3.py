#%%
# Summary: Clean the sector_fuel_codes_to_names mapping and export to a new CSV.
from __future__ import annotations

from pathlib import Path
from typing import List, Tuple

import pandas as pd
#%%

#%%
######### CONSTANTS (UNLIKELY TO CHANGE) #########
ENABLE_DEBUG_BREAKPOINTS = True
REQUIRED_COLUMNS = {"9th_label", "9th_column", "esto_label", "esto_column", "name"}
#%%

#%%
######### FUNCTIONS #########
def try_debug_breakpoint() -> None:
    """Trigger a debug breakpoint when enabled (safe to call anywhere).

    Inputs:
        None.
    Outputs:
        None.
    Side effects:
        May pause execution in a debugger.
    """
    if not ENABLE_DEBUG_BREAKPOINTS:
        return
    try:
        breakpoint()
    except Exception as breakpoint_exc:
        print(f"Debug breakpoint failed: {breakpoint_exc}")


def load_mapping_sheet(workbook_path: Path, sheet_name: str) -> pd.DataFrame:
    """Load the mapping sheet from the workbook.

    Inputs:
        workbook_path: Path to the Excel workbook.
        sheet_name: Sheet name to read.
    Outputs:
        Dataframe with the sheet contents.
    Side effects:
        Reads from disk.
    """
    try:
        df = pd.read_excel(workbook_path, sheet_name=sheet_name, dtype=str).fillna("")
        missing = REQUIRED_COLUMNS - set(df.columns)
        if missing:
            raise ValueError(f"Missing required columns: {sorted(missing)}")
        return df
    except Exception as exc:
        print(f"Failed to read mapping sheet {sheet_name} from {workbook_path}: {exc}")
        try_debug_breakpoint()
        raise


def fix_known_mismatches(df: pd.DataFrame) -> pd.DataFrame:
    """Remove incorrect ESTO mappings for known 9th labels.

    Inputs:
        df: Mapping dataframe.
    Outputs:
        Updated dataframe with mismatch fixes applied.
    Side effects:
        None.
    """
    try:
        df_out = df.copy()
        df_out.loc[df_out["9th_label"] == "09_x_04_biomass", ["esto_label", "esto_column"]] = ""
        df_out.loc[df_out["9th_label"] == "09_13_06_others", ["esto_label", "esto_column"]] = ""
        return df_out
    except Exception as exc:
        print(f"Failed to apply mismatch fixes: {exc}")
        try_debug_breakpoint()
        raise


def count_segments(label: str) -> int:
    """Count number of underscore segments (more detail = higher count).

    Inputs:
        label: Label string.
    Outputs:
        Integer segment count (or -1 for blank labels).
    Side effects:
        None.
    """
    if pd.isna(label) or label == "":
        return -1
    return label.count("_")


def dedupe_esto_labels(df: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    """Deduplicate ESTO labels, preferring more detailed 9th labels.

    Inputs:
        df: Mapping dataframe.
    Outputs:
        Tuple of (updated dataframe, removed row count).
    Side effects:
        None.
    """
    try:
        df_out = df.copy()
        duplicate_mask = df_out["esto_label"].duplicated(keep=False) & (df_out["esto_label"] != "")
        duplicates = df_out[duplicate_mask].copy()

        if duplicates.empty:
            return df_out, 0

        rows_to_drop: List[int] = []
        for esto_label in duplicates["esto_label"].unique():
            group = df_out[df_out["esto_label"] == esto_label].copy()

            group["detail_score"] = group["9th_label"].apply(count_segments)
            group["has_9th"] = group["9th_label"].apply(
                lambda x: 0 if pd.isna(x) or x == "" else 1
            )
            group = group.sort_values(["has_9th", "detail_score"], ascending=[False, False])

            rows_to_drop.extend(group.index[1:].tolist())

        df_out = df_out.drop(rows_to_drop)
        return df_out, len(rows_to_drop)
    except Exception as exc:
        print(f"Failed to deduplicate ESTO labels: {exc}")
        try_debug_breakpoint()
        raise


def add_missing_labels(
    df: pd.DataFrame,
    missing_9th: List[Tuple[str, str]],
    missing_esto: List[Tuple[str, str]],
) -> Tuple[pd.DataFrame, int, int]:
    """Add missing 9th and ESTO labels to the mapping.

    Inputs:
        df: Mapping dataframe.
        missing_9th: List of (label, column) tuples to add for 9th.
        missing_esto: List of (label, column) tuples to add for ESTO.
    Outputs:
        Tuple of (updated dataframe, count_9th_added, count_esto_added).
    Side effects:
        None.
    """
    try:
        df_out = df.copy()
        name_map_9th = dict(zip(df_out["9th_label"], df_out["name"]))
        name_map_esto = dict(zip(df_out["esto_label"], df_out["name"]))

        new_rows_9th = []
        for label, column in missing_9th:
            if label not in df_out["9th_label"].values:
                new_rows_9th.append(
                    {
                        "9th_label": label,
                        "9th_column": column,
                        "esto_label": "",
                        "esto_column": "",
                        "name": name_map_9th.get(label, ""),
                    }
                )

        new_rows_esto = []
        for label, column in missing_esto:
            if label not in df_out["esto_label"].values:
                new_rows_esto.append(
                    {
                        "9th_label": "",
                        "9th_column": "",
                        "esto_label": label,
                        "esto_column": column,
                        "name": name_map_esto.get(label, ""),
                    }
                )

        df_out = pd.concat([df_out, pd.DataFrame(new_rows_9th), pd.DataFrame(new_rows_esto)], ignore_index=True)
        return df_out, len(new_rows_9th), len(new_rows_esto)
    except Exception as exc:
        print(f"Failed to add missing labels: {exc}")
        try_debug_breakpoint()
        raise


def save_output_csv(df: pd.DataFrame, output_path: Path) -> None:
    """Save the cleaned mapping to a CSV.

    Inputs:
        df: Mapping dataframe.
        output_path: CSV output path.
    Outputs:
        None.
    Side effects:
        Writes to disk.
    """
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_csv(output_path, index=False)
        print(f"Cleaned mapping saved to: {output_path}")
        print(f"Total rows: {len(df)}")
    except Exception as exc:
        print(f"Failed to write CSV to {output_path}: {exc}")
        try_debug_breakpoint()
        raise
#%%

#%%
######### CONSTANTS (LIKELY TO CHANGE) #########
WORKBOOK_PATH = Path(r"C:\Users\Work\github\leap_utilities\config\sector_fuel_codes_to_names.xlsx")
SOURCE_SHEET_NAME = "code_to_name"
OUTPUT_CSV_PATH = Path(r"C:\Users\Work\github\leap_utilities\outputs\sector_fuel_codes_to_names_cleaned.csv")

MISSING_9TH = [
    ("10_01_02_gas_works_plants", "sub2sectors"),
    ("10_01_03_liquefaction_regasification_plants", "sub2sectors"),
    ("10_01_04_gastoliquids_plants", "sub2sectors"),
    ("10_01_05_coke_ovens", "sub2sectors"),
    ("10_01_07_blast_furnaces", "sub2sectors"),
    ("10_01_09_bkb_pb_plants", "sub2sectors"),
    ("10_01_10_liquefaction_plants_coal_to_oil", "sub2sectors"),
    ("10_01_11_oil_refineries", "sub2sectors"),
    ("14_03_02_01_fs", "sub3sectors"),
    ("15_02_01_02_01_diesel_engine", "sub4sectors"),
    ("15_02_01_02_02_gasoline_engine", "sub4sectors"),
    ("15_02_01_02_03_battery_ev", "sub4sectors"),
    ("15_02_01_02_04_compressed_natual_gas", "sub4sectors"),
    ("15_02_01_02_05_plugin_hybrid_ev_gasoline", "sub4sectors"),
    ("15_02_01_02_06_plugin_hybrid_ev_diesel", "sub4sectors"),
    ("15_02_01_03_01_diesel_engine", "sub4sectors"),
    ("15_02_01_03_02_gasoline_engine", "sub4sectors"),
    ("15_02_01_03_03_battery_ev", "sub4sectors"),
    ("15_02_01_03_04_compressed_natual_gas", "sub4sectors"),
    ("15_02_01_03_05_plugin_hybrid_ev_gasoline", "sub4sectors"),
    ("15_02_01_03_06_plugin_hybrid_ev_diesel", "sub4sectors"),
    ("15_02_01_04_01_diesel_engine", "sub4sectors"),
    ("15_02_01_04_02_gasoline_engine", "sub4sectors"),
    ("15_02_01_04_03_battery_ev", "sub4sectors"),
    ("15_02_01_04_04_compressed_natual_gas", "sub4sectors"),
    ("15_02_01_04_05_plugin_hybrid_ev_gasoline", "sub4sectors"),
    ("15_02_01_04_06_plugin_hybrid_ev_diesel", "sub4sectors"),
    ("15_02_01_passenger", "sub2sectors"),
    ("15_02_02_01_01_diesel_engine", "sub4sectors"),
    ("15_02_02_01_02_gasoline_engine", "sub4sectors"),
    ("15_02_02_01_03_battery_ev", "sub4sectors"),
    ("15_02_02_01_04_compressed_natual_gas", "sub4sectors"),
    ("15_02_02_01_05_plugin_hybrid_ev_gasoline", "sub4sectors"),
    ("15_02_02_01_06_plugin_hybrid_ev_diesel", "sub4sectors"),
    ("15_02_02_01_two_wheeler", "sub3sectors"),
    ("15_02_02_02_01_diesel_engine", "sub4sectors"),
    ("15_02_02_02_02_gasoline_engine", "sub4sectors"),
    ("15_02_02_02_03_battery_ev", "sub4sectors"),
    ("15_02_02_02_04_compressed_natual_gas", "sub4sectors"),
    ("15_02_02_02_05_plugin_hybrid_ev_gasoline", "sub4sectors"),
    ("15_02_02_02_06_plugin_hybrid_ev_diesel", "sub4sectors"),
    ("15_02_02_02_light_vehicle", "sub3sectors"),
    ("15_02_02_03_01_diesel_engine", "sub4sectors"),
    ("15_02_02_03_02_gasoline_engine", "sub4sectors"),
    ("15_02_02_03_03_battery_ev", "sub4sectors"),
    ("15_02_02_03_04_compressed_natual_gas", "sub4sectors"),
    ("15_02_02_03_05_plugin_hybrid_ev_gasoline", "sub4sectors"),
    ("15_02_02_03_06_plugin_hybrid_ev_diesel", "sub4sectors"),
    ("15_02_02_03_light_truck", "sub3sectors"),
    ("15_02_02_04_01_diesel_engine", "sub4sectors"),
    ("15_02_02_04_02_gasoline_engine", "sub4sectors"),
    ("15_02_02_04_03_battery_ev", "sub4sectors"),
    ("15_02_02_04_04_compressed_natual_gas", "sub4sectors"),
    ("15_02_02_04_05_plugin_hybrid_ev_gasoline", "sub4sectors"),
    ("15_02_02_04_06_plugin_hybrid_ev_diesel", "sub4sectors"),
    ("15_02_02_04_heavy_truck", "sub3sectors"),
    ("15_02_02_freight", "sub2sectors"),
    ("15_03_01_passenger", "sub2sectors"),
    ("15_03_02_freight", "sub2sectors"),
    ("15_04_01_passenger", "sub2sectors"),
    ("15_04_02_freight", "sub2sectors"),
    ("15_05_other_biomass", "subfuels"),
    ("16_others", "sectors"),
    ("17_02_industry_sector", "sub1sectors"),
    ("17_03_transport_sector", "sub1sectors"),
    ("17_04_other_sector", "sub1sectors"),
    ("18_01_01_01_subcritical", "sub3sectors"),
    ("18_01_01_02_superultracritical", "sub3sectors"),
    ("18_01_01_03_advultracritical", "sub3sectors"),
    ("18_01_01_04_ccs", "sub3sectors"),
    ("18_01_01_coal_power", "sub2sectors"),
    ("18_01_02_01_gasturbine", "sub3sectors"),
    ("18_01_02_02_combinedcycle", "sub3sectors"),
    ("18_01_02_03_ccs", "sub3sectors"),
    ("18_01_02_gas_power", "sub2sectors"),
    ("18_01_03_oil", "sub2sectors"),
    ("18_01_04_nuclear", "sub2sectors"),
    ("18_01_05_01_large", "sub3sectors"),
    ("18_01_05_02_mediumsmall", "sub3sectors"),
    ("18_01_05_03_pump", "sub3sectors"),
    ("18_01_05_hydro", "sub2sectors"),
    ("18_01_06_biomass", "sub2sectors"),
    ("18_01_07_geothermal", "sub2sectors"),
    ("18_01_08_01_utility", "sub3sectors"),
    ("18_01_08_02_rooftop", "sub3sectors"),
    ("18_01_08_03_csp", "sub3sectors"),
    ("18_01_08_solar", "sub2sectors"),
    ("18_01_09_01_onshore", "sub3sectors"),
    ("18_01_09_02_offshore", "sub3sectors"),
    ("18_01_09_wind", "sub2sectors"),
    ("18_01_10_otherrenewable", "sub2sectors"),
    ("18_01_11_otherfuel", "sub2sectors"),
    ("18_01_12_storage", "sub2sectors"),
    ("18_01_electricity_plants", "sub1sectors"),
    ("18_02_01_coal", "sub2sectors"),
    ("18_02_02_gas", "sub2sectors"),
    ("18_02_03_oil", "sub2sectors"),
    ("18_02_04_biomass", "sub2sectors"),
    ("18_02_chp_plants", "sub1sectors"),
    ("19_01_chp_plants", "sub1sectors"),
    ("22_total_combustion_emissions", "sectors"),
]

MISSING_ESTO = [
    ("01 Coal", "products"),
    ("02 Coal products", "products"),
    ("03 Peat", "products"),
    ("04 Peat products", "products"),
    ("05 Oil shale and oil sands", "products"),
    ("06 Crude oil & NGL", "products"),
    ("07 Petroleum products", "products"),
    ("07.01 Aviation gasoline", "products"),
    ("08 Gas", "products"),
    ("08.01 Natural gas", "products"),
    ("08.02 LNG", "products"),
    ("08.03 Gas works gas", "products"),
    ("08.99 Non-specified Gas", "products"),
    ("09 Nuclear", "products"),
    ("09.02.01 Electricity plants", "flows"),
    ("09.02.02 CHP plants", "flows"),
    ("09.02.03 Heat plants", "flows"),
    ("10 Hydro", "products"),
    ("10.01.02 Gas works plants", "flows"),
    ("10.01.03 Liquefaction/regasification plants", "flows"),
    ("10.01.04 Gas-to-liquids plants", "flows"),
    ("10.01.05 Coke ovens", "flows"),
    ("10.01.07 Blast furnaces", "flows"),
    ("10.01.08 Patent fuel plants", "flows"),
    ("10.01.09 BKB/PB plants", "flows"),
    ("10.01.11 Oil refineries", "flows"),
    ("11 Geothermal", "products"),
    ("12 Solar", "products"),
    ("13 Tide, wave, ocean", "products"),
    ("14 Wind", "products"),
    ("15 Solid biomass", "products"),
    ("15.01 Fuelwood & woodwaste", "products"),
    ("15.02 Bagasse", "products"),
    ("15.03 Charcoal", "products"),
    ("15.04 Black liqour", "products"),
    ("15.05 Other biomass", "products"),
    ("16 Others", "products"),
    ("16.01 Biogas", "products"),
    ("16.02 Industrial waste", "products"),
    ("16.03 Municipal solid waste (renewable)", "products"),
    ("16.04 Municipal solid waste (non-renewable)", "products"),
    ("16.05 Biogasoline", "products"),
    ("17.02 Industry sector", "flows"),
    ("17.03 Transport sector", "flows"),
    ("17.04 Other sector", "flows"),
    ("18 Heat", "products"),
    ("19 Total", "products"),
    ("19.01 MAP CHP plants", "flows"),
    ("19.03 AP CHP plants", "flows"),
]

RUN_CLEANING = True
#%%

#%%
######### RUN CLEANING (TOGGLE) #########
if RUN_CLEANING:
    try:
        source_df = load_mapping_sheet(WORKBOOK_PATH, SOURCE_SHEET_NAME)
        fixed_df = fix_known_mismatches(source_df)
        fixed_df, removed_count = dedupe_esto_labels(fixed_df)
        fixed_df, added_9th_count, added_esto_count = add_missing_labels(
            fixed_df, MISSING_9TH, MISSING_ESTO
        )

        print(f"Removed {removed_count} duplicate rows")
        print(f"Added {added_9th_count} missing 9th labels")
        print(f"Added {added_esto_count} missing ESTO labels")

        save_output_csv(fixed_df, OUTPUT_CSV_PATH)
    except Exception as exc:
        print(f"Failed to clean mapping: {exc}")
        try_debug_breakpoint()
#%%

