# LEAP utilities import/export workflow

## Overview

- **Data sources**: `transformation_analysis_utils.py` (analytics) and the `transformation_entry.py` entrypoint (via `transformation_workflow.py`) read the cleansed 9th/ESTO tables in `data/merged_file_energy_ALL_20250814_pre_trump.csv` and `data/00APEC_2024_low.csv` (updated to 2025 when that data becomes available), plus the subtotal mapping workbook `config/ESTO_subtotal_mapping.xlsx`. For exact 9th edition projection matching, switch to `data/merged_file_energy_ALL_20251106.csv` and `data/merged_file_energy_00_APEC_20251106`.
- **Code/label mappings**: Optional name lookups and subtotal helpers live under `config/`, including `config/sector_fuel_codes_to_names*.xlsx` and the subtotal mapping workbook.
- **Purpose**: The workflows derive LEAP parameters (outputs, feedstocks, auxiliaries, losses, efficiencies, imports/exports) under `Transformation` and `Resources/Primary`, then map the Excel exports back into LEAP to update branch trees.
- **Execution order**: Run the transformation extractor before (or alongside) the supply extractor. Each script writes its own XLSX under `outputs/leap_exports/`. Use the matching mapping script afterwards to push the data into LEAP (scenario/economy chunks are embedded in file names and sheet metadata).


## Typical workflow & modeling connotations

1. **Pick the sectors you want to fill**: Use the workflow scripts that correspond to the LEAP sectors you are populating. Common choices:
   - `codebase/transformation_workflow.py` for Transformation inputs/outputs, losses, efficiencies.
   - `codebase/power_workflow.py` for direct import of power model exports (`data/power export.xlsx`) with scenario alignment and audit reporting.
   - `codebase/supply_workflow.py` for Resources → Primary imports/exports.
   - `codebase/transfers_workflow.py` for Transfers flows.
   - `codebase/minor_demand_workflow.py` for Minor demand branches.
   You can run only the workflows you need; each produces its own LEAP-ready workbook.
2. **Derive numbers**: Each workflow calls its underlying analysis helpers (e.g., `transformation_analysis_utils.py`, `supply_data_pipeline.py`) to compute the values that should appear in LEAP.
3. **Verify**: Inspect printed summaries (for example the transformation LEAP structure block or fuel/flow summaries) to confirm the right fuels, losses, and totals are selected. Use summary CSVs if the workflow exposes a save toggle.
4. **Export**: Every workflow writes an XLSX in LEAP log format (`Branch Path`, `Scenario`, `Measure`, etc.), with scenario/economy context encoded in the file name.
5. **Import into LEAP**: Use the matching workflow import helper (for example `transformation_workflow.import_transformation_workbook_to_leap()`) so LEAP creates the branch skeleton and fills measures. Make sure scenario and economy names in the file match the LEAP scenario keys and economy keys you intend to update.
6. **Modeling impact**: After import, LEAP sees:
   - `Transformation` branches where `Process Efficiency`, `Feedstock Fuel Share`, and `Auxiliary Fuel Use` are rooted under each combustion/transformation technology, allowing dispatch rules to know the true outputs and losses.
   - `Resources → Primary` branches showing the imported/exported volumes and unmet demand shares for each fuel, which can be used by downstream demand modules.
   - Additional sector branches from transfers/minor demand workflows that fill out non-transformation parts of the model when needed.

The other modules in this repo mostly support these workflow files (shared constants, data loaders, mapping utilities, and LEAP import/export helpers).

## transformation_analysis_utils.py and transformation_workflow.py

### Intent and data pipeline

- Loads the 9th (reference-focused) and ESTO (Matt) datasets, normalizes `1980`–`2070` year columns to integers, drops subtotals, and adds the synthetic `ALL` economy when `INCLUDE_ALL_ECONOMIES` is `True`.
- `MAJOR_SECTOR_CONFIG` declares each transformation flow: dataset key (`"ninth"` for LNG, `"esto"` for others), flow codes, loss references, and navigation hints (subsector codes, titles).
- `CODE_TO_NAME_MAPPING` (if enabled) uses `config/sector_fuel_codes_to_names*.xlsx` to show human-readable sector/fuel names in logs and exports.
- Note that while we wait for Matthew to update ESTO's data to include LNG/Nautral gas splits for OECD economies, we will use the 9th dataset for LNG liquefaction/regasification analysis and ESTO for gas processing and coal/charcoal/nonspecified solid fuel transformations. The goal is to eventually consolidate everything into ESTO once those flows are available there.

### Processing steps

1. For each sector in `ANALYSIS_REGISTRY`, `run_analysis_for_sector` grabs the right dataset and runs the sector-specific analyser (`analyze_lng_liquefaction_regas`, `analyze_gas_processing`, or the flow-based `summarize_transformation_flows`).
2. `summarize_transformation_flows` isolates positive (`outputs`) and negative (`feedstocks`) fuel rows, pulls loss/own-use data via `build_loss_context`, and builds per-economy/year series that feed into:
   - `compute_efficiency_by_year` (output / (feedstock + losses))
   - `build_auxiliary_ratios_by_year` (per-fuel auxiliary shares)
   - `build_process_record` (aggregation of output, feedstock, auxiliary, loss, import/export targets, and shares)
3. A `PROCESS_RECORDS` list gathers every process; `save_transformation_summaries` optionally writes `transformation_process_summary.csv` / `transformation_detail_summary.csv` for diagnostics.
4. `save_transformation_export` runs `build_transformation_log_rows`, `finalise_export_df`, `build_expression_export_df`, and `save_export_files` so the XLSX aligns with LEAP’s log export format. The default file name is `transformation_leap_imports_{economy}_{scenario}.xlsx`.

### Output shape and LEAP modeling impact

- The export fills the `Transformation` branch tree. Typical rows include:
  - `Output Fuels` entries with values, import/export targets, and `Units=Petajoule`/`Gigajoule`.
  - Processes under each sector with `Process Efficiency`, `Feedstock Fuel Share`, and `Auxiliary Fuel Use`. Auxiliary rows inherit `DEFAULT_AUXILIARY_UNITS` (`Gigajoule`) and `Per...=Gigajoule`.
  - `Dispatch Rule`/`Process Share` rows that feed `fill_branches_from_export_file` into a Demand-technology tree (the new `transformation_entry.py` entrypoint controls when LEAP gets called).
- The script prints `print_leap_structure_block` per flow so you can sanity-check fuel splits.
- `TRANSFORMATION_OUTPUT_VARIABLES` controls which series go into the log (outputs, import/export targets, feedstock shares, efficiencies, auxiliary ratios, loss totals).

### Customisation knobs

- Toggle analyses (LNG, gas works, coal, charcoal, nonspecified) via `RUN_*` constants near the bottom.
- Control exported scenarios with `SCENARIOS_TO_EXPORT` and override specific years in `SCENARIO_EXPORT_OVERRIDES`.
- Use `INCLUDE_ALL_FEEDSTOCKS_AS_AUXILIARY` to treat every non-primary feedstock as an auxiliary fuel so LEAP sees them under `Auxiliary Fuels`.
- Update `EXPORT_MODEL_NAME`, `EXPORT_REGION`, or `EXPORT_OUTPUT_DIR` to match a different LEAP project.

### transformation_entry.py

- This file is called by `transformation_workflow.py` during the normal run. It is the standard place that starts the XLSX export and, if you choose, the LEAP import.
- Use this file when you want the usual flow without extra setup.
- Older functions (`prepare_transformation_exports`, `run_transformation_pipeline`, `run_transformation_leap_import`) are kept for older notebooks; they do the same work by calling `transformation_workflow.py`.

## supply_workflow.py

- Acts as the user-facing entrypoint for the supply export/import workflow. `run_supply_export_and_import` (and the `SUPPLY_RUN_LEAP_IMPORT` environment toggle) first calls `assemble_supply_workbooks()` to regenerate the XLSX from the ESTO/9th tables, then optionally invokes `supply_data_pipeline.run_supply_leap_import()` with the configured scenario to push those values into LEAP.
- Legacy names (`quick_supply_export`, `run_supply_pipeline`) remain available for older scripts.
- Use `SUPPLY_IMPORT_SCENARIO` to override the LEAP scenario when you need to inject the export into something other than the default (`Target`). If the flag/variable is absent, the helper still writes the workbook but leaves the LEAP step for you to run later (you can re-run the helper later with the env var turned on or call `supply_data_pipeline.run_supply_leap_import` directly from a notebook).

## power_workflow.py

- Standalone import workflow for power-sector LEAP export files, currently configured for `data/power export.xlsx` (`Export` sheet).
- Copies the source workbook into `outputs/leap_exports/power_export_prepared_{economy}_{scenario}.xlsx` before any edits, so the source export is unchanged.
- Aligns scenarios by duplicating `Optimization` rows into `Reference` and `Target` while preserving `Current Accounts`.
- Validates fuel labels under `Output Fuels` / `Feedstock Fuels` / `Auxiliary Fuels` against cleaned ESTO products (`data/00APEC_2024_low.csv`) and writes `intermediate_data/power_fuel_validation_report.csv`.
- Supports explicit `SKIP_VARIABLES` for control/optimization variables and writes a consolidated fill audit report to `intermediate_data/power_fill_audit_report.csv`.
- Supports inline hardcoded post-fill overrides (`HARDCODED_VARIABLE_OVERRIDES`) and writes status rows to `intermediate_data/power_hardcoded_values_report.csv`.
- Branch creation/fill uses the same `create_branches_from_export_file` and `fill_branches_from_export_file` flow as the refining/industry mapping scripts, including stale-child checks during fill.

## supply_data_pipeline.py

### Scope

- Focuses on `Resources` section within LEAP structure. This requires imports/exports to be captured per fuel.
- Builds a dynamic supply fuel config (not `MAJOR_SECTOR_CONFIG`) by reading the `ESTO` sheet of `sector_fuel_codes_to_names*.xlsx`. Every ESTO product becomes a `Resources → Primary` branch unless excluded by `EXCLUDED_ESTO_PREFIXES`.
- Defines a secondary-product classification (`SECONDARY_ESTO_PRODUCT_PREFIXES/EXACT`), but this classification is not currently used to include/exclude fuels in the supply export; it is available if you want to add that filter later.
- Uses `FLOW_CODES_BY_DATASET` to keep the 9th and ESTO flow labels (`01 Production`, `02 Imports`, etc.) in sync with whichever dataset key is active (default is `"esto"`).
- Generates three LEAP measures (`Imports`, `Exports`, `Unmet Requirements`) for each fuel in `SUPPLY_MEASURES`; the last is a placeholder with `value=0` and `per=MeetWithImports`.

### Processing flow

1. Loads the 9th/ESTO tables, applies the subtotal mapping, optionally saves the subtotal-labeled ESTO file, and filters to the Reference scenario.
2. Adds the `ALL` economy rows if `INCLUDE_ALL_ECONOMIES` is `True`.
3. `build_supply_log_rows` loops through every configured fuel, uses `get_flow_total_for_fuel` to sum base-year import/export values, and coerces each scalar into a per-year dict via `coerce_value_by_year` so the log matches LEAP’s yearly format. Export totals are normalized with `normalize_supply_flow_total` so LEAP always sees positive exports even if the source balance records a negative value.
4. Logs are finalised with `finalise_export_df`, then persisted via `save_export_files` to `outputs/leap_exports/supply_leap_imports_{economy}_{scenarios}.xlsx`. Scenario list defaults to `["Current Accounts", "Reference", "Target"]`. You can override the export folder during testing or reruns with `SUPPLY_LEAP_EXPORT_DIR=/tmp/custom_dir`.

### Modeling implications

- The generated file describes `Resources → Primary → [Fuel]` branches. Each branch gets:
  - `Imports` and `Exports` measured in `Gigajoule`.
  - `Unmet Requirements` as a `Percent` with `Per...=MeetWithImports`, which gives LEAP a placeholder for unmet demand.
- Because the script only uses the base year, each branch resets year columns to that value (other years are copies via `coerce_value_by_year`), so every branch reflects the single-year view typical for source flows.
- You can use `RUN_LIST_FUELS` to print unique fuel/product combinations for debugging or to discover new codes.

## transfers_workflow.py

### Scope

- Treats ESTO `08.*` Transfers flows as transformation-style processes so they can be exported with the same LEAP structure as the transformation workflows.
- Uses per-economy mappings in `TRANSFER_PROCESS_CONFIG` (and `TRANSFER_CATEGORY_TEMPLATES` as a fallback) to group inputs/outputs; prefers subflows (`08.01`–`08.03`) when they have data.
- Drops subtotals before any transfer logic runs, and reuses `transformation_analysis_utils.py` helpers to build process records and export workbooks.

### Processing flow

1. Pulls the ESTO reference data from `transformation_analysis_utils.py`, filters transfer flow rows, and applies the per-economy process configuration.
2. Builds process records (inputs/outputs, optional output targets) using the same sign conventions as the transformation pipeline (negative = input, positive = output).
3. Merges duplicate process rows, optionally consolidates output series/targets, and saves the workbook as `outputs/leap_exports/transfer_leap_imports_{economy}_{scenario}.xlsx`.
4. Optionally imports into LEAP via `run_transfer_export_and_import` / `import_transfer_workbook_to_leap`, which uses the standard branch-creation helpers.

### Modeling implications

- Creates `Transfers` processes under the Transformation-style branch structure, which keeps transfer activity visible without editing the core transformation modules.
- If you change `TRANSFER_CATEGORY_TEMPLATES`, you can rerun `codebase/scrapbook/transfers_mapping_exploration.py` and paste the printed `TRANSFER_PROCESS_CONFIG` back into `codebase/transfers_workflow.py` (per `AGENTS.md`). Make manual changes only when you understand the mapping logic. There are utilities within `codebase/scrapbook/transfers_mapping_exploration.py` that help with identifying flows and working out appropriate groupings.

## minor_demand_workflow.py

### Scope

- Builds a small demand tree under `Demand → Other sector` for minor sectors (Agriculture, Fishing, Non-specified others) using ESTO flows and 9th projections.
- Uses `config/ninth_pairs_to_esto_pairs.xlsx` to allocate 9th projections down to ESTO flow/product pairs (same allocation logic as `ninth_projection_mapping.py`).
- Exports a LEAP workbook that mirrors the `data/industry export.xlsx` schema, with an expression-based sheet for LEAP import.

### Processing flow

1. Loads ESTO and 9th datasets, drops subtotals, and filters 9th to the projection scenario.
2. Filters the 9th↔ESTO mapping to the minor-demand flows in `MINOR_DEMAND_FLOW_CONFIG`.
3. Allocates 9th projections to ESTO pairs (`build_esto_projection_table`) and builds per-fuel Activity Level rows plus Final Energy Intensity rows.
4. Writes `outputs/leap_exports/minor_demand_export_{economy}_{scenario}.xlsx` and (optionally) creates/fills LEAP branches via `create_branches_from_export_file` and `fill_branches_from_export_file`.

### Modeling implications

- Activity Level is projection-driven; Final Energy Intensity is a placeholder by default (`INTENSITY_MODE="uniform"` and `DEFAULT_INTENSITY=1.0`), so total energy can exceed sector totals when multiple fuels are present.
- Switch `INTENSITY_MODE` to `fuel_share` or `custom` if you need intensities that sum to 1 or match calibrated values.
