# LEAP utilities import/export workflow

## Overview

- **Data sources**: `transformation_analysis_utils.py` (analytics) and the lean `transformation_entry.py` entrypoint (which in turn uses `transformation_workflow.py`) read the cleansed 9th/ESTO tables in `data/merged_file_energy_ALL_20250814.csv` and `data/00APEC_2024_low.csv`, plus the subtotal mapping workbook `config/ESTO_subtotal_mapping.xlsx`. The transformation workflow also relies on optional code-to-name tables in `config/sector_fuel_codes_to_names*.xlsx`.
- **Purpose**: The two workflow modules derive the parameters (`output` fuels, feedstocks, auxiliaries, losses, efficiencies, imports/exports) that the LEAP model expects under `Transformation` and `Resources/Primary`. The complementary mapping scripts translate the generated Excel exports back into LEAP so the new branch tree reflects those parameters.
- **Execution order**: Run the transformation extractor before (or alongside) the supply extractor. Each script produces its own XLSX under `outputs/leap_exports/`. Use the matching mapping script afterwards to push the data into LEAP (chunks of scenarios/economies are embedded in the file names and sheet metadata).

## transformation_analysis_utils.py and transformation_workflow.py

### Intent and data pipeline

- Loads both the 9th (reference-focused) and ESTO (Matt) datasets, normalizes the `1980`–`2070` columns to integers, drops subtotal rows, and adds the synthetic `ALL` economy when `INCLUDE_ALL_ECONOMIES` is `True`.
- `MAJOR_SECTOR_CONFIG` declares every transformation flow of interest: dataset key (`"ninth"` for LNG, `"esto"` for others), explicit transformation flow codes, loss references, and navigation hints (subsector codes, titles).
- `CODE_TO_NAME_MAPPING` (if enabled) uses `config/sector_fuel_codes_to_names*.xlsx` to show human-readable sector/fuel names in logs and exports.

### Processing steps

1. For each sector in `ANALYSIS_REGISTRY`, `run_analysis_for_sector` grabs the right dataset and runs the sector-specific analyser (`analyze_lng_liquefaction_regas`, `analyze_gas_processing`, or the flow-based `summarize_transformation_flows`).
2. `summarize_transformation_flows` isolates positive (`outputs`) and negative (`feedstocks`) fuel rows, pulls loss/own-use data via `build_loss_context`, and builds per-economy/year series that feed into:
   - `compute_efficiency_by_year` (output / (feedstock + losses))
   - `build_auxiliary_ratios_by_year` (per-fuel auxiliary shares)
   - `build_process_record` (aggregation of output, feedstock, auxiliary, loss, import/export targets, and shares)
3. A `PROCESS_RECORDS` list gathers every process; `save_transformation_summaries` optionally writes `transformation_process_summary.csv` / `transformation_detail_summary.csv` for diagnostics.
4. `save_transformation_export` runs `build_transformation_log_rows`, `finalise_export_df`, `build_expression_export_df`, and `save_export_files` so the XLSX aligns with LEAP’s log export format. The default file name is `transformation_leap_imports_{economy}_{scenario}.xlsx`.

### Output shape and LEAP modeling impact

- The created export fills the `Transformation` branch tree. Typical rows include:
  - `Output Fuels` entries with values, import/export targets, and `Units=Petajoule`/`Gigajoule`.
  - Processes under each sector with `Process Efficiency`, `Feedstock Fuel Share`, and `Auxiliary Fuel Use`. Auxiliary rows inherit `DEFAULT_AUXILIARY_UNITS` (`Gigajoule`) and `Per...=Gigajoule`.
  - `Dispatch Rule`/`Process Share` rows that feed `fill_branches_from_export_file` into a Demand-technology tree (the new `transformation_entry.py` entrypoint controls when LEAP gets called).
- The script prints `print_leap_structure_block` for each flow so you can eyeball fuel splits before trusting the export file.
- `TRANSFORMATION_OUTPUT_VARIABLES` controls which series go into the log (outputs, import/export targets, feedstock shares, efficiencies, auxiliary ratios, loss totals), enabling quick toggling.

### Customisation knobs

- Toggle analyses (LNG, gas works, coal, charcoal, nonspecified) via `RUN_*` constants near the bottom.
- Control exported scenarios with `SCENARIOS_TO_EXPORT` and override specific years in `SCENARIO_EXPORT_OVERRIDES`.
- Use `INCLUDE_ALL_FEEDSTOCKS_AS_AUXILIARY` to treat every non-primary feedstock as an auxiliary fuel so LEAP sees them under `Auxiliary Fuels`.
- Update `EXPORT_MODEL_NAME`, `EXPORT_REGION`, or `EXPORT_OUTPUT_DIR` to match a different LEAP project.

### transformation_entry.py

- The new user-facing entrypoint that mirrors `supply_workflow.py`: it calls `transformation_workflow.prepare_transformation_exports()` to create the XLSX and — when `include_leap_import=True` — invokes `transformation_workflow.run_transformation_leap_import()` so LEAP receives the same data. A tiny set of constants near the bottom allows quick toggling of economies, scenarios, and the import behavior.
- All heavy logic (export filename formatting, process record collection, scenario/location validation, branch creation/fill) lives in `transformation_workflow.py`, which in turn wraps `transformation_analysis_utils.py` plus the LEAP API helpers. That module can still be imported if you need finer control, but typical users just call `run_with_notebook_config()` from `transformation_entry.py`.

## supply_data_pipeline.py

### Scope

- Focuses on `Resources → Primary` so imports/exports can be captured per fuel for `Base Year = 2022`.
- Builds a dynamic `MAJOR_SECTOR_CONFIG` by reading the `ESTO` sheet of `sector_fuel_codes_to_names*.xlsx`, classifying secondary products (`SECONDARY_ESTO_PRODUCT_PREFIXES/EXACT`) so transformation leftovers can be excluded if desired.
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

## supply_workflow.py

- Acts as the user-facing entrypoint for the supply export/import workflow. `run_supply_pipeline` (and the `SUPPLY_RUN_LEAP_IMPORT` environment toggle) first calls `quick_supply_export()` to regenerate the XLSX from the ESTO/9th tables, then optionally invokes `supply_data_pipeline.run_supply_leap_import()` with the configured scenario to push those values into LEAP.
- Use `SUPPLY_IMPORT_SCENARIO` to override the LEAP scenario when you need to inject the export into something other than the default (`Target`). If the flag/variable is absent, the helper still writes the workbook but leaves the LEAP step for you to run later (you can re-run the helper later with the env var turned on or call `supply_data_pipeline.run_supply_leap_import` directly from a notebook).

## Typical workflow & modeling connotations

1. **Derive numbers**: Run `transformation_analysis_workflow.py` to compute transformation efficiencies, feedstock/auxiliary shares, and import/export targets; run `supply_data_pipeline.py` to get imports/exports tied to the same base data.
2. **Verify**: Inspect the printed `LEAP structure block` (transformation) or fuel/flow summaries to confirm the right fuels and losses are selected. Use the summary CSVs if `SAVE_SUMMARY_TABLES` is enabled.
3. **Export**: Each script writes an XLSX that is in the LEAP log format (`Branch Path`, `Scenario`, `Measure`, etc.).
4. **Import into LEAP**: Run `transformation_entry.py` (or directly `transformation_workflow.run_transformation_leap_import()` when you already have the export file) and `supply_workflow.py` (with `SUPPLY_RUN_LEAP_IMPORT=1`) to create the branch skeleton and fill the measures. Ensure the scenario names in the file match the scenario keys you intend to update in LEAP.
5. **Modeling impact**: After import, LEAP sees:
   - `Transformation` branches where `Process Efficiency`, `Feedstock Fuel Share`, and `Auxiliary Fuel Use` are rooted under each combustion/transformation technology, allowing dispatch rules to know the true outputs and losses.
   - `Resources → Primary` branches showing the imported/exported volumes and unmet demand shares for each fuel, which can be used by downstream demand modules.

By keeping the constants at the bottom of each script in sync with the LEAP project identifiers (`EXPORT_REGION`, `EXPORT_MODEL_NAME`, scenario labels), you can re-run the full process whenever the ESTO/9th data refreshes. Review `AGENTS.md` for any additional repo-specific practices when tweaking these scripts.
