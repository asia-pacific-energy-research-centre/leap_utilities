# AGENTS.md

These are project-level instructions for Codex (and similar agents).

## When editing draw.io diagrams

- See `AGENTS_DRAWIO.md` for draw.io-specific requirements.

## Small guide for humans

- Put instructions here that you want Codex to follow every time it edits this repo.
- Keep rules short and specific; avoid large, complex policies.
- For file-specific rules, include path globs like `docs/leap-system*.drawio`.
- When updating transfer category mappings, re-run `leap_utils/scrapbook/transfers_mapping_exploration.py`
  and paste the printed `TRANSFER_PROCESS_CONFIG` into `leap_utils/transfers_workflow.py`.

## LEAP Export File Structure

- See `C:\\Users\\Work\\.codex\\AGENTS_LEAP_EXPORT.md` for LEAP export structure requirements.

## Balance Table Structures (ESTO vs 9th)

- See `C:\\Users\\Work\\.codex\\AGENTS_BALANCE_TABLES.md` for balance table structure details.

## Balance Table Structures (ESTO vs 9th)

These two balance tables are the core inputs for `leap_utils/transformation_analysis_workflow.py`.
Keep this structure in mind when adding new transformations or debugging data issues.

### 9th structure (sector/fuel hierarchy)

- Source file: `data/merged_file_energy_ALL_20250814.csv` (loaded as "9th" in the script).
- Key columns:
  - `scenarios`, `economy`
  - Sector hierarchy: `sectors`, `sub1sectors`, `sub2sectors`, `sub3sectors`, `sub4sectors`
  - Fuel hierarchy: `fuels`, `subfuels`
  - Subtotal flags: `subtotal_layout`, `subtotal_results`
  - Year columns (as strings before normalization): `1980` ... `2070`
- Coding style:
  - Codes use underscores, e.g., `09_06_gas_processing_plants`, `10_01_03_liquefaction_regasification_plants`.
  - `"x"` means "not used" for a given hierarchy level.
- Usage in transformations:
  - Supports detailed subsector selection (e.g., LNG uses `sub2sectors` and `subfuels`).
  - Filtered to `scenarios == reference` before calculations.
- Subtotals are removed using the subtotal mapping in `config/ESTO_subtotal_mapping.xlsx`.

### ESTO (Matt) structure (flow/product table)

- Source file: `data/00APEC_2024_low.csv` (loaded as "ESTO (Matt) data" in the script).
- Key columns:
  - `economy`
  - `flows` (balance rows like production, transformation, own use, losses)
  - `products` (fuel/product codes)
  - Year columns: `1990` ... `2022`
- Coding style:
  - Economy codes are compact (e.g., `01AUS`), normalized to `01_AUS` to align with 9th.
  - Flow codes match the 09/10 transformation and loss lists (e.g., `09.08.01 Coke ovens`, `10.01.05 Coke ovens`).
- Usage in transformations:
  - Used for most transformation flows when sector detail is not required.
  - No `sub*sectors` columns are present, so selection is done via `flows` and `products`.

### Shared sign conventions (both tables)

- Positive values represent outputs from a transformation flow.
- Negative values represent inputs to a transformation flow (feedstock or auxiliary fuels).
- Loss/own-use flows are treated as auxiliary fuel use (absolute values are used in ratios).
