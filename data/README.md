# Data Folder Guide

This folder holds model inputs, manually exported LEAP workbooks, reference
tables, and local caches used by the workflow scripts. Most generated workflow
outputs should go under `outputs/`, not here.

## Main Reference Tables

These CSVs are the common historical/projection data sources used across
mapping, dashboard, demand, supply, and transformation workflows.

### ESTO Historical Tables

- `00APEC_2024_low.csv`
  - Historical ESTO-style balance data used by older supply, transformation,
    industry, buildings, power, refining, and minor-demand workflows.
  - Key columns are `economy`, `flows`, `products`, and year columns such as
    `1990` through the latest base year in the file.

- `00APEC_2024_low_with_subtotals.csv`
  - Same 2024 ESTO source with subtotal labels added.
  - Used where workflows need to identify subtotal rows explicitly, especially
    transfer, detailed-balance, and older mapping checks.

- `00APEC_2025_low.csv`
  - Newer ESTO-style historical table.
  - Keep when comparing behavior across 2024 vs 2025 data vintages.

- `00APEC_2025_low_with_subtotals.csv`
  - Current preferred ESTO historical source for dashboard and balance-table
    comparison workflows.
  - Used by `codebase/leap_results_dashboard*_workflow.py`,
    `codebase/leap_balance_to_esto_long_workflow.py`,
    `codebase/leap_results_workflow.py`, and balance-demand logic in
    `codebase/results_supply_link_workflow.py`.

### 9th Projection Tables

- `merged_file_energy_ALL_20250814_pre_trump.csv`
  - Older 9th projection input used by several established transformation,
    supply, industry, and minor-demand workflows.
  - Keep this file because some workflow defaults still point to it.

- `merged_file_energy_ALL_20251106.csv`
  - Current preferred 9th projection table for exact 9th edition matching.
  - Used by the dashboard, balance-table, mapping-refresh, and
    `results_supply_link` balance-demand paths.

- `merged_file_energy_00_APEC_20251106.csv`
  - APEC aggregate version of the current 9th projection data.
  - Used by mapping and comparison preparation scripts that need aggregate
    projection rows.

- `merged_file_energy_ALL_20251106 - for chatgpt.csv`
  - Review/export copy for external inspection.
  - Do not treat it as the workflow source of truth unless a script is changed
    to point to it explicitly.

## LEAP Import Template Workbooks

These are workbook-shaped inputs that mirror LEAP Analysis-view import/export
structure. They are used as templates or reference schemas when building manual
LEAP import workbooks.

- `full model export.xlsx`
  - Main full-model Analysis-view reference workbook.
  - Used to align branch paths, variables, scenarios, regions, workbook IDs,
    and field mappings for combined/full-model exports.
  - `results_supply_link_workflow.py` also uses it for verification and for
    supply root classification.

- `industry export.xlsx`
  - LEAP import/export template for industry demand branches.
  - Used by `codebase/industry_workflow.py` and as the template for
    minor-demand workflows.

- `buildings export.xlsx`, `dummy buildings export.xlsx`,
  `buildings_dummy_20_USA todo add fuels to buildings export then import as banches into leap.xlsx`
  - Buildings-sector templates and working variants.
  - Used by `codebase/buildings_workflow.py` and
    `codebase/buildings_dummy_workflow.py`.

- `power export.xlsx`
  - Power-sector import workbook used by `codebase/power_workflow.py`.

- `refining model export.xlsx`
  - Refining import workbook used by `codebase/refining_workflow.py`.

- `detailed balance table output example.xlsx`
  - Template/example workbook for detailed balance-table generation.

## LEAP Results Inputs

### `leap balances exports/`

Manual Energy Balance exports from LEAP. These are now the main source for
balance-demand extraction and dashboard-independent LEAP balance tables.

See `leap balances exports/README.md` for filename rules and extraction
details. In short, workflows read workbooks like:

```text
leap balances exports/20_USA/full model output all years 04092026 REF.xlsx
leap balances exports/20_USA/full model output all years 04092026 TGT.xlsx
```

The extractor converts balance sheets into long rows, converts values to
petajoules, then maps LEAP sector/fuel pairs to ESTO flow/product pairs using
`config/leap_mappings.xlsx`.

### `leap results tables/`

Rendered LEAP Results-view workbook templates and refreshed outputs. These were
the older source for dashboard/result workflows and are still used by
`codebase/leap_results_workflow.py`, old extraction probes, and some comparison
utilities.

Typical active files are:

```text
leap results tables/transformation_results_20_USA_Reference.xlsx
leap results tables/transformation_results_20_USA_Target.xlsx
leap results tables/supply_results_20_USA_Reference.xlsx
leap results tables/supply_results_20_USA_Target.xlsx
leap results tables/industry_results_20_USA_Reference.xlsx
leap results tables/industry_results_20_USA_Target.xlsx
leap results tables/buildings_results_20_USA_Reference.xlsx
leap results tables/buildings_results_20_USA_Target.xlsx
```

Files under `leap results tables/processed tables/` are derived helper tables
for dashboards, such as transformation auxiliary own-use and derived metrics.

## Other Reference Inputs

- `Data for comparison  - APERC outlooks .xlsx`
  - External comparison workbook used by older APERC reference aggregation and
    mapping preparation scripts.

- `usa proejcted simplifeid.csv`
  - Older simplified USA projection artifact.
  - Treat as reference/scratch unless a workflow explicitly points to it.

- `population/`
  - World Population Prospects 2024 files.
  - Used as reference data for workflows or checks that need population
    indicators. These are external input files, not generated outputs.

## Cache, Archive, and Scratch Areas

- `.cache/`
  - Local pandas cache files for expensive reference-table loads.
  - Safe to regenerate. Do not edit manually.

- `archive/`
  - Old source files, backups, and damaged/corrupted workbook copies kept for
    provenance.
  - Workflows should not normally read from here unless explicitly configured.

- `temp/`
  - Scratch mapping and unmapped-label artifacts.
  - Safe to clean only after confirming no active mapping task depends on the
    files.

## Editing Rules

- Prefer adding new generated artifacts under `outputs/`, not `data/`.
- Keep canonical input filenames stable unless you also update every workflow
  constant that references them.
- When replacing a canonical CSV or workbook, archive the old copy first.
- Keep files that are manually exported from LEAP in the matching LEAP folder:
  Energy Balance exports go in `leap balances exports/`; Results-view exports
  go in `leap results tables/`.
- For current balance/dashboard work, default to
  `00APEC_2025_low_with_subtotals.csv` and
  `merged_file_energy_ALL_20251106.csv` unless the workflow explicitly requires
  an older data vintage.
