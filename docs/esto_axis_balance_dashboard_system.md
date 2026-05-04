# ESTO-Axis Balance Dashboard System

This document explains how the ESTO-axis balance dashboard works from a first-use and systems perspective. The dashboard compares three sources on the same ESTO balance-table axis:

- LEAP balance exports, mapped from LEAP sector/fuel rows to ESTO flow/product rows.
- ESTO base-year balance data, read directly by ESTO flow/product.
- 9th edition projection data, read by 9th sector/fuel pairs that are linked back through the LEAP mapping lineage.

The main workflow is `codebase/leap_results_dashboard_balance_estoaxis_workflow.py`.

## Where the Code Lives

- `codebase/leap_results_dashboard_balance_estoaxis_workflow.py`
  The notebook-safe entrypoint. It defines paths, years, stage toggles, visible series, and the ordered run workflow.

- `codebase/utilities/leap_results_dashboard_balance.py`
  The main implementation module for ESTO-axis balance extraction, mapping, comparison building, rendering, lineage audits, and coverage checks.

- `codebase/utilities/leap_results_dashboard_utils.py`
  Shared dashboard helpers for loading/pulling ESTO and 9th series, preparing render data, and writing chart files.

- `codebase/utilities/leap_results_dashboard_v2/*`
  Shared comparison diagnostics, output writing, and chart ledger helpers reused by the ESTO-axis dashboard.

The workflow file should stay readable and high level. Most data logic belongs in `leap_results_dashboard_balance.py`.

## Main Inputs

- `config/leap_comparison_dashboard_template_v2.json`
  Controls the dashboard page structure, chart groups, requested ESTO flows, and requested products. It is the dashboard authoring file.

- `config/leap_mappings.xlsx`
  The primary mapping workbook. The ESTO-axis dashboard uses:
  - `leap_combined_esto`: LEAP sector/fuel to ESTO flow/product.
  - `leap_combined_ninth`: LEAP sector/fuel to 9th sector/fuel.

- `config/master_config.xlsx`, sheet `ninth_pairs_to_esto_pairs`
  The direct canonical ESTO-to-9th mapping table. This is only used when a dashboard template section explicitly enables `use_esto_to_ninth_mapping`.

- LEAP balance export workbooks
  Resolved by `resolve_balance_export_workbook()` in `codebase/utilities/leap_balance_export_resolver.py`.

- ESTO base table
  Usually `data/00APEC_2025_low_with_subtotals.csv`.

- 9th projection table
  Usually `data/merged_file_energy_ALL_20251106.csv`.

## High-Level Flow

1. The workflow reads the dashboard template and builds an ESTO-axis structure.
2. The LEAP balance export workbooks are extracted into long rows.
3. LEAP rows are mapped to ESTO flow/product pairs using `leap_combined_esto`.
4. The mapped LEAP rows are grouped by scenario, year, ESTO flow, and ESTO product.
5. A comparison table is built with LEAP, ESTO base-year, and 9th projection rows.
6. Transformation rows are split into input and output measures for rendering.
7. Charts and dashboard pages are rendered.
8. Audit files and ledgers are written so charted values can be traced.

## Workflow Controls

The main workflow has stage flags near the top of `codebase/leap_results_dashboard_balance_estoaxis_workflow.py`:

- `STAGE_EXTRACT`
  Reads REF/TGT LEAP balance workbooks and maps them to ESTO rows. This is usually the slowest stage.

- `STAGE_COMPARE`
  Builds the LEAP/ESTO/9th comparison tables.

- `STAGE_WRITE_OUTPUTS`
  Writes comparison tables, simple mapped balance tables, and diagnostics.

- `STAGE_RENDER_DASHBOARDS`
  Renders HTML chart files and dashboard pages.

- `STAGE_WRITE_COVERAGE`
  Writes coverage, runtime issue, and mapping check outputs.

Other important controls:

- `BASE_YEAR`, `MAX_OUTPUT_YEAR`, and `PROJECTION_YEARS`
  Define the base/projection year range. The current setup uses base year 2022 and projects through 2060.

- `SCENARIO_MAP`
  Maps dashboard scenario names to 9th table scenario codes, for example `Target -> target`.

- `VISIBLE_COMPARISON_SERIES`
  Filters which source/scenario series are written and rendered. The workflow may build more data internally than it displays.

- `FAIL_ON_UNMAPPED_BALANCE_ROWS`
  Controls whether unresolved mapped-balance issues should fail the workflow. The workflow still writes outputs before raising/reporting this issue.

## Mapping Source of Truth

The workflow relies first on the LEAP mapping lineage:

```text
ESTO flow/product <- LEAP sector/fuel -> 9th sector/fuel
```

This lineage is built from active rows in `leap_combined_esto` and `leap_combined_ninth`.

Rows are active when:

- `remove_row` is not true.
- `duplicate_to_remove` is not true, if that column exists.
- The LEAP sector path, LEAP fuel, and target pair columns are populated.

The direct `ninth_pairs_to_esto_pairs` table is not the default source for charted 9th rows. It is a fallback for template sections that opt in with `use_esto_to_ninth_mapping`.

The mapping workbook is also used as an audit artifact. The columns `pair_mapping_cardinality`, `subtotal_alignment`, `many_to_many_is_ok`, `remove_row`, `remove_row_reason`, and `Note` are part of the mapping decision record, not just decoration.

## Why the Full Mapping Crosswalk Exists

Earlier versions effectively used only LEAP balance rows that survived into `mapping_status` to decide which 9th rows to query. That meant a valid 9th projection row could be missing if the corresponding LEAP export row was zero or absent.

The current workflow builds a wider active crosswalk from the mapping workbook before projection rows are created. This means:

- A 9th row can appear when it is mapped and nonzero, even if the LEAP row is zero or absent.
- LEAP rows still use the same LEAP-to-ESTO mapping rules as before.
- The dashboard target universe is no longer restricted to nonzero LEAP rows.

To avoid flooding the dashboard with structural zero mappings, the expanded crosswalk is filtered to 9th sector/fuel pairs that are nonzero in the requested projection economy, scenarios, and years.

## How LEAP Rows Are Built

The LEAP balance export is extracted by `TemplateBalanceExtractor`.

For each balance row:

1. The LEAP sector path and fuel name are normalized.
2. The row is matched to `leap_combined_esto`.
3. Rows without a required ESTO flow/product are recorded as runtime mapping issues.
4. Mapped rows are grouped by:
   - scenario
   - year
   - ESTO flow
   - ESTO product
5. Values are summed.
6. The grouped rows become the LEAP comparator series.

The mapping extraction keeps a pre-group mapped detail table as well as grouped dashboard rows. Use the pre-group table when you need to understand which original LEAP balance rows contributed to a grouped ESTO pair.

The grouped output is written to:

- `outputs/dashboards/leap_results_dashboard_balance_estoaxis/USA/leap_long.csv`
- `outputs/balance_tables/leap_balance_to_esto_long/USA/supporting_files/leap_balance_mapped_detail_long.csv`

## How 9th Rows Are Built

The 9th comparator uses the ESTO-to-LEAP-to-9th target map.

The target map is seeded from:

1. `mapping_status`, which represents LEAP-backed mapped rows.
2. The active workbook crosswalk from `leap_combined_esto` joined to `leap_combined_ninth`.
3. Direct canonical ESTO-to-9th mappings, only where the template explicitly opts in.

The 9th table is pre-filtered by economy and scenario, then prepared once so repeated sector/fuel lookups are fast.

The expanded workbook crosswalk is filtered to nonzero 9th pairs for the requested projection economy, scenarios, and years. This keeps the target universe complete enough to catch nonzero projection rows, but avoids creating chart rows for every structural zero in the mapping workbook.

For each chart group, the renderer:

1. Gets all 9th sector/fuel targets for the ESTO flow/product.
2. Pulls the projection series for those targets.
3. Sums target series for the chart row.
4. Records component rows for audit.

If the same 9th sector/fuel pair is claimed more than once in the same ESTO flow group and scenario, the renderer claims it once to reduce double counting. LEAP-backed chart rows get priority over workbook-only template rows when claiming shared pairs.

This priority matters when one 9th pair maps to several ESTO products. A row that exists in the LEAP-backed `mapping_status` claims the shared 9th pair before a template-only product row.

## How ESTO Rows Are Built

ESTO base-year values are pulled directly from the ESTO base table using:

- `economy`
- `flows`
- `products`
- base year column

The ESTO value is one value per ESTO flow/product/year. ESTO subtotal rows are avoided for normal chart rows unless a subtotal row is explicitly part of the chart structure.

ESTO values are base-year only in the dashboard. Projection years come from LEAP and 9th; ESTO is used as the historical/base comparator.

## Transformation Input and Output Splitting

Transformation charts are split for readability:

- Negative transformation values are inputs.
- Positive transformation values are outputs.
- Input charts display negative values as positive magnitudes.
- Output charts display positive values.

This split can happen after the main comparison table is built, so `comparison_long.csv` is not always enough to understand rendered chart semantics. For rendered chart debugging, use:

- `supporting_files/charting/chart_line_mapping_ledger.csv`
- `supporting_files/charting/chart_total_component_ledger.csv`

For example, a positive `17 Electricity` transformation value belongs in an output chart. It should not be interpreted as a nonzero input just because the unsplit comparison table still has a generic balance measure.

## Subtotal Handling

Subtotal rows are not globally removed from every stage.

Current behavior:

- LEAP subtotal flags are carried through from the mapping workbook and inferred from obvious total/subtotal naming.
- 9th projection lookups prefer non-subtotal rows when detailed rows exist.
- If an ESTO pair maps to both always-subtotal and non-subtotal 9th pairs, the always-subtotal targets are dropped.
- If a row is marked as a 9th subtotal row, projection output for that row is blanked in rendering.

The goal is to avoid summing a subtotal row and its children in the same chart.

## Many-to-Many Handling

The mapping workbook can contain intentional one-to-many, many-to-one, and many-to-many relationships.

Important points:

- The workflow does not automatically allocate one source value across many targets.
- A LEAP source row mapped to multiple ESTO pairs can contribute to multiple ESTO rows.
- The renderer has projection-side duplicate protection so one 9th pair is not counted multiple times within the same ESTO flow group and scenario.
- Many-to-many mappings should be marked intentionally in the workbook with `many_to_many_is_ok` and explained in `Note`.

The safest way to audit many-to-many behavior is to inspect:

- `mapping_lineage_audit.csv`
- `chart_line_mapping_ledger.csv`
- `dashboard_comparator_pair_coverage.xlsx`
- `leap_mapping_duplicate_mappings.csv`

Do not assume `many_to_many_is_ok=True` means the workflow has allocated values. It means the mapping author accepted the relationship. The rendered projection path still has duplicate-claim protection, but LEAP and ESTO values remain tied to their mapped ESTO rows.

## Dashboard Template Behavior

The dashboard template controls what the user sees.

Common graph specifications:

- `aggregate_graphs`
  Creates total charts for one or more ESTO flows.

- `by_fuel_graphs`
  Creates fuel/product charts for selected ESTO flows.

- `products: "All"`
  Lets the renderer include products that are visible from LEAP, ESTO, or mapped nonzero 9th projection data.

- `use_esto_to_ninth_mapping`
  Allows direct canonical ESTO-to-9th fallback mapping for that specific template section. Use this carefully because it can include canonical pairs that are not intended for LEAP-lineage comparison.

- `about_page`
  Adds a human-readable About page to the rendered dashboard. This is for high-level orientation, not detailed system documentation.

The template is both a navigation file and a chart allowlist. If a page or chart group is absent from the template, the renderer should not invent it except for explicitly supported fallback/empty-page behavior.

## Rendered Dashboard Versus Data Tables

The workflow writes several data tables before rendering charts. The dashboard can further transform those rows for display:

- It splits transformation balances into Inputs and Outputs.
- It converts input values to positive magnitudes for input charts.
- It can collapse multi-flow template rows, such as sections that combine main activity and autoproducer flows.
- It can hide source/scenario combinations excluded by `VISIBLE_COMPARISON_SERIES`.

Use `comparison_long.csv` for broad analysis and `chart_line_mapping_ledger.csv` for rendered-chart semantics.

## Main Outputs

Primary outputs:

- `comparison_long.csv`
  Long comparison table across sources.

- `comparison_wide.csv`
  Wide comparison table with one source column per comparator.

- `mapping_status.xlsx`
  Mapping status and availability for chart rows.

- `dashboards/`
  Rendered HTML dashboard pages.

- `charts/`
  Individual chart HTML files.

Key audit outputs:

- `supporting_files/mapping/mapping_lineage_audit.csv`
  Row-level mapping lineage at audit years.

  By default this is sampled at `BASE_YEAR`, `BASE_YEAR + 1`, and `MAX_OUTPUT_YEAR`, not every year.

- `supporting_files/charting/chart_line_mapping_ledger.csv`
  Rows actually exposed to rendered charts.

- `supporting_files/charting/chart_total_component_ledger.csv`
  Components used in total chart rows.

- `supporting_files/checks/dashboard_comparator_pair_coverage.xlsx`
  Coverage audit for exposed ESTO and 9th comparator pairs.

- `supporting_files/checks/comparison_gap_diagnostics.csv`
  Large gaps between comparator series.

- `supporting_files/runtime/balance_runtime_issues.csv`
  Runtime mapping issues, including unmapped LEAP balance rows.

- `supporting_files/checks/ninth_mapping_data_coverage.xlsx`
  Checks whether active 9th mapping workbook pairs have data in the requested projection slice.

- `supporting_files/mapping/balance_missing_mapping_candidates.xlsx`
  Candidate rows for filling missing mapping workbook entries.

## Practical Debugging Workflow

When a chart value looks wrong:

1. Find the chart in `comparison_long.csv` by dashboard section, measure, fuel, source, scenario, and year.
2. Check `chart_line_mapping_ledger.csv` to see what the renderer actually exposed.
3. Check `mapping_lineage_audit.csv` for source sector/fuel rows behind the value.
4. Check `mapping_status.xlsx` for the ESTO flow/product and 9th sector/fuel targets.
5. Check the active rows in `leap_combined_esto` and `leap_combined_ninth`.
6. If the row is a transformation chart, confirm whether the value belongs to Inputs or Outputs based on sign.

When checking a rendered dashboard value, prefer this order:

1. `chart_line_mapping_ledger.csv`
2. `mapping_lineage_audit.csv`
3. `mapping_status.xlsx`
4. `comparison_long.csv`

`comparison_long.csv` is still important, but it is upstream of some render-time transformation logic.

## Common Failure Modes

- A valid 9th row is missing because the LEAP export row is zero or absent.
  The active workbook crosswalk should now prevent this when the 9th pair is mapped and nonzero.

- A direct ESTO-to-9th fallback brings in an unexpected pair.
  Check for `use_esto_to_ninth_mapping` in the template.

- A subtotal and child row are both present.
  Check subtotal flags in the mapping workbook and 9th data.

- A many-to-many mapping duplicates a value.
  Check whether the mapping is intentionally marked and whether the duplicate appears in rendered chart ledgers.

- `comparison_long.csv` does not explain the rendered chart value.
  Use `chart_line_mapping_ledger.csv`, especially for transformation input/output charts.

- A chart row appears with 9th data but no LEAP data.
  Check whether it came from the active workbook crosswalk. This is expected when a 9th pair is mapped and nonzero but the corresponding LEAP export row is zero or absent.

- A chart row appears in the data table but not the dashboard.
  Check `VISIBLE_COMPARISON_SERIES`, the dashboard template, and the rendered chart exposure files.
