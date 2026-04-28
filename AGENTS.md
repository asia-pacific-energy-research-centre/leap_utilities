# AGENTS.md

These are project-level instructions for Codex (and similar agents).

## When editing draw.io diagrams

- See `AGENTS_DRAWIO.md` for draw.io-specific requirements.

## Small guide for humans

- Put instructions here that you want Codex to follow every time it edits this repo.
- Keep rules short and specific; avoid large, complex policies.
- For file-specific rules, include path globs like `docs/leap-system*.drawio`.
- Workflow-file pattern for small projects: create/maintain one `*_workflow.py` entry script per task area and make it notebook-safe.
- In workflow scripts, always define `REPO_ROOT = Path(__file__).resolve().parents[1]` (or correct repo level), add it to `sys.path` only if missing, and resolve all relative paths via a `_resolve()` helper against `REPO_ROOT`.
- Why: notebooks run with arbitrary CWD, so this prevents `FileNotFoundError` and import failures.
- Normalize user-provided path strings by replacing `\\` with `/` before `Path(...)` when needed.
- When updating transfer category mappings, re-run `codebase/scrapbook/transfers_mapping_exploration.py`
  and paste the printed `TRANSFER_PROCESS_CONFIG` into `codebase/transfers_workflow.py`.
- When referring to files in replies, prefer paths relative to the active repo root
  (for example, `outputs/example.csv`) instead of absolute `/mnt/c/...` or
  `C:\...` paths. Use absolute paths only for files outside the repo or when needed
  to disambiguate.

## Output clarity

- Keep output folders small and easy to inspect.
- Prefer a few clearly named primary outputs.
- Do not create extra files unless they serve a clear human-facing purpose.
- Keep primary outputs narrow: include important columns only.
- Put debug-heavy or trace-heavy artifacts in `extra_detail` or `diagnostics`, not beside the main outputs.
- Make sure there is a clear file for inspecting errors when needed.

## Dashboard Comparison Policies (Scoped)

Apply this section only when editing dashboard mapping/comparison code
(`codebase/leap_results_dashboard*_workflow.py`, `codebase/utilities/leap_results_dashboard*.py`).

### Purpose

- Build a dashboard that compares LEAP results, 9th projections, and ESTO base-year values with transparent mapping lineage.
- Treat mapping quality as a core output, not a side artifact, because these mappings are reused for future analysis.

### Mapping Model

- Maintain explicit mapping logic for each dataset path (LEAP, 9th, ESTO), then reconcile into comparable series.
- Allow one-to-many and many-to-one mappings when they are intentional and auditable.
- Raise an error on many-to-many mapping situations (ambiguous crosswalks that cannot be deterministically resolved).
- Prefer unique sector+product canonical resolution before aggregated canonical fallback.
- Use aggregated canonical mappings only as fallback when unique product-level resolution is unavailable.

### Hierarchy and Display Level

- Dashboard pages should follow sector hierarchy (sector, subsector, deeper subsectors) and fuels, including total charts.
- Include total-by-fuel-set charts for each subsector grouping and overall sector totals.
- Show the lowest category level that exists across all compared datasets for that charted group.
- If one dataset is only available at a parent level, aggregate children from the other datasets up to that parent level.
- When parent level is shown for comparability, do not also show conflicting child-level lines for that same comparison group.

### Aggregation and Totals

- Aggregation to parent must be comprehensive and non-duplicative.
- Totals should be computed by summing the displayed child components at the resolved comparison level.
- Do not use `max`-based bucket dedupe for totals once child values are allocated/disaggregated.
- Resolve duplicate-assignment risk upstream in mapping/allocation logic, not by suppressing totals downstream.

### Diagnostics and Auditability

- Always produce mapping diagnostics and line-level ledgers so each plotted value can be traced to mapping decisions.
- Keep outputs that make aggregate composition auditable (for example, total-component ledgers and mapping-status artifacts).
- Distinguish mapping differences from methodological differences in interpretation; mismatches are expected in some cases and should be explainable.

### Validation Checklist

- Run the dashboard workflow and require successful completion:
  - `python3 codebase/leap_results_dashboard_v2_workflow.py`
- Confirm required outputs exist:
  - `outputs/dashboards/leap_results_dashboard_v2/USA/comparison_long.csv`
  - `outputs/dashboards/leap_results_dashboard_v2/USA/mapping_status.xlsx`
  - `outputs/dashboards/leap_results_dashboard_v2/USA/chart_line_mapping_ledger.csv`
  - `outputs/dashboards/leap_results_dashboard_v2/USA/chart_total_component_ledger.csv`
  - `outputs/dashboards/leap_results_dashboard_v2/USA/comparison_gap_diagnostics.csv`
  - `outputs/dashboards/leap_results_dashboard_v2/USA/mapping_rundown_by_sheet.csv`
- Confirm no unresolved mapping policy violations:
  - no many-to-many mapping situations left unresolved
  - no parent/child double-display conflicts for the same comparison group
- Confirm totals logic:
  - comparator `Total` series equals sum of displayed child comparator rows at the resolved comparison level
  - no `max`-based bucket suppression used as a total-construction shortcut after child allocation/disaggregation
- Spot-check at least one known-problem sheet (for example `Non-specified industry`) using ledgers:
  - verify mapping rows are explainable
  - verify total component sums match plotted totals

## LEAP Export File Structure

- See `C:\\Users\\Work\\.codex\\AGENTS_LEAP_EXPORT.md` for LEAP export structure requirements.

## Balance Table Structures (ESTO vs 9th)

- See `C:\\Users\\Work\\.codex\\AGENTS_BALANCE_TABLES.md` for balance table structure details.

These two balance tables are the core inputs for `codebase/transformation_analysis_workflow.py`.
Keep this structure in mind when adding new transformations or debugging data issues.

### 9th structure (sector/fuel hierarchy)

- Source file: `data/merged_file_energy_ALL_20250814_pre_trump.csv` (loaded as "9th" in the script).
  - Use `data/merged_file_energy_ALL_20251106.csv` and `data/merged_file_energy_00_APEC_20251106` when you need to exactly match 9th edition projections.
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

## Python Environment

- This repo's `.venv` is a WSL-created venv (`home = /usr/bin` in `pyvenv.cfg`) and cannot be used from Windows shells (PowerShell, cmd, or the Bash tool when running in a Git-Bash context on Windows).
- Use `/c/Users/Work/miniconda3/python.exe` for all Python scripts run via the Bash tool (Git-Bash on Windows).
- Do **not** attempt to activate `.venv/bin/activate` from the Bash tool — it will fail silently or error.
- Do **not** use PowerShell's `python` or `py` aliases — output is swallowed and exit codes are unreliable.
