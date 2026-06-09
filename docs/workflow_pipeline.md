# Results Supply Link Workflow — Pipeline Reference

This document describes the full data pipeline in `codebase/results_supply_link_workflow.py`
and related modules. It is intended as a working reference for understanding what each step
does, what data it consumes, and how adjustments propagate through to the LEAP model.

---

## Model structure context

The LEAP transformation side is divided into four broad groups:

### 1. Power
Least-cost optimisation (NEMO via LEAP). Most complex. Requires more iteration and checking
because LEAP sits between the user and the underlying solver, making some variables less
transparent. Not managed by this workflow — handled separately.

### 2. Refining
Simulation model with capacity data. Represents how a real refining system supplies multiple
oil products from one system. Capacity and product split estimation happens outside LEAP.
Relatively straightforward compared to power.

**Current status in workflow:** Partially included via a fallback path only.
`_load_refinery_fallback_table` reads LEAP results workbooks directly. Oil refineries
(`09.07`) is **not** in `ANALYSIS_REGISTRY` and has no ESTO-derived process record
(no efficiency, no feedstock shares, no auxiliary ratios). Multi-output nature
(gasoline, diesel, jet fuel, etc.) requires bespoke analysis — see the hydrogen
transformation function as a template.

### 3. Other transformation sectors
Simpler simulation models, similar in structure to refining, with capacity constraints.
Some (LNG liquefaction/regasification, gas processing) are important for specific
economies. Others (non-specified transformation, patent fuel plants) often replicate
historical ESTO balance relationships with unusual parameters (efficiency > 100%,
counter-intuitive inputs/outputs) and are best left largely unchanged.

Sectors currently in `ANALYSIS_REGISTRY`:
- LNG liquefaction / regasification
- Gas works plants / natural gas blending
- Coal: coke ovens, blast furnaces, patent fuel plants, BKB/PB plants, liquefaction
- Electric boilers, chemical heat for electricity, petrochemical industry, gas-to-liquids,
  biofuels processing
- Charcoal processing
- Non-specified transformation
- Hydrogen transformation

### 4. Transfers
Upstream liquids transfers, Refinery and blending transfers, Transfers unallocated.
Created to make sense of ESTO "Transfers" fuel flows that couldn't be cleanly split.
Generally unusual parameters; best left unchanged unless refining/oil extraction changes.

---

## Full pipeline (one run)

### PRE-RUN (optional data collection)

**`refresh_fuel_branch_catalog_from_leap`**
Probes the live LEAP model for all existing transformation/supply fuel branches.
Output is a CSV at `LEAP_FUEL_BRANCH_PROBE_OUTPUT_PATH`. Used by the zero-fill step
to know which branches need to be explicitly cleared.

**`_run_leap_results_template_scrape`**
Reads LEAP Results API and writes results tables to disk for use in downstream steps.
Required before balance demand loading if results have changed since the last scrape.

---

### DATA LOADING

**`load_balance_demand_inputs`**
Reads LEAP balance export workbooks → `demand_table` and `sector_demand_table`.
This is what LEAP says demand is, per economy/sector/fuel/year.

**`build_transformation_balance_table` / `build_transformation_sector_table`**
Reads LEAP results for transformation outputs and sector-level totals.
Provides `max_transformation_output` and `constrained_transformation_output` columns
used in the capacity unmet pass.

**`collect_transformation_rows` → `build_transformation_rows`** (ESTO analysis)
Runs all registered ESTO sector analyses. For each economy × sector:
- Identifies primary input/output fuels from ESTO balance
- Derives efficiency as `output / (total_feedstock_input + own_use_losses)`
- Derives feedstock shares (multi-feedstock: all inputs as shares of total)
- Derives auxiliary ratios from ESTO own-use/loss sectors (ratio per unit of output)
- Computes `own_use_ratios`: for fuels appearing in both feedstock and own-use data,
  stores `mean_own_use / (mean_own_use + mean_feedstock)` for use in later feedback
- Produces process records stored in `transformation_process_records`

Feedstock method (controlled by `FEEDSTOCK_METHOD`):
- `FEEDSTOCK_METHOD_MULTI` (current default): all inputs are feedstocks in one process;
  auxiliary fuels come from own-use/loss data only
- `FEEDSTOCK_METHOD_SPLIT`: one process per feedstock fuel, losses allocated by share
- `FEEDSTOCK_METHOD_SINGLE_AUX`: primary feedstock only; others become auxiliary

**`_refresh_transformation_measures_from_leap_results`** *(if enabled)*
Overwrites `feedstock_values` and `output_values` in process records with actual values
from the LEAP Results API (`Inputs` and `Outputs by Output Fuel` / `Outputs by Feedstock
Fuel`). This replaces the ESTO-derived values with what LEAP actually computed.

**`_apply_own_use_ratio_feedback`** *(called inside the above)*
After feedstock values are refreshed from LEAP, recalibrates `auxiliary_ratios` for
fuels that serve as both feedstock and own-use.
- Formula: `estimated_own_use(year) = leap_feedstock(year) × ratio / (1 − ratio)`
- New aux ratio: `estimated_own_use / total_output`
- Limitation: LEAP does not expose own-use separately from transformation inputs in its
  energy balance output, so the estimate is ESTO-ratio-based only and cannot be
  cross-checked against actual LEAP own-use values.

**`prepare_projected_supply_table` / `prepare_supply_primary_table`**
Builds ESTO/9th-derived supply projections and primary supply assets.

**`load_leap_constraint_tables`**
Reads capacity and production constraint templates (max output, max production per
module, per economy, per year).

---

### RECONCILIATION

**`build_reconciliation_table`**
Central balance step. Combines:
- `demand_table` (what LEAP says demand is)
- `transformation_table` (what transformation can supply)
- `supply_projection_table` (what supply resources can provide)
- Constraint tables

Computes:
- `adjusted_imports`: residual gap that needs to come from imports after domestic
  supply + transformation are maximised
- `adjusted_exports`: surplus available for export
- `max_transformation_output`, `constrained_transformation_output`

**`apply_trade_split_between_transformation_and_supply`**
Historically split import/export residuals between transformation targets and supply.
Currently a no-op for the transformation side (`_use_legacy_trade_split_mode()` = False,
so transformation always gets 0). Still needed because it creates the
`supply_imports_residual` and `supply_exports_residual` columns used by downstream steps.

**`reset_supply_and_transformation_import_export_to_zero`** *(if enabled)*
Zeroes stale LEAP values before writing new ones. Covers:
- Supply import/export targets
- Transformation import/export targets
- Auxiliary fuel use and feedstock fuel shares for branches in the catalog
  (zero-fills any catalog branch not explicitly set in this run, preventing old
  values from persisting across runs)

---

### CAPACITY UNMET ITERATIVE PASS *(optional, multi-pass)*

**Trigger condition:** `unmet_proxy = observed_imports − adjusted_imports > 0`
If LEAP is importing more than the reconciliation said it should, that gap is treated
as a proxy for unmet transformation capacity (the module can't produce what's needed).

**Per fuel / economy / year:**
1. Check remaining capacity headroom:
   `headroom = max_transformation_output − constrained_output − prior_pass_additions`
2. Allocate `min(unmet_proxy, headroom)` as an output uplift
3. Find eligible transformation modules for that fuel via `process_catalog`
   (ranked by preference, e.g. base-year output share)
4. Convert output uplift → capacity addition:
   `capacity_addition = output_uplift / module_yield`
   where yield = output PJ per PJ of installed capacity from ESTO
5. Add to cumulative state persisted in `CAPACITY_UNMET_STATE_PATH` (JSON on disk)

**Configuration source:** the capacity-unmet priority lists and cap limits are
loaded from `config/results_supply_link_config.json`. That file is the static
policy/config input; `CAPACITY_UNMET_STATE_PATH` is the mutable runtime state
file written under `outputs/.../runtime/`.

**Between passes:**
State persists across runs. Each new run reads cumulative additions from the state file,
merges with fresh LEAP results, computes the delta, and writes an updated import workbook.
Convergence is reached when `observed_imports ≈ adjusted_imports` for all constrained fuels.

**Manual loop:**
1. Run workflow → generates LEAP import workbook with exogenous capacity additions
2. Import workbook into LEAP
3. Recalculate LEAP model
4. Refresh results tables (scrape)
5. Re-run workflow → reads new observed_imports, computes new delta
6. Repeat until converged

**Balanced variant** (`_run_capacity_unmet_iterative_balanced_pass`):
Also manages exports and applies a `max_production` ceiling to prevent capacity from
overshooting demand. Two pass modes:
- `baseline_seed`: first pass only — writes imports=0 with baseline exports/capacity,
  no residual allocation from existing results
- `results_update`: subsequent passes — reads LEAP results and layers in residuals

---

### OUTPUT GENERATION

**`_build_transformation_supply_fuel_catalog_df`**
Builds the branch catalog from three sources (priority order):
1. Full model export file at `FULL_MODEL_EXPORT_CATALOG_PATH` (if enabled)
2. LEAP probe output CSV (`LEAP_FUEL_BRANCH_PROBE_OUTPUT_PATH`)
3. Previously generated transformation/supply export workbooks

If a fuel branch is not in any of these sources, zero-fill rows won't be generated for it.
Bootstrapping problem: a branch never written to a workbook will never appear in source 3,
so it must first appear in the LEAP model (source 1 or 2).

**`save_transformation_leap_export`** → **`build_aux_fuel_zero_rows`**
Writes the transformation LEAP import workbook. Before writing data rows, generates
zero-fill rows for:
- Auxiliary Fuels (`Auxiliary Fuel Use`) — all catalog branches not already set
- Feedstock Fuels (`Feedstock Fuel Share`) — all catalog branches not already set

This ensures fuels that don't exist in a given economy/module are explicitly cleared
to 0, preventing stale values from prior runs.

**`run_results_linked_leap_import`** *(if enabled)*
Pushes the generated workbooks into LEAP via the COM API. Runs sequentially across
scenarios, with Current Accounts handled first when reset mode is active.

---

## Key adjustment points summary

| Step | What changes | Source data |
|---|---|---|
| ESTO analysis | Efficiency, feedstock shares, aux ratios | ESTO historical balance |
| LEAP results refresh | Feedstock values, output values | LEAP Results API |
| Own-use feedback | Aux ratios recalibrated | LEAP feedstock × ESTO own-use ratio |
| Reconciliation | Import/export residuals | LEAP demand − supply − transformation |
| Trade split | Routes all residuals to supply (currently) | Reconciliation |
| Zero-fill | Aux fuel and feedstock branches cleared | Branch catalog |
| Capacity unmet pass | Exogenous capacity, observed imports | LEAP observed vs expected imports |
| LEAP import | All adjusted values written to LEAP | Generated workbooks |

---

## Known gaps / future work

- **Refining not ESTO-analysed:** Oil refineries has no process record. Multi-output
  analysis (similar to hydrogen) needed before it can be included in the standard path.
- **Efficiency not refreshed from LEAP:** If LEAP's actual process efficiency diverges
  from ESTO-derived values, that difference is not fed back. The ESTO value is used
  unchanged in all passes.
- **Own-use not separable from LEAP balance:** LEAP aggregates auxiliary fuel use
  into transformation inputs in its energy balance, so own-use feedback estimates
  rely entirely on ESTO ratios and cannot be verified against LEAP output.
- **Capacity unmet convergence not guaranteed:** If a fuel's capacity headroom is
  fully consumed before demand is met, the remaining gap persists as unconstrained
  imports with no further allocation possible without manual capacity increases.
