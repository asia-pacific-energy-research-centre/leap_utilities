# Full Model Workflow Post-Run Guide

Note: this guide explains the manual LEAP checks and follow-up actions to run
after `codebase/full_model_workflow_notebook.py` completes. It records checks
that are expected to happen in the LEAP GUI or by inspecting generated import
workbooks. Use it to catch skipped imports, unit-scale issues, and known manual
settings before treating a full model run as complete.

Companion guide for `codebase/full_model_workflow_notebook.py`.
Use this after running `run_all_workflows()` to capture manual LEAP steps that are expected or commonly required.

## Quick use

1. Run the notebook end-to-end.
2. Open LEAP and apply the checklist below.
3. Mark each item as done for the economy/scenario you just processed.

## Post-run checklist

### 1) Confirm LEAP import actually ran

- [ ] Confirm LEAP API was available during run.
- [ ] If you saw messages like `LEAP API unavailable ... skipping branch creation/fill`, import the generated workbook(s) manually in LEAP.

Typical output files are in `outputs/leap_exports/standalone/`, `outputs/leap_exports/results_supply_link/`, and `outputs/leap_exports/combined/`.

### 2) Industry share/saturation scale fix (known issue)

Context: documented in comments in `codebase/industry_workflow.py` and README.

- [ ] For Industry share-based variables (especially `Activity Level` where shares must sum to 100%), check `Scale` in LEAP GUI.
- [ ] If scale is wrong/missing, re-select `Units` as `Share` on fuel leaf nodes.
- [ ] For top-level saturation-style variables, re-select `Units` as `Saturation`.
- [ ] Spot-check intensity variables for missing scale/unit behavior.
- [ ] Validate in LEAP results tables that totals and shares behave correctly.

Directive summary: if `%`-style scale is wrong after import, reset unit in GUI (`Share`/`Saturation`) so LEAP re-applies scale correctly.

### 3) Refining variables intentionally skipped by code

Context: `SKIP_VARIABLES` in `codebase/refining_workflow.py`.

These are intentionally not set by `fill_branches_from_export_file` and must be reviewed/set manually where needed:

- [ ] `Dispatchable`
- [ ] `Optimize`
- [ ] `Surplus Rule`
- [ ] `Shortfall Rule`
- [ ] `Priority Output`
- [ ] `Dispatch Rule`

### 4) Optimization/capacity variable units

- [ ] For transformation sectors using optimization (including oil refining), confirm units and `Per...` are correct for capacity-style variables.
- [ ] Explicitly verify variables like `Exogenous Capacity` and `Endogenous Capacity` in LEAP GUI where present.
- [ ] Ensure unit choices are consistent across scenarios (Current Accounts/Reference/Target).

Why: API-based branch creation can leave default/blank unit metadata for some branch types, which can produce incorrect behavior unless corrected in LEAP.

### 5) Branches/nodes the LEAP API may not create

Context: constraints in `codebase/functions/leap_core.py`.

If branch creation logs warned about manual creation, create the missing nodes in LEAP, then re-run fill/import.

Known cases:

- [ ] Stock-based demand branches (`(road)`) other than `Fuel (road)` must be created manually.
- [ ] Demand fuel branches cannot be auto-created directly by API (`AddDemandFuel` is unavailable).
- [ ] Missing Transformation process categories may need manual creation before child branches can be filled.
- [ ] Transformation processes that fail API creation must be added manually, then data can be filled.

### 6) Supply unmet requirements (if excluded by config)

Context: `SUPPLY_INCLUDE_UNMET_REQUIREMENTS = False` in `codebase/configuration/workflow_config.py` leaves it for manual LEAP setup.

- [ ] If this flag was `False`, manually create/configure `Unmet Requirements` measures in LEAP where needed.

### 7) General unit/scale sanity pass

- [ ] For newly created technology branches, verify unit metadata (some are created with blank defaults by API helpers).
- [ ] Check `Units`, `Scale`, and `Per...` on representative branches in each workflow module (Transformation, Supply, Transfers, Minor demand, Industry).
- [ ] Resolve any share/saturation normalization warnings by correcting units/scale and rechecking totals.

## Suggested run log template

Copy this block per run:

```text
Date:
Economy:
Scenarios:
Notebook commit/version:

[ ] LEAP API import confirmed (or manual import done)
[ ] Industry share/saturation scales corrected
[ ] Refining skipped control variables set
[ ] Optimization capacity units verified (incl. Exogenous/Endogenous where used)
[ ] Missing API-uncreatable nodes created manually
[ ] Supply Unmet Requirements handled (if excluded)
[ ] Final unit/scale sanity pass complete

Notes:
```

## Sources for this checklist

- `codebase/full_model_workflow_notebook.py`
- `codebase/industry_workflow.py`
- `codebase/refining_workflow.py`
- `codebase/configuration/workflow_config.py`
- `codebase/functions/leap_core.py`
- `README.md`
