# Workflow Inventory

Last reviewed: 2026-04-14

This is a cleanup-oriented inventory of the workflow entry scripts.
Current workflows remain in `codebase/`; old/probe/scaffold workflows live in
`codebase/old_workflows/`.

## Keep As Core Entry Points

- `full_model_workflow_notebook.py` - orchestrates the broader model run.
- `results_supply_link_workflow.py` - main linked results/transformation/supply workflow.
- `leap_results_dashboard_balance_estoaxis_workflow.py` - selected balance-dashboard entrypoint.
- `leap_results_dashboard_v2_workflow.py` - current dashboard workflow; required by `AGENTS.md`.
- `leap_results_workflow.py` - LEAP Results export/template-fill API used by V2 and replica tooling.
- `minor_demand_workflow.py` - used directly by `full_model_workflow_notebook.py`.
- `transformation_workflow.py`, `transfers_workflow.py`, `hydrogen_transformation_workflow.py`, `supply_workflow.py` - direct dependencies of the kept full-model/results-supply entrypoints.
- `transformation_entry.py` - convenience entrypoint used by `full_model_workflow_notebook.py`.

## Kept With A Specific Case

- `leap_results_dashboard_workflow.py` - V1 dashboard. It looks superseded by V2, but V2 support utilities still import functions/constants from it.
- `leap_results_dashboard_balance_workflow.py` - balance dashboard workflow; covered by tests.
- `power_workflow.py` - README/docs describe it as the standalone power import path, and tests import it.
- `buildings_workflow.py`, `industry_workflow.py`, `refining_workflow.py` - standalone sector import workflows. They are not direct dependencies of the kept three, but they are still user-facing entrypoints for sector-specific imports.

## Moved To `codebase/old_workflows/`

These are retained for reference, but no longer sit beside the active entrypoints:

- `aperc_reference_aggregation_workflow.py` - large prototype workflow with prototype output folder; keep only if this analysis is still active.
- `detailed_balance_from_esto_workflow.py`, `energy_balance_template_extract_workflow.py` - balance-template utilities.
- `leap_favorites_transplant_workflow.py` - LEAP area/favorites maintenance tool.
- `leap_results_extraction_replica_workflow.py` - specialized replica extraction workflow.
- `leap_results_template_year_axis_audit_workflow.py` - small audit wrapper around `leap_results_workflow.py`.
- `leap_transformation_losses_own_use_workflow.py` - docstring says it is legacy and no longer part of the official dashboard path.
- `ninth_to_esto_mapping_coverage_workflow.py` - mapping diagnostics wrapper.
- `synthetic_reference_mapping_suggestions_workflow.py` - dashboard mapping maintenance helper.
- `leap_balance_mapping_scaffold_workflow.py` - scaffold generator, likely only needed while rebuilding balance mappings.
- `leap_individual_mapping_scaffold_workflow.py` - scaffold generator, likely only needed while rebuilding individual mappings.
- `leap_results_variable_probe_workflow.py` - LEAP Results probe/debug workflow.
- `transformation_leap_probe_workflow.py` - transformation-results probe/debug workflow.
- `unmet_requirements_results_probe_workflow.py` - Resources/Unmet Requirements probe/debug workflow.

## Already Moved Out Of `codebase/`

- `codebase/radar chart - prc.py` moved to `codebase/scrapbook/radar_chart_prc.py`; it is a one-off plotting script, not a workflow entrypoint.
