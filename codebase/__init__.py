"""
Expose the main LEAP utility helpers used by workflows and notebooks.

This package initializer re-exports common branch, Excel import/export, API,
and comparison functions so callers can import them from `codebase`. Keep this
file limited to stable public helpers; workflow-specific logic should stay in
the workflow modules.
"""

from .functions.leap_core import (
    connect_to_leap,
    safe_branch_call,
    build_expr,
    safe_set_variable,
    build_expression_from_mapping,
    ensure_branch_exists,
    diagnose_measures_in_leap_branch,
    create_branches_from_export_file,
    fill_branches_from_export_file,
    create_transformation_module,
    create_transformation_process,
    create_transformation_output,
    create_transformation_feedstock,
    create_simple_transformation_process,
    get_resource_branch_for_fuel,
    ensure_fuel_exists,
    ensure_unit_exists,
)

from .functions.leap_excel_io import (
    create_import_instructions_sheet,
    finalise_export_df,
    save_export_files,
    join_and_check_import_structure_matches_export_structure,
    separate_current_accounts_from_scenario,
    copy_energy_spreadsheet_into_leap_import_file,
)

from .functions.leap_exports import (
    build_workbook_filename,
    build_export_dataframe,
    save_workbook,
    build_and_save_workbook,
    list_scenarios,
    validate_region,
    find_workbook,
)

from .functions.leap_api import (
    is_available,
    connect,
    import_workbook,
)
from .functions import leap_exports as leap_exports
from .functions import leap_api as leap_api

from .functions.energy_use_reconciliation import (
    build_branch_rules_from_mapping,
    reconcile_energy_use,
)
from .functions.leap_series_comparison import (
    ComparisonRunConfig,
    ComparisonArtifacts,
    run_leap_series_comparison,
)

from .configuration.config import region_id_name_dict, scenario_dict

__all__ = [
    # core
    "connect_to_leap",
    "safe_branch_call",
    "build_expr",
    "safe_set_variable",
    "build_expression_from_mapping",
    "ensure_branch_exists",
    "diagnose_measures_in_leap_branch",
    "create_branches_from_export_file",
    "fill_branches_from_export_file",
    "create_transformation_module",
    "create_transformation_process",
    "create_transformation_output",
    "create_transformation_feedstock",
    "create_simple_transformation_process",
    "get_resource_branch_for_fuel",
    "ensure_fuel_exists",
    "ensure_unit_exists",
    # excel io
    "create_import_instructions_sheet",
    "finalise_export_df",
    "save_export_files",
    "join_and_check_import_structure_matches_export_structure",
    "separate_current_accounts_from_scenario",
    "copy_energy_spreadsheet_into_leap_import_file",
    # packaged export helpers
    "build_workbook_filename",
    "build_export_dataframe",
    "save_workbook",
    "build_and_save_workbook",
    "list_scenarios",
    "validate_region",
    "find_workbook",
    "leap_exports",
    "leap_api",
    # packaged LEAP API helpers
    "is_available",
    "connect",
    "import_workbook",
    # reconciliation
    "build_branch_rules_from_mapping",
    "reconcile_energy_use",
    # series comparison
    "ComparisonRunConfig",
    "ComparisonArtifacts",
    "run_leap_series_comparison",
    # economy config
    "region_id_name_dict",
    "scenario_dict",
]
