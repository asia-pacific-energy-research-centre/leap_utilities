from __future__ import annotations

from pathlib import Path

from codebase.functions.leap_core import connect_to_leap, is_leap_api_available
from codebase.utilities.workflow_common import import_workbook_to_leap as _import_workbook_to_leap


def is_available() -> bool:
    """Return True when LEAP COM API is available in this environment."""
    return is_leap_api_available()


def connect(force_rebuild: bool = True):
    """Return a connected LEAP application object."""
    return connect_to_leap(force_rebuild=force_rebuild)


def import_workbook(
    export_path: Path,
    *,
    sheet_name: str,
    scenario: str | None,
    region: str | None,
    create_branches: bool = True,
    fill_branches: bool = True,
    include_current_accounts: bool = True,
    default_branch_type: tuple | None = None,
    branch_type_mapping: dict | None = None,
    branch_root: str | None = None,
    branch_path_col: str | None = None,
    raise_on_missing_branch: bool = False,
) -> Path:
    """Create/fill LEAP branches from a workbook using the shared import routine."""
    return _import_workbook_to_leap(
        export_path=export_path,
        sheet_name=sheet_name,
        scenario=scenario,
        region=region,
        create_branches=create_branches,
        fill_branches=fill_branches,
        include_current_accounts=include_current_accounts,
        default_branch_type=default_branch_type,
        branch_type_mapping=branch_type_mapping,
        branch_root=branch_root,
        branch_path_col=branch_path_col,
        raise_on_missing_branch=raise_on_missing_branch,
    )
