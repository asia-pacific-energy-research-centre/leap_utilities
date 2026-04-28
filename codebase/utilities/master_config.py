from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[2]
MASTER_CONFIG_PATH = REPO_ROOT / "config" / "master_config.xlsx"

LEGACY_FILE_DEFAULT_SHEETS = {
    "ESTO_subtotal_mapping.xlsx": "ESTO_subtotal_mapping",
    "ninth_pairs_to_esto_pairs.xlsx": "ninth_pairs_to_esto_pairs",
    "leap_results_explicit_reassignments.csv": "leap_explicit_reassignments",
    "leap_results_sheet_map.csv": "leap_results_sheet_map",
    "leap_results_x_hierarchy_overrides.csv": "leap_x_hierarchy_overrides",
    "ninth_sector_fuel_pairs.csv": "ninth_sector_fuel_pairs",
    "synthetic_reference_rows.csv": "synthetic_reference_rows",
}

LEGACY_WORKBOOK_SHEETS = {
    ("sector_fuel_codes_to_names.xlsx", "9th"): "sector_fuel_codes_9th",
    ("sector_fuel_codes_to_names.xlsx", "ESTO"): "sector_fuel_codes_ESTO",
    ("sector_fuel_codes_to_names.xlsx", "code_to_name"): "sector_fuel_code_to_name",
    ("sector_fuel_codes_to_names.xlsx", "ESTO_LEAP_names"): "sector_fuel_ESTO_LEAP_names",
    ("leap_mappings.xlsx", "leap_combined_esto"): "leap_combined_esto",
    ("leap_mappings.xlsx", "leap_combined_ninth"): "leap_combined_ninth",
    ("independent product flow mappings.xlsx", "product"): "independent_product_mapping",
    ("independent product flow mappings.xlsx", "flow"): "independent_flow_mapping",
}

MASTER_SHEET_ALIASES = {
    "9th": "sector_fuel_codes_9th",
    "ESTO": "sector_fuel_codes_ESTO",
    "code_to_name": "sector_fuel_code_to_name",
    "ESTO_LEAP_names": "sector_fuel_ESTO_LEAP_names",
    "leap_combined_esto": "leap_combined_esto",
    "leap_combined_ninth": "leap_combined_ninth",
    "product": "independent_product_mapping",
    "flow": "independent_flow_mapping",
}


def _master_sheet_exists(sheet_name: str | None) -> bool:
    if not sheet_name or not MASTER_CONFIG_PATH.exists():
        return False
    try:
        return str(sheet_name) in pd.ExcelFile(MASTER_CONFIG_PATH).sheet_names
    except Exception:
        return False


def resolve_master_config_sheet(path: str | Path, sheet_name: str | None = None) -> str | None:
    """Return the master_config.xlsx sheet that supersedes a legacy config file."""
    source = Path(path)
    if source.name == MASTER_CONFIG_PATH.name:
        if sheet_name is None:
            return None
        return MASTER_SHEET_ALIASES.get(str(sheet_name), str(sheet_name))

    if sheet_name is not None:
        mapped = LEGACY_WORKBOOK_SHEETS.get((source.name, str(sheet_name)))
        if mapped:
            return mapped

    return LEGACY_FILE_DEFAULT_SHEETS.get(source.name)


def config_table_exists(path: str | Path, sheet_name: str | None = None) -> bool:
    """Return True if a standalone config file or its master sheet exists."""
    master_sheet = resolve_master_config_sheet(path, sheet_name)
    if _master_sheet_exists(master_sheet):
        return True
    return Path(path).exists()


def read_master_config_sheet(sheet_name: str, **kwargs: Any) -> pd.DataFrame:
    kwargs.pop("encoding", None)
    return pd.read_excel(MASTER_CONFIG_PATH, sheet_name=sheet_name, **kwargs)


def read_config_table(
    path: str | Path,
    sheet_name: str | None = None,
    **kwargs: Any,
) -> pd.DataFrame:
    """Read a config table from master_config.xlsx or from a standalone file.

    Legacy mapping files that have been consolidated into config/master_config.xlsx
    are resolved by filename and, for multi-sheet workbooks, by source sheet name.
    Any path not represented in master_config.xlsx is read normally.
    """
    path = Path(path)
    master_sheet = resolve_master_config_sheet(path, sheet_name)
    if _master_sheet_exists(master_sheet):
        return read_master_config_sheet(master_sheet, **kwargs)

    if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        if sheet_name is not None:
            kwargs["sheet_name"] = sheet_name
        return pd.read_excel(path, **kwargs)

    kwargs.pop("sheet_name", None)
    return pd.read_csv(path, **kwargs)
