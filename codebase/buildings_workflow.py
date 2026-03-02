#%%
from __future__ import annotations

import sys
from pathlib import Path
from typing import Sequence

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if REPO_ROOT.exists() and str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.configuration.config import (
    BRANCH_DEMAND_CATEGORY,
    BRANCH_DEMAND_TECHNOLOGY,
)
from codebase.configuration.buildings_technology_mapping import BUILDINGS_TECHNOLOGY_MAPPING
from codebase.functions.buildings_fuel_remap import remap_buildings_export_fuels
from codebase.functions.leap_core import (
    connect_to_leap,
    create_branches_from_export_file,
    fill_branches_from_export_file,
)
from codebase.functions.leap_exports import list_scenarios as list_export_scenarios
from codebase.functions.leap_excel_io import read_export_sheet

CREATE_BRANCHES_FROM_EXPORT_FILE = True
FILL_BRANCHES_FROM_EXPORT_FILE = True
HANDLE_CURRENT_ACCOUNTS_TOO = True
SET_UNITS = True

LEAP_EXPORT_FILENAME = "../data/buildings export.xlsx"
SHEET_NAME = "Export"
ECONOMY = "20_USA"
BASE_YEAR = 2022
SCENARIOS = ["Reference"]
REGION = "United States of America"

REMAP_FUELS = True
SERIES_FORMAT_POLICY = "preserve"  # preserve | expression | year_columns
ESTO_DATA_PATH = "../data/00APEC_2024_low.csv"
ESTO_SUBTOTAL_MAPPING_PATH = "../config/ESTO_subtotal_mapping.xlsx"
REMAP_OUTPUT_PATH = (
    REPO_ROOT / "outputs" / "leap_exports" / "buildings_export_remapped_20_USA.xlsx"
)
REMAP_REPORT_PATH = "../intermediate_data/buildings_fuel_remap_report.csv"
REMAP_VALIDATION_PATH = "../intermediate_data/buildings_fuel_remap_validation.csv"


def _normalize_expected(values: Sequence[str]) -> list[str]:
    seen = set()
    out: list[str] = []
    for value in values:
        text = str(value).strip()
        if not text:
            continue
        key = text.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(text)
    return out


def _assert_export_matches_expected(
    export_filename: str | Path,
    sheet_name: str,
    expected_scenarios: Sequence[str],
    expected_region: str,
) -> None:
    _, data, _ = read_export_sheet(export_filename, sheet_name)
    if "Scenario" not in data.columns or "Region" not in data.columns:
        raise ValueError("Export sheet must include Scenario and Region columns.")

    present_scenarios = sorted(
        {
            str(value).strip()
            for value in data["Scenario"].dropna().astype(str)
            if str(value).strip()
        },
        key=lambda value: value.lower(),
    )
    expected = _normalize_expected(expected_scenarios)
    expected_set = {value.lower() for value in expected}
    present_set = {value.lower() for value in present_scenarios}
    if present_set != expected_set:
        raise ValueError(
            "Scenario mismatch in buildings export. "
            f"Expected={expected} Present={present_scenarios}"
        )

    regions = sorted(
        {
            str(value).strip()
            for value in data["Region"].dropna().astype(str)
            if str(value).strip()
        },
        key=lambda value: value.lower(),
    )
    if len(regions) != 1 or regions[0].lower() != str(expected_region).strip().lower():
        raise ValueError(
            "Region mismatch in buildings export. "
            f"Expected='{expected_region}' Present={regions}"
        )


def _discover_fill_scenarios(export_filename: str | Path, sheet_name: str) -> list[str]:
    raw_scenarios = list_export_scenarios(Path(export_filename), sheet_name=sheet_name)
    available_by_key = {
        str(item).strip().lower(): str(item).strip()
        for item in raw_scenarios
        if str(item).strip()
    }
    resolved = []
    for scenario in SCENARIOS:
        key = str(scenario).strip().lower()
        if key not in available_by_key:
            raise ValueError(f"Scenario '{scenario}' is missing from export workbook.")
        resolved.append(available_by_key[key])
    return resolved


L = connect_to_leap()

_assert_export_matches_expected(
    export_filename=LEAP_EXPORT_FILENAME,
    sheet_name=SHEET_NAME,
    expected_scenarios=[*SCENARIOS, "Current Accounts"],
    expected_region=REGION,
)

if REMAP_FUELS:
    remap_buildings_export_fuels(
        input_path=LEAP_EXPORT_FILENAME,
        output_path=REMAP_OUTPUT_PATH,
        mapping_csv_path=None,
        mapping_dict=BUILDINGS_TECHNOLOGY_MAPPING,
        esto_data_path=ESTO_DATA_PATH,
        subtotal_mapping_path=ESTO_SUBTOTAL_MAPPING_PATH,
        economy=ECONOMY,
        base_year=BASE_YEAR,
        sheet_name=SHEET_NAME,
        include_extra_nonspecified=True,
        report_path=REMAP_REPORT_PATH,
        validation_path=REMAP_VALIDATION_PATH,
        output_series_format=SERIES_FORMAT_POLICY,
    )
    LEAP_EXPORT_FILENAME = REMAP_OUTPUT_PATH
    SHEET_NAME = "LEAP"

scenarios_to_fill = _discover_fill_scenarios(LEAP_EXPORT_FILENAME, SHEET_NAME)

if CREATE_BRANCHES_FROM_EXPORT_FILE:
    create_branches_from_export_file(
        L,
        LEAP_EXPORT_FILENAME,
        sheet_name=SHEET_NAME,
        branch_path_col="Branch Path",
        scenario=None,
        region=REGION,
        branch_type_mapping=None,
        default_branch_type=(
            BRANCH_DEMAND_CATEGORY,
            BRANCH_DEMAND_CATEGORY,
            BRANCH_DEMAND_TECHNOLOGY,
        ),
        RAISE_ERROR_ON_FAILED_BRANCH_CREATION=True,
    )

if FILL_BRANCHES_FROM_EXPORT_FILE:
    for idx, scenario_name in enumerate(scenarios_to_fill):
        include_current_accounts = HANDLE_CURRENT_ACCOUNTS_TOO and idx == 0
        fill_branches_from_export_file(
            L,
            LEAP_EXPORT_FILENAME,
            sheet_name=SHEET_NAME,
            scenario=scenario_name,
            region=REGION,
            RAISE_ERROR_ON_FAILED_SET=True,
            SET_UNITS=SET_UNITS,
            HANDLE_CURRENT_ACCOUNTS_TOO=include_current_accounts,
            CHECK_STALE_CHILD_BRANCHES=True,
            PROMPT_DELETE_STALE_BRANCHES=True,
        )
