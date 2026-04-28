
#%%
"""
Generate a LEAP buildings dummy export workbook from ESTO and 9th data.

The generated workbook mirrors ``data/dummy buildings export.xlsx`` but expands
``Demand\\Buildings dummy`` into:
- Datacentres
- Residential
- Commercial and public services

Each child branch receives one technology/fuel branch for every aggregate fuel
present in the selected source data. Values are written as LEAP ``Data(...)``
expressions on the ``Total Energy`` variable.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Iterable, Sequence

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[1]
if REPO_ROOT.exists() and str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.configuration.all_products_and_flows import ESTO_PRODUCT_LIST
from codebase.configuration.config import region_id_name_dict, scenario_dict
from codebase.functions.leap_excel_io import read_export_sheet, write_export_sheet
from codebase.functions.leap_expressions import build_data_expression
from codebase.functions.leap_labels import clean_fuel_label_for_leap
from codebase.functions.ninth_projection_mapping import normalize_economy_key
from codebase.utilities.output_paths import STANDALONE_LEAP_EXPORTS_ROOT


TEMPLATE_EXPORT_PATH = REPO_ROOT / "data" / "dummy buildings export.xlsx"
OUTPUT_EXPORT_PATH = STANDALONE_LEAP_EXPORTS_ROOT / "buildings_dummy_20_USA.xlsx"
DIAGNOSTICS_PATH = (
    STANDALONE_LEAP_EXPORTS_ROOT / "supporting_files" / "buildings_dummy_20_USA_diagnostics.csv"
)
SHEET_NAME = "Export"

ECONOMY = "20_USA"
REGION = "United States of America"
BASE_YEAR = 2022
SCENARIOS = ["Reference"]
INCLUDE_CURRENT_ACCOUNTS = True

NINTH_DATA_PATH = REPO_ROOT / "data" / "merged_file_energy_ALL_20251106.csv"
ESTO_DATA_PATH = REPO_ROOT / "data" / "00APEC_2024_low.csv"
PROJECTION_START_YEAR = 2023
PROJECTION_END_YEAR = 2070

ROOT_BRANCH = r"Demand\Buildings dummy"
SOURCE_BRANCHES = {
    "Datacentres": {
        "ninth_sub2sectors": [
            "16_01_03_ai_training",
            "16_01_04_traditional_data_centres",
        ],
        # ESTO has no datacentres split in 00APEC_2024_low.csv; use 9th base-year fallback.
        "esto_flows": [],
    },
    "Residential": {
        "ninth_sub2sectors": ["16_01_02_residential"],
        "esto_flows": ["16.02 Residential"],
    },
    "Commercial and public services": {
        "ninth_sub2sectors": ["16_01_01_commercial_and_public_services"],
        "esto_flows": ["16.01 Commercial and public services"],
    },
}

CURRENT_ACCOUNT_LABELS = {"current accounts", "current account"}
EXCLUDED_FUEL_CODES = {"19_total", "20_total_renewables", "21_modern_renewables"}
EXCLUDED_ESTO_PRODUCTS = {"19 Total", "20 Total Renewables", "21 Modern renewables"}

VARIABLE_IDS = {
    "Demand Cost": 776,
    "Activity Level": 2027,
    "Total Energy": 2042,
}


def _normalize_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _repo_path(path: str | Path) -> Path:
    path = Path(str(path).replace("\\", "/"))
    return path if path.is_absolute() else (REPO_ROOT / path).resolve()


def _year_columns(columns: Iterable[object]) -> list[int]:
    years = []
    for col in columns:
        text = str(col).strip()
        if text.isdigit() and len(text) == 4:
            years.append(int(text))
    return sorted(set(years))


def _truthy(value: object) -> bool:
    return str(value).strip().lower() in {"1", "true", "t", "yes", "y"}


def _code_key(value: object) -> str:
    return _normalize_text(value).lower().replace(".", "_").replace("&", "and").replace("-", "_").replace(" ", "_")


def _clean_ninth_code_label(value: object) -> str:
    text = _normalize_text(value)
    if not text:
        return ""
    return clean_fuel_label_for_leap(text.replace("_", " "))


def _build_fuel_lookup() -> tuple[dict[str, str], dict[str, str], list[str]]:
    product_by_ninth_fuel: dict[str, str] = {}
    label_by_product: dict[str, str] = {}
    ordered_labels: list[str] = []

    for product in ESTO_PRODUCT_LIST:
        product = _normalize_text(product)
        if not product or product in EXCLUDED_ESTO_PRODUCTS:
            continue
        product_key = _code_key(product)
        product_by_ninth_fuel[product_key] = product
        label = clean_fuel_label_for_leap(product)
        label_by_product[product] = label
        if label and label not in ordered_labels:
            ordered_labels.append(label)

    return product_by_ninth_fuel, label_by_product, ordered_labels


PRODUCT_BY_NINTH_FUEL, LABEL_BY_ESTO_PRODUCT, FUEL_LABEL_ORDER = _build_fuel_lookup()


def _scenario_to_ninth_key(scenario_name: str) -> str:
    # Current 9th data uses lower-case scenario names.
    return _normalize_text(scenario_name).lower()


def _scenario_id(scenario_name: str, template_data: pd.DataFrame) -> object:
    scenario_key = _normalize_text(scenario_name).lower()
    template_rows = template_data[
        template_data.get("Scenario", pd.Series(dtype=object))
        .astype(str)
        .str.strip()
        .str.lower()
        .eq(scenario_key)
    ]
    if not template_rows.empty and "ScenarioID" in template_rows.columns:
        value = template_rows["ScenarioID"].dropna()
        if not value.empty:
            return value.iloc[0]

    for configured_name, payload in scenario_dict.items():
        if configured_name.lower() == scenario_key:
            return payload.get("scenario_id", pd.NA)
    return pd.NA


def _region_id(economy: str) -> object:
    payload = region_id_name_dict.get(economy, {})
    return payload.get("region_id", pd.NA)


def _load_ninth() -> pd.DataFrame:
    path = _repo_path(NINTH_DATA_PATH)
    if not path.exists():
        raise FileNotFoundError(f"9th data file not found: {path}")
    df = pd.read_csv(path)
    df = df.rename(columns={col: int(col) for col in df.columns if str(col).isdigit()})
    df["economy_key"] = df["economy"].apply(normalize_economy_key)
    return df


def _load_esto() -> pd.DataFrame:
    path = _repo_path(ESTO_DATA_PATH)
    if not path.exists():
        raise FileNotFoundError(f"ESTO data file not found: {path}")
    df = pd.read_csv(path)
    df = df.rename(columns={col: int(col) for col in df.columns if str(col).isdigit()})
    df["economy_key"] = df["economy"].apply(normalize_economy_key)
    df["flows"] = df["flows"].astype(str).str.strip()
    df["products"] = df["products"].astype(str).str.strip()
    return df


def _aggregate_ninth_fuel_series(
    ninth_df: pd.DataFrame,
    *,
    economy: str,
    scenario_name: str,
    sub2sectors: Sequence[str],
    years: Sequence[int],
) -> dict[str, dict[int, float]]:
    if not years:
        return {}

    subset = ninth_df[
        (ninth_df["economy_key"] == normalize_economy_key(economy))
        & (ninth_df["scenarios"].astype(str).str.strip().str.lower() == _scenario_to_ninth_key(scenario_name))
        & (ninth_df["sub1sectors"].astype(str).str.strip() == "16_01_buildings")
        & (ninth_df["sub2sectors"].astype(str).str.strip().isin(sub2sectors))
    ].copy()
    if subset.empty:
        return {}

    if "subtotal_results" in subset.columns:
        subset = subset[~subset["subtotal_results"].map(_truthy)].copy()

    year_cols = [year for year in years if year in subset.columns]
    if not year_cols:
        return {}

    subset["fuel_code"] = subset["fuels"].astype(str).str.strip()
    subset = subset[~subset["fuel_code"].isin(EXCLUDED_FUEL_CODES)].copy()
    for year in year_cols:
        subset[year] = pd.to_numeric(subset[year], errors="coerce").fillna(0.0)

    grouped = subset.groupby("fuel_code", dropna=False)[year_cols].sum()
    out: dict[str, dict[int, float]] = {}
    for fuel_code, row in grouped.iterrows():
        fuel_code = _normalize_text(fuel_code)
        if not fuel_code or fuel_code.lower() == "x":
            continue
        product = PRODUCT_BY_NINTH_FUEL.get(fuel_code.lower())
        label = LABEL_BY_ESTO_PRODUCT.get(product, "") if product else _clean_ninth_code_label(fuel_code)
        if not label:
            continue
        out[label] = {int(year): max(float(row[year]), 0.0) for year in year_cols}
    return out


def _aggregate_esto_base_values(
    esto_df: pd.DataFrame,
    *,
    economy: str,
    flows: Sequence[str],
    base_year: int,
) -> dict[str, float]:
    if not flows or base_year not in esto_df.columns:
        return {}

    subset = esto_df[
        (esto_df["economy_key"] == normalize_economy_key(economy))
        & (esto_df["flows"].isin(flows))
        & (~esto_df["products"].isin(EXCLUDED_ESTO_PRODUCTS))
    ].copy()
    if subset.empty:
        return {}

    aggregate_products = {
        product
        for product in ESTO_PRODUCT_LIST
        if product not in EXCLUDED_ESTO_PRODUCTS and "." not in product.split(" ", 1)[0]
    }
    subset = subset[subset["products"].isin(aggregate_products)].copy()
    subset[base_year] = pd.to_numeric(subset[base_year], errors="coerce").fillna(0.0)
    grouped = subset.groupby("products", dropna=False)[base_year].sum()

    out: dict[str, float] = {}
    for product, value in grouped.items():
        label = LABEL_BY_ESTO_PRODUCT.get(_normalize_text(product), clean_fuel_label_for_leap(product))
        if label:
            out[label] = max(float(value), 0.0)
    return out


def _all_relevant_fuels(
    base_values: dict[str, float],
    scenario_series: dict[str, dict[int, float]],
) -> list[str]:
    present = set(base_values) | set(scenario_series)
    ordered = [label for label in FUEL_LABEL_ORDER if label in present]
    extras = sorted(present - set(ordered), key=str.lower)
    return ordered + extras


def _make_levels(branch_path: str, columns: Sequence[str]) -> dict[str, object]:
    parts = [part for part in branch_path.split("\\") if part]
    level_cols = [col for col in columns if str(col).startswith("Level ")]
    values: dict[str, object] = {}
    for idx, col in enumerate(level_cols):
        values[col] = parts[idx] if idx < len(parts) else pd.NA
    return values


def _make_row(
    *,
    columns: Sequence[str],
    branch_path: str,
    variable: str,
    scenario: str,
    scenario_id: object,
    expression: object,
    scale: object = pd.NA,
    units: object = pd.NA,
    per: object = pd.NA,
) -> dict[str, object]:
    row = {col: pd.NA for col in columns}
    row.update(
        {
            "BranchID": pd.NA,
            "VariableID": VARIABLE_IDS.get(variable, pd.NA),
            "ScenarioID": scenario_id,
            "RegionID": _region_id(ECONOMY),
            "Branch Path": branch_path,
            "Variable": variable,
            "Scenario": scenario,
            "Region": REGION,
            "Scale": scale,
            "Units": units,
            "Per...": per,
            "Expression": expression,
        }
    )
    row.update(_make_levels(branch_path, columns))
    return row


def _add_category_rows(
    rows: list[dict[str, object]],
    *,
    columns: Sequence[str],
    branch_path: str,
    scenario: str,
    scenario_id_value: object,
    root: bool = False,
) -> None:
    rows.append(
        _make_row(
            columns=columns,
            branch_path=branch_path,
            variable="Demand Cost",
            scenario=scenario,
            scenario_id=scenario_id_value,
            expression=0,
            units="U.S. Dollar",
        )
    )
    rows.append(
        _make_row(
            columns=columns,
            branch_path=branch_path,
            variable="Activity Level",
            scenario=scenario,
            scenario_id=scenario_id_value,
            expression=0 if root else 100,
            scale=pd.NA if root else "%",
            units="No data" if root else "Share",
        )
    )


def _add_fuel_rows(
    rows: list[dict[str, object]],
    *,
    columns: Sequence[str],
    branch_path: str,
    scenario: str,
    scenario_id_value: object,
    expression: object,
) -> None:
    rows.append(
        _make_row(
            columns=columns,
            branch_path=branch_path,
            variable="Demand Cost",
            scenario=scenario,
            scenario_id=scenario_id_value,
            expression=0,
            units="U.S. Dollar",
        )
    )
    rows.append(
        _make_row(
            columns=columns,
            branch_path=branch_path,
            variable="Total Energy",
            scenario=scenario,
            scenario_id=scenario_id_value,
            expression=expression,
            units="Petajoule",
        )
    )


def build_buildings_dummy_export() -> tuple[Path, pd.DataFrame]:
    template_path = _repo_path(TEMPLATE_EXPORT_PATH)
    header_rows, template_data, columns = read_export_sheet(template_path, SHEET_NAME)
    ninth_df = _load_ninth()
    esto_df = _load_esto()

    projection_years = [
        year
        for year in range(PROJECTION_START_YEAR, PROJECTION_END_YEAR + 1)
        if year in _year_columns(ninth_df.columns)
    ]
    scenarios = [_normalize_text(scenario) for scenario in SCENARIOS if _normalize_text(scenario)]
    workbook_scenarios = (["Current Accounts"] if INCLUDE_CURRENT_ACCOUNTS else []) + scenarios

    rows: list[dict[str, object]] = []
    diagnostics: list[dict[str, object]] = []

    base_by_branch: dict[str, dict[str, float]] = {}
    scenario_by_branch: dict[tuple[str, str], dict[str, dict[int, float]]] = {}
    for branch_name, source in SOURCE_BRANCHES.items():
        base_values = _aggregate_esto_base_values(
            esto_df,
            economy=ECONOMY,
            flows=source["esto_flows"],
            base_year=BASE_YEAR,
        )
        fallback_base = _aggregate_ninth_fuel_series(
            ninth_df,
            economy=ECONOMY,
            scenario_name=scenarios[0] if scenarios else "Reference",
            sub2sectors=source["ninth_sub2sectors"],
            years=[BASE_YEAR],
        )
        for label, series in fallback_base.items():
            base_values.setdefault(label, float(series.get(BASE_YEAR, 0.0)))
        base_by_branch[branch_name] = base_values

        for scenario in scenarios:
            scenario_by_branch[(branch_name, scenario)] = _aggregate_ninth_fuel_series(
                ninth_df,
                economy=ECONOMY,
                scenario_name=scenario,
                sub2sectors=source["ninth_sub2sectors"],
                years=projection_years,
            )

    for scenario in workbook_scenarios:
        scenario_id_value = _scenario_id(scenario, template_data)
        _add_category_rows(
            rows,
            columns=columns,
            branch_path=ROOT_BRANCH,
            scenario=scenario,
            scenario_id_value=scenario_id_value,
            root=True,
        )
        for branch_name in SOURCE_BRANCHES:
            category_path = f"{ROOT_BRANCH}\\{branch_name}"
            _add_category_rows(
                rows,
                columns=columns,
                branch_path=category_path,
                scenario=scenario,
                scenario_id_value=scenario_id_value,
            )

            base_values = base_by_branch.get(branch_name, {})
            scenario_series = (
                {}
                if scenario.lower() in CURRENT_ACCOUNT_LABELS
                else scenario_by_branch.get((branch_name, scenario), {})
            )
            for fuel_label in _all_relevant_fuels(base_values, scenario_series):
                fuel_path = f"{category_path}\\{fuel_label}"
                if scenario.lower() in CURRENT_ACCOUNT_LABELS:
                    value = float(base_values.get(fuel_label, 0.0))
                    expression = build_data_expression({BASE_YEAR: value})
                else:
                    series = scenario_series.get(fuel_label)
                    if series is None:
                        series = {year: 0.0 for year in projection_years}
                    expression = build_data_expression(series)
                _add_fuel_rows(
                    rows,
                    columns=columns,
                    branch_path=fuel_path,
                    scenario=scenario,
                    scenario_id_value=scenario_id_value,
                    expression=expression,
                )
                diagnostics.append(
                    {
                        "branch": branch_name,
                        "fuel": fuel_label,
                        "scenario": scenario,
                        "source": "ESTO base-year with 9th fallback"
                        if scenario.lower() in CURRENT_ACCOUNT_LABELS
                        else "9th projection",
                        "expression": expression,
                    }
                )

    output = pd.DataFrame(rows).reindex(columns=columns)
    output_path = _repo_path(OUTPUT_EXPORT_PATH)
    write_export_sheet(
        path=output_path,
        sheet_name=SHEET_NAME,
        header_rows=header_rows,
        columns=columns,
        data=output,
    )

    diagnostics_df = pd.DataFrame(diagnostics)
    diagnostics_path = _repo_path(DIAGNOSTICS_PATH)
    diagnostics_path.parent.mkdir(parents=True, exist_ok=True)
    diagnostics_df.to_csv(diagnostics_path, index=False)

    return output_path, diagnostics_df


def main() -> None:
    output_path, diagnostics_df = build_buildings_dummy_export()
    branch_count = diagnostics_df[["branch", "fuel"]].drop_duplicates().shape[0]
    print(f"[INFO] Wrote buildings dummy export: {output_path}")
    print(f"[INFO] Wrote diagnostics: {_repo_path(DIAGNOSTICS_PATH)}")
    print(f"[INFO] Fuel branches generated: {branch_count}")


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True, style: str | None = None) -> None:  # noqa: ARG001
        return

#%%
if __name__ == "__main__":  # pragma: no cover
    main()
    _emit_completion_beep(success=True, style="chime")
#%%
