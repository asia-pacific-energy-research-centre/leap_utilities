#%%
"""
Draft transfer analysis scaffolding.

Purpose:
- Treat ESTO 08.* Transfers flows as Transformation-style processes for LEAP.
- Build process_records compatible with transformation exports.
- Keep logic isolated (no edits to existing transformation modules).

Notes:
- Inputs are negative, outputs are positive in balance tables.
- Prefer subflows (08.01/08.02/08.03) when they have nonzero data; fallback to 08 Transfers.
- Transfers are economy-specific: update TRANSFER_PROCESS_CONFIG with explicit mappings.
- Subtotals are dropped before any transfer logic runs.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Iterable, Sequence

import pandas as pd

# Allow the repository root to be importable regardless of the working directory.
REPO_ROOT = Path(__file__).resolve().parents[1]
CURRENT_DIR = Path.cwd()
if CURRENT_DIR != REPO_ROOT:
    os.chdir(REPO_ROOT)
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from leap_utils import transformation_analysis_utils as core
from leap_utils.leap_core import (
    connect_to_leap,
    create_branches_from_export_file,
    fill_branches_from_export_file,
    is_leap_api_available,
)
from leap_utils.config import (
    BRANCH_DEMAND_CATEGORY,
    BRANCH_DEMAND_TECHNOLOGY,
)

# --- Configuration ---
TRANSFER_FLOW_CODES = [
    "08 Transfers",
    "08.01 Recycled products",
    "08.02 Interproduct transfers",
    "08.03 Products transferred",
    "08.99 Transformation nonspecified"
]

# Prefer subflows when they have nonzero data.
TRANSFER_SUBFLOWS = [
    "08.01 Recycled products",
    "08.02 Interproduct transfers",
    "08.03 Products transferred",
    "08.99 Transformation nonspecified"
]

# If True, filter subtotal rows immediately before transfer calculations.
DROP_SUBTOTALS_FIRST = True

# Category templates that help organize transfers when per-economy mappings are missing.
# These are broad, optional groupings based on the requested breakdowns.
TRANSFER_CATEGORY_TEMPLATES = [
    {
        "category": "Upstream liquids transfers",
        "inputs": [
            "08.01 Natural gas",
            "06.02 Natural gas liquids",
            "06.01 Crude oil",
            "06 Crude oil & NGL",
            "06.05 Other hydrocarbons",
        ],
        "outputs": [
            "07.09 LPG",
            "07.11 Ethane",
            "06.05 Other hydrocarbons",
        ],
    },
    {
        "category": "Refinery & blending transfers",
        "inputs": [
            "06.04 Additives/  oxygenates",
            "07.03 Naphtha",
            "07 Petroleum products",
            "07.17 Other products",
            "07.02 Aviation gasoline",
            "07.12 White spirit SBP",
            "07.13 Lubricants",
            "07.15 Paraffin  waxes",
            "07.08 Fuel oil",
            "07.06 Kerosene",
            "07.07 Gas/diesel oil",
            "07.14 Bitumen",
            "07.05 Kerosene type jet fuel",
            "07.09 LPG",
            "07.01 Motor gasoline",
            "07.16 Petroleum coke",
            "07.10 Refinery gas (not liquefied)",
        ],
        "outputs": [
            "07.10 Refinery gas (not liquefied)",
            "07.13 Lubricants",
            "07.16 Petroleum coke",
            "07.02 Aviation gasoline",
            "07.10 Refinery gas (not liquefied)",
            "07.16 Petroleum coke",
            "07.01 Motor gasoline",
            "07.07 Gas/diesel oil",
            "07.05 Kerosene type jet fuel",
            "07.06 Kerosene",
            "07.08 Fuel oil",
            "07.14 Bitumen",
            "06.03 Refinery feedstocks",
            "07.03 Naphtha",
            "07.17 Other products",
            "07.15 Paraffin  waxes",
            "07.12 White spirit SBP",
        ],
    },
    {
        "category": "Others",
        "inputs": [],
        "outputs": [],
        "mode": "others",
    },
]

# Economy-specific mapping. Each entry is a list of process configs per flow.
# Replace these placeholders with real transfer groupings per economy.
# Note: When TRANSFER_CATEGORY_TEMPLATES changes, re-run
# `leap_utils/scrapbook/transfers_mapping_exploration.py` and paste the printed
# TRANSFER_PROCESS_CONFIG output here so categories stay aligned.
TRANSFER_PROCESS_CONFIG: dict[str, dict[str, list[dict]]] = {
    "00_APEC": {
        "transfer_flows_combined": [
            {
                "process": "Upstream liquids transfers",
                "inputs": [
                    "06.02 Natural gas liquids",
                    "06.05 Other hydrocarbons"
                ],
                "outputs": [
                    "06.01 Crude oil",
                    "07.09 LPG",
                    "07.11 Ethane"
                ]
            },
            {
                "process": "Refinery & blending transfers",
                "inputs": [
                    "06.04 Additives/  oxygenates",
                    "07.03 Naphtha",
                    "07.05 Kerosene type jet fuel",
                    "07.06 Kerosene",
                    "07.08 Fuel oil",
                    "07.12 White spirit SBP",
                    "07.14 Bitumen",
                    "07.15 Paraffin  waxes",
                    "07.17 Other products"
                ],
                "outputs": [
                    "07.01 Motor gasoline",
                    "07.02 Aviation gasoline",
                    "07.07 Gas/diesel oil",
                    "06.03 Refinery feedstocks",
                    "07.10 Refinery gas (not liquefied)",
                    "07.13 Lubricants",
                    "07.16 Petroleum coke"
                ]
            }
        ]
    },
    "01_AUS": {
        "transfer_flows_combined": [
            {
                "process": "Upstream & refinery transfers",
                "inputs": [
                    "06.02 Natural gas liquids"
                ],
                "outputs": [
                    "06.01 Crude oil",
                    "06.03 Refinery feedstocks",
                    "07.09 LPG",
                    "07.11 Ethane",
                    "07.17 Other products"
                ]
            }
        ]
    },
    "02_BD": {
        "transfer_flows_combined": [
            {
                "process": "Refinery & blending transfers",
                "inputs": [
                    "07.01 Motor gasoline",
                    "07.03 Naphtha"
                ],
                "outputs": [
                    "06.03 Refinery feedstocks",
                    "07.17 Other products"
                ]
            }
        ]
    },
    "03_CDA": {
        "transfer_flows_combined": [
            {
                "process": "Upstream liquids transfers",
                "inputs": [
                    "06.02 Natural gas liquids",
                    "06.05 Other hydrocarbons"
                ],
                "outputs": [
                    "07.09 LPG",
                    "07.11 Ethane"
                ]
            },
            {
                "process": "Refinery & blending transfers",
                "inputs": [
                    "06.04 Additives/  oxygenates",
                    "07.02 Aviation gasoline",
                    "07.03 Naphtha",
                    "07.05 Kerosene type jet fuel",
                    "07.08 Fuel oil",
                    "07.12 White spirit SBP",
                    "07.14 Bitumen",
                    "07.17 Other products"
                ],
                "outputs": [
                    "07.01 Motor gasoline",
                    "07.07 Gas/diesel oil",
                    "06.03 Refinery feedstocks",
                    "07.10 Refinery gas (not liquefied)",
                    "07.13 Lubricants",
                    "07.16 Petroleum coke"
                ]
            }
        ]
    },
    "04_CHL": {
        "transfer_flows_combined": [
            {
                "process": "Upstream & refinery transfers",
                "inputs": [
                    "06.02 Natural gas liquids",
                    "07.01 Motor gasoline",
                    "07.02 Aviation gasoline",
                    "07.03 Naphtha",
                    "07.05 Kerosene type jet fuel",
                    "07.06 Kerosene",
                    "07.07 Gas/diesel oil",
                    "07.08 Fuel oil",
                    "07.09 LPG",
                    "07.17 Other products"
                ],
                "outputs": [
                    "06.03 Refinery feedstocks"
                ]
            }
        ]
    },
    "08_JPN": {
        "transfer_flows_combined": [
            {
                "process": "Upstream & refinery transfers",
                "inputs": [
                    "06.05 Other hydrocarbons",
                    "07.05 Kerosene type jet fuel",
                    "07.06 Kerosene",
                    "07.08 Fuel oil",
                    "07.09 LPG",
                    "07.13 Lubricants",
                    "07.14 Bitumen",
                    "07.15 Paraffin  waxes",
                    "07.16 Petroleum coke"
                ],
                "outputs": [
                    "07.01 Motor gasoline",
                    "07.03 Naphtha",
                    "07.07 Gas/diesel oil",
                    "07.17 Other products"
                ]
            }
        ]
    },
    "09_ROK": {
        "transfer_flows_combined": [
            {
                "process": "Upstream & refinery transfers",
                "inputs": [
                    "06.04 Additives/  oxygenates",
                    "07.03 Naphtha",
                    "07.06 Kerosene",
                    "07.08 Fuel oil",
                    "07.12 White spirit SBP",
                    "07.13 Lubricants",
                    "07.15 Paraffin  waxes",
                    "07.17 Other products"
                ],
                "outputs": [
                    "06.03 Refinery feedstocks",
                    "07.01 Motor gasoline",
                    "07.02 Aviation gasoline",
                    "07.05 Kerosene type jet fuel",
                    "07.07 Gas/diesel oil",
                    "07.09 LPG",
                    "07.10 Refinery gas (not liquefied)",
                    "07.14 Bitumen",
                    "07.16 Petroleum coke"
                ]
            }
        ]
    },
    "11_MEX": {
        "transfer_flows_combined": [
            {
                "process": "Upstream liquids transfers",
                "inputs": [
                    "06.02 Natural gas liquids"
                ],
                "outputs": [
                    
                    "07.09 LPG",
                    "07.11 Ethane"
                ]
            },
            {
                "process": "Refinery & blending transfers",
                "inputs": [
                    "07.03 Naphtha"
                ],
                "outputs": [
                    "06.03 Refinery feedstocks",
                ]
            }
        ]
    },
    "12_NZ": {
        "transfer_flows_combined": [
            {
                "process": "Upstream liquids transfers",
                "inputs": [
                    "06.02 Natural gas liquids"
                ],
                "outputs": [
                    "07.09 LPG"
                ]
            },
            {
                "process": "Refinery & blending transfers",
                "inputs": [
                    "07.01 Motor gasoline",
                    "07.05 Kerosene type jet fuel",
                    "07.07 Gas/diesel oil",
                    "07.08 Fuel oil",
                    "07.14 Bitumen",
                    "07.17 Other products"
                ],
                "outputs": [
                    "06.03 Refinery feedstocks",
                ]
            }
        ]
    },
    "13_PNG": {
        "transfer_flows_combined": [
            {
                "process": "Upstream & refinery transfers",
                "inputs": [
                    "07.03 Naphtha",
                    "07.06 Kerosene"
                ],
                "outputs": [
                    "07.01 Motor gasoline",
                    "07.05 Kerosene type jet fuel"
                ]
            }
        ]
    },
    "14_PE": {
        "transfer_flows_combined": [
            {
                "process": "Upstream liquids transfers",
                "inputs": [
                    "06.02 Natural gas liquids"
                ],
                "outputs": [
                    "07.09 LPG"
                ]
            },
            {
                "process": "Refinery & blending transfers",
                "inputs": [
                    "07.05 Kerosene type jet fuel",
                    "07.07 Gas/diesel oil",
                    "07.08 Fuel oil"
                ],
                "outputs": [
                    "07.01 Motor gasoline",
                    "07.03 Naphtha",
                    "07.06 Kerosene",
                    "06.03 Refinery feedstocks",
                ]
            }
        ]
    },
    "18_CT": {
        "transfer_flows_combined": [
            {
                "process": "Upstream & refinery transfers",
                "inputs": [
                    "07.03 Naphtha",
                    "07.05 Kerosene type jet fuel",
                    "07.06 Kerosene",
                    "07.07 Gas/diesel oil",
                    "07.08 Fuel oil",
                    "06.03 Refinery feedstocks"
                    "07.12 White spirit SBP",
                    "07.13 Lubricants"
                    "07.09 LPG"
                ],
                "outputs": [
                    "06.04 Additives/  oxygenates",
                    "07.01 Motor gasoline",
                    "07.17 Other products"
                ]
            }
        ]
    },
    "20_USA": {
        "transfer_flows_combined": [
            {
                "process": "Upstream liquids transfers",
                "inputs": [
                    "06.02 Natural gas liquids"
                ],
                "outputs": [
                    "07.09 LPG",
                    "07.11 Ethane"
                ]
            },
            {
                "process": "Refinery & blending transfers",
                "inputs": [
                    "06.04 Additives/  oxygenates",
                    "07.02 Aviation gasoline",
                    "07.06 Kerosene",
                    "07.08 Fuel oil"
                ],
                "outputs": [
                    "07.01 Motor gasoline",
                    "07.05 Kerosene type jet fuel",
                    "07.07 Gas/diesel oil",
                    "06.03 Refinery feedstocks",
                    "07.14 Bitumen",
                    "07.17 Other products"
                ]
            }
        ]
    },
    "21_VN": {
        "transfer_flows_combined": [
            {
                "process": "Upstream & refinery transfers",
                "inputs": [
                    "06.02 Natural gas liquids"
                ],
                "outputs": [
                    "07.09 LPG"
                ]
            }
        ]
    }
}

TRANSFER_ECONOMY_CONFIG_ALIASES = {
    "ALL_ECONOMIES": "00_APEC",
}


def _sum_series(series_list: Iterable[pd.Series]) -> pd.Series:
    """Sum a list of pandas Series, aligning indices and filling missing with 0."""
    total = None
    for series in series_list:
        if series is None or series.empty:
            continue
        total = series if total is None else total.add(series, fill_value=0.0)
    return total if total is not None else pd.Series(dtype=float)


def _flow_has_nonzero(flow_rows: pd.DataFrame, year_cols: list[int]) -> bool:
    """Return True if any nonzero value exists in the flow rows."""
    if flow_rows.empty:
        return False
    return (flow_rows[year_cols] != 0).any().any()

def _combine_flow_rows(
    data: pd.DataFrame, economy: str, flow_codes: Iterable[str]
) -> pd.DataFrame:
    """Return concatenated rows for the requested flows."""
    frames = [
        core.select_flow_rows(data, economy, flow_code) for flow_code in flow_codes
    ]
    frames = [frame for frame in frames if frame is not None and not frame.empty]
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def select_transfer_flows(
    data: pd.DataFrame, year_cols: list[int], economy: str
) -> list[str]:
    """Prefer subflows when they have data; fallback to aggregate."""
    subflow_hits = []
    for flow_code in TRANSFER_SUBFLOWS:
        rows = core.select_flow_rows(data, economy, flow_code)
        if _flow_has_nonzero(rows, year_cols):
            subflow_hits.append(flow_code)
    if subflow_hits:
        return subflow_hits
    aggregate_rows = core.select_flow_rows(data, economy, "08 Transfers")
    if _flow_has_nonzero(aggregate_rows, year_cols):
        return ["08 Transfers"]
    return []

def _resolve_transfer_io_labels(
    process_config: dict,
    totals: pd.Series,
) -> tuple[list[str], list[str]]:
    """Assign labels to inputs/outputs based on sign in totals."""
    label_keys = ("inputs", "outputs", "products", "fuels", "labels")
    labels: list[str] = []
    for key in label_keys:
        values = process_config.get(key, [])
        if not values:
            continue
        for value in values:
            label = str(value).strip()
            if label:
                labels.append(label)
    if not labels:
        return [], []
    seen = set()
    unique_labels = [label for label in labels if not (label in seen or seen.add(label))]
    inputs = [label for label in unique_labels if totals.get(label, 0.0) < 0]
    outputs = [label for label in unique_labels if totals.get(label, 0.0) > 0]
    return inputs, outputs


def _normalize_transfer_process_name(process_config: dict, flow_code: str) -> str:
    """Return a standardized process name aligned to the three transfer categories."""
    raw = (
        process_config.get("category")
        or process_config.get("process")
        or flow_code
    )
    text = str(raw).strip()
    lowered = text.lower()
    if "upstream" in lowered and ("refinery" in lowered or "blending" in lowered):
        return TRANSFER_PROCESS_NAMES["upstream_and_refinery"]
    if "upstream" in lowered:
        return TRANSFER_PROCESS_NAMES["upstream_liquids"]
    if "refinery" in lowered or "blending" in lowered:
        return TRANSFER_PROCESS_NAMES["refinery_blending"]
    return text


def _build_template_processes(
    flow_rows: pd.DataFrame,
    year_cols: list[int],
    start_year: int,
) -> list[dict]:
    """Create process configs from category templates using nonzero products."""
    totals, _ = core.summarize_fuel_totals(
        flow_rows, year_cols, start_year, allow_all_years_fallback=True
    )
    processes: list[dict] = []
    matched_inputs: set[str] = set()
    matched_outputs: set[str] = set()
    for template in TRANSFER_CATEGORY_TEMPLATES:
        if template.get("mode") == "others":
            continue
        inputs = [
            label for label in template["inputs"] if totals.get(label, 0.0) < 0
        ]
        outputs = [
            label for label in template["outputs"] if totals.get(label, 0.0) > 0
        ]
        if not inputs or not outputs:
            continue
        matched_inputs.update(inputs)
        matched_outputs.update(outputs)
        processes.append(
            {
                "process": template["category"],
                "category": template["category"],
                "inputs": inputs,
                "outputs": outputs,
            }
        )
    others_template = next(
        (template for template in TRANSFER_CATEGORY_TEMPLATES if template.get("mode") == "others"),
        None,
    )
    if others_template is not None:
        other_inputs = [
            label
            for label, value in totals.items()
            if value < 0 and label not in matched_inputs
        ]
        other_outputs = [
            label
            for label, value in totals.items()
            if value > 0 and label not in matched_outputs
        ]
        if other_inputs and other_outputs:
            processes.append(
                {
                    "process": others_template["category"],
                    "category": others_template["category"],
                    "inputs": other_inputs,
                    "outputs": other_outputs,
                }
            )
    return processes


def _build_process_records_for_mapping(
    flow_rows: pd.DataFrame,
    year_cols: list[int],
    start_year: int,
    economy: str,
    flow_code: str,
    process_config: dict,
    sector_title: str,
    use_output_targets: bool = False,
) -> list[dict]:
    """Build process records for a configured transfer mapping."""
    timeseries, _ = core.summarize_fuel_timeseries(
        flow_rows, year_cols, start_year, allow_all_years_fallback=True
    )
    totals, _ = core.summarize_fuel_totals(
        flow_rows, year_cols, start_year, allow_all_years_fallback=True
    )
    input_labels, output_labels = _resolve_transfer_io_labels(process_config, totals)
    if not input_labels or not output_labels:
        return []

    output_series_map = {
        label: core.ensure_full_year_series(
            core.get_label_timeseries(timeseries, label),
            core.EXPORT_BASE_YEAR,
            core.EXPORT_FINAL_YEAR,
        )
        for label in output_labels
    }
    input_series_map = {
        label: core.ensure_full_year_series(
            core.get_label_timeseries(timeseries, label).abs(),
            core.EXPORT_BASE_YEAR,
            core.EXPORT_FINAL_YEAR,
        )
        for label in input_labels
    }
    total_output = _sum_series(output_series_map.values())
    total_input = _sum_series(input_series_map.values())

    if total_output.empty or total_input.empty:
        return []

    efficiency_series = core.safe_divide_series(total_output, total_input)
    feedstock_shares = {
        label: core.safe_divide_series(series, total_input).to_dict()
        for label, series in input_series_map.items()
    }
    feedstock_values = {
        label: core.series_to_year_dict(series, core.EXPORT_BASE_YEAR, core.EXPORT_FINAL_YEAR)
        for label, series in input_series_map.items()
    }
    output_values = {
        label: core.series_to_year_dict(series, core.EXPORT_BASE_YEAR, core.EXPORT_FINAL_YEAR)
        for label, series in output_series_map.items()
    }
    output_import_targets: dict = {}
    output_export_targets: dict = {}
    if use_output_targets:
        output_import_targets, output_export_targets = core.gather_output_target_dicts(
            economy,
            list(output_series_map.keys()),
            core.EXPORT_BASE_YEAR,
            core.EXPORT_FINAL_YEAR,
        )
        zero_target = core.build_value_by_year(0.0, core.EXPORT_BASE_YEAR, core.EXPORT_FINAL_YEAR)
        for label in output_series_map.keys():
            if label not in output_import_targets:
                output_import_targets[label] = dict(zero_target)
            if label not in output_export_targets:
                output_export_targets[label] = dict(zero_target)
    process_name = _normalize_transfer_process_name(process_config, flow_code)
    record = core.build_process_record(
        economy,
        sector_title,
        process_name,
        output_values,
        feedstock_values,
        core.series_to_year_dict(
            efficiency_series, core.EXPORT_BASE_YEAR, core.EXPORT_FINAL_YEAR
        ),
        auxiliary_ratios={},
        loss_values={},
        loss_total=0.0,
        feedstock_shares=feedstock_shares,
        input_total=total_input.sum(),
        output_import_targets=output_import_targets,
        output_export_targets=output_export_targets,
    )
    return [record]


def build_transfer_process_records(
    economy: str,
    sector_title: str = "Transfers",
    start_year: int = core.YEAR_START_FOR_ANALYSIS,
    process_config: dict | None = None,
    use_output_targets: bool = False,
    data_override: pd.DataFrame | None = None,
    year_cols_override: list[int] | None = None,
) -> list[dict]:
    """Return transfer process records for the given economy."""
    data = data_override if data_override is not None else core.esto_data
    if DROP_SUBTOTALS_FIRST:
        data = core.filter_matt_subtotals(data)
        data = core.filter_total_energy_rows(data)
    year_cols = year_cols_override or core.esto_year_cols
    records: list[dict] = []
    flow_codes = select_transfer_flows(data, year_cols, economy)
    if not flow_codes:
        print(f"No nonzero transfer flows for {economy}.")
        return records
    config_source = process_config or TRANSFER_PROCESS_CONFIG
    economy_config = config_source.get(economy)
    if not economy_config:
        alias = TRANSFER_ECONOMY_CONFIG_ALIASES.get(economy)
        if alias:
            economy_config = config_source.get(alias, {})
    if economy_config is None:
        economy_config = {}
    handled_flows: set[str] = set()
    combined_processes = economy_config.get(TRANSFER_COMBINED_FLOW_KEY)
    if combined_processes:
        combined_rows = _combine_flow_rows(data, economy, flow_codes)
        if not combined_rows.empty:
            for process_cfg in combined_processes:
                records.extend(
                    _build_process_records_for_mapping(
                        combined_rows,
                        year_cols,
                        start_year,
                        economy,
                        TRANSFER_COMBINED_FLOW_KEY,
                        process_cfg,
                        sector_title,
                        use_output_targets=use_output_targets,
                    )
                )
            if records:
                handled_flows.update(flow_codes)
    for flow_code in flow_codes:
        if flow_code in handled_flows:
            continue
        flow_rows = core.select_flow_rows(data, economy, flow_code)
        if flow_rows.empty:
            continue
        flow_processes = economy_config.get(flow_code)
        if not flow_processes:
            flow_processes = _build_template_processes(flow_rows, year_cols, start_year)
        if not flow_processes:
            # Final fallback: treat all positives as outputs, all negatives as inputs.
            totals, _ = core.summarize_fuel_totals(
                flow_rows, year_cols, start_year, allow_all_years_fallback=True
            )
            negatives = [label for label, value in totals.items() if value < 0]
            positives = [label for label, value in totals.items() if value > 0]
            flow_processes = [
                {
                    "process": flow_code,
                    "inputs": negatives,
                    "outputs": positives,
                }
            ]
        for process_cfg in flow_processes:
            records.extend(
                _build_process_records_for_mapping(
                    flow_rows,
                    year_cols,
                    start_year,
                        economy,
                        flow_code,
                        process_cfg,
                        sector_title,
                        use_output_targets=use_output_targets,
                    )
                )
    return records


def save_transfer_export(
    process_records: list[dict],
    scenarios: list[str] | None = None,
    output_dir: str | None = None,
    filename_template: str | None = None,
) -> str | None:
    """Save a LEAP export workbook for transfer process records."""
    if not process_records:
        print("No transfer process records to export.")
        return None
    scenario_list = scenarios or list(core.SCENARIOS_TO_EXPORT)
    economy = process_records[0].get("economy", "economy")
    output_dir = output_dir or core.EXPORT_OUTPUT_DIR
    filename = (filename_template or EXPORT_FILENAME_TEMPLATE).format(
        economy=core.format_filename_segment(economy),
        scenario=core.format_filename_segment("_".join(scenario_list)),
    )
    return core.save_transformation_export(
        process_records,
        core.EXPORT_REGION,
        core.EXPORT_BASE_YEAR,
        core.EXPORT_FINAL_YEAR,
        core.code_to_name_mapping,
        output_dir,
        filename,
        core.EXPORT_MODEL_NAME,
        scenario_list,
    )

def _format_scenario_segment(scenarios: Sequence[str]) -> str:
    tokens = [core.format_filename_segment(segment) for segment in scenarios if segment]
    sanitized = "_".join(token for token in tokens if token)
    return sanitized or "scenarios"


def _build_export_filename(
    economy_label: str,
    scenarios: Sequence[str],
    template: str | None = None,
) -> str:
    template = template or EXPORT_FILENAME_TEMPLATE
    scenario_segment = _format_scenario_segment(scenarios)
    economy_segment = core.format_filename_segment(economy_label)
    try:
        return template.format(economy=economy_segment, scenario=scenario_segment)
    except Exception as exc:
        print(f"Failed to format transfer export filename: {exc}")
        return EXPORT_FILENAME_TEMPLATE.format(economy=economy_segment, scenario=scenario_segment)


def _infer_primary_economy(process_records: Sequence[dict]) -> str:
    for record in process_records:
        economy = record.get("economy")
        if economy:
            return economy
    if core.ECONOMIES_TO_ANALYZE:
        return core.ECONOMIES_TO_ANALYZE[0]
    return "economy"


def prepare_transfer_exports(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    export_output_dir: Path | str | None = None,
    filename_template: str | None = None,
    process_config: dict | None = None,
    start_year: int = core.YEAR_START_FOR_ANALYSIS,
    include_output_series: bool = False,
    use_output_targets: bool = False,
    include_all_economies: bool = False,
    aggregate_economy_label: str | None = None,
    build_export: bool = core.BUILD_LEAP_EXPORT,
) -> list[Path]:
    """Build transfer process records and emit the LEAP workbook."""
    if not build_export:
        print("BUILD_LEAP_EXPORT is False; skipping workbook generation.")
        return []
    aggregate_label = aggregate_economy_label or "ALL_ECONOMIES"
    data_override = None
    year_cols_override = None
    previous_import_export_data = None
    previous_import_export_years = None
    import_export_override = False
    if include_all_economies:
        data_override = core.add_all_economy_total(
            core.esto_data,
            core.esto_year_cols,
            aggregate_label,
        )
        year_cols_override = core.esto_year_cols
        economy_list = [aggregate_label]
    else:
        economy_list = list(economies or core.ECONOMIES_TO_ANALYZE)
    process_records: list[dict] = []
    try:
        if include_all_economies and use_output_targets:
            previous_import_export_data = core.ESTO_IMPORT_EXPORT_REFERENCE_DATA
            previous_import_export_years = core.ESTO_IMPORT_EXPORT_YEAR_COLS
            core.ESTO_IMPORT_EXPORT_REFERENCE_DATA = data_override
            core.ESTO_IMPORT_EXPORT_YEAR_COLS = year_cols_override or core.esto_year_cols
            import_export_override = True
        for economy in economy_list:
            process_records.extend(
                build_transfer_process_records(
                    economy,
                    start_year=start_year,
                    process_config=process_config,
                    use_output_targets=use_output_targets,
                    data_override=data_override,
                    year_cols_override=year_cols_override,
                )
            )
    finally:
        if import_export_override:
            core.ESTO_IMPORT_EXPORT_REFERENCE_DATA = previous_import_export_data
            core.ESTO_IMPORT_EXPORT_YEAR_COLS = previous_import_export_years
    if not process_records:
        print("No transfer records were generated; nothing to export.")
        return []
    process_records = _merge_transfer_process_records(process_records)
    _consolidate_transfer_outputs(process_records, include_output_series, use_output_targets)
    scenario_list = list(scenarios or core.SCENARIOS_TO_EXPORT)
    output_dir_path = Path(export_output_dir or core.EXPORT_OUTPUT_DIR)
    output_dir_path.mkdir(parents=True, exist_ok=True)
    economy_label = _infer_primary_economy(process_records)
    export_filename = _build_export_filename(economy_label, scenario_list, filename_template)
    previous_output_setting = core.INCLUDE_OUTPUT_SERIES_IN_LEAP_EXPORT
    previous_output_config = dict(core.TRANSFORMATION_OUTPUT_VARIABLES)
    core.INCLUDE_OUTPUT_SERIES_IN_LEAP_EXPORT = bool(include_output_series)
    core.TRANSFORMATION_OUTPUT_VARIABLES["output"] = bool(include_output_series)
    core.TRANSFORMATION_OUTPUT_VARIABLES["output_import_target"] = bool(use_output_targets)
    core.TRANSFORMATION_OUTPUT_VARIABLES["output_export_target"] = bool(use_output_targets)
    try:
        export_path = core.save_transformation_export(
            process_records,
            core.EXPORT_REGION,
            core.EXPORT_BASE_YEAR,
            core.EXPORT_FINAL_YEAR,
            core.code_to_name_mapping,
            str(output_dir_path),
            export_filename,
            core.EXPORT_MODEL_NAME,
            scenario_list,
        )
    finally:
        core.INCLUDE_OUTPUT_SERIES_IN_LEAP_EXPORT = previous_output_setting
        core.TRANSFORMATION_OUTPUT_VARIABLES = previous_output_config
    return [Path(export_path)] if export_path else []


def _sum_year_dicts(series_list: Iterable[dict]) -> dict:
    """Sum year->value dicts, aligning years."""
    totals: dict[int, float] = {}
    for series in series_list:
        if not series:
            continue
        for year, value in series.items():
            if value is None:
                continue
            totals[int(year)] = totals.get(int(year), 0.0) + float(value)
    return totals


def _consolidate_transfer_outputs(
    process_records: list[dict],
    include_output_series: bool,
    use_output_targets: bool,
) -> None:
    """Ensure transfer output values/targets are aggregated to avoid duplicates."""
    if not process_records or not (include_output_series or use_output_targets):
        return
    grouped: dict[tuple[str, str], list[dict]] = {}
    for record in process_records:
        key = (record.get("economy"), record.get("sector_title"))
        grouped.setdefault(key, []).append(record)
    for _, records in grouped.items():
        if len(records) < 2:
            continue
        output_values_by_label: dict[str, list[dict]] = {}
        import_targets_by_label: dict[str, list[dict]] = {}
        export_targets_by_label: dict[str, list[dict]] = {}
        for record in records:
            for label, values in (record.get("output_values") or {}).items():
                output_values_by_label.setdefault(label, []).append(values)
            for label, values in (record.get("output_import_targets") or {}).items():
                import_targets_by_label.setdefault(label, []).append(values)
            for label, values in (record.get("output_export_targets") or {}).items():
                export_targets_by_label.setdefault(label, []).append(values)
        aggregated_outputs = {
            label: _sum_year_dicts(values)
            for label, values in output_values_by_label.items()
            if values
        }
        aggregated_imports = {
            label: _sum_year_dicts(values)
            for label, values in import_targets_by_label.items()
            if values
        }
        aggregated_exports = {
            label: _sum_year_dicts(values)
            for label, values in export_targets_by_label.items()
            if values
        }
        carrier = records[0]
        carrier["output_values"] = aggregated_outputs if include_output_series else {}
        if use_output_targets:
            carrier["output_import_targets"] = aggregated_imports
            carrier["output_export_targets"] = aggregated_exports
        else:
            carrier["output_import_targets"] = {}
            carrier["output_export_targets"] = {}
        for record in records[1:]:
            record["output_values"] = {}
            record["output_import_targets"] = {}
            record["output_export_targets"] = {}


def _read_unique_column(export_path: Path, column: str) -> list[str]:
    for header in (2, 0):
        try:
            df = pd.read_excel(
                export_path, sheet_name=SHEET_NAME, header=header, usecols=[column]
            )
        except Exception:
            continue
        if column not in df.columns:
            continue
        seen: list[str] = []
        for value in df[column].dropna().astype(str):
            if value not in seen:
                seen.append(value)
        if seen:
            return seen
    return []


def _sum_year_dicts(series_list: Iterable[dict]) -> dict:
    """Sum year->value dicts, aligning years."""
    totals: dict[int, float] = {}
    for series in series_list:
        if not series:
            continue
        for year, value in series.items():
            if value is None:
                continue
            totals[int(year)] = totals.get(int(year), 0.0) + float(value)
    return totals


def _sum_label_series(label_map: dict[str, dict]) -> pd.Series:
    """Sum dict-of-year series across labels."""
    total = pd.Series(dtype=float)
    for series in (label_map or {}).values():
        if not series:
            continue
        total = total.add(pd.Series(series, dtype=float), fill_value=0.0)
    return total


def _merge_transfer_process_records(process_records: list[dict]) -> list[dict]:
    """Merge records that share economy/sector/process to avoid duplicate LEAP rows."""
    if not process_records:
        return process_records
    grouped: dict[tuple[str, str, str], list[dict]] = {}
    for record in process_records:
        key = (
            record.get("economy"),
            record.get("sector_title"),
            record.get("process_name"),
        )
        grouped.setdefault(key, []).append(record)
    merged_records: list[dict] = []
    for _, records in grouped.items():
        if len(records) == 1:
            merged_records.append(records[0])
            continue
        output_values_by_label: dict[str, list[dict]] = {}
        feedstock_values_by_label: dict[str, list[dict]] = {}
        import_targets_by_label: dict[str, list[dict]] = {}
        export_targets_by_label: dict[str, list[dict]] = {}
        for record in records:
            for label, values in (record.get("output_values") or {}).items():
                output_values_by_label.setdefault(label, []).append(values)
            for label, values in (record.get("feedstock_values") or {}).items():
                feedstock_values_by_label.setdefault(label, []).append(values)
            for label, values in (record.get("output_import_targets") or {}).items():
                import_targets_by_label.setdefault(label, []).append(values)
            for label, values in (record.get("output_export_targets") or {}).items():
                export_targets_by_label.setdefault(label, []).append(values)
        aggregated_outputs = {
            label: _sum_year_dicts(values)
            for label, values in output_values_by_label.items()
            if values
        }
        aggregated_feedstocks = {
            label: _sum_year_dicts(values)
            for label, values in feedstock_values_by_label.items()
            if values
        }
        aggregated_imports = {
            label: _sum_year_dicts(values)
            for label, values in import_targets_by_label.items()
            if values
        }
        aggregated_exports = {
            label: _sum_year_dicts(values)
            for label, values in export_targets_by_label.items()
            if values
        }
        total_output_series = _sum_label_series(aggregated_outputs)
        total_input_series = _sum_label_series(aggregated_feedstocks)
        efficiency_series = core.safe_divide_series(total_output_series, total_input_series)
        feedstock_shares = {
            label: core.safe_divide_series(pd.Series(series, dtype=float), total_input_series).to_dict()
            for label, series in aggregated_feedstocks.items()
        }
        carrier = dict(records[0])
        carrier["output_values"] = aggregated_outputs
        carrier["feedstock_values"] = aggregated_feedstocks
        carrier["feedstock_shares"] = feedstock_shares
        carrier["efficiency"] = core.series_to_year_dict(
            efficiency_series, core.EXPORT_BASE_YEAR, core.EXPORT_FINAL_YEAR
        )
        carrier["input_total"] = float(total_input_series.sum()) if not total_input_series.empty else 0.0
        carrier["output_import_targets"] = aggregated_imports
        carrier["output_export_targets"] = aggregated_exports
        merged_records.append(carrier)
    return merged_records


def get_available_scenarios(export_path: Path) -> list[str]:
    return _read_unique_column(export_path, "Scenario")


def ensure_region_in_export(export_path: Path, region: str) -> None:
    regions = _read_unique_column(export_path, "Region")
    if not regions:
        print(f"Warning: 'Region' column missing from {export_path.name}; skipping region check.")
        return
    if region not in regions:
        raise ValueError(
            f"Requested region '{region}' not present in {export_path.name}; available: {regions}"
        )


def locate_transfer_export(
    directory: Path | str | None = None, filename: str | None = None
) -> Path:
    directory_path = Path(directory or core.EXPORT_OUTPUT_DIR)
    if filename:
        candidate = directory_path / filename
        if candidate.exists():
            return candidate
        raise FileNotFoundError(f"Specified transfer export missing: {candidate}")
    matches = sorted(directory_path.glob(f"{EXPORT_FILENAME_PREFIX}*.xlsx"))
    if not matches:
        raise FileNotFoundError(f"No transfer exports detected in {directory_path}")
    return matches[-1]


def run_transfer_leap_import(
    export_directory: Path | str | None = None,
    filename: str | None = None,
    scenario_to_run: str | None = None,
    region: str | None = None,
    include_current_accounts: bool = True,
    create_branches: bool = True,
    fill_branches: bool = True,
    raise_on_missing_branch: bool = False,
) -> Path:
    """Connect to LEAP, create branches, and fill data from the transfer export."""
    export_path = locate_transfer_export(export_directory, filename)
    available = get_available_scenarios(export_path)
    scenario_choice = scenario_to_run or (available[0] if available else None)
    if scenario_choice and scenario_choice not in available:
        raise ValueError(
            f"Scenario '{scenario_choice}' not found in {export_path.name}; options {available}"
        )
    target_region = region or core.EXPORT_REGION
    ensure_region_in_export(export_path, target_region)

    leap_conn = connect_to_leap()
    if leap_conn is None:
        raise RuntimeError("Unable to connect to LEAP.")
    if create_branches:
        create_branches_from_export_file(
            leap_conn,
            export_path,
            sheet_name=SHEET_NAME,
            branch_root=None,
            default_branch_type=(
                BRANCH_DEMAND_CATEGORY,
                BRANCH_DEMAND_CATEGORY,
                BRANCH_DEMAND_TECHNOLOGY,
            ),
            RAISE_ERROR_ON_FAILED_BRANCH_CREATION=raise_on_missing_branch,
        )
    if fill_branches:
        fill_branches_from_export_file(
            leap_conn,
            export_path,
            sheet_name=SHEET_NAME,
            scenario=scenario_choice,
            region=target_region,
            HANDLE_CURRENT_ACCOUNTS_TOO=include_current_accounts,
        )
    return export_path


def run_transfer_pipeline(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    include_leap_import: bool = False,
    import_scenario: str | None = None,
    region: str | None = None,
    handle_current_accounts: bool = True,
    create_branches: bool = True,
    fill_branches: bool = True,
    include_all_economies: bool = False,
    aggregate_economy_label: str | None = None,
    **export_kwargs,
) -> list[Path]:
    """Run exports and optionally push the workbook into LEAP."""
    exports = prepare_transfer_exports(
        economies=economies,
        scenarios=scenarios,
        export_output_dir=export_kwargs.get("export_output_dir"),
        filename_template=export_kwargs.get("filename_template"),
        process_config=export_kwargs.get("process_config"),
        start_year=export_kwargs.get("start_year", core.YEAR_START_FOR_ANALYSIS),
        include_output_series=export_kwargs.get("include_output_series", False),
        use_output_targets=export_kwargs.get("use_output_targets", False),
        include_all_economies=include_all_economies,
        aggregate_economy_label=aggregate_economy_label,
        build_export=export_kwargs.get("build_export", core.BUILD_LEAP_EXPORT),
    )
    if not exports or not include_leap_import:
        return exports
    scenario_choice = import_scenario or (scenarios or core.SCENARIOS_TO_EXPORT)[0]
    if not LEAP_API_AVAILABLE:
        print("[INFO] LEAP API unavailable in this environment; skipping branch creation/fill.")
        return exports
    run_transfer_leap_import(
        export_directory=exports[0].parent,
        filename=exports[0].name,
        scenario_to_run=scenario_choice,
        region=region or core.EXPORT_REGION,
        include_current_accounts=handle_current_accounts,
        create_branches=create_branches,
        fill_branches=fill_branches,
    )
    return exports

#%%

EXPORT_FILENAME_TEMPLATE = "transfer_leap_imports_{economy}_{scenario}.xlsx"
EXPORT_FILENAME_PREFIX = "transfer_leap_imports_"
SHEET_NAME = "LEAP"
TRANSFER_COMBINED_FLOW_KEY = "transfer_flows_combined"
TRANSFER_PROCESS_NAMES = {
    "upstream_and_refinery": "Upstream & refinery transfers",
    "upstream_liquids": "Upstream liquids transfers",
    "refinery_blending": "Refinery & blending transfers",
}
LEAP_API_AVAILABLE = is_leap_api_available()

#%%
# Simple notebook-focused configuration block.
ECONOMIES = list(core.ECONOMIES_TO_ANALYZE)
SCENARIOS = list(core.SCENARIOS_TO_EXPORT)
INCLUDE_LEAP_IMPORT = LEAP_API_AVAILABLE
IMPORT_SCENARIO = SCENARIOS[0] if SCENARIOS else None
CURRENT_ACCOUNTS = True
INCLUDE_OUTPUT_SERIES = False
USE_OUTPUT_TARGETS = True
INCLUDE_ALL_ECONOMIES = True
AGGREGATE_ECONOMY_LABEL = "ALL_ECONOMIES"

#%%
if __name__ == "__main__":
    exports = run_transfer_pipeline(
        economies=ECONOMIES,
        scenarios=SCENARIOS,
        include_leap_import=INCLUDE_LEAP_IMPORT,
        import_scenario=IMPORT_SCENARIO,
        handle_current_accounts=CURRENT_ACCOUNTS,
        include_output_series=INCLUDE_OUTPUT_SERIES,
        use_output_targets=USE_OUTPUT_TARGETS,
        include_all_economies=INCLUDE_ALL_ECONOMIES,
        aggregate_economy_label=AGGREGATE_ECONOMY_LABEL,
    )
    if exports:
        print(f"Transfer export saved to: {exports[0]}")
#%%
