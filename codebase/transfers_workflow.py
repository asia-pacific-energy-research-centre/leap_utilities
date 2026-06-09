#%%
"""
Draft transfer analysis scaffolding.

This script converts ESTO transfer flows into LEAP transformation-style process
records and export workbooks.
It handles economy-specific transfer mappings, optional unallocated-process
fallback behavior, and import dispatch for the generated transfer workbook.

Purpose:
- Treat ESTO 08.* Transfers flows as Transformation-style processes for LEAP.
- Build process_records compatible with transformation exports.
- Keep logic isolated (no edits to existing transformation modules).

Notes:
- Inputs are negative, outputs are positive in balance tables.
- Prefer subflows (08.01/08.02/08.03) when they have nonzero data; fallback to 08 Transfers.
- Transfers are economy-specific: update TRANSFER_PROCESS_CONFIG with explicit mappings.
- Subtotals are dropped before any transfer logic runs.

Most user-editable settings live in `codebase/workflow_config.py`.
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

from codebase.functions import transformation_analysis_utils as core
from codebase.configuration import workflow_config as workflow_cfg
from codebase.functions import leap_api, leap_exports
from codebase.functions.analysis_input_write_dispatcher import (
    get_analysis_input_write_mode,
)
from codebase.configuration.config import (
    BRANCH_DEMAND_CATEGORY,
    BRANCH_DEMAND_TECHNOLOGY,
)
from codebase.utilities import workflow_common

# --- Configuration ---
TRANSFER_FLOW_CODES = [
    "08 Transfers",
    "08.01 Recycled products",
    "08.02 Interproduct transfers",
    "08.03 Products transferred",
    "08.99 Transfers nonspecified"
]

# Prefer subflows when they have nonzero data.
TRANSFER_SUBFLOWS = [
    "08.01 Recycled products",
    "08.02 Interproduct transfers",
    "08.03 Products transferred",
    "08.99 Transfers nonspecified"
]

# If True, filter subtotal rows immediately before transfer calculations.
DROP_SUBTOTALS_FIRST = True
DEFAULT_SCENARIOS = list(workflow_cfg.TRANSFERS_DEFAULT_SCENARIOS)

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
        "category": "Transfers unallocated",
        "inputs": [],
        "outputs": [],
        "mode": "others",
    },
]

# Economy-specific mapping. Each entry is a list of process configs per flow.
# Replace these placeholders with real transfer groupings per economy.
# Note: When TRANSFER_CATEGORY_TEMPLATES changes, re-run
# `codebase/scrapbook/transfers_mapping_exploration.py` and paste the printed
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
                    "07.03 Naphtha",
                    "07.06 Kerosene",
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
                    "07.03 Naphtha",
                    "07.06 Kerosene",
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
                    "07.03 Naphtha",
                    "07.06 Kerosene"
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
                    "07.03 Naphtha",
                    "07.06 Kerosene"
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
                    "06.03 Refinery feedstocks",
                    "07.12 White spirit SBP",
                    "07.13 Lubricants",
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
                    "07.03 Naphtha",
                    "07.05 Kerosene type jet fuel",
                    "07.06 Kerosene",
                    "07.07 Gas/diesel oil",
                    "06.03 Refinery feedstocks",
                    "07.14 Bitumen",
                    "07.17 Other products"
                ]
            }
        ],
        "unallocated_policy": {
            "enabled": True,
            "process_name": "Transfers unallocated",
            # Trigger merge when any configured transfer process exceeds this ratio.
            "max_efficiency_ratio": 50.0,
            # When triggered, merge all included transfer categories (not only bad rows).
            "merge_all_when_triggered": True,
        },
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
        return TRANSFER_PROCESS_NAMES["unallocated"]
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
    if not _template_processes_cover_all(totals, processes):
        return []
    return processes


def _template_processes_cover_all(
    totals: pd.Series,
    processes: list[dict],
) -> bool:
    """Return True if template processes cover all nonzero inputs/outputs."""
    if not processes:
        return False
    nonzero_inputs = {label for label, value in totals.items() if value < 0}
    nonzero_outputs = {label for label, value in totals.items() if value > 0}
    covered_inputs: set[str] = set()
    covered_outputs: set[str] = set()
    for process in processes:
        covered_inputs.update(process.get("inputs", []))
        covered_outputs.update(process.get("outputs", []))
    return nonzero_inputs.issubset(covered_inputs) and nonzero_outputs.issubset(covered_outputs)


def _build_process_records_for_mapping(
    flow_rows: pd.DataFrame,
    year_cols: list[int],
    start_year: int,
    economy: str,
    flow_code: str,
    process_config: dict,
    sector_title: str,
    use_output_targets: bool = False,
    feedstock_method: str | None = None,
) -> list[dict]:
    """Build process records for a configured transfer mapping."""
    method = core.resolve_feedstock_method(feedstock_method)
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

    output_import_targets: dict = {}
    output_export_targets: dict = {}
    if use_output_targets:
        output_import_targets, output_export_targets = core.gather_output_target_dicts(
            economy,
            list(output_series_map.keys()),
            core.EXPORT_BASE_YEAR,
            core.EXPORT_FINAL_YEAR,
            output_series_by_fuel=output_series_map,
        )
        zero_target = core.build_value_by_year(0.0, core.EXPORT_BASE_YEAR, core.EXPORT_FINAL_YEAR)
        for label in output_series_map.keys():
            if label not in output_import_targets:
                output_import_targets[label] = dict(zero_target)
            if label not in output_export_targets:
                output_export_targets[label] = dict(zero_target)

    process_name = _normalize_transfer_process_name(process_config, flow_code)

    if method == core.FEEDSTOCK_METHOD_SPLIT:
        feedstock_labels = list(input_series_map.keys())
        records: list[dict] = []
        for idx, feedstock_label in enumerate(feedstock_labels):
            input_series = input_series_map[feedstock_label]
            share_series = core.build_input_share_series(
                input_series,
                total_input,
                fallback_to_one=(idx == 0),
            )
            allocated_outputs = {
                label: series.mul(share_series, fill_value=0.0)
                for label, series in output_series_map.items()
            }
            allocated_output_total = _sum_series(allocated_outputs.values())
            efficiency_series = core.safe_divide_series(
                allocated_output_total,
                input_series,
            )
            output_values = {
                label: core.series_to_year_dict(series, core.EXPORT_BASE_YEAR, core.EXPORT_FINAL_YEAR)
                for label, series in allocated_outputs.items()
            }
            output_import_targets_split = {
                label: core.scale_year_dict_by_share(values, share_series)
                for label, values in (output_import_targets or {}).items()
            }
            output_export_targets_split = {
                label: core.scale_year_dict_by_share(values, share_series)
                for label, values in (output_export_targets or {}).items()
            }
            process_label = (
                process_name
                if len(feedstock_labels) == 1
                else f"{process_name} - {feedstock_label}"
            )
            records.append(
                core.build_process_record(
                    economy,
                    sector_title,
                    process_label,
                    output_values,
                    {
                        feedstock_label: core.series_to_year_dict(
                            input_series, core.EXPORT_BASE_YEAR, core.EXPORT_FINAL_YEAR
                        )
                    },
                    core.series_to_year_dict(
                        efficiency_series, core.EXPORT_BASE_YEAR, core.EXPORT_FINAL_YEAR
                    ),
                    auxiliary_ratios={},
                    loss_values={},
                    loss_total=0.0,
                    feedstock_shares={feedstock_label: 1.0},
                    input_total=float(input_series.sum()),
                    output_import_targets=output_import_targets_split,
                    output_export_targets=output_export_targets_split,
                )
            )
        return records

    if method == core.FEEDSTOCK_METHOD_MULTI:
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

    primary_input = max(input_series_map, key=lambda label: input_series_map[label].sum())
    primary_series = input_series_map[primary_input]
    other_feedstocks = [label for label in input_series_map if label != primary_input]
    auxiliary_ratios = core.build_auxiliary_ratios_by_year(
        timeseries,
        other_feedstocks,
        total_output,
    )
    efficiency_series = core.safe_divide_series(total_output, primary_series)
    record = core.build_process_record(
        economy,
        sector_title,
        process_name,
        {
            label: core.series_to_year_dict(series, core.EXPORT_BASE_YEAR, core.EXPORT_FINAL_YEAR)
            for label, series in output_series_map.items()
        },
        {
            primary_input: core.series_to_year_dict(
                primary_series, core.EXPORT_BASE_YEAR, core.EXPORT_FINAL_YEAR
            )
        },
        core.series_to_year_dict(
            efficiency_series, core.EXPORT_BASE_YEAR, core.EXPORT_FINAL_YEAR
        ),
        auxiliary_ratios=auxiliary_ratios,
        loss_values={},
        loss_total=0.0,
        feedstock_shares={primary_input: 1.0},
        input_total=float(primary_series.sum()),
        output_import_targets=output_import_targets,
        output_export_targets=output_export_targets,
    )
    return [record]


def _sum_label_series_dict(label_map: dict[str, dict]) -> pd.Series:
    """Return total year series across label->year maps."""
    total = pd.Series(dtype=float)
    for values in (label_map or {}).values():
        if not values:
            continue
        total = total.add(pd.Series(values, dtype=float), fill_value=0.0)
    return total


def _max_efficiency_ratio(record: dict) -> float:
    """Return the maximum efficiency ratio found in a process record."""
    efficiency_map = record.get("efficiency")
    if not isinstance(efficiency_map, dict) or not efficiency_map:
        return 0.0
    ratios = [float(value) for value in efficiency_map.values() if value is not None]
    if not ratios:
        return 0.0
    return float(max(ratios))


def _normalized_name_set(values: Iterable[object] | None) -> set[str]:
    """Return normalized lowercase process-name tokens."""
    if not values:
        return set()
    out: set[str] = set()
    for value in values:
        token = str(value or "").strip().lower()
        if token:
            out.add(token)
    return out


def _apply_unallocated_policy(
    records: list[dict],
    policy: dict | None,
) -> list[dict]:
    """Collapse selected transfer rows into one unallocated process when triggered."""
    if not records:
        return records
    if not isinstance(policy, dict) or not policy:
        return records
    if not bool(policy.get("enabled", False)):
        return records

    include_names = _normalized_name_set(policy.get("include_processes"))
    exclude_names = _normalized_name_set(policy.get("exclude_processes"))

    def _is_included(record: dict) -> bool:
        name = str(record.get("process_name") or "").strip().lower()
        if not name:
            return False
        if include_names and name not in include_names:
            return False
        if exclude_names and name in exclude_names:
            return False
        return True

    candidate_records = [record for record in records if _is_included(record)]
    if not candidate_records:
        return records

    max_efficiency_ratio = policy.get("max_efficiency_ratio")
    max_efficiency_limit = (
        float(max_efficiency_ratio) if max_efficiency_ratio is not None else None
    )
    min_input_total = policy.get("min_input_total")
    min_input_limit = float(min_input_total) if min_input_total is not None else None

    bad_records: list[dict] = []
    for record in candidate_records:
        record_is_bad = False
        if max_efficiency_limit is not None and _max_efficiency_ratio(record) > max_efficiency_limit:
            record_is_bad = True
        if min_input_limit is not None:
            input_total = float(record.get("input_total") or 0.0)
            if input_total < min_input_limit:
                record_is_bad = True
        if record_is_bad:
            bad_records.append(record)
    if not bad_records:
        return records

    merge_all = bool(policy.get("merge_all_when_triggered", True))
    merge_targets = candidate_records if merge_all else bad_records
    if not merge_targets:
        return records

    output_values_by_label: dict[str, list[dict]] = {}
    feedstock_values_by_label: dict[str, list[dict]] = {}
    import_targets_by_label: dict[str, list[dict]] = {}
    export_targets_by_label: dict[str, list[dict]] = {}
    for record in merge_targets:
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

    total_output_series = _sum_label_series_dict(aggregated_outputs)
    total_input_series = _sum_label_series_dict(aggregated_feedstocks)
    efficiency_series = core.safe_divide_series(total_output_series, total_input_series)
    feedstock_shares = {
        label: core.safe_divide_series(pd.Series(series, dtype=float), total_input_series).to_dict()
        for label, series in aggregated_feedstocks.items()
    }

    carrier = dict(merge_targets[0])
    carrier["process_name"] = str(policy.get("process_name") or "Transfers unallocated")
    carrier["sector_title"] = carrier["process_name"]
    carrier["output_values"] = aggregated_outputs
    carrier["feedstock_values"] = aggregated_feedstocks
    carrier["feedstock_shares"] = feedstock_shares
    carrier["efficiency"] = core.series_to_year_dict(
        efficiency_series,
        core.EXPORT_BASE_YEAR,
        core.EXPORT_FINAL_YEAR,
    )
    carrier["input_total"] = (
        float(total_input_series.sum()) if not total_input_series.empty else 0.0
    )
    carrier["output_import_targets"] = aggregated_imports
    carrier["output_export_targets"] = aggregated_exports

    merged_target_ids = {id(record) for record in merge_targets}
    output_rows: list[dict] = [record for record in records if id(record) not in merged_target_ids]
    output_rows.append(carrier)
    return output_rows


def build_transfer_rows(
    economy: str,
    sector_title: str = "Transfers",
    start_year: int = core.YEAR_START_FOR_ANALYSIS,
    process_config: dict | None = None,
    use_output_targets: bool = False,
    feedstock_method: str | None = None,
    data_override: pd.DataFrame | None = None,
    year_cols_override: list[int] | None = None,
) -> list[dict]:
    """Return transfer rows for the given economy."""
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
    unallocated_policy = economy_config.get("unallocated_policy", DEFAULT_TRANSFER_UNALLOCATED_POLICY)

    def _sector_title_for_process(process_cfg: dict, fallback_flow_code: str) -> str:
        if not SPLIT_TRANSFER_SECTORS:
            return str(sector_title)
        return _normalize_transfer_process_name(process_cfg, fallback_flow_code)
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
                        _sector_title_for_process(process_cfg, TRANSFER_COMBINED_FLOW_KEY),
                        use_output_targets=use_output_targets,
                        feedstock_method=feedstock_method,
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
                    "process": TRANSFER_PROCESS_NAMES["unallocated"],
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
                        _sector_title_for_process(process_cfg, flow_code),
                        use_output_targets=use_output_targets,
                        feedstock_method=feedstock_method,
                    )
                )
    return _apply_unallocated_policy(records, unallocated_policy)


def save_transfer_export(
    process_records: list[dict],
    scenarios: list[str] | None = None,
    output_dir: str | None = None,
    filename_template: str | None = None,
) -> str | None:
    """Save a LEAP export workbook for transfer process records."""
    if not process_records:
        print("No transfer rows to export.")
        return None
    scenario_list = workflow_common.normalize_workflow_scenarios(
        scenarios,
        DEFAULT_SCENARIOS,
    )
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

def format_export_filename(
    economy_label: str,
    scenarios: Sequence[str],
    template: str | None = None,
) -> str:
    template = template or EXPORT_FILENAME_TEMPLATE
    return leap_exports.build_workbook_filename(
        economy_label=economy_label,
        scenarios=scenarios,
        template=template,
        fallback_template=EXPORT_FILENAME_TEMPLATE,
    )


def _infer_primary_economy(rows: Sequence[dict]) -> str:
    for row in rows:
        economy = row.get("economy")
        if economy:
            return economy
    if core.ECONOMIES_TO_ANALYZE:
        return core.ECONOMIES_TO_ANALYZE[0]
    return "economy"


def assemble_transfer_workbook(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    export_output_dir: Path | str | None = None,
    filename_template: str | None = None,
    process_config: dict | None = None,
    start_year: int = core.YEAR_START_FOR_ANALYSIS,
    include_output_series: bool = False,
    use_output_targets: bool = False,
    feedstock_method: str | None = None,
    aggregate_economy_label: str | None = None,
    build_export: bool = core.BUILD_LEAP_EXPORT,
) -> list[Path]:
    """Build transfer rows and emit the LEAP workbook."""
    if not build_export:
        print("BUILD_LEAP_EXPORT is False; skipping workbook generation.")
        return []
    economy_list = workflow_common.normalize_economies(economies or core.ECONOMIES_TO_ANALYZE)
    should_aggregate, aggregate_label, _ = workflow_common.resolve_aggregate_economy(
        economy_list,
        aggregate_label=aggregate_economy_label or workflow_cfg.TRANSFERS_AGGREGATE_ECONOMY_LABEL,
    )
    data_override = None
    year_cols_override = None
    previous_import_export_data = None
    previous_import_export_years = None
    import_export_override = False
    if should_aggregate:
        data_override = core.add_all_economy_total(
            core.esto_data,
            core.esto_year_cols,
            aggregate_label,
        )
        year_cols_override = core.esto_year_cols
        economy_list = [aggregate_label]
    rows: list[dict] = []
    original_feedstock_method = core.FEEDSTOCK_METHOD
    if feedstock_method is not None:
        core.FEEDSTOCK_METHOD = core.resolve_feedstock_method(feedstock_method)
    try:
        if should_aggregate and use_output_targets:
            previous_import_export_data = core.ESTO_IMPORT_EXPORT_REFERENCE_DATA
            previous_import_export_years = core.ESTO_IMPORT_EXPORT_YEAR_COLS
            core.ESTO_IMPORT_EXPORT_REFERENCE_DATA = data_override
            core.ESTO_IMPORT_EXPORT_YEAR_COLS = year_cols_override or core.esto_year_cols
            import_export_override = True
        for economy in economy_list:
            rows.extend(
                build_transfer_rows(
                    economy,
                    start_year=start_year,
                    process_config=process_config,
                    use_output_targets=use_output_targets,
                    feedstock_method=core.FEEDSTOCK_METHOD,
                    data_override=data_override,
                    year_cols_override=year_cols_override,
                )
            )
    finally:
        if import_export_override:
            core.ESTO_IMPORT_EXPORT_REFERENCE_DATA = previous_import_export_data
            core.ESTO_IMPORT_EXPORT_YEAR_COLS = previous_import_export_years
        core.FEEDSTOCK_METHOD = original_feedstock_method
    if not rows:
        print("No transfer rows were generated; nothing to export.")
        return []
    rows = merge_transfer_rows(rows)
    consolidate_transfer_output_rows(rows, include_output_series, use_output_targets)
    scenario_list = workflow_common.normalize_workflow_scenarios(
        scenarios,
        DEFAULT_SCENARIOS,
    )
    output_dir_path = Path(export_output_dir or core.EXPORT_OUTPUT_DIR)
    output_dir_path.mkdir(parents=True, exist_ok=True)
    economy_label = _infer_primary_economy(rows)
    export_filename = format_export_filename(economy_label, scenario_list, filename_template)
    previous_output_setting = core.INCLUDE_OUTPUT_SERIES_IN_LEAP_EXPORT
    previous_output_config = dict(core.TRANSFORMATION_OUTPUT_VARIABLES)
    core.INCLUDE_OUTPUT_SERIES_IN_LEAP_EXPORT = bool(include_output_series)
    core.TRANSFORMATION_OUTPUT_VARIABLES["output"] = bool(include_output_series)
    core.TRANSFORMATION_OUTPUT_VARIABLES["output_import_target"] = bool(use_output_targets)
    core.TRANSFORMATION_OUTPUT_VARIABLES["output_export_target"] = bool(use_output_targets)
    try:
        export_path = core.save_transformation_export(
            rows,
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


def consolidate_transfer_output_rows(
    rows: list[dict],
    include_output_series: bool,
    use_output_targets: bool,
) -> None:
    """Ensure transfer output values/targets are aggregated to avoid duplicates."""
    if not rows or not (include_output_series or use_output_targets):
        return
    grouped: dict[tuple[str, str], list[dict]] = {}
    for record in rows:
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


def merge_transfer_rows(rows: list[dict]) -> list[dict]:
    """Merge rows that share economy/sector/process to avoid duplicate LEAP rows."""
    if not rows:
        return rows
    grouped: dict[tuple[str, str, str], list[dict]] = {}
    for record in rows:
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


def list_export_scenarios(export_path: Path) -> list[str]:
    return leap_exports.list_scenarios(export_path, sheet_name=SHEET_NAME)


def validate_export_region(export_path: Path, region: str) -> None:
    return leap_exports.validate_region(export_path, region, sheet_name=SHEET_NAME)


def find_transfer_workbook(
    directory: Path | str | None = None, filename: str | None = None
) -> Path:
    directory_path = Path(directory or core.EXPORT_OUTPUT_DIR)
    return leap_exports.find_workbook(
        directory=directory_path,
        prefix=EXPORT_FILENAME_PREFIX,
        filename=filename,
    )


def import_transfer_workbook_to_leap(
    export_directory: Path | str | None = None,
    filename: str | None = None,
    scenario_to_run: str | None = None,
    region: str | None = None,
    include_current_accounts: bool = False,
    create_branches: bool = True,
    fill_branches: bool = True,
    raise_on_missing_branch: bool = False,
) -> Path:
    """Connect to LEAP, create branches, and fill data from the transfer export."""
    if (
        str(scenario_to_run or "").strip().lower() in {"current accounts", "current account"}
        and not include_current_accounts
    ):
        raise ValueError(
            "Direct transfer LEAP import for 'Current Accounts' is disabled "
            "unless include_current_accounts=True is passed explicitly."
        )
    export_path = find_transfer_workbook(export_directory, filename)
    target_region = region or core.EXPORT_REGION
    return leap_api.import_workbook(
        export_path=export_path,
        sheet_name=SHEET_NAME,
        scenario=scenario_to_run,
        region=target_region,
        create_branches=create_branches,
        fill_branches=fill_branches,
        include_current_accounts=include_current_accounts,
        default_branch_type=(
            BRANCH_DEMAND_CATEGORY,
            BRANCH_DEMAND_CATEGORY,
            BRANCH_DEMAND_TECHNOLOGY,
        ),
        raise_on_missing_branch=raise_on_missing_branch,
    )


def run_transfer_export_and_import(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    include_leap_import: bool = False,
    import_scenario: str | Sequence[str] | None = None,
    region: str | None = None,
    handle_current_accounts: bool = True,
    create_branches: bool = True,
    fill_branches: bool = True,
    aggregate_economy_label: str | None = None,
    feedstock_method: str | None = None,
    **export_kwargs,
) -> list[Path]:
    """Run exports and optionally push the workbook into LEAP."""
    _print_reset_reminder_for_import(include_leap_import)
    exports = assemble_transfer_workbook(
        economies=economies,
        scenarios=scenarios,
        export_output_dir=export_kwargs.get("export_output_dir"),
        filename_template=export_kwargs.get("filename_template"),
        process_config=export_kwargs.get("process_config"),
        start_year=export_kwargs.get("start_year", core.YEAR_START_FOR_ANALYSIS),
        include_output_series=export_kwargs.get("include_output_series", False),
        use_output_targets=export_kwargs.get("use_output_targets", False),
        feedstock_method=feedstock_method,
        aggregate_economy_label=aggregate_economy_label,
        build_export=export_kwargs.get("build_export", core.BUILD_LEAP_EXPORT),
    )
    if not exports or not include_leap_import:
        return exports
    scenario_list = workflow_common.normalize_workflow_scenarios(
        scenarios,
        DEFAULT_SCENARIOS,
    )
    scenario_choices = workflow_common.resolve_import_scenarios(
        scenario_list,
        import_scenario,
    )
    if get_analysis_input_write_mode() == "api" and not LEAP_API_AVAILABLE:
        print("[INFO] LEAP API unavailable in this environment; skipping branch creation/fill.")
        return exports
    for index, scenario_choice in enumerate(scenario_choices):
        import_transfer_workbook_to_leap(
            export_directory=exports[0].parent,
            filename=exports[0].name,
            scenario_to_run=scenario_choice,
            region=region or core.EXPORT_REGION,
            include_current_accounts=handle_current_accounts and index == 0,
            create_branches=create_branches and index == 0,
            fill_branches=fill_branches,
        )
    return exports


# Legacy names kept for compatibility.
def build_transfer_process_records(
    economy: str,
    sector_title: str = "Transfers",
    start_year: int = core.YEAR_START_FOR_ANALYSIS,
    process_config: dict | None = None,
    use_output_targets: bool = False,
    data_override: pd.DataFrame | None = None,
    year_cols_override: list[int] | None = None,
) -> list[dict]:
    return build_transfer_rows(
        economy=economy,
        sector_title=sector_title,
        start_year=start_year,
        process_config=process_config,
        use_output_targets=use_output_targets,
        data_override=data_override,
        year_cols_override=year_cols_override,
    )


def prepare_transfer_exports(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    export_output_dir: Path | str | None = None,
    filename_template: str | None = None,
    process_config: dict | None = None,
    start_year: int = core.YEAR_START_FOR_ANALYSIS,
    include_output_series: bool = False,
    use_output_targets: bool = False,
    feedstock_method: str | None = None,
    aggregate_economy_label: str | None = None,
    build_export: bool = core.BUILD_LEAP_EXPORT,
) -> list[Path]:
    return assemble_transfer_workbook(
        economies=economies,
        scenarios=scenarios,
        export_output_dir=export_output_dir,
        filename_template=filename_template,
        process_config=process_config,
        start_year=start_year,
        include_output_series=include_output_series,
        use_output_targets=use_output_targets,
        feedstock_method=feedstock_method,
        aggregate_economy_label=aggregate_economy_label,
        build_export=build_export,
    )


def run_transfer_pipeline(
    economies: Iterable[str] | None = None,
    scenarios: Sequence[str] | None = None,
    include_leap_import: bool = False,
    import_scenario: str | Sequence[str] | None = None,
    region: str | None = None,
    handle_current_accounts: bool = True,
    create_branches: bool = True,
    fill_branches: bool = True,
    aggregate_economy_label: str | None = None,
    **export_kwargs,
) -> list[Path]:
    return run_transfer_export_and_import(
        economies=economies,
        scenarios=scenarios,
        include_leap_import=include_leap_import,
        import_scenario=import_scenario,
        region=region,
        handle_current_accounts=handle_current_accounts,
        create_branches=create_branches,
        fill_branches=fill_branches,
        aggregate_economy_label=aggregate_economy_label,
        **export_kwargs,
    )


def locate_transfer_export(
    directory: Path | str | None = None, filename: str | None = None
) -> Path:
    return find_transfer_workbook(directory=directory, filename=filename)


def get_available_scenarios(export_path: Path) -> list[str]:
    return list_export_scenarios(export_path)


def ensure_region_in_export(export_path: Path, region: str) -> None:
    return validate_export_region(export_path, region)


def run_transfer_leap_import(
    export_directory: Path | str | None = None,
    filename: str | None = None,
    scenario_to_run: str | None = None,
    region: str | None = None,
    include_current_accounts: bool = False,
    create_branches: bool = True,
    fill_branches: bool = True,
    raise_on_missing_branch: bool = False,
) -> Path:
    return import_transfer_workbook_to_leap(
        export_directory=export_directory,
        filename=filename,
        scenario_to_run=scenario_to_run,
        region=region,
        include_current_accounts=include_current_accounts,
        create_branches=create_branches,
        fill_branches=fill_branches,
        raise_on_missing_branch=raise_on_missing_branch,
    )


def _merge_transfer_process_records(process_records: list[dict]) -> list[dict]:
    return merge_transfer_rows(process_records)


def _consolidate_transfer_outputs(
    process_records: list[dict],
    include_output_series: bool,
    use_output_targets: bool,
) -> None:
    return consolidate_transfer_output_rows(
        process_records,
        include_output_series,
        use_output_targets,
    )

#%%

EXPORT_FILENAME_TEMPLATE = workflow_cfg.TRANSFERS_EXPORT_FILENAME_TEMPLATE
EXPORT_FILENAME_PREFIX = workflow_cfg.TRANSFERS_EXPORT_FILENAME_PREFIX
SHEET_NAME = workflow_cfg.TRANSFERS_SHEET_NAME
TRANSFER_COMBINED_FLOW_KEY = "transfer_flows_combined"
TRANSFER_PROCESS_NAMES = {
    "upstream_and_refinery": "Upstream & refinery transfers",
    "upstream_liquids": "Upstream liquids transfers",
    "refinery_blending": "Refinery & blending transfers",
    "unallocated": "Transfers unallocated",
}
DEFAULT_TRANSFER_UNALLOCATED_POLICY = {
    "enabled": True,
    "process_name": TRANSFER_PROCESS_NAMES["unallocated"],
    "max_efficiency_ratio": 50.0,
    "merge_all_when_triggered": True,
}
LEAP_API_AVAILABLE = leap_api.is_available()


def get_transfer_sector_titles() -> set[str]:
    """Return all possible LEAP sector titles that the transfers workflow can produce.

    Used by zero-fill logic to identify catalog branches that belong to transfers
    even when a specific economy had no transfer data in the current run.
    """
    titles: set[str] = set()
    # Generic fallback title (SPLIT_TRANSFER_SECTORS = False)
    titles.add("Transfers")
    # All named category/process titles (SPLIT_TRANSFER_SECTORS = True)
    titles.update(TRANSFER_PROCESS_NAMES.values())
    # Any category names declared in the templates that fall outside TRANSFER_PROCESS_NAMES
    for template in TRANSFER_CATEGORY_TEMPLATES:
        cat = template.get("category")
        if cat:
            titles.add(str(cat))
    return titles


#%%
# Simple notebook-focused configuration block.
ECONOMIES = (
    list(workflow_cfg.TRANSFERS_NOTEBOOK_ECONOMIES)
    if workflow_cfg.TRANSFERS_NOTEBOOK_ECONOMIES is not None
    else list(core.ECONOMIES_TO_ANALYZE)
)
SCENARIOS = (
    list(workflow_cfg.TRANSFERS_NOTEBOOK_SCENARIOS)
    if workflow_cfg.TRANSFERS_NOTEBOOK_SCENARIOS is not None
    else list(DEFAULT_SCENARIOS)
)
INCLUDE_LEAP_IMPORT = (
    workflow_cfg.TRANSFERS_NOTEBOOK_INCLUDE_LEAP_IMPORT
    if workflow_cfg.TRANSFERS_NOTEBOOK_INCLUDE_LEAP_IMPORT is not None
    else (LEAP_API_AVAILABLE if get_analysis_input_write_mode() == "api" else True)
)
IMPORT_SCENARIOS = [
    scenario.lower()
    for scenario in SCENARIOS
    if scenario.lower() not in {"current accounts", "current account"}
]


def _print_reset_reminder_for_import(include_leap_import: bool) -> None:
    """Remind users that standalone transfer import does not clear stale trade targets."""
    if not include_leap_import:
        return
    print(
        "[WARN] Reset reminder: standalone transfers workflow import does not perform a global "
        "supply/transformation trade reset. If you need a clean rerun, run "
        "codebase/results_supply_link_workflow.py with "
        "MAIN_RUN_RESET_SUPPLY_AND_TRANSFORMATION_IMPORT_EXPORT=True."
    )


CURRENT_ACCOUNTS = workflow_cfg.TRANSFERS_NOTEBOOK_CURRENT_ACCOUNTS
INCLUDE_OUTPUT_SERIES = False
USE_OUTPUT_TARGETS = True
AGGREGATE_ECONOMY_LABEL = workflow_cfg.TRANSFERS_AGGREGATE_ECONOMY_LABEL
SPLIT_TRANSFER_SECTORS = True

#%%
if __name__ == "__main__":
    exports = run_transfer_export_and_import(
        economies=ECONOMIES,
        scenarios=SCENARIOS,
        include_leap_import=INCLUDE_LEAP_IMPORT,
        import_scenario=IMPORT_SCENARIOS,
        handle_current_accounts=CURRENT_ACCOUNTS,
        include_output_series=INCLUDE_OUTPUT_SERIES,
        use_output_targets=USE_OUTPUT_TARGETS,
        aggregate_economy_label=AGGREGATE_ECONOMY_LABEL,
    )
    if exports:
        print(f"Transfer export saved to: {exports[0]}")
#%%


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
