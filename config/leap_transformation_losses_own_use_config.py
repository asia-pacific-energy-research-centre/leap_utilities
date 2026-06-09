from __future__ import annotations

from pathlib import Path

# Legacy power-only configuration.
# This file is kept for reference and is not part of the current
# leap_utilities transformation workflow path.


SCENARIOS = ("Reference", "Target")
REGIONS = ()
YEAR_RANGE = tuple(range(2022, 2060 + 1))
UNIT = "Petajoules"
OUTPUT_ROOT = Path("outputs/leap_transformation_losses_own_use")
PROJECTION_TABLE_PATH = Path("data/merged_file_energy_ALL_20251106.csv")


RESULT_DIMENSION_HINTS = {
    "fuel": ("Fuel", "Feedstock Fuel", "Output Fuel"),
    "input_type": ("Input Type",),
    "output_type": ("Output Type",),
}


RESULT_MEMBER_HINTS = {
    "feedstock_input": ("Feedstock fuels", "Feedstock Fuels", "Feedstock"),
    "aux_total": ("Auxiliary fuels", "Auxiliary Fuels", "Auxiliary"),
    "aux_from_outputs": (
        "Auxiliary fuels from outputs",
        "Auxiliary Fuels from Outputs",
        "From module outputs",
    ),
    "aux_from_other": (
        "Auxiliary fuels from other modules/imports",
        "Auxiliary Fuels from Other Modules/Imports",
        "From other modules/imports",
    ),
    "net_output": ("Net production", "Net Production"),
    "output_for_auxiliary_use": (
        "Production for auxiliary fuel use",
        "Production for Auxiliary Fuel Use",
        "Auxiliary fuel use",
    ),
}


# Simple dict-based manual overrides for LEAP result member labels.
# Start here when the API does not expose dimension members cleanly.
# Edit these labels to exactly match the model's Results filter text.
MANUAL_RESULT_MEMBER_OVERRIDES = {
    "Electricity Generation": {
        "Inputs": {
            "feedstock_input": "",
            "aux_total": "",
            "aux_from_outputs": "",
            "aux_from_other_modules_or_imports": "",
        },
        "Outputs by Output Fuel": {
            "net_output": "",
            "output_for_auxiliary_use": "",
        },
    },
    "Transmission and Distribution": {
        "Inputs": {
            "feedstock_input": "",
            "aux_total": "",
            "aux_from_outputs": "",
            "aux_from_other_modules_or_imports": "",
        },
        "Outputs by Output Fuel": {
            "net_output": "",
            "output_for_auxiliary_use": "",
        },
    },
}


MODULES = (
    {
        "branch": r"Transformation\Electricity Generation\Processes",
        "module_name": "Electricity Generation",
        "sector_group": "power_generation",
        "dashboard_metrics": (
            "feedstock_input",
            "aux_total",
            "aux_from_outputs",
            "aux_from_other_modules_or_imports",
            "net_output",
            "output_for_auxiliary_use",
            "gross_output",
            "conversion_loss",
            "own_use_internal",
            "own_use_total",
        ),
    },
    {
        "branch": r"Transformation\Transmission and Distribution\Processes",
        "module_name": "Transmission and Distribution",
        "sector_group": "transmission_distribution",
        "dashboard_metrics": (
            "feedstock_input",
            "net_output",
            "gross_output",
            "explicit_loss_module_loss",
            "total_loss_reported",
        ),
    },
)


DASHBOARD_SHEET_DEFINITIONS = {
    "power_generation_conversion_loss": {
        "metric": "conversion_loss",
        "sheet_name": "power_generation_conversion_loss",
        "measure": "Generation conversion loss (PJ)",
        "fuel_label": "Total",
        "sector_code_9th": "",
        "sector_name": "Power generation conversion losses",
        "leap_variable": "Derived Transformation Metric",
    },
    "power_generation_own_use_internal": {
        "metric": "own_use_internal",
        "sheet_name": "power_generation_own_use_internal",
        "measure": "Power plant own use (internal) (PJ)",
        "fuel_label": "Total",
        "sector_code_9th": "",
        "sector_name": "Power plant own use",
        "leap_variable": "Derived Transformation Metric",
    },
    "power_generation_own_use_total": {
        "metric": "own_use_total",
        "sheet_name": "power_generation_own_use_total",
        "measure": "Power plant own use (total) (PJ)",
        "fuel_label": "Total",
        "sector_code_9th": "",
        "sector_name": "Power plant own use",
        "leap_variable": "Derived Transformation Metric",
    },
    "td_loss": {
        "metric": "total_loss_reported",
        "sheet_name": "td_loss",
        "measure": "Transmission and distribution loss (PJ)",
        "fuel_label": "Electricity",
        "sector_code_9th": "10_02_transmission_and_distribution_losses",
        "sector_name": "Transmission and distribution losses",
        "leap_variable": "Derived Transformation Metric",
    },
}


NINTH_MERGE_CANDIDATES = (
    {
        "module": "Transmission and Distribution",
        "metric": "total_loss_reported",
        "candidate_sector_code": "10_02_transmission_and_distribution_losses",
        "fuel_code": "17_electricity",
        "join_kind": "direct_loss_sector",
        "notes": "Cleanest existing 9th loss join. Current LEAP magnitude should be checked against 9th electricity T&D losses.",
    },
    {
        "module": "Electricity Generation",
        "metric": "own_use_total",
        "candidate_sector_code": "10_01_01_electricity_chp_and_heat_plants",
        "fuel_code": "17_electricity",
        "join_kind": "direct_own_use_sector_with_scope_gap",
        "notes": "Closest 9th own-use bucket, but comparator includes CHP and heat plants beyond electricity generation alone.",
    },
    {
        "module": "Electricity Generation",
        "metric": "own_use_internal",
        "candidate_sector_code": "10_01_01_electricity_chp_and_heat_plants",
        "fuel_code": "17_electricity",
        "join_kind": "direct_own_use_sector_with_scope_gap",
        "notes": "Closest 9th own-use bucket, but comparator includes CHP and heat plants beyond electricity generation alone.",
    },
    {
        "module": "Electricity Generation",
        "metric": "conversion_loss",
        "candidate_sector_code": "",
        "fuel_code": "",
        "join_kind": "no_direct_10xx_join",
        "notes": "Generation conversion loss is not the same concept as 10.01 own use. Compare separately or derive a 9th-side gap from generation inputs minus outputs.",
    },
)
