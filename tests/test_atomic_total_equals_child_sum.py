from __future__ import annotations

import pandas as pd

from codebase.utilities.leap_results_dashboard_v2.atomic_engine import build_atomic_outputs
from codebase.utilities.leap_results_dashboard_v2.comparison_engine import (
    build_total_rows_for_charts,
    filter_full_comparator_chart_rows,
)
from codebase.utilities.leap_results_dashboard_v2.models import AtomicSettings


def test_atomic_backed_totals_equal_sum_of_displayed_children() -> None:
    comparison_long = pd.DataFrame(
        [
            {"economy": "20_USA", "scenario": "Reference", "sheet": "A", "fuel_label": "Coal", "source": "leap", "year": 2030, "value": 10.0},
            {"economy": "20_USA", "scenario": "Reference", "sheet": "A", "fuel_label": "Coal", "source": "base", "year": 2022, "value": 6.0},
            {"economy": "20_USA", "scenario": "Reference", "sheet": "A", "fuel_label": "Coal", "source": "projection", "year": 2030, "value": 8.0},
            {"economy": "20_USA", "scenario": "Reference", "sheet": "A", "fuel_label": "Gas", "source": "leap", "year": 2030, "value": 4.0},
            {"economy": "20_USA", "scenario": "Reference", "sheet": "A", "fuel_label": "Gas", "source": "base", "year": 2022, "value": 2.0},
            {"economy": "20_USA", "scenario": "Reference", "sheet": "A", "fuel_label": "Gas", "source": "projection", "year": 2030, "value": 3.0},
        ]
    )
    mapping_status = pd.DataFrame(
        [
            {"sheet": "A", "fuel_label": "Coal", "sector_code_9th": "14_03_11_nonspecified_industry", "ninth_fuel_code": "01_x_thermal_coal", "esto_flow": "14.03.11 Non-specified industry", "esto_product": "01.02 Other bituminous coal", "mapping_source": "canonical"},
            {"sheet": "A", "fuel_label": "Gas", "sector_code_9th": "14_03_11_nonspecified_industry", "ninth_fuel_code": "08_01_natural_gas", "esto_flow": "14.03.11 Non-specified industry", "esto_product": "08.01 Natural gas", "mapping_source": "canonical"},
        ]
    )
    sheet_map = pd.DataFrame([{"sheet_name": "A", "sector_name": "A"}])
    canonical_pairs = pd.DataFrame(
        columns=["9th_sector", "9th_fuel", "esto_flow", "esto_product"]
    )
    base_df = pd.DataFrame(
        [
            {
                "economy": "20USA",
                "flows": "14.03.11 Non-specified industry",
                "products": "01.02 Other bituminous coal",
                "2022": 6.0,
            },
            {
                "economy": "20USA",
                "flows": "14.03.11 Non-specified industry",
                "products": "08.01 Natural gas",
                "2022": 2.0,
            },
        ]
    )
    ninth_df = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "scenarios": "reference",
                "sectors": "14_03_11_nonspecified_industry",
                "sub1sectors": "x",
                "sub2sectors": "x",
                "sub3sectors": "x",
                "sub4sectors": "x",
                "fuels": "01_coal",
                "subfuels": "01_x_thermal_coal",
                "2030": 8.0,
            },
            {
                "economy": "20_USA",
                "scenarios": "reference",
                "sectors": "14_03_11_nonspecified_industry",
                "sub1sectors": "x",
                "sub2sectors": "x",
                "sub3sectors": "x",
                "sub4sectors": "x",
                "fuels": "08_gas",
                "subfuels": "08_01_natural_gas",
                "2030": 3.0,
            },
        ]
    )
    leap_long = pd.DataFrame(
        [
            {
                "sheet_name": "A",
                "fuel_label": "Coal",
                "scenario": "Reference",
                "year": 2030,
                "leap_value": 10.0,
                "economy": "USA",
            },
            {
                "sheet_name": "A",
                "fuel_label": "Gas",
                "scenario": "Reference",
                "year": 2030,
                "leap_value": 4.0,
                "economy": "USA",
            },
        ]
    )
    settings = AtomicSettings(enabled=True, rollout_mode="shadow", many_to_many_policy="error", write_shadow_outputs=True)
    atomic = build_atomic_outputs(
        comparison_long=comparison_long,
        mapping_status=mapping_status,
        sheet_map=sheet_map,
        canonical_pairs=canonical_pairs,
        base_df=base_df,
        ninth_df=ninth_df,
        leap_long=leap_long,
        base_economy="20USA",
        projection_economy="20_USA",
        settings=settings,
    )
    strict = filter_full_comparator_chart_rows(atomic["atomic_comparison_long"])
    totals = build_total_rows_for_charts(strict, mapping_status)
    chart = pd.concat(
        [strict[strict["fuel_label"].ne("Total")], totals],
        ignore_index=True,
        sort=False,
    )
    chart = filter_full_comparator_chart_rows(chart).drop_duplicates(
        subset=["economy", "sheet", "fuel_label", "scenario", "source", "year"],
        keep="first",
    )
    child = chart[(chart["fuel_label"] != "Total") & (chart["source"] == "projection")]
    child_sum = child.groupby(["economy", "sheet", "scenario", "source", "year"], as_index=False)["value"].sum()
    total = chart[(chart["fuel_label"] == "Total") & (chart["source"] == "projection")]
    merged = total.merge(child_sum, on=["economy", "sheet", "scenario", "source", "year"], suffixes=("_total", "_child"))
    assert len(merged) == 1
    assert float(merged["value_total"].iloc[0]) == float(merged["value_child"].iloc[0])


def test_total_base_rows_merge_scenario_split_children_once() -> None:
    chart_rows = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "Electricity plants",
                "measure": "Electricity generation inputs (PJ)",
                "fuel_label": "Coal",
                "scenario": "Reference",
                "source": "base",
                "year": 2022,
                "value": 6.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Electricity plants",
                "measure": "Electricity generation inputs (PJ)",
                "fuel_label": "Coal",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 8.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Electricity plants",
                "measure": "Electricity generation inputs (PJ)",
                "fuel_label": "Gas",
                "scenario": "Target",
                "source": "base",
                "year": 2022,
                "value": 2.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Electricity plants",
                "measure": "Electricity generation inputs (PJ)",
                "fuel_label": "Gas",
                "scenario": "Target",
                "source": "projection",
                "year": 2030,
                "value": 3.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Electricity plants",
                "measure": "Electricity generation inputs (PJ)",
                "fuel_label": "Biomass",
                "scenario": "Reference",
                "source": "base",
                "year": 2022,
                "value": 1.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Electricity plants",
                "measure": "Electricity generation inputs (PJ)",
                "fuel_label": "Biomass",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 1.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Electricity plants",
                "measure": "Electricity generation inputs (PJ)",
                "fuel_label": "Biomass",
                "scenario": "Target",
                "source": "base",
                "year": 2022,
                "value": 1.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Electricity plants",
                "measure": "Electricity generation inputs (PJ)",
                "fuel_label": "Biomass",
                "scenario": "Target",
                "source": "projection",
                "year": 2030,
                "value": 2.0,
            },
        ]
    )

    totals = build_total_rows_for_charts(chart_rows, pd.DataFrame())
    base_totals = totals[
        (totals["fuel_label"] == "Total")
        & (totals["source"] == "base")
        & (totals["year"] == 2022)
    ].sort_values("scenario")

    assert list(base_totals["scenario"]) == ["Reference", "Target"]
    assert list(base_totals["value"]) == [9.0, 9.0]
