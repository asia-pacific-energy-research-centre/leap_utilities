from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from codebase.leap_results_dashboard_v2_workflow import _build_chart_input_for_rendering
from codebase.leap_results_workflow import validate_component_rows_present
from codebase.utilities.leap_results_dashboard_v2.comparison_engine import (
    _common_level_only_filter,
    _fail_fast_leaf_holes,
    aggregate_leap_for_shared_sector_groups,
    build_chart_line_mapping_ledger,
    build_total_component_ledger,
    filter_full_comparator_chart_rows,
)
from codebase.utilities.leap_results_dashboard_v2.atomic_engine import (
    _build_atomic_edge_candidates,
    _prepare_line_rows,
)
from codebase.utilities.leap_results_dashboard_v2.mapping_engine import annotate_mapping_status
from codebase.utilities.leap_results_dashboard_v2.shadow_compare import compare_outputs


def test_annotate_mapping_status_sets_aggregated_flag() -> None:
    df = pd.DataFrame(
        [
            {"sheet": "A", "fuel_label": "Coal", "mapping_source": "canonical_aggregated", "mapping_note": "", "sector_match_method": ""},
            {"sheet": "B", "fuel_label": "Gas", "mapping_source": "canonical", "mapping_note": "", "sector_match_method": ""},
        ]
    )
    out = annotate_mapping_status(df)
    assert bool(out.loc[out["sheet"] == "A", "aggregated_mapping"].iloc[0])
    assert not bool(out.loc[out["sheet"] == "B", "aggregated_mapping"].iloc[0])


def test_common_level_only_filter_keeps_leap_and_full_comparator_groups() -> None:
    comparison_long = pd.DataFrame(
        [
            {"sheet": "A", "fuel_label": "Coal", "scenario": "Reference", "source": "leap", "year": 2030, "value": 10},
            {"sheet": "A", "fuel_label": "Coal", "scenario": "Reference", "source": "base", "year": 2022, "value": 9},
            {"sheet": "A", "fuel_label": "Coal", "scenario": "Reference", "source": "projection", "year": 2030, "value": 11},
            {"sheet": "B", "fuel_label": "Gas", "scenario": "Reference", "source": "leap", "year": 2030, "value": 7},
            {"sheet": "B", "fuel_label": "Gas", "scenario": "Reference", "source": "base", "year": 2022, "value": 6},
        ]
    )
    filtered = _common_level_only_filter(comparison_long, pd.DataFrame())
    assert len(filtered[(filtered["sheet"] == "A") & (filtered["source"] != "leap")]) == 2
    assert len(filtered[(filtered["sheet"] == "B") & (filtered["source"] != "leap")]) == 0
    assert len(filtered[(filtered["sheet"] == "B") & (filtered["source"] == "leap")]) == 1


def test_fail_fast_leaf_holes_raises_on_missing_comparator_values() -> None:
    comparison_long = pd.DataFrame(
        [
            {"sheet": "A", "fuel_label": "Coal", "scenario": "Reference", "source": "base", "year": 2022, "value": None},
            {"sheet": "A", "fuel_label": "Coal", "scenario": "Reference", "source": "projection", "year": 2030, "value": 1.0},
        ]
    )
    mapping_status = pd.DataFrame([{"sheet": "A", "fuel_label": "Coal", "mapped": True}])
    with pytest.raises(RuntimeError, match="residual missing comparator"):
        _fail_fast_leaf_holes(comparison_long, mapping_status)


def test_validate_component_rows_present_raises_when_total_has_no_nonzero_children() -> None:
    table_df = pd.DataFrame(
        [
            ["Outputs by Feedstock Fuel", "", "", ""],
            ["Scenario: Reference, Region: United States of America", "", "", ""],
            ["Branch: Transformation\\Transfers unallocated\\Processes", "", "", ""],
            ["Units: Petajoules", "", "", ""],
            ["", "", "", ""],
            ["Fuel", 2022, 2023, 2024],
            ["Total", 5.0, 6.0, 7.0],
            ["Natural Gas", 0.0, 0.0, 0.0],
            ["Coal", 0.0, 0.0, 0.0],
        ]
    )
    with pytest.raises(RuntimeError, match="Suspicious LEAP results table detected"):
        validate_component_rows_present(table_df, context="test.xlsx/transfers_unallocated_out_feed")


def test_shadow_compare_writes_summary(tmp_path: Path) -> None:
    v1_dir = tmp_path / "v1"
    v2_dir = tmp_path / "v2"
    v1_dir.mkdir()
    v2_dir.mkdir()
    pd.DataFrame(
        [
            {"sheet": "A", "fuel_label": "Coal", "scenario": "Reference", "source": "leap", "year": 2030, "value": 1.0}
        ]
    ).to_csv(v1_dir / "comparison_long.csv", index=False)
    pd.DataFrame(
        [
            {"sheet": "A", "fuel_label": "Coal", "scenario": "Reference", "source": "leap", "year": 2030, "value": 2.0}
        ]
    ).to_csv(v2_dir / "comparison_long.csv", index=False)

    out = compare_outputs(v1_output_dir=v1_dir, v2_output_dir=v2_dir, out_path=tmp_path / "summary.csv")
    assert out.exists()
    summary = pd.read_csv(out)
    assert (summary["metric"] == "value_diff_rows").any()


def test_filter_full_comparator_chart_rows_drops_incomplete_groups() -> None:
    comp = pd.DataFrame(
        [
            {"sheet": "A", "fuel_label": "Coal", "scenario": "Reference", "source": "leap", "year": 2030, "value": 1.0},
            {"sheet": "A", "fuel_label": "Coal", "scenario": "Reference", "source": "base", "year": 2022, "value": 1.0},
            {"sheet": "A", "fuel_label": "Coal", "scenario": "Reference", "source": "projection", "year": 2030, "value": 1.0},
            {"sheet": "B", "fuel_label": "Gas", "scenario": "Reference", "source": "leap", "year": 2030, "value": 1.0},
            {"sheet": "B", "fuel_label": "Gas", "scenario": "Reference", "source": "base", "year": 2022, "value": 1.0},
        ]
    )
    out = filter_full_comparator_chart_rows(comp)
    assert set(out["sheet"].unique()) == {"A"}


def test_aggregate_leap_for_shared_sector_groups_sums_across_sheets() -> None:
    comp = pd.DataFrame(
        [
            {"economy": "20_USA", "sheet": "Agriculture", "fuel_label": "Electricity", "scenario": "Reference", "source": "leap", "year": 2030, "value": 10.0},
            {"economy": "20_USA", "sheet": "Fishing", "fuel_label": "Electricity", "scenario": "Reference", "source": "leap", "year": 2030, "value": 0.0},
            {"economy": "20_USA", "sheet": "Agriculture", "fuel_label": "Electricity", "scenario": "Reference", "source": "projection", "year": 2030, "value": 20.0},
        ]
    )
    status = pd.DataFrame(
        [
            {"sheet": "Agriculture", "sector_code_9th": "16_02_agriculture_and_fishing"},
            {"sheet": "Fishing", "sector_code_9th": "16_02_agriculture_and_fishing"},
        ]
    )
    out = aggregate_leap_for_shared_sector_groups(comp, status)
    chk = out[(out["source"] == "leap") & (out["fuel_label"] == "Electricity") & (out["scenario"] == "Reference") & (out["year"] == 2030)]
    vals = chk.set_index("sheet")["value"].to_dict()
    assert vals["Agriculture"] == 10.0
    assert vals["Fishing"] == 10.0


def test_atomic_engine_expands_explicit_aggregated_components() -> None:
    comparison_long = pd.DataFrame(
        [
            {
                "economy": "20USA",
                "sheet": "elecgen_inputs",
                "measure": "Electricity generation inputs (PJ)",
                "fuel_label": "Coal Bituminous",
                "scenario": "Reference",
                "source": "base",
                "year": 2022,
                "value": 8781.5,
            },
            {
                "economy": "20_USA",
                "sheet": "elecgen_inputs",
                "measure": "Electricity generation inputs (PJ)",
                "fuel_label": "Coal Bituminous",
                "scenario": "Reference",
                "source": "projection",
                "year": 2023,
                "value": 6737.6,
            },
        ]
    )
    mapping_status = pd.DataFrame(
        [
            {
                "sheet": "elecgen_inputs",
                "fuel_label": "Coal Bituminous",
                "sector_code_9th": "09_01_electricity_plants",
                "ninth_fuel_code": "2 fuels (aggregated)",
                "esto_flow": "09.01.01 Electricity plants",
                "esto_product": "4 products (aggregated)",
                "mapping_source": "explicit",
                "sector_match_method": "manual_override",
                "projection_parent_sector_code": "",
                "projection_fuel_codes_detail": '["01_x_thermal_coal", "01_05_lignite"]',
                "projection_targets_detail": (
                    '[{"sector_code_9th":"09_01_01_coal_power","ninth_fuel_code":"01_x_thermal_coal"},'
                    '{"sector_code_9th":"09_01_01_coal_power","ninth_fuel_code":"01_05_lignite"}]'
                ),
                "base_targets_detail": (
                    '[{"esto_flow":"09.01.01 Electricity plants","esto_product":"01.02 Other bituminous coal"},'
                    '{"esto_flow":"09.01.01 Electricity plants","esto_product":"01.03 Sub-bituminous coal"},'
                    '{"esto_flow":"09.01.01 Electricity plants","esto_product":"01.04 Anthracite"},'
                    '{"esto_flow":"09.01.01 Electricity plants","esto_product":"01.05 Lignite"}]'
                ),
            }
        ]
    )
    resolved_levels = pd.DataFrame(
        [
            {
                "sheet": "elecgen_inputs",
                "resolved_node_id": "elecgen_inputs",
                "resolved_node_level": "sheet",
            }
        ]
    )

    line_rows = _prepare_line_rows(comparison_long, mapping_status, resolved_levels)
    edges = _build_atomic_edge_candidates(line_rows=line_rows, canonical_pairs=pd.DataFrame())

    base_edges = edges[edges["source_family"] == "base"]
    projection_edges = edges[edges["source_family"] == "projection"]

    assert set(base_edges["esto_product"]) == {
        "01.02 Other bituminous coal",
        "01.03 Sub-bituminous coal",
        "01.04 Anthracite",
        "01.05 Lignite",
    }
    assert set(projection_edges["sector_node"]) == {"09_01_01_coal_power"}
    assert set(projection_edges["fuel_node"]) == {"01_x_thermal_coal", "01_05_lignite"}


def test_build_chart_input_for_rendering_keeps_partially_mapped_rows() -> None:
    comparison_long = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "elecgen_inputs",
                "measure": "Electricity generation inputs (PJ)",
                "fuel_label": "Biomass",
                "scenario": "Reference",
                "source": "leap",
                "year": 2030,
                "value": 200.0,
            },
            {
                "economy": "20USA",
                "sheet": "elecgen_inputs",
                "measure": "Electricity generation inputs (PJ)",
                "fuel_label": "Biomass",
                "scenario": "Reference",
                "source": "base",
                "year": 2022,
                "value": 162.3,
            },
            {
                "economy": "20_USA",
                "sheet": "elecgen_inputs",
                "measure": "Electricity generation inputs (PJ)",
                "fuel_label": "Biomass",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": float("nan"),
            },
        ]
    )
    mapping_status = pd.DataFrame(
        [
            {
                "sheet": "elecgen_inputs",
                "measure": "Electricity generation inputs (PJ)",
                "fuel_label": "Biomass",
                "mapped": False,
                "partially_mapped": True,
                "has_any_mapping": True,
            }
        ]
    )
    sheet_map = pd.DataFrame(
        [
            {
                "sheet_name": "elecgen_inputs",
                "sector_name": "Electricity plants",
                "notes": "Electricity generation inputs",
            }
        ]
    )

    out = _build_chart_input_for_rendering(
        comparison_long=comparison_long,
        mapping_status=mapping_status,
        sheet_map=sheet_map,
    )

    biomass = out[out["fuel_label"].eq("Biomass")].copy()
    assert not biomass.empty
    assert "elecgen_inputs" in set(biomass["sheet"]) or "Electricity plants" in set(biomass["sheet"])
    assert bool(biomass["force_show_chart"].fillna(False).any())


def test_total_component_ledger_flags_exact_duplicate_comparator_keys_without_max_dedupe() -> None:
    chart_rows = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "Buildings",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 100.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Buildings",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity duplicate",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 40.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Buildings",
                "measure": "Demand (PJ)",
                "fuel_label": "Gas",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 25.0,
            },
        ]
    )
    mapping_status = pd.DataFrame(
        [
            {
                "sheet": "Buildings",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity",
                "sector_code_9th": "16_01_buildings",
                "ninth_fuel_code": "17_electricity",
                "projection_parent_sector_code": "16_01",
            },
            {
                "sheet": "Buildings",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity duplicate",
                "sector_code_9th": "16_01_buildings",
                "ninth_fuel_code": "17_electricity",
                "projection_parent_sector_code": "16_01",
            },
            {
                "sheet": "Buildings",
                "measure": "Demand (PJ)",
                "fuel_label": "Gas",
                "sector_code_9th": "16_01_buildings",
                "ninth_fuel_code": "08_01_natural_gas",
                "projection_parent_sector_code": "",
            },
        ]
    )

    ledger = build_total_component_ledger(chart_rows, mapping_status)
    dup = ledger[ledger["exact_comparator_key"] == "16_01_buildings|17_electricity"].sort_values("member_fuel_label")

    assert list(dup["member_value"]) == [40.0, 100.0]
    assert list(dup["component_selected_value"]) == [40.0, 100.0]
    assert set(dup["duplicate_exact_comparator_key_count"]) == {2}
    assert set(dup["component_included_in_total"]) == {True}
    assert set(dup["is_selected_max"]) == {True}
    assert set(dup["projection_parent_sector_code"]) == {"16_01"}
    assert set(dup["is_leaf_level"]) == {False}


def test_chart_line_mapping_ledger_marks_first_duplicate_exact_comparator_key() -> None:
    chart_rows = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "Buildings",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 100.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Buildings",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity duplicate",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 40.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Buildings",
                "measure": "Demand (PJ)",
                "fuel_label": "Total",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 140.0,
            },
        ]
    )
    mapping_status = pd.DataFrame(
        [
            {
                "sheet": "Buildings",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity",
                "sector_code_9th": "16_01_buildings",
                "ninth_fuel_code": "17_electricity",
                "projection_parent_sector_code": "16_01",
            },
            {
                "sheet": "Buildings",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity duplicate",
                "sector_code_9th": "16_01_buildings",
                "ninth_fuel_code": "17_electricity",
                "projection_parent_sector_code": "16_01",
            },
        ]
    )

    ledger = build_chart_line_mapping_ledger(chart_rows, mapping_status)
    dup = ledger[
        (ledger["sheet"] == "Buildings")
        & (ledger["source"] == "projection")
        & (ledger["fuel_label"] != "Total")
    ].sort_values("fuel_label")

    assert set(dup["aggregate_group_key"]) == {"16_01_buildings|17_electricity"}
    assert set(dup["duplicate_exact_comparator_key_count"]) == {2}
    assert list(dup["first_of_aggregate"]) == [True, False]
    assert list(dup["first_of_aggregate_or_non_aggregate"]) == [True, False]

    total = ledger[ledger["fuel_label"] == "Total"]
    assert int(total["total_component_bucket_count"].iloc[0]) == 1


def test_build_chart_input_for_rendering_drops_parent_comparator_when_child_exists() -> None:
    comparison_long = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "Child sheet",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity parent",
                "scenario": "Reference",
                "source": "leap",
                "year": 2030,
                "value": 130.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Child sheet",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity child",
                "scenario": "Reference",
                "source": "leap",
                "year": 2030,
                "value": 90.0,
            },
            {
                "economy": "20USA",
                "sheet": "Child sheet",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity parent",
                "scenario": "Reference",
                "source": "base",
                "year": 2022,
                "value": 120.0,
            },
            {
                "economy": "20USA",
                "sheet": "Child sheet",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity child",
                "scenario": "Reference",
                "source": "base",
                "year": 2022,
                "value": 80.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Child sheet",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity parent",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 120.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Child sheet",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity child",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 80.0,
            },
        ]
    )
    mapping_status = pd.DataFrame(
        [
            {
                "sheet": "Child sheet",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity parent",
                "mapped": True,
                "partially_mapped": False,
                "has_any_mapping": True,
                "sector_code_9th": "16_01",
                "ninth_fuel_code": "17_electricity",
                "projection_parent_sector_code": "",
            },
            {
                "sheet": "Child sheet",
                "measure": "Demand (PJ)",
                "fuel_label": "Electricity child",
                "mapped": True,
                "partially_mapped": False,
                "has_any_mapping": True,
                "sector_code_9th": "16_01_buildings",
                "ninth_fuel_code": "17_electricity",
                "projection_parent_sector_code": "16_01",
            },
        ]
    )
    sheet_map = pd.DataFrame(
        [
            {
                "sheet_name": "Child sheet",
                "sector_name": "Buildings",
                "notes": "",
            }
        ]
    )

    out = _build_chart_input_for_rendering(
        comparison_long=comparison_long,
        mapping_status=mapping_status,
        sheet_map=sheet_map,
    )

    proj = out[(out["source"] == "projection") & (out["fuel_label"] != "Total")].copy()
    assert set(proj["fuel_label"]) == {"Electricity child"}

    total = out[(out["source"] == "projection") & (out["fuel_label"] == "Total")]
    assert len(total) == 1
    assert float(total["value"].iloc[0]) == 80.0


def test_build_chart_input_for_rendering_drops_parent_fuel_when_numeric_child_exists() -> None:
    comparison_long = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Gas aggregate",
                "scenario": "Reference",
                "source": "leap",
                "year": 2030,
                "value": 70.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Gas child",
                "scenario": "Reference",
                "source": "leap",
                "year": 2030,
                "value": 50.0,
            },
            {
                "economy": "20USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Gas aggregate",
                "scenario": "Reference",
                "source": "base",
                "year": 2022,
                "value": 60.0,
            },
            {
                "economy": "20USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Gas child",
                "scenario": "Reference",
                "source": "base",
                "year": 2022,
                "value": 40.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Gas aggregate",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 60.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Gas child",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 40.0,
            },
        ]
    )
    mapping_status = pd.DataFrame(
        [
            {
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Gas aggregate",
                "mapped": True,
                "partially_mapped": False,
                "has_any_mapping": True,
                "sector_code_9th": "07_supply",
                "ninth_fuel_code": "08",
                "projection_parent_sector_code": "",
            },
            {
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Gas child",
                "mapped": True,
                "partially_mapped": False,
                "has_any_mapping": True,
                "sector_code_9th": "07_supply",
                "ninth_fuel_code": "08_01_natural_gas",
                "projection_parent_sector_code": "",
            },
        ]
    )
    sheet_map = pd.DataFrame([{"sheet_name": "Supply sheet", "sector_name": "Supply", "notes": ""}])

    out = _build_chart_input_for_rendering(
        comparison_long=comparison_long,
        mapping_status=mapping_status,
        sheet_map=sheet_map,
    )

    proj = out[(out["source"] == "projection") & (out["fuel_label"] != "Total")].copy()
    assert set(proj["fuel_label"]) == {"Gas child"}


def test_build_chart_input_for_rendering_does_not_treat_x_fuel_as_numeric_parent() -> None:
    comparison_long = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Other petroleum products",
                "scenario": "Reference",
                "source": "leap",
                "year": 2030,
                "value": 90.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Ethane",
                "scenario": "Reference",
                "source": "leap",
                "year": 2030,
                "value": 20.0,
            },
            {
                "economy": "20USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Other petroleum products",
                "scenario": "Reference",
                "source": "base",
                "year": 2022,
                "value": 70.0,
            },
            {
                "economy": "20USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Ethane",
                "scenario": "Reference",
                "source": "base",
                "year": 2022,
                "value": 10.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Other petroleum products",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 70.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Ethane",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 10.0,
            },
        ]
    )
    mapping_status = pd.DataFrame(
        [
            {
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Other petroleum products",
                "mapped": True,
                "partially_mapped": False,
                "has_any_mapping": True,
                "sector_code_9th": "07_supply",
                "ninth_fuel_code": "07_x_other_petroleum_products",
                "projection_parent_sector_code": "",
            },
            {
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Ethane",
                "mapped": True,
                "partially_mapped": False,
                "has_any_mapping": True,
                "sector_code_9th": "07_supply",
                "ninth_fuel_code": "07_11_ethane",
                "projection_parent_sector_code": "",
            },
        ]
    )
    sheet_map = pd.DataFrame([{"sheet_name": "Supply sheet", "sector_name": "Supply", "notes": ""}])

    out = _build_chart_input_for_rendering(
        comparison_long=comparison_long,
        mapping_status=mapping_status,
        sheet_map=sheet_map,
    )

    proj = out[(out["source"] == "projection") & (out["fuel_label"] != "Total")].copy()
    assert set(proj["fuel_label"]) == {"Other petroleum products", "Ethane"}


def test_build_chart_input_for_rendering_uses_x_override_for_parent_aggregate() -> None:
    comparison_long = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Thermal coal aggregate",
                "scenario": "Reference",
                "source": "leap",
                "year": 2030,
                "value": 55.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Bituminous coal",
                "scenario": "Reference",
                "source": "leap",
                "year": 2030,
                "value": 45.0,
            },
            {
                "economy": "20USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Thermal coal aggregate",
                "scenario": "Reference",
                "source": "base",
                "year": 2022,
                "value": 50.0,
            },
            {
                "economy": "20USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Bituminous coal",
                "scenario": "Reference",
                "source": "base",
                "year": 2022,
                "value": 40.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Thermal coal aggregate",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 50.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Bituminous coal",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 40.0,
            },
        ]
    )
    mapping_status = pd.DataFrame(
        [
            {
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Thermal coal aggregate",
                "mapped": True,
                "partially_mapped": False,
                "has_any_mapping": True,
                "sector_code_9th": "07_supply",
                "ninth_fuel_code": "01_x_thermal_coal",
                "projection_parent_sector_code": "",
            },
            {
                "sheet": "Supply sheet",
                "measure": "Supply (PJ)",
                "fuel_label": "Bituminous coal",
                "mapped": True,
                "partially_mapped": False,
                "has_any_mapping": True,
                "sector_code_9th": "07_supply",
                "ninth_fuel_code": "01_02_other_bituminous_coal",
                "projection_parent_sector_code": "",
            },
        ]
    )
    sheet_map = pd.DataFrame([{"sheet_name": "Supply sheet", "sector_name": "Supply", "notes": ""}])

    out = _build_chart_input_for_rendering(
        comparison_long=comparison_long,
        mapping_status=mapping_status,
        sheet_map=sheet_map,
    )

    proj = out[(out["source"] == "projection") & (out["fuel_label"] != "Total")].copy()
    assert set(proj["fuel_label"]) == {"Bituminous coal"}
