from pathlib import Path

import pandas as pd

from codebase.utilities import leap_results_dashboard_utils as dashboard_utils
from codebase.utilities.leap_results_dashboard_utils import (
    _aggregate_display_rows_to_total,
    _prepare_render_long,
    _infer_fuel_from_label_fallback,
    build_comparisons,
    load_explicit_sector_fuel_mappings,
)


def test_infer_fuel_from_label_fallback_maps_known_problem_labels():
    assert _infer_fuel_from_label_fallback("White spirit SBP") == "07_x_other_petroleum_products"
    assert _infer_fuel_from_label_fallback("PetProd nonspecified") == "07_x_other_petroleum_products"
    assert _infer_fuel_from_label_fallback("Other sources") == "16_09_other_sources"
    assert _infer_fuel_from_label_fallback("Hydrogen") == "16_x_hydrogen"
    assert _infer_fuel_from_label_fallback("Anthracite") == ""
    assert _infer_fuel_from_label_fallback("BKB and PB") == ""
    assert _infer_fuel_from_label_fallback("Blast furnace gas") == ""
    assert _infer_fuel_from_label_fallback("Coal tar") == ""
    assert _infer_fuel_from_label_fallback("Coke oven coke") == ""
    assert _infer_fuel_from_label_fallback("Coke oven gas") == ""
    assert _infer_fuel_from_label_fallback("Electricity") == ""
    assert _infer_fuel_from_label_fallback("Gas coke") == ""
    assert _infer_fuel_from_label_fallback("Kerosene type jet fuel") == ""
    assert _infer_fuel_from_label_fallback("Other recovered gases") == ""
    assert _infer_fuel_from_label_fallback("Patent fuel") == ""
    assert _infer_fuel_from_label_fallback("Sub bituminous coal") == ""


def test_infer_fuel_from_label_fallback_returns_blank_for_unknown_label():
    assert _infer_fuel_from_label_fallback("Passenger non road") == ""


def test_build_dashboards_renders_supply_overview_row(tmp_path, monkeypatch):
    charts_dir = tmp_path / "charts"
    charts_dir.mkdir()

    for stem in [
        "Exports__Exports__PJ__Total",
        "Imports__Imports__PJ__Total",
        "Production__Production__PJ__Total",
    ]:
        (charts_dir / f"{stem}.html").write_text("<html><body>chart</body></html>", encoding="utf-8")

    def _fake_make_chart(
        sheet: str,
        fuel: str,
        sub: pd.DataFrame,
        charts_dir: Path,
        backend: str = "plotly",
        *,
        display_sheet: str | None = None,
        file_sheet: str | None = None,
    ) -> Path:
        chart_path = charts_dir / f"{dashboard_utils._safe_token(file_sheet or sheet)}__{dashboard_utils._safe_token(fuel)}.html"
        chart_path.write_text("<html><body>chart</body></html>", encoding="utf-8")
        return chart_path

    monkeypatch.setattr(dashboard_utils, "make_chart", _fake_make_chart)

    comparison_long = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "Production",
                "measure": "Production (PJ)",
                "fuel_label": "Coal",
                "scenario": "Reference",
                "source": "leap",
                "year": year,
                "value": value,
            }
            for year, value in [(2022, 100.0), (2023, 105.0)]
        ]
        + [
            {
                "economy": "20_USA",
                "sheet": "Imports",
                "measure": "Imports (PJ)",
                "fuel_label": "Coal",
                "scenario": "Reference",
                "source": "leap",
                "year": year,
                "value": value,
            }
            for year, value in [(2022, 20.0), (2023, 24.0)]
        ]
        + [
            {
                "economy": "20_USA",
                "sheet": "Exports",
                "measure": "Exports (PJ)",
                "fuel_label": "Coal",
                "scenario": "Reference",
                "source": "leap",
                "year": year,
                "value": value,
            }
            for year, value in [(2022, 10.0), (2023, 11.0)]
        ]
    )

    dashboard_utils.build_dashboards(
        output_dir=tmp_path / "out",
        comparison_long=comparison_long,
        charts_dir=charts_dir,
        mapping_status=None,
    )

    supply_html = (tmp_path / "out" / "dashboards" / "node__Supply.html").read_text(encoding="utf-8")
    assert "Measure: Supply overview (PJ)" in supply_html
    assert 'href="#sec-Supply_overview"' in supply_html

    supply_overview_section = supply_html.split('<section id="sec-Supply_overview"', 1)[1].split("</section>", 1)[0]
    assert "TPES is calculated here as production + imports - exports." in supply_overview_section
    for label in ["TPES", "Exports", "Imports", "Production"]:
        assert f">{label}</div>" in supply_overview_section


def test_build_dashboards_supply_tpes_uses_single_projection_family_trace(tmp_path, monkeypatch):
    charts_dir = tmp_path / "charts"
    charts_dir.mkdir()

    def _fake_make_chart(
        sheet: str,
        fuel: str,
        sub: pd.DataFrame,
        charts_dir: Path,
        backend: str = "plotly",
        *,
        display_sheet: str | None = None,
        file_sheet: str | None = None,
    ) -> Path:
        chart_path = charts_dir / f"{dashboard_utils._safe_token(file_sheet or sheet)}__{dashboard_utils._safe_token(fuel)}.html"
        chart_path.write_text("<html><body>chart</body></html>", encoding="utf-8")
        return chart_path

    monkeypatch.setattr(dashboard_utils, "make_chart", _fake_make_chart)

    comparison_long = pd.DataFrame(
        [
            {"economy": "20_USA", "sheet": "Production", "measure": "Production (PJ)", "fuel_label": "Total", "scenario": "Reference", "source": "projection", "year": 2030, "value": 100.0},
            {"economy": "20_USA", "sheet": "Production", "measure": "Production (PJ)", "fuel_label": "Total", "scenario": "Reference", "source": "projection_estimated", "year": 2030, "value": 10.0},
            {"economy": "20_USA", "sheet": "Imports", "measure": "Imports (PJ)", "fuel_label": "Total", "scenario": "Reference", "source": "projection", "year": 2030, "value": 50.0},
            {"economy": "20_USA", "sheet": "Imports", "measure": "Imports (PJ)", "fuel_label": "Total", "scenario": "Reference", "source": "projection_estimated", "year": 2030, "value": 5.0},
            {"economy": "20_USA", "sheet": "Exports", "measure": "Exports (PJ)", "fuel_label": "Total", "scenario": "Reference", "source": "projection", "year": 2030, "value": 20.0},
            {"economy": "20_USA", "sheet": "Exports", "measure": "Exports (PJ)", "fuel_label": "Total", "scenario": "Reference", "source": "projection_estimated", "year": 2030, "value": 2.0},
            {"economy": "20_USA", "sheet": "Production", "measure": "Production (PJ)", "fuel_label": "Coal", "scenario": "Reference", "source": "projection", "year": 2030, "value": 100.0},
            {"economy": "20_USA", "sheet": "Production", "measure": "Production (PJ)", "fuel_label": "Gas", "scenario": "Reference", "source": "projection_estimated", "year": 2030, "value": 10.0},
            {"economy": "20_USA", "sheet": "Imports", "measure": "Imports (PJ)", "fuel_label": "Coal", "scenario": "Reference", "source": "projection", "year": 2030, "value": 50.0},
            {"economy": "20_USA", "sheet": "Imports", "measure": "Imports (PJ)", "fuel_label": "Gas", "scenario": "Reference", "source": "projection_estimated", "year": 2030, "value": 5.0},
            {"economy": "20_USA", "sheet": "Exports", "measure": "Exports (PJ)", "fuel_label": "Coal", "scenario": "Reference", "source": "projection", "year": 2030, "value": 20.0},
            {"economy": "20_USA", "sheet": "Exports", "measure": "Exports (PJ)", "fuel_label": "Gas", "scenario": "Reference", "source": "projection_estimated", "year": 2030, "value": 2.0},
        ]
    )

    dashboard_utils.build_dashboards(
        output_dir=tmp_path / "out",
        comparison_long=comparison_long,
        charts_dir=charts_dir,
        mapping_status=None,
    )

    tpes_chart = (charts_dir / "node__Supply__TPES_excl__bunkers__PJ__Total.html").read_text(encoding="utf-8")
    assert "9th projection est/real REF" in tpes_chart
    assert "9th projection est REF" not in tpes_chart
    assert "9th projection REF" not in tpes_chart


def test_build_dashboards_keeps_supply_flow_branch_for_single_child_sheet(tmp_path, monkeypatch):
    charts_dir = tmp_path / "charts"
    charts_dir.mkdir()

    def _fake_make_chart(
        sheet: str,
        fuel: str,
        sub: pd.DataFrame,
        charts_dir: Path,
        backend: str = "plotly",
        *,
        display_sheet: str | None = None,
        file_sheet: str | None = None,
    ) -> Path:
        chart_path = charts_dir / f"{dashboard_utils._safe_token(file_sheet or sheet)}__{dashboard_utils._safe_token(fuel)}.html"
        chart_path.write_text("<html><body>chart</body></html>", encoding="utf-8")
        return chart_path

    monkeypatch.setattr(dashboard_utils, "make_chart", _fake_make_chart)

    comparison_long = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "production primary",
                "measure": "Indigenous Production (Petajoules)",
                "fuel_label": "Total",
                "scenario": "Reference",
                "source": "leap",
                "year": 2022,
                "value": 100.0,
            },
            {
                "economy": "20_USA",
                "sheet": "imports primary",
                "measure": "Imports (Petajoules)",
                "fuel_label": "Total",
                "scenario": "Reference",
                "source": "leap",
                "year": 2022,
                "value": 50.0,
            },
            {
                "economy": "20_USA",
                "sheet": "exports primary",
                "measure": "Exports (Petajoules)",
                "fuel_label": "Total",
                "scenario": "Reference",
                "source": "leap",
                "year": 2022,
                "value": 20.0,
            },
        ]
    )
    mapping_status = pd.DataFrame(
        [
            {
                "sheet": "production primary",
                "measure": "Indigenous Production (Petajoules)",
                "fuel_label": "Total",
                "sector_code_9th": "01_production",
                "mapped": True,
                "base_mapping_complete": True,
                "projection_mapping_complete": True,
            },
            {
                "sheet": "imports primary",
                "measure": "Imports (Petajoules)",
                "fuel_label": "Total",
                "sector_code_9th": "02_imports",
                "mapped": True,
                "base_mapping_complete": True,
                "projection_mapping_complete": True,
            },
            {
                "sheet": "exports primary",
                "measure": "Exports (Petajoules)",
                "fuel_label": "Total",
                "sector_code_9th": "03_exports",
                "mapped": True,
                "base_mapping_complete": True,
                "projection_mapping_complete": True,
            },
        ]
    )

    dashboard_utils.build_dashboards(
        output_dir=tmp_path / "out",
        comparison_long=comparison_long,
        charts_dir=charts_dir,
        mapping_status=mapping_status,
    )

    supply_html = (tmp_path / "out" / "dashboards" / "node__Supply.html").read_text(encoding="utf-8")
    assert 'href="#sec-Primary_resource_production"' in supply_html
    assert 'href="#sec-Primary_resource_exports"' in supply_html
    assert 'href="#sec-Primary_resource_imports"' in supply_html
    assert 'class="jump-chip" data-level="2" data-kind="sheet">Primary resource production</a>' not in supply_html
    assert 'class="jump-chip" data-level="3" data-kind="sheet">Primary resource production</a>' in supply_html


def test_aggregate_display_rows_to_total_preserves_projection_sources_by_default():
    frame = pd.DataFrame(
        [
            {"economy": "20_USA", "sheet": "elecgen_inputs", "measure": "Electricity generation inputs (PJ)", "fuel_label": "Coal", "scenario": "Reference", "source": "projection", "year": 2030, "value": 100.0},
            {"economy": "20_USA", "sheet": "heat_inputs", "measure": "Transformation heat inputs (PJ)", "fuel_label": "Gas", "scenario": "Reference", "source": "projection_estimated", "year": 2030, "value": 10.0},
            {"economy": "20_USA", "sheet": "elecgen_inputs", "measure": "Electricity generation inputs (PJ)", "fuel_label": "Coal", "scenario": "", "source": "base", "year": 2022, "value": 50.0},
            {"economy": "20_USA", "sheet": "heat_inputs", "measure": "Transformation heat inputs (PJ)", "fuel_label": "Gas", "scenario": "Target", "source": "base_estimated", "year": 2022, "value": 5.0},
        ]
    )

    total = _aggregate_display_rows_to_total(
        frame,
        title="Power",
        measure_value="Summary (PJ)",
        collapse_base_family=True,
    )

    projection_sources = set(total.loc[total["year"].eq(2030), "source"].astype(str))
    assert projection_sources == {"projection", "projection_estimated"}

    base_rows = total.loc[total["year"].eq(2022)].copy()
    assert set(base_rows["source"].astype(str)) == {"base_mixed"}
    assert float(base_rows["value"].iloc[0]) == 50.0


def test_aggregate_display_rows_to_total_can_collapse_projection_family():
    frame = pd.DataFrame(
        [
            {"economy": "20_USA", "sheet": "Production", "measure": "Production (PJ)", "fuel_label": "Total", "scenario": "Reference", "source": "projection", "year": 2030, "value": 100.0},
            {"economy": "20_USA", "sheet": "Imports", "measure": "Imports (PJ)", "fuel_label": "Total", "scenario": "Reference", "source": "projection_estimated", "year": 2030, "value": 10.0},
        ]
    )

    total = _aggregate_display_rows_to_total(
        frame,
        title="Supply",
        measure_value="TPES excl. bunkers (PJ)",
        collapse_projection_family=True,
    )

    assert set(total["source"].astype(str)) == {"projection_mixed"}
    assert float(total["value"].iloc[0]) == 110.0


def test_prepare_render_long_derives_totals_without_projection_mixing():
    frame = pd.DataFrame(
        [
            {"economy": "20_USA", "sheet": "Buildings", "measure": "Demand (PJ)", "fuel_label": "Electricity", "scenario": "Reference", "source": "projection", "year": 2030, "value": 100.0},
            {"economy": "20_USA", "sheet": "Buildings", "measure": "Demand (PJ)", "fuel_label": "Gas", "scenario": "Reference", "source": "projection_estimated", "year": 2030, "value": 10.0},
            {"economy": "20_USA", "sheet": "Buildings", "measure": "Demand (PJ)", "fuel_label": "Electricity", "scenario": "", "source": "base", "year": 2022, "value": 50.0},
            {"economy": "20_USA", "sheet": "Buildings", "measure": "Demand (PJ)", "fuel_label": "Gas", "scenario": "Target", "source": "base_estimated", "year": 2022, "value": 5.0},
        ]
    )

    render_long = _prepare_render_long(frame)
    totals = render_long.loc[render_long["fuel_label"].eq("Total")].copy()

    projection_totals = totals.loc[totals["year"].eq(2030)].copy()
    assert set(projection_totals["source"].astype(str)) == {"projection", "projection_estimated"}
    assert set(projection_totals["value"].astype(float)) == {100.0, 10.0}

    base_totals = totals.loc[totals["year"].eq(2022)].copy()
    assert set(base_totals["source"].astype(str)) == {"base_mixed"}
    assert float(base_totals["value"].iloc[0]) == 50.0


def test_build_dashboards_routes_feedstock_output_pages_under_power(tmp_path, monkeypatch):
    charts_dir = tmp_path / "charts"
    charts_dir.mkdir()

    def _write_chart(sheet: str, measure: str, fuel: str) -> None:
        sheet_slug = dashboard_utils._safe_token(f"{sheet}__{measure}")
        fuel_slug = dashboard_utils._safe_token(fuel)
        (charts_dir / f"{sheet_slug}__{fuel_slug}.html").write_text("<html><body>chart</body></html>", encoding="utf-8")

    _write_chart("elecgen_out_feed", "Electricity generation outputs by feedstock (PJ)", "Biomass")
    _write_chart("heat_out_feed", "Transformation heat output by feedstock (PJ)", "Coal Bituminous")

    def _fake_make_chart(
        sheet: str,
        fuel: str,
        sub: pd.DataFrame,
        charts_dir: Path,
        backend: str = "plotly",
        *,
        display_sheet: str | None = None,
        file_sheet: str | None = None,
    ) -> Path:
        chart_path = charts_dir / f"{dashboard_utils._safe_token(file_sheet or sheet)}__{dashboard_utils._safe_token(fuel)}.html"
        chart_path.write_text("<html><body>chart</body></html>", encoding="utf-8")
        return chart_path

    monkeypatch.setattr(dashboard_utils, "make_chart", _fake_make_chart)

    comparison_long = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "elecgen_out_feed",
                "measure": "Electricity generation outputs by feedstock (PJ)",
                "fuel_label": "Biomass",
                "scenario": "Reference",
                "source": "leap",
                "year": 2023,
                "value": 10.0,
            },
            {
                "economy": "20_USA",
                "sheet": "heat_out_feed",
                "measure": "Transformation heat output by feedstock (PJ)",
                "fuel_label": "Coal Bituminous",
                "scenario": "Reference",
                "source": "leap",
                "year": 2023,
                "value": 5.0,
            },
        ]
    )
    mapping_status = pd.DataFrame(
        [
            {
                "sheet": "elecgen_out_feed",
                "measure": "Electricity generation outputs by feedstock (PJ)",
                "fuel_label": "Biomass",
                "sector_code_9th": "18_01_electricity_plants",
            },
            {
                "sheet": "heat_out_feed",
                "measure": "Transformation heat output by feedstock (PJ)",
                "fuel_label": "Coal Bituminous",
                "sector_code_9th": "18_02_chp_plants",
            },
        ]
    )

    dashboard_utils.build_dashboards(
        output_dir=tmp_path / "out",
        comparison_long=comparison_long,
        charts_dir=charts_dir,
        mapping_status=mapping_status,
    )

    dashboards_dir = tmp_path / "out" / "dashboards"
    assert not (dashboards_dir / "node__Electricity_output_in_GWh.html").exists()
    power_html = (dashboards_dir / "node__Power.html").read_text(encoding="utf-8")
    assert "Electricity generation outputs by feedstock" in power_html
    assert "Transformation heat output by feedstock" in power_html


def test_build_dashboards_matches_legacy_chart_stems_for_blank_measure_alias_sheets(tmp_path, monkeypatch):
    charts_dir = tmp_path / "charts"
    charts_dir.mkdir()
    (charts_dir / "Residential__Electricity.html").write_text("<html><body>chart</body></html>", encoding="utf-8")

    def _fake_make_chart(
        sheet: str,
        fuel: str,
        sub: pd.DataFrame,
        charts_dir: Path,
        backend: str = "plotly",
        *,
        display_sheet: str | None = None,
        file_sheet: str | None = None,
    ) -> Path:
        chart_path = charts_dir / f"{dashboard_utils._safe_token(file_sheet or sheet)}__{dashboard_utils._safe_token(fuel)}.html"
        chart_path.write_text("<html><body>chart</body></html>", encoding="utf-8")
        return chart_path

    monkeypatch.setattr(dashboard_utils, "make_chart", _fake_make_chart)

    comparison_long = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "Residential",
                "measure": "",
                "fuel_label": "Electricity",
                "scenario": "Reference",
                "source": "leap",
                "year": 2023,
                "value": 10.0,
            }
        ]
    )

    dashboard_utils.build_dashboards(
        output_dir=tmp_path / "out",
        comparison_long=comparison_long,
        charts_dir=charts_dir,
        mapping_status=None,
    )

    dashboards_dir = tmp_path / "out" / "dashboards"
    index_html = (dashboards_dir / "index.html").read_text(encoding="utf-8")
    assert "Buildings" in index_html
    assert (dashboards_dir / "node__Buildings.html").exists()


def test_build_charts_preserves_precomputed_total_rows(tmp_path, monkeypatch):
    charts_dir = tmp_path / "charts"
    charts_dir.mkdir()
    seen_subsets: dict[tuple[str, str], set[str]] = {}

    def _fake_make_chart(
        sheet: str,
        fuel: str,
        sub: pd.DataFrame,
        charts_dir: Path,
        backend: str = "plotly",
        *,
        display_sheet: str | None = None,
        file_sheet: str | None = None,
    ) -> Path:
        seen_subsets[(str(sheet), str(fuel))] = set(sub["source"].astype(str).tolist())
        chart_path = charts_dir / f"{dashboard_utils._safe_token(file_sheet or sheet)}__{dashboard_utils._safe_token(fuel)}.html"
        chart_path.write_text("<html><body>chart</body></html>", encoding="utf-8")
        return chart_path

    monkeypatch.setattr(dashboard_utils, "make_chart", _fake_make_chart)

    comparison_long = pd.DataFrame(
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
                "fuel_label": "Gas",
                "scenario": "Reference",
                "source": "projection_estimated",
                "year": 2030,
                "value": 25.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Buildings",
                "measure": "Demand (PJ)",
                "fuel_label": "Total",
                "scenario": "Reference",
                "source": "projection",
                "year": 2030,
                "value": 100.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Buildings",
                "measure": "Demand (PJ)",
                "fuel_label": "Total",
                "scenario": "Reference",
                "source": "projection_estimated",
                "year": 2030,
                "value": 25.0,
            },
        ]
    )

    dashboard_utils.build_charts(comparison_long=comparison_long, charts_dir=charts_dir, backend="plotly")

    assert seen_subsets[("Buildings", "Total")] == {"projection", "projection_estimated"}
    assert "projection_mixed" not in seen_subsets[("Buildings", "Total")]


def test_build_dashboards_collapses_demand_and_energy_totals_for_demand_nodes(tmp_path, monkeypatch):
    charts_dir = tmp_path / "charts"
    charts_dir.mkdir()
    (charts_dir / "Construction__Energy__PJ__Electricity.html").write_text("<html><body>chart</body></html>", encoding="utf-8")
    (charts_dir / "Chemical__incl__petrochemical__Electricity.html").write_text("<html><body>chart</body></html>", encoding="utf-8")

    def _fake_make_chart(
        sheet: str,
        fuel: str,
        sub: pd.DataFrame,
        charts_dir: Path,
        backend: str = "plotly",
        *,
        display_sheet: str | None = None,
        file_sheet: str | None = None,
    ) -> Path:
        chart_path = charts_dir / f"{dashboard_utils._safe_token(file_sheet or sheet)}__{dashboard_utils._safe_token(fuel)}.html"
        chart_path.write_text("<html><body>chart</body></html>", encoding="utf-8")
        return chart_path

    monkeypatch.setattr(dashboard_utils, "make_chart", _fake_make_chart)

    comparison_long = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "Construction",
                "measure": "Energy (PJ)",
                "fuel_label": "Electricity",
                "scenario": "Reference",
                "source": "leap",
                "year": 2023,
                "value": 10.0,
            },
            {
                "economy": "20_USA",
                "sheet": "Chemical (incl. petrochemical)",
                "measure": "",
                "fuel_label": "Electricity",
                "scenario": "Reference",
                "source": "leap",
                "year": 2023,
                "value": 8.0,
            },
        ]
    )
    mapping_status = pd.DataFrame(
        [
            {
                "sheet": "Construction",
                "measure": "Energy (PJ)",
                "fuel_label": "Electricity",
                "sector_code_9th": "14_02_construction",
            },
            {
                "sheet": "Chemical (incl. petrochemical)",
                "measure": "",
                "fuel_label": "Electricity",
                "sector_code_9th": "14_03_02_chemical_incl_petrochemical",
            },
        ]
    )

    dashboard_utils.build_dashboards(
        output_dir=tmp_path / "out",
        comparison_long=comparison_long,
        charts_dir=charts_dir,
        mapping_status=mapping_status,
    )

    industry_html = (tmp_path / "out" / "dashboards" / "node__Industry_sector.html").read_text(encoding="utf-8")
    assert "Demand (PJ)" in industry_html
    assert "Energy (PJ)" not in industry_html


def test_build_dashboards_places_map_power_sheets_under_power_output_nodes(tmp_path, monkeypatch):
    charts_dir = tmp_path / "charts"
    charts_dir.mkdir()
    (charts_dir / "MAP_electricity_plants__Total.html").write_text("<html><body>chart</body></html>", encoding="utf-8")
    (charts_dir / "MAP_CHP_plants__electricity__Total.html").write_text("<html><body>chart</body></html>", encoding="utf-8")

    def _fake_make_chart(
        sheet: str,
        fuel: str,
        sub: pd.DataFrame,
        charts_dir: Path,
        backend: str = "plotly",
        *,
        display_sheet: str | None = None,
        file_sheet: str | None = None,
    ) -> Path:
        chart_path = charts_dir / f"{dashboard_utils._safe_token(file_sheet or sheet)}__{dashboard_utils._safe_token(fuel)}.html"
        chart_path.write_text("<html><body>chart</body></html>", encoding="utf-8")
        return chart_path

    monkeypatch.setattr(dashboard_utils, "make_chart", _fake_make_chart)

    comparison_long = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "MAP electricity plants",
                "measure": "",
                "fuel_label": "Total",
                "scenario": "Reference",
                "source": "leap",
                "year": 2023,
                "value": 10.0,
            },
            {
                "economy": "20_USA",
                "sheet": "MAP CHP plants (electricity)",
                "measure": "",
                "fuel_label": "Total",
                "scenario": "Reference",
                "source": "leap",
                "year": 2023,
                "value": 5.0,
            },
        ]
    )

    dashboard_utils.build_dashboards(
        output_dir=tmp_path / "out",
        comparison_long=comparison_long,
        charts_dir=charts_dir,
        mapping_status=None,
    )

    power_html = (tmp_path / "out" / "dashboards" / "node__Power.html").read_text(encoding="utf-8")
    assert "Electricity plants (electricity output)" in power_html
    assert "CHP plants (electricity output)" in power_html
    assert '<optgroup label="MAP electricity plants">' not in power_html
    assert '<optgroup label="MAP CHP plants (electricity)">' not in power_html


def test_build_dashboards_places_transfer_process_sheets_under_other_transformation(tmp_path, monkeypatch):
    charts_dir = tmp_path / "charts"
    charts_dir.mkdir()

    def _fake_make_chart(
        sheet: str,
        fuel: str,
        sub: pd.DataFrame,
        charts_dir: Path,
        backend: str = "plotly",
        *,
        display_sheet: str | None = None,
        file_sheet: str | None = None,
    ) -> Path:
        chart_path = charts_dir / f"{dashboard_utils._safe_token(file_sheet or sheet)}__{dashboard_utils._safe_token(fuel)}.html"
        chart_path.write_text("<html><body>chart</body></html>", encoding="utf-8")
        return chart_path

    monkeypatch.setattr(dashboard_utils, "make_chart", _fake_make_chart)

    comparison_long = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "sheet": "transfers_out_fuel",
                "measure": "",
                "fuel_label": "Motor gasoline",
                "scenario": "Reference",
                "source": "leap",
                "year": 2023,
                "value": 10.0,
            },
            {
                "economy": "20_USA",
                "sheet": "transfers_unallocated_inputs",
                "measure": "",
                "fuel_label": "Crude oil",
                "scenario": "Reference",
                "source": "leap",
                "year": 2023,
                "value": 4.0,
            },
            {
                "economy": "20_USA",
                "sheet": "refinery_blending_inputs",
                "measure": "",
                "fuel_label": "Naphtha",
                "scenario": "Reference",
                "source": "leap",
                "year": 2023,
                "value": 3.0,
            },
            {
                "economy": "20_USA",
                "sheet": "upstream_liquids_inputs",
                "measure": "",
                "fuel_label": "Natural gas liquids",
                "scenario": "Reference",
                "source": "leap",
                "year": 2023,
                "value": 2.0,
            },
        ]
    )
    mapping_status = pd.DataFrame(
        [
            {
                "sheet": "transfers_out_fuel",
                "measure": "",
                "fuel_label": "Motor gasoline",
                "sector_code_9th": "08_transfers",
                "mapped": True,
                "base_mapping_complete": True,
                "projection_mapping_complete": True,
            },
            {
                "sheet": "transfers_unallocated_inputs",
                "measure": "",
                "fuel_label": "Crude oil",
                "sector_code_9th": "08_transfers",
                "mapped": True,
                "base_mapping_complete": True,
                "projection_mapping_complete": True,
            },
            {
                "sheet": "refinery_blending_inputs",
                "measure": "",
                "fuel_label": "Naphtha",
                "sector_code_9th": "08_transfers",
                "mapped": True,
                "base_mapping_complete": True,
                "projection_mapping_complete": True,
            },
            {
                "sheet": "upstream_liquids_inputs",
                "measure": "",
                "fuel_label": "Natural gas liquids",
                "sector_code_9th": "08_transfers",
                "mapped": True,
                "base_mapping_complete": True,
                "projection_mapping_complete": True,
            },
        ]
    )

    dashboard_utils.build_dashboards(
        output_dir=tmp_path / "out",
        comparison_long=comparison_long,
        charts_dir=charts_dir,
        mapping_status=mapping_status,
    )

    supply_path = tmp_path / "out" / "dashboards" / "node__Supply.html"
    supply_html = supply_path.read_text(encoding="utf-8") if supply_path.exists() else ""
    assert "Transformation transfer outputs by product" not in supply_html
    assert "Transfers unallocated inputs" not in supply_html
    assert "Refinery blending inputs" not in supply_html
    assert "Upstream liquids inputs" not in supply_html

    transfers_html = (tmp_path / "out" / "dashboards" / "node__Other_transformation__Transfers.html").read_text(
        encoding="utf-8"
    )
    assert "Transformation transfer outputs by product" in transfers_html
    assert "Transfers unallocated inputs" in transfers_html
    assert "Refinery blending inputs" in transfers_html
    assert "Upstream liquids inputs" in transfers_html
    assert ">    transfers_out_fuel<" not in transfers_html
    assert ">    transfers_unallocated_inputs<" not in transfers_html
    assert ">    refinery_blending_inputs<" not in transfers_html
    assert ">    upstream_liquids_inputs<" not in transfers_html


def test_build_comparisons_groups_multirow_explicit_mappings_without_projection_double_count() -> None:
    leap_long = pd.DataFrame(
        [
            {
                "economy": "USA",
                "scenario": "Reference",
                "sheet_name": "elecgen_inputs",
                "sector_code_9th": "09_01_electricity_plants",
                "sector_name": "Electricity plants",
                "fuel_label": "Coal Bituminous",
                "year": 2022,
                "leap_value": 100.0,
            },
            {
                "economy": "USA",
                "scenario": "Reference",
                "sheet_name": "elecgen_inputs",
                "sector_code_9th": "09_01_electricity_plants",
                "sector_name": "Electricity plants",
                "fuel_label": "Coal Bituminous",
                "year": 2023,
                "leap_value": 110.0,
            },
        ]
    )
    sheet_map = pd.DataFrame(
        [
            {
                "sheet_name": "elecgen_inputs",
                "sector_code_9th": "09_01_electricity_plants",
                "sector_name": "Electricity plants",
                "notes": "Electricity generation inputs",
                "category_type": "fuel",
            }
        ]
    )
    explicit_mappings = pd.DataFrame(
        [
            {
                "sheet_name": "elecgen_inputs",
                "fuel_label": "Coal Bituminous",
                "sector_code_9th": "09_01_electricity_plants",
                "ninth_fuel_code": "01_x_thermal_coal",
                "esto_flow": "09.01.01 Electricity plants",
                "esto_product": "01.02 Other bituminous coal",
                "mapping_note": "aggregate coal power inputs",
            },
            {
                "sheet_name": "elecgen_inputs",
                "fuel_label": "Coal Bituminous",
                "sector_code_9th": "09_01_electricity_plants",
                "ninth_fuel_code": "01_x_thermal_coal",
                "esto_flow": "09.01.01 Electricity plants",
                "esto_product": "01.03 Sub-bituminous coal",
                "mapping_note": "aggregate coal power inputs",
            },
            {
                "sheet_name": "elecgen_inputs",
                "fuel_label": "Coal Bituminous",
                "sector_code_9th": "09_01_electricity_plants",
                "ninth_fuel_code": "01_05_lignite",
                "esto_flow": "09.01.01 Electricity plants",
                "esto_product": "01.05 Lignite",
                "mapping_note": "aggregate coal power inputs",
            },
        ]
    )
    base_df = pd.DataFrame(
        [
            {
                "economy": "20USA",
                "flows": "09.01.01 Electricity plants",
                "products": "01.02 Other bituminous coal",
                "2022": -3.0,
            },
            {
                "economy": "20USA",
                "flows": "09.01.01 Electricity plants",
                "products": "01.03 Sub-bituminous coal",
                "2022": -4.0,
            },
            {
                "economy": "20USA",
                "flows": "09.01.01 Electricity plants",
                "products": "01.05 Lignite",
                "2022": -1.0,
            },
        ]
    )
    ninth_df = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "scenarios": "reference",
                "sectors": "09_transformation",
                "sub1sectors": "09_01_electricity_plants",
                "sub2sectors": "09_01_01_coal_power",
                "sub3sectors": "x",
                "sub4sectors": "x",
                "fuels": "01_coal",
                "subfuels": "01_x_thermal_coal",
                "subtotal_layout": False,
                "subtotal_results": False,
                "2023": -10.0,
            },
            {
                "economy": "20_USA",
                "scenarios": "reference",
                "sectors": "09_transformation",
                "sub1sectors": "09_01_electricity_plants",
                "sub2sectors": "09_01_01_coal_power",
                "sub3sectors": "x",
                "sub4sectors": "x",
                "fuels": "01_coal",
                "subfuels": "01_05_lignite",
                "subtotal_layout": False,
                "subtotal_results": False,
                "2023": -2.0,
            },
        ]
    )
    ninth_pairs = pd.DataFrame(
        [
            {
                "9th_sector": "09_01_01_coal_power",
                "9th_fuel": "01_x_thermal_coal",
                "esto_flow": "09.01.01 Electricity plants",
                "esto_product": "01.02 Other bituminous coal",
            },
            {
                "9th_sector": "09_01_01_coal_power",
                "9th_fuel": "01_x_thermal_coal",
                "esto_flow": "09.01.01 Electricity plants",
                "esto_product": "01.03 Sub-bituminous coal",
            },
            {
                "9th_sector": "09_01_01_coal_power",
                "9th_fuel": "01_05_lignite",
                "esto_flow": "09.01.01 Electricity plants",
                "esto_product": "01.05 Lignite",
            },
        ]
    )

    comparison_long, _, mapping_status = build_comparisons(
        leap_long=leap_long,
        sheet_map=sheet_map,
        fuel_mapping={},
        sector_flow_mapping={},
        ninth_pairs=ninth_pairs,
        base_df=base_df,
        ninth_df=ninth_df,
        explicit_mappings=explicit_mappings,
        base_year=2022,
        base_economy="20USA",
        projection_economy="20_USA",
        projection_years=(2023,),
        scenario_map={"reference": "reference"},
    )

    mapped = mapping_status.iloc[0]
    assert mapped["mapping_source"] == "explicit"
    assert mapped["ninth_fuel_code"] == "2 fuels (aggregated)"
    assert mapped["esto_product"] == "3 products (aggregated)"
    assert "aggregated explicit targets" in mapped["mapping_note"]
    assert "09_01_01_coal_power" in mapped["projection_targets_detail"]
    assert bool(mapped["mapped"])

    base_value = comparison_long.loc[comparison_long["source"] == "base", "value"].iloc[0]
    projection_value = comparison_long.loc[comparison_long["source"] == "projection", "value"].iloc[0]
    assert base_value == 8.0
    assert projection_value == 12.0


def test_explicit_power_input_aggregate_rows_exist() -> None:
    explicit = load_explicit_sector_fuel_mappings()
    power = explicit[
        (explicit["sheet_name"] == "elecgen_inputs")
        & (explicit["sector_code_9th"] == "09_01_electricity_plants")
    ].copy()

    coal = power[power["fuel_label"] == "Coal Bituminous"]
    biomass = power[power["fuel_label"] == "Biomass"]
    solar = power[power["fuel_label"] == "Solar"]

    assert set(coal["ninth_fuel_code"]) == {"01_x_thermal_coal", "01_05_lignite"}
    assert set(coal["esto_product"]) == {
        "01.02 Other bituminous coal",
        "01.03 Sub-bituminous coal",
        "01.04 Anthracite",
        "01.05 Lignite",
    }
    assert set(biomass["ninth_fuel_code"]) == {"15_solid_biomass_unallocated", "15_05_other_biomass"}
    assert set(biomass["esto_product"]) == {"15 Solid biomass"}
    assert set(solar["ninth_fuel_code"]) == {"12_01_of_which_photovoltaics", "12_x_other_solar"}
    assert set(solar["esto_product"]) == {"12 Solar"}


def test_explicit_projection_targets_stay_child_scoped_when_one_component_is_missing() -> None:
    leap_long = pd.DataFrame(
        [
            {
                "economy": "USA",
                "scenario": "Reference",
                "sheet_name": "elecgen_inputs",
                "sector_code_9th": "09_01_electricity_plants",
                "sector_name": "Electricity plants",
                "fuel_label": "Biomass",
                "year": 2022,
                "leap_value": 100.0,
            },
            {
                "economy": "USA",
                "scenario": "Reference",
                "sheet_name": "elecgen_inputs",
                "sector_code_9th": "09_01_electricity_plants",
                "sector_name": "Electricity plants",
                "fuel_label": "Biomass",
                "year": 2023,
                "leap_value": 100.0,
            },
        ]
    )
    sheet_map = pd.DataFrame(
        [
            {
                "sheet_name": "elecgen_inputs",
                "sector_code_9th": "09_01_electricity_plants",
                "sector_name": "Electricity plants",
                "notes": "Electricity generation inputs",
                "category_type": "fuel",
            }
        ]
    )
    explicit_mappings = pd.DataFrame(
        [
            {
                "sheet_name": "elecgen_inputs",
                "fuel_label": "Biomass",
                "sector_code_9th": "09_01_electricity_plants",
                "ninth_fuel_code": "15_solid_biomass_unallocated",
                "esto_flow": "09.01.01 Electricity plants",
                "esto_product": "15 Solid biomass",
                "mapping_note": "aggregate biomass power inputs",
            },
            {
                "sheet_name": "elecgen_inputs",
                "fuel_label": "Biomass",
                "sector_code_9th": "09_01_electricity_plants",
                "ninth_fuel_code": "15_05_other_biomass",
                "esto_flow": "09.01.01 Electricity plants",
                "esto_product": "15 Solid biomass",
                "mapping_note": "aggregate biomass power inputs",
            },
        ]
    )
    base_df = pd.DataFrame(
        [
            {
                "economy": "20USA",
                "flows": "09.01.01 Electricity plants",
                "products": "15 Solid biomass",
                "2022": -162.3,
            },
        ]
    )
    ninth_df = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "scenarios": "reference",
                "sectors": "09_total_transformation_sector",
                "sub1sectors": "09_01_electricity_plants",
                "sub2sectors": "09_01_06_biomass",
                "sub3sectors": "x",
                "sub4sectors": "x",
                "fuels": "15_solid_biomass",
                "subfuels": "15_solid_biomass_unallocated",
                "subtotal_layout": False,
                "subtotal_results": False,
                "2023": -147.176893,
            },
        ]
    )

    comparison_long, _, mapping_status = build_comparisons(
        leap_long=leap_long,
        sheet_map=sheet_map,
        fuel_mapping={},
        sector_flow_mapping={},
        ninth_pairs=pd.DataFrame(),
        base_df=base_df,
        ninth_df=ninth_df,
        explicit_mappings=explicit_mappings,
        base_year=2022,
        base_economy="20USA",
        projection_economy="20_USA",
        projection_years=(2023,),
        scenario_map={"reference": "reference"},
    )

    status = mapping_status.iloc[0]
    assert not bool(status["projection_parent_fallback"])
    projection_rows = comparison_long[
        (comparison_long["sheet"] == "elecgen_inputs")
        & (comparison_long["fuel_label"] == "Biomass")
        & (comparison_long["source"] == "projection")
    ]
    assert len(projection_rows) == 1
    assert projection_rows["value"].iloc[0] == 147.176893
    assert comparison_long[
        (comparison_long["sheet"] == "Electricity plants")
        & (comparison_long["fuel_label"] == "Biomass")
        & (comparison_long["source"] == "projection")
    ].empty


def test_shared_projection_bucket_split_is_marked_estimated() -> None:
    leap_long = pd.DataFrame(
        [
            {
                "economy": "USA",
                "scenario": "Reference",
                "sheet_name": "Food and tobacco",
                "sector_code_9th": "14_03_07_food_beverages_and_tobacco",
                "sector_name": "Food and tobacco",
                "fuel_label": "Anthracite",
                "year": 2022,
                "leap_value": 6.0,
            },
            {
                "economy": "USA",
                "scenario": "Reference",
                "sheet_name": "Food and tobacco",
                "sector_code_9th": "14_03_07_food_beverages_and_tobacco",
                "sector_name": "Food and tobacco",
                "fuel_label": "Anthracite",
                "year": 2023,
                "leap_value": 6.0,
            },
            {
                "economy": "USA",
                "scenario": "Reference",
                "sheet_name": "Food and tobacco",
                "sector_code_9th": "14_03_07_food_beverages_and_tobacco",
                "sector_name": "Food and tobacco",
                "fuel_label": "Other bituminous coal",
                "year": 2022,
                "leap_value": 6.0,
            },
            {
                "economy": "USA",
                "scenario": "Reference",
                "sheet_name": "Food and tobacco",
                "sector_code_9th": "14_03_07_food_beverages_and_tobacco",
                "sector_name": "Food and tobacco",
                "fuel_label": "Other bituminous coal",
                "year": 2023,
                "leap_value": 6.0,
            },
            {
                "economy": "USA",
                "scenario": "Reference",
                "sheet_name": "Food and tobacco",
                "sector_code_9th": "14_03_07_food_beverages_and_tobacco",
                "sector_name": "Food and tobacco",
                "fuel_label": "Sub bituminous coal",
                "year": 2022,
                "leap_value": 6.0,
            },
            {
                "economy": "USA",
                "scenario": "Reference",
                "sheet_name": "Food and tobacco",
                "sector_code_9th": "14_03_07_food_beverages_and_tobacco",
                "sector_name": "Food and tobacco",
                "fuel_label": "Sub bituminous coal",
                "year": 2023,
                "leap_value": 6.0,
            },
        ]
    )
    sheet_map = pd.DataFrame(
        [
            {
                "sheet_name": "Food and tobacco",
                "sector_code_9th": "14_03_07_food_beverages_and_tobacco",
                "sector_name": "Food and tobacco",
                "notes": "Industry demand",
                "category_type": "fuel",
            }
        ]
    )
    fuel_mapping = {
        "anthracite": {"esto_product": "01.04 Anthracite", "ninth_fuel": ""},
        "other bituminous coal": {"esto_product": "01.02 Other bituminous coal", "ninth_fuel": ""},
        "sub bituminous coal": {"esto_product": "01.03 Sub-bituminous coal", "ninth_fuel": ""},
    }
    ninth_pairs = pd.DataFrame(
        [
            {
                "9th_sector": "14_03_07_food_beverages_and_tobacco",
                "9th_fuel": "01_x_thermal_coal",
                "esto_flow": "14.03.07 Food, beverages and tobacco",
                "esto_product": "01.02 Other bituminous coal",
            },
            {
                "9th_sector": "14_03_07_food_beverages_and_tobacco",
                "9th_fuel": "01_x_thermal_coal",
                "esto_flow": "14.03.07 Food, beverages and tobacco",
                "esto_product": "01.03 Sub-bituminous coal",
            },
            {
                "9th_sector": "14_03_07_food_beverages_and_tobacco",
                "9th_fuel": "01_x_thermal_coal",
                "esto_flow": "14.03.07 Food, beverages and tobacco",
                "esto_product": "01.04 Anthracite",
            },
        ]
    )
    base_df = pd.DataFrame(
        [
            {
                "economy": "20USA",
                "flows": "14.03.07 Food, beverages and tobacco",
                "products": "01.04 Anthracite",
                "2022": -1.0,
            },
            {
                "economy": "20USA",
                "flows": "14.03.07 Food, beverages and tobacco",
                "products": "01.02 Other bituminous coal",
                "2022": -3.0,
            },
            {
                "economy": "20USA",
                "flows": "14.03.07 Food, beverages and tobacco",
                "products": "01.03 Sub-bituminous coal",
                "2022": -6.0,
            },
        ]
    )
    ninth_df = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "scenarios": "reference",
                "sectors": "14_industry_sector",
                "sub1sectors": "14_03_manufacturing",
                "sub2sectors": "14_03_07_food_beverages_and_tobacco",
                "sub3sectors": "x",
                "sub4sectors": "x",
                "fuels": "01_coal",
                "subfuels": "01_x_thermal_coal",
                "subtotal_layout": False,
                "subtotal_results": False,
                "2023": -10.0,
            },
        ]
    )

    comparison_long, _, _ = build_comparisons(
        leap_long=leap_long,
        sheet_map=sheet_map,
        fuel_mapping=fuel_mapping,
        sector_flow_mapping={},
        ninth_pairs=ninth_pairs,
        base_df=base_df,
        ninth_df=ninth_df,
        explicit_mappings=pd.DataFrame(),
        base_year=2022,
        base_economy="20USA",
        projection_economy="20_USA",
        projection_years=(2023,),
        scenario_map={"reference": "reference"},
    )

    projection_rows = comparison_long[
        (comparison_long["sheet"] == "Food and tobacco")
        & (comparison_long["source"] == "projection_estimated")
        & (comparison_long["year"] == 2023)
    ].copy()
    assert set(projection_rows["fuel_label"]) == {
        "Anthracite",
        "Other bituminous coal",
        "Sub bituminous coal",
    }
    values = projection_rows.set_index("fuel_label")["value"].to_dict()
    assert values["Anthracite"] == 1.0
    assert values["Other bituminous coal"] == 3.0
    assert values["Sub bituminous coal"] == 6.0
    assert projection_rows["value"].sum() == 10.0


def test_shared_base_target_split_is_marked_estimated() -> None:
    leap_long = pd.DataFrame(
        [
            {
                "economy": "USA",
                "scenario": "Reference",
                "sheet_name": "hydrogen_out_fuel",
                "sector_code_9th": "09_13_hydrogen_transformation",
                "sector_name": "Hydrogen transformation",
                "fuel_label": "Hydrogen",
                "year": 2023,
                "leap_value": 6.0,
            },
            {
                "economy": "USA",
                "scenario": "Reference",
                "sheet_name": "hydrogen_out_fuel",
                "sector_code_9th": "09_13_hydrogen_transformation",
                "sector_name": "Hydrogen transformation",
                "fuel_label": "Efuel",
                "year": 2023,
                "leap_value": 3.0,
            },
            {
                "economy": "USA",
                "scenario": "Reference",
                "sheet_name": "hydrogen_out_fuel",
                "sector_code_9th": "09_13_hydrogen_transformation",
                "sector_name": "Hydrogen transformation",
                "fuel_label": "Ammonia",
                "year": 2023,
                "leap_value": 0.0,
            },
        ]
    )
    sheet_map = pd.DataFrame(
        [
            {
                "sheet_name": "hydrogen_out_fuel",
                "sector_code_9th": "09_13_hydrogen_transformation",
                "sector_name": "Hydrogen transformation",
                "notes": "Hydrogen outputs by fuel",
                "category_type": "fuel",
            }
        ]
    )
    fuel_mapping = {
        "hydrogen": {
            "ninth_fuel": "16_x_hydrogen",
            "esto_product": "16.09 Other sources",
            "esto_flow": "",
            "mapping_source": "codebook_fallback",
            "flow_source": "",
            "fuel_source": "inferred",
        },
        "efuel": {
            "ninth_fuel": "16_x_efuel",
            "esto_product": "16.09 Other sources",
            "esto_flow": "",
            "mapping_source": "codebook_fallback",
            "flow_source": "",
            "fuel_source": "inferred",
        },
        "ammonia": {
            "ninth_fuel": "16_x_ammonia",
            "esto_product": "16.09 Other sources",
            "esto_flow": "",
            "mapping_source": "codebook_fallback",
            "flow_source": "",
            "fuel_source": "inferred",
        },
    }
    ninth_pairs = pd.DataFrame(
        [
            {
                "9th_sector": "09_13_hydrogen_transformation",
                "9th_fuel": "16_x_hydrogen",
                "esto_flow": "09.12 Non-specified transformation",
                "esto_product": "16.09 Other sources",
            },
            {
                "9th_sector": "09_13_hydrogen_transformation",
                "9th_fuel": "16_x_efuel",
                "esto_flow": "09.12 Non-specified transformation",
                "esto_product": "16.09 Other sources",
            },
            {
                "9th_sector": "09_13_hydrogen_transformation",
                "9th_fuel": "16_x_ammonia",
                "esto_flow": "09.12 Non-specified transformation",
                "esto_product": "16.09 Other sources",
            },
        ]
    )
    base_df = pd.DataFrame(
        [
            {
                "economy": "20USA",
                "flows": "09.12 Non-specified transformation",
                "products": "16.09 Other sources",
                "2022": 9.0,
            },
        ]
    )
    ninth_df = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "scenarios": "reference",
                "sectors": "09_total_transformation_sector",
                "sub1sectors": "09_13_hydrogen_transformation",
                "sub2sectors": "x",
                "sub3sectors": "x",
                "sub4sectors": "x",
                "fuels": "16_others",
                "subfuels": "16_x_hydrogen",
                "subtotal_layout": False,
                "subtotal_results": False,
                "2023": 6.0,
            },
            {
                "economy": "20_USA",
                "scenarios": "reference",
                "sectors": "09_total_transformation_sector",
                "sub1sectors": "09_13_hydrogen_transformation",
                "sub2sectors": "x",
                "sub3sectors": "x",
                "sub4sectors": "x",
                "fuels": "16_others",
                "subfuels": "16_x_efuel",
                "subtotal_layout": False,
                "subtotal_results": False,
                "2023": 3.0,
            },
            {
                "economy": "20_USA",
                "scenarios": "reference",
                "sectors": "09_total_transformation_sector",
                "sub1sectors": "09_13_hydrogen_transformation",
                "sub2sectors": "x",
                "sub3sectors": "x",
                "sub4sectors": "x",
                "fuels": "16_others",
                "subfuels": "16_x_ammonia",
                "subtotal_layout": False,
                "subtotal_results": False,
                "2023": 0.0,
            },
        ]
    )

    comparison_long, _, _ = build_comparisons(
        leap_long=leap_long,
        sheet_map=sheet_map,
        fuel_mapping=fuel_mapping,
        sector_flow_mapping={},
        ninth_pairs=ninth_pairs,
        base_df=base_df,
        ninth_df=ninth_df,
        explicit_mappings=pd.DataFrame(),
        base_year=2022,
        base_economy="20USA",
        projection_economy="20_USA",
        projection_years=(2023,),
        scenario_map={"reference": "reference"},
    )

    base_rows = comparison_long[
        (comparison_long["sheet"] == "hydrogen_out_fuel")
        & (comparison_long["source"] == "base_estimated")
        & (comparison_long["year"] == 2022)
    ].copy()
    assert set(base_rows["fuel_label"]) == {"Hydrogen", "Efuel", "Ammonia"}
    values = base_rows.set_index("fuel_label")["value"].to_dict()
    assert values["Hydrogen"] == 6.0
    assert values["Efuel"] == 3.0
    assert values["Ammonia"] == 0.0
    assert base_rows["value"].sum() == 9.0
