from __future__ import annotations

import os
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from codebase.utilities import leap_results_dashboard_balance as balance_mod
from codebase.leap_results_dashboard_balance_workflow import run_workflow
from codebase.utilities.leap_results_dashboard_v2.comparison_engine import build_total_component_ledger


def _write_minimal_balance_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    *,
    scenario: str,
    year: int,
    units: str,
    flow: str = "Production",
    fuel: str = "Total",
    value: float = 1.0,
) -> None:
    ws.cell(1, 1, 'LEAP Area "Test Economy"')
    ws.cell(2, 1, f"Scenario: {scenario}, Year: {year}, Units: {units}")
    ws.cell(3, 2, fuel)
    ws.cell(4, 1, flow)
    ws.cell(4, 2, value)


def _make_minimal_workbook(path: Path, *, units_2023: str = "Petajoule", value_2023: float = 1.0) -> None:
    wb = openpyxl.Workbook()
    ws_2023 = wb.active
    ws_2023.title = "EBal|2023"
    _write_minimal_balance_sheet(
        ws_2023,
        scenario="Reference",
        year=2023,
        units=units_2023,
        value=value_2023,
    )

    ws_2022 = wb.create_sheet("EBal|2022")
    _write_minimal_balance_sheet(
        ws_2022,
        scenario="Reference",
        year=2022,
        units="Petajoule",
        value=2.0,
    )
    wb.save(path)


def test_extract_balance_workbook_combines_year_sheets_and_metadata(tmp_path: Path) -> None:
    workbook_path = tmp_path / "mini_balance.xlsx"
    _make_minimal_workbook(workbook_path)

    out = balance_mod._extract_balance_workbook(
        workbook_path,
        template_sheet="EBal|2023",
        mapping_pairs_path=balance_mod.DEFAULT_MAPPING_PAIRS_PATH,
        codebook_path=balance_mod.DEFAULT_CODEBOOK_PATH,
    )

    mapped = out["mapped_long"]
    assert set(pd.to_numeric(mapped["year"], errors="coerce").dropna().astype(int).unique()) == {2022, 2023}
    assert {"EBal|2022", "EBal|2023"}.issubset(set(mapped["source_sheet"].astype(str)))
    assert out["report"]["summary"]["selected_sheet_count"] == 2


def test_extract_balance_workbook_converts_units_to_petajoule(tmp_path: Path) -> None:
    workbook_path = tmp_path / "mini_balance_units.xlsx"
    _make_minimal_workbook(workbook_path, units_2023="Gigajoule", value_2023=1000.0)

    out = balance_mod._extract_balance_workbook(
        workbook_path,
        template_sheet="EBal|2023",
        mapping_pairs_path=balance_mod.DEFAULT_MAPPING_PAIRS_PATH,
        codebook_path=balance_mod.DEFAULT_CODEBOOK_PATH,
    )

    mapped = out["mapped_long"].copy()
    row_2023 = mapped[pd.to_numeric(mapped["year"], errors="coerce").eq(2023)].iloc[0]
    assert row_2023["units_petajoule"] == "Petajoule"
    assert pytest.approx(float(row_2023["value_petajoule"]), rel=1e-9, abs=1e-12) == 0.001


def test_load_balance_leap_long_filters_full_mappings_and_is_deterministic(monkeypatch: pytest.MonkeyPatch) -> None:
    ref_rows = pd.DataFrame(
        [
            {
                "scenario": "ref",
                "year": 2022,
                "leap_sector": "flow_a",
                "leap_fuel": "fuel_a",
                "esto_flow": "esto_f1",
                "esto_product": "esto_p1",
                "leap_sector_name": "Flow A",
                "leap_fuel_name": "Fuel A",
                "value_petajoule": 1.0,
                "source_sheet": "EBal|2022",
                "source_workbook": "ref.xlsx",
            },
            {
                "scenario": "ref",
                "year": 2022,
                "leap_sector": "flow_a",
                "leap_fuel": "fuel_a",
                "esto_flow": "esto_f1",
                "esto_product": "esto_p1",
                "leap_sector_name": "Flow A",
                "leap_fuel_name": "Fuel A",
                "value_petajoule": 2.0,
                "source_sheet": "EBal|2022",
                "source_workbook": "ref.xlsx",
            },
            {
                "scenario": "ref",
                "year": 2022,
                "leap_sector": "flow_incomplete",
                "leap_fuel": "fuel_a",
                "esto_flow": "esto_f_incomplete",
                "esto_product": "",
                "leap_sector_name": "Flow Incomplete",
                "leap_fuel_name": "Fuel A",
                "value_petajoule": 3.0,
                "source_sheet": "EBal|2022",
                "source_workbook": "ref.xlsx",
            },
            {
                "scenario": "ref",
                "year": 2023,
                "leap_sector": "flow_conflict",
                "leap_fuel": "fuel_a",
                "esto_flow": "esto_conflict_1",
                "esto_product": "esto_p1",
                "leap_sector_name": "Flow Conflict",
                "leap_fuel_name": "Fuel A",
                "value_petajoule": 1.0,
                "source_sheet": "EBal|2023",
                "source_workbook": "ref.xlsx",
            },
            {
                "scenario": "ref",
                "year": 2023,
                "leap_sector": "flow_conflict",
                "leap_fuel": "fuel_a",
                "esto_flow": "esto_conflict_2",
                "esto_product": "esto_p1",
                "leap_sector_name": "Flow Conflict",
                "leap_fuel_name": "Fuel A",
                "value_petajoule": 1.0,
                "source_sheet": "EBal|2023",
                "source_workbook": "ref.xlsx",
            },
            {
                "scenario": "ref",
                "year": 2022,
                "leap_sector": "flow_unmapped",
                "leap_fuel": "fuel_a",
                "esto_flow": "esto_fu",
                "esto_product": "esto_pu",
                "leap_sector_name": "Flow Unmapped",
                "leap_fuel_name": "Fuel A",
                "value_petajoule": 4.0,
                "source_sheet": "EBal|2022",
                "source_workbook": "ref.xlsx",
            },
        ]
    )
    tgt_rows = pd.DataFrame(
        [
            {
                "scenario": "tgt",
                "year": 2022,
                "leap_sector": "flow_a",
                "leap_fuel": "fuel_a",
                "esto_flow": "esto_f1",
                "esto_product": "esto_p1",
                "leap_sector_name": "Flow A",
                "leap_fuel_name": "Fuel A",
                "value_petajoule": 5.0,
                "source_sheet": "EBal|2022",
                "source_workbook": "tgt.xlsx",
            }
        ]
    )

    calls = {"n": 0}

    def _fake_extract(*args, **kwargs):  # noqa: ANN002, ANN003
        calls["n"] += 1
        rows = ref_rows if calls["n"] == 1 else tgt_rows
        return {
            "template_sheet": "EBal|2023",
            "raw_long": rows.copy(),
            "mapped_long": rows.copy(),
            "coverage": pd.DataFrame(),
            "unit_diag": pd.DataFrame(),
            "report": {"summary": {"selected_sheet_count": 2}},
        }

    monkeypatch.setattr(balance_mod, "_extract_balance_workbook", _fake_extract)

    structure = {
        "sheet_catalog": {
            "sheet_a": {"measure": "Energy balance (PJ)"},
            "sheet_conflict": {"measure": "Energy balance (PJ)"},
        },
        "flow_to_sheet": {
            "flow_a": "sheet_a",
            "flow_conflict": "sheet_conflict",
        },
    }

    result1 = balance_mod.load_balance_leap_long(
        ref_workbook_path=balance_mod.DEFAULT_REF_WORKBOOK_PATH,
        tgt_workbook_path=balance_mod.DEFAULT_TGT_WORKBOOK_PATH,
        mapping_pairs_path=balance_mod.DEFAULT_MAPPING_PAIRS_PATH,
        codebook_path=balance_mod.DEFAULT_CODEBOOK_PATH,
        structure_config=structure,
        known_issues={"mapping_overrides": [], "label_overrides": {}, "row_filters": {}},
        projection_economy="20_USA",
    )
    calls["n"] = 0
    result2 = balance_mod.load_balance_leap_long(
        ref_workbook_path=balance_mod.DEFAULT_REF_WORKBOOK_PATH,
        tgt_workbook_path=balance_mod.DEFAULT_TGT_WORKBOOK_PATH,
        mapping_pairs_path=balance_mod.DEFAULT_MAPPING_PAIRS_PATH,
        codebook_path=balance_mod.DEFAULT_CODEBOOK_PATH,
        structure_config=structure,
        known_issues={"mapping_overrides": [], "label_overrides": {}, "row_filters": {}},
        projection_economy="20_USA",
    )

    leap_long = result1["leap_long"]
    assert set(leap_long["scenario"].unique()) == {"Reference", "Target"}
    ref_sum = leap_long[
        leap_long["scenario"].eq("Reference")
        & leap_long["sector_code_9th"].eq("flow_a")
        & leap_long["year"].eq(2022)
    ]["leap_value"].iloc[0]
    assert float(ref_sum) == 3.0

    issue_reasons = set(result1["issues"]["reason"].dropna().astype(str))
    assert "incomplete_mapping" in issue_reasons
    assert "mapping_conflict_after_aggregation" in issue_reasons
    assert "flow_not_in_structure_config" in issue_reasons

    pd.testing.assert_frame_equal(
        result1["leap_long"].reset_index(drop=True),
        result2["leap_long"].reset_index(drop=True),
    )


def test_load_balance_leap_long_applies_overrides_only_from_json(monkeypatch: pytest.MonkeyPatch) -> None:
    rows = pd.DataFrame(
        [
            {
                "scenario": "ref",
                "year": 2022,
                "leap_sector": "flow_x",
                "leap_fuel": "fuel_x",
                "esto_flow": "",
                "esto_product": "",
                "leap_sector_name": "Flow X",
                "leap_fuel_name": "Fuel X",
                "value_petajoule": 1.0,
                "source_sheet": "EBal|2022",
                "source_workbook": "ref.xlsx",
            }
        ]
    )

    def _fake_extract(*args, **kwargs):  # noqa: ANN002, ANN003
        return {
            "template_sheet": "EBal|2022",
            "raw_long": rows.copy(),
            "mapped_long": rows.copy(),
            "coverage": pd.DataFrame(),
            "unit_diag": pd.DataFrame(),
            "report": {"summary": {"selected_sheet_count": 1}},
        }

    monkeypatch.setattr(balance_mod, "_extract_balance_workbook", _fake_extract)

    structure = {
        "sheet_catalog": {"sheet_x": {"measure": "Energy balance (PJ)"}},
        "flow_to_sheet": {"flow_x": "sheet_x"},
    }

    no_override = balance_mod.load_balance_leap_long(
        ref_workbook_path=balance_mod.DEFAULT_REF_WORKBOOK_PATH,
        tgt_workbook_path=balance_mod.DEFAULT_TGT_WORKBOOK_PATH,
        mapping_pairs_path=balance_mod.DEFAULT_MAPPING_PAIRS_PATH,
        codebook_path=balance_mod.DEFAULT_CODEBOOK_PATH,
        structure_config=structure,
        known_issues={"mapping_overrides": [], "label_overrides": {}, "row_filters": {}},
        projection_economy="20_USA",
    )
    assert no_override["leap_long"].empty

    with_override = balance_mod.load_balance_leap_long(
        ref_workbook_path=balance_mod.DEFAULT_REF_WORKBOOK_PATH,
        tgt_workbook_path=balance_mod.DEFAULT_TGT_WORKBOOK_PATH,
        mapping_pairs_path=balance_mod.DEFAULT_MAPPING_PAIRS_PATH,
        codebook_path=balance_mod.DEFAULT_CODEBOOK_PATH,
        structure_config=structure,
        known_issues={
            "mapping_overrides": [
                {
                    "active": True,
                    "match": {"leap_flow": "flow_x", "leap_product": "fuel_x"},
                    "set": {"esto_flow": "esto_flow_x", "esto_product": "esto_product_x"},
                }
            ],
            "label_overrides": {},
            "row_filters": {},
        },
        projection_economy="20_USA",
    )
    assert not with_override["leap_long"].empty
    assert int(with_override["override_report"]["applied_rows"].sum()) > 0


def test_render_balance_dashboards_respects_structure_and_empty_notice(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def _fake_build_charts(comparison_long, charts_dir, backend="plotly", hide_leap_only_charts=False):  # noqa: ANN001
        charts_dir.mkdir(parents=True, exist_ok=True)
        p = charts_dir / "SheetMapped__Energy_balance__PJ__Coal.html"
        p.write_text("<html><body>chart</body></html>", encoding="utf-8")
        return [p]

    monkeypatch.setattr(balance_mod, "build_charts", _fake_build_charts)

    comparison_long = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "scenario": "Reference",
                "sheet": "SheetMapped",
                "measure": "Energy balance (PJ)",
                "fuel_label": "Coal",
                "source": "leap",
                "year": 2022,
                "value": 1.0,
            }
        ]
    )
    structure = {
        "page_tree": [
            {
                "id": "group",
                "label": "Group",
                "children": [
                    {"id": "mapped", "label": "Mapped", "children": []},
                    {"id": "empty", "label": "Empty", "children": []},
                ],
            }
        ],
        "sheet_catalog": {
            "SheetMapped": {
                "display_label": "SheetMapped",
                "path": ["Group", "Mapped"],
                "measure": "Energy balance (PJ)",
                "sort_order": 0,
            }
        },
        "flow_to_sheet": {},
        "empty_page_notice": "No mapped data on this page.",
    }

    out = balance_mod.render_balance_dashboards(
        comparison_long=comparison_long,
        mapping_status=pd.DataFrame(),
        structure_config=structure,
        output_dir=tmp_path,
        chart_backend="plotly",
        hide_leap_only_charts=False,
    )

    dashboard_index = Path(out["dashboard_index"])
    assert dashboard_index.exists()

    empty_page = Path(out["dashboards_dir"]) / balance_mod._page_filename_from_path(("Group", "Empty"))
    assert empty_page.exists()
    html = empty_page.read_text(encoding="utf-8")
    assert "No mapped data on this page." in html


def test_template_esto_axis_records_strips_code_prefix_from_fuel_label() -> None:
    """Template-declared fuel_label must match the LEAP-observed path (code prefix stripped).

    LEAP-observed rows set fuel_label via _strip_esto_code_prefix(esto_product), e.g.
    esto_product="08.01 Natural gas" → fuel_label="Natural gas".
    Template rows must produce the same label so the two paths deduplicate correctly and
    the safety guard does not fire on phantom duplicate ESTO pairs.
    """
    template = {
        "defaults": {"measure": "Energy balance (PJ)"},
        "Energy": {
            "graphs": {
                "natural_gas": {
                    "esto_flow": "08 Natural gas",
                    "products": ["08.01 Natural gas", "Total"],
                }
            }
        },
    }

    import unittest.mock as mock

    with mock.patch.object(balance_mod, "_load_dashboard_template_allowlist", return_value=template):
        result = balance_mod._dashboard_template_esto_axis_records(
            None,
            scenario_names=["Reference"],
        )

    assert not result.empty, "Expected at least one row from template"
    coded_label_rows = result[result["fuel_label"].str.contains(r"^\d{2}", regex=True)]
    assert coded_label_rows.empty, (
        f"fuel_label must not contain ESTO code prefixes; got: {coded_label_rows['fuel_label'].tolist()}"
    )
    natural_gas = result[result["esto_product"].str.contains("Natural gas", case=False)]
    assert not natural_gas.empty, "Expected a Natural gas row"
    assert (natural_gas["fuel_label"] == "Natural gas").all(), (
        f"Expected fuel_label='Natural gas', got: {natural_gas['fuel_label'].tolist()}"
    )
    total_rows = result[result["esto_product"] == "Total"]
    assert not total_rows.empty, "Expected a Total row"
    assert (total_rows["fuel_label"] == "Total").all(), (
        f"Expected fuel_label='Total' for Total product, got: {total_rows['fuel_label'].tolist()}"
    )


def test_template_esto_axis_records_emit_page_and_chart_group_fields() -> None:
    template = {
        "defaults": {"measure": "Energy balance (PJ)"},
        "Buildings": {
            "Commercial": {
                "graphs": {
                    "Electricity": {
                        "esto_flow": "16.01 Commercial and public services",
                        "products": ["17 Electricity"],
                    }
                }
            }
        },
    }

    import unittest.mock as mock

    with mock.patch.object(balance_mod, "_load_dashboard_template_allowlist", return_value=template):
        result = balance_mod._dashboard_template_esto_axis_records(
            None,
            scenario_names=["Reference"],
        )

    row = result.iloc[0]
    assert row["sheet"] == "esto__16_01__Commercial_and_public_services"
    assert row["page_label"] == "Buildings"
    assert row["page_key"] == "buildings"
    assert row["chart_group_label"] == "Commercial"
    assert row["chart_group_key"]


def test_render_balance_dashboards_exposes_multiple_chart_groups_on_one_page(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def _fake_build_charts(comparison_long, charts_dir, backend="plotly", hide_leap_only_charts=False):  # noqa: ANN001
        charts_dir.mkdir(parents=True, exist_ok=True)
        written = []
        for (sheet, measure, fuel), _ in comparison_long.groupby(["sheet", "measure", "fuel_label"]):
            file_name = f"{balance_mod._safe_token(f'{sheet}__{measure}')}__{balance_mod._safe_token(fuel)}.html"
            path = charts_dir / file_name
            path.write_text("<html><body>chart</body></html>", encoding="utf-8")
            written.append(path)
        return written

    monkeypatch.setattr(balance_mod, "build_charts", _fake_build_charts)

    template = {
        "defaults": {"measure": "Energy balance (PJ)"},
        "Buildings": {
            "Commercial": {
                "graphs": {"Electricity": {"esto_flow": "16.01 Commercial and public services", "products": ["17 Electricity"]}}
            },
            "Residential": {
                "graphs": {"Electricity": {"esto_flow": "16.02 Residential", "products": ["17 Electricity"]}}
            },
        },
    }
    structure = balance_mod.build_esto_axis_structure_from_dashboard_template(None)
    with monkeypatch.context() as m:
        m.setattr(balance_mod, "_load_dashboard_template_allowlist", lambda _path: template)
        structure = balance_mod.build_esto_axis_structure_from_dashboard_template(None)
        comparison_long = pd.DataFrame(
            [
                {
                    "economy": "20_USA",
                    "scenario": "Target",
                    "sheet": "esto__16_01__Commercial_and_public_services",
                    "measure": "Energy balance (PJ)",
                    "fuel_label": "Electricity",
                    "source": "leap",
                    "year": 2030,
                    "value": 1.0,
                },
                {
                    "economy": "20_USA",
                    "scenario": "Target",
                    "sheet": "esto__16_02__Residential",
                    "measure": "Energy balance (PJ)",
                    "fuel_label": "Electricity",
                    "source": "leap",
                    "year": 2030,
                    "value": 2.0,
                },
            ]
        )
        mapping_status = pd.DataFrame(
            [
                {
                    "sheet": "esto__16_01__Commercial_and_public_services",
                    "measure": "Energy balance (PJ)",
                    "fuel_label": "Electricity",
                    "esto_flow": "16.01 Commercial and public services",
                    "esto_product": "17 Electricity",
                },
                {
                    "sheet": "esto__16_02__Residential",
                    "measure": "Energy balance (PJ)",
                    "fuel_label": "Electricity",
                    "esto_flow": "16.02 Residential",
                    "esto_product": "17 Electricity",
                },
            ]
        )
        out = balance_mod.render_balance_dashboards(
            comparison_long=comparison_long,
            mapping_status=mapping_status,
            structure_config=structure,
            output_dir=tmp_path,
            chart_backend="plotly",
            chart_navigation_guide_path=None,
        )

    exposure = pd.read_csv(out["chart_group_exposure"])
    assert set(exposure["page_label"]) == {"Buildings"}
    assert {"Commercial", "Residential"}.issubset(set(exposure["chart_group_label"]))
    assert exposure["chart_group_key"].nunique() == 2


def test_total_component_ledger_dedup_scopes_to_chart_group_key() -> None:
    chart_rows = pd.DataFrame(
        [
            {"economy": "20_USA", "sheet": "Legacy", "page_key": "buildings", "page_label": "Buildings", "chart_group_key": "commercial", "chart_group_label": "Commercial", "measure": "Energy balance (PJ)", "fuel_label": "Electricity", "scenario": "Target", "source": "projection", "year": 2030, "value": 1.0},
            {"economy": "20_USA", "sheet": "Legacy", "page_key": "buildings", "page_label": "Buildings", "chart_group_key": "residential", "chart_group_label": "Residential", "measure": "Energy balance (PJ)", "fuel_label": "Electricity", "scenario": "Target", "source": "projection", "year": 2030, "value": 2.0},
        ]
    )
    mapping_status = pd.DataFrame(
        [
            {"sheet": "Legacy", "chart_group_key": "commercial", "measure": "Energy balance (PJ)", "fuel_label": "Electricity", "sector_code_9th": "16_01_commercial", "ninth_fuel_code": "17_electricity"},
            {"sheet": "Legacy", "chart_group_key": "residential", "measure": "Energy balance (PJ)", "fuel_label": "Electricity", "sector_code_9th": "16_02_residential", "ninth_fuel_code": "17_electricity"},
        ]
    )

    ledger = build_total_component_ledger(chart_rows, mapping_status)

    assert set(ledger["chart_group_key"]) == {"commercial", "residential"}
    assert set(ledger["duplicate_exact_comparator_key_count"]) == {1}


def test_attach_chart_groups_uses_legacy_sheet_fallback(tmp_path: Path) -> None:
    chart_groups_path = tmp_path / "chart_group_exposure.csv"
    pd.DataFrame(
        [
            {
                "chart_group_id": "chart::legacy.html",
                "dashboard_path": "Legacy Page",
                "sheet": "LegacySheet",
                "measure": "Energy balance (PJ)",
                "fuel_label": "Coal",
                "chart_file": "charts/legacy.html",
                "section_id": "sec-legacy",
                "section_label": "Legacy Page",
                "entry_kind": "direct",
            }
        ]
    ).to_csv(chart_groups_path, index=False)
    exposure = pd.DataFrame(
        [
            {
                "sheet": "LegacySheet",
                "measure": "Energy balance (PJ)",
                "fuel_label": "Coal",
                "source": "base",
                "year": 2022,
                "value": 1.0,
            }
        ]
    )

    out = balance_mod.attach_chart_groups_to_dashboard_exposure(exposure, chart_groups_path)

    assert out["chart_group_id"].iloc[0] == "chart::legacy.html"
    assert out["chart_group_key"].iloc[0]


@pytest.mark.integration
def test_balance_dashboard_workflow_integration_usa_inputs() -> None:
    if os.getenv("RUN_BALANCE_INTEGRATION", "").strip().lower() not in {"1", "true", "yes", "y"}:
        pytest.skip("Set RUN_BALANCE_INTEGRATION=1 to run full integration workflow test.")

    result = run_workflow()
    out_dir = Path(result["comparison_long"]).resolve().parent

    required = [
        out_dir / "comparison_long.csv",
        out_dir / "comparison_wide.csv",
        out_dir / "mapping_status.xlsx",
        out_dir / "chart_line_mapping_ledger.csv",
        out_dir / "chart_total_component_ledger.csv",
        out_dir / "comparison_gap_diagnostics.csv",
        out_dir / "dashboards/index.html",
    ]
    for path in required:
        assert path.exists(), f"Missing expected output: {path}"

    comparison_long = pd.read_csv(out_dir / "comparison_long.csv")
    assert {"leap", "base", "projection"}.issubset(set(comparison_long["source"].dropna().astype(str).unique()))
    years = pd.to_numeric(comparison_long["year"], errors="coerce").dropna().astype(int)
    assert int(years.min()) <= 2022
    assert int(years.max()) >= 2060

    leap_long = pd.read_csv(out_dir / "leap_long.csv")
    assert set(leap_long["leap_units"].dropna().astype(str).unique()) == {"Petajoule"}
