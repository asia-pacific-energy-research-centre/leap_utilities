from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import pytest

from codebase import results_supply_link_workflow as workflow
from codebase.configuration import workflow_config as workflow_cfg


def _minimal_reconciliation_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "scenario": "Reference",
                "esto_product": "17 Electricity",
                "year": 2030,
                "adjusted_imports": 2.0,
                "max_transformation_output": 20.0,
                "constrained_transformation_output": 5.0,
            }
        ]
    )


def _write_balance_table_csv(
    path: Path,
    *,
    observed_imports: float,
    observed_exports: float = 0.0,
) -> None:
    pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "scenario": "Reference",
                "year": 2030,
                "esto_product": "17 Electricity",
                "balance_component": "adjusted_imports",
                "value": observed_imports,
            },
            {
                "economy": "20_USA",
                "scenario": "Reference",
                "year": 2030,
                "esto_product": "17 Electricity",
                "balance_component": "adjusted_exports",
                "value": -abs(observed_exports),
            },
        ]
    ).to_csv(path, index=False)


def test_current_accounts_resolves_to_target_when_reference_absent() -> None:
    reconciliation = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "scenario": "Target",
                "esto_product": "17 Electricity",
                "year": 2030,
            }
        ]
    )
    assert (
        workflow._resolve_reconciliation_scenario_key(reconciliation, "Current Accounts")
        == "target"
    )


def test_build_supply_overrides_capacity_unmet_iterative(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(workflow, "TRADE_TARGET_EXPORT_MODE", "capacity_unmet_iterative", raising=False)
    reconciliation = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "scenario": "Reference",
                "esto_product": "17 Electricity",
                "year": 2030,
                "adjusted_imports": 12.0,
                "adjusted_exports": 3.5,
                "projected_exports": 4.0,
            }
        ]
    )
    overrides = workflow.build_supply_overrides(reconciliation)
    payload = overrides["20_USA"]["Reference"]["17 Electricity"]
    assert payload["imports"][2030] == pytest.approx(0.0)
    assert payload["exports"][2030] == pytest.approx(3.5)


def test_build_supply_overrides_capacity_unmet_iterative_balanced(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(workflow, "TRADE_TARGET_EXPORT_MODE", "capacity_unmet_iterative_balanced", raising=False)
    monkeypatch.setattr(
        workflow,
        "_CAPACITY_UNMET_RUNTIME_EXPORT_ADJUSTMENTS",
        {"20_usa|reference|17 electricity|2030": 1.25},
        raising=False,
    )
    monkeypatch.setattr(
        workflow,
        "_CAPACITY_UNMET_RUNTIME_PRIMARY_ADDITIONS",
        {"20_usa|reference|17 electricity|2030": 2.0},
        raising=False,
    )
    reconciliation = pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "scenario": "Reference",
                "esto_product": "17 Electricity",
                "year": 2030,
                "adjusted_imports": 12.0,
                "adjusted_exports": 3.5,
                "projected_exports": 4.0,
                "constrained_production": 5.0,
                "max_production": 20.0,
            }
        ]
    )
    overrides = workflow.build_supply_overrides(reconciliation)
    payload = overrides["20_USA"]["Reference"]["17 Electricity"]
    assert payload["imports"][2030] == pytest.approx(0.0)
    assert payload["exports"][2030] == pytest.approx(4.75)
    assert payload["production"][2030] == pytest.approx(7.0)
    assert payload["max_production"][2030] == pytest.approx(20.0)


def test_capacity_unmet_iterative_requires_workbook_mode(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(workflow, "TRADE_TARGET_EXPORT_MODE", "capacity_unmet_iterative", raising=False)
    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_WRITE_MODE", "api", raising=False)
    with pytest.raises(ValueError, match="iterative unmet modes require"):
        workflow.run_results_linked_transformation_supply_workflow(
            economies=["20_USA"],
            scenario_names=["Reference"],
            include_leap_import=False,
            use_direct_leap_results_for_demand=False,
            scrape_leap_results_for_demand=False,
        )


def test_capacity_unmet_iterative_balanced_requires_workbook_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(workflow, "TRADE_TARGET_EXPORT_MODE", "capacity_unmet_iterative_balanced", raising=False)
    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_WRITE_MODE", "api", raising=False)
    with pytest.raises(ValueError, match="iterative unmet modes require"):
        workflow.run_results_linked_transformation_supply_workflow(
            economies=["20_USA"],
            scenario_names=["Reference"],
            include_leap_import=False,
            use_direct_leap_results_for_demand=False,
            scrape_leap_results_for_demand=False,
        )


def test_capacity_unmet_iterative_same_results_guard(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    reconciliation = _minimal_reconciliation_df()
    state_path = tmp_path / "state.json"
    balance_csv = tmp_path / "balance_table_20_USA_25042026_REF_2030.csv"
    _write_balance_table_csv(balance_csv, observed_imports=10.0)
    signature_payload = {
        "source": "balance_tables",
        "files": [workflow._build_results_signature(balance_csv)],
    }
    state_payload = {
        "version": 1,
        "cumulative_capacity_additions": {},
        "cumulative_output_additions": {},
        "last_results_signatures": {
            "20_usa|reference": signature_payload
        },
        "passes": [],
    }
    state_path.write_text(json.dumps(state_payload), encoding="utf-8")

    process_catalog = pd.DataFrame(
        [
            {
                "record_index": 0,
                "economy": "20_USA",
                "module": "Electricity generation",
                "process": "Gas plants",
                "instance": 1,
                "esto_product": "17 Electricity",
                "year": 2030,
                "product_output": 10.0,
                "module_total_output": 20.0,
                "yield": 0.5,
            }
        ]
    )
    monkeypatch.setattr(workflow, "_build_capacity_process_catalog", lambda records: (process_catalog, []))
    monkeypatch.setattr(workflow, "_build_label_to_esto_product_lookup", lambda: {})
    workflow._run_capacity_unmet_iterative_pass(
        reconciliation_table=reconciliation,
        process_records=[{}],
        economies=["20_USA"],
        scenarios=["Reference"],
        results_dir=[balance_csv],
        state_path=state_path,
        allow_same_results_reuse=False,
    )
    assert "detected no new LEAP results artifacts" in capsys.readouterr().out


def test_collect_observed_trade_prefers_balance_tables(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    balance_dir = tmp_path / "balance_tables"
    balance_dir.mkdir()
    pd.DataFrame(
        [
            {
                "economy": "20_USA",
                "scenario": "Reference",
                "year": 2030,
                "esto_product": "17 Electricity",
                "balance_component": "adjusted_imports",
                "value": 8.0,
            },
            {
                "economy": "20_USA",
                "scenario": "Reference",
                "year": 2030,
                "esto_product": "17 Electricity",
                "balance_component": "adjusted_exports",
                "value": 1.5,
            },
        ]
    ).to_csv(balance_dir / "balance_table_2030.csv", index=False)

    def _fail_if_legacy_lookup_used(**kwargs):
        raise AssertionError("legacy workbook lookup should not run when balance tables exist")

    monkeypatch.setattr(workflow, "_select_supply_results_workbook", _fail_if_legacy_lookup_used)
    monkeypatch.setattr(workflow, "_read_supply_results_import_sheet", _fail_if_legacy_lookup_used)
    monkeypatch.setattr(workflow, "_read_supply_results_export_sheet", _fail_if_legacy_lookup_used)

    observed, signature_map, unmatched = workflow._collect_observed_trade_from_supply_results(
        scenario_pairs=[("20_USA", "reference")],
        label_to_product={},
        results_dir=balance_dir,
        include_exports=True,
    )

    assert unmatched == []
    assert len(signature_map) == 1
    assert observed["observed_imports"].sum() == pytest.approx(8.0)
    assert observed["observed_exports"].sum() == pytest.approx(1.5)


def test_save_year_balance_tables_writes_csv_and_archives_old_dates(tmp_path: Path) -> None:
    reconciliation = _minimal_reconciliation_df()
    output_dir = tmp_path / "yearly_balance_tables"
    old_csv = output_dir / "balance_table_20_USA_01012026_REF_2030.csv"
    old_xlsx = output_dir / "balance_table_20_USA_01012026_REF_2030.xlsx"
    output_dir.mkdir()
    old_csv.write_text("old", encoding="utf-8")
    old_xlsx.write_text("old", encoding="utf-8")

    paths = workflow.save_year_balance_tables(
        reconciliation,
        years=[2030],
        output_dir=output_dir,
        economies=["20_USA"],
        scenarios=["Reference"],
    )
    csv_paths = sorted(output_dir.glob("balance_table_20_USA_*_REF_2030.csv"))
    xlsx_paths = sorted(output_dir.glob("balance_table_20_USA_*_REF_2030.xlsx"))
    assert len(csv_paths) == 1
    assert xlsx_paths == []
    assert any(path.suffix == ".csv" for path in paths)
    assert not any(path.suffix == ".xlsx" for path in paths)
    assert output_dir.joinpath("archive", old_csv.name).exists()
    assert output_dir.joinpath("archive", old_xlsx.name).exists()
    existing_text = paths[0].read_text(encoding="utf-8")
    second_paths = workflow.save_year_balance_tables(
        reconciliation,
        years=[2030],
        output_dir=output_dir,
        economies=["20_USA"],
        scenarios=["Reference"],
    )
    assert second_paths == paths
    assert paths[0].read_text(encoding="utf-8") == existing_text


def test_capacity_unmet_iterative_allocates_and_persists(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    reconciliation = _minimal_reconciliation_df()
    state_path = tmp_path / "state.json"
    balance_csv = tmp_path / "balance_table_20_USA_25042026_REF_2030.csv"
    _write_balance_table_csv(balance_csv, observed_imports=8.0)

    process_catalog = pd.DataFrame(
        [
            {
                "record_index": 0,
                "economy": "20_USA",
                "module": "Electricity generation",
                "process": "Gas plants",
                "instance": 1,
                "esto_product": "17 Electricity",
                "year": 2030,
                "product_output": 10.0,
                "module_total_output": 20.0,
                "yield": 0.5,
            }
        ]
    )
    monkeypatch.setattr(workflow, "_build_capacity_process_catalog", lambda records: (process_catalog, []))
    monkeypatch.setattr(workflow, "_build_label_to_esto_product_lookup", lambda: {})
    summary = workflow._run_capacity_unmet_iterative_pass(
        reconciliation_table=reconciliation,
        process_records=[{}],
        economies=["20_USA"],
        scenarios=["Reference"],
        results_dir=[balance_csv],
        state_path=state_path,
        allow_same_results_reuse=False,
    )
    assert summary["allocated_output_total"] == pytest.approx(6.0)
    assert summary["clipped_output_total"] == pytest.approx(0.0)
    payload = json.loads(state_path.read_text(encoding="utf-8"))
    cumulative = payload["cumulative_capacity_additions"]
    assert len(cumulative) == 1
    only_value = list(cumulative.values())[0]
    assert float(only_value) == pytest.approx(12.0)
