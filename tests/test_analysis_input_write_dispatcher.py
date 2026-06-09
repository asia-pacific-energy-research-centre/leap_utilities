from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

from codebase.configuration import workflow_config as workflow_cfg
from codebase.functions import analysis_input_write_dispatcher as dispatcher
from codebase.functions import leap_core
from codebase.utilities import workflow_common


def _write_leap_like_workbook(
    path: Path,
    *,
    sheet_name: str = "LEAP",
    data_rows: list[dict[str, object]] | None = None,
) -> None:
    columns = [
        "Branch Path",
        "Variable",
        "Scenario",
        "Region",
        "Scale",
        "Units",
        "Per...",
        "Expression",
        "Level 1",
        "Level 2",
    ]
    header0 = {col: "" for col in columns}
    header0["Branch Path"] = "Area:"
    header0["Variable"] = "test_model"
    header0["Scenario"] = "Ver:"
    header0["Region"] = "2"
    header1 = {col: "" for col in columns}
    rows = data_rows or []
    frame = pd.DataFrame(rows, columns=columns)
    output = pd.concat(
        [
            pd.DataFrame([header0], columns=columns),
            pd.DataFrame([header1], columns=columns),
            pd.DataFrame([columns], columns=columns),
            frame,
        ],
        ignore_index=True,
    )
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
        output.to_excel(writer, sheet_name=sheet_name, index=False, header=False)


def _write_mapping_workbook(path: Path, rows: list[dict[str, object]]) -> None:
    columns = [
        "enabled",
        "match_scope",
        "branch_path",
        "variable",
        "units",
        "scale",
        "per",
        "confidence",
        "notes",
    ]
    frame = pd.DataFrame(rows, columns=columns)
    with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
        frame.to_excel(writer, sheet_name="field_mappings", index=False)


def _read_units(path: Path, *, sheet_name: str = "LEAP") -> str:
    _, _, _, data = dispatcher._read_workbook_sheet(path, sheet_name)
    return str(data["Units"].iloc[0]).strip()


def test_get_analysis_input_write_mode_validation(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_WRITE_MODE", None, raising=False)
    assert dispatcher.get_analysis_input_write_mode() == "api"

    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_WRITE_MODE", "api", raising=False)
    assert dispatcher.get_analysis_input_write_mode() == "api"

    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_WRITE_MODE", "workbook", raising=False)
    assert dispatcher.get_analysis_input_write_mode() == "workbook"

    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_WRITE_MODE", "invalid-mode", raising=False)
    with pytest.raises(ValueError, match="Invalid ANALYSIS_INPUT_WRITE_MODE"):
        dispatcher.get_analysis_input_write_mode()


def test_dispatch_analysis_input_write_api_passthrough(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "api_passthrough.xlsx"
    _write_leap_like_workbook(
        workbook,
        data_rows=[
            {
                "Branch Path": r"Demand\A",
                "Variable": "Activity Level",
                "Scenario": "Reference",
                "Region": "United States",
                "Scale": "",
                "Units": "Petajoule",
                "Per...": "",
                "Expression": "Data(2022,1)",
                "Level 1": "Demand",
                "Level 2": "A",
            }
        ],
    )
    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_WRITE_MODE", "api", raising=False)
    called = {"count": 0}

    def _callback() -> str:
        called["count"] += 1
        return "ok"

    result = dispatcher.dispatch_analysis_input_write(
        export_path=workbook,
        sheet_name="LEAP",
        scenario="Reference",
        region="United States",
        context_label="test_api_passthrough",
        run_api_write=_callback,
    )
    assert result["mode"] == "api"
    assert called["count"] == 1
    assert result["api_result"] == "ok"


def test_workbook_mode_mapping_overrides_template(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "workbook_mode.xlsx"
    template = tmp_path / "template.xlsx"
    mapping = tmp_path / "mapping.xlsx"

    _write_leap_like_workbook(
        workbook,
        data_rows=[
            {
                "Branch Path": r"Demand\A",
                "Variable": "Activity Level",
                "Scenario": "Reference",
                "Region": "United States",
                "Scale": "",
                "Units": "OldUnit",
                "Per...": "",
                "Expression": "Data(2022,1)",
                "Level 1": "Demand",
                "Level 2": "A",
            }
        ],
    )
    _write_leap_like_workbook(
        template,
        data_rows=[
            {
                "Branch Path": r"Demand\A",
                "Variable": "Activity Level",
                "Scenario": "Reference",
                "Region": "United States",
                "Scale": "",
                "Units": "TemplateUnit",
                "Per...": "",
                "Expression": "Data(2022,1)",
                "Level 1": "Demand",
                "Level 2": "A",
            }
        ],
    )
    _write_mapping_workbook(
        mapping,
        rows=[
            {
                "enabled": True,
                "match_scope": "variable",
                "branch_path": "",
                "variable": "Activity Level",
                "units": "MappedUnit",
                "scale": "",
                "per": "",
                "confidence": "known",
                "notes": "test",
            }
        ],
    )

    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_WRITE_MODE", "workbook", raising=False)
    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_FIELD_MAPPING_PATH", mapping, raising=False)
    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_FIELD_MAPPING_SHEET", "field_mappings", raising=False)
    monkeypatch.setattr(
        workflow_cfg,
        "ANALYSIS_INPUT_CANONICAL_TEMPLATE_PATHS",
        [template],
        raising=False,
    )

    result = dispatcher.dispatch_analysis_input_write(
        export_path=workbook,
        sheet_name="LEAP",
        scenario="Reference",
        region="United States",
        context_label="test_workbook_mode",
    )
    assert result["mode"] == "workbook"
    assert Path(result["summary_path"]).exists()
    assert _read_units(workbook) == "MappedUnit"


def test_workbook_mode_fails_fast_on_unresolved_required_fields(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "workbook_unresolved.xlsx"
    canonical_template = tmp_path / "canonical_template.xlsx"
    mapping = tmp_path / "mapping_empty.xlsx"

    _write_leap_like_workbook(
        workbook,
        data_rows=[
            {
                "Branch Path": r"Demand\B",
                "Variable": "Unknown Variable",
                "Scenario": "Reference",
                "Region": "United States",
                "Scale": "",
                "Units": "",
                "Per...": "",
                "Expression": "Data(2022,1)",
                "Level 1": "Demand",
                "Level 2": "B",
            }
        ],
    )
    _write_leap_like_workbook(
        canonical_template,
        data_rows=[
            {
                "Branch Path": r"Demand\C",
                "Variable": "Activity Level",
                "Scenario": "Reference",
                "Region": "United States",
                "Scale": "",
                "Units": "TemplateUnit",
                "Per...": "",
                "Expression": "Data(2022,1)",
                "Level 1": "Demand",
                "Level 2": "C",
            }
        ],
    )
    _write_mapping_workbook(mapping, rows=[])

    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_WRITE_MODE", "workbook", raising=False)
    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_FIELD_MAPPING_PATH", mapping, raising=False)
    monkeypatch.setattr(
        workflow_cfg,
        "ANALYSIS_INPUT_CANONICAL_TEMPLATE_PATHS",
        [canonical_template],
        raising=False,
    )

    with pytest.raises(ValueError, match="Fields needing confirmation"):
        dispatcher.validate_workbook_for_manual_import(
            workbook,
            sheet_name="LEAP",
            scenario="Reference",
            region="United States",
        )


def test_mapping_match_precedence_branch_variable_variable_branch_template(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "precedence.xlsx"
    template = tmp_path / "template.xlsx"
    mapping = tmp_path / "mapping.xlsx"

    _write_leap_like_workbook(
        workbook,
        data_rows=[
            {
                "Branch Path": r"Demand\A",
                "Variable": "Activity Level",
                "Scenario": "Reference",
                "Region": "United States",
                "Scale": "",
                "Units": "",
                "Per...": "",
                "Expression": "Data(2022,1)",
                "Level 1": "Demand",
                "Level 2": "A",
            },
            {
                "Branch Path": r"Demand\B",
                "Variable": "Activity Level",
                "Scenario": "Reference",
                "Region": "United States",
                "Scale": "",
                "Units": "",
                "Per...": "",
                "Expression": "Data(2022,2)",
                "Level 1": "Demand",
                "Level 2": "B",
            },
        ],
    )
    _write_leap_like_workbook(
        template,
        data_rows=[
            {
                "Branch Path": r"Demand\A",
                "Variable": "Activity Level",
                "Scenario": "Reference",
                "Region": "United States",
                "Scale": "",
                "Units": "TemplateUnitA",
                "Per...": "",
                "Expression": "Data(2022,1)",
                "Level 1": "Demand",
                "Level 2": "A",
            },
            {
                "Branch Path": r"Demand\B",
                "Variable": "Activity Level",
                "Scenario": "Reference",
                "Region": "United States",
                "Scale": "",
                "Units": "TemplateUnitB",
                "Per...": "",
                "Expression": "Data(2022,2)",
                "Level 1": "Demand",
                "Level 2": "B",
            },
        ],
    )
    _write_mapping_workbook(
        mapping,
        rows=[
            {
                "enabled": True,
                "match_scope": "branch_variable",
                "branch_path": r"Demand\A",
                "variable": "Activity Level",
                "units": "BranchVariableUnit",
                "scale": "",
                "per": "",
                "confidence": "known",
                "notes": "",
            },
            {
                "enabled": True,
                "match_scope": "variable",
                "branch_path": "",
                "variable": "Activity Level",
                "units": "VariableUnit",
                "scale": "",
                "per": "",
                "confidence": "known",
                "notes": "",
            },
            {
                "enabled": True,
                "match_scope": "branch",
                "branch_path": r"Demand\B",
                "variable": "",
                "units": "BranchUnit",
                "scale": "",
                "per": "",
                "confidence": "known",
                "notes": "",
            },
        ],
    )

    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_WRITE_MODE", "workbook", raising=False)
    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_FIELD_MAPPING_PATH", mapping, raising=False)
    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_FIELD_MAPPING_SHEET", "field_mappings", raising=False)
    monkeypatch.setattr(
        workflow_cfg,
        "ANALYSIS_INPUT_CANONICAL_TEMPLATE_PATHS",
        [template],
        raising=False,
    )

    summary = dispatcher.validate_workbook_for_manual_import(
        workbook,
        sheet_name="LEAP",
        scenario="Reference",
        region="United States",
    )
    _, _, _, data = dispatcher._read_workbook_sheet(workbook, "LEAP")
    units_by_branch = {
        str(row["Branch Path"]).strip(): str(row["Units"]).strip()
        for _, row in data.iterrows()
    }
    assert units_by_branch[r"Demand\A"] == "BranchVariableUnit"
    assert units_by_branch[r"Demand\B"] == "VariableUnit"
    assert summary["fields_needing_confirmation"] == []


def test_workbook_mode_low_level_write_guards(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_WRITE_MODE", "workbook", raising=False)

    with pytest.raises(RuntimeError, match="reads are disabled in workbook mode"):
        dispatcher.ensure_analysis_view_api_read_allowed("test_read_guard")

    with pytest.raises(RuntimeError, match="reads are disabled in workbook mode"):
        leap_core.safe_branch_call(
            leap_obj=None,
            branch_path=r"Demand\\A",
        )

    with pytest.raises(RuntimeError, match="disabled in workbook mode"):
        leap_core.safe_set_variable(
            L=None,
            obj=None,
            varname="Activity Level",
            expr="Data(2022,1)",
        )

    with pytest.raises(RuntimeError, match="disabled in workbook mode"):
        leap_core.create_branches_from_export_file(
            L=None,
            leap_export_filename="dummy.xlsx",
        )

    with pytest.raises(RuntimeError, match="disabled in workbook mode"):
        leap_core.fill_branches_from_export_file(
            L=None,
            leap_export_filename="dummy.xlsx",
        )


def test_workflow_common_import_skips_api_writes_in_workbook_mode(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "workbook_mode_import.xlsx"
    template = tmp_path / "template.xlsx"
    mapping = tmp_path / "mapping.xlsx"
    _write_leap_like_workbook(
        workbook,
        data_rows=[
            {
                "Branch Path": r"Demand\A",
                "Variable": "Activity Level",
                "Scenario": "Reference",
                "Region": "United States",
                "Scale": "",
                "Units": "",
                "Per...": "",
                "Expression": "Data(2022,1)",
                "Level 1": "Demand",
                "Level 2": "A",
            }
        ],
    )
    _write_leap_like_workbook(
        template,
        data_rows=[
            {
                "Branch Path": r"Demand\A",
                "Variable": "Activity Level",
                "Scenario": "Reference",
                "Region": "United States",
                "Scale": "",
                "Units": "TemplateUnit",
                "Per...": "",
                "Expression": "Data(2022,1)",
                "Level 1": "Demand",
                "Level 2": "A",
            }
        ],
    )
    _write_mapping_workbook(
        mapping,
        rows=[
            {
                "enabled": True,
                "match_scope": "variable",
                "branch_path": "",
                "variable": "Activity Level",
                "units": "MappedUnit",
                "scale": "",
                "per": "",
                "confidence": "known",
                "notes": "",
            }
        ],
    )
    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_WRITE_MODE", "workbook", raising=False)
    monkeypatch.setattr(workflow_cfg, "ANALYSIS_INPUT_FIELD_MAPPING_PATH", mapping, raising=False)
    monkeypatch.setattr(
        workflow_cfg,
        "ANALYSIS_INPUT_CANONICAL_TEMPLATE_PATHS",
        [template],
        raising=False,
    )

    def _unexpected_call(*args, **kwargs):
        raise AssertionError("API write function should not be called in workbook mode.")

    monkeypatch.setattr(workflow_common, "connect_to_leap", _unexpected_call)
    monkeypatch.setattr(workflow_common, "create_branches_from_export_file", _unexpected_call)
    monkeypatch.setattr(workflow_common, "fill_branches_from_export_file", _unexpected_call)

    result_path = workflow_common.import_workbook_to_leap(
        export_path=workbook,
        sheet_name="LEAP",
        scenario="Reference",
        region="United States",
        create_branches=True,
        fill_branches=True,
        include_current_accounts=False,
    )
    assert result_path == workbook
