from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook

from codebase.utilities.leap_balance_export_resolver import (
    load_leap_balance_activity_table,
    resolve_balance_export_workbook,
)


def _touch(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("", encoding="utf-8")


def test_resolve_balance_export_workbook_uses_latest_date_id(tmp_path: Path) -> None:
    export_dir = tmp_path / "20_USA"
    _touch(export_dir / "full model output all years 492026 TGT.xlsx")
    expected = export_dir / "full model output all years 4212026 TGT.xlsx"
    _touch(expected)

    resolved = resolve_balance_export_workbook(
        economy="20_USA",
        scenario="Target",
        exports_root=tmp_path,
    )

    assert resolved == expected


def test_resolve_balance_export_workbook_honors_explicit_date_id(tmp_path: Path) -> None:
    export_dir = tmp_path / "20_USA"
    expected = export_dir / "full model output all years 492026 REF.xlsx"
    _touch(expected)
    _touch(export_dir / "full model output all years 4212026 REF.xlsx")

    resolved = resolve_balance_export_workbook(
        economy="20_USA",
        scenario="ref",
        date_id="492026",
        exports_root=tmp_path,
    )

    assert resolved == expected


def test_resolve_balance_export_workbook_reports_missing_match(tmp_path: Path) -> None:
    try:
        resolve_balance_export_workbook(
            economy="20_USA",
            scenario="REF",
            exports_root=tmp_path,
        )
    except FileNotFoundError as exc:
        assert "20_USA" in str(exc)
        assert "REF" in str(exc)
    else:
        raise AssertionError("missing balance-export workbook did not raise")


def _write_balance_workbook(path: Path, *, units: str, electricity_value: float) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "EBal|2060"
    sheet.append(['Energy Balance for Area "Test"', None, None])
    sheet.append([f"Scenario: Target, Year: 2060, Units: {units}", None, None])
    sheet.append([None, "Electricity", "Natural gas"])
    sheet.append(["Imports", electricity_value, 2.0])
    sheet.append(["Production", 3.0, 4.0])
    workbook.save(path)


def test_load_leap_balance_activity_table_normalizes_thousand_petajoule_to_pj(tmp_path: Path) -> None:
    pj_path = tmp_path / "pj.xlsx"
    thousand_pj_path = tmp_path / "thousand_pj.xlsx"
    _write_balance_workbook(pj_path, units="Petajoule", electricity_value=1200.0)
    _write_balance_workbook(thousand_pj_path, units="Thousand Petajoule", electricity_value=1.2)

    pj = load_leap_balance_activity_table(
        pj_path,
        balance_rows=["Imports"],
        fuels=["Electricity"],
    )
    thousand_pj = load_leap_balance_activity_table(
        thousand_pj_path,
        balance_rows=["Imports"],
        fuels=["Electricity"],
    )

    assert pj.loc[0, "value"] == 1200.0
    assert thousand_pj.loc[0, "value"] == 1200.0
