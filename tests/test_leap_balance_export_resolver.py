from __future__ import annotations

from pathlib import Path

from codebase.utilities.leap_balance_export_resolver import resolve_balance_export_workbook


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
