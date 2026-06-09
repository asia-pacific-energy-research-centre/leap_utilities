from pathlib import Path

import pandas as pd

from codebase import leap_exports


def test_build_workbook_filename_formats_economy_and_scenarios() -> None:
    filename = leap_exports.build_workbook_filename(
        economy_label="ALL ECONOMIES",
        scenarios=["Reference", "Target", "Current Accounts"],
        template="transformation_leap_imports_{economy}_{scenario}.xlsx",
    )
    assert (
        filename
        == "transformation_leap_imports_ALL_ECONOMIES_Reference_Target_Current_Accounts.xlsx"
    )


def test_find_and_validate_export_helpers(tmp_path: Path) -> None:
    old_path = tmp_path / "transfer_leap_imports_001.xlsx"
    new_path = tmp_path / "transfer_leap_imports_002.xlsx"
    old_path.touch()

    df = pd.DataFrame(
        {
            "Branch Path": ["Demand\\A"],
            "Variable": ["Activity Level"],
            "Scenario": ["Target"],
            "Region": ["United States"],
        }
    )
    with pd.ExcelWriter(new_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="LEAP", index=False)

    found = leap_exports.find_workbook(
        directory=tmp_path,
        prefix="transfer_leap_imports_",
    )
    assert found == new_path
    assert leap_exports.list_scenarios(new_path, sheet_name="LEAP") == ["Target"]
    leap_exports.validate_region(new_path, "United States", sheet_name="LEAP")
