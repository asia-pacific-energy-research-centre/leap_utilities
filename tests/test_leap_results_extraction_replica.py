from pathlib import Path

import pandas as pd

from codebase.utilities.leap_results_extraction_replica import (
    apply_strict_meta,
    compare_tables,
)


def test_apply_strict_meta_for_transformation_sheet_suffixes() -> None:
    template = Path("transformation_results_20_USA_Target.xlsx")
    base = {
        "variable": "Inputs",
        "legend_label": "Branch",
        "scenario": "Target",
        "branch": "Transformation\\Electricity Generation\\Processes",
    }
    result = apply_strict_meta(template, "elecgen_inputs", base)
    assert result["variable"] == "Inputs"
    assert result["legend_label"] == "Fuel"

    result = apply_strict_meta(template, "elecgen_out_feed", base)
    assert result["variable"] == "Outputs by Feedstock Fuel"
    assert result["legend_label"] == "Fuel"

    result = apply_strict_meta(template, "elecgen_out_fuel", base)
    assert result["variable"] == "Outputs by Output Fuel"
    assert result["legend_label"] == "Fuel"


def test_compare_tables_reports_match_and_mismatch() -> None:
    a = pd.DataFrame([["Fuel", 2022], ["Gas", 1.0]])
    b = pd.DataFrame([["Fuel", 2022], ["Gas", 1.0]])
    status, diff, note = compare_tables(a, b)
    assert status == "match"
    assert diff == 0.0
    assert note == ""

    c = pd.DataFrame([["Fuel", 2022], ["Gas", 2.0]])
    status, diff, note = compare_tables(a, c)
    assert status == "value_mismatch"
    assert diff == 1.0
    assert "first mismatch" in note
