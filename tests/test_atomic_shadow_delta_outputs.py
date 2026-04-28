from __future__ import annotations

import pandas as pd

from codebase.utilities.leap_results_dashboard_v2.atomic_engine import build_shadow_delta_reports


def test_build_shadow_delta_reports_emits_series_totals_and_summary() -> None:
    legacy = pd.DataFrame(
        [
            {"economy": "20_USA", "sheet": "A", "fuel_label": "Coal", "scenario": "Reference", "source": "projection", "year": 2030, "value": 10.0},
            {"economy": "20_USA", "sheet": "A", "fuel_label": "Total", "scenario": "Reference", "source": "projection", "year": 2030, "value": 10.0},
        ]
    )
    atomic = pd.DataFrame(
        [
            {"economy": "20_USA", "sheet": "A", "fuel_label": "Coal", "scenario": "Reference", "source": "projection", "year": 2030, "value": 12.0},
            {"economy": "20_USA", "sheet": "A", "fuel_label": "Total", "scenario": "Reference", "source": "projection", "year": 2030, "value": 12.0},
        ]
    )
    out = build_shadow_delta_reports(legacy_chart_input=legacy, atomic_chart_input=atomic)
    assert {"atomic_shadow_delta_series", "atomic_shadow_delta_totals", "atomic_shadow_delta_summary"} <= set(out.keys())

    series = out["atomic_shadow_delta_series"]
    coal = series[series["fuel_label"] == "Coal"].iloc[0]
    assert float(coal["delta"]) == 2.0

    totals = out["atomic_shadow_delta_totals"]
    assert len(totals) == 1
    assert totals["fuel_label"].iloc[0] == "Total"

    summary = out["atomic_shadow_delta_summary"]
    assert not summary.empty
    assert "max_abs_delta" in summary.columns
