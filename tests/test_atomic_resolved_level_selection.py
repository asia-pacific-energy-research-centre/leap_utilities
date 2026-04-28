from __future__ import annotations

import pandas as pd

from codebase.utilities.leap_results_dashboard_v2.atomic_engine import resolve_comparison_level


def test_resolve_comparison_level_uses_sector_group_parent_for_grouped_sheets() -> None:
    comparison_long = pd.DataFrame(
        [
            {"sheet": "SubA", "fuel_label": "Coal", "scenario": "Reference", "source": "leap", "year": 2030, "value": 1.0},
            {"sheet": "SubB", "fuel_label": "Coal", "scenario": "Reference", "source": "leap", "year": 2030, "value": 1.0},
            {"sheet": "Standalone", "fuel_label": "Gas", "scenario": "Reference", "source": "leap", "year": 2030, "value": 1.0},
        ]
    )
    sheet_map = pd.DataFrame(
        [
            {"sheet_name": "SubA", "sector_name": "Parent"},
            {"sheet_name": "SubB", "sector_name": "Parent"},
            {"sheet_name": "Standalone", "sector_name": "Standalone"},
        ]
    )
    out = resolve_comparison_level(comparison_long, sheet_map)
    lookup = out.set_index("sheet")["resolved_node_id"].to_dict()
    assert lookup["SubA"] == "Parent"
    assert lookup["SubB"] == "Parent"
    assert lookup["Standalone"] == "Standalone"
