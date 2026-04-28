import json
import re
from pathlib import Path

import pandas as pd

from codebase.utilities.leap_results_dashboard_utils import make_chart


def test_make_chart_uses_compact_legend_labels(tmp_path: Path):
    rows: list[dict[str, object]] = []
    for scenario, offset in [("Reference", 0.0), ("Target", 100.0)]:
        rows.extend(
            [
                {"year": 2022, "value": 10.0 + offset, "source": "leap"},
                {"year": 2023, "value": 11.0 + offset, "source": "leap"},
                {"year": 2022, "value": 20.0 + offset, "source": "projection"},
                {"year": 2023, "value": 21.0 + offset, "source": "projection"},
                {"year": 2022, "value": 30.0 + offset, "source": "projection_mixed"},
                {"year": 2023, "value": 31.0 + offset, "source": "projection_mixed"},
                {"year": 2022, "value": 40.0 + offset, "source": "projection_estimated"},
                {"year": 2023, "value": 41.0 + offset, "source": "projection_estimated"},
                {"year": 2022, "value": 50.0 + offset, "source": "base"},
                {"year": 2022, "value": 60.0 + offset, "source": "base_mixed"},
                {"year": 2022, "value": 70.0 + offset, "source": "base_estimated"},
            ]
        )
        for row in rows[-11:]:
            row.update({"economy": "20_USA", "sheet": "Industry", "measure": "Energy (PJ)", "fuel_label": "Coal", "scenario": scenario})

    subset = pd.DataFrame(rows)
    chart_path = make_chart("Industry", "Coal", subset, tmp_path, backend="plotly")

    assert chart_path is not None
    html = chart_path.read_text(encoding="utf-8")
    names = {json.loads(f'"{name}"') for name in re.findall(r'"name":"([^"]+)"', html)}

    for expected in [
        "LEAP REF",
        "LEAP TGT",
        "9th projection REF",
        "9th projection TGT",
        "9th projection est/real REF",
        "9th projection est/real TGT",
        "9th projection est REF",
        "9th projection est TGT",
        "Base 2022 REF",
        "Base 2022 TGT",
        "Base est/real REF",
        "Base est/real TGT",
        "Base est REF",
        "Base est TGT",
    ]:
        assert expected in names

    for unexpected in [
        "LEAP (Reference)",
        "LEAP (Target)",
        "9th projection (Reference)",
        "9th projection (Target)",
        "9th projection (estimated + real, Reference)",
        "9th projection (estimated + real, Target)",
        "9th projection (estimated, Reference)",
        "9th projection (estimated, Target)",
        "Base (2022, Reference)",
        "Base (2022, Target)",
        "Base (estimated + real, Reference)",
        "Base (estimated + real, Target)",
        "Base (estimated, Reference)",
        "Base (estimated, Target)",
    ]:
        assert unexpected not in names
