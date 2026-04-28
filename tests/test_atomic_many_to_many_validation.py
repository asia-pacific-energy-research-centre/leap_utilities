from __future__ import annotations

import pandas as pd

from codebase.utilities.leap_results_dashboard_v2.atomic_engine import (
    find_unresolved_many_to_many_components,
)


def test_find_unresolved_many_to_many_components_flags_unresolved_component() -> None:
    edges = pd.DataFrame(
        [
            {
                "sheet": "A",
                "scenario": "Reference",
                "source_family": "base",
                "resolved_node_id": "A",
                "line_key": "L1",
                "atomic_key": "X1",
                "edge_reason": "unresolved_mapping",
            },
            {
                "sheet": "A",
                "scenario": "Reference",
                "source_family": "base",
                "resolved_node_id": "A",
                "line_key": "L1",
                "atomic_key": "X2",
                "edge_reason": "unresolved_mapping",
            },
            {
                "sheet": "A",
                "scenario": "Reference",
                "source_family": "base",
                "resolved_node_id": "A",
                "line_key": "L2",
                "atomic_key": "X1",
                "edge_reason": "unresolved_mapping",
            },
        ]
    )
    out = find_unresolved_many_to_many_components(edges)
    assert len(out) == 1
    row = out.iloc[0]
    assert int(row["line_count"]) == 2
    assert int(row["atomic_count"]) == 2


def test_find_unresolved_many_to_many_components_ignores_deterministic_component() -> None:
    edges = pd.DataFrame(
        [
            {
                "sheet": "A",
                "scenario": "Reference",
                "source_family": "projection",
                "resolved_node_id": "A",
                "line_key": "L1",
                "atomic_key": "X1",
                "edge_reason": "base_share_allocation",
            },
            {
                "sheet": "A",
                "scenario": "Reference",
                "source_family": "projection",
                "resolved_node_id": "A",
                "line_key": "L2",
                "atomic_key": "X1",
                "edge_reason": "base_share_allocation",
            },
            {
                "sheet": "A",
                "scenario": "Reference",
                "source_family": "projection",
                "resolved_node_id": "A",
                "line_key": "L2",
                "atomic_key": "X2",
                "edge_reason": "equal_split_allocation",
            },
        ]
    )
    out = find_unresolved_many_to_many_components(edges)
    assert out.empty
