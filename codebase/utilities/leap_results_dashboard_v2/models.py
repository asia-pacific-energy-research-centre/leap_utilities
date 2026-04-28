from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Sequence


@dataclass(frozen=True)
class DashboardV2Settings:
    mapping_graph_mode: str = "common_level_only"
    mapping_precedence: str = "explicit_canonical_fallback"
    ambiguous_policy: str = "aggregate"
    leaf_hole_policy: str = "fail_fast"


@dataclass(frozen=True)
class AtomicSettings:
    enabled: bool = True
    rollout_mode: str = "shadow"
    many_to_many_policy: str = "error"
    write_shadow_outputs: bool = True


@dataclass(frozen=True)
class WorkflowPaths:
    leap_results_dir: Path
    output_dir: Path
    mapping_views_dir: Path


@dataclass(frozen=True)
class WorkflowInputs:
    economy_token: str
    scenarios: Sequence[str]
    base_year: int
    projection_years: Sequence[int]
    base_economy: str
    projection_economy: str
