from __future__ import annotations

from pathlib import Path

import pandas as pd

from codebase.mappings.canonical_mapping import (
    ConfigTableRef,
    build_sector_to_esto_flow_lookup,
    load_canonical_pairs,
    load_fuel_aliases,
    load_sheet_map,
)
from codebase.utilities.leap_results_dashboard_utils import (
    load_explicit_sector_fuel_mappings,
    load_explicit_sector_reassignments,
)


def load_mapping_inputs(
    *,
    sheet_map_path: Path,
    backup_mappings_path: Path,
    codebook_path: Path,
    canonical_pairs_path: ConfigTableRef,
    explicit_mappings_path: Path,
    explicit_reassignments_path: Path,
) -> dict[str, object]:
    sheet_map = load_sheet_map(sheet_map_path)
    fuel_aliases = load_fuel_aliases(backup_mappings_path, codebook_path)
    explicit_mappings = load_explicit_sector_fuel_mappings(explicit_mappings_path)
    explicit_reassignments = load_explicit_sector_reassignments(explicit_reassignments_path)
    sector_flow_mapping = build_sector_to_esto_flow_lookup(codebook_path)
    canonical_pairs, canonical_conflicts = load_canonical_pairs(canonical_pairs_path, strict=False)
    return {
        "sheet_map": sheet_map,
        "fuel_aliases": fuel_aliases,
        "explicit_mappings": explicit_mappings,
        "explicit_reassignments": explicit_reassignments,
        "sector_flow_mapping": sector_flow_mapping,
        "canonical_pairs": canonical_pairs,
        "canonical_conflicts": canonical_conflicts,
    }


def write_canonical_conflicts(conflicts: pd.DataFrame, mapping_views_dir: Path) -> Path | None:
    if conflicts.empty:
        return None
    mapping_views_dir.mkdir(parents=True, exist_ok=True)
    out = mapping_views_dir / "mapping_conflicts.csv"
    conflicts.to_csv(out, index=False)
    return out
