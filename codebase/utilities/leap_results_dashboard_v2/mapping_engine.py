from __future__ import annotations

import pandas as pd


def annotate_mapping_status(mapping_status: pd.DataFrame) -> pd.DataFrame:
    """Add machine-readable audit flags used by V2 diagnostics and shadow compares."""
    if mapping_status.empty:
        out = mapping_status.copy()
        out["aggregated_mapping"] = False
        return out

    out = mapping_status.copy()
    for col in ["mapping_source", "mapping_note", "sector_match_method"]:
        if col not in out.columns:
            out[col] = ""
        out[col] = out[col].fillna("").astype(str).str.strip().str.lower()

    out["aggregated_mapping"] = (
        out["mapping_source"].isin({"canonical_aggregated", "category_sector"})
        | out["sector_match_method"].eq("aggregated_canonical_targets")
        | out["mapping_note"].str.contains("aggregated", regex=False)
    )
    out["mapping_precedence"] = "explicit_canonical_fallback"
    out["ambiguous_policy"] = "aggregate"
    return out
