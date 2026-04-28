from __future__ import annotations

from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]
OUTPUTS_ROOT = REPO_ROOT / "outputs"

# High-signal top-level output categories.
DASHBOARDS_ROOT = OUTPUTS_ROOT / "dashboards"
MAPPINGS_ROOT = OUTPUTS_ROOT / "mappings"
LEAP_EXPORTS_ROOT = OUTPUTS_ROOT / "leap_exports"
BALANCE_TABLES_ROOT = OUTPUTS_ROOT / "balance_tables"
REFERENCE_ANALYSIS_ROOT = OUTPUTS_ROOT / "reference_analysis"
LEAP_RESULTS_ROOT = OUTPUTS_ROOT / "leap_results"

# Shared LEAP export subcategories.
STANDALONE_LEAP_EXPORTS_ROOT = LEAP_EXPORTS_ROOT / "standalone"
INTEGRATED_LEAP_EXPORTS_ROOT = LEAP_EXPORTS_ROOT / "results_supply_link"
COMBINED_LEAP_EXPORTS_ROOT = LEAP_EXPORTS_ROOT / "combined"


def ensure_output_categories() -> None:
    for path in [
        OUTPUTS_ROOT,
        DASHBOARDS_ROOT,
        MAPPINGS_ROOT,
        LEAP_EXPORTS_ROOT,
        BALANCE_TABLES_ROOT,
        REFERENCE_ANALYSIS_ROOT,
        LEAP_RESULTS_ROOT,
        STANDALONE_LEAP_EXPORTS_ROOT,
        INTEGRATED_LEAP_EXPORTS_ROOT,
        COMBINED_LEAP_EXPORTS_ROOT,
    ]:
        path.mkdir(parents=True, exist_ok=True)
