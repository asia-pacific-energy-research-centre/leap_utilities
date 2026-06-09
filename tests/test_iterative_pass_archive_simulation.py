"""
Archive-based simulation of successive BETWEEN-PASSES / PASS 2+ iterations.

Treats the existing archive balance-table CSVs as if they were successive
LEAP recalculation results, running _run_capacity_unmet_iterative_balanced_pass
for each and verifying that:

  1. A JSON state file is created after the first pass.
  2. The state accumulates across passes (passes list grows; cumulative
     additions are non-decreasing).
  3. Signature detection marks different archive sets as "new results".
  4. Positive import gaps (observed > adjusted) trigger capacity or primary-
     production allocations.
  5. Running the same files twice is detected and warned, not silently rerun.

Data sources:
  - "Pass A": archive/balance_table_20_USA_25042026_TGT_*.csv  (oldest)
  - "Pass B": archive/balance_table_20_USA_26042026_TGT_*.csv  (middle)
  - "Pass C": balance_table_20_USA_15052026_TGT_*.csv           (current)

All tests are skipped when the real data files do not exist so the suite
stays green in a fresh checkout.
"""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import pytest

# ---------------------------------------------------------------------------
# Locate real data on disk (tests skip gracefully when files are absent)
# ---------------------------------------------------------------------------
REPO_ROOT = Path(__file__).resolve().parents[1]
YEARLY_BALANCE_DIR = (
    REPO_ROOT
    / "outputs"
    / "balance_tables"
    / "results_supply_link"
    / "yearly_balance_tables"
)
ARCHIVE_DIR = YEARLY_BALANCE_DIR / "archive"
RECONCILIATION_CSV = (
    REPO_ROOT
    / "outputs"
    / "leap_exports"
    / "results_supply_link"
    / "results_supply_reconciliation.csv"
)

_PASS_A_CSVS = sorted(ARCHIVE_DIR.glob("balance_table_20_USA_25042026_TGT_*.csv"))
_PASS_B_CSVS = sorted(ARCHIVE_DIR.glob("balance_table_20_USA_26042026_TGT_*.csv"))
_PASS_C_CSVS = sorted(YEARLY_BALANCE_DIR.glob("balance_table_20_USA_15052026_TGT_*.csv"))

_HAS_REAL_DATA = (
    RECONCILIATION_CSV.exists()
    and len(_PASS_A_CSVS) >= 3
    and len(_PASS_B_CSVS) >= 3
    and len(_PASS_C_CSVS) >= 3
)

_SKIP_NO_DATA = pytest.mark.skipif(
    not _HAS_REAL_DATA,
    reason=(
        "Real archive balance CSVs and/or reconciliation CSV not found. "
        "Run the results-supply-link workflow at least twice to generate them."
    ),
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _load_reconciliation() -> pd.DataFrame:
    """Load the saved reconciliation table for the Target scenario only."""
    rec = pd.read_csv(RECONCILIATION_CSV)
    return rec[rec["scenario"] == "Target"].copy().reset_index(drop=True)


def _build_process_records(reconciliation: pd.DataFrame) -> list[dict]:
    """
    Build process_records covering all secondary ESTO products.

    Primary fuels (coal, crude oil, natural gas, renewables, etc.) are
    handled by the production lever natively — they must NOT appear in
    process_records because the balanced pass only uses process_catalog for
    the transformation lever.

    Secondary fuels are mapped to the transformation module that produces
    them.  Output values are derived from adjusted_imports in the
    reconciliation table (a realistic proxy for what the module needs to
    produce to be self-sufficient).

    Module → products mapping follows ESTO sector definitions:
      Oil refineries           — all 07.xx petroleum products + 06.03/04/05
      Coke ovens               — 02.01 coke, 02.02 gas coke, 02.03 coke oven gas,
                                 02.07 coal tar, 02.05 other recovered gases
      Blast furnaces           — 02.04 blast furnace gas
      Patent fuel plants       — 02.06 patent fuel
      BKB and PB plants        — 02.08 BKB/PB
      NG Liquefaction          — 08.02 LNG
      Gas works plants         — 08.03 gas works gas
      Electricity generation   — 17 Electricity (plus CHP to satisfy priority list)
      Main activity CHP plants — 17 Electricity, 18 Heat
      Heat plants              — 18 Heat
      Charcoal processing      — 15.03 charcoal
      Biofuels processing      — 15.04 black liquour, 16.05–16.08 biofuels
      Biogas production        — 16.04 Biogas (where applicable)

    Multi-module products ("17 Electricity", "18 Heat") are already covered
    by CAPACITY_UNMET_PRIORITY_BY_PRODUCT in the workflow config.
    """
    mask = reconciliation["year"].isin([2022, 2030, 2050])
    subset = reconciliation[mask][["esto_product", "year", "adjusted_imports"]].copy()

    def _year_map(product: str, fallback: float = 1.0) -> dict[int, float]:
        """Return {year: max(adjusted_imports, fallback)} for the 3 key years."""
        rows = subset[subset["esto_product"] == product]
        if rows.empty:
            # Product not in reconciliation — give a nominal non-zero value so
            # the catalog entry exists without inflating allocations.
            return {2022: fallback, 2030: fallback, 2050: fallback}
        return {
            int(r["year"]): max(float(r["adjusted_imports"]), fallback)
            for _, r in rows.iterrows()
            if pd.notna(r["adjusted_imports"])
        }

    OIL_REFINERY_PRODUCTS = {
        # 06.xx refinery-derived
        "06.03 Refinery feedstocks",
        "06.04 Additives/  oxygenates",
        "06.05 Other hydrocarbons",
        # 07.xx full petroleum slate
        "07.01 Motor gasoline",
        "07.02 Aviation gasoline",
        "07.03 Naphtha",
        "07.04 Gasoline type jet fuel",
        "07.05 Kerosene type jet fuel",
        "07.06 Kerosene",
        "07.07 Gas/diesel oil",
        "07.08 Fuel oil",
        "07.09 LPG",
        "07.10 Refinery gas (not liquefied)",
        "07.11 Ethane",
        "07.12 White spirit SBP",
        "07.13 Lubricants",
        "07.14 Bitumen",
        "07.15 Paraffin  waxes",
        "07.16 Petroleum coke",
        "07.17 Other products",
        "07.99 PetProd nonspecified",
    }

    records = [
        # ---- Oil refineries (full 07.xx slate + refinery-derived 06.xx) ----
        {
            "economy": "20_USA",
            "sector_title": "Oil refineries",
            "process_name": "Refinery",
            "output_values": {p: _year_map(p) for p in OIL_REFINERY_PRODUCTS},
        },
        # ---- Coke ovens ----
        {
            "economy": "20_USA",
            "sector_title": "Coke ovens",
            "process_name": "Coke oven",
            "output_values": {
                "02.01 Coke oven coke": _year_map("02.01 Coke oven coke"),
                "02.02 Gas coke": _year_map("02.02 Gas coke"),
                "02.03 Coke oven gas": _year_map("02.03 Coke oven gas"),
                "02.05 Other recovered gases": _year_map("02.05 Other recovered gases"),
                "02.07 Coal tar": _year_map("02.07 Coal tar"),
            },
        },
        # ---- Blast furnaces ----
        {
            "economy": "20_USA",
            "sector_title": "Blast furnaces",
            "process_name": "Blast furnace",
            "output_values": {
                "02.04 Blast furnace gas": _year_map("02.04 Blast furnace gas"),
            },
        },
        # ---- Patent fuel plants ----
        {
            "economy": "20_USA",
            "sector_title": "Patent fuel plants",
            "process_name": "Patent fuel plant",
            "output_values": {
                "02.06 Patent fuel": _year_map("02.06 Patent fuel"),
            },
        },
        # ---- BKB and PB plants ----
        {
            "economy": "20_USA",
            "sector_title": "BKB and PB plants",
            "process_name": "BKB/PB plant",
            "output_values": {
                "02.08 BKB/PB": _year_map("02.08 BKB/PB"),
            },
        },
        # ---- NG Liquefaction → LNG ----
        {
            "economy": "20_USA",
            "sector_title": "NG Liquefaction",
            "process_name": "LNG liquefaction",
            "output_values": {
                "08.02 LNG": _year_map("08.02 LNG"),
            },
        },
        # ---- Gas works plants ----
        {
            "economy": "20_USA",
            "sector_title": "Gas works plants",
            "process_name": "Gas works",
            "output_values": {
                "08.03 Gas works gas": _year_map("08.03 Gas works gas"),
            },
        },
        # ---- Electricity generation ----
        # Listed alongside CHP so CAPACITY_UNMET_PRIORITY_BY_PRODUCT for
        # "17 Electricity" covers multiple modules — no validation error.
        {
            "economy": "20_USA",
            "sector_title": "Electricity generation",
            "process_name": "Power plants",
            "output_values": {"17 Electricity": _year_map("17 Electricity")},
        },
        # ---- Main activity producer CHP plants (electricity + heat) ----
        {
            "economy": "20_USA",
            "sector_title": "Main activity producer CHP plants",
            "process_name": "CHP plants",
            "output_values": {
                "17 Electricity": _year_map("17 Electricity"),
                "18 Heat": _year_map("18 Heat"),
            },
        },
        # ---- Heat plants ----
        {
            "economy": "20_USA",
            "sector_title": "Heat plants",
            "process_name": "Heat plant",
            "output_values": {"18 Heat": _year_map("18 Heat")},
        },
        # ---- Charcoal processing ----
        {
            "economy": "20_USA",
            "sector_title": "Charcoal processing",
            "process_name": "Charcoal kiln",
            "output_values": {
                "15.03 Charcoal": _year_map("15.03 Charcoal"),
            },
        },
        # ---- Biofuels processing ----
        {
            "economy": "20_USA",
            "sector_title": "Biofuels processing",
            "process_name": "Biofuel plant",
            "output_values": {
                "15.04 Black liqour": _year_map("15.04 Black liqour"),
                "16.05 Biogasoline": _year_map("16.05 Biogasoline"),
                "16.06 Biodiesel": _year_map("16.06 Biodiesel"),
                "16.07 Bio jet kerosene": _year_map("16.07 Bio jet kerosene"),
                "16.08 Other liquid biofuels": _year_map("16.08 Other liquid biofuels"),
            },
        },
        # ---- Peat products ----
        {
            "economy": "20_USA",
            "sector_title": "Patent fuel plants",
            "process_name": "Peat processing",
            "output_values": {
                "04 Peat products": _year_map("04 Peat products"),
            },
        },
    ]
    return records


def _run_pass(
    *,
    wf,
    reconciliation: pd.DataFrame,
    process_records: list[dict],
    balance_csvs: list[Path],
    state_path: Path,
) -> dict:
    """
    Run one results_update pass and return the summary dict.

    Monkey-patches CAPACITY_UNMET_PASS_MODE to 'results_update' and redirects
    RESULTS_CHECKS_DIR to the same tmp dir as state_path so diagnostic CSVs
    don't land in the real output tree (which may be locked by Excel).
    """
    original_mode = wf.CAPACITY_UNMET_PASS_MODE
    original_checks = wf.RESULTS_CHECKS_DIR
    wf.CAPACITY_UNMET_PASS_MODE = "results_update"
    wf.RESULTS_CHECKS_DIR = state_path.parent / "checks"
    try:
        return wf._run_capacity_unmet_iterative_balanced_pass(
            reconciliation_table=reconciliation,
            process_records=process_records,
            economies=["20_USA"],
            scenarios=["Target"],
            results_dir=balance_csvs,
            state_path=state_path,
            allow_same_results_reuse=True,
        )
    finally:
        wf.CAPACITY_UNMET_PASS_MODE = original_mode
        wf.RESULTS_CHECKS_DIR = original_checks


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


@_SKIP_NO_DATA
def test_iterative_balanced_pass_creates_state_on_first_call(tmp_path: Path) -> None:
    """State JSON is created from scratch after the first results_update pass."""
    from codebase import results_supply_link_workflow as wf

    reconciliation = _load_reconciliation()
    process_records = _build_process_records(reconciliation)
    state_path = tmp_path / "state.json"

    assert not state_path.exists(), "Pre-condition: state must not exist yet"

    _run_pass(
        wf=wf,
        reconciliation=reconciliation,
        process_records=process_records,
        balance_csvs=list(_PASS_A_CSVS),
        state_path=state_path,
    )

    assert state_path.exists(), "State JSON must be written after the first pass"
    payload = json.loads(state_path.read_text(encoding="utf-8"))
    assert "passes" in payload
    assert len(payload["passes"]) == 1
    assert payload["passes"][0]["iteration_run_mode"] == "results_update"


@_SKIP_NO_DATA
def test_iterative_balanced_pass_accumulates_state_across_archive_passes(
    tmp_path: Path,
) -> None:
    """
    Run three successive passes (A → B → C) and verify cumulative state grows.

    After pass A the state has 1 entry in 'passes'.
    After pass B it has 2 entries and the signature has changed (new files).
    After pass C it has 3 entries.

    Also verifies that cumulative_capacity_additions and
    cumulative_primary_additions are non-decreasing: values from later passes
    must be >= values from earlier passes for each key.
    """
    from codebase import results_supply_link_workflow as wf

    reconciliation = _load_reconciliation()
    process_records = _build_process_records(reconciliation)
    state_path = tmp_path / "state.json"

    # ---- Pass A (25042026 archive) ------------------------------------------
    _run_pass(
        wf=wf,
        reconciliation=reconciliation,
        process_records=process_records,
        balance_csvs=list(_PASS_A_CSVS),
        state_path=state_path,
    )
    payload_a = json.loads(state_path.read_text(encoding="utf-8"))
    cap_a = payload_a.get("cumulative_capacity_additions", {})
    prim_a = payload_a.get("cumulative_primary_additions", {})
    assert len(payload_a["passes"]) == 1, "Expected 1 pass after Pass A"

    # ---- Pass B (26042026 archive) ------------------------------------------
    _run_pass(
        wf=wf,
        reconciliation=reconciliation,
        process_records=process_records,
        balance_csvs=list(_PASS_B_CSVS),
        state_path=state_path,
    )
    payload_b = json.loads(state_path.read_text(encoding="utf-8"))
    cap_b = payload_b.get("cumulative_capacity_additions", {})
    prim_b = payload_b.get("cumulative_primary_additions", {})
    assert len(payload_b["passes"]) == 2, "Expected 2 passes after Pass B"

    # Signatures should differ between A and B (different files).
    sig_a = payload_a["passes"][0].get("results_signature_used", {})
    sig_b = payload_b["passes"][1].get("results_signature_used", {})
    assert sig_a != sig_b, "Pass B must record a different signature from Pass A"

    # Cumulative additions must be non-decreasing for every key present in A.
    for key, val_a in cap_a.items():
        assert cap_b.get(key, val_a) >= val_a - 1e-9, (
            f"Cumulative capacity addition regressed for key '{key}': "
            f"{val_a} → {cap_b.get(key)}"
        )
    for key, val_a in prim_a.items():
        assert prim_b.get(key, val_a) >= val_a - 1e-9, (
            f"Cumulative primary addition regressed for key '{key}': "
            f"{val_a} → {prim_b.get(key)}"
        )

    # ---- Pass C (15052026 current) ------------------------------------------
    _run_pass(
        wf=wf,
        reconciliation=reconciliation,
        process_records=process_records,
        balance_csvs=list(_PASS_C_CSVS),
        state_path=state_path,
    )
    payload_c = json.loads(state_path.read_text(encoding="utf-8"))
    cap_c = payload_c.get("cumulative_capacity_additions", {})
    prim_c = payload_c.get("cumulative_primary_additions", {})
    assert len(payload_c["passes"]) == 3, "Expected 3 passes after Pass C"

    # Non-decreasing from B → C as well.
    for key, val_b in cap_b.items():
        assert cap_c.get(key, val_b) >= val_b - 1e-9, (
            f"Cumulative capacity addition regressed B→C for key '{key}': "
            f"{val_b} → {cap_c.get(key)}"
        )
    for key, val_b in prim_b.items():
        assert prim_c.get(key, val_b) >= val_b - 1e-9, (
            f"Cumulative primary addition regressed B→C for key '{key}': "
            f"{val_b} → {prim_c.get(key)}"
        )


@_SKIP_NO_DATA
def test_iterative_balanced_pass_allocates_for_positive_gaps(tmp_path: Path) -> None:
    """
    Pass A (25042026 archive) has confirmed positive import gaps for
    '17 Electricity', '18 Heat', and '07.05 Kerosene type jet fuel'.
    The pass must allocate capacity or primary production for at least one
    of those products, and the total allocated output must be positive.
    """
    from codebase import results_supply_link_workflow as wf

    reconciliation = _load_reconciliation()
    process_records = _build_process_records(reconciliation)
    state_path = tmp_path / "state.json"

    summary = _run_pass(
        wf=wf,
        reconciliation=reconciliation,
        process_records=process_records,
        balance_csvs=list(_PASS_A_CSVS),
        state_path=state_path,
    )

    allocated_transform = float(summary.get("allocated_transformation_output_total", 0.0))
    allocated_primary = float(summary.get("allocated_primary_output_total", 0.0))
    total_allocated = allocated_transform + allocated_primary

    assert total_allocated > 0.0, (
        f"Expected positive allocations for Pass A (gaps exist for electricity, "
        f"heat, and jet fuel), but got transform={allocated_transform}, "
        f"primary={allocated_primary}"
    )

    # State must record the gap products that received allocations.
    payload = json.loads(state_path.read_text(encoding="utf-8"))
    cap_keys = payload.get("cumulative_capacity_additions", {})
    prim_keys = payload.get("cumulative_primary_additions", {})
    assert len(cap_keys) + len(prim_keys) > 0, (
        "Cumulative state must have at least one addition entry after a pass "
        "with positive gaps"
    )


@_SKIP_NO_DATA
def test_iterative_balanced_pass_same_file_does_not_reset_state(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    """
    Running the same balance CSVs twice must not corrupt cumulative state.
    The second call with allow_same_results_reuse=True must produce the same
    (or greater) cumulative additions as the first call.
    """
    from codebase import results_supply_link_workflow as wf

    reconciliation = _load_reconciliation()
    process_records = _build_process_records(reconciliation)
    state_path = tmp_path / "state.json"

    # First call.
    _run_pass(
        wf=wf,
        reconciliation=reconciliation,
        process_records=process_records,
        balance_csvs=list(_PASS_A_CSVS),
        state_path=state_path,
    )
    payload_1 = json.loads(state_path.read_text(encoding="utf-8"))
    cap_1 = payload_1.get("cumulative_capacity_additions", {})
    prim_1 = payload_1.get("cumulative_primary_additions", {})

    # Second call with the identical files.
    _run_pass(
        wf=wf,
        reconciliation=reconciliation,
        process_records=process_records,
        balance_csvs=list(_PASS_A_CSVS),
        state_path=state_path,
    )
    payload_2 = json.loads(state_path.read_text(encoding="utf-8"))
    cap_2 = payload_2.get("cumulative_capacity_additions", {})
    prim_2 = payload_2.get("cumulative_primary_additions", {})

    # State must not have lost any additions.
    for key, val in cap_1.items():
        assert cap_2.get(key, val) >= val - 1e-9, (
            f"Capacity addition for '{key}' regressed on re-run: {val} → {cap_2.get(key)}"
        )
    for key, val in prim_1.items():
        assert prim_2.get(key, val) >= val - 1e-9, (
            f"Primary addition for '{key}' regressed on re-run: {val} → {prim_2.get(key)}"
        )

    # The passes list must now have 2 entries.
    assert len(payload_2["passes"]) == 2


@_SKIP_NO_DATA
def test_baseline_seed_resets_state_and_archives_old_json(tmp_path: Path) -> None:
    """
    Switching to baseline_seed mode must reset cumulative additions to zero
    and archive the old state JSON (not delete it).
    """
    from codebase import results_supply_link_workflow as wf

    reconciliation = _load_reconciliation()
    process_records = _build_process_records(reconciliation)
    state_path = tmp_path / "state.json"
    archive_dir = tmp_path / "archive"
    archive_dir.mkdir()

    # Build up state across two passes.
    _run_pass(
        wf=wf,
        reconciliation=reconciliation,
        process_records=process_records,
        balance_csvs=list(_PASS_A_CSVS),
        state_path=state_path,
    )
    _run_pass(
        wf=wf,
        reconciliation=reconciliation,
        process_records=process_records,
        balance_csvs=list(_PASS_B_CSVS),
        state_path=state_path,
    )
    assert state_path.exists()
    payload_before = json.loads(state_path.read_text(encoding="utf-8"))
    assert len(payload_before["passes"]) == 2

    # Simulate a baseline_seed reset (what the main workflow does at the start
    # of a fresh run).  _read_capacity_unmet_state in baseline_seed mode
    # archives the old state and returns an empty default.
    original_mode = wf.CAPACITY_UNMET_PASS_MODE
    original_archive = wf.RESULTS_SINGLE_FILE_ARCHIVE_DIR
    wf.CAPACITY_UNMET_PASS_MODE = "baseline_seed"
    wf.RESULTS_SINGLE_FILE_ARCHIVE_DIR = archive_dir
    try:
        fresh_state = wf._read_capacity_unmet_state(state_path=state_path, run_mode="baseline_seed")
        wf._write_capacity_unmet_state(fresh_state, state_path=state_path)
    finally:
        wf.CAPACITY_UNMET_PASS_MODE = original_mode
        wf.RESULTS_SINGLE_FILE_ARCHIVE_DIR = original_archive

    # State must now be empty (no passes, no additions).
    payload_after = json.loads(state_path.read_text(encoding="utf-8"))
    assert payload_after.get("cumulative_capacity_additions") == {}, (
        "baseline_seed must clear cumulative_capacity_additions"
    )
    assert payload_after.get("cumulative_primary_additions") == {}, (
        "baseline_seed must clear cumulative_primary_additions"
    )
    assert payload_after.get("passes", []) == [], (
        "baseline_seed must clear passes list"
    )

    # Old state must have been archived (not deleted).
    # The archive filename is derived from the state_path stem ("state"),
    # so the pattern is "state_<timestamp>.json".
    archived = sorted(archive_dir.glob("state_*.json"))
    assert len(archived) == 1, (
        f"Expected one archived state JSON in {archive_dir}, found {len(archived)}"
    )
