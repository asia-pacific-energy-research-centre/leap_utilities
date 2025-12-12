"""Utilities for reconciling model outputs with ESTO energy use totals.

This version avoids classes to keep the workflow simple and approachable. The
functions mirror the validation approach used in
``validate_final_energy_use_for_base_year_equals_esto_totals`` while remaining
sector-agnostic so they can be reused for transport, buildings, supply, or
other sectors.

Core ideas
----------
- Energy use for each LEAP branch is calculated by multiplying together the
  relevant input variables (e.g., activity * intensity or stock * mileage *
  efficiency). The variables can be overridden per-branch if needed.
- Totals are compared with ESTO values for the same keys, and any mismatch is
  resolved by scaling the base-year inputs proportionally.
- Everything is driven by plain functions and dictionaries, keeping the module
  flexible without requiring users to understand custom classes.

Example (transport)
-------------------
>>> from transport_branch_mappings import ESTO_SECTOR_FUEL_TO_LEAP_BRANCH_MAP
>>> from transport_measure_catalog import get_leap_branch_to_analysis_type_mapping
>>> export_df = pd.read_excel("../results/USA_transport_leap_export_Target.xlsx")
>>> esto_totals = {('15_02_road', '07_petroleum_products', '07_01_motor_gasoline'): 100.0}
>>> branch_rules = build_branch_rules_from_mapping(
...     ESTO_SECTOR_FUEL_TO_LEAP_BRANCH_MAP,
...     get_leap_branch_to_analysis_type_mapping,
...     root="Demand",
... )
>>> adjusted_df, summary = reconcile_energy_use(
...     export_df=export_df,
...     base_year=2022,
...     branch_mapping_rules=branch_rules,
...     esto_energy_totals=esto_totals,
... )
"""

from __future__ import annotations

from typing import Callable, Dict, List, Mapping, Optional, Sequence, Tuple

import pandas as pd

# Optional transport-specific helper; if missing, validation will be skipped
try:  # pragma: no cover - optional dependency
    from transport_branch_mappings import identify_missing_esto_mappings_for_leap_branches
except Exception:  # pragma: no cover - optional dependency
    identify_missing_esto_mappings_for_leap_branches = None

# ---------------------------------------------------------------------------
# Helpers for branch paths and defaults
# ---------------------------------------------------------------------------

def build_branch_path(branch_tuple: Tuple[str, ...], root: str = "Demand") -> str:
    """Convert a tuple of branch labels into a LEAP-style branch path string."""

    segments = [root, *branch_tuple]
    cleaned_segments = [segment for segment in segments if segment]
    return "\\".join(cleaned_segments)


DEFAULT_STRATEGIES: Dict[str, Sequence[str]] = {
    "Intensity": ["Activity Level", "Final Energy Intensity"],
    "Stock": ["Stock", "Mileage", "Fuel Economy"],
}


def get_leap_branch_to_analysis_type_mapping(leap_branch, leap_branch_to_analysis_type_mapping):
    """Return the analysis type associated with a LEAP branch path. This is the default function to use in build_branch_rules_from_mapping but transport uses a custom one."""
    analysis_type = leap_branch_to_analysis_type_mapping.get(leap_branch)
    if analysis_type is not None:
        return analysis_type
    else:
        breakpoint()
        raise ValueError(f"No analysis type mapping found for LEAP branch: {leap_branch}")


def build_branch_rules_from_mapping(
    esto_to_leap_mapping: Mapping[Tuple[str, ...], Sequence[Tuple[str, ...]]],
    unmappable_branches: Sequence[Tuple[str, ...]],
    all_leap_branches: Sequence[Tuple[str, ...]],
    analysis_type_lookup: Callable[[Tuple[str, ...]], str],
    root: str = "Demand",
) -> Dict[Tuple[str, ...], List[Dict[str, object]]]:
    """Construct branch rules from an ESTO → LEAP mapping.

    Each rule is a plain dictionary so users can inspect or edit them without
    needing to understand custom classes.

    An example ESTO to LEAP mapping is the one in `transport_branch_mappings.py`
    > ESTO_SECTOR_FUEL_TO_LEAP_BRANCH_MAP:

    ("15_01_domestic_air_transport", "07_petroleum_products", "07_01_motor_gasoline"): [
        ("Nonspecified transport", "Gasoline")
    ],
    ("15_01_domestic_air_transport", "07_petroleum_products", "07_02_aviation_gasoline"): [
        ("Passenger non road", "Air", "Aviation gasoline"),
        ("Freight non road", "Air", "Aviation gasoline")
    ],

    Note that the ESTO to LEAP mapping should come with associated lists and
    mappings to help determine whether all branches are mapped (i.e.
    identify_missing_esto_mappings_for_leap_branches() needs a list of all LEAP
    branches and those that cant be mapped (such as new fuel types)).
    """
    if identify_missing_esto_mappings_for_leap_branches is not None:
        identify_missing_esto_mappings_for_leap_branches(
            esto_to_leap_mapping, unmappable_branches, all_leap_branches
        )

    rules: Dict[Tuple[str, ...], List[Dict[str, object]]] = {}
    for esto_key, leap_branches in esto_to_leap_mapping.items():
        rules[esto_key] = [
            {
                "branch_tuple": branch,
                "calculation_strategy": analysis_type_lookup(branch),
                "root": root,
                "input_variables_override": None,
            }
            for branch in leap_branches
        ]
    return rules


# ---------------------------------------------------------------------------
# Energy calculations
# ---------------------------------------------------------------------------

def _default_input_series_provider(
    export_df: pd.DataFrame,
    base_year: int | str,
    branch_path: str,
    input_variables: Sequence[str],
) -> List[pd.Series]:
    """Return the series needed for energy calculation using simple Branch/Variable masking."""

    if base_year not in export_df.columns:
        breakpoint()
        raise KeyError(f"Base year column '{base_year}' not found in export data.")

    series_list: List[pd.Series] = []
    for variable in input_variables:
        mask = (export_df["Branch Path"] == branch_path) & (export_df["Variable"] == variable)
        series = export_df.loc[mask, base_year]
        if series.empty:
            breakpoint()
            raise ValueError(
                f"No values found for variable '{variable}' on branch '{branch_path}' in {base_year}."
            )
        series_list.append(series)
    return series_list


def calculate_branch_energy(
    export_df: pd.DataFrame,
    base_year: int | str,
    rule: Mapping[str, object],
    strategies: Mapping[str, Sequence[str]],
    combination_fn: Optional[Callable[[List[pd.Series]], pd.Series]] = None,
    energy_fn: Optional[
        Callable[[pd.DataFrame, int | str, Mapping[str, object], Mapping[str, Sequence[str]], Optional[Callable]], float]
    ] = None,
    input_series_provider: Optional[
        Callable[[pd.DataFrame, int | str, str, Sequence[str]], List[pd.Series]]
    ] = None,
) -> float:
    """
    Calculate energy for a single branch rule.

    Users can override behaviour in two ways:
    - Provide ``energy_fn`` to fully replace the calculation (e.g., complex power generation logic).
    - Provide ``input_series_provider`` to customize how input variables are pulled, while leaving the
      final multiplication/combination logic intact.
    """

    if energy_fn:
        return float(energy_fn(export_df, base_year, rule, strategies, combination_fn))

    input_vars = rule.get("input_variables_override") or strategies[rule["calculation_strategy"]]
    branch_path = build_branch_path(rule["branch_tuple"], root=rule.get("root", "Demand"))
    provider = input_series_provider or _default_input_series_provider
    series_list = provider(export_df, base_year, branch_path, input_vars)

    if not series_list:
        return 0.0

    if combination_fn:
        combined = combination_fn(series_list)
    else:
        combined = series_list[0]
        for additional in series_list[1:]:
            combined = combined * additional

    return float(combined.sum())


# ---------------------------------------------------------------------------
# Reconciliation logic
# ---------------------------------------------------------------------------

def _compute_scale_factor(leap_total: float, esto_total: float) -> float:
    if leap_total == 0:
        return 1.0
    return esto_total / leap_total


def get_adjustment_year_columns(
    export_df: pd.DataFrame,
    base_year: int | str,
    include_future_years: bool = False,
    apply_adjustments_to_past_years: bool = False,
) -> List[int | str]:
    """Return a list of year columns to adjust, including the base year."""
    if not include_future_years:
        return [base_year]
    
    years: List[int | str] = []
    year_columns = [
        col
        for col in export_df.columns
        if len(str(col)) == 4 and str(col).isdigit()
    ]
    #in case we are considering years before and after the base year we want to sort them properly.. not that this causes a potential issue where any afdjustments to historical data may not result in the same energy use as is expected.
    for col in sorted(year_columns, key=lambda c: int(str(c))):
        if not apply_adjustments_to_past_years and int(col) < int(base_year):
            continue
        years.append(col)

    return years


def _apply_proportional_adjustment(
    export_df: pd.DataFrame,
    base_year: int | str,
    rule: Mapping[str, object],
    scale_factor: float,
    strategies: Mapping[str, Sequence[str]],
    year_columns: Optional[Sequence[int | str]] = None,
) -> None:
    """Scale the base-year inputs for a branch by the provided factor.

    Generic, sector-agnostic fallback:
      energy ∝ ∏(input variables)
      ⇒ to scale energy by `scale_factor`, multiply each input by `scale_factor`.
      
    NOTE THIS PROBABLY WON'T WORK FOR ANY SECTOR SINCE THEY ARE ALL SOMEWHAT COMPLEX AND NEED THEIR OWN ADJUSTMENT LOGIC
    """
    years_to_adjust = list(year_columns or [base_year])
    input_vars = rule.get("input_variables_override") or strategies[rule["calculation_strategy"]]
    branch_path = build_branch_path(rule["branch_tuple"], root=rule.get("root", "Demand"))
    for variable in input_vars:
        mask = (export_df["Branch Path"] == branch_path) & (export_df["Variable"] == variable)
        if not mask.any():
            continue
        for year_col in years_to_adjust:
            if year_col not in export_df.columns:
                continue
            export_df.loc[mask, year_col] = export_df.loc[mask, year_col] * scale_factor
        
def reconcile_energy_use(
    export_df: pd.DataFrame,
    base_year: int | str,
    branch_mapping_rules: Mapping[Tuple[str, ...], Sequence[Mapping[str, object]]],
    esto_energy_totals: Mapping[Tuple[str, ...], float],
    strategies: Optional[Mapping[str, Sequence[str]]] = None,
    tolerance: float = 1e-6,
    adjustment_fn: Optional[
        Callable[
            [pd.DataFrame, int | str, Mapping[str, object], float, Mapping[str, Sequence[str]], Optional[Sequence[int | str]]],
            None,
        ]
    ] = None,
    combination_fn: Optional[Callable[[List[pd.Series]], pd.Series]] = None,
    energy_fn: Optional[
        Callable[[pd.DataFrame, int | str, Mapping[str, object], Mapping[str, Sequence[str]], Optional[Callable]], float]
    ] = None,
    input_series_provider: Optional[
        Callable[[pd.DataFrame, int | str, str, Sequence[str]], List[pd.Series]]
    ] = None,
    apply_adjustments_to_future_years: bool = False,
    apply_adjustments_to_past_years: bool = False,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Compare modelled totals with ESTO totals and scale inputs when needed.

    Parameters
    ----------
    export_df: LEAP export data.
    base_year: Column name for the base year to reconcile.
    branch_mapping_rules: Mapping of ESTO keys to lists of branch rules.
    esto_energy_totals: ESTO totals keyed by the same ESTO keys.
    strategies: Optional mapping of strategy names to input-variable lists.
    tolerance: Difference threshold before adjustments are applied.
    adjustment_fn: Optional custom adjustment function with signature
        (export_df, base_year, rule, scale_factor, strategies, year_columns).
    apply_adjustments_to_future_years: When True, apply the same scale factor to
        all year columns after `base_year` (if present) instead of only scaling
        the base year.
    apply_adjustments_to_past_years: When True, apply the same scale factor to
        all year columns before `base_year` (if present) instead of only scaling
        the base year. - this is only relevant if `apply_adjustments_to_future_years` is also True.
    combination_fn: Optional custom function for combining input variables into
        energy use.
    energy_fn: Optional function that fully overrides how energy is calculated
        for each rule. Signature: (export_df, base_year, rule, strategies, combination_fn) -> float.
    input_series_provider: Optional function that returns the series to combine
        for a rule when you only need to override how inputs are fetched. Signature:
        (export_df, base_year, branch_path, input_variables) -> List[pd.Series].
    """
    
    working_df = export_df.copy()
    strategy_lookup = {**DEFAULT_STRATEGIES, **(strategies or {})}
    if energy_fn and not adjustment_fn:
        breakpoint()
        raise ValueError("Provide a custom adjustment_fn when using a custom energy_fn so scale factors are applied correctly.")
    adjust = adjustment_fn or _apply_proportional_adjustment
    adjustment_year_columns = get_adjustment_year_columns(
        working_df, base_year, include_future_years=apply_adjustments_to_future_years, apply_adjustments_to_past_years=apply_adjustments_to_past_years
    )

    results = []
    for esto_key, rules in branch_mapping_rules.items():
        leap_total = 0.0
        adjusted_paths: List[str] = []

        for rule in rules:
            try:
                energy = calculate_branch_energy(
                    working_df,
                    base_year,
                    rule,
                    strategy_lookup,
                    combination_fn,
                    energy_fn=energy_fn,
                    input_series_provider=input_series_provider,
                )
            except Exception:
                breakpoint()
                energy = 0.0
            leap_total += energy

        esto_total = float(esto_energy_totals.get(esto_key, 0.0))
        scale_factor = _compute_scale_factor(leap_total, esto_total)

        if abs(leap_total - esto_total) > tolerance and scale_factor != 1.0:
            for rule in rules:
                try:
                    adjust(working_df, base_year, rule, scale_factor, strategy_lookup, adjustment_year_columns)
                    adjusted_paths.append(build_branch_path(rule["branch_tuple"], root=rule.get("root", "Demand")))
                except Exception:
                    breakpoint()
                    continue

        results.append(
            {
                "ESTO Key": " | ".join(esto_key),
                "LEAP Energy Use": leap_total,
                "ESTO Energy Use": esto_total,
                "Scale Factor": scale_factor,
                "Adjusted Branches": ", ".join(adjusted_paths),
            }
        )

    summary_df = pd.DataFrame(results)
    return working_df, summary_df


# ---------------------------------------------------------------------------
# Adjustment reporting helpers
# ---------------------------------------------------------------------------

def _build_change_table_for_years(
    original_df: pd.DataFrame,
    adjusted_df: pd.DataFrame,
    years: Sequence[int | str],
    tol: float = 1e-9,
) -> pd.DataFrame:
    """Internal: build a long-form table of value changes for the provided years."""

    if not years:
        return pd.DataFrame(columns=["Branch Path", "Variable", "Year", "Original", "Adjusted", "Abs Change", "Pct Change"])

    meta_cols = [col for col in ("Scenario", "Economy") if col in original_df.columns]
    base_cols = ["Branch Path", "Variable", *meta_cols]
    frames: List[pd.DataFrame] = []

    for year_col in years:
        if year_col not in original_df.columns or year_col not in adjusted_df.columns:
            continue

        orig_vals = pd.to_numeric(original_df[year_col], errors="coerce")
        adj_vals = pd.to_numeric(adjusted_df[year_col], errors="coerce")
        diff = adj_vals - orig_vals

        mask = ~((orig_vals.isna()) & (adj_vals.isna())) & diff.abs().gt(tol)
        if not mask.any():
            continue

        pct_change = diff / orig_vals.replace(0, pd.NA)

        frame = pd.DataFrame(
            {
                "Branch Path": original_df.loc[mask, "Branch Path"],
                "Variable": original_df.loc[mask, "Variable"],
                "Year": year_col,
                "Original": orig_vals.loc[mask],
                "Adjusted": adj_vals.loc[mask],
                "Abs Change": diff.loc[mask],
                "Pct Change": pct_change.loc[mask],
            }
        )

        for meta_col in meta_cols:
            frame[meta_col] = original_df.loc[mask, meta_col]

        frames.append(frame)

    if not frames:
        return pd.DataFrame(columns=["Branch Path", "Variable", "Year", "Original", "Adjusted", "Abs Change", "Pct Change", *meta_cols])

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.sort_values(by="Abs Change", key=lambda s: s.abs(), ascending=False)
    return combined.reset_index(drop=True)


def build_adjustment_change_tables(
    original_df: pd.DataFrame,
    adjusted_df: pd.DataFrame,
    base_year: int | str,
    include_future_years: bool = False,
    tol: float = 1e-9,
) -> Tuple[pd.DataFrame, Optional[pd.DataFrame]]:
    """
    Create tables describing how LEAP inputs changed during reconciliation.

    Returns
    -------
    base_year_changes: DataFrame with differences for the base_year column.
    future_year_changes: DataFrame with differences for all columns after base_year
        when include_future_years is True. Otherwise None.
    """

    base_changes = _build_change_table_for_years(original_df, adjusted_df, [base_year], tol=tol)

    future_changes = None
    if include_future_years:
        future_years = [
            year for year in get_adjustment_year_columns(adjusted_df, base_year, include_future_years=True) if year != base_year
        ]
        future_changes = _build_change_table_for_years(original_df, adjusted_df, future_years, tol=tol)

    return base_changes, future_changes


# ---------------------------------------------------------------------------
# Convenience helper
# ---------------------------------------------------------------------------

def build_esto_totals_from_dataframe(
    esto_df: pd.DataFrame,
    key_columns: Sequence[str],
    value_column: str,
) -> Dict[Tuple[str, ...], float]:
    """Convert an ESTO DataFrame into a tuple-keyed mapping of totals."""

    return {tuple(row[col] for col in key_columns): float(row[value_column]) for _, row in esto_df.iterrows()}

