from __future__ import annotations

from pathlib import Path
from typing import Iterable, Sequence

import pandas as pd

DEFAULT_SCENARIO = "reference"
NINTH_SECTOR_COLS = [
    "sub4sectors",
    "sub3sectors",
    "sub2sectors",
    "sub1sectors",
    "sectors",
]
NINTH_FUEL_COLS = ["subfuels", "fuels"]


def normalize_economy_key(value: str | None) -> str:
    """Return a canonical economy key for cross-dataset joins."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none"}:
        return ""
    return text.replace("_", "").upper()


def _clean_label_series(series: pd.Series) -> pd.Series:
    cleaned = series.fillna("").astype(str).str.strip()
    return cleaned.mask(cleaned.str.lower() == "x", pd.NA)


def add_ninth_pair_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Add most-specific sector/fuel columns for 9th data."""
    working = df.copy()
    sector_cols = [col for col in NINTH_SECTOR_COLS if col in working.columns]
    fuel_cols = [col for col in NINTH_FUEL_COLS if col in working.columns]
    if sector_cols:
        sector_values = pd.DataFrame(
            {col: _clean_label_series(working[col]) for col in sector_cols}
        )
        working["9th_sector"] = sector_values.bfill(axis=1).iloc[:, 0].fillna("")
    else:
        working["9th_sector"] = ""
    if fuel_cols:
        fuel_values = pd.DataFrame(
            {col: _clean_label_series(working[col]) for col in fuel_cols}
        )
        working["9th_fuel"] = fuel_values.bfill(axis=1).iloc[:, 0].fillna("")
    else:
        working["9th_fuel"] = ""
    return working


def filter_ninth_projection_rows(
    df: pd.DataFrame, scenario: str = DEFAULT_SCENARIO
) -> pd.DataFrame:
    """Filter 9th data to the reference scenario and non-subtotal rows."""
    working = df.copy()
    if scenario and "scenarios" in working.columns:
        scenario_key = str(scenario).strip().lower()
        working = working[
            working["scenarios"].astype(str).str.strip().str.lower() == scenario_key
        ]
    if "subtotal_results" in working.columns:
        flag = (
            working["subtotal_results"]
            .fillna(False)
            .astype(str)
            .str.strip()
            .str.lower()
            .isin({"1", "true", "yes", "y", "t"})
        )
        working = working[~flag]
    return working


def build_ninth_projection_series(
    ninth_df: pd.DataFrame, projection_years: Sequence[int]
) -> pd.DataFrame:
    """Aggregate projected-year values by economy + 9th pair."""
    if not projection_years or ninth_df.empty:
        return pd.DataFrame()
    year_cols = [year for year in projection_years if year in ninth_df.columns]
    if not year_cols:
        return pd.DataFrame()
    working = ninth_df.copy()
    working = working[(working["9th_sector"] != "") & (working["9th_fuel"] != "")]
    if working.empty:
        return pd.DataFrame()
    for year in year_cols:
        working[year] = pd.to_numeric(working[year], errors="coerce").fillna(0.0)
    grouped = (
        working.groupby(["economy_key", "9th_sector", "9th_fuel"], dropna=False)[year_cols]
        .sum()
        .reset_index()
    )
    return grouped


def build_esto_base_year_values(
    esto_df: pd.DataFrame, base_year: int
) -> pd.DataFrame:
    """Return base-year values per economy/flow/product."""
    if esto_df.empty:
        return pd.DataFrame()
    year_col = base_year if base_year in esto_df.columns else str(base_year)
    if year_col not in esto_df.columns:
        return pd.DataFrame()
    working = esto_df.copy()
    working["economy_key"] = working["economy"].apply(normalize_economy_key)
    working["esto_flow"] = working["flows"].astype(str).str.strip()
    working["esto_product"] = working["products"].astype(str).str.strip()
    working[year_col] = pd.to_numeric(working[year_col], errors="coerce").fillna(0.0)
    grouped = (
        working.groupby(["economy_key", "esto_flow", "esto_product"], dropna=False)[
            year_col
        ]
        .sum()
        .reset_index()
        .rename(columns={year_col: "base_value"})
    )
    grouped["base_value_abs"] = grouped["base_value"].abs()
    return grouped


def compute_esto_base_year_shares(
    base_values: pd.DataFrame,
    economy_key: str,
    esto_flow: str,
    esto_products: Sequence[str],
) -> dict[str, float]:
    """Return absolute-share splits for a flow/product set in a given economy."""
    if not esto_products:
        return {}
    if base_values is None or base_values.empty:
        return {product: 1.0 / len(esto_products) for product in esto_products}
    subset = base_values[
        (base_values["economy_key"] == economy_key)
        & (base_values["esto_flow"] == esto_flow)
        & (base_values["esto_product"].isin(esto_products))
    ]
    grouped = (
        subset.groupby("esto_product", dropna=False)["base_value_abs"]
        .sum()
        .reindex(esto_products)
        .fillna(0.0)
    )
    total = float(grouped.sum())
    if total <= 0:
        return {product: 1.0 / len(esto_products) for product in esto_products}
    return {product: float(grouped.loc[product]) / total for product in esto_products}


def _build_conservation_diagnostics(
    source_by_pair: pd.DataFrame,
    allocated_rows: pd.DataFrame,
    year_cols: Sequence[int],
    tolerance: float = 1e-6,
) -> pd.DataFrame:
    """Return diagnostics proving allocation conserves source totals by 9th pair.

    For each (economy_key, 9th_sector, 9th_fuel), this compares:
    - source values: the original 9th series
    - allocated values: sum across mapped ESTO rows
    """
    if source_by_pair.empty or allocated_rows.empty or not year_cols:
        return pd.DataFrame()
    key_cols = ["economy_key", "9th_sector", "9th_fuel"]
    source_by_pair = source_by_pair[key_cols + list(year_cols)].copy()
    allocated_by_pair = (
        allocated_rows.groupby(key_cols, dropna=False)[list(year_cols)]
        .sum()
        .reset_index()
    )
    source_indexed = source_by_pair.set_index(key_cols)
    allocated_indexed = allocated_by_pair.set_index(key_cols)
    diff = allocated_indexed[list(year_cols)] - source_indexed[list(year_cols)]
    abs_diff = diff.abs()
    max_abs_diff = abs_diff.max(axis=1)
    mismatch_mask = max_abs_diff > tolerance
    if not mismatch_mask.any():
        return pd.DataFrame()

    mismatch_diff = diff[mismatch_mask]
    mismatch_abs = abs_diff[mismatch_mask]
    worst_years = mismatch_abs.idxmax(axis=1)
    rows = []
    for key, row in mismatch_diff.iterrows():
        worst_year = int(worst_years.loc[key])
        source_value = float(source_indexed.loc[key, worst_year])
        allocated_value = float(allocated_indexed.loc[key, worst_year])
        error_value = float(row[worst_year])
        rows.append(
            {
                "economy_key": key[0],
                "9th_sector": key[1],
                "9th_fuel": key[2],
                "diagnostic_type": "conservation_mismatch",
                "worst_year": worst_year,
                "source_value_worst_year": source_value,
                "allocated_value_worst_year": allocated_value,
                "allocation_error_worst_year": error_value,
                "max_abs_allocation_error": float(mismatch_abs.loc[key].max()),
                "sum_abs_allocation_error": float(mismatch_abs.loc[key].sum()),
                "year_count_above_tolerance": int((mismatch_abs.loc[key] > tolerance).sum()),
            }
        )
    return pd.DataFrame(rows)


def _resolve_sign_stable_flow_set(
    mapping: pd.DataFrame,
    sign_stable_flows: Iterable[str] | str | None,
) -> set[str]:
    """Return normalized sign-stable flow names from iterable or mode string.

    Accepted string modes:
    - "all" / "*": apply sign-stable routing to every mapped ESTO flow.
    - "off" / "none" / "": disable sign-stable routing.
    - Any other string: treated as a single flow name.
    """
    if sign_stable_flows is None:
        return set()
    if isinstance(sign_stable_flows, str):
        mode = sign_stable_flows.strip().lower()
        if mode in {"", "off", "none", "false"}:
            return set()
        if mode in {"all", "*"}:
            return {
                str(flow).strip()
                for flow in mapping["esto_flow"].dropna().astype(str)
                if str(flow).strip()
                and str(flow).strip().lower() not in {"nan", "none"}
            }
        return {sign_stable_flows.strip()}
    return {
        str(flow).strip()
        for flow in sign_stable_flows
        if str(flow).strip() and str(flow).strip().lower() not in {"nan", "none"}
    }


def allocate_ninth_projection_to_esto(
    mapping_df: pd.DataFrame,
    ninth_series: pd.DataFrame,
    base_values: pd.DataFrame,
    projection_years: Sequence[int],
    sign_stable_flows: Iterable[str] | str | None = None,
    strict_conservation: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Allocate 9th projections to ESTO pairs using base-year share rules.

    How allocation works
    --------------------
    1. Build a source series by (economy_key, 9th_sector, 9th_fuel).
    2. Use the mapping table to fan each source series out to one or more
       ESTO (flow, product) rows.
    3. Compute shares from base-year ESTO magnitudes.

    Legacy mode (default)
    ---------------------
    - Share is based on absolute base-year values.
    - Positive source values can be distributed into base-year-negative rows.
    - This preserves totals but can flip detailed row signs.

    Optional sign-stable mode (`sign_stable_flows`)
    ----------------------------------------------
    - Accepts a flow list or a mode string ("all", "off"/"none").
    - Triggered by ESTO flows listed in `sign_stable_flows`.
    - Once triggered for any mapped row, it is applied to the whole
      (economy_key, 9th_sector, 9th_fuel) source pair to preserve totals.
    - Positive source years are split only across base-year-positive targets.
    - Negative source years are split only across base-year-negative targets.
    - If no same-sign targets exist for a given sign, it falls back to legacy
      shares for that sign/year to avoid dropping totals.

    Concrete example (08_JPN, coal products split)
    ----------------------------------------------
    Source pair: (09_08_coal_transformation, 02_coal_products)
    Source 2023 value: +877.277
    Base-year target signs include:
    - 09.08.01 Coke ovens | 02.01 Coke oven coke = +833.544
    - 09.08.02 Blast furnaces | 02.01 Coke oven coke = -693.349

    Legacy split allocates to both rows by abs-share:
    - Coke ovens coke: 335.251
    - Blast furnaces coke: 278.865

    Sign-stable split (positive source) allocates only to positive-sign targets:
    - Coke ovens coke increases to 491.481
    - Blast furnaces coke becomes 0.0
    - Total remains +877.277

    Tradeoff
    --------
    Sign-stable mode reduces sign-flip artifacts from aggregated mappings but may
    suppress legitimate sign transitions in future years. Keep it scoped to flows
    known to be affected by aggregation artifacts.
    """
    if mapping_df.empty or ninth_series.empty or not projection_years:
        return pd.DataFrame(), pd.DataFrame()
    mapping = mapping_df.copy()
    mapping["9th_sector"] = mapping["9th_sector"].fillna("").astype(str).str.strip()
    mapping["9th_fuel"] = mapping["9th_fuel"].fillna("").astype(str).str.strip()
    mapping["esto_flow"] = mapping["esto_flow"].fillna("").astype(str).str.strip()
    mapping["esto_product"] = mapping["esto_product"].fillna("").astype(str).str.strip()
    mapping = mapping[(mapping["9th_sector"] != "") & (mapping["9th_fuel"] != "")]
    if mapping.empty:
        return pd.DataFrame(), pd.DataFrame()
    sign_stable_flow_set = _resolve_sign_stable_flow_set(mapping, sign_stable_flows)

    base_values = base_values.copy()
    if not base_values.empty:
        base_values["esto_flow"] = base_values["esto_flow"].astype(str).str.strip()
        base_values["esto_product"] = base_values["esto_product"].astype(str).str.strip()
        base_values["economy_key"] = base_values["economy_key"].astype(str).str.strip()

    apec_base = (
        base_values.groupby(["esto_flow", "esto_product"], dropna=False)["base_value_abs"]
        .sum()
        .reset_index()
    )
    mapping_apec = mapping.merge(apec_base, on=["esto_flow", "esto_product"], how="left")
    mapping_apec["base_value_abs"] = mapping_apec["base_value_abs"].fillna(0.0)
    mapping_apec["apec_group_total"] = mapping_apec.groupby(
        ["9th_sector", "9th_fuel"], dropna=False
    )["base_value_abs"].transform("sum")
    mapping_apec["apec_share"] = 0.0
    apec_mask = mapping_apec["apec_group_total"] > 0
    mapping_apec.loc[apec_mask, "apec_share"] = (
        mapping_apec.loc[apec_mask, "base_value_abs"]
        / mapping_apec.loc[apec_mask, "apec_group_total"]
    )

    merged = mapping.merge(
        ninth_series, on=["9th_sector", "9th_fuel"], how="inner"
    )
    merged = merged.merge(
        base_values[
            [
                "economy_key",
                "esto_flow",
                "esto_product",
                "base_value",
                "base_value_abs",
            ]
        ],
        on=["economy_key", "esto_flow", "esto_product"],
        how="left",
    )
    merged["base_value"] = pd.to_numeric(merged["base_value"], errors="coerce").fillna(0.0)
    merged["base_value_abs"] = merged["base_value_abs"].fillna(0.0)
    merged = merged.merge(
        mapping_apec[
            [
                "9th_sector",
                "9th_fuel",
                "esto_flow",
                "esto_product",
                "apec_group_total",
                "apec_share",
            ]
        ],
        on=["9th_sector", "9th_fuel", "esto_flow", "esto_product"],
        how="left",
    )
    merged["apec_group_total"] = merged["apec_group_total"].fillna(0.0)
    merged["apec_share"] = merged["apec_share"].fillna(0.0)
    merged["group_total"] = merged.groupby(
        ["economy_key", "9th_sector", "9th_fuel"], dropna=False
    )["base_value_abs"].transform("sum")
    # Equal-share fallback must be per economy + 9th pair (not global across economies).
    merged["group_count"] = merged.groupby(
        ["economy_key", "9th_sector", "9th_fuel"], dropna=False
    )["esto_flow"].transform("count").astype(float)
    merged["share"] = 0.0
    merged["share_source"] = "economy"
    economy_mask = merged["group_total"] > 0
    merged.loc[economy_mask, "share"] = (
        merged.loc[economy_mask, "base_value_abs"]
        / merged.loc[economy_mask, "group_total"]
    )
    fallback_mask = ~economy_mask
    apec_mask = fallback_mask & (merged["apec_group_total"] > 0)
    merged.loc[apec_mask, "share"] = merged.loc[apec_mask, "apec_share"]
    merged.loc[apec_mask, "share_source"] = "apec"
    equal_mask = fallback_mask & ~apec_mask
    merged.loc[equal_mask, "share"] = (
        1.0
        / merged.loc[equal_mask, "group_count"].replace(0, pd.NA)
    )
    merged.loc[equal_mask, "share_source"] = "equal"
    merged["share"] = merged["share"].fillna(0.0)
    merged["apply_sign_stable"] = False
    merged["apply_sign_stable_pair"] = False
    if sign_stable_flow_set:
        merged["apply_sign_stable"] = merged["esto_flow"].isin(sign_stable_flow_set)
        key_cols = ["economy_key", "9th_sector", "9th_fuel"]
        merged["apply_sign_stable_pair"] = (
            merged.groupby(key_cols, dropna=False)["apply_sign_stable"]
            .transform("max")
            .astype(bool)
        )
        # Build sign-specific share pools from base-year values.
        merged["base_pos_abs"] = merged["base_value_abs"].where(merged["base_value"] > 0, 0.0)
        merged["base_neg_abs"] = merged["base_value_abs"].where(merged["base_value"] < 0, 0.0)
        merged["group_positive_total"] = merged.groupby(
            ["economy_key", "9th_sector", "9th_fuel"], dropna=False
        )["base_pos_abs"].transform("sum")
        merged["group_negative_total"] = merged.groupby(
            ["economy_key", "9th_sector", "9th_fuel"], dropna=False
        )["base_neg_abs"].transform("sum")
        merged["positive_share"] = 0.0
        merged["negative_share"] = 0.0
        positive_mask = (merged["base_value"] > 0) & (merged["group_positive_total"] > 0)
        merged.loc[positive_mask, "positive_share"] = (
            merged.loc[positive_mask, "base_value_abs"]
            / merged.loc[positive_mask, "group_positive_total"]
        )
        negative_mask = (merged["base_value"] < 0) & (merged["group_negative_total"] > 0)
        merged.loc[negative_mask, "negative_share"] = (
            merged.loc[negative_mask, "base_value_abs"]
            / merged.loc[negative_mask, "group_negative_total"]
        )

    year_cols = [year for year in projection_years if year in merged.columns]
    for year in year_cols:
        merged[year] = pd.to_numeric(merged[year], errors="coerce").fillna(0.0)
    # Capture original source series before replacing year columns with allocated values.
    source_by_pair = (
        merged[["economy_key", "9th_sector", "9th_fuel"] + year_cols]
        .drop_duplicates(subset=["economy_key", "9th_sector", "9th_fuel"])
        .copy()
    )
    for year in year_cols:
        source = merged[year]
        allocated = source * merged["share"]
        if sign_stable_flow_set:
            stable_mask = merged["apply_sign_stable_pair"]
            src_positive = stable_mask & (source > 0)
            src_negative = stable_mask & (source < 0)
            has_positive_group = merged["group_positive_total"] > 0
            has_negative_group = merged["group_negative_total"] > 0
            # Route source values by sign when sign-stable mode is enabled.
            # If a sign pool does not exist, legacy allocation remains in place.
            allocated.loc[src_positive & has_positive_group] = (
                source.loc[src_positive & has_positive_group]
                * merged.loc[src_positive & has_positive_group, "positive_share"]
            )
            allocated.loc[src_negative & has_negative_group] = (
                source.loc[src_negative & has_negative_group]
                * merged.loc[src_negative & has_negative_group, "negative_share"]
            )
        merged[year] = allocated

    projection_df = (
        merged.groupby(["economy_key", "esto_flow", "esto_product"], dropna=False)[
            year_cols
        ]
        .sum()
        .reset_index()
    )
    diagnostics = merged.loc[
        merged["share_source"] != "economy",
        [
            "economy_key",
            "9th_sector",
            "9th_fuel",
            "esto_flow",
            "esto_product",
            "share_source",
            "group_total",
            "apec_group_total",
            "base_value_abs",
            "share",
            "apply_sign_stable",
            "apply_sign_stable_pair",
        ],
    ].copy()
    if not diagnostics.empty:
        diagnostics["diagnostic_type"] = "share_fallback"

    conservation_diagnostics = _build_conservation_diagnostics(
        source_by_pair,
        merged,
        year_cols,
        tolerance=1e-6,
    )
    if not conservation_diagnostics.empty:
        max_err = float(conservation_diagnostics["max_abs_allocation_error"].max())
        message = (
            "Allocation conservation check failed for "
            f"{len(conservation_diagnostics)} source pairs. "
            f"Max abs allocation error={max_err:.6e}"
        )
        if strict_conservation:
            sample_cols = [
                "economy_key",
                "9th_sector",
                "9th_fuel",
                "worst_year",
                "allocation_error_worst_year",
                "max_abs_allocation_error",
            ]
            sample = (
                conservation_diagnostics.sort_values(
                    "max_abs_allocation_error", ascending=False
                )[sample_cols]
                .head(10)
                .to_string(index=False)
            )
            raise ValueError(f"{message}\nTop mismatches:\n{sample}")
        print(f"[WARN] {message}")
        diagnostics = pd.concat([diagnostics, conservation_diagnostics], ignore_index=True, sort=False)

    if sign_stable_flow_set:
        diagnostics["sign_stable_mode"] = diagnostics["apply_sign_stable"].map(
            {True: "enabled", False: "disabled"}
        )
    return projection_df, diagnostics


def build_esto_projection_table(
    ninth_data: pd.DataFrame,
    esto_data: pd.DataFrame,
    mapping_path: str | Path,
    base_year: int,
    projection_years: Sequence[int],
    scenario: str = DEFAULT_SCENARIO,
    sign_stable_flows: Iterable[str] | str | None = None,
    strict_conservation: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return projected ESTO values plus allocation diagnostics.

    Args:
        sign_stable_flows:
            Optional ESTO flow names to allocate with sign-stable routing.
            Use `[]` or `"off"` for pure legacy abs-share behavior.
            Use `"all"` to apply sign-stable routing to every mapped ESTO flow.
        strict_conservation:
            If True, raise ValueError when allocated totals do not match source totals.
    """
    mapping_path = Path(mapping_path)
    if not mapping_path.exists():
        return pd.DataFrame(), pd.DataFrame()
    if mapping_path.suffix.lower() in {".xlsx", ".xls"}:
        mapping_df = pd.read_excel(mapping_path, dtype=str).fillna("")
    else:
        mapping_df = pd.read_csv(mapping_path, dtype=str).fillna("")
    if mapping_df.empty:
        return pd.DataFrame(), pd.DataFrame()
    ninth_filtered = filter_ninth_projection_rows(ninth_data, scenario=scenario)
    ninth_pairs = add_ninth_pair_columns(ninth_filtered)
    ninth_pairs["economy_key"] = ninth_pairs["economy"].apply(normalize_economy_key)
    ninth_series = build_ninth_projection_series(ninth_pairs, projection_years)
    base_values = build_esto_base_year_values(esto_data, base_year)
    return allocate_ninth_projection_to_esto(
        mapping_df,
        ninth_series,
        base_values,
        projection_years,
        sign_stable_flows=sign_stable_flows,
        strict_conservation=strict_conservation,
    )


def merge_projection_into_esto(
    esto_df: pd.DataFrame,
    projection_df: pd.DataFrame,
    projection_years: Sequence[int],
) -> pd.DataFrame:
    """Return an ESTO dataframe with projection years appended."""
    if not projection_years:
        return esto_df
    if projection_df is None or projection_df.empty:
        working = esto_df.copy()
        print(
            f"[INFO] No 9th projection data available; adding empty projection-year columns "
            f"({min(projection_years)}–{max(projection_years)}) to ESTO base-year data."
        )
        for year in projection_years:
            if year not in working.columns:
                working[year] = 0.0
        return working
    working = esto_df.copy()
    print(
        f"[INFO] Merging 9th projections into ESTO data for years "
        f"{min(projection_years)}–{max(projection_years)}."
    )
    working["economy_key"] = working["economy"].apply(normalize_economy_key)
    working["flows"] = working["flows"].astype(str).str.strip()
    working["products"] = working["products"].astype(str).str.strip()

    proj = projection_df.copy()
    proj["esto_flow"] = proj["esto_flow"].astype(str).str.strip()
    proj["esto_product"] = proj["esto_product"].astype(str).str.strip()
    proj_cols = [year for year in projection_years if year in proj.columns]
    if not proj_cols:
        return esto_df
    proj = proj.rename(columns={year: f"{year}_proj" for year in proj_cols})

    merged = working.merge(
        proj,
        left_on=["economy_key", "flows", "products"],
        right_on=["economy_key", "esto_flow", "esto_product"],
        how="left",
    )
    for year in proj_cols:
        proj_col = f"{year}_proj"
        merged[year] = merged[proj_col].fillna(0.0)
    drop_cols = [
        "economy_key",
        "esto_flow",
        "esto_product",
    ] + [f"{year}_proj" for year in proj_cols]
    merged = merged.drop(columns=[col for col in drop_cols if col in merged.columns])

    base_cols = [col for col in esto_df.columns if col not in proj_cols]
    existing_years = [col for col in base_cols if str(col).isdigit()]
    non_year_cols = [col for col in base_cols if col not in existing_years]
    ordered_years = sorted(set(existing_years + proj_cols))
    ordered_cols = non_year_cols + ordered_years
    merged = merged[ordered_cols]
    return merged


def build_projection_lookup(projection_df: pd.DataFrame) -> pd.DataFrame | None:
    """Return a MultiIndex lookup for projection values."""
    if projection_df is None or projection_df.empty:
        return None
    grouped = (
        projection_df.groupby(
            ["economy_key", "esto_flow", "esto_product"], dropna=False
        )
        .sum(numeric_only=True)
    )
    return grouped
