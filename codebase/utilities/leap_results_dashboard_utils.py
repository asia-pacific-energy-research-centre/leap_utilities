#%%
"""
Utility helpers for the LEAP Results dashboard workflow.

Functions are designed for notebook-first usage with clear toggles and small,
composable pieces. The helpers:
- load sheet/sector mappings
- normalize fuel labels using canonical codebooks (with optional backup overrides)
- parse LEAP result workbooks (template-style sheets)
- pull reference (ESTO) and projection (9th) series
- assemble comparison DataFrames and lightweight status diagnostics
- generate simple charts and HTML dashboards (style reused from leap_transport)
"""
from __future__ import annotations

import math
import os
import re
import sys
from pathlib import Path
from typing import Optional, Sequence

import pandas as pd

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.functions.ninth_projection_mapping import normalize_economy_key
from codebase.mappings.canonical_mapping import (
    build_sector_to_esto_flow_lookup as shared_build_sector_to_esto_flow_lookup,
    clean_token as shared_clean_token,
    load_canonical_pairs as shared_load_canonical_pairs,
    load_fuel_aliases as shared_load_fuel_aliases,
    load_sheet_map as shared_load_sheet_map,
    map_fuel_label as shared_map_fuel_label,
    normalize_label as shared_normalize_label,
    split_sector_codes as shared_split_sector_codes,
)

# Stable paths (overridable by caller)
DEFAULT_SHEET_MAP = Path("config/leap_results_sheet_map.csv")
DEFAULT_BACKUP_LEAP_MAPPINGS = Path("config/backup_leap_mappings.xlsx")
DEFAULT_CODEBOOK = Path("config/sector_fuel_codes_to_names.xlsx")
DEFAULT_NINTH_FUEL_PAIRS = Path("config/ninth_sector_fuel_pairs.csv")
DEFAULT_NINTH_TO_ESTO = Path("config/ninth_pairs_to_esto_pairs.xlsx")


# -----------------------------------------------------------------------------
# Helpers: path / repo handling
# -----------------------------------------------------------------------------
REPO_ROOT = Path(__file__).resolve().parents[2]


def ensure_repo_root() -> None:
    """Force cwd to repo root so relative paths work in notebooks."""
    cwd = Path.cwd()
    if cwd != REPO_ROOT:
        os.chdir(REPO_ROOT)


# -----------------------------------------------------------------------------
# Mapping loaders
# -----------------------------------------------------------------------------
def load_sheet_map(path: Path = DEFAULT_SHEET_MAP) -> pd.DataFrame:
    """Read sheet→sector map and return active rows with normalized names."""
    return shared_load_sheet_map(path)


def _split_sector_codes(raw_value: object) -> list[str]:
    """
    Parse one-to-many sector mapping tokens from a sheet-map cell.
    Accepted separators: ',', ';', '|', 'AND'.
    """
    return shared_split_sector_codes(raw_value)


def _build_codebook_lookup(codebook_path: Path) -> dict[str, str]:
    """Map human-readable names to 9th fuel codes."""
    df = pd.read_excel(codebook_path, sheet_name="code_to_name")
    lookup: dict[str, str] = {}
    for _, row in df.iterrows():
        name = str(row.get("name") or "").strip().lower()
        code = str(row.get("9th_label") or "").strip()
        if name and code:
            lookup[name] = code
    return lookup


def _clean_token(value: object) -> str:
    return shared_clean_token(value)


def _build_name_to_esto_product(codebook_path: Path) -> dict[str, str]:
    """Map human-readable names to ESTO product codes."""
    df = pd.read_excel(codebook_path, sheet_name="code_to_name")
    lookup: dict[str, str] = {}
    for _, row in df.iterrows():
        name = str(row.get("name") or "").strip().lower()
        est = str(row.get("esto_label") or "").strip()
        if name and est:
            lookup[name] = est
    return lookup


def build_sector_to_esto_flow_lookup(codebook_path: Path = DEFAULT_CODEBOOK) -> dict[str, str]:
    """Map 9th sector codes to ESTO flow labels from code_to_name."""
    return shared_build_sector_to_esto_flow_lookup(codebook_path)


def _build_leap_esto_lookup(codebook_path: Path) -> dict[str, str]:
    """
    Build mapping from LEAP fuel label (as used in LEAP exports) to 9th fuel code
    by chaining through ESTO_LEAP_names (LEAP label -> ESTO label) and code_to_name
    (ESTO label -> 9th label).
    """
    df_leap = pd.read_excel(codebook_path, sheet_name="ESTO_LEAP_names")
    df_code = pd.read_excel(codebook_path, sheet_name="code_to_name")
    esto_to_9th = {
        str(r["esto_label"]).strip().lower(): str(r["9th_label"]).strip()
        for _, r in df_code.iterrows()
        if pd.notna(r.get("esto_label")) and pd.notna(r.get("9th_label"))
    }
    lookup: dict[str, str] = {}
    for _, row in df_leap.iterrows():
        if str(row.get("category")).strip().lower() != "products":
            continue
        leap_label = str(row.get("leap_name") or "").strip().lower()
        esto_label = str(row.get("original_label") or "").strip().lower()
        if not leap_label or not esto_label:
            continue
        ninth = esto_to_9th.get(esto_label, "")
        if ninth:
            lookup[leap_label] = ninth
    return lookup


def _normalize_label(value: object) -> str:
    """Lowercase and collapse whitespace for robust text joins."""
    return shared_normalize_label(value)


def load_canonical_pairs(
    path: Path = DEFAULT_NINTH_TO_ESTO,
    *,
    strict: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Load canonical 9th->ESTO pairs and return (clean_pairs, conflicts).
    Hard conflict: same (9th_sector, 9th_fuel) maps to inconsistent
    (esto_flow, esto_product).
    """
    return shared_load_canonical_pairs(path=path, strict=strict)


def _extract_esto_code(esto_label: str) -> str:
    """Extract ESTO code prefix from labels like '07.12 White spirit SBP'."""
    match = re.match(r"^\s*(\d{2}(?:\.\d{2})?)\b", str(esto_label or ""))
    if not match:
        return ""
    return match.group(1).replace(".", "_")


def _build_ninth_fuel_lookup(ninth_fuel_pairs_path: Path) -> tuple[dict[str, str], dict[str, str]]:
    """
    Build lookups from ESTO-like numeric prefixes to valid 9th fuel codes.
    Returns:
    - exact_lookup: 07_01 -> 07_01_motor_gasoline (if available in pairs)
    - group_lookup: 07 -> 07_petroleum_products
    """
    pairs = pd.read_csv(ninth_fuel_pairs_path)
    exact_lookup: dict[str, str] = {}
    group_lookup: dict[str, str] = {}

    # Group mapping (e.g., 07 -> 07_petroleum_products) comes only from fuels.
    if "fuels" in pairs.columns:
        fuel_codes = (
            pairs["fuels"]
            .astype(str)
            .str.strip()
            .replace({"": pd.NA, "x": pd.NA})
            .dropna()
            .unique()
            .tolist()
        )
        for code in sorted(fuel_codes):
            m2 = re.match(r"^(\d{2})_", code)
            if m2:
                group_lookup.setdefault(m2.group(1), code)

    # Exact mapping (e.g., 07_01 -> 07_01_motor_gasoline) from detailed fuel codes.
    cols = [c for c in ["fuel_pair", "subfuels"] if c in pairs.columns]
    candidates = (
        pd.concat([pairs[c].astype(str) for c in cols], ignore_index=True)
        .str.strip()
        .replace({"": pd.NA, "x": pd.NA})
        .dropna()
        .unique()
        .tolist()
    )
    for code in sorted(candidates):
        m = re.match(r"^(\d{2}(?:_\d{2})?)_", code)
        if m:
            exact_lookup.setdefault(m.group(1), code)
    return exact_lookup, group_lookup


def load_fuel_aliases(
    alias_path: Path | str | None = DEFAULT_BACKUP_LEAP_MAPPINGS,
    codebook_path: Path = DEFAULT_CODEBOOK,
) -> dict[str, dict[str, str]]:
    """
    Build a mapping from LEAP fuel labels to ESTO products (+ optional explicit overrides).
    Priority:
    1) codebook-driven mapping (ESTO_LEAP_names + code_to_name name harmonization)
    2) explicit backup overrides (optional, wins)
    Returns dict keyed by normalized leap fuel label.
    """
    return shared_load_fuel_aliases(alias_path=alias_path, codebook_path=codebook_path)


def map_fuel_label(
    fuel_label: str,
    fuel_mapping: dict[str, dict[str, str]],
    fallback_codebook: dict[str, str] | None = None,
) -> dict[str, str]:
    """
    Return mapping hints for a LEAP fuel label.
    """
    return shared_map_fuel_label(
        fuel_label=fuel_label,
        fuel_mapping=fuel_mapping,
        fallback_codebook=fallback_codebook,
    )


# -----------------------------------------------------------------------------
# LEAP workbook parsing
# -----------------------------------------------------------------------------
def parse_template_sheet(sheet: pd.DataFrame) -> dict:
    """
    Extract metadata and series from a template-style LEAP results sheet.
    Expected structure (matching leap_results_workflow template refills):
    row0: variable
    row1: "Scenario: X, Region: Y"
    row2: "Branch: ..."
    row3: "Units: ..."
    row5: header with legend label + years
    row6+: legend members with values per year
    """
    meta: dict[str, object] = {}
    meta["variable"] = str(sheet.iloc[0, 0]).strip()
    scenario_region = str(sheet.iloc[1, 0])
    for part in scenario_region.split(","):
        if "Scenario:" in part:
            meta["scenario"] = part.split(":", 1)[1].strip()
        if "Region:" in part:
            meta["region"] = part.split(":", 1)[1].strip()
    meta["branch"] = str(sheet.iloc[2, 0]).split(":", 1)[-1].strip()
    meta["units"] = str(sheet.iloc[3, 0]).split(":", 1)[-1].strip()
    meta["legend_label"] = str(sheet.iloc[5, 0]).strip()
    def _parse_year_token(value: object) -> int | None:
        if pd.isna(value):
            return None
        # Direct numeric parse first.
        try:
            num = float(value)
            if math.isfinite(num):
                direct = int(round(num))
                if 1900 <= direct <= 2200:
                    return direct
                # Some LEAP exports encode years as tiny scientific notation, e.g. 2.022e-12.
                for factor in (1e3, 1e6, 1e9, 1e12, 1e15):
                    scaled = int(round(num * factor))
                    if 1900 <= scaled <= 2200:
                        return scaled
        except Exception:
            pass

        # String fallback: extract a 4-digit year if present.
        text = str(value).strip()
        m = re.search(r"\b(19\d{2}|20\d{2}|21\d{2})\b", text)
        if m:
            return int(m.group(1))
        return None

    year_cols: list[tuple[int, int]] = []
    for col_idx, val in enumerate(sheet.iloc[5, 1:], start=1):
        year_int = _parse_year_token(val)
        if year_int is None:
            # Skip non-year tokens such as "Total"
            continue
        year_cols.append((col_idx, year_int))
    years = [y for _, y in year_cols]
    meta["years"] = years
    records: list[dict] = []
    for _, row in sheet.iloc[6:, :].iterrows():
        fuel = str(row.iloc[0]).strip()
        if not fuel or pd.isna(fuel):
            break
        fuel_lower = fuel.lower()
        if fuel_lower in {
            "total",
            "demand total",
            "international transport",
            "freight road",
            "freight non road",
            "nonspecified transport",
        }:
            continue
        for col_idx, year in year_cols:
            val = row.iloc[col_idx]
            try:
                num = float(val)
            except Exception:
                num = float("nan")
            records.append(
                {
                    "fuel_label": fuel,
                    "year": year,
                    "leap_value": num,
                }
            )
    return {"meta": meta, "records": pd.DataFrame(records)}


def load_leap_workbook(
    workbook: Path,
    sheet_map: pd.DataFrame,
    expected_scenario: Optional[str] = None,
) -> pd.DataFrame:
    """
    Load all mapped sheets from a LEAP results workbook into long form.
    Columns: economy, scenario, sheet_name, sector_code_9th, fuel_label, year, leap_value
    """
    xl = pd.ExcelFile(workbook)
    rows: list[pd.DataFrame] = []
    for _, mapping in sheet_map.iterrows():
        sheet_name = mapping["sheet_name"]
        if sheet_name not in xl.sheet_names:
            continue
        sheet_df = xl.parse(sheet_name, header=None)
        parsed = parse_template_sheet(sheet_df)
        meta = parsed["meta"]
        scenario = str(meta.get("scenario") or expected_scenario or "").strip()
        region = str(meta.get("region") or "").strip()
        df = parsed["records"].copy()
        df["sheet_name"] = sheet_name
        df["sector_code_9th"] = mapping["sector_code_9th"]
        df["sector_name"] = mapping.get("sector_name", "")
        df["scenario"] = scenario
        df["region"] = region
        # Infer economy code from filename tokens if region missing
        economy_token = None
        m = re.search(r"_([A-Z]{3})_", workbook.name)
        if m:
            economy_token = m.group(1)
        df["economy"] = economy_token or region or ""
        rows.append(df)
    if not rows:
        return pd.DataFrame(
            columns=[
                "economy",
                "scenario",
                "sheet_name",
                "sector_code_9th",
                "sector_name",
                "fuel_label",
                "year",
                "leap_value",
            ]
        )
    return pd.concat(rows, ignore_index=True)


# -----------------------------------------------------------------------------
# Reference/projection data handlers
# -----------------------------------------------------------------------------
SECTOR_COLUMNS = ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]
FUEL_COLUMNS = ["fuels", "subfuels"]


def _filter_ninth_by_sector_fuel(
    ninth_df: pd.DataFrame,
    sector_code: str,
    fuel_code: str,
    economy_code: str,
    scenario: str,
) -> pd.DataFrame:
    """Return rows for matching economy, scenario, sector, and fuel."""
    working = ninth_df.copy()
    working = working[(working["economy"] == economy_code) & (working["scenarios"] == scenario)]
    # Remove subtotal rows to avoid double-counting when aggregating projection series.
    # In merged 9th data, subtotal rows often duplicate detailed rows (or contain totals),
    # which can inflate sector/fuel sums substantially.
    if "subtotal_results" in working.columns:
        working = working[~working["subtotal_results"].astype(str).str.lower().isin({"true", "1", "yes"})]
    if "subtotal_layout" in working.columns:
        working = working[~working["subtotal_layout"].astype(str).str.lower().isin({"true", "1", "yes"})]
    if working.empty:
        return working
    sector_mask = False
    for col in SECTOR_COLUMNS:
        if col in working.columns:
            sector_mask |= working[col].astype(str).str.lower() == sector_code.lower()
    working = working[sector_mask]
    if working.empty:
        return working
    if fuel_code:
        fuel_mask = False
        for col in FUEL_COLUMNS:
            if col in working.columns:
                fuel_mask |= working[col].astype(str).str.lower() == fuel_code.lower()
        working = working[fuel_mask]
    return working


def _extract_year_series(df: pd.DataFrame, years: Sequence[int]) -> pd.Series:
    """Return a series indexed by year with numeric values (NaN preserved)."""
    if df.empty:
        return pd.Series(dtype="float64", index=years)
    year_col_map: dict[int, object] = {}
    for year in years:
        if year in df.columns:
            year_col_map[int(year)] = year
            continue
        year_str = str(int(year))
        if year_str in df.columns:
            year_col_map[int(year)] = year_str
    if not year_col_map:
        return pd.Series(dtype="float64", index=years)
    # sum across matching rows (common convention in 9th data)
    summed = df[list(year_col_map.values())].apply(pd.to_numeric, errors="coerce").sum()
    summed.index = [int(str(col)) for col in summed.index]
    # Reindex to requested years to keep shape stable.
    return summed.reindex([int(y) for y in years])


def pull_projection_series(
    ninth_df: pd.DataFrame,
    sector_code: str,
    fuel_code: str,
    economy_code: str,
    scenario: str,
    projection_years: Sequence[int],
) -> pd.Series:
    filtered = _filter_ninth_by_sector_fuel(ninth_df, sector_code, fuel_code, economy_code, scenario)
    return _extract_year_series(filtered, projection_years)


def pull_base_year_value(
    esto_df: pd.DataFrame,
    base_year: int,
    economy_code: str,
    esto_flow: str,
    esto_product: str,
) -> float:
    working = esto_df.copy()
    working = working[(working["economy"] == economy_code)]
    if esto_flow:
        working = working[working["flows"].astype(str).str.lower() == esto_flow.lower()]
    if eso_product := esto_product:
        working = working[working["products"].astype(str).str.lower() == eso_product.lower()]
    # If parent-flow exact match is unavailable (e.g., subtotal rows removed upstream),
    # fallback to summing child flows under that parent code (e.g., 14.03.*).
    if working.empty and esto_flow and eso_product:
        parent = str(esto_flow).strip().lower()
        parent_code_match = re.match(r"^(\d+(?:\.\d+)*)", parent)
        parent_code = parent_code_match.group(1) if parent_code_match else ""
        fallback = esto_df.copy()
        fallback = fallback[(fallback["economy"] == economy_code)]
        fallback = fallback[fallback["products"].astype(str).str.lower() == eso_product.lower()]
        if parent_code:
            flow_codes = fallback["flows"].astype(str).str.extract(r"^(\d+(?:\.\d+)*)", expand=False).fillna("")
            fallback = fallback[flow_codes.str.startswith(parent_code + ".")]
        else:
            fallback = fallback[fallback["flows"].astype(str).str.lower().str.startswith(parent + ".")]
        working = fallback
    try:
        return float(pd.to_numeric(working[str(base_year)], errors="coerce").sum())
    except Exception:
        return float("nan")


def aggregate_esto_by_ninth_pairs(
    esto_df: pd.DataFrame,
    ninth_pairs: pd.DataFrame,
    base_year: int,
    economy_code: str,
) -> pd.DataFrame:
    """
    Attach 9th sector/fuel codes to ESTO flows/products using pairs file, then aggregate.
    Returns long DataFrame: economy, scenario, sheet, fuel_label, source, year, value, ninth_sector, ninth_fuel.
    """
    working = esto_df.copy()
    working["flows_norm"] = working["flows"].astype(str).str.strip().str.lower()
    working["products_norm"] = working["products"].astype(str).str.strip().str.lower()
    pairs = ninth_pairs.copy()
    pairs["esto_flow_norm"] = pairs["esto_flow"].astype(str).str.strip().str.lower()
    pairs["esto_product_norm"] = pairs["esto_product"].astype(str).str.strip().str.lower()

    merged = working.merge(
        pairs[["esto_flow_norm", "esto_product_norm", "9th_sector", "9th_fuel"]],
        left_on=["flows_norm", "products_norm"],
        right_on=["esto_flow_norm", "esto_product_norm"],
        how="inner",
    )
    if merged.empty:
        return pd.DataFrame(columns=["economy", "ninth_sector", "ninth_fuel", "year", "value"])

    year_cols = [str(base_year)] + [c for c in merged.columns if c.isdigit()]
    value_cols = [c for c in year_cols if c in merged.columns]
    melted = merged.melt(
        id_vars=["economy", "9th_sector", "9th_fuel"],
        value_vars=value_cols,
        var_name="year",
        value_name="value",
    )
    melted["year"] = pd.to_numeric(melted["year"], errors="coerce")
    melted = melted[(melted["economy"] == economy_code) & melted["year"].notna()]
    melted["value"] = pd.to_numeric(melted["value"], errors="coerce")
    agg = (
        melted.groupby(["economy", "9th_sector", "9th_fuel", "year"], as_index=False)["value"]
        .sum(min_count=1)
    )
    return agg


# -----------------------------------------------------------------------------
# Comparison assembly
# -----------------------------------------------------------------------------
def build_comparisons(
    leap_long: pd.DataFrame,
    sheet_map: pd.DataFrame,
    fuel_mapping: dict[str, dict[str, str]],
    sector_flow_mapping: dict[str, str],
    ninth_pairs: pd.DataFrame,
    base_df: pd.DataFrame,
    ninth_df: pd.DataFrame,
    *,
    base_year: int,
    base_economy: str,
    projection_economy: str,
    projection_years: Sequence[int],
    scenario_map: dict[str, str],
    use_esto_agg_only: bool = False,
    sibling_comparator_mode: str = "none",
    include_sibling_parent_totals: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Return comparison_long, comparison_wide, mapping_status.
    - comparison_long: economy, scenario, sheet, fuel, source, year, value
    - comparison_wide: economy, scenario, sheet, fuel, year, leap_value, base_value, projection_value
    - mapping_status: per fuel mapping diagnostics
    """
    status_rows: list[dict] = []
    long_rows: list[dict] = []
    pairs, _ = load_canonical_pairs(DEFAULT_NINTH_TO_ESTO, strict=False) if ninth_pairs.empty else (ninth_pairs.copy(), pd.DataFrame())
    for col in ["9th_sector", "9th_fuel", "esto_flow", "esto_product"]:
        pairs[col] = pairs[col].map(_clean_token)
    for col in ["sector_match_method", "fuel_match_method", "mapping_note"]:
        if col not in pairs.columns:
            pairs[col] = ""
        pairs[col] = pairs[col].map(_clean_token)
    pairs = pairs[(pairs["9th_sector"] != "") & (pairs["9th_fuel"] != "")]
    pairs = pairs[(pairs["esto_flow"] != "") & (pairs["esto_product"] != "")]
    pairs["sector_norm"] = pairs["9th_sector"].str.lower()
    pairs["fuel_norm"] = pairs["9th_fuel"].str.lower()
    pairs["esto_flow_norm"] = pairs["esto_flow"].str.lower()
    pairs["esto_product_norm"] = pairs["esto_product"].str.lower()
    projection_cache: dict[tuple[str, str, str], pd.Series] = {}
    base_cache: dict[tuple[str, str], float] = {}
    scenario_values = {str(v).strip().lower() for v in scenario_map.values()}
    ninth_projection_df = ninth_df.copy()
    if "economy" in ninth_projection_df.columns:
        ninth_projection_df = ninth_projection_df[ninth_projection_df["economy"] == projection_economy]
    if "scenarios" in ninth_projection_df.columns and scenario_values:
        ninth_projection_df = ninth_projection_df[
            ninth_projection_df["scenarios"].astype(str).str.strip().str.lower().isin(scenario_values)
        ]

    def _resolve_sector_flow(sector_codes: list[str]) -> str:
        for sector_code in sector_codes:
            key = str(sector_code or "").strip().lower()
            if not key:
                continue
            if key in sector_flow_mapping:
                return sector_flow_mapping[key]
            # Fallback: detailed sector -> nearest mapped parent (e.g., 15_02_01_* -> 15_02_road).
            prefix_match = re.match(r"^(\d{2}_\d{2})_", key)
            if prefix_match:
                prefix = prefix_match.group(1) + "_"
                candidates = [k for k in sector_flow_mapping if k.startswith(prefix)]
                if candidates:
                    best = min(candidates, key=len)
                    return sector_flow_mapping.get(best, "")
        return ""

    def _format_scenario_label(value: object) -> str:
        raw = str(value or "").strip()
        if not raw:
            return raw
        lowered = raw.lower()
        if lowered == "reference":
            return "Reference"
        if lowered == "target":
            return "Target"
        return raw

    def _sector_root_group(raw_sector_codes: object) -> str:
        """
        Build a coarse sector-family key so sibling branches (e.g., 15_02_01_*)
        can share one parent comparator without mixing unrelated families.
        """
        codes = _split_sector_codes(raw_sector_codes)
        if not codes:
            return ""
        roots: list[str] = []
        for code in codes:
            token = str(code or "").strip().lower()
            if not token:
                continue
            m3 = re.match(r"^(\d{2}_\d{2}_\d{2})", token)
            if m3:
                roots.append(m3.group(1))
                continue
            m2 = re.match(r"^(\d{2}_\d{2})", token)
            if m2:
                roots.append(m2.group(1))
                continue
            roots.append(token)
        if not roots:
            return ""
        return " | ".join(sorted(set(roots)))

    def _preserve_signed_values(esto_flow: str) -> bool:
        """
        Keep signed values only for flows where input/output direction matters,
        notably TPES/supply and transformation/own-use style balances.
        """
        flow = str(esto_flow or "").strip().lower()
        if not flow:
            return False
        if "tpes" in flow or "total primary energy supply" in flow:
            return True
        # ESTO numbering conventions: 07.* supply, 09.* transformation, 10.* own use.
        if flow.startswith(("07", "09", "10")):
            return True
        if "transformation" in flow:
            return True
        return False

    def _sector_match_subset(df: pd.DataFrame, sector_codes: list[str]) -> pd.DataFrame:
        if df.empty:
            return df
        scored: list[pd.DataFrame] = []
        for sector in sector_codes:
            key = str(sector or "").strip().lower()
            if not key:
                continue
            exact = df[df["sector_norm"] == key].copy()
            if not exact.empty:
                exact["match_priority"] = 0
                scored.append(exact)
            child = df[df["sector_norm"].str.startswith(key + "_")].copy()
            if not child.empty:
                child["match_priority"] = 1
                scored.append(child)
        if not scored:
            return df.iloc[0:0].copy()
        merged = pd.concat(scored, ignore_index=True).drop_duplicates()
        best = int(merged["match_priority"].min())
        return merged[merged["match_priority"] == best].drop(columns=["match_priority"], errors="ignore")

    def _canonical_by_sector_and_fuel(sector_codes: list[str], fuel_code: str) -> pd.DataFrame:
        if not fuel_code:
            return pd.DataFrame(columns=pairs.columns)
        subset = pairs[pairs["fuel_norm"] == fuel_code.strip().lower()]
        return _sector_match_subset(subset, sector_codes)

    def _canonical_by_sector_and_product(sector_codes: list[str], esto_product: str) -> pd.DataFrame:
        if not esto_product:
            return pd.DataFrame(columns=pairs.columns)
        subset = pairs[pairs["esto_product_norm"] == esto_product.strip().lower()]
        return _sector_match_subset(subset, sector_codes)

    def _choose_single_candidate(df: pd.DataFrame) -> tuple[str, str, str, str]:
        if df.empty:
            return "", "", "", ""
        unique_pairs = df[["9th_fuel", "esto_flow", "esto_product"]].drop_duplicates().sort_values(
            ["9th_fuel", "esto_flow", "esto_product"]
        )
        if len(unique_pairs) != 1:
            return "", "", "", ""
        row = unique_pairs.iloc[0]
        match = df[
            (df["9th_fuel"] == row["9th_fuel"])
            & (df["esto_flow"] == row["esto_flow"])
            & (df["esto_product"] == row["esto_product"])
        ].copy()
        if not match.empty and "sector_match_method" in match.columns:
            match = match.sort_values(["sector_match_method", "fuel_match_method", "mapping_note"], na_position="last")
            method = _clean_token(match.iloc[0].get("sector_match_method"))
        else:
            method = ""
        return _clean_token(row["9th_fuel"]), _clean_token(row["esto_flow"]), _clean_token(row["esto_product"]), method

    def _infer_fuel_from_flow_product(sector_codes: list[str], esto_flow: str, esto_product: str) -> str:
        if not esto_flow or not esto_product:
            return ""
        subset = pairs[
            (pairs["esto_flow_norm"] == esto_flow.strip().lower())
            & (pairs["esto_product_norm"] == esto_product.strip().lower())
        ]
        match = _sector_match_subset(subset, sector_codes)
        vals = sorted(match["9th_fuel"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())
        if len(vals) == 1:
            return vals[0]
        # Fallback 1: global flow+product unique fuel across canonical pairs.
        vals_global = sorted(
            subset["9th_fuel"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist()
        )
        if len(vals_global) == 1:
            return vals_global[0]
        # Fallback 2: product-only unique fuel (independent of flow/sector).
        prod_subset = pairs[pairs["esto_product_norm"] == esto_product.strip().lower()]
        vals_prod = sorted(
            prod_subset["9th_fuel"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist()
        )
        if len(vals_prod) == 1:
            return vals_prod[0]
        return ""

    for (sheet, fuel), sub in leap_long.groupby(["sheet_name", "fuel_label"], dropna=False):
        sheet_row = sheet_map.loc[sheet_map["sheet_name"] == sheet].iloc[0]
        sector_codes = _split_sector_codes(sheet_row.get("sector_code_9th"))
        if not sector_codes:
            sector_codes = [str(sheet_row.get("sector_code_9th") or "").strip()]
        sheet_flow_override = ""
        if "esto_flow_override" in sheet_map.columns:
            sheet_flow_override = _clean_token(sheet_row.get("esto_flow_override"))

        mapped_hint = map_fuel_label(fuel, fuel_mapping)
        ninth_fuel = _clean_token(mapped_hint.get("ninth_fuel"))
        esto_product_hint = _clean_token(mapped_hint.get("esto_product"))
        esto_flow_hint = _clean_token(mapped_hint.get("esto_flow"))
        mapping_source = _clean_token(mapped_hint.get("mapping_source")) or "canonical"
        flow_source = _clean_token(mapped_hint.get("flow_source"))
        fuel_source = _clean_token(mapped_hint.get("fuel_source"))
        sector_match_method = ""
        mapping_note = ""

        c_by_fuel = _canonical_by_sector_and_fuel(sector_codes, ninth_fuel)
        fuel_match_count = len(c_by_fuel[["9th_fuel", "esto_flow", "esto_product"]].drop_duplicates()) if not c_by_fuel.empty else 0
        c_ninth, c_flow, c_prod, c_sector_method = _choose_single_candidate(c_by_fuel)

        if c_ninth and c_flow and c_prod:
            ninth_fuel, esto_flow, esto_product = c_ninth, c_flow, c_prod
            sector_match_method = c_sector_method
            mapping_source = "override" if mapping_source == "override" else "canonical"
            if flow_source != "override":
                flow_source = "canonical"
            if fuel_source != "override":
                fuel_source = "canonical"
        else:
            if fuel_match_count > 1 and ninth_fuel:
                mapping_note = "ambiguous canonical matches for sector+fuel"
            c_by_product = _canonical_by_sector_and_product(sector_codes, esto_product_hint)
            prod_match_count = len(c_by_product[["9th_fuel", "esto_flow", "esto_product"]].drop_duplicates()) if not c_by_product.empty else 0
            p_ninth, p_flow, p_prod, p_sector_method = _choose_single_candidate(c_by_product)
            if p_ninth and p_flow and p_prod:
                ninth_fuel = ninth_fuel or p_ninth
                esto_flow = _clean_token(esto_flow_hint) or p_flow
                esto_product = _clean_token(esto_product_hint) or p_prod
                sector_match_method = p_sector_method
                if mapping_source != "override":
                    mapping_source = "codebook_fallback"
                if flow_source != "override":
                    flow_source = "canonical"
                fuel_source = "override" if fuel_source == "override" else "inferred"
            else:
                if prod_match_count > 1 and esto_product_hint:
                    mapping_note = "ambiguous canonical matches for sector+esto_product"
                esto_flow = _clean_token(esto_flow_hint)
                esto_product = _clean_token(esto_product_hint)
                if sheet_flow_override:
                    esto_flow = sheet_flow_override
                    flow_source = "sheet_override"
                elif not esto_flow:
                    fallback_flow = _resolve_sector_flow(sector_codes)
                    if fallback_flow:
                        esto_flow = _clean_token(fallback_flow)
                        flow_source = "sector_fallback"
                if not ninth_fuel:
                    inferred = _infer_fuel_from_flow_product(sector_codes, esto_flow, esto_product)
                    if inferred:
                        ninth_fuel = inferred
                        fuel_source = "inferred"
                if not mapping_source:
                    mapping_source = "override" if fuel_source == "override" else "codebook_fallback"

        if not flow_source and esto_flow:
            flow_source = "canonical"
        if not fuel_source and ninth_fuel:
            fuel_source = "canonical"
        mapped_flag = bool(ninth_fuel or esto_flow or esto_product)
        preserve_sign = _preserve_signed_values(esto_flow)
        status_rows.append(
            {
                "sheet": sheet,
                "fuel_label": fuel,
                "sector_code_9th": " | ".join(sector_codes),
                "ninth_fuel_code": ninth_fuel,
                "esto_flow": esto_flow,
                "esto_product": esto_product,
                "mapped": mapped_flag,
                "mapping_source": mapping_source or "",
                "flow_source": flow_source or "",
                "fuel_source": fuel_source or "",
                "sector_match_method": sector_match_method or "",
                "mapping_note": mapping_note,
            }
        )

        for scenario_key, sub_scenario in sub.groupby("scenario", dropna=False):
            scenario_label = str(scenario_key or "").strip()
            projection_scenario = scenario_map.get(scenario_label.lower(), "reference")
            scenario_display = _format_scenario_label(scenario_label or projection_scenario)

            # LEAP series
            leap_economy_raw = normalize_economy_key(sub_scenario["economy"].iloc[0])
            leap_economy = leap_economy_raw
            raw_token = str(leap_economy_raw or "").strip().upper()
            proj_token = str(projection_economy or "").strip().upper()
            if raw_token and proj_token.endswith(raw_token):
                leap_economy = projection_economy
            for _, row in sub_scenario.iterrows():
                long_rows.append(
                    {
                        "economy": leap_economy,
                        "scenario": scenario_display,
                        "sheet": sheet,
                        "fuel_label": fuel,
                        "source": "leap",
                        "year": int(row["year"]),
                        "value": float(row["leap_value"]),
                    }
                )

            if use_esto_agg_only:
                # ESTO-aggregated reference only (no projection)
                base_key = (esto_flow.strip().lower(), esto_product.strip().lower())
                if base_key not in base_cache:
                    base_cache[base_key] = pull_base_year_value(
                        base_df,
                        base_year=base_year,
                        economy_code=base_economy,
                        esto_flow=esto_flow,
                        esto_product=esto_product,
                    )
                base_value = base_cache[base_key]
                if not preserve_sign and not pd.isna(base_value):
                    base_value = abs(float(base_value))
                long_rows.append(
                    {
                        "economy": base_economy,
                        "scenario": scenario_display,
                        "sheet": sheet,
                        "fuel_label": fuel,
                        "source": "esto_aggregated",
                        "year": base_year,
                        "value": base_value,
                    }
                )
            else:
                # Base year
                base_key = (esto_flow.strip().lower(), esto_product.strip().lower())
                if base_key not in base_cache:
                    base_cache[base_key] = pull_base_year_value(
                        base_df,
                        base_year=base_year,
                        economy_code=base_economy,
                        esto_flow=esto_flow,
                        esto_product=esto_product,
                    )
                base_value = base_cache[base_key]
                if not preserve_sign and not pd.isna(base_value):
                    base_value = abs(float(base_value))
                long_rows.append(
                    {
                        "economy": base_economy,
                        "scenario": scenario_display,
                        "sheet": sheet,
                        "fuel_label": fuel,
                        "source": "base",
                        "year": base_year,
                        "value": base_value,
                    }
                )

                # Projection series
                proj_parts: list[pd.Series] = []
                if ninth_fuel:
                    for sector_code in sector_codes:
                        cache_key = (
                            str(sector_code or "").strip().lower(),
                            str(ninth_fuel or "").strip().lower(),
                            str(projection_scenario or "").strip().lower(),
                        )
                        if cache_key not in projection_cache:
                            projection_cache[cache_key] = pull_projection_series(
                                ninth_projection_df,
                                sector_code=sector_code,
                                fuel_code=ninth_fuel,
                                economy_code=projection_economy,
                                scenario=projection_scenario,
                                projection_years=projection_years,
                            ).reindex(projection_years)
                        proj_part = projection_cache[cache_key]
                        proj_parts.append(proj_part.reindex(projection_years))
                if proj_parts:
                    proj_series = pd.concat(proj_parts, axis=1).sum(axis=1, min_count=1)
                else:
                    proj_series = pd.Series(dtype="float64", index=projection_years)
                if not preserve_sign:
                    proj_series = proj_series.abs()
                for year, val in proj_series.items():
                    long_rows.append(
                        {
                            "economy": projection_economy,
                            "scenario": scenario_display,
                            "sheet": sheet,
                            "fuel_label": fuel,
                            "source": "projection",
                            "year": int(year),
                            "value": float(val) if not pd.isna(val) else float("nan"),
                        }
                    )

    status_df = pd.DataFrame(status_rows)
    if not status_df.empty:
        code_df = pd.read_excel(DEFAULT_CODEBOOK, sheet_name="code_to_name")
        code_df["esto_label_norm"] = code_df.get("esto_label", "").map(_clean_token).str.lower()
        code_df["esto_column_norm"] = code_df.get("esto_column", "").map(_clean_token).str.lower()
        flow_depth_lookup: dict[str, int] = {}
        for _, row in code_df.iterrows():
            if row["esto_column_norm"] != "flows":
                continue
            flow = str(row["esto_label_norm"] or "").strip()
            ninth_label = str(_clean_token(row.get("9th_label", "")) or "").strip()
            if flow and ninth_label:
                depth = len([p for p in ninth_label.split("_") if p])
                if flow not in flow_depth_lookup or depth < flow_depth_lookup[flow]:
                    flow_depth_lookup[flow] = depth
        status_df = status_df.copy()
        status_df["esto_flow_norm"] = status_df["esto_flow"].fillna("").astype(str).str.strip().str.lower()
        status_df["flow_source"] = status_df["flow_source"].fillna("").astype(str).str.strip().str.lower()
        status_df["sector_match_method"] = status_df["sector_match_method"].fillna("").astype(str).str.strip().str.lower()
        status_df["esto_flow_depth"] = status_df["esto_flow_norm"].map(flow_depth_lookup)
        status_df["sector_codes_list"] = status_df["sector_code_9th"].map(_split_sector_codes)
        status_df["min_sector_depth"] = status_df["sector_codes_list"].map(
            lambda xs: min((len([p for p in str(x).split("_") if p]) for x in xs), default=9999)
        )
        status_df["uses_parent_flow"] = (
            pd.to_numeric(status_df["esto_flow_depth"], errors="coerce").fillna(9999)
            < pd.to_numeric(status_df["min_sector_depth"], errors="coerce").fillna(9999)
        )
        status_df["allow_parent_estimate"] = (
            (status_df["flow_source"] == "sector_fallback")
            | status_df["sector_match_method"].str.startswith("code_ancestor_")
            | (status_df["sector_match_method"] == "independent_reverse")
        )
        strict_parent_flow = status_df[
            status_df["uses_parent_flow"]
            & status_df["esto_flow_norm"].ne("")
            & ~status_df["allow_parent_estimate"]
        ].copy()
        if not strict_parent_flow.empty:
            examples = (
                strict_parent_flow[
                    ["sheet", "fuel_label", "sector_code_9th", "esto_flow", "flow_source", "sector_match_method"]
                ]
                .drop_duplicates()
                .head(10)
                .to_dict("records")
            )
            raise RuntimeError(
                "Parent-flow canonical mappings detected. These rows map to a shallower ESTO flow and must be fixed "
                f"in the mapping files instead of being estimated. Total rows: {len(strict_parent_flow)}. "
                f"Examples: {examples}"
            )

    comparison_long = pd.DataFrame(long_rows)
    if not comparison_long.empty and str(sibling_comparator_mode or "").strip().lower() == "allocate_by_leap_share":
        if not status_df.empty:
            code_df = pd.read_excel(DEFAULT_CODEBOOK, sheet_name="code_to_name")
            code_df["esto_label_norm"] = code_df.get("esto_label", "").map(_clean_token).str.lower()
            code_df["esto_column_norm"] = code_df.get("esto_column", "").map(_clean_token).str.lower()
            code_df["name_clean"] = code_df.get("name", "").map(_clean_token)
            flow_name_lookup: dict[str, str] = {}
            flow_depth_lookup: dict[str, int] = {}
            for _, row in code_df.iterrows():
                if row["esto_column_norm"] != "flows":
                    continue
                flow = str(row["esto_label_norm"] or "").strip()
                name = str(row["name_clean"] or "").strip()
                if flow and name and flow not in flow_name_lookup:
                    flow_name_lookup[flow] = name
                ninth_label = str(_clean_token(row.get("9th_label", "")) or "").strip()
                if flow and ninth_label:
                    depth = len([p for p in ninth_label.split("_") if p])
                    if flow not in flow_depth_lookup or depth < flow_depth_lookup[flow]:
                        flow_depth_lookup[flow] = depth

            status_df = status_df.copy()
            status_df["sheet"] = status_df["sheet"].astype(str)
            status_df["fuel_label"] = status_df["fuel_label"].astype(str)
            status_df["esto_flow_norm"] = status_df["esto_flow"].fillna("").astype(str).str.strip().str.lower()
            status_df["flow_source"] = status_df["flow_source"].fillna("").astype(str).str.strip().str.lower()
            status_df["sector_match_method"] = status_df["sector_match_method"].fillna("").astype(str).str.strip().str.lower()
            status_df["effective_parent_name"] = status_df["esto_flow_norm"].map(flow_name_lookup).fillna("")
            status_df["esto_flow_depth"] = status_df["esto_flow_norm"].map(flow_depth_lookup)
            status_df["sector_codes_list"] = status_df["sector_code_9th"].map(_split_sector_codes)
            status_df["min_sector_depth"] = status_df["sector_codes_list"].map(
                lambda xs: min((len([p for p in str(x).split("_") if p]) for x in xs), default=9999)
            )
            status_df["uses_parent_flow"] = (
                pd.to_numeric(status_df["esto_flow_depth"], errors="coerce").fillna(9999)
                < pd.to_numeric(status_df["min_sector_depth"], errors="coerce").fillna(9999)
            )
            status_df["allow_parent_estimate"] = (
                (status_df["flow_source"] == "sector_fallback")
                | status_df["sector_match_method"].str.startswith("code_ancestor_")
                | (status_df["sector_match_method"] == "independent_reverse")
            )
            status_df = (
                status_df.sort_values(["sheet", "fuel_label"])
                .drop_duplicates(subset=["sheet", "fuel_label"], keep="first")
                [["sheet", "fuel_label", "esto_flow_norm", "flow_source", "effective_parent_name", "min_sector_depth", "uses_parent_flow", "allow_parent_estimate"]]
            )

            comp = comparison_long.copy()
            comp["sheet"] = comp["sheet"].astype(str)
            comp["fuel_label"] = comp["fuel_label"].astype(str)
            comp["value"] = pd.to_numeric(comp["value"], errors="coerce")
            comp = comp.merge(status_df, on=["sheet", "fuel_label"], how="left")
            comp["esto_flow_norm"] = comp["esto_flow_norm"].fillna("")
            comp["flow_source"] = comp["flow_source"].fillna("")
            comp["effective_parent_name"] = comp["effective_parent_name"].fillna("")
            comp["min_sector_depth"] = pd.to_numeric(comp["min_sector_depth"], errors="coerce").fillna(9999)
            comp["uses_parent_flow"] = comp["uses_parent_flow"].fillna(False).astype(bool)
            comp["allow_parent_estimate"] = comp["allow_parent_estimate"].fillna(False).astype(bool)

            fallback_mask = comp["uses_parent_flow"] & comp["allow_parent_estimate"] & (comp["effective_parent_name"] != "")
            group_cols = ["scenario", "fuel_label", "year", "effective_parent_name"]

            if fallback_mask.any():
                fallback_rows = comp[fallback_mask].copy()
                promoted_base_rows = pd.DataFrame()

                # Estimate ESTO base branch points using direct-detail shares.
                share_rows = pd.DataFrame()
                for direct_source, share_col in [("projection", "detail_share"), ("leap", "detail_share")]:
                    direct_rows = fallback_rows[fallback_rows["source"] == direct_source].copy()
                    if direct_rows.empty:
                        continue
                    min_depth = (
                        direct_rows.groupby(group_cols, as_index=False)["min_sector_depth"]
                        .min()
                        .rename(columns={"min_sector_depth": "group_min_depth"})
                    )
                    parent_rows = direct_rows.merge(min_depth, on=group_cols, how="left")
                    parent_rows = parent_rows[parent_rows["min_sector_depth"] == parent_rows["group_min_depth"]].copy()
                    totals = (
                        parent_rows.groupby(group_cols, as_index=False)["value"]
                        .sum(min_count=1)
                        .rename(columns={"value": "parent_total"})
                    )
                    shares = direct_rows.merge(totals, on=group_cols, how="left")
                    shares[share_col] = shares["value"] / shares["parent_total"]
                    shares.loc[
                        ~shares[share_col].replace([float("inf"), float("-inf")], pd.NA).notna(),
                        share_col,
                    ] = pd.NA
                    share_rows = pd.concat(
                        [
                            share_rows,
                            shares[
                                ["sheet"] + group_cols + ["esto_flow_norm", "flow_source", "min_sector_depth", "economy", share_col]
                            ].rename(columns={share_col: "detail_share"})
                        ],
                        ignore_index=True,
                        sort=False,
                    )
                if not share_rows.empty:
                    share_rows = share_rows.dropna(subset=["detail_share"]).drop_duplicates(
                        subset=["sheet"] + group_cols,
                        keep="first",
                    )
                    base_rows = fallback_rows[fallback_rows["source"] == "base"].copy()
                    if not base_rows.empty:
                        base_parent = (
                            base_rows.groupby(group_cols, as_index=False)["value"]
                            .agg(
                                first_non_null=lambda s: s.dropna().iloc[0] if len(s.dropna()) else float("nan"),
                                unique_non_null=lambda s: s.dropna().nunique(),
                            )
                        )
                        base_parent = base_parent[base_parent["unique_non_null"] <= 1].copy()
                        if not base_parent.empty:
                            base_econ = base_rows.groupby(group_cols, as_index=False)["economy"].first()
                            promoted_base_rows = base_parent.merge(base_econ, on=group_cols, how="left")
                            promoted_base_rows["sheet"] = promoted_base_rows["effective_parent_name"]
                            promoted_base_rows["source"] = "base"
                            promoted_base_rows["value"] = promoted_base_rows["first_non_null"]
                        est_rows = share_rows.merge(
                            base_parent[group_cols + ["first_non_null"]],
                            on=group_cols,
                            how="inner",
                        )
                        if not est_rows.empty:
                            est_rows["value"] = est_rows["detail_share"] * est_rows["first_non_null"]
                            est_rows["source"] = "base_estimated"
                            keep_cols = [
                                "economy",
                                "scenario",
                                "sheet",
                                "fuel_label",
                                "source",
                                "year",
                                "value",
                                "esto_flow_norm",
                                "flow_source",
                                "effective_parent_name",
                                "min_sector_depth",
                            ]
                            comp = comp[~((fallback_mask) & (comp["source"] == "base"))].copy()
                            comp = pd.concat([comp, est_rows[keep_cols]], ignore_index=True, sort=False)

                # Promote true parent-level charts named by the effective parent category (e.g. Road).
                if include_sibling_parent_totals:
                    promoted = comp[(comp["effective_parent_name"] != "") & comp["source"].isin(["leap", "projection"])].copy()
                    if not promoted.empty:
                        min_depth = promoted.groupby(group_cols, as_index=False)["min_sector_depth"].min().rename(columns={"min_sector_depth": "group_min_depth"})
                        promoted = promoted.merge(min_depth, on=group_cols, how="left")
                        promoted = promoted[promoted["min_sector_depth"] == promoted["group_min_depth"]].copy()
                        promoted = (
                            promoted.groupby(group_cols + ["source"], as_index=False)["value"]
                            .sum(min_count=1)
                        )
                        # preserve economy code per source where possible
                        econ_ref = (
                            comp[(comp["effective_parent_name"] != "") & comp["source"].isin(["leap", "projection"])]
                            .groupby(group_cols + ["source"], as_index=False)["economy"]
                            .first()
                        )
                        promoted = promoted.merge(econ_ref, on=group_cols + ["source"], how="left")
                        promoted["sheet"] = promoted["effective_parent_name"]
                        keep_cols = ["economy", "scenario", "sheet", "fuel_label", "source", "year", "value"]
                        if not promoted_base_rows.empty:
                            promoted = pd.concat(
                                [promoted[keep_cols], promoted_base_rows[keep_cols]],
                                ignore_index=True,
                                sort=False,
                            )
                        else:
                            promoted = promoted[keep_cols]
                        comp = pd.concat([comp, promoted], ignore_index=True, sort=False)

            comparison_long = comp.drop(
                columns=["esto_flow_norm", "flow_source", "effective_parent_name", "min_sector_depth"],
                errors="ignore",
            )

    if comparison_long.empty:
        comparison_wide = pd.DataFrame()
    else:
        comparison_wide = (
            comparison_long.pivot_table(
                index=["economy", "scenario", "sheet", "fuel_label", "year"],
                columns="source",
                values="value",
                aggfunc="first",
            )
            .reset_index()
        )
    mapping_status = pd.DataFrame(status_rows)
    if use_esto_agg_only and not mapping_status.empty:
        mapping_status["projection_available"] = False
    return comparison_long, comparison_wide, mapping_status


# -----------------------------------------------------------------------------
# Charting and dashboards
# -----------------------------------------------------------------------------
def _safe_token(value: object) -> str:
    text = str(value).strip() if value is not None else ""
    if not text:
        return "item"
    safe = "".join(ch if ch.isalnum() or ch in {"_", "-"} else "_" for ch in text)
    return safe.strip("_") or "item"


def make_chart(
    sheet: str,
    fuel: str,
    subset: pd.DataFrame,
    output_dir: Path,
    backend: str = "plotly",
    display_sheet: str | None = None,
    file_sheet: str | None = None,
) -> Path | None:
    """Generate a simple comparison chart for one sheet/fuel."""
    output_dir.mkdir(parents=True, exist_ok=True)
    display_name = display_sheet or sheet
    sheet_slug = _safe_token((file_sheet or sheet).replace("\\", "_"))
    fuel_slug = _safe_token(fuel)
    out_png = output_dir / f"{sheet_slug}__{fuel_slug}.png"
    out_html = output_dir / f"{sheet_slug}__{fuel_slug}.html"

    def _format_scenario_label(value: object) -> str:
        raw = str(value or "").strip()
        if not raw:
            return "Scenario"
        lowered = raw.lower()
        if lowered == "reference":
            return "Reference"
        if lowered == "target":
            return "Target"
        return raw

    def _series_by_scenario(source_name: str) -> dict[str, pd.Series]:
        out: dict[str, pd.Series] = {}
        src = subset[subset["source"] == source_name].copy()
        if src.empty:
            return out
        for scenario, g in src.groupby("scenario", dropna=False):
            scen_label = _format_scenario_label(scenario)
            s = (
                g.sort_values("year")
                .groupby("year", as_index=True)["value"]
                .first()
            )
            out[scen_label] = pd.to_numeric(s, errors="coerce")
        return out

    leap_by_scenario = _series_by_scenario("leap")
    base_by_scenario = _series_by_scenario("base")
    base_est_by_scenario = _series_by_scenario("base_estimated")
    proj_by_scenario = _series_by_scenario("projection")
    proj_est_by_scenario = _series_by_scenario("projection_estimated")

    def _render_static() -> Path | None:
        try:
            import matplotlib.pyplot as plt

            plt.figure(figsize=(7, 4))
            for scen, leap_s in sorted(leap_by_scenario.items()):
                if not leap_s.empty:
                    label = "LEAP" if len(leap_by_scenario) == 1 else f"LEAP ({scen})"
                    plt.plot(leap_s.index, leap_s.values, label=label, marker="o")
            for scen, proj_s in sorted(proj_by_scenario.items()):
                if not proj_s.empty:
                    label = "Projection" if len(proj_by_scenario) == 1 else f"Projection ({scen})"
                    plt.plot(proj_s.index, proj_s.values, label=label, marker="o")
            for scen, proj_s in sorted(proj_est_by_scenario.items()):
                if not proj_s.empty:
                    label = "Projection (estimated)" if len(proj_est_by_scenario) == 1 else f"Projection (estimated, {scen})"
                    plt.plot(proj_s.index, proj_s.values, label=label, marker="x", linestyle="--")
            for scen, base_s in sorted(base_by_scenario.items()):
                if not base_s.empty:
                    label = "Base (2022)" if len(base_by_scenario) == 1 else f"Base (2022, {scen})"
                    plt.scatter(base_s.index, base_s.values, label=label, marker="D")
            for scen, base_s in sorted(base_est_by_scenario.items()):
                if not base_s.empty:
                    label = "Base (estimated)" if len(base_est_by_scenario) == 1 else f"Base (estimated, {scen})"
                    plt.scatter(base_s.index, base_s.values, label=label, marker="X")
            plt.title(f"{display_name} – {fuel}")
            plt.xlabel("Year")
            plt.ylabel("Energy")
            plt.legend()
            plt.tight_layout()
            plt.savefig(out_png, dpi=150)
            plt.close()
            return out_png
        except Exception as exc:  # noqa: BLE001
            print(f"[WARN] Failed to render static chart for {display_name}/{fuel}: {exc}")
            return None

    try:
        if backend == "plotly":
            import plotly.graph_objects as go

            fig = go.Figure()
            for scen, leap_s in sorted(leap_by_scenario.items()):
                if leap_s.empty:
                    continue
                name = "LEAP" if len(leap_by_scenario) == 1 else f"LEAP ({scen})"
                fig.add_trace(go.Scatter(x=leap_s.index, y=leap_s.values, mode="lines+markers", name=name))
            for scen, proj_s in sorted(proj_by_scenario.items()):
                if proj_s.empty:
                    continue
                name = "Projection" if len(proj_by_scenario) == 1 else f"Projection ({scen})"
                fig.add_trace(go.Scatter(x=proj_s.index, y=proj_s.values, mode="lines+markers", name=name))
            for scen, proj_s in sorted(proj_est_by_scenario.items()):
                if proj_s.empty:
                    continue
                name = "Projection (estimated)" if len(proj_est_by_scenario) == 1 else f"Projection (estimated, {scen})"
                fig.add_trace(
                    go.Scatter(x=proj_s.index, y=proj_s.values, mode="lines+markers", line=dict(dash="dash"), marker=dict(symbol="x"), name=name)
                )
            for scen, base_s in sorted(base_by_scenario.items()):
                if base_s.empty:
                    continue
                name = "Base (2022)" if len(base_by_scenario) == 1 else f"Base (2022, {scen})"
                fig.add_trace(
                    go.Scatter(
                        x=base_s.index,
                        y=base_s.values,
                        mode="markers",
                        marker=dict(size=10, symbol="diamond"),
                        name=name,
                    )
                )
            for scen, base_s in sorted(base_est_by_scenario.items()):
                if base_s.empty:
                    continue
                name = "Base (estimated)" if len(base_est_by_scenario) == 1 else f"Base (estimated, {scen})"
                fig.add_trace(
                    go.Scatter(
                        x=base_s.index,
                        y=base_s.values,
                        mode="markers",
                        marker=dict(size=10, symbol="x"),
                        name=name,
                    )
                )
            fig.update_layout(
                title=f"{display_name} – {fuel}",
                xaxis_title="Year",
                yaxis_title="Energy",
                template="plotly_white",
            )
            fig.write_html(out_html, include_plotlyjs="cdn", full_html=True)
            return out_html
        return _render_static()
    except Exception as exc:  # noqa: BLE001
        if backend == "plotly":
            print(f"[WARN] Plotly chart failed for {display_name}/{fuel}; falling back to static: {exc}")
            return _render_static()
        print(f"[WARN] Failed to render chart for {display_name}/{fuel}: {exc}")
        return None


def _append_total_rows(comparison_long: pd.DataFrame) -> pd.DataFrame:
    if comparison_long.empty:
        return comparison_long.copy()

    base = comparison_long.copy()
    base["value"] = pd.to_numeric(base["value"], errors="coerce")
    base_no_total = base[base["fuel_label"].astype(str) != "Total"].copy()
    if base_no_total.empty:
        return base

    totals = (
        base_no_total.groupby(["sheet", "scenario", "source", "year"], dropna=False)
        .agg(value=("value", lambda s: s.sum(min_count=1)))
        .reset_index()
    )
    totals["fuel_label"] = "Total"

    return pd.concat([base, totals], ignore_index=True, sort=False)


def build_charts(
    comparison_long: pd.DataFrame,
    charts_dir: Path,
    backend: str = "plotly",
) -> list[Path]:
    written: list[Path] = []
    if comparison_long.empty:
        return written
    render_long = _append_total_rows(comparison_long)
    for (sheet, fuel), sub in render_long.groupby(["sheet", "fuel_label"]):
        out = make_chart(sheet, fuel, sub, charts_dir, backend=backend)
        if out:
            written.append(out)
    return written


def build_dashboards(
    output_dir: Path,
    comparison_long: pd.DataFrame,
    charts_dir: Path,
    mapping_status: pd.DataFrame | None = None,
) -> Path | None:
    """
    Reuse the lightweight dashboard style from leap_transport (_build_sheet_dashboards).
    """
    if comparison_long.empty or not charts_dir.exists():
        print("[INFO] No charts available for dashboard rendering.")
        return None

    render_long = _append_total_rows(comparison_long)

    dashboards_dir = output_dir / "dashboards"
    dashboards_dir.mkdir(parents=True, exist_ok=True)

    # Flag sheet/fuel pairs with a meaningful base-year mismatch so they stand out in the UI.
    base_issue_lookup: dict[tuple[str, str], dict[str, float | str]] = {}
    base_probe_year = 2022
    base_diag = comparison_long.copy()
    base_diag["value"] = pd.to_numeric(base_diag["value"], errors="coerce")
    base_diag = base_diag[
        (pd.to_numeric(base_diag["year"], errors="coerce") == base_probe_year)
        & (base_diag["source"].isin(["leap", "base", "base_estimated"]))
    ].copy()
    if not base_diag.empty:
        base_diag["base_compare"] = base_diag["source"].replace({"base_estimated": "base"})
        base_wide = (
            base_diag.pivot_table(
                index=["sheet", "fuel_label"],
                columns="base_compare",
                values="value",
                aggfunc="first",
            )
            .reset_index()
        )
        if "leap" not in base_wide.columns:
            base_wide["leap"] = pd.NA
        if "base" not in base_wide.columns:
            base_wide["base"] = pd.NA
        base_wide["abs_gap"] = (pd.to_numeric(base_wide["leap"], errors="coerce") - pd.to_numeric(base_wide["base"], errors="coerce")).abs()
        base_wide["magnitude"] = (
            pd.concat(
                [
                    pd.to_numeric(base_wide["leap"], errors="coerce").abs(),
                    pd.to_numeric(base_wide["base"], errors="coerce").abs(),
                ],
                axis=1,
            )
            .max(axis=1)
        )
        denom = pd.to_numeric(base_wide["leap"], errors="coerce").abs().where(lambda s: s > 1e-9)
        alt = pd.to_numeric(base_wide["base"], errors="coerce").abs().where(lambda s: s > 1e-9)
        denom = denom.fillna(alt).fillna(1.0)
        base_wide["gap_ratio"] = base_wide["abs_gap"] / denom
        magnitude_p90 = float(base_wide["magnitude"].dropna().quantile(0.9)) if base_wide["magnitude"].notna().any() else 0.0
        magnitude_p50 = float(base_wide["magnitude"].dropna().quantile(0.5)) if base_wide["magnitude"].notna().any() else 0.0
        sig = base_wide[
            pd.to_numeric(base_wide["leap"], errors="coerce").notna()
            & pd.to_numeric(base_wide["base"], errors="coerce").notna()
            & (base_wide["gap_ratio"] >= 0.10)
        ].copy()
        for _, row in sig.iterrows():
            pct = float(row["gap_ratio"]) * 100.0
            if pct >= 200.0:
                severity = "Extreme"
            elif pct >= 50.0:
                severity = "High"
            else:
                severity = "Moderate"
            magnitude = float(row["magnitude"]) if pd.notna(row["magnitude"]) else 0.0
            if magnitude_p90 > 0 and magnitude >= magnitude_p90:
                impact = "major"
            elif magnitude_p50 > 0 and magnitude >= magnitude_p50:
                impact = "medium"
            else:
                impact = "minor"
            base_issue_lookup[(str(row["sheet"]), str(row["fuel_label"]))] = {
                "pct": pct,
                "severity": severity,
                "impact": impact,
                "label": f"Base-year gap: {severity}",
            }

    magnitude_lookup = (
        render_long.assign(value_abs=pd.to_numeric(render_long["value"], errors="coerce").abs())
        .groupby(["sheet", "fuel_label"], dropna=False)["value_abs"]
        .max()
        .to_dict()
    )

    code_df = pd.read_excel(DEFAULT_CODEBOOK, sheet_name="code_to_name")
    code_df["9th_label_clean"] = code_df.get("9th_label", "").map(_clean_token)
    code_df["9th_column_clean"] = code_df.get("9th_column", "").map(_clean_token).str.lower()
    code_df["name_clean"] = code_df.get("name", "").map(_clean_token)
    sector_rows = code_df[code_df["9th_column_clean"].isin({"sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"})].copy()

    def _numeric_seq(code: str) -> tuple[int, ...]:
        parts: list[int] = []
        for token in str(code or "").strip().split("_"):
            if token.isdigit():
                parts.append(int(token))
            else:
                break
        return tuple(parts)

    sector_rows["num_seq"] = sector_rows["9th_label_clean"].map(_numeric_seq)
    seq_to_name = {
        tuple(row["num_seq"]): str(row["name_clean"]).strip()
        for _, row in sector_rows.iterrows()
        if tuple(row["num_seq"]) and str(row["name_clean"]).strip()
    }
    name_to_seq = {
        str(row["name_clean"]).strip().lower(): tuple(row["num_seq"])
        for _, row in sector_rows.iterrows()
        if str(row["name_clean"]).strip() and tuple(row["num_seq"])
    }

    status_by_sheet: dict[str, list[str]] = {}
    if mapping_status is not None and not mapping_status.empty:
        ms = mapping_status.copy()
        ms["sheet"] = ms["sheet"].astype(str)
        for sheet, g in ms.groupby("sheet", dropna=False):
            codes: list[str] = []
            for raw in g["sector_code_9th"].dropna().astype(str):
                codes.extend(_split_sector_codes(raw))
            if codes:
                status_by_sheet[str(sheet)] = codes

    def _common_prefix(seqs: list[tuple[int, ...]]) -> tuple[int, ...]:
        if not seqs:
            return ()
        prefix = list(seqs[0])
        for seq in seqs[1:]:
            keep = 0
            for a, b in zip(prefix, seq):
                if a != b:
                    break
                keep += 1
            prefix = prefix[:keep]
            if not prefix:
                break
        return tuple(prefix)

    def _sheet_path(sheet: str) -> list[str]:
        codes = status_by_sheet.get(sheet, [])
        seqs = [_numeric_seq(code) for code in codes if _numeric_seq(code)]
        if seqs:
            seq = _common_prefix(seqs)
        else:
            seq = name_to_seq.get(str(sheet).strip().lower(), ())
        if not seq:
            return [sheet]
        names: list[str] = []
        for i in range(1, len(seq) + 1):
            prefix = seq[:i]
            name = seq_to_name.get(prefix)
            if name:
                names.append(name)
        if not names:
            return [sheet]
        if names[-1] != sheet:
            names.append(sheet)
        return names

    sheet_paths: dict[str, list[str]] = {}
    for sheet in render_long["sheet"].astype(str).drop_duplicates():
        sheet_paths[sheet] = _sheet_path(sheet)

    def _sheet_sort_key(sheet: str) -> tuple:
        path = sheet_paths.get(sheet, [sheet])
        return tuple([p.lower() for p in path] + [sheet.lower()])

    def _menu_label(sheet: str) -> str:
        path = sheet_paths.get(sheet, [sheet])
        depth = max(0, len(path) - 1)
        return f"{'  ' * depth}{sheet}"

    sheet_to_entries: dict[str, list[tuple[str, Path]]] = {}
    chart_files: dict[str, Path] = {}
    for p in charts_dir.glob("*"):
        if p.suffix.lower() not in {".png", ".svg", ".html"}:
            continue
        stem = p.stem
        existing = chart_files.get(stem)
        if existing is None or p.stat().st_mtime > existing.stat().st_mtime:
            chart_files[stem] = p
    for (sheet, fuel), _ in render_long.groupby(["sheet", "fuel_label"]):
        sheet_slug = _safe_token(sheet.replace("\\", "_"))
        fuel_slug = _safe_token(fuel)
        stem = f"{sheet_slug}__{fuel_slug}"
        if stem in chart_files:
            sheet_to_entries.setdefault(sheet, []).append((fuel, chart_files[stem]))

    if not sheet_to_entries:
        print("[INFO] No matching chart files were found for dashboards.")
        return None

    ordered_sheets = sorted(sheet_to_entries.items(), key=lambda item: _sheet_sort_key(item[0]))
    sheet_entries_lookup = {sheet: entries for sheet, entries in ordered_sheets}
    sheet_path_tuples = {sheet: tuple(sheet_paths.get(sheet, [sheet])) for sheet in sheet_to_entries}

    node_paths: set[tuple[str, ...]] = set()
    for path in sheet_path_tuples.values():
        for i in range(1, len(path) + 1):
            node_paths.add(path[:i])

    def _path_sort_key(path: tuple[str, ...]) -> tuple[str, ...]:
        return tuple(part.lower() for part in path)

    node_to_sheet: dict[tuple[str, ...], str] = {path: sheet for sheet, path in sheet_path_tuples.items()}

    def _dashboard_filename(path: tuple[str, ...]) -> str:
        if path in node_to_sheet:
            sheet = node_to_sheet[path]
            return f"{_safe_token(sheet.replace('\\', '_'))}.html"
        slug = "__".join(_safe_token(part.replace("\\", "_")) for part in path)
        return f"node__{slug}.html"

    path_to_file = {path: dashboards_dir / _dashboard_filename(path) for path in sorted(node_paths, key=_path_sort_key)}

    def _descendant_sheets(path: tuple[str, ...]) -> list[str]:
        out = [sheet for sheet, sheet_path in sheet_path_tuples.items() if sheet_path[: len(path)] == path]
        return sorted(out, key=_sheet_sort_key)

    def _leaf_descendant_sheets(path: tuple[str, ...]) -> list[str]:
        descendants = _descendant_sheets(path)
        leafs: list[str] = []
        for sheet in descendants:
            sheet_path = sheet_path_tuples[sheet]
            has_child = any(
                other != sheet and sheet_path == sheet_path_tuples[other][: len(sheet_path)]
                for other in descendants
            )
            if not has_child:
                leafs.append(sheet)
        return sorted(leafs, key=_sheet_sort_key)

    def _render_cards(section_sheet: str, entries: list[tuple[str, Path]]) -> str:
        if not entries:
            return ""
        entries = sorted(
            entries,
            key=lambda item: (
                0 if str(item[0]) == "Total" else 1,
                -float(magnitude_lookup.get((section_sheet, item[0]), 0.0) or 0.0),
                str(item[0]).lower(),
            ),
        )
        cols = 3 if len(entries) >= 3 else max(1, len(entries))
        cards = []
        for fuel, png_path in entries:
            issue = base_issue_lookup.get((section_sheet, fuel))
            rel_chart = os.path.relpath(png_path, start=dashboards_dir).replace("\\", "/")
            if png_path.suffix.lower() == ".html":
                chart_markup = (
                    f'<iframe src="{rel_chart}" '
                    f'title="{section_sheet} – {fuel}" '
                    'style="width:100%;height:420px;border:1px solid #d0d7de;background:#fff;" loading="lazy"></iframe>'
                )
            else:
                chart_markup = (
                    f'<img src="{rel_chart}" alt="{section_sheet} – {fuel}" '
                    'style="max-width:100%;height:auto;" loading="lazy" />'
                )
            card_style = "margin:8px;padding:8px;border:1px solid #d0d7de;border-radius:8px;background:#fff;box-shadow:0 1px 2px rgba(0,0,0,0.05);"
            issue_badge = ""
            if issue:
                severity = str(issue.get("severity", "Moderate"))
                impact = str(issue.get("impact", "minor"))
                palette = {
                    "Moderate": {"border": "rgba(217,119,6,{a})", "bg": "rgba(245,158,11,{b})", "text": "#92400e"},
                    "High": {"border": "rgba(234,88,12,{a})", "bg": "rgba(249,115,22,{b})", "text": "#9a3412"},
                    "Extreme": {"border": "rgba(220,38,38,{a})", "bg": "rgba(220,38,38,{b})", "text": "#991b1b"},
                }.get(severity, {"border": "rgba(217,119,6,{a})", "bg": "rgba(245,158,11,{b})", "text": "#92400e"})
                alpha = {"minor": "0.30", "medium": "0.42", "major": "0.58"}.get(impact, "0.30")
                bg_alpha = {"minor": "0.04", "medium": "0.07", "major": "0.11"}.get(impact, "0.04")
                card_style = (
                    "margin:8px;padding:8px;border:2px solid "
                    + palette["border"].format(a=alpha)
                    + ";border-radius:8px;background:"
                    + palette["bg"].format(b=bg_alpha)
                    + ";box-shadow:0 1px 2px rgba(0,0,0,0.05);"
                )
                issue_badge = (
                    f'<div style="margin-top:4px;color:{palette["text"]};font-size:12px;font-weight:600;">'
                    f'{issue["label"]}</div>'
                )
            cards.append(
                f"""
<figure style="{card_style}">
  <figcaption style="font-weight:600;margin-bottom:4px;">{fuel}</figcaption>
  {issue_badge}
  {chart_markup}
</figure>
"""
            )
        return f'<div class="grid" style="display:grid;gap:12px;grid-template-columns:repeat({cols}, minmax(220px, 1fr));">{"".join(cards)}</div>'

    ordered_node_paths = sorted(node_paths, key=_path_sort_key)
    page_chart_counts: dict[tuple[str, ...], int] = {}
    for path in ordered_node_paths:
        desc = _descendant_sheets(path)
        page_chart_counts[path] = sum(len(sheet_entries_lookup.get(sheet, [])) for sheet in desc)

    nav_groups: dict[str, list[str]] = {}
    nav_group_order: list[str] = []
    for path in ordered_node_paths:
        top_group = path[0] if path else "Other"
        if top_group not in nav_groups:
            nav_groups[top_group] = []
            nav_group_order.append(top_group)
        selected_placeholder = "__SELECTED__"
        depth = max(0, len(path) - 1)
        label = f"{'  ' * depth}{path[-1]}"
        nav_groups[top_group].append((path, label, selected_placeholder))

    page_files: list[tuple[tuple[str, ...], Path, int]] = []
    for path in ordered_node_paths:
        dashboard_file = path_to_file[path]
        title = path[-1]
        desc_sheets = _descendant_sheets(path)
        own_sheet = node_to_sheet.get(path)

        sections: list[str] = []
        if not own_sheet:
            leaf_sheets = _leaf_descendant_sheets(path)
            if leaf_sheets:
                node_subset = render_long[
                    render_long["sheet"].astype(str).isin(leaf_sheets)
                    & (render_long["fuel_label"].astype(str) != "Total")
                ].copy()
                if not node_subset.empty:
                    node_total = (
                        node_subset.groupby(["scenario", "source", "year"], dropna=False)
                        .agg(value=("value", lambda s: pd.to_numeric(s, errors="coerce").sum(min_count=1)))
                        .reset_index()
                    )
                    node_total["sheet"] = title
                    node_total["fuel_label"] = "Total"
                    node_chart = make_chart(
                        title,
                        "Total",
                        node_total,
                        charts_dir,
                        backend="plotly",
                        display_sheet=title,
                        file_sheet=f"node__{'__'.join(_safe_token(part.replace('\\', '_')) for part in path)}",
                    )
                    if node_chart:
                        sections.append(
                            f'<section><h2 style="margin:18px 0 8px 0;">{title} total</h2>'
                            '<p style="margin:0 0 10px 0;color:#4b5563;font-size:13px;">'
                            'Aggregated from non-overlapping leaf descendant categories.'
                            '</p>'
                            f'{_render_cards(title, [("Total", node_chart)])}</section>'
                        )
        if own_sheet and own_sheet in sheet_entries_lookup:
            own_cards = _render_cards(own_sheet, sheet_entries_lookup[own_sheet])
            if own_cards:
                sections.append(f'<section><h2 style="margin:18px 0 8px 0;">{own_sheet}</h2>{own_cards}</section>')
        for child_sheet in desc_sheets:
            if child_sheet == own_sheet:
                continue
            cards = _render_cards(child_sheet, sheet_entries_lookup.get(child_sheet, []))
            if not cards:
                continue
            child_path = sheet_path_tuples[child_sheet]
            child_file = path_to_file.get(child_path)
            heading = f'<a href="{child_file.name}">{child_sheet}</a>' if child_file else child_sheet
            sections.append(f'<section><h2 style="margin:22px 0 8px 0;">{heading}</h2>{cards}</section>')
        body_content = "".join(sections) if sections else '<p>No charts available for this category.</p>'

        nav_options = []
        for top_group in nav_group_order:
            opts = []
            for other_path, label, _ in nav_groups[top_group]:
                selected = " selected" if other_path == path else ""
                opts.append(f'<option value="{path_to_file[other_path].name}"{selected}>{label}</option>')
            nav_options.append(f'<optgroup label="{top_group}">{"".join(opts)}</optgroup>')

        breadcrumb_parts = []
        for i in range(1, len(path) + 1):
            prefix = path[:i]
            file_path = path_to_file.get(prefix)
            name = prefix[-1]
            if file_path:
                breadcrumb_parts.append(f'<a href="{file_path.name}">{name}</a>')
            else:
                breadcrumb_parts.append(name)
        breadcrumb = " > ".join(breadcrumb_parts)

        html_doc = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>{title} – LEAP Results Dashboard</title>
  <style>
    body {{ font-family: Segoe UI, Arial, sans-serif; margin: 24px; background: #f4f6f8; color: #111; }}
    h1 {{ margin: 0; }}
    a {{ color: #0b3d5c; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
    .page-header {{
      position: sticky;
      top: 0;
      z-index: 100;
      margin: -24px -24px 18px -24px;
      padding: 16px 24px 12px 24px;
      background: rgba(244, 246, 248, 0.96);
      backdrop-filter: blur(8px);
      border-bottom: 1px solid #d8dee4;
      box-shadow: 0 1px 0 rgba(0, 0, 0, 0.04);
    }}
  </style>
</head>
<body>
  <div class="page-header">
    <h1>{title}</h1>
    <div style="margin:6px 0 12px 0;color:#4b5563;">{breadcrumb}</div>
    <div>
      <label for="dashboard-picker" style="font-weight:600;margin-right:8px;">View dashboard:</label>
      <select id="dashboard-picker" onchange="if (this.value) window.location.href=this.value;" style="padding:6px 10px;border:1px solid #c5ccd3;border-radius:6px;background:#fff;">
        <option value="index.html">Index</option>
        {''.join(nav_options)}
      </select>
    </div>
  </div>
  {body_content}
</body>
</html>
"""
        dashboard_file.write_text(html_doc, encoding="utf-8")
        page_files.append((path, dashboard_file, page_chart_counts.get(path, 0)))

    tree: dict[str, dict] = {}
    for path in ordered_node_paths:
        node = tree
        for part in path:
            node = node.setdefault(part, {})

    def _render_tree(node: dict, depth: int = 0, path_prefix: tuple[str, ...] = ()) -> str:
        items: list[str] = []
        for name in sorted(node.keys(), key=lambda item: item.lower()):
            child = node[name]
            current_path = path_prefix + (name,)
            file_path = path_to_file.get(current_path)
            count = page_chart_counts.get(current_path, 0)
            label_html = f'<a href="{file_path.name}">{name}</a> <span style="color:#4b5563;">({count} charts)</span>' if file_path else name
            child_html = _render_tree(child, depth + 1, current_path)
            section_style = (
                "margin:10px 0 6px 0;padding:8px 10px;border-left:3px solid #c5ccd3;background:#fff;"
                if depth == 0
                else "margin:6px 0 6px 14px;padding-left:10px;border-left:1px solid #d8dee4;"
            )
            items.append(f'<li style="{section_style}">{label_html}{child_html}</li>')
        if not items:
            return ""
        return f'<ul style="list-style:none;margin:{6 if depth else 0}px 0 0 0;padding:0;">{"".join(items)}</ul>'

    links_html = _render_tree(tree)
    issue_rows = []
    severity_rank = {"Extreme": 3, "High": 2, "Moderate": 1}
    impact_rank = {"major": 3, "medium": 2, "minor": 1}
    for (sheet, fuel), issue in sorted(
        base_issue_lookup.items(),
        key=lambda item: (
            -severity_rank.get(str(item[1].get("severity", "")), 0),
            -impact_rank.get(str(item[1].get("impact", "")), 0),
            -float(item[1]["pct"]),
            item[0][0].lower(),
            item[0][1].lower(),
        ),
    ):
        sheet_path = sheet_path_tuples.get(sheet)
        file_path = path_to_file.get(sheet_path) if sheet_path else None
        sheet_link = f'<a href="{file_path.name}">{sheet}</a>' if file_path else sheet
        issue_rows.append(
            f'<li><span style="color:#b91c1c;font-weight:600;">{issue["label"]} ({str(issue.get("impact", "minor")).capitalize()} impact)</span>: {sheet_link} – {fuel}</li>'
        )
    issues_html = (
        '<section style="margin:18px 0 24px 0;padding:12px 14px;border:1px solid rgba(220,38,38,0.25);'
        'background:rgba(220,38,38,0.04);border-radius:10px;">'
        '<h2 style="margin:0 0 8px 0;color:#991b1b;font-size:20px;">Significant Base-Year Differences</h2>'
        f'<ul style="margin:0;padding-left:20px;line-height:1.6;">{"".join(issue_rows[:100])}</ul>'
        f'{"<p style=\"margin:8px 0 0 0;color:#7f1d1d;\">Showing first 100 issues.</p>" if len(issue_rows) > 100 else ""}'
        '</section>'
        if issue_rows
        else ""
    )
    index_file = dashboards_dir / "index.html"
    index_file.write_text(
        f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>LEAP Results Dashboards</title>
  <style>
    body {{ font-family: Segoe UI, Arial, sans-serif; margin: 24px; background: #f4f6f8; color: #111; }}
    h1 {{ font-size: 32px; margin-top: 0; }}
    a {{ color: #0b3d5c; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
  </style>
</head>
<body>
  <h1>LEAP Results Dashboards</h1>
  <p>{len(page_files)} dashboards generated.</p>
  {issues_html}
  {links_html}
</body>
</html>
""",
        encoding="utf-8",
    )
    print(f"[INFO] Generated dashboards: {index_file}")
    return index_file


# -----------------------------------------------------------------------------
# Diagnostics / checks
# -----------------------------------------------------------------------------
def basic_checks(
    sheet_map: pd.DataFrame,
    fuel_mapping: dict[str, dict[str, str]],
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    *,
    fuel_coverage_threshold: float = 0.95,
) -> dict[str, object]:
    """Return summary diagnostics."""
    coverage = 0.0
    if not mapping_status.empty:
        coverage = mapping_status["mapped"].mean()

    has_all_sheets = sheet_map["sheet_name"].isin(comparison_long["sheet"].unique()).mean() == 1.0 if not comparison_long.empty else False

    issues = []
    if coverage < fuel_coverage_threshold:
        issues.append(f"Fuel mapping coverage {coverage:.2%} below target {fuel_coverage_threshold:.0%}.")
    if not has_all_sheets:
        issues.append("Some mapped sheets missing in comparison output.")

    return {
        "fuel_mapping_coverage": coverage,
        "all_sheets_present": has_all_sheets,
        "issues": issues,
    }


#%%
