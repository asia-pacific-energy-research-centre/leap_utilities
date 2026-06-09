"""
check_missing_branch_esto.py

Validates whether missing LEAP branch paths have supporting data in ESTO 2024/2025.

Run from repo root:
    python scripts/check_missing_branch_esto.py
"""

import sys
import os

# Ensure repo root is on path
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO_ROOT)

import pandas as pd

# ---------------------------------------------------------------------------
# Sector → ESTO flow code hints (substring match against flows column)
# ---------------------------------------------------------------------------
SECTOR_TO_FLOW_HINTS = {
    'NG Liquefaction': ['09.06.02'],
    'LNG regasification': ['09.06.02'],
    'Oil Refining': ['09.07'],
    'Gas works plants': ['09.06.01'],
    'BKB and PB plants': ['09.08.04'],
    'Blast furnaces': ['09.08.02'],
    'Hydrogen transformation': ['09.13'],
    'Non specified transformation': ['09.12'],
    'Gas to liquids plants': ['09.06.04'],
    'Petrochemical industry': ['09.09'],
    'Patent fuel plants': ['09.08.03'],
    'Refinery and blending transfers': ['08.01', '08.02'],
    'Transfers unallocated': ['08.99'],
    'Transfers': ['08.01', '08.02', '08.99'],
    'Natural gas blending plants': ['09.06.03'],
    'Coke ovens': ['09.08.01'],
    'Coal liquefaction': ['09.08.05'],
    'Charcoal production': ['09.11'],
    'Biofuels processing': ['09.10'],
    'Electricity Generation': ['09.01', '09.02'],
    'Transmission and Distribution': ['10.02'],
}

LOSS_FLOW_PREFIX = '10.'

# ESTO aggregate/subtotal fuel labels that should never be LEAP branch fuels.
# These are group-level products (e.g. "08 Gas" = sum of Natural gas + LNG + ...).
# Any branch whose leaf fuel label matches one of these is excluded from the
# required-for-LEAP-export list.
ESTO_SUBTOTAL_FUEL_LABELS = {
    'gas', 'coal', 'oil products', 'petroleum products',
    'crude oil & ngl', 'solid biomass', 'others', 'electricity', 'heat',
    'nuclear', 'hydro', 'geothermal', 'solar', 'wind',
    'peat', 'peat products', 'coal products',
    'oil shale and oil sands', 'tide, wave, ocean',
}

# ---------------------------------------------------------------------------
# Fuel name → ESTO product substring hints
# Many LEAP fuel labels don't exactly match ESTO product codes; we map common
# ones here.  The fallback is case-insensitive substring matching.
# ---------------------------------------------------------------------------
FUEL_PRODUCT_HINTS = {
    'natural gas': ['natural gas', '08.01'],
    'lng': ['lng', '08.02'],
    'gas': ['08 gas', '08.01', '08.02', '08.03', '08.99'],
    'crude oil': ['crude oil', '06.01'],
    'ngl': ['06.02', 'natural gas liquids'],
    'lpg': ['07.09', 'lpg'],
    'motor gasoline': ['07.01', 'motor gasoline'],
    'gas and diesel oil': ['07.07', 'gas/diesel'],
    'diesel': ['07.07', 'gas/diesel'],
    'naphtha': ['07.03', 'naphtha'],
    'fuel oil': ['07.08', 'fuel oil'],
    'kerosene': ['07.06', 'kerosene'],
    'kerosene type jet fuel': ['07.05', 'kerosene type jet'],
    'aviation gasoline': ['07.02', 'aviation gasoline'],
    'refinery feedstocks': ['06.03', 'refinery feedstock'],
    'refinery gas not liquefied': ['07.10', 'refinery gas'],
    'petroleum coke': ['07.16', 'petroleum coke'],
    'other products': ['07.17', 'other products'],
    'other hydrocarbons': ['06.05', 'other hydrocarbons'],
    'additives and oxygenates': ['06.04', 'additives'],
    'natural gas liquids': ['06.02', 'natural gas liquids'],
    'ethane': ['07.11', 'ethane'],
    'white spirit sbp': ['07.12', 'white spirit'],
    'lubricants': ['07.13', 'lubricants'],
    'bitumen': ['07.14', 'bitumen'],
    'petprod nonspecified': ['07.99', 'petprod nonspecified'],
    'coking coal': ['01.01', 'coking coal'],
    'other bituminous coal': ['01.02', 'other bituminous coal'],
    'sub bituminous coal': ['01.03', 'sub-bituminous'],
    'lignite': ['01.05', 'lignite'],
    'coal': ['01 coal', '01.'],
    'coke oven coke': ['02.01', 'coke oven coke'],
    'hydrogen': ['16.12', 'hydrogen'],
    'electricity': ['17 electricity'],
    'heat': ['18 heat'],
    'biomass': ['15 solid biomass', '15.'],
    'biogas': ['16.01', 'biogas'],
    'biodiesel': ['16.06', 'biodiesel'],
    'biogasoline': ['16.05', 'biogasoline'],
    'other liquid biofuels': ['16.08', 'other liquid biofuels'],
    'municipal solid waste': ['16.03', '16.04', 'municipal solid waste'],
}


def load_esto_data():
    """Load augmented ESTO data (includes 9th projection years up to 2060)."""
    print("Loading ESTO data (this may take a moment)...")
    from codebase.transformation_workflow import core
    # Use raw ESTO data (actual historical measurements, not 9th-edition projections)
    esto = core.normalize_esto_economy_codes(core.esto_data_raw.copy())
    print(f"  ESTO raw data loaded: {esto.shape[0]} rows")
    year_cols = sorted([c for c in esto.columns if isinstance(c, int)])
    print(f"  Year range: {min(year_cols)}–{max(year_cols)}")
    return esto


def load_missing_branches():
    path = os.path.join(
        REPO_ROOT,
        'outputs', 'leap_exports', 'results_supply_link',
        'supporting_files', 'checks', 'missing_branch_ids.csv'
    )
    df = pd.read_csv(path)
    print(f"Loaded missing branch IDs: {len(df)} rows from {path}")
    return df


def load_full_model_export():
    path = os.path.join(REPO_ROOT, 'data', 'full model export.xlsx')
    df = pd.read_excel(path, sheet_name='Export', header=2)
    print(f"Loaded full model export: {len(df)} rows")
    return df


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def extract_path_parts(branch_path: str):
    """
    Return (sector, fuel, category) from a Transformation branch path.

    Examples:
      Transformation\\NG Liquefaction\\Processes\\Liquefaction\\Feedstock Fuels\\Natural gas
        → ('NG Liquefaction', 'Natural gas', 'Feedstock Fuels')

      Transformation\\Oil refineries\\Output Fuels\\Motor gasoline
        → ('Oil refineries', 'Motor gasoline', 'Output Fuels')

      Transformation\\NG Regasification\\Processes\\Regasification
        → ('NG Regasification', None, None)   ← no fuel segment
    """
    parts = branch_path.split('\\')
    if len(parts) < 2:
        return None, None, 'Other'

    sector = parts[1] if len(parts) > 1 else None
    fuel = parts[-1] if len(parts) > 2 else None

    category = 'Other'
    for cat in ('Feedstock Fuels', 'Auxiliary Fuels', 'Output Fuels'):
        if cat in parts:
            category = cat
            break

    # If the last segment IS the category label (no fuel after it), fuel = None
    if fuel in ('Feedstock Fuels', 'Auxiliary Fuels', 'Output Fuels', 'Processes'):
        fuel = None

    return sector, fuel, category


def get_flow_codes_for_sector(sector: str) -> list:
    """Return list of flow code prefixes to search for a given sector."""
    return SECTOR_TO_FLOW_HINTS.get(sector, [])


def get_product_hints(fuel: str) -> list:
    """Return list of substrings to test against ESTO products column."""
    if not fuel:
        return []
    key = fuel.lower().strip()
    # direct match
    if key in FUEL_PRODUCT_HINTS:
        return FUEL_PRODUCT_HINTS[key]
    # partial key match
    for k, v in FUEL_PRODUCT_HINTS.items():
        if k in key or key in k:
            return v
    # fallback: use the fuel name itself as a substring
    return [key]


def filter_esto_by_flow(esto: pd.DataFrame, flow_hints: list) -> pd.DataFrame:
    """Filter ESTO rows where 'flows' contains any of the given prefixes."""
    if not flow_hints:
        return esto.iloc[0:0]  # empty
    mask = pd.Series(False, index=esto.index)
    for hint in flow_hints:
        mask |= esto['flows'].str.contains(hint, case=False, na=False)
    return esto[mask]


def filter_esto_by_product(esto: pd.DataFrame, product_hints: list) -> pd.DataFrame:
    """Filter ESTO rows where 'products' matches any hint."""
    if not product_hints:
        return esto.iloc[0:0]
    mask = pd.Series(False, index=esto.index)
    for hint in product_hints:
        mask |= esto['products'].str.contains(hint, case=False, na=False)
    return esto[mask]


def check_year_values(rows: pd.DataFrame, year: int, sign: str) -> bool:
    """
    Check whether the given rows have values in the requested year matching sign.

    sign: 'negative' | 'positive' | 'nonzero'
    Returns True if any matching value found.
    """
    if year not in rows.columns or rows.empty:
        return False
    vals = pd.to_numeric(rows[year], errors='coerce').dropna()
    if vals.empty:
        return False
    if sign == 'negative':
        return bool((vals < 0).any())
    elif sign == 'positive':
        return bool((vals > 0).any())
    else:  # nonzero
        return bool((vals != 0).any())


def check_any_year_values(rows: pd.DataFrame, sign: str, all_year_cols: list) -> tuple:
    """Check across ALL available years (all-years fallback). Returns (found, first_year_with_data)."""
    for yr in sorted(all_year_cols, reverse=True):  # most recent first
        if yr not in rows.columns:
            continue
        vals = pd.to_numeric(rows[yr], errors='coerce').dropna()
        if vals.empty:
            continue
        if sign == 'negative' and bool((vals < 0).any()):
            return True, yr
        elif sign == 'positive' and bool((vals > 0).any()):
            return True, yr
        elif sign == 'nonzero' and bool((vals != 0).any()):
            return True, yr
    return False, None


def check_esto_for_row(esto: pd.DataFrame, sector: str, fuel: str,
                       category: str, year: int, all_year_cols: list) -> tuple:
    """
    Returns (result, note) where result is:
      'has_data'        — data found in the requested year
      'historical_only' — no data in requested year but found in earlier years
                          (mirrors the all-years fallback in the transformation code)
      'no_data'         — no matching ESTO data in any year
      'unknown'         — cannot determine (no flow/product mapping)
    """
    if category == 'Other' or not sector or not fuel:
        return 'unknown', 'Non-transformation path or missing sector/fuel'

    flow_hints = get_flow_codes_for_sector(sector)
    product_hints = get_product_hints(fuel)

    if not flow_hints:
        return 'unknown', f'No flow mapping for sector "{sector}"'
    if not product_hints:
        return 'unknown', f'No product hints for fuel "{fuel}"'

    if category == 'Auxiliary Fuels':
        loss_rows = esto[esto['flows'].str.startswith(LOSS_FLOW_PREFIX, na=False)]
        sector_loss_hints = _get_loss_flow_hints(sector)
        if sector_loss_hints:
            sector_loss_rows = filter_esto_by_flow(loss_rows, sector_loss_hints)
            if not sector_loss_rows.empty:
                loss_rows = sector_loss_rows
        prod_rows = filter_esto_by_product(loss_rows, product_hints)
        sign = 'nonzero'
    elif category == 'Feedstock Fuels':
        flow_rows = filter_esto_by_flow(esto, flow_hints)
        prod_rows = filter_esto_by_product(flow_rows, product_hints)
        sign = 'negative'
    elif category == 'Output Fuels':
        flow_rows = filter_esto_by_flow(esto, flow_hints)
        prod_rows = filter_esto_by_product(flow_rows, product_hints)
        sign = 'positive'
    else:
        return 'unknown', 'Unhandled category'

    if year not in esto.columns:
        return 'unknown', f'Year {year} not in ESTO data'

    # Primary check: requested year
    if check_year_values(prod_rows, year, sign):
        return 'has_data', f'{sign} values found in {year}'

    # Fallback: check all historical years (mirrors allow_all_years_fallback=True)
    hist_found, hist_year = check_any_year_values(prod_rows, sign, all_year_cols)
    if hist_found:
        return 'historical_only', (
            f'No {sign} values in {year} but found in {hist_year} '
            f'(historical only — no projection-period data)'
        )

    return 'no_data', (
        f'No {sign} values for fuel "{fuel}" in any year '
        f'(searched {len(prod_rows)} product-matching rows)'
    )


# Map sector names to own-use flow code substrings
SECTOR_LOSS_FLOW_HINTS = {
    'NG Liquefaction': ['10.01.03'],
    'LNG regasification': ['10.01.03'],
    'Oil Refining': ['10.01.11'],
    'Gas works plants': ['10.01.02'],
    'BKB and PB plants': ['10.01.09'],
    'Blast furnaces': ['10.01.07'],
    'Hydrogen transformation': ['10.01.19'],
    'Non specified transformation': ['10.01.17'],
    'Gas to liquids plants': ['10.01.04'],
    'Petrochemical industry': ['10.01'],
    'Patent fuel plants': ['10.01.08'],
    'Refinery and blending transfers': ['10.01.11'],
    'Transfers unallocated': ['10.01.17'],
    'Transfers': ['10.01.11', '10.01.17'],
    'Natural gas blending plants': ['10.01'],
    'Coke ovens': ['10.01.05'],
    'Electricity Generation': ['10.01.01'],
    'Transmission and Distribution': ['10.02'],
}


def _get_loss_flow_hints(sector: str) -> list:
    return SECTOR_LOSS_FLOW_HINTS.get(sector, [])


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    # Load data
    esto = load_esto_data()
    missing = load_missing_branches()

    all_year_cols = sorted([c for c in esto.columns if isinstance(c, int)])
    latest_year  = all_year_cols[-1] if all_year_cols else None
    prev_year    = all_year_cols[-2] if len(all_year_cols) >= 2 else latest_year
    print(f"  Using latest ESTO years: {prev_year} and {latest_year}")

    # Parse each branch path
    records = []
    for _, row in missing.iterrows():
        bp = row['Branch Path']
        variable = row['Variable']
        affected = row['Affected Economies']

        sector, fuel, category = extract_path_parts(str(bp))

        # Only do ESTO check for Transformation paths
        if not str(bp).startswith('Transformation'):
            category = 'Other'

        # ESTO checks for the two most recent years available, with all-years fallback
        result_2024, note_2024 = check_esto_for_row(esto, sector, fuel, category, prev_year,   all_year_cols)
        result_2025, note_2025 = check_esto_for_row(esto, sector, fuel, category, latest_year, all_year_cols)

        # Combined note (use 2024 note as primary)
        note = note_2024

        records.append({
            'Branch Path': bp,
            'Variable': variable,
            'Affected Economies': affected,
            'Fuel': fuel or '',
            'Sector': sector or '',
            'Category': category,
            f'esto_check_{prev_year}': result_2024,
            f'esto_check_{latest_year}': result_2025,
            'note': note,
        })

    out_df = pd.DataFrame(records)
    col_prev   = f'esto_check_{prev_year}'
    col_latest = f'esto_check_{latest_year}'

    # Save full analysis
    out_path = os.path.join(
        REPO_ROOT,
        'outputs', 'leap_exports', 'results_supply_link',
        'supporting_files', 'checks', 'missing_branch_ids_esto_check.csv'
    )
    out_df.to_csv(out_path, index=False)
    print(f"\nSaved full analysis ({len(out_df)} rows) to:\n  {out_path}")

    # Save action list: only branches with active ESTO data in recent years.
    # Excludes:
    #   historical_only — data existed pre-2022 but not in current period; no current need
    #   no_data         — no ESTO evidence; branch may be from model structure only
    #   unknown/Other   — non-transformation or unmapped sector
    required = out_df[
        (out_df[col_latest] == 'has_data') |
        (out_df[col_prev]   == 'has_data')
    ][['Branch Path', 'Variable', 'Affected Economies', 'Sector', 'Fuel',
       'Category', col_prev, col_latest, 'note']].copy()
    # Drop ESTO aggregate/subtotal fuels — they are group-level products that
    # should never appear as individual LEAP branch fuels.
    subtotal_mask = required['Fuel'].str.lower().isin(ESTO_SUBTOTAL_FUEL_LABELS)
    if subtotal_mask.any():
        dropped = required[subtotal_mask][['Branch Path', 'Fuel']].drop_duplicates()
        print(f"\n[INFO] Excluding {subtotal_mask.sum()} ESTO subtotal fuel rows from required list:")
        for _, r in dropped.iterrows():
            print(f"  {r['Fuel']} | {r['Branch Path']}")
        required = required[~subtotal_mask].reset_index(drop=True)
    required_path = os.path.join(
        REPO_ROOT,
        'outputs', 'leap_exports', 'results_supply_link',
        'supporting_files', 'checks', 'branches_to_add_to_leap_export.csv'
    )
    required.to_csv(required_path, index=False)
    print(f"Saved required-for-LEAP-export list ({len(required)} rows) to:\n  {required_path}")

    # -----------------------------------------------------------------------
    # Summary
    # -----------------------------------------------------------------------

    print("\n" + "=" * 70)
    print(f"SUMMARY: Category × {col_prev}")
    print("=" * 70)
    print(out_df.groupby(['Category', col_prev]).size().reset_index(name='count').to_string(index=False))

    print("\n" + "=" * 70)
    print(f"SUMMARY: Category × {col_latest}")
    print("=" * 70)
    print(out_df.groupby(['Category', col_latest]).size().reset_index(name='count').to_string(index=False))

    trans_df = out_df[out_df['Category'].isin(['Feedstock Fuels', 'Auxiliary Fuels', 'Output Fuels'])]
    if not trans_df.empty:
        print("\n" + "=" * 70)
        print(f"TRANSFORMATION ROWS — {latest_year} data status")
        print("=" * 70)
        for cat in ['Feedstock Fuels', 'Auxiliary Fuels', 'Output Fuels']:
            sub = trans_df[trans_df['Category'] == cat]
            if sub.empty:
                continue
            print(f"\n  {cat} ({len(sub)} rows):")
            vc = sub[col_latest].value_counts()
            for status, cnt in vc.items():
                pct = 100 * cnt / len(sub)
                print(f"    {status:10s}: {cnt:3d} ({pct:.0f}%)")
            # Show no_data rows
            no_data = sub[sub[col_latest] == 'no_data']
            if not no_data.empty:
                print(f"    no_data rows:")
                for _, r in no_data.iterrows():
                    print(f"      {r['Sector']:35s} | {r['Fuel']:30s} | {r['note'][:60]}")

    print("\nDone.")


if __name__ == '__main__':
    main()
