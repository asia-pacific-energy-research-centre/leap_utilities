import pandas as pd
from pathlib import Path


def _truthy(value: object) -> bool:
    return str(value or '').strip().lower() in {'1', 'true', 't', 'yes', 'y', 'on'}


def _deepest_non_x(row: pd.Series, columns: list[str]) -> str:
    for column in reversed(columns):
        value = str(row.get(column, '')).strip()
        if value and value.lower() != 'x':
            return value
    return ''


def _normalize_hierarchy_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = df[columns].copy()
    for column in columns:
        out[column] = out[column].astype(str).str.strip()
    out = out.replace('', pd.NA)
    out = out.mask(out.apply(lambda s: s.str.lower().eq('x')))
    return out


def _list_unique(values: pd.Series, limit: int = 12) -> str:
    cleaned = sorted(set(str(v).strip() for v in values if str(v).strip()))
    shown = cleaned[:limit]
    if len(cleaned) > limit:
        shown.append(f'...+{len(cleaned)-limit} more')
    return ' | '.join(shown)


status_path = Path('outputs/mappings/mapping_checks/ninth_pairs_join_vs_master_status.csv')
mismatch_path = Path('outputs/mappings/mapping_checks/ninth_pairs_join_vs_master_mismatches.csv')
bridge_path = Path('outputs/mappings/mapping_checks/ninth_pairs_join_master_bridge.csv')

# --- ESTO sums, excluding subtotals ---
esto = pd.read_csv('data/00APEC_2025_low_with_subtotals.csv', dtype=object).fillna('')
esto['flows'] = esto.get('flows', '').astype(str).str.strip()
esto['products'] = esto.get('products', '').astype(str).str.strip()
if 'is_subtotal' in esto.columns:
    esto['_is_subtotal'] = esto['is_subtotal'].map(_truthy)
else:
    esto['_is_subtotal'] = False
esto_non_subtotal = esto[~esto['_is_subtotal']].copy()

y_esto = [c for c in esto_non_subtotal.columns if str(c).isdigit()]
esto_non_subtotal['_abs'] = (
    esto_non_subtotal[y_esto]
    .apply(pd.to_numeric, errors='coerce')
    .fillna(0.0)
    .abs()
    .sum(axis=1)
)
esto_lookup = (
    esto_non_subtotal[(esto_non_subtotal['flows'] != '') & (esto_non_subtotal['products'] != '')]
    .groupby(['flows', 'products'], as_index=False)['_abs']
    .sum()
    .rename(columns={'flows': 'esto_flow', 'products': 'esto_product', '_abs': 'esto_all_years_abs_sum'})
)

# --- Ninth sums, excluding subtotals with year-aware logic ---
ninth = pd.read_csv('data/merged_file_energy_ALL_20251106.csv', dtype=object, low_memory=False).fillna('')
sector_cols = ['sectors', 'sub1sectors', 'sub2sectors', 'sub3sectors', 'sub4sectors']
fuel_cols = ['fuels', 'subfuels']
for column in sector_cols + fuel_cols:
    if column not in ninth.columns:
        ninth[column] = ''
    ninth[column] = ninth[column].astype(str).str.strip()

sector_h = _normalize_hierarchy_columns(ninth, sector_cols)
fuel_h = _normalize_hierarchy_columns(ninth, fuel_cols)
ninth['ninth_sector'] = sector_h[list(reversed(sector_cols))].bfill(axis=1).iloc[:, 0].fillna('')
ninth['ninth_fuel'] = fuel_h[list(reversed(fuel_cols))].bfill(axis=1).iloc[:, 0].fillna('')
ninth['_subtotal_results'] = ninth.get('subtotal_results', '').map(_truthy)
ninth['_subtotal_layout'] = ninth.get('subtotal_layout', '').map(_truthy)

y_ninth = [c for c in ninth.columns if str(c).isdigit()]
ninth_long = ninth[['ninth_sector', 'ninth_fuel', '_subtotal_results', '_subtotal_layout', *y_ninth]].melt(
    id_vars=['ninth_sector', 'ninth_fuel', '_subtotal_results', '_subtotal_layout'],
    value_vars=y_ninth,
    var_name='year',
    value_name='value',
)
ninth_long['year'] = pd.to_numeric(ninth_long['year'], errors='coerce').astype('Int64')
ninth_long['value'] = pd.to_numeric(ninth_long['value'], errors='coerce').fillna(0.0)

# Drop subtotal_results only for years > 2022, and subtotal_layout only for years <= 2022
ninth_long = ninth_long[
    ~((ninth_long['year'] > 2022) & (ninth_long['_subtotal_results']))
    & ~((ninth_long['year'] <= 2022) & (ninth_long['_subtotal_layout']))
].copy()
ninth_long['_abs'] = ninth_long['value'].abs()

ninth_lookup = (
    ninth_long[(ninth_long['ninth_sector'] != '') & (ninth_long['ninth_fuel'] != '')]
    .groupby(['ninth_sector', 'ninth_fuel'], as_index=False)['_abs']
    .sum()
    .rename(columns={'_abs': 'ninth_all_years_abs_sum'})
)

# --- Apply sums to status outputs ---
for path in [status_path, mismatch_path]:
    df = pd.read_csv(path).fillna('')
    for column in ['ninth_sector', 'ninth_fuel', 'esto_flow', 'esto_product']:
        df[column] = df[column].astype(str).str.strip()
    for column in ['esto_all_years_abs_sum', 'ninth_all_years_abs_sum']:
        if column in df.columns:
            df = df.drop(columns=[column])

    df = df.merge(esto_lookup, on=['esto_flow', 'esto_product'], how='left')
    df = df.merge(ninth_lookup, on=['ninth_sector', 'ninth_fuel'], how='left')
    df['esto_all_years_abs_sum'] = pd.to_numeric(df['esto_all_years_abs_sum'], errors='coerce').fillna(0.0)
    df['ninth_all_years_abs_sum'] = pd.to_numeric(df['ninth_all_years_abs_sum'], errors='coerce').fillna(0.0)
    df.to_csv(path, index=False)
    print(f'updated {path} rows={len(df)}')

# --- Build mismatch cause bridge: what the other set maps instead ---
status = pd.read_csv(status_path).fillna('')
for column in ['ninth_sector', 'ninth_fuel', 'esto_flow', 'esto_product', 'status']:
    status[column] = status[column].astype(str).str.strip()

join_only = status[status['status'] == 'only_in_join_implied'].copy()
master_only = status[status['status'] == 'only_in_master'].copy()

def _decorate(base: pd.DataFrame, other: pd.DataFrame, base_label: str, other_label: str) -> pd.DataFrame:
    out = base.copy()

    same_ninth = (
        other.groupby(['ninth_sector', 'ninth_fuel'], as_index=False)
        .agg(
            other_same_ninth_count=('status', 'size'),
            other_same_ninth_pairs=('esto_flow', lambda s: _list_unique(s + ' -> ' + other.loc[s.index, 'esto_product'])),
        )
    )
    same_esto = (
        other.groupby(['esto_flow', 'esto_product'], as_index=False)
        .agg(
            other_same_esto_count=('status', 'size'),
            other_same_esto_pairs=('ninth_sector', lambda s: _list_unique(s + ' -> ' + other.loc[s.index, 'ninth_fuel'])),
        )
    )

    out = out.merge(same_ninth, on=['ninth_sector', 'ninth_fuel'], how='left')
    out = out.merge(same_esto, on=['esto_flow', 'esto_product'], how='left')
    out['other_same_ninth_count'] = pd.to_numeric(out['other_same_ninth_count'], errors='coerce').fillna(0).astype(int)
    out['other_same_esto_count'] = pd.to_numeric(out['other_same_esto_count'], errors='coerce').fillna(0).astype(int)
    out['other_same_ninth_pairs'] = out['other_same_ninth_pairs'].fillna('')
    out['other_same_esto_pairs'] = out['other_same_esto_pairs'].fillna('')

    def _cause(row: pd.Series) -> str:
        n = int(row['other_same_ninth_count'])
        e = int(row['other_same_esto_count'])
        if n > 0 and e > 0:
            return 'both_ninth_and_esto_have_alternative_mappings'
        if n > 0:
            return 'different_esto_target_for_same_ninth_pair'
        if e > 0:
            return 'different_ninth_source_for_same_esto_pair'
        return 'no_direct_counterpart_in_other_set'

    out['mismatch_cause_hint'] = out.apply(_cause, axis=1)
    out['base_set'] = base_label
    out['other_set'] = other_label
    return out


bridge_join = _decorate(join_only, master_only, 'only_in_join_implied', 'only_in_master')
bridge_master = _decorate(master_only, join_only, 'only_in_master', 'only_in_join_implied')
bridge = pd.concat([bridge_join, bridge_master], ignore_index=True)
bridge.to_csv(bridge_path, index=False)
print(f'updated {bridge_path} rows={len(bridge)}')

summary = (
    status.assign(
        esto_nz=status['esto_all_years_abs_sum'] > 0,
        ninth_nz=status['ninth_all_years_abs_sum'] > 0,
    )
    .groupby(['status', 'esto_nz', 'ninth_nz'])
    .size()
    .reset_index(name='rows')
)
print('\nstatus nonzero breakdown:')
print(summary.to_string(index=False))

cause_summary = bridge.groupby(['base_set', 'mismatch_cause_hint']).size().reset_index(name='rows')
print('\nmismatch cause bridge summary:')
print(cause_summary.to_string(index=False))
