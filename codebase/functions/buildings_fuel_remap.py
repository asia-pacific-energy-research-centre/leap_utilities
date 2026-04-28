from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal

import pandas as pd

from codebase.utilities.master_config import read_config_table
from codebase.functions.leap_excel_io import read_export_sheet, save_export_files
from codebase.functions.leap_labels import clean_fuel_label_for_leap
from codebase.functions.leap_series_adapter import (
    SeriesFormat,
    collect_available_years,
    detect_series_format,
    extract_row_series,
    inject_row_series,
    scale_row_series,
    sum_rows_series,
)
from codebase.functions.ninth_projection_mapping import (
    build_esto_base_year_values,
    compute_esto_base_year_shares,
    normalize_economy_key,
)
from codebase.scrapbook.utilities import (
    apply_matt_subtotal_mapping,
    filter_matt_subtotals,
    load_augmented_reference_tables,
)
from codebase.utilities.workflow_common import archive_config_dir_once_per_day

CURRENT_ACCOUNT_LABELS = {"current accounts", "current account"}
FLOW_BY_SECTOR_PREFIX = {
    "RESIDENTIAL": "16.02 Residential",
    "SERVICES": "16.01 Commercial and public services",
}


@dataclass
class MappingRow:
    sector_key: str
    end_use: str
    technology: str
    canonical_technology: str
    mapping_mode: str
    esto_flow: str
    target_products: list[str]
    notes: str


def _normalize_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _normalize_key(value: object) -> str:
    return _normalize_text(value).lower()


def _normalize_year_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {col: int(col) for col in df.columns if str(col).isdigit()}
    return df.rename(columns=rename_map)


def _load_esto_data(path: Path, subtotal_mapping_path: Path) -> pd.DataFrame:
    archive_config_dir_once_per_day()
    df, _ = load_augmented_reference_tables(
        esto_path=path,
        ninth_path=Path("data/merged_file_energy_ALL_20251106.csv"),
        subtotal_mapping_path=subtotal_mapping_path,
        synthetic_rules_path=Path("config/synthetic_reference_rows.csv"),
        cache_dir=Path("data/.cache/buildings_reference_tables"),
        apply_esto_subtotal_map=True,
        filter_esto_subtotals_flag=True,
        filter_ninth_subtotals_flag=False,
    )
    df = _normalize_year_columns(df)
    df["economy_key"] = df["economy"].apply(normalize_economy_key)
    df["flows"] = df["flows"].astype(str).str.strip()
    df["products"] = df["products"].astype(str).str.strip()
    return df


def _load_mapping(mapping_path: Path) -> dict[tuple[str, str, str], MappingRow]:
    df = read_config_table(mapping_path).fillna("")
    rows: dict[tuple[str, str, str], MappingRow] = {}
    for _, row in df.iterrows():
        sector_key = _normalize_text(row.get("sector_key"))
        end_use = _normalize_text(row.get("end_use"))
        technology = _normalize_text(row.get("technology"))
        if not (sector_key and end_use and technology):
            continue
        canonical = _normalize_text(row.get("canonical_technology")) or technology
        mapping_mode = _normalize_key(row.get("mapping_mode")) or "direct"
        esto_flow = _normalize_text(row.get("esto_flow"))
        target_products_raw = _normalize_text(row.get("target_products"))
        target_products = [part.strip() for part in target_products_raw.split(";") if part.strip()]
        notes = _normalize_text(row.get("notes"))
        key = (_normalize_key(sector_key), _normalize_key(end_use), _normalize_key(technology))
        incoming = MappingRow(
            sector_key=sector_key,
            end_use=end_use,
            technology=technology,
            canonical_technology=canonical,
            mapping_mode=mapping_mode,
            esto_flow=esto_flow,
            target_products=target_products,
            notes=notes,
        )
        existing = rows.get(key)
        # Keep explicit/direct mappings when case-variant alias rows collide.
        if existing is None:
            rows[key] = incoming
        elif existing.mapping_mode == "alias" and incoming.mapping_mode != "alias":
            rows[key] = incoming
        elif existing.mapping_mode != "alias" and incoming.mapping_mode == "alias":
            continue
        else:
            rows[key] = incoming
    return rows


def _load_mapping_from_dict(
    mapping_dict: dict[str, dict[str, dict[str, dict[str, object]]]],
) -> dict[tuple[str, str, str], MappingRow]:
    rows: dict[tuple[str, str, str], MappingRow] = {}
    for sector_key, end_uses in mapping_dict.items():
        for end_use, technologies in end_uses.items():
            for technology, payload in technologies.items():
                payload = payload or {}
                target_value = payload.get("target_products", [])
                if isinstance(target_value, str):
                    target_products = [part.strip() for part in target_value.split(";") if part.strip()]
                elif isinstance(target_value, (list, tuple, set)):
                    target_products = [
                        _normalize_text(part) for part in target_value if _normalize_text(part)
                    ]
                else:
                    target_products = []

                canonical = _normalize_text(payload.get("canonical_technology")) or _normalize_text(technology)
                mapping_mode = _normalize_key(payload.get("mapping_mode")) or "direct"
                esto_flow = _normalize_text(payload.get("esto_flow"))
                notes = _normalize_text(payload.get("notes"))

                key = (
                    _normalize_key(sector_key),
                    _normalize_key(end_use),
                    _normalize_key(technology),
                )
                incoming = MappingRow(
                    sector_key=_normalize_text(sector_key),
                    end_use=_normalize_text(end_use),
                    technology=_normalize_text(technology),
                    canonical_technology=canonical,
                    mapping_mode=mapping_mode,
                    esto_flow=esto_flow,
                    target_products=target_products,
                    notes=notes,
                )
                existing = rows.get(key)
                # Keep explicit/direct mappings when case-variant alias rows collide.
                if existing is None:
                    rows[key] = incoming
                elif existing.mapping_mode == "alias" and incoming.mapping_mode != "alias":
                    rows[key] = incoming
                elif existing.mapping_mode != "alias" and incoming.mapping_mode == "alias":
                    continue
                else:
                    rows[key] = incoming
    return rows


def _resolve_mapping(
    mapping_rows: dict[tuple[str, str, str], MappingRow],
    sector: str,
    end_use: str,
    technology: str,
) -> MappingRow | None:
    key = (_normalize_key(sector), _normalize_key(end_use), _normalize_key(technology))
    current = mapping_rows.get(key)
    visited: set[tuple[str, str, str]] = set()
    while current and current.mapping_mode == "alias":
        if key in visited:
            return None
        visited.add(key)
        key = (
            _normalize_key(current.sector_key),
            _normalize_key(current.end_use),
            _normalize_key(current.canonical_technology),
        )
        current = mapping_rows.get(key)
    return current


def _is_buildings_technology_row(
    row: pd.Series,
    level_cols: list[str],
) -> tuple[bool, str, str, str, int]:
    levels = [_normalize_text(row.get(col)) for col in level_cols]
    if len(levels) < 4:
        return False, "", "", "", -1
    if levels[0] != "Demand":
        return False, "", "", "", -1
    sector, end_use, technology = levels[1], levels[2], levels[3]
    if not sector or not end_use or not technology:
        return False, "", "", "", -1
    last_non_empty = max((idx for idx, value in enumerate(levels) if value), default=-1)
    if last_non_empty != 3:
        return False, "", "", "", -1
    return True, sector, end_use, technology, 3


def _update_branch_path(row: pd.Series, level_cols: list[str], level_idx: int, label: str) -> pd.Series:
    levels = [_normalize_text(row.get(col)) for col in level_cols]
    if level_idx >= len(levels):
        return row
    levels[level_idx] = clean_fuel_label_for_leap(label)
    updated = row.copy()
    updated["Branch Path"] = "\\".join([value for value in levels if value])
    for col, value in zip(level_cols, levels):
        updated[col] = value if value else pd.NA
    return updated


def _rebuild_level_columns(df: pd.DataFrame, level_cols: list[str]) -> pd.DataFrame:
    working = df.copy()
    parts = working["Branch Path"].fillna("").astype(str).str.split("\\")
    for idx, col in enumerate(level_cols):
        working[col] = parts.str.get(idx).fillna("").replace({"": pd.NA})
    return working


def _drop_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.columns.is_unique:
        return df
    return df.loc[:, ~df.columns.duplicated()].copy()


def _detect_extra_products_for_flow(
    esto_df: pd.DataFrame,
    base_year: int,
    economy_key: str,
    flow: str,
    mapped_products: set[str],
) -> list[str]:
    if base_year not in esto_df.columns:
        return []
    subset = esto_df[
        (esto_df["economy_key"] == economy_key)
        & (esto_df["flows"] == flow)
    ].copy()
    if subset.empty:
        return []
    subset[base_year] = pd.to_numeric(subset[base_year], errors="coerce").fillna(0.0)
    products = set(subset.loc[subset[base_year] != 0, "products"].astype(str).tolist())
    return sorted(products - mapped_products)


def _product_code(product_label: str) -> str:
    text = _normalize_text(product_label)
    if not text:
        return ""
    return text.split(" ", 1)[0].strip()


def _available_products_for_flow(
    esto_df: pd.DataFrame,
    economy_key: str,
    flow: str,
    base_year: int,
) -> list[str]:
    subset = esto_df[
        (esto_df["economy_key"] == economy_key)
        & (esto_df["flows"] == flow)
    ].copy()
    if subset.empty:
        return []
    if base_year in subset.columns:
        subset[base_year] = pd.to_numeric(subset[base_year], errors="coerce").fillna(0.0)
        subset = subset.loc[subset[base_year] != 0]
    products = sorted({_normalize_text(value) for value in subset["products"].tolist() if _normalize_text(value)})
    return products


def _expand_target_products_for_flow(
    targets: list[str],
    available_products: list[str],
) -> list[str]:
    if not targets:
        return []
    available = [_normalize_text(item) for item in available_products if _normalize_text(item)]
    if not available:
        return list(targets)
    available_set = set(available)
    available_by_code: dict[str, list[str]] = {}
    for product in available:
        code = _product_code(product)
        if not code:
            continue
        available_by_code.setdefault(code, []).append(product)

    expanded: list[str] = []
    seen: set[str] = set()
    for raw_target in targets:
        target = _normalize_text(raw_target)
        if not target:
            continue
        code = _product_code(target)
        # Parent codes (e.g. "01 Coal") are expanded to leaf-level codes when present.
        child_candidates = sorted(
            [
                product
                for product_code, products in available_by_code.items()
                if code and "." not in code and product_code.startswith(f"{code}.")
                for product in products
            ]
        )
        candidates: list[str]
        if child_candidates:
            candidates = child_candidates
        elif target in available_set:
            candidates = [target]
        elif code and code in available_by_code:
            candidates = sorted(available_by_code[code])
        else:
            candidates = [target]

        for candidate in candidates:
            if candidate in seen:
                continue
            seen.add(candidate)
            expanded.append(candidate)
    return expanded


def _flow_for_sector(sector_key: str) -> str:
    sector = _normalize_text(sector_key).upper()
    for prefix, flow in FLOW_BY_SECTOR_PREFIX.items():
        if sector.startswith(prefix):
            return flow
    return ""


def remap_buildings_export_fuels(
    input_path: str | Path,
    output_path: str | Path,
    mapping_csv_path: str | Path | None,
    esto_data_path: str | Path,
    subtotal_mapping_path: str | Path,
    economy: str,
    base_year: int,
    mapping_dict: dict[str, dict[str, dict[str, dict[str, object]]]] | None = None,
    sheet_name: str = "Export",
    include_extra_nonspecified: bool = True,
    report_path: str | Path | None = None,
    validation_path: str | Path | None = None,
    output_series_format: Literal["preserve", "expression", "year_columns"] = "preserve",
) -> dict[str, object]:
    input_path = Path(input_path)
    output_path = Path(output_path)
    mapping_csv = Path(mapping_csv_path) if mapping_csv_path is not None else None
    esto_data_path = Path(esto_data_path)
    subtotal_mapping_path = Path(subtotal_mapping_path)

    header_rows, df, columns = read_export_sheet(input_path, sheet_name)
    series_format = detect_series_format(df)
    year_cols = collect_available_years(df, series_format)
    level_cols = [col for col in columns if str(col).startswith("Level ")]
    id_cols = [col for col in ["BranchID", "VariableID", "ScenarioID", "RegionID"] if col in columns]

    if mapping_dict is not None:
        mapping_rows = _load_mapping_from_dict(mapping_dict)
    elif mapping_csv is not None:
        mapping_rows = _load_mapping(mapping_csv)
    else:
        raise ValueError("Provide either mapping_dict or mapping_csv_path.")
    esto_df = _load_esto_data(esto_data_path, subtotal_mapping_path)
    economy_key = normalize_economy_key(economy)
    base_values = build_esto_base_year_values(esto_df, base_year)

    issues: dict[str, object] = {
        "unmapped_technologies": [],
        "extra_products_by_flow": {},
    }

    available_products_cache: dict[str, list[str]] = {}

    mapped_products_by_flow: dict[str, set[str]] = {}
    others_rows_by_flow: dict[str, list[MappingRow]] = {}
    for mapping in mapping_rows.values():
        if mapping.mapping_mode in {"direct", "split_base_year"}:
            flow = mapping.esto_flow or _flow_for_sector(mapping.sector_key)
            if flow:
                if flow not in available_products_cache:
                    available_products_cache[flow] = _available_products_for_flow(
                        esto_df=esto_df,
                        economy_key=economy_key,
                        flow=flow,
                        base_year=base_year,
                    )
                mapping.target_products = _expand_target_products_for_flow(
                    targets=mapping.target_products,
                    available_products=available_products_cache[flow],
                )
                mapped_products_by_flow.setdefault(flow, set()).update(mapping.target_products)
        if _normalize_key(mapping.canonical_technology) in {"others", "non-specified", "non specified"}:
            flow = mapping.esto_flow or _flow_for_sector(mapping.sector_key)
            if flow:
                others_rows_by_flow.setdefault(flow, []).append(mapping)

    for flow, mapped_products in mapped_products_by_flow.items():
        extras = _detect_extra_products_for_flow(
            esto_df=esto_df,
            base_year=base_year,
            economy_key=economy_key,
            flow=flow,
            mapped_products=mapped_products,
        )
        if extras:
            issues["extra_products_by_flow"][flow] = extras
            if include_extra_nonspecified:
                for mapping in others_rows_by_flow.get(flow, []):
                    existing = set(mapping.target_products)
                    mapping.target_products.extend([product for product in extras if product not in existing])

    new_rows: list[pd.Series] = []
    for _, row in df.iterrows():
        is_tech, sector, end_use, technology, tech_idx = _is_buildings_technology_row(row, level_cols)
        if not is_tech:
            new_rows.append(row)
            continue
        mapping = _resolve_mapping(mapping_rows, sector, end_use, technology)
        if mapping is None:
            issues["unmapped_technologies"].append(
                {"sector": sector, "end_use": end_use, "technology": technology}
            )
            new_rows.append(row)
            continue

        flow = mapping.esto_flow or _flow_for_sector(sector)
        if flow and flow not in available_products_cache:
            available_products_cache[flow] = _available_products_for_flow(
                esto_df=esto_df,
                economy_key=economy_key,
                flow=flow,
                base_year=base_year,
            )
        targets = _expand_target_products_for_flow(
            targets=mapping.target_products,
            available_products=available_products_cache.get(flow, []),
        )
        if not targets:
            issues["unmapped_technologies"].append(
                {"sector": sector, "end_use": end_use, "technology": technology, "reason": "empty target_products"}
            )
            new_rows.append(row)
            continue

        shares: dict[str, float]
        fallback_share: float | None = None
        if mapping.mapping_mode == "split_base_year":
            shares = compute_esto_base_year_shares(
                base_values,
                economy_key=economy_key,
                esto_flow=flow,
                esto_products=targets,
            )
            fallback_share = (sum(shares.values()) / len(shares)) if shares else None
        else:
            shares = {target: 1.0 / len(targets) for target in targets}

        for target in targets:
            share = shares.get(target, 0.0 if mapping.mapping_mode == "split_base_year" else (1.0 / len(targets)))
            updated = _update_branch_path(row, level_cols, tech_idx, target)
            updated = scale_row_series(
                updated,
                scale_by=share,
                base_year=base_year,
                series_format=series_format,
                year_cols=year_cols if year_cols else None,
                fallback_share=fallback_share,
            )
            for id_col in id_cols:
                updated[id_col] = pd.NA
            new_rows.append(updated)

    mapped_df = pd.DataFrame([row.tolist() for row in new_rows], columns=columns)
    key_cols = [col for col in ["Branch Path", "Variable", "Scenario", "Region"] if col in mapped_df.columns]
    if key_cols:
        grouped_rows: list[pd.Series] = []
        for _, group in mapped_df.groupby(key_cols, dropna=False):
            if len(group) == 1:
                grouped_rows.append(group.iloc[0])
                continue
            merged = group.iloc[0].copy()
            summed = sum_rows_series(
                group,
                series_format=series_format,
                year_cols=year_cols if year_cols else None,
            )
            merged = inject_row_series(
                merged,
                series=summed,
                series_format=series_format,
                year_cols=year_cols if year_cols else None,
            )
            for id_col in id_cols:
                merged[id_col] = pd.NA
            grouped_rows.append(merged)
        mapped_df = pd.DataFrame(
            [row.tolist() for row in grouped_rows],
            columns=mapped_df.columns,
        )

    mapped_df = _rebuild_level_columns(mapped_df, level_cols)
    mapped_df = _drop_duplicate_columns(mapped_df)

    if output_series_format not in {"preserve", "expression", "year_columns"}:
        raise ValueError("output_series_format must be one of: preserve, expression, year_columns.")
    target_format: SeriesFormat = (
        series_format
        if output_series_format == "preserve"
        else "expression"
        if output_series_format == "expression"
        else "year_columns"
    )

    if target_format != series_format:
        if target_format == "expression":
            expressions: list[str] = []
            for _, row in mapped_df.iterrows():
                series = extract_row_series(
                    row,
                    series_format="year_columns",
                    year_cols=year_cols if year_cols else None,
                )
                expression = inject_row_series(
                    pd.Series({"Expression": ""}),
                    series=series if isinstance(series, dict) else {},
                    series_format="expression",
                ).get("Expression", "")
                expressions.append(expression)
            mapped_df["Expression"] = expressions
            mapped_df = mapped_df.drop(columns=[col for col in mapped_df.columns if str(col).isdigit()], errors="ignore")
            years = year_cols
        else:
            years = collect_available_years(mapped_df, "expression")
            if not years:
                years = [int(base_year)]
            for year in years:
                mapped_df[year] = pd.NA
            for idx, row in mapped_df.iterrows():
                series = extract_row_series(row, "expression", year_cols=years)
                mapped_df.loc[idx] = inject_row_series(
                    mapped_df.loc[idx],
                    series=series if isinstance(series, dict) else {},
                    series_format="year_columns",
                    year_cols=years,
                )
            mapped_df = mapped_df.drop(columns=["Expression"], errors="ignore")
    else:
        years = collect_available_years(mapped_df, target_format)

    if target_format == "expression":
        viewing_years = years or [int(base_year)]
        viewing_rows = []
        for _, row in mapped_df.iterrows():
            series = extract_row_series(row, "expression", year_cols=viewing_years)
            values = {year: (series.get(year, pd.NA) if isinstance(series, dict) else pd.NA) for year in viewing_years}
            viewing_rows.append(values)
        viewing_df = pd.concat(
            [mapped_df.drop(columns=["Expression"], errors="ignore"), pd.DataFrame(viewing_rows, index=mapped_df.index)],
            axis=1,
        )
    else:
        viewing_df = mapped_df.copy()
        viewing_years = years

    viewing_df = _drop_duplicate_columns(viewing_df)

    model_name = "Model"
    if "Variable" in header_rows.columns:
        values = header_rows["Variable"].dropna().astype(str).str.strip()
        values = values[values != ""]
        if not values.empty:
            model_name = values.iloc[0]

    save_export_files(
        leap_export_df=mapped_df,
        export_df_for_viewing=viewing_df,
        leap_export_filename=output_path,
        base_year=base_year,
        final_year=max(viewing_years) if viewing_years else base_year,
        model_name=model_name,
    )

    if report_path:
        report_rows: list[dict[str, object]] = []
        for item in issues["unmapped_technologies"]:
            report_rows.append({"issue_type": "unmapped_technology", **item})
        for flow, products in issues["extra_products_by_flow"].items():
            for product in products:
                report_rows.append(
                    {
                        "issue_type": "extra_esto_product",
                        "flow": flow,
                        "product": product,
                    }
                )
        if report_rows:
            report_path = Path(report_path)
            report_path.parent.mkdir(parents=True, exist_ok=True)
            pd.DataFrame(report_rows).to_csv(report_path, index=False)

    if validation_path:
        validation_rows = []
        for variable in sorted(mapped_df.get("Variable", pd.Series(dtype=str)).dropna().astype(str).unique()):
            validation_rows.append({"variable": variable, "status": "present"})
        validation_path = Path(validation_path)
        validation_path.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame(validation_rows).to_csv(validation_path, index=False)

    return issues


__all__ = ["remap_buildings_export_fuels"]
