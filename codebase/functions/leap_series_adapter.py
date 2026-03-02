from __future__ import annotations

from typing import Iterable, Literal

import pandas as pd

from codebase.functions.leap_expressions import (
    build_data_expression,
    collect_expression_years,
    expression_to_series,
    parse_expression,
    scale_expression,
    sum_expressions,
)

SeriesFormat = Literal["expression", "year_columns"]


def _coerce_year(value: object) -> int | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, int):
        return value if 1000 <= value <= 3000 else None
    if isinstance(value, float) and value.is_integer():
        year = int(value)
        return year if 1000 <= year <= 3000 else None
    text = str(value).strip()
    if text.isdigit() and len(text) == 4:
        return int(text)
    return None


def _year_columns_from_columns(columns: Iterable[object]) -> list[object]:
    year_cols: list[tuple[int, object]] = []
    for col in columns:
        year = _coerce_year(col)
        if year is None:
            continue
        year_cols.append((year, col))
    return [original for _, original in sorted(year_cols, key=lambda item: item[0])]


def _resolve_year_columns(df: pd.DataFrame, year_cols: Iterable[int | str] | None) -> list[tuple[int, object]]:
    if year_cols is None:
        candidates = _year_columns_from_columns(df.columns)
    else:
        candidates = list(year_cols)
    resolved: list[tuple[int, object]] = []
    for col in candidates:
        if col in df.columns:
            year = _coerce_year(col)
            if year is not None:
                resolved.append((year, col))
            continue
        text = str(col).strip()
        if text in df.columns:
            year = _coerce_year(text)
            if year is not None:
                resolved.append((year, text))
            continue
        year = _coerce_year(col)
        if year is not None and year in df.columns:
            resolved.append((year, year))
    return sorted(dict(resolved).items(), key=lambda item: item[0])


def _get_value_for_year_label(row: pd.Series, col: object) -> object:
    if col in row.index:
        return row.get(col)
    year = _coerce_year(col)
    if year is None:
        return row.get(col)
    if year in row.index:
        return row.get(year)
    text = str(year)
    if text in row.index:
        return row.get(text)
    return pd.NA


def _set_value_for_year_label(updated: pd.Series, col: object, value: object) -> None:
    if col in updated.index:
        updated[col] = value
        return
    year = _coerce_year(col)
    if year is None:
        updated[col] = value
        return
    if year in updated.index:
        updated[year] = value
        return
    text = str(year)
    if text in updated.index:
        updated[text] = value
        return
    updated[col] = value


def _resolve_df_column_label(df: pd.DataFrame, col: object) -> object | None:
    if col in df.columns:
        return col
    year = _coerce_year(col)
    if year is None:
        return None
    if year in df.columns:
        return year
    text = str(year)
    if text in df.columns:
        return text
    return None


def detect_series_format(df: pd.DataFrame) -> SeriesFormat:
    has_expression_col = "Expression" in df.columns
    expression_populated = False
    if has_expression_col:
        expression_populated = (
            df["Expression"].astype("string").fillna("").str.strip() != ""
        ).any()

    year_cols = _year_columns_from_columns(df.columns)
    year_populated = False
    if year_cols:
        numeric = df[year_cols].apply(pd.to_numeric, errors="coerce")
        year_populated = numeric.notna().any(axis=None)

    if expression_populated and year_populated:
        raise ValueError(
            "Detected mixed series format: both non-empty Expression and year columns are present."
        )
    if expression_populated:
        return "expression"
    if year_cols:
        return "year_columns"
    if has_expression_col:
        return "expression"
    raise ValueError("Unable to detect series format: no Expression column or year columns found.")


def collect_available_years(df: pd.DataFrame, series_format: SeriesFormat) -> list[int]:
    if series_format == "expression":
        years = collect_expression_years(df.get("Expression", pd.Series(dtype=object)).tolist())
        return sorted(int(year) for year in years)
    resolved = _resolve_year_columns(df, year_cols=None)
    return [year for year, _col in resolved]


def extract_row_series(
    row: pd.Series,
    series_format: SeriesFormat,
    year_cols: Iterable[int | str] | None = None,
) -> dict[int, float] | float | None:
    if series_format == "expression":
        years = [int(year) for year in (year_cols or [])]
        if years:
            return expression_to_series(row.get("Expression"), years=years, base_year=years[0] if years else None)
        mode, payload = parse_expression(row.get("Expression"))
        if mode == "series" and isinstance(payload, dict):
            return {int(year): float(value) for year, value in payload.items()}
        if mode == "const" and payload is not None:
            return float(payload)
        if mode == "empty":
            return {}
        return None

    series: dict[int, float] = {}
    if year_cols is None:
        candidate_cols = [col for col in row.index if _coerce_year(col) is not None]
    else:
        candidate_cols = list(year_cols)
    for col in candidate_cols:
        year = _coerce_year(col)
        if year is None:
            continue
        value = pd.to_numeric(_get_value_for_year_label(row, col), errors="coerce")
        if pd.isna(value):
            value = pd.to_numeric(_get_value_for_year_label(row, str(col)), errors="coerce")
        if pd.isna(value):
            continue
        series[int(year)] = float(value)
    return series


def scale_row_series(
    row: pd.Series,
    scale_by: float | dict[int, float],
    base_year: int,
    series_format: SeriesFormat,
    year_cols: Iterable[int | str] | None = None,
    fallback_share: float | None = None,
) -> pd.Series:
    updated = row.copy()
    if series_format == "expression":
        updated["Expression"] = scale_expression(
            row.get("Expression"),
            share=scale_by,
            base_year=base_year,
            fallback_share=fallback_share,
        )
        return updated

    if year_cols is None:
        candidate_cols = [col for col in row.index if _coerce_year(col) is not None]
    else:
        candidate_cols = list(year_cols)
    if isinstance(scale_by, dict):
        fallback = fallback_share if fallback_share is not None else 0.0
        for col in candidate_cols:
            year = _coerce_year(col)
            if year is None:
                continue
            current = pd.to_numeric(_get_value_for_year_label(row, col), errors="coerce")
            if pd.isna(current):
                current = pd.to_numeric(_get_value_for_year_label(row, str(col)), errors="coerce")
            if pd.isna(current):
                continue
            _set_value_for_year_label(
                updated,
                col,
                float(current) * float(scale_by.get(int(year), fallback)),
            )
        return updated

    scalar = float(scale_by)
    for col in candidate_cols:
        current = pd.to_numeric(_get_value_for_year_label(row, col), errors="coerce")
        if pd.isna(current):
            current = pd.to_numeric(_get_value_for_year_label(row, str(col)), errors="coerce")
        if pd.isna(current):
            continue
        _set_value_for_year_label(updated, col, float(current) * scalar)
    return updated


def sum_rows_series(
    rows: pd.DataFrame,
    series_format: SeriesFormat,
    year_cols: Iterable[int | str] | None = None,
) -> str | dict[int, float]:
    if rows.empty:
        return {} if series_format == "year_columns" else ""

    if series_format == "expression":
        return sum_expressions(rows.get("Expression", pd.Series(dtype=object)).tolist())

    if year_cols is None:
        candidates = [col for col in rows.columns if _coerce_year(col) is not None]
    else:
        candidates = list(year_cols)
    sums: dict[int, float] = {}
    for col in candidates:
        year = _coerce_year(col)
        if year is None:
            continue
        resolved_col = _resolve_df_column_label(rows, col)
        if resolved_col is None:
            continue
        values = pd.to_numeric(rows[resolved_col], errors="coerce")
        total = float(values.fillna(0.0).sum())
        sums[int(year)] = total
    return sums


def inject_row_series(
    row: pd.Series,
    series: dict[int, float] | float | str | None,
    series_format: SeriesFormat,
    year_cols: Iterable[int | str] | None = None,
) -> pd.Series:
    updated = row.copy()
    if series is None:
        return updated

    if series_format == "expression":
        if isinstance(series, dict):
            updated["Expression"] = build_data_expression(
                {int(year): float(value) for year, value in sorted(series.items())}
            )
            return updated
        mode, payload = parse_expression(series)
        if mode == "empty":
            updated["Expression"] = ""
            return updated
        if mode == "series" and isinstance(payload, dict):
            updated["Expression"] = build_data_expression(
                {int(year): float(value) for year, value in sorted(payload.items())}
            )
            return updated
        if mode == "const" and payload is not None:
            updated["Expression"] = str(float(payload))
            return updated
        updated["Expression"] = str(series)
        return updated

    if year_cols is None:
        target_cols = [col for col in row.index if _coerce_year(col) is not None]
    else:
        target_cols = list(year_cols)

    if isinstance(series, dict):
        for col in target_cols:
            year = _coerce_year(col)
            if year is None:
                continue
            if int(year) in series:
                _set_value_for_year_label(updated, col, float(series[int(year)]))
        return updated

    for col in target_cols:
        if _coerce_year(col) is None:
            continue
        _set_value_for_year_label(updated, col, float(series))
    return updated


__all__ = [
    "SeriesFormat",
    "detect_series_format",
    "extract_row_series",
    "scale_row_series",
    "sum_rows_series",
    "inject_row_series",
    "collect_available_years",
]
