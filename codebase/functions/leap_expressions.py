from __future__ import annotations

from typing import Iterable

import pandas as pd


def parse_expression(expr: object) -> tuple[str, float | dict[int, float] | None]:
    if expr is None or (isinstance(expr, float) and pd.isna(expr)):
        return "empty", None
    text = str(expr).strip()
    if not text:
        return "empty", None
    if text.startswith("Interp(") or text.startswith("Data("):
        body = text[text.find("(") + 1 :]
        if body.endswith(")"):
            body = body[:-1]
        parts = [part.strip() for part in body.split(",") if part.strip()]
        if len(parts) % 2 != 0:
            return "unknown", None
        values: dict[int, float] = {}
        for idx in range(0, len(parts), 2):
            try:
                year = int(float(parts[idx]))
                value = float(parts[idx + 1])
            except ValueError:
                return "unknown", None
            values[year] = value
        return "series", values
    try:
        return "const", float(text)
    except ValueError:
        return "unknown", None


def _format_number(value: float) -> str:
    return f"{value:.12g}"


def build_data_expression(values_by_year: dict[int, float]) -> str:
    parts = [f"{int(year)},{_format_number(value)}" for year, value in sorted(values_by_year.items())]
    return f"Data({', '.join(parts)})"


def build_data_expression_from_row(row: pd.Series, year_cols: Iterable[int]) -> str:
    parts = []
    for year in year_cols:
        value = row.get(year)
        if value is None or pd.isna(value):
            value = 0.0
        parts.append(f"{int(year)},{float(value)}")
    return f"Data({', '.join(parts)})"


def scale_expression(
    expr: object,
    share: float | dict[int, float],
    base_year: int | None = None,
    fallback_share: float | None = None,
) -> str:
    mode, payload = parse_expression(expr)
    if mode == "empty":
        return ""
    if mode == "unknown" or payload is None:
        return str(expr)
    if isinstance(share, dict):
        share_by_year = share
    else:
        share_by_year = None

    if mode == "const":
        value = float(payload)
        if share_by_year is None:
            scale = float(share)
        else:
            scale = share_by_year.get(base_year) if base_year is not None else None
            if scale is None:
                scale = fallback_share if fallback_share is not None else 0.0
        return _format_number(value * scale)

    values_by_year = payload
    if share_by_year is None:
        scaled = {year: value * float(share) for year, value in values_by_year.items()}
        return build_data_expression(scaled)

    fallback = fallback_share if fallback_share is not None else 0.0
    scaled = {}
    for year, value in values_by_year.items():
        scale = share_by_year.get(year, fallback)
        scaled[year] = value * scale
    return build_data_expression(scaled)


def sum_expressions(expressions: Iterable[object]) -> str:
    series_values: list[dict[int, float]] = []
    const_sum = 0.0
    for expr in expressions:
        mode, payload = parse_expression(expr)
        if mode == "empty":
            continue
        if mode == "const" and payload is not None:
            const_sum += float(payload)
        elif mode == "series" and isinstance(payload, dict):
            series_values.append(payload)
        else:
            return str(expr)

    if not series_values:
        return _format_number(const_sum)

    years = sorted({year for series in series_values for year in series})
    combined = {}
    for year in years:
        total = const_sum
        for series in series_values:
            total += series.get(year, 0.0)
        combined[year] = total
    return build_data_expression(combined)


def expression_to_series(
    expr: object,
    years: Iterable[int] | None = None,
    base_year: int | None = None,
) -> dict[int, float] | None:
    mode, payload = parse_expression(expr)
    if mode == "unknown":
        return None
    if mode == "empty":
        return {}
    if mode == "series" and isinstance(payload, dict):
        if years is None:
            return dict(payload)
        return {int(year): float(payload.get(int(year), 0.0)) for year in years}
    if mode == "const" and payload is not None:
        value = float(payload)
        if years is not None:
            return {int(year): value for year in years}
        if base_year is None:
            return {}
        return {int(base_year): value}
    return {}


def collect_expression_years(expressions: Iterable[object]) -> set[int]:
    years: set[int] = set()
    for expr in expressions:
        mode, payload = parse_expression(expr)
        if mode == "series" and isinstance(payload, dict):
            years.update(payload.keys())
    return years


__all__ = [
    "parse_expression",
    "build_data_expression",
    "build_data_expression_from_row",
    "scale_expression",
    "sum_expressions",
    "expression_to_series",
    "collect_expression_years",
]
