from __future__ import annotations

from collections import defaultdict, deque
import json
from typing import Any

import pandas as pd

from .models import AtomicSettings


def _normalize_text(value: object) -> str:
    text = str(value or "").strip()
    if text.lower() == "nan":
        return ""
    return text


def _canonical_scenario(value: object) -> str:
    text = _normalize_text(value).lower()
    if text == "reference":
        return "Reference"
    if text == "target":
        return "Target"
    return text.title() if text else ""


def _measure_is_input_only(measure: object) -> bool:
    text = _normalize_text(measure).lower()
    return bool(text and "input" in text and "output" not in text)


def _measure_is_output_only(measure: object) -> bool:
    text = _normalize_text(measure).lower()
    return bool(text and "output" in text and "input" not in text)


def _transformation_sign_role_from_measure(measure: object) -> str:
    if _measure_is_input_only(measure):
        return "input"
    if _measure_is_output_only(measure):
        return "output"
    return ""


def _sheet_is_export_flow(sheet: object, measure: object = "") -> bool:
    sheet_text = _normalize_text(sheet).lower()
    measure_text = _normalize_text(measure).lower()
    return sheet_text.startswith("exports") or "export" in measure_text


def _split_sector_codes(raw_value: object) -> list[str]:
    text = _normalize_text(raw_value)
    if not text:
        return []
    parts = pd.Series([text]).str.split(r"\s*(?:,|;|\||\band\b)\s*", regex=True).iloc[0]
    out: list[str] = []
    seen: set[str] = set()
    for part in parts:
        token = _normalize_text(part)
        if not token:
            continue
        key = token.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(token)
    return out


def _source_family(source: str) -> str:
    src = _normalize_text(source).lower()
    if src == "leap":
        return "leap"
    if src in {"base", "base_estimated", "base_mixed"}:
        return "base"
    if src in {"projection", "projection_estimated", "projection_mixed"}:
        return "projection"
    return src


def _parse_json_list(value: object) -> list[Any]:
    text = _normalize_text(value)
    if not text:
        return []
    try:
        parsed = json.loads(text)
    except Exception:
        return []
    return parsed if isinstance(parsed, list) else []


def _parse_base_targets_detail(value: object) -> list[tuple[str, str]]:
    targets: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for item in _parse_json_list(value):
        if isinstance(item, dict):
            flow = _normalize_text(item.get("esto_flow", ""))
            product = _normalize_text(item.get("esto_product", ""))
        elif isinstance(item, (list, tuple)) and len(item) >= 2:
            flow = _normalize_text(item[0])
            product = _normalize_text(item[1])
        else:
            continue
        if not flow and not product:
            continue
        key = (flow.lower(), product.lower())
        if key in seen:
            continue
        seen.add(key)
        targets.append((flow, product))
    return targets


def _parse_projection_fuel_codes_detail(value: object) -> list[str]:
    fuels: list[str] = []
    seen: set[str] = set()
    for item in _parse_json_list(value):
        token = _normalize_text(item)
        if not token:
            continue
        key = token.lower()
        if key in seen:
            continue
        seen.add(key)
        fuels.append(token)
    return fuels


def _parse_projection_targets_detail(value: object) -> list[tuple[str, str]]:
    targets: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for item in _parse_json_list(value):
        if isinstance(item, dict):
            sector = _normalize_text(item.get("sector_code_9th", ""))
            fuel = _normalize_text(item.get("ninth_fuel_code", ""))
        elif isinstance(item, (list, tuple)) and len(item) >= 2:
            sector = _normalize_text(item[0])
            fuel = _normalize_text(item[1])
        else:
            continue
        if not sector and not fuel:
            continue
        key = (sector.lower(), fuel.lower())
        if key in seen:
            continue
        seen.add(key)
        targets.append((sector, fuel))
    return targets


def _allocation_group_cols() -> list[str]:
    """
    Scope shared-unit allocation to the current measure view.

    Summary sheets and measure-specific sheets can intentionally reference the
    same raw comparator unit. Those are alternative dashboard views rather than
    sibling components, so they should not split one atomic unit across each
    other.
    """
    return ["source_family", "economy", "scenario", "year", "atomic_key", "measure"]


def resolve_comparison_level(
    comparison_long: pd.DataFrame,
    sheet_map: pd.DataFrame,
) -> pd.DataFrame:
    """
    Resolve chart comparison node by sheet using current final-category grouping behavior.

    If multiple sheets share the same final category label in sheet_map, resolve
    those sheet rows to the shared parent node.
    """
    if comparison_long.empty:
        return pd.DataFrame(columns=["sheet", "resolved_node_id", "resolved_node_level"])

    sheets = (
        comparison_long.get("sheet", pd.Series(dtype=str))
        .astype(str)
        .dropna()
        .drop_duplicates()
        .tolist()
    )
    out = pd.DataFrame({"sheet": sheets})
    out["resolved_node_id"] = out["sheet"]
    out["resolved_node_level"] = "sheet"

    if sheet_map is None or sheet_map.empty:
        return out

    sm = sheet_map.copy()
    for col in ["sheet_name", "sector_name", "final_category_name", "final category name"]:
        if col not in sm.columns:
            sm[col] = ""
        sm[col] = sm[col].fillna("").astype(str).str.strip()
    final_category = sm["final_category_name"].where(
        sm["final_category_name"].ne(""),
        sm["final category name"],
    )
    sm["grouped_category_name"] = final_category.where(
        final_category.ne(""),
        sm["sector_name"],
    )
    sm = sm[sm["sheet_name"].ne("") & sm["grouped_category_name"].ne("")]
    if sm.empty:
        return out

    grouped = (
        sm.groupby("grouped_category_name", as_index=False)["sheet_name"]
        .nunique()
        .rename(columns={"sheet_name": "sheet_count"})
    )
    grouped = grouped[grouped["sheet_count"] > 1]
    if grouped.empty:
        return out

    grouped_category_names = set(grouped["grouped_category_name"].astype(str).tolist())
    child_map = sm[sm["grouped_category_name"].isin(grouped_category_names)][["sheet_name", "grouped_category_name"]].drop_duplicates()
    child_lookup = {
        str(row.sheet_name): str(row.grouped_category_name)
        for row in child_map.itertuples(index=False)
    }

    out["resolved_node_id"] = out["sheet"].map(lambda s: child_lookup.get(str(s), str(s)))
    out["resolved_node_level"] = out["sheet"].map(
        lambda s: "sector_group_parent" if str(s) in child_lookup else "sheet"
    )
    return out


def _line_key(row: pd.Series) -> str:
    return "||".join(
        [
            str(row.get("economy", "")),
            str(row.get("sheet", "")),
            str(row.get("fuel_label", "")),
            str(row.get("scenario", "")),
            str(row.get("source", "")),
            str(int(row.get("year", 0))),
        ]
    )


def _canonical_targets_for_base_row(row: pd.Series, canonical_pairs: pd.DataFrame) -> list[tuple[str, str]]:
    explicit_targets = _parse_base_targets_detail(row.get("base_targets_detail", ""))
    if explicit_targets:
        return explicit_targets
    flow = _normalize_text(row.get("esto_flow", ""))
    product = _normalize_text(row.get("esto_product", ""))
    mapping_source = _normalize_text(row.get("mapping_source", "")).lower()
    sector_code = _normalize_text(row.get("sector_code_9th", ""))
    fuel_code = _normalize_text(row.get("ninth_fuel_code", ""))

    if mapping_source == "canonical_aggregated" and sector_code and fuel_code and not canonical_pairs.empty:
        cp = canonical_pairs.copy()
        for col in ["9th_sector", "9th_fuel", "esto_flow", "esto_product"]:
            if col not in cp.columns:
                cp[col] = ""
            cp[col] = cp[col].fillna("").astype(str).str.strip()
        subset = cp[
            cp["9th_sector"].eq(sector_code)
            & cp["9th_fuel"].eq(fuel_code)
            & cp["esto_flow"].ne("")
            & cp["esto_product"].ne("")
        ][["esto_flow", "esto_product"]].drop_duplicates()
        targets = [(str(r.esto_flow), str(r.esto_product)) for r in subset.itertuples(index=False)]
        if targets:
            return targets

    if flow and product:
        return [(flow, product)]
    if flow or product:
        return [(flow, product)]
    return []


def _prepare_line_rows(
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    resolved_levels: pd.DataFrame,
) -> pd.DataFrame:
    if comparison_long.empty:
        return pd.DataFrame()

    comp = comparison_long.copy()
    comp = comp[comp.get("fuel_label", "").astype(str).ne("Total")].copy()
    comp["value"] = pd.to_numeric(comp.get("value"), errors="coerce")
    comp["sheet"] = comp["sheet"].astype(str)
    comp["fuel_label"] = comp["fuel_label"].astype(str)
    comp["source"] = comp["source"].astype(str)
    comp["scenario"] = comp["scenario"].astype(str).map(_canonical_scenario)
    comp["year"] = pd.to_numeric(comp["year"], errors="coerce").fillna(0).astype(int)
    comp["source_family"] = comp["source"].map(_source_family)
    comp["line_key"] = comp.apply(_line_key, axis=1)

    ms = mapping_status.copy() if mapping_status is not None else pd.DataFrame()
    map_cols: list[str] = [
        "sheet",
        "fuel_label",
        "sector_code_9th",
        "ninth_fuel_code",
        "esto_flow",
        "esto_product",
        "mapping_source",
        "sector_match_method",
        "projection_parent_sector_code",
        "projection_fuel_codes_detail",
        "projection_targets_detail",
        "base_targets_detail",
    ]
    for col in map_cols:
        if col not in ms.columns:
            ms[col] = ""
        ms[col] = ms[col].fillna("").astype(str).str.strip()
    map_rows = ms[map_cols].drop_duplicates(subset=["sheet", "fuel_label"], keep="first")

    comp = comp.merge(map_rows, on=["sheet", "fuel_label"], how="left")
    comp = comp.merge(resolved_levels, on="sheet", how="left")
    comp["resolved_node_id"] = comp["resolved_node_id"].fillna(comp["sheet"]).astype(str)
    comp["resolved_node_level"] = comp["resolved_node_level"].fillna("sheet").astype(str)
    return comp


def _build_atomic_edge_candidates(
    *,
    line_rows: pd.DataFrame,
    canonical_pairs: pd.DataFrame,
) -> pd.DataFrame:
    if line_rows.empty:
        return pd.DataFrame(
            columns=[
                "line_key",
                "atomic_key",
                "source_family",
                "weight",
                "line_to_atomic_weight",
                "edge_reason",
                "economy",
                "scenario",
                "year",
                "sheet",
                "measure",
                "fuel_label",
                "source",
                "resolved_node_id",
                "resolved_node_level",
                "esto_flow",
                "esto_product",
                "sector_node",
                "fuel_node",
                "line_value",
                "atomic_value",
                "edge_contribution",
            ]
        )

    edge_rows: list[dict[str, Any]] = []
    for row in line_rows.itertuples(index=False):
        source_family = _normalize_text(row.source_family).lower()
        base_edge = {
            "line_key": str(row.line_key),
            "source_family": source_family,
            "economy": str(row.economy),
            "scenario": _canonical_scenario(row.scenario),
            "year": int(row.year),
            "sheet": str(row.sheet),
            "measure": str(getattr(row, "measure", "")),
            "fuel_label": str(row.fuel_label),
            "source": str(row.source),
            "resolved_node_id": str(row.resolved_node_id),
            "resolved_node_level": str(row.resolved_node_level),
            "esto_flow": "",
            "esto_product": "",
            "sector_node": "",
            "fuel_node": "",
            "line_value": float(row.value) if pd.notna(row.value) else float("nan"),
        }

        if source_family == "leap":
            atomic_key = (
                f"leap|{row.economy}|{_canonical_scenario(row.scenario)}|{int(row.year)}|"
                f"{row.sheet}|{row.fuel_label}"
            )
            edge_rows.append(
                {
                    **base_edge,
                    "atomic_key": atomic_key,
                    "line_to_atomic_weight": 1.0,
                    "edge_reason": "direct",
                }
            )
            continue

        if source_family == "base":
            targets = _canonical_targets_for_base_row(pd.Series(row._asdict()), canonical_pairs)
            if targets:
                w = 1.0 / float(len(targets))
                reason = "direct" if len(targets) == 1 else "equal_split_allocation"
                for flow, prod in targets:
                    atomic_key = (
                        f"base|{row.economy}|{_canonical_scenario(row.scenario)}|{int(row.year)}|"
                        f"{flow}|{prod}"
                    )
                    edge_rows.append(
                        {
                            **base_edge,
                            "atomic_key": atomic_key,
                            "line_to_atomic_weight": w,
                            "edge_reason": reason,
                            "esto_flow": str(flow),
                            "esto_product": str(prod),
                        }
                    )
            else:
                atomic_key = (
                    f"base|{row.economy}|{_canonical_scenario(row.scenario)}|{int(row.year)}|"
                    f"__unmapped__|{row.sheet}|{row.fuel_label}"
                )
                edge_rows.append(
                    {
                        **base_edge,
                        "atomic_key": atomic_key,
                        "line_to_atomic_weight": 1.0,
                        "edge_reason": "unresolved_mapping",
                        "esto_flow": _normalize_text(row.esto_flow),
                        "esto_product": _normalize_text(row.esto_product),
                    }
                )
            continue

        if source_family == "projection":
            explicit_projection_targets = _parse_projection_targets_detail(
                getattr(row, "projection_targets_detail", "")
            )
            if explicit_projection_targets:
                combos = explicit_projection_targets
            else:
                sector_codes = _split_sector_codes(_normalize_text(row.sector_code_9th))
                if not sector_codes:
                    sector_codes = _split_sector_codes(_normalize_text(row.projection_parent_sector_code))
                fuel_nodes = _parse_projection_fuel_codes_detail(getattr(row, "projection_fuel_codes_detail", ""))
                if not fuel_nodes:
                    fuel_node = _normalize_text(row.ninth_fuel_code)
                    fuel_nodes = [fuel_node] if fuel_node else []
                combos = [(sector, fuel_node) for sector in sector_codes for fuel_node in fuel_nodes]
            if combos:
                w = 1.0 / float(len(combos))
                reason = "direct" if len(combos) == 1 else "equal_split_allocation"
                for sector, fuel_node in combos:
                    atomic_key = (
                        f"projection|{row.economy}|{_canonical_scenario(row.scenario)}|{int(row.year)}|"
                        f"{sector}|{fuel_node}"
                    )
                    edge_rows.append(
                        {
                            **base_edge,
                            "atomic_key": atomic_key,
                            "line_to_atomic_weight": w,
                            "edge_reason": reason,
                            "sector_node": str(sector),
                            "fuel_node": fuel_node,
                        }
                    )
            else:
                atomic_key = (
                    f"projection|{row.economy}|{_canonical_scenario(row.scenario)}|{int(row.year)}|"
                    f"__unmapped__|{row.sheet}|{row.fuel_label}"
                )
                edge_rows.append(
                    {
                        **base_edge,
                        "atomic_key": atomic_key,
                        "line_to_atomic_weight": 1.0,
                        "edge_reason": "unresolved_mapping",
                        "sector_node": _normalize_text(row.sector_code_9th),
                        "fuel_node": _normalize_text(row.ninth_fuel_code),
                    }
                )
            continue

        atomic_key = (
            f"{source_family}|{row.economy}|{_canonical_scenario(row.scenario)}|{int(row.year)}|"
            f"{row.sheet}|{row.fuel_label}"
        )
        edge_rows.append(
            {
                **base_edge,
                "atomic_key": atomic_key,
                "line_to_atomic_weight": 1.0,
                "edge_reason": "direct",
            }
        )

    edges = pd.DataFrame(edge_rows)
    if edges.empty:
        return edges

    edges = edges.drop_duplicates(subset=["line_key", "atomic_key"], keep="first").reset_index(drop=True)
    edges["line_to_atomic_weight"] = pd.to_numeric(edges["line_to_atomic_weight"], errors="coerce").fillna(0.0)
    edges["weight"] = pd.NA
    edges["line_value"] = pd.to_numeric(edges["line_value"], errors="coerce")
    edges["atomic_value"] = pd.NA
    edges["edge_contribution"] = pd.NA
    return edges


def _year_columns(df: pd.DataFrame, years: set[int]) -> list[str]:
    out: list[str] = []
    for col in df.columns:
        token = str(col).strip()
        if token.isdigit() and int(token) in years:
            out.append(token)
    return out


def _base_value_lookup(base_df: pd.DataFrame, *, base_economy: str, years: set[int]) -> dict[tuple[int, str, str, str], float]:
    if base_df is None or base_df.empty or not years:
        return {}

    df = base_df.copy()
    for col in ["economy", "flows", "products"]:
        if col not in df.columns:
            return {}
        df[col] = df[col].fillna("").astype(str).str.strip()

    if base_economy:
        df = df[df["economy"].eq(str(base_economy).strip())]
    if df.empty:
        return {}

    year_cols = _year_columns(df, years)
    if not year_cols:
        return {}

    long_df = df.melt(
        id_vars=["flows", "products"],
        value_vars=year_cols,
        var_name="year",
        value_name="value",
    )
    long_df["year"] = pd.to_numeric(long_df["year"], errors="coerce").astype("Int64")
    long_df["value"] = pd.to_numeric(long_df["value"], errors="coerce")
    long_df = long_df.dropna(subset=["year"])
    long_df["sign_role"] = ""
    is_transformation = long_df["flows"].astype(str).str.strip().str.lower().str.startswith("09")
    long_df.loc[is_transformation & long_df["value"].lt(0), "sign_role"] = "input"
    long_df.loc[is_transformation & long_df["value"].gt(0), "sign_role"] = "output"

    raw_df = (
        long_df.groupby(["year", "flows", "products"], as_index=False)["value"]
        .sum(min_count=1)
        .assign(sign_role="")
    )
    directional_df = long_df[long_df["sign_role"].ne("")].groupby(
        ["year", "flows", "products", "sign_role"], as_index=False
    )["value"].sum(min_count=1)
    value_df = pd.concat([raw_df, directional_df], ignore_index=True, sort=False)
    return {
        (int(r.year), str(r.flows), str(r.products), str(r.sign_role)): float(r.value) if pd.notna(r.value) else float("nan")
        for r in value_df.itertuples(index=False)
    }


def _projection_value_lookup(
    ninth_df: pd.DataFrame,
    *,
    projection_economy: str,
    years: set[int],
) -> dict[tuple[str, int, str, str, str], float]:
    if ninth_df is None or ninth_df.empty or not years:
        return {}

    df = ninth_df.copy()
    for col in ["economy", "scenarios", "sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors", "fuels", "subfuels"]:
        if col not in df.columns:
            return {}
        df[col] = df[col].fillna("").astype(str).str.strip()

    if projection_economy:
        df = df[df["economy"].eq(str(projection_economy).strip())]
    if df.empty:
        return {}

    df["scenario"] = df["scenarios"].map(_canonical_scenario)
    df["sector_node"] = df["sectors"]
    for col in ["sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]:
        valid = df[col].ne("") & ~df[col].str.lower().eq("x")
        df.loc[valid, "sector_node"] = df.loc[valid, col]
    df["fuel_node"] = df["subfuels"]
    fuel_fallback = df["fuel_node"].eq("") | df["fuel_node"].str.lower().eq("x")
    df.loc[fuel_fallback, "fuel_node"] = df.loc[fuel_fallback, "fuels"]

    year_cols = _year_columns(df, years)
    if not year_cols:
        return {}

    long_df = df.melt(
        id_vars=["scenario", "sector_node", "fuel_node"],
        value_vars=year_cols,
        var_name="year",
        value_name="value",
    )
    long_df["year"] = pd.to_numeric(long_df["year"], errors="coerce").astype("Int64")
    long_df["value"] = pd.to_numeric(long_df["value"], errors="coerce")
    long_df = long_df.dropna(subset=["year"])
    long_df["sign_role"] = ""
    is_transformation = long_df["sector_node"].astype(str).str.strip().str.lower().str.startswith("09_")
    long_df.loc[is_transformation & long_df["value"].lt(0), "sign_role"] = "input"
    long_df.loc[is_transformation & long_df["value"].gt(0), "sign_role"] = "output"

    raw_df = (
        long_df.groupby(["scenario", "year", "sector_node", "fuel_node"], as_index=False)["value"]
        .sum(min_count=1)
        .assign(sign_role="")
    )
    directional_df = long_df[long_df["sign_role"].ne("")].groupby(
        ["scenario", "year", "sector_node", "fuel_node", "sign_role"], as_index=False
    )["value"].sum(min_count=1)
    value_df = pd.concat([raw_df, directional_df], ignore_index=True, sort=False)
    return {
        (str(r.scenario), int(r.year), str(r.sector_node), str(r.fuel_node), str(r.sign_role)): float(r.value) if pd.notna(r.value) else float("nan")
        for r in value_df.itertuples(index=False)
    }


def _leap_value_lookup(leap_long: pd.DataFrame) -> dict[tuple[str, int, str, str], float]:
    if leap_long is None or leap_long.empty:
        return {}

    df = leap_long.copy()
    if "sheet" not in df.columns:
        if "sheet_name" in df.columns:
            df["sheet"] = df["sheet_name"]
        else:
            df["sheet"] = ""
    if "value" not in df.columns:
        if "leap_value" in df.columns:
            df["value"] = df["leap_value"]
        else:
            df["value"] = pd.NA
    if "scenario" not in df.columns:
        df["scenario"] = ""
    if "fuel_label" not in df.columns:
        df["fuel_label"] = ""
    if "year" not in df.columns:
        return {}

    df["scenario"] = df["scenario"].map(_canonical_scenario)
    df["year"] = pd.to_numeric(df["year"], errors="coerce").astype("Int64")
    df["sheet"] = df["sheet"].fillna("").astype(str).str.strip()
    df["fuel_label"] = df["fuel_label"].fillna("").astype(str).str.strip()
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    df = df.dropna(subset=["year"])
    df = (
        df.groupby(["scenario", "year", "sheet", "fuel_label"], as_index=False)["value"]
        .sum(min_count=1)
    )
    return {
        (str(r.scenario), int(r.year), str(r.sheet), str(r.fuel_label)): float(r.value) if pd.notna(r.value) else float("nan")
        for r in df.itertuples(index=False)
    }


def _build_atomic_units_from_raw(
    *,
    edges: pd.DataFrame,
    base_df: pd.DataFrame,
    ninth_df: pd.DataFrame,
    leap_long: pd.DataFrame,
    base_economy: str,
    projection_economy: str,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if edges.empty:
        empty_base = pd.DataFrame(columns=["economy", "scenario", "year", "esto_flow", "esto_product", "value"])
        empty_proj = pd.DataFrame(columns=["economy", "scenario", "year", "sector_node", "fuel_node", "value"])
        empty_leap = pd.DataFrame(columns=["economy", "scenario", "year", "sheet", "fuel_label", "value"])
        empty_lookup = pd.DataFrame(columns=["source_family", "economy", "scenario", "year", "atomic_key", "atomic_value"])
        return empty_base, empty_proj, empty_leap, empty_lookup

    required_years = set(pd.to_numeric(edges["year"], errors="coerce").dropna().astype(int).tolist())
    base_lookup = _base_value_lookup(base_df, base_economy=base_economy, years=required_years)
    projection_lookup = _projection_value_lookup(
        ninth_df,
        projection_economy=projection_economy,
        years=required_years,
    )
    leap_lookup = _leap_value_lookup(leap_long)

    base_edges = edges[edges["source_family"].eq("base")][
        ["economy", "scenario", "year", "sheet", "measure", "fuel_label", "esto_flow", "esto_product", "atomic_key"]
    ].drop_duplicates()
    if not base_edges.empty:
        base_edges["value"] = base_edges.apply(
            lambda r: base_lookup.get(
                (
                    int(r["year"]),
                    _normalize_text(r["esto_flow"]),
                    _normalize_text(r["esto_product"]),
                    _transformation_sign_role_from_measure(r.get("measure", "")),
                ),
                float("nan"),
            ),
            axis=1,
        )
    else:
        base_edges["value"] = pd.Series(dtype=float)
    base_units = base_edges[["economy", "scenario", "year", "sheet", "measure", "fuel_label", "esto_flow", "esto_product", "value"]].copy()

    proj_edges = edges[edges["source_family"].eq("projection")][
        ["economy", "scenario", "year", "sheet", "measure", "fuel_label", "sector_node", "fuel_node", "atomic_key"]
    ].drop_duplicates()
    if not proj_edges.empty:
        proj_edges["value"] = proj_edges.apply(
            lambda r: projection_lookup.get(
                (
                    _canonical_scenario(r["scenario"]),
                    int(r["year"]),
                    _normalize_text(r["sector_node"]),
                    _normalize_text(r["fuel_node"]),
                    _transformation_sign_role_from_measure(r.get("measure", "")),
                ),
                float("nan"),
            ),
            axis=1,
        )
    else:
        proj_edges["value"] = pd.Series(dtype=float)
    projection_units = proj_edges[["economy", "scenario", "year", "sheet", "measure", "fuel_label", "sector_node", "fuel_node", "value"]].copy()

    leap_edges = edges[edges["source_family"].eq("leap")][
        ["economy", "scenario", "year", "sheet", "fuel_label", "atomic_key", "line_value"]
    ].drop_duplicates()
    if not leap_edges.empty:
        leap_edges["value"] = leap_edges.apply(
            lambda r: leap_lookup.get(
                (
                    _canonical_scenario(r["scenario"]),
                    int(r["year"]),
                    _normalize_text(r["sheet"]),
                    _normalize_text(r["fuel_label"]),
                ),
                float("nan"),
            ),
            axis=1,
        )
        leap_edges["value"] = pd.to_numeric(leap_edges["value"], errors="coerce").where(
            pd.to_numeric(leap_edges["value"], errors="coerce").notna(),
            pd.to_numeric(leap_edges["line_value"], errors="coerce"),
        )
    else:
        leap_edges["value"] = pd.Series(dtype=float)
    leap_units = leap_edges[["economy", "scenario", "year", "sheet", "fuel_label", "value"]].copy()

    unit_lookup = pd.concat(
        [
            base_edges.assign(source_family="base").rename(columns={"value": "atomic_value"})[
                ["source_family", "economy", "scenario", "year", "atomic_key", "atomic_value"]
            ],
            proj_edges.assign(source_family="projection").rename(columns={"value": "atomic_value"})[
                ["source_family", "economy", "scenario", "year", "atomic_key", "atomic_value"]
            ],
            leap_edges.assign(source_family="leap").rename(columns={"value": "atomic_value"})[
                ["source_family", "economy", "scenario", "year", "atomic_key", "atomic_value"]
            ],
        ],
        ignore_index=True,
        sort=False,
    )

    return base_units, projection_units, leap_units, unit_lookup


def _connected_components(lines: list[str], atomics: list[str], pairs: list[tuple[str, str]]) -> list[dict[str, set[str]]]:
    line_adj: dict[str, set[str]] = defaultdict(set)
    atomic_adj: dict[str, set[str]] = defaultdict(set)
    for ln, at in pairs:
        line_adj[ln].add(at)
        atomic_adj[at].add(ln)

    all_line = set(lines)
    all_atomic = set(atomics)
    seen_line: set[str] = set()
    seen_atomic: set[str] = set()
    out: list[dict[str, set[str]]] = []

    for start in all_line:
        if start in seen_line:
            continue
        q = deque([("line", start)])
        comp_lines: set[str] = set()
        comp_atomics: set[str] = set()
        while q:
            kind, node = q.popleft()
            if kind == "line":
                if node in seen_line:
                    continue
                seen_line.add(node)
                comp_lines.add(node)
                for nxt in line_adj.get(node, set()):
                    if nxt not in seen_atomic:
                        q.append(("atomic", nxt))
            else:
                if node in seen_atomic:
                    continue
                seen_atomic.add(node)
                comp_atomics.add(node)
                for nxt in atomic_adj.get(node, set()):
                    if nxt not in seen_line:
                        q.append(("line", nxt))
        out.append({"lines": comp_lines, "atomics": comp_atomics})

    for start in all_atomic:
        if start in seen_atomic:
            continue
        out.append({"lines": set(), "atomics": {start}})
        seen_atomic.add(start)

    return out


def find_unresolved_many_to_many_components(edges: pd.DataFrame) -> pd.DataFrame:
    if edges.empty:
        return pd.DataFrame(
            columns=[
                "sheet",
                "scenario",
                "source_family",
                "resolved_node_id",
                "component_id",
                "line_count",
                "atomic_count",
                "all_deterministic",
                "edge_reasons",
                "sample_line_keys",
                "sample_atomic_keys",
            ]
        )

    deterministic = {
        "direct",
        "line_value_share_allocation",
        "base_share_allocation",
        "equal_split_allocation",
    }
    rows: list[dict[str, Any]] = []
    grp_cols = ["sheet", "scenario", "source_family", "resolved_node_id"]
    for keys, g in edges.groupby(grp_cols, dropna=False):
        sheet, scenario, source_family, resolved_node_id = keys
        group_reasons = set(g["edge_reason"].fillna("").astype(str).tolist())
        if group_reasons and group_reasons.issubset(deterministic):
            continue
        pairs = list(zip(g["line_key"].astype(str), g["atomic_key"].astype(str)))
        comps = _connected_components(
            lines=g["line_key"].astype(str).tolist(),
            atomics=g["atomic_key"].astype(str).tolist(),
            pairs=pairs,
        )
        for cid, comp in enumerate(comps, start=1):
            line_keys = sorted(comp["lines"])
            atomic_keys = sorted(comp["atomics"])
            if len(line_keys) <= 1 or len(atomic_keys) <= 1:
                continue
            sub = g[
                g["line_key"].astype(str).isin(line_keys)
                & g["atomic_key"].astype(str).isin(atomic_keys)
            ]
            reasons = sorted(set(sub["edge_reason"].fillna("").astype(str).tolist()))
            all_deterministic = all(reason in deterministic for reason in reasons)
            if all_deterministic:
                continue
            rows.append(
                {
                    "sheet": sheet,
                    "scenario": scenario,
                    "source_family": source_family,
                    "resolved_node_id": resolved_node_id,
                    "component_id": cid,
                    "line_count": len(line_keys),
                    "atomic_count": len(atomic_keys),
                    "all_deterministic": all_deterministic,
                    "edge_reasons": " | ".join(reasons),
                    "sample_line_keys": " | ".join(line_keys[:5]),
                    "sample_atomic_keys": " | ".join(atomic_keys[:5]),
                }
            )
    return pd.DataFrame(rows)


def _initial_weight_allocation(edges: pd.DataFrame) -> pd.DataFrame:
    if edges.empty:
        return edges

    out = edges.copy()
    group_cols = _allocation_group_cols()
    out["weight"] = pd.to_numeric(out["weight"], errors="coerce")
    for _, g in out.groupby(group_cols, dropna=False):
        idx = g.index
        line_count = g["line_key"].astype(str).nunique()
        if line_count <= 1:
            out.loc[idx, "weight"] = 1.0
            continue

        unresolved = g["edge_reason"].fillna("").astype(str).eq("unresolved_mapping").any()
        n = len(idx)
        out.loc[idx, "weight"] = 1.0 / float(n) if n else 0.0
        if not unresolved:
            out.loc[idx, "edge_reason"] = "equal_split_allocation"
    return out


def _build_base_line_lookup(edges: pd.DataFrame) -> dict[tuple[str, str, str], float]:
    if edges.empty:
        return {}

    base = edges[edges["source_family"].eq("base")].copy()
    if base.empty:
        return {}
    base["edge_contribution"] = pd.to_numeric(base["edge_contribution"], errors="coerce")
    vals = (
        base.groupby(["sheet", "measure", "fuel_label", "scenario"], as_index=False)["edge_contribution"]
        .sum(min_count=1)
    )
    return {
        (str(r.sheet), str(r.measure), str(r.fuel_label), str(r.scenario)): (
            float(r.edge_contribution) if pd.notna(r.edge_contribution) else float("nan")
        )
        for r in vals.itertuples(index=False)
    }


def _apply_projection_base_share_weights(
    edges: pd.DataFrame,
    *,
    base_line_lookup: dict[tuple[str, str, str, str], float],
) -> pd.DataFrame:
    if edges.empty:
        return edges

    out = edges.copy()
    group_cols = _allocation_group_cols()
    for keys, g in out.groupby(group_cols, dropna=False):
        source_family = str(keys[0] or "")
        if source_family != "projection":
            continue
        idx = g.index
        line_count = g["line_key"].astype(str).nunique()
        if line_count <= 1:
            out.loc[idx, "weight"] = 1.0
            continue

        unresolved = g["edge_reason"].fillna("").astype(str).eq("unresolved_mapping").any()
        if unresolved:
            continue

        shares = []
        for r in g.itertuples(index=False):
            key = (str(r.sheet), str(r.measure), str(r.fuel_label), str(r.scenario))
            val = base_line_lookup.get(key, float("nan"))
            num = abs(float(val)) if pd.notna(val) else 0.0
            shares.append(num)
        denom = float(sum(shares))
        if denom > 0:
            out.loc[idx, "weight"] = [s / denom for s in shares]
            out.loc[idx, "edge_reason"] = "base_share_allocation"
        else:
            n = len(idx)
            out.loc[idx, "weight"] = 1.0 / float(n) if n else 0.0
            out.loc[idx, "edge_reason"] = "equal_split_allocation"
    return out


def _apply_atomic_values(edges: pd.DataFrame, unit_lookup: pd.DataFrame) -> pd.DataFrame:
    if edges.empty:
        return edges

    out = edges.copy()
    lookup_cols = ["source_family", "economy", "scenario", "year", "atomic_key", "atomic_value"]
    join_keys = ["source_family", "economy", "scenario", "year", "atomic_key"]
    values = unit_lookup[lookup_cols].drop_duplicates(subset=join_keys, keep="first") if not unit_lookup.empty else pd.DataFrame(columns=lookup_cols)
    out = out.drop(columns=["atomic_value", "edge_contribution"], errors="ignore")
    out = out.merge(values, on=join_keys, how="left")
    out["weight"] = pd.to_numeric(out["weight"], errors="coerce")
    out["atomic_value"] = pd.to_numeric(out["atomic_value"], errors="coerce")
    out["edge_contribution"] = out["atomic_value"] * out["weight"]
    return out


def _build_atomic_comparison_long(comparison_long: pd.DataFrame, edges: pd.DataFrame) -> pd.DataFrame:
    if comparison_long.empty:
        return comparison_long.copy()

    comp = comparison_long.copy()
    comp["source"] = comp["source"].astype(str)
    comp["scenario"] = comp["scenario"].astype(str).map(_canonical_scenario)
    comp["year"] = pd.to_numeric(comp["year"], errors="coerce").fillna(0).astype(int)
    comp["line_key"] = comp.apply(_line_key, axis=1)
    comp["__legacy_value"] = pd.to_numeric(comp["value"], errors="coerce")

    if edges.empty:
        out = comp.drop(columns=["line_key"], errors="ignore")
        out["value"] = out["__legacy_value"]
        return out.drop(columns=["__legacy_value"], errors="ignore")

    line_vals = (
        edges.groupby("line_key", as_index=False)["edge_contribution"]
        .sum(min_count=1)
        .rename(columns={"edge_contribution": "atomic_value"})
    )
    edge_work = edges.copy()
    shared_cols = _allocation_group_cols()
    edge_work["shared_line_count"] = edge_work.groupby(shared_cols)["line_key"].transform("nunique")
    edge_work["line_allocated"] = (
        edge_work["shared_line_count"].gt(1)
        & pd.to_numeric(edge_work["atomic_value"], errors="coerce").notna()
    )
    line_flags = (
        edge_work.groupby("line_key", as_index=False)
        .agg(
            line_has_atomic_value=("atomic_value", lambda s: pd.to_numeric(s, errors="coerce").notna().any()),
            line_all_unresolved=("edge_reason", lambda s: s.fillna("").astype(str).eq("unresolved_mapping").all()),
            atomic_allocated=("line_allocated", "max"),
        )
    )
    out = comp.merge(line_vals, on="line_key", how="left")
    out = out.merge(line_flags, on="line_key", how="left")
    out["value"] = pd.to_numeric(out["atomic_value"], errors="coerce")

    # When a line has only unresolved mappings, retain legacy values so shadow
    # comparisons can focus on resolvable mapping differences.
    unresolved_fallback = (
        out["value"].isna()
        & out["line_all_unresolved"].fillna(False)
        & ~out["line_has_atomic_value"].fillna(False)
    )
    out.loc[unresolved_fallback, "value"] = out.loc[unresolved_fallback, "__legacy_value"]

    # Totals are recomputed later for chart rendering; keep existing values here.
    total_mask = out["fuel_label"].astype(str).eq("Total")
    out.loc[total_mask, "value"] = out.loc[total_mask, "__legacy_value"]

    # Match dashboard display convention: comparator values on input-only
    # sheets should be shown as positive magnitudes even though the raw
    # transformation tables store them as negative inputs.
    if "measure" in out.columns:
        input_only = out["measure"].map(_measure_is_input_only)
        export_only = out.apply(lambda r: _sheet_is_export_flow(r.get("sheet", ""), r.get("measure", "")), axis=1)
        comparator_rows = out["source"].astype(str).isin(
            ["base", "base_estimated", "base_mixed", "projection", "projection_estimated", "projection_mixed"]
        )
        display_positive = (input_only | export_only) & comparator_rows
        out.loc[display_positive, "value"] = pd.to_numeric(
            out.loc[display_positive, "value"],
            errors="coerce",
        ).abs()

    out = out.drop(
        columns=["line_key", "atomic_value", "__legacy_value", "line_has_atomic_value", "line_all_unresolved"],
        errors="ignore",
    )
    return out


def _build_atomic_comparison_wide(comparison_long: pd.DataFrame) -> pd.DataFrame:
    if comparison_long.empty:
        return pd.DataFrame(
            columns=[
                "economy",
                "scenario",
                "sheet",
                "fuel_label",
                "year",
                "leap_value",
                "base_value",
                "projection_value",
            ]
        )

    comp = comparison_long.copy()
    comp["value"] = pd.to_numeric(comp["value"], errors="coerce")
    piv = (
        comp.pivot_table(
            index=["economy", "scenario", "sheet", "fuel_label", "year"],
            columns="source",
            values="value",
            aggfunc="first",
        )
        .reset_index()
    )

    def _coalesce(df: pd.DataFrame, cols: list[str]) -> pd.Series:
        existing = [c for c in cols if c in df.columns]
        if not existing:
            return pd.Series(float("nan"), index=df.index)
        out = pd.to_numeric(df[existing[0]], errors="coerce")
        for col in existing[1:]:
            out = out.combine_first(pd.to_numeric(df[col], errors="coerce"))
        return out

    piv["leap_value"] = _coalesce(piv, ["leap"])
    piv["base_value"] = _coalesce(piv, ["base", "base_estimated", "base_mixed"])
    piv["projection_value"] = _coalesce(piv, ["projection", "projection_estimated", "projection_mixed"])
    cols = ["economy", "scenario", "sheet", "fuel_label", "year", "leap_value", "base_value", "projection_value"]
    return piv[cols].copy()


def build_atomic_outputs(
    *,
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    sheet_map: pd.DataFrame,
    canonical_pairs: pd.DataFrame,
    base_df: pd.DataFrame,
    ninth_df: pd.DataFrame,
    leap_long: pd.DataFrame,
    base_economy: str,
    projection_economy: str,
    settings: AtomicSettings,
) -> dict[str, pd.DataFrame]:
    resolved_levels = resolve_comparison_level(comparison_long, sheet_map)
    line_rows = _prepare_line_rows(comparison_long, mapping_status, resolved_levels)
    edges = _build_atomic_edge_candidates(
        line_rows=line_rows,
        canonical_pairs=canonical_pairs if canonical_pairs is not None else pd.DataFrame(),
    )

    base_units, projection_units, leap_units, unit_lookup = _build_atomic_units_from_raw(
        edges=edges,
        base_df=base_df if base_df is not None else pd.DataFrame(),
        ninth_df=ninth_df if ninth_df is not None else pd.DataFrame(),
        leap_long=leap_long if leap_long is not None else pd.DataFrame(),
        base_economy=str(base_economy or "").strip(),
        projection_economy=str(projection_economy or "").strip(),
    )

    edges = _initial_weight_allocation(edges)
    edges = _apply_atomic_values(edges, unit_lookup)
    base_line_lookup = _build_base_line_lookup(edges)
    edges = _apply_projection_base_share_weights(edges, base_line_lookup=base_line_lookup)
    edges = _apply_atomic_values(edges, unit_lookup)

    many_to_many_errors = find_unresolved_many_to_many_components(edges)
    atomic_long = _build_atomic_comparison_long(comparison_long, edges)
    atomic_wide = _build_atomic_comparison_wide(atomic_long)

    validation_rows = [
        {"metric": "atomic_edges", "value": float(len(edges)), "details": ""},
        {"metric": "atomic_base_units", "value": float(len(base_units)), "details": ""},
        {"metric": "atomic_projection_units", "value": float(len(projection_units)), "details": ""},
        {"metric": "atomic_leap_units", "value": float(len(leap_units)), "details": ""},
        {"metric": "unresolved_many_to_many_components", "value": float(len(many_to_many_errors)), "details": ""},
    ]
    if not edges.empty:
        grp_cols = ["source_family", "economy", "scenario", "year", "atomic_key"]
        weight_chk = edges.groupby(grp_cols, as_index=False)["weight"].sum(min_count=1)
        bad_weight = weight_chk[~weight_chk["weight"].round(8).eq(1.0)]
        missing_unit_edges = edges["atomic_value"].isna().sum()
        missing_line_values = (
            atomic_long[atomic_long["fuel_label"].astype(str).ne("Total")]["value"]
            .isna()
            .sum()
        )
        validation_rows.append(
            {"metric": "atomic_weight_sum_violations", "value": float(len(bad_weight)), "details": ""}
        )
        validation_rows.append(
            {"metric": "atomic_edges_missing_unit_values", "value": float(missing_unit_edges), "details": ""}
        )
        validation_rows.append(
            {"metric": "atomic_lines_missing_values", "value": float(missing_line_values), "details": ""}
        )
    atomic_validation_report = pd.DataFrame(validation_rows)

    return {
        "atomic_base_units": base_units,
        "atomic_projection_units": projection_units,
        "atomic_leap_units": leap_units,
        "atomic_mapping_edges": edges,
        "atomic_comparison_long": atomic_long,
        "atomic_comparison_wide": atomic_wide,
        "atomic_validation_report": atomic_validation_report,
        "atomic_many_to_many_errors": many_to_many_errors,
        "resolved_level_lookup": resolved_levels,
    }


def build_shadow_delta_reports(
    *,
    legacy_chart_input: pd.DataFrame,
    atomic_chart_input: pd.DataFrame,
) -> dict[str, pd.DataFrame]:
    keys = ["economy", "sheet", "fuel_label", "scenario", "source", "year"]
    legacy = legacy_chart_input.copy()
    atomic = atomic_chart_input.copy()
    for frame in [legacy, atomic]:
        frame["value"] = pd.to_numeric(frame.get("value"), errors="coerce")
        for col in keys:
            if col not in frame.columns:
                frame[col] = ""
    legacy = legacy[keys + ["value"]].rename(columns={"value": "value_legacy"})
    atomic = atomic[keys + ["value"]].rename(columns={"value": "value_atomic"})

    series = legacy.merge(atomic, on=keys, how="outer")
    series["delta"] = pd.to_numeric(series["value_atomic"], errors="coerce") - pd.to_numeric(
        series["value_legacy"], errors="coerce"
    )
    series["abs_delta"] = series["delta"].abs()

    totals = series[series["fuel_label"].astype(str).eq("Total")].copy()
    summary = (
        series.groupby(["source", "sheet"], as_index=False)
        .agg(
            rows=("delta", "size"),
            max_abs_delta=("abs_delta", "max"),
            p95_abs_delta=("abs_delta", lambda s: pd.to_numeric(s, errors="coerce").quantile(0.95)),
            sum_abs_delta=("abs_delta", "sum"),
        )
        .sort_values(["sum_abs_delta", "max_abs_delta"], ascending=False)
        .reset_index(drop=True)
    )

    return {
        "atomic_shadow_delta_series": series,
        "atomic_shadow_delta_totals": totals,
        "atomic_shadow_delta_summary": summary,
    }
