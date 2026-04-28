from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd

from codebase.functions.leap_results_functions import (
    connect_leap,
    ensure_calculated,
    fetch_values_rs,
    list_dimensions,
)


RAW_RESULT_COLUMNS = [
    "scenario",
    "region",
    "year",
    "sector_group",
    "module",
    "branch",
    "process",
    "fuel",
    "variable",
    "dimension_name",
    "member_name",
    "filter_str",
    "value",
    "unit",
    "notes",
]

NORMALIZED_COLUMNS = [
    "scenario",
    "region",
    "year",
    "sector_group",
    "module",
    "branch",
    "process",
    "fuel",
    "metric",
    "value",
    "unit",
    "source_variable",
    "source_filter",
    "notes",
    "qa_negative_conversion_loss",
    "qa_aux_exceeds_gross_output",
    "qa_feedstock_missing_but_output_present",
    "qa_zero_output_with_positive_aux",
    "qa_balance_gap",
]

DASHBOARD_COLUMNS = [
    "economy",
    "scenario",
    "region",
    "sheet_name",
    "sector_code_9th",
    "sector_name",
    "fuel_label",
    "year",
    "leap_value",
    "leap_variable",
    "leap_units",
    "measure",
    "leap_scale_note",
]

QA_COLUMNS = [
    "scenario",
    "region",
    "year",
    "sector_group",
    "module",
    "process",
    "fuel",
    "metric",
    "qa_flag",
    "qa_value",
    "details",
]

DISCOVERY_COLUMNS = ["dimension_name", "member_name"]


def empty_raw_result_pulls() -> pd.DataFrame:
    return pd.DataFrame(columns=RAW_RESULT_COLUMNS)


def empty_normalized_long() -> pd.DataFrame:
    return pd.DataFrame(columns=NORMALIZED_COLUMNS)


def empty_dashboard_leap_long() -> pd.DataFrame:
    return pd.DataFrame(columns=DASHBOARD_COLUMNS)


def empty_qa_flags() -> pd.DataFrame:
    return pd.DataFrame(columns=QA_COLUMNS)


def empty_dimension_discovery() -> pd.DataFrame:
    return pd.DataFrame(columns=DISCOVERY_COLUMNS)


def scale_raw_energy_values(raw_df: pd.DataFrame) -> pd.DataFrame:
    if raw_df.empty:
        return raw_df
    out = raw_df.copy()
    out["value"] = pd.to_numeric(out["value"], errors="coerce")
    units = out.get("unit")
    if units is None:
        return out
    normalized_units = units.fillna("").astype(str).str.strip().str.lower()
    max_abs = out["value"].abs().max()
    if pd.isna(max_abs) or float(max_abs) < 1e6:
        return out
    scale_mask = normalized_units.eq("petajoules")
    if scale_mask.any():
        out.loc[scale_mask, "value"] = pd.to_numeric(out.loc[scale_mask, "value"], errors="coerce") * 1e-6
    return out


def _extract_member_name(member_obj: object) -> str:
    for attr in ("Name", "name", "Label", "label"):
        try:
            value = getattr(member_obj, attr)
        except Exception:
            continue
        text = str(value or "").strip()
        if text:
            return text
    return str(member_obj or "").strip()


def discover_result_dimensions(app) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    try:
        dims = getattr(app, "Dimensions")
        dim_count = int(getattr(dims, "Count", 0))
    except Exception:
        return empty_dimension_discovery()

    for i in range(1, dim_count + 1):
        try:
            dim = dims.Item(i)
            dim_name = str(getattr(dim, "Name", "") or "").strip()
        except Exception:
            continue
        if not dim_name:
            continue
        member_names: list[str] = []
        try:
            members = getattr(dim, "Members")
            member_count = int(getattr(members, "Count", 0))
            for j in range(1, member_count + 1):
                member_name = _extract_member_name(members.Item(j))
                if member_name:
                    member_names.append(member_name)
        except Exception:
            member_names = []
        if not member_names:
            rows.append({"dimension_name": dim_name, "member_name": ""})
            continue
        for member_name in member_names:
            rows.append({"dimension_name": dim_name, "member_name": member_name})
    return pd.DataFrame(rows, columns=DISCOVERY_COLUMNS)


def _resolve_dimension_name(discovery: pd.DataFrame, hints: Iterable[str]) -> str:
    if discovery.empty:
        return ""
    names = discovery["dimension_name"].fillna("").astype(str).tolist()
    lowered = {name.lower(): name for name in names if name}
    for hint in hints:
        key = str(hint or "").strip().lower()
        if key in lowered:
            return lowered[key]
    for hint in hints:
        key = str(hint or "").strip().lower()
        for name in names:
            if key and key in str(name).strip().lower():
                return str(name).strip()
    return ""


def _resolve_member_name(discovery: pd.DataFrame, dimension_name: str, hints: Iterable[str]) -> str:
    if discovery.empty or not dimension_name:
        return ""
    subset = discovery[discovery["dimension_name"].astype(str).str.strip().str.lower() == dimension_name.strip().lower()]
    if subset.empty:
        return ""
    names = subset["member_name"].fillna("").astype(str).tolist()
    lowered = {name.lower(): name for name in names if name}
    for hint in hints:
        key = str(hint or "").strip().lower()
        if key in lowered:
            return lowered[key]
    for hint in hints:
        key = str(hint or "").strip().lower()
        for name in names:
            if key and key in str(name).strip().lower():
                return str(name).strip()
    return ""


def _resolve_member_name_with_override(
    discovery: pd.DataFrame,
    *,
    dimension_name: str,
    hints: Iterable[str],
    override_value: str = "",
) -> str:
    override = str(override_value or "").strip()
    if override:
        return override
    return _resolve_member_name(discovery, dimension_name, hints)


def _coerce_bool(series: pd.Series) -> pd.Series:
    return series.fillna(False).astype(bool)


def normalize_raw_results(raw_df: pd.DataFrame) -> pd.DataFrame:
    if raw_df.empty:
        return empty_normalized_long()

    df = scale_raw_energy_values(raw_df)
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    base_keys = ["scenario", "region", "year", "sector_group", "module", "branch", "process", "fuel"]

    metric_rows: list[dict[str, object]] = []
    for key_values, group in df.groupby(base_keys, dropna=False):
        key_map = dict(zip(base_keys, key_values if isinstance(key_values, tuple) else (key_values,), strict=False))
        values = (
            group.groupby("notes", dropna=False)["value"]
            .sum(min_count=1)
            .to_dict()
        )
        feedstock_input = values.get("feedstock_input")
        aux_total = values.get("aux_total")
        aux_from_outputs = values.get("aux_from_outputs")
        aux_from_other = values.get("aux_from_other_modules_or_imports")
        net_output = values.get("net_output")
        output_for_auxiliary_use = values.get("output_for_auxiliary_use")
        explicit_loss_module_loss = values.get("explicit_loss_module_loss")

        gross_output = None
        if pd.notna(net_output) or pd.notna(output_for_auxiliary_use):
            gross_output = (0.0 if pd.isna(net_output) else float(net_output)) + (
                0.0 if pd.isna(output_for_auxiliary_use) else float(output_for_auxiliary_use)
            )
        own_use_internal = aux_from_outputs
        own_use_total = aux_total
        conversion_loss = None
        if pd.notna(feedstock_input) and pd.notna(gross_output):
            conversion_loss = float(feedstock_input) - float(gross_output)
        total_loss_reported = explicit_loss_module_loss
        if pd.isna(total_loss_reported) and pd.notna(conversion_loss):
            total_loss_reported = conversion_loss

        derived_metrics = {
            "feedstock_input": feedstock_input,
            "aux_total": aux_total,
            "aux_from_outputs": aux_from_outputs,
            "aux_from_other_modules_or_imports": aux_from_other,
            "net_output": net_output,
            "output_for_auxiliary_use": output_for_auxiliary_use,
            "gross_output": gross_output,
            "conversion_loss": conversion_loss,
            "explicit_loss_module_loss": explicit_loss_module_loss,
            "total_loss_reported": total_loss_reported,
            "own_use_internal": own_use_internal,
            "own_use_total": own_use_total,
        }

        qa_negative_conversion_loss = pd.notna(conversion_loss) and float(conversion_loss) < 0
        qa_aux_exceeds_gross_output = (
            pd.notna(aux_from_outputs) and pd.notna(gross_output) and float(aux_from_outputs) > float(gross_output)
        )
        qa_feedstock_missing_but_output_present = pd.isna(feedstock_input) and pd.notna(gross_output) and float(gross_output) > 0
        qa_zero_output_with_positive_aux = pd.notna(aux_total) and float(aux_total) > 0 and (
            pd.isna(gross_output) or float(gross_output) == 0
        )
        qa_balance_gap = False

        for metric, value in derived_metrics.items():
            metric_rows.append(
                {
                    **key_map,
                    "metric": metric,
                    "value": value,
                    "unit": group["unit"].dropna().astype(str).iloc[0] if group["unit"].notna().any() else "",
                    "source_variable": "Derived Transformation Metric",
                    "source_filter": "",
                    "notes": "",
                    "qa_negative_conversion_loss": qa_negative_conversion_loss,
                    "qa_aux_exceeds_gross_output": qa_aux_exceeds_gross_output,
                    "qa_feedstock_missing_but_output_present": qa_feedstock_missing_but_output_present,
                    "qa_zero_output_with_positive_aux": qa_zero_output_with_positive_aux,
                    "qa_balance_gap": qa_balance_gap,
                }
            )
    out = pd.DataFrame(metric_rows, columns=NORMALIZED_COLUMNS)
    for col in [
        "qa_negative_conversion_loss",
        "qa_aux_exceeds_gross_output",
        "qa_feedstock_missing_but_output_present",
        "qa_zero_output_with_positive_aux",
        "qa_balance_gap",
    ]:
        out[col] = _coerce_bool(out[col])
    return out


def build_dashboard_leap_long(
    normalized_df: pd.DataFrame,
    *,
    economy: str,
    dashboard_sheet_definitions: dict[str, dict[str, str]],
    unit: str,
) -> pd.DataFrame:
    if normalized_df.empty:
        return empty_dashboard_leap_long()

    rows: list[dict[str, object]] = []
    for definition in dashboard_sheet_definitions.values():
        metric = str(definition.get("metric") or "").strip()
        if not metric:
            continue
        subset = normalized_df[normalized_df["metric"].astype(str).str.strip() == metric].copy()
        if subset.empty:
            continue
        grouped = (
            subset.groupby(["scenario", "region", "year"], dropna=False)["value"]
            .sum(min_count=1)
            .reset_index()
        )
        for row in grouped.itertuples(index=False):
            rows.append(
                {
                    "economy": economy,
                    "scenario": row.scenario,
                    "region": row.region,
                    "sheet_name": str(definition.get("sheet_name") or ""),
                    "sector_code_9th": str(definition.get("sector_code_9th") or ""),
                    "sector_name": str(definition.get("sector_name") or ""),
                    "fuel_label": str(definition.get("fuel_label") or "Total"),
                    "year": row.year,
                    "leap_value": row.value,
                    "leap_variable": str(definition.get("leap_variable") or "Derived Transformation Metric"),
                    "leap_units": unit,
                    "measure": str(definition.get("measure") or ""),
                    "leap_scale_note": "Derived from LEAP transformation losses/own-use extraction",
                }
            )
    return pd.DataFrame(rows, columns=DASHBOARD_COLUMNS)


def build_qa_flags(normalized_df: pd.DataFrame) -> pd.DataFrame:
    if normalized_df.empty:
        return empty_qa_flags()
    rows: list[dict[str, object]] = []
    flag_columns = [
        "qa_negative_conversion_loss",
        "qa_aux_exceeds_gross_output",
        "qa_feedstock_missing_but_output_present",
        "qa_zero_output_with_positive_aux",
        "qa_balance_gap",
    ]
    for item in normalized_df.itertuples(index=False):
        for flag in flag_columns:
            if bool(getattr(item, flag, False)):
                rows.append(
                    {
                        "scenario": item.scenario,
                        "region": item.region,
                        "year": item.year,
                        "sector_group": item.sector_group,
                        "module": item.module,
                        "process": item.process,
                        "fuel": item.fuel,
                        "metric": item.metric,
                        "qa_flag": flag,
                        "qa_value": item.value,
                        "details": "",
                    }
                )
    return pd.DataFrame(rows, columns=QA_COLUMNS)


def extract_transformation_losses_own_use(
    *,
    scenarios: Iterable[str],
    regions: Iterable[str],
    years: Iterable[int],
    modules: Iterable[dict[str, object]],
    result_dimension_hints: dict[str, tuple[str, ...]],
    result_member_hints: dict[str, tuple[str, ...]],
    manual_result_member_overrides: dict[str, dict[str, dict[str, str]]] | None,
    unit: str,
    visible: bool = False,
    reuse_running: bool = True,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    app = connect_leap(visible=visible, reuse_running=reuse_running)
    app.ActiveView = "Results"
    ensure_calculated(app, force=False)

    discovery = discover_result_dimensions(app)
    fuel_dimension = _resolve_dimension_name(discovery, result_dimension_hints.get("fuel", ()))
    input_type_dimension = _resolve_dimension_name(discovery, result_dimension_hints.get("input_type", ()))
    output_type_dimension = _resolve_dimension_name(discovery, result_dimension_hints.get("output_type", ()))

    rows: list[dict[str, object]] = []
    for module_cfg in modules:
        branch_path = str(module_cfg.get("branch") or "").strip()
        module_name = str(module_cfg.get("module_name") or "").strip()
        sector_group = str(module_cfg.get("sector_group") or "").strip()
        if not branch_path:
            continue
        branch_obj = app.Branch(branch_path)
        process_name = getattr(branch_obj, "Name", module_name) or module_name
        module_overrides = (manual_result_member_overrides or {}).get(module_name, {})
        input_overrides = module_overrides.get("Inputs", {})
        output_overrides = module_overrides.get("Outputs by Output Fuel", {})
        variable_map = {
            "Inputs": {
                "feedstock_input": _resolve_member_name_with_override(
                    discovery,
                    dimension_name=input_type_dimension,
                    hints=result_member_hints.get("feedstock_input", ()),
                    override_value=input_overrides.get("feedstock_input", ""),
                ),
                "aux_total": _resolve_member_name_with_override(
                    discovery,
                    dimension_name=input_type_dimension,
                    hints=result_member_hints.get("aux_total", ()),
                    override_value=input_overrides.get("aux_total", ""),
                ),
                "aux_from_outputs": _resolve_member_name_with_override(
                    discovery,
                    dimension_name=input_type_dimension,
                    hints=result_member_hints.get("aux_from_outputs", ()),
                    override_value=input_overrides.get("aux_from_outputs", ""),
                ),
                "aux_from_other_modules_or_imports": _resolve_member_name_with_override(
                    discovery,
                    dimension_name=input_type_dimension,
                    hints=result_member_hints.get("aux_from_other", ()),
                    override_value=input_overrides.get("aux_from_other_modules_or_imports", ""),
                ),
            },
            "Outputs by Output Fuel": {
                "net_output": _resolve_member_name_with_override(
                    discovery,
                    dimension_name=output_type_dimension,
                    hints=result_member_hints.get("net_output", ()),
                    override_value=output_overrides.get("net_output", ""),
                ),
                "output_for_auxiliary_use": _resolve_member_name_with_override(
                    discovery,
                    dimension_name=output_type_dimension,
                    hints=result_member_hints.get("output_for_auxiliary_use", ()),
                    override_value=output_overrides.get("output_for_auxiliary_use", ""),
                ),
            },
        }

        for variable_name, metric_members in variable_map.items():
            try:
                variable_obj = branch_obj.Variable(variable_name)
            except Exception:
                continue
            for scenario in scenarios:
                for region in regions:
                    for year in years:
                        for metric_name, member_name in metric_members.items():
                            filter_parts: list[str] = []
                            if fuel_dimension:
                                filter_parts.append(fuel_dimension + "=Total")
                            if variable_name == "Inputs" and input_type_dimension and member_name:
                                filter_parts.append(input_type_dimension + "=" + member_name)
                            if variable_name == "Outputs by Output Fuel" and output_type_dimension and member_name:
                                filter_parts.append(output_type_dimension + "=" + member_name)
                            filter_str = "; ".join(part for part in filter_parts if part)
                            try:
                                value = fetch_values_rs(
                                    variable_obj,
                                    region=region,
                                    scenario=scenario,
                                    year=int(year),
                                    unit="",
                                    filter_str=filter_str,
                                )
                            except Exception:
                                continue
                            rows.append(
                                {
                                    "scenario": scenario,
                                    "region": region,
                                    "year": int(year),
                                    "sector_group": sector_group,
                                    "module": module_name,
                                    "branch": branch_path,
                                    "process": str(process_name),
                                    "fuel": "Total",
                                    "variable": variable_name,
                                    "dimension_name": input_type_dimension if variable_name == "Inputs" else output_type_dimension,
                                    "member_name": member_name,
                                    "filter_str": filter_str,
                                    "value": value,
                                    "unit": unit,
                                    "notes": metric_name,
                                }
                            )

        if sector_group == "transmission_distribution":
            try:
                inputs_obj = branch_obj.Variable("Inputs")
                outputs_obj = branch_obj.Variable("Outputs by Output Fuel")
            except Exception:
                continue
            for scenario in scenarios:
                for region in regions:
                    for year in years:
                        try:
                            input_value = fetch_values_rs(inputs_obj, region=region, scenario=scenario, year=int(year), unit="", filter_str="")
                            output_value = fetch_values_rs(outputs_obj, region=region, scenario=scenario, year=int(year), unit="", filter_str="")
                        except Exception:
                            continue
                        rows.append(
                            {
                                "scenario": scenario,
                                "region": region,
                                "year": int(year),
                                "sector_group": sector_group,
                                "module": module_name,
                                "branch": branch_path,
                                "process": str(process_name),
                                "fuel": "Total",
                                "variable": "Derived explicit T&D loss",
                                "dimension_name": "",
                                "member_name": "",
                                "filter_str": "",
                                "value": float(input_value) - float(output_value),
                                "unit": unit,
                                "notes": "explicit_loss_module_loss",
                            }
                        )

    raw_df = pd.DataFrame(rows, columns=RAW_RESULT_COLUMNS)
    raw_df = scale_raw_energy_values(raw_df)
    return raw_df, discovery


def write_losses_own_use_outputs(
    *,
    output_dir: Path,
    raw_df: pd.DataFrame,
    normalized_df: pd.DataFrame,
    dashboard_df: pd.DataFrame,
    qa_df: pd.DataFrame,
    discovery_df: pd.DataFrame,
) -> dict[str, str]:
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "raw_result_pulls": output_dir / "raw_result_pulls.csv",
        "normalized_long": output_dir / "normalized_long.csv",
        "dashboard_leap_long": output_dir / "dashboard_leap_long.csv",
        "qa_flags": output_dir / "qa_flags.csv",
        "dimension_discovery": output_dir / "dimension_discovery.csv",
    }
    raw_df.to_csv(paths["raw_result_pulls"], index=False)
    normalized_df.to_csv(paths["normalized_long"], index=False)
    dashboard_df.to_csv(paths["dashboard_leap_long"], index=False)
    qa_df.to_csv(paths["qa_flags"], index=False)
    discovery_df.to_csv(paths["dimension_discovery"], index=False)
    return {key: str(path) for key, path in paths.items()}


def build_ninth_merge_exploration(
    normalized_df: pd.DataFrame,
    *,
    projection_df: pd.DataFrame,
    projection_economy: str,
    scenario: str,
    merge_candidates: Iterable[dict[str, str]],
) -> pd.DataFrame:
    columns = [
        "module",
        "metric",
        "candidate_sector_code",
        "fuel_code",
        "join_kind",
        "scenario",
        "year",
        "leap_value",
        "ninth_value",
        "ninth_value_abs",
        "abs_gap",
        "ratio",
        "notes",
    ]
    if normalized_df.empty or projection_df.empty:
        return pd.DataFrame(columns=columns)

    proj = projection_df.copy()
    proj = proj[proj["economy"].astype(str).str.upper().eq(str(projection_economy).upper())].copy()
    proj = proj[proj["scenarios"].astype(str).str.lower().eq(str(scenario).lower())].copy()
    if proj.empty:
        return pd.DataFrame(columns=columns)

    year_cols = [col for col in proj.columns if str(col).isdigit()]
    rows: list[dict[str, object]] = []
    for candidate in merge_candidates:
        module = str(candidate.get("module") or "")
        metric = str(candidate.get("metric") or "")
        sector_code = str(candidate.get("candidate_sector_code") or "")
        fuel_code = str(candidate.get("fuel_code") or "")
        join_kind = str(candidate.get("join_kind") or "")
        notes = str(candidate.get("notes") or "")

        leap_subset = normalized_df[
            normalized_df["module"].astype(str).eq(module)
            & normalized_df["metric"].astype(str).eq(metric)
            & normalized_df["scenario"].astype(str).str.lower().eq(str(scenario).lower())
        ].copy()
        if leap_subset.empty:
            continue
        leap_year = leap_subset.groupby("year", as_index=False)["value"].sum(min_count=1)

        if not sector_code:
            for row in leap_year.itertuples(index=False):
                rows.append(
                    {
                        "module": module,
                        "metric": metric,
                        "candidate_sector_code": sector_code,
                        "fuel_code": fuel_code,
                        "join_kind": join_kind,
                        "scenario": scenario,
                        "year": int(row.year),
                        "leap_value": row.value,
                        "ninth_value": pd.NA,
                        "ninth_value_abs": pd.NA,
                        "abs_gap": pd.NA,
                        "ratio": pd.NA,
                        "notes": notes,
                    }
                )
            continue

        sector_mask = False
        for col in ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]:
            sector_mask = sector_mask | proj[col].astype(str).str.lower().eq(sector_code.lower())
        ninth_subset = proj[sector_mask].copy()
        if fuel_code:
            fuel_mask = False
            for col in ["fuels", "subfuels"]:
                fuel_mask = fuel_mask | ninth_subset[col].astype(str).str.lower().eq(fuel_code.lower())
            ninth_subset = ninth_subset[fuel_mask].copy()

        if ninth_subset.empty:
            for row in leap_year.itertuples(index=False):
                rows.append(
                    {
                        "module": module,
                        "metric": metric,
                        "candidate_sector_code": sector_code,
                        "fuel_code": fuel_code,
                        "join_kind": join_kind,
                        "scenario": scenario,
                        "year": int(row.year),
                        "leap_value": row.value,
                        "ninth_value": pd.NA,
                        "ninth_value_abs": pd.NA,
                        "abs_gap": pd.NA,
                        "ratio": pd.NA,
                        "notes": notes,
                    }
                )
            continue

        ninth_year = ninth_subset[year_cols].sum(axis=0, numeric_only=True)
        ninth_year.index = ninth_year.index.astype(int)
        merged = leap_year.copy()
        merged["year"] = pd.to_numeric(merged["year"], errors="coerce").astype("Int64")
        merged["ninth_value"] = merged["year"].map(ninth_year.to_dict())
        merged["ninth_value_abs"] = pd.to_numeric(merged["ninth_value"], errors="coerce").abs()
        merged["abs_gap"] = pd.to_numeric(merged["value"], errors="coerce") - pd.to_numeric(merged["ninth_value_abs"], errors="coerce")
        denom = pd.to_numeric(merged["ninth_value_abs"], errors="coerce").replace({0: pd.NA})
        merged["ratio"] = pd.to_numeric(merged["value"], errors="coerce") / denom
        for row in merged.itertuples(index=False):
            rows.append(
                {
                    "module": module,
                    "metric": metric,
                    "candidate_sector_code": sector_code,
                    "fuel_code": fuel_code,
                    "join_kind": join_kind,
                    "scenario": scenario,
                    "year": int(row.year),
                        "leap_value": row.value,
                        "ninth_value": row.ninth_value,
                        "ninth_value_abs": row.ninth_value_abs,
                        "abs_gap": row.abs_gap,
                        "ratio": row.ratio,
                        "notes": notes,
                }
            )
    return pd.DataFrame(rows, columns=columns)


def build_ninth_mapped_charts(merge_df: pd.DataFrame, *, output_dir: Path) -> dict[str, str]:
    dashboards_dir = output_dir / "dashboards"
    charts_dir = output_dir / "mapped_charts"
    dashboards_dir.mkdir(parents=True, exist_ok=True)
    charts_dir.mkdir(parents=True, exist_ok=True)

    out_page = dashboards_dir / "mapped_9th_comparisons.html"
    if merge_df.empty:
        out_page.write_text(
            "<html><body><h2>No merge exploration data</h2><p>No mapped LEAP-vs-9th comparison charts were generated.</p></body></html>",
            encoding="utf-8",
        )
        return {"mapped_comparison_dashboard": str(out_page)}

    try:
        import plotly.graph_objects as go
    except Exception as exc:
        out_page.write_text(
            f"<html><body><h2>Plotly unavailable</h2><p>{exc}</p></body></html>",
            encoding="utf-8",
        )
        return {"mapped_comparison_dashboard": str(out_page)}

    df = merge_df.copy()
    df["year"] = pd.to_numeric(df["year"], errors="coerce")
    df["leap_value"] = pd.to_numeric(df["leap_value"], errors="coerce")
    df["ninth_value"] = pd.to_numeric(df["ninth_value"], errors="coerce")
    df["ninth_value_abs"] = pd.to_numeric(df.get("ninth_value_abs"), errors="coerce")
    df = df[df["year"].notna() & df["leap_value"].notna()].copy()

    created: list[tuple[str, str, str, Path, str]] = []
    group_cols = ["module", "metric", "candidate_sector_code", "fuel_code", "join_kind", "notes"]
    for key_values, group in df.groupby(group_cols, dropna=False):
        module, metric, sector_code, fuel_code, join_kind, notes = key_values
        fig = go.Figure()
        subgroup = group.sort_values("year", kind="mergesort")
        fig.add_trace(
            go.Scatter(
                x=subgroup["year"].astype(int).tolist(),
                y=subgroup["leap_value"].tolist(),
                mode="lines+markers",
                name="LEAP derived",
            )
        )
        if subgroup["ninth_value_abs"].notna().any():
            fig.add_trace(
                go.Scatter(
                    x=subgroup["year"].astype(int).tolist(),
                    y=subgroup["ninth_value_abs"].tolist(),
                    mode="lines+markers",
                    name="9th comparator (abs)",
                )
            )
        if subgroup["ninth_value"].notna().any():
            fig.add_trace(
                go.Scatter(
                    x=subgroup["year"].astype(int).tolist(),
                    y=subgroup["ninth_value"].tolist(),
                    mode="lines",
                    name="9th comparator (signed)",
                    line={"dash": "dot"},
                )
            )
        title = f"{module} | {metric}"
        if str(sector_code or "").strip():
            title += f" | {sector_code}"
        fig.update_layout(
            template="plotly_white",
            title=title,
            xaxis_title="Year",
            yaxis_title="PJ",
            legend_title="Series",
        )
        safe_name = "__".join(_safe_filename_token(part) for part in [module, metric, sector_code, fuel_code] if str(part).strip())
        chart_path = charts_dir / f"{safe_name}.html"
        fig.write_html(chart_path, include_plotlyjs="cdn")
        created.append((str(module), str(metric), str(sector_code), chart_path, str(notes)))

    lines = [
        "<html><head><title>Mapped LEAP vs 9th Losses</title>",
        "<style>",
        "body { font-family: Arial, sans-serif; margin: 24px; background: #f6f7f9; color: #1f2937; }",
        ".chart-card { background: #fff; border: 1px solid #d1d5db; border-radius: 10px; padding: 16px; margin: 0 0 18px 0; }",
        ".chart-title { font-size: 18px; font-weight: 600; margin: 0 0 8px 0; }",
        ".chart-meta { font-size: 13px; color: #4b5563; margin: 0 0 10px 0; }",
        "iframe { width: 100%; height: 560px; border: 0; background: #fff; }",
        "</style></head><body>",
        "<h1>Mapped LEAP vs 9th Losses / Own Use</h1>",
        "<p>These are simple dict-based candidate joins, not the standard dashboard mapping chain.</p>",
    ]
    for module, metric, sector_code, chart_path, notes in created:
        rel = chart_path.relative_to(dashboards_dir.parent).as_posix()
        lines.extend(
            [
                "<section class=\"chart-card\">",
                f"<div class=\"chart-title\">{module} | {metric}</div>",
                f"<div class=\"chart-meta\">Candidate 9th sector: {sector_code or 'none'} | {notes}</div>",
                f"<iframe src=\"../{rel}\" loading=\"lazy\"></iframe>",
                "</section>",
            ]
        )
    lines.append("</body></html>")
    out_page.write_text("\n".join(lines), encoding="utf-8")
    return {
        "mapped_comparison_dashboard": str(out_page),
        "mapped_chart_count": str(len(created)),
    }


def build_inspection_charts(normalized_df: pd.DataFrame, *, output_dir: Path) -> dict[str, str]:
    charts_dir = output_dir / "charts"
    dashboards_dir = output_dir / "dashboards"
    charts_dir.mkdir(parents=True, exist_ok=True)
    dashboards_dir.mkdir(parents=True, exist_ok=True)

    if normalized_df.empty:
        index_path = dashboards_dir / "index.html"
        combined_path = dashboards_dir / "all_charts.html"
        index_path.write_text(
            (
                "<html><body><h2>No losses/own-use data</h2>"
                "<p>The normalized dataset is empty, so no inspection charts were generated.</p></body></html>"
            ),
            encoding="utf-8",
        )
        combined_path.write_text(
            (
                "<html><body><h2>No losses/own-use data</h2>"
                "<p>The normalized dataset is empty, so no inspection charts were generated.</p></body></html>"
            ),
            encoding="utf-8",
        )
        return {"dashboard_index": str(index_path), "dashboard_all_charts": str(combined_path)}

    try:
        import plotly.graph_objects as go
    except Exception as exc:
        index_path = dashboards_dir / "index.html"
        combined_path = dashboards_dir / "all_charts.html"
        index_path.write_text(
            (
                "<html><body><h2>Plotly unavailable</h2>"
                f"<p>Inspection charts were skipped: {exc}</p></body></html>"
            ),
            encoding="utf-8",
        )
        combined_path.write_text(
            (
                "<html><body><h2>Plotly unavailable</h2>"
                f"<p>Inspection charts were skipped: {exc}</p></body></html>"
            ),
            encoding="utf-8",
        )
        return {"dashboard_index": str(index_path), "dashboard_all_charts": str(combined_path)}

    df = normalized_df.copy()
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    df["year"] = pd.to_numeric(df["year"], errors="coerce")
    df = df[df["value"].notna() & df["year"].notna()].copy()
    if df.empty:
        index_path = dashboards_dir / "index.html"
        combined_path = dashboards_dir / "all_charts.html"
        index_path.write_text(
            "<html><body><h2>No plottable data</h2><p>All values were empty after numeric coercion.</p></body></html>",
            encoding="utf-8",
        )
        combined_path.write_text(
            "<html><body><h2>No plottable data</h2><p>All values were empty after numeric coercion.</p></body></html>",
            encoding="utf-8",
        )
        return {"dashboard_index": str(index_path), "dashboard_all_charts": str(combined_path)}

    created: list[tuple[str, str, str, str, Path]] = []
    group_cols = ["sector_group", "module", "metric", "region"]
    for key_values, group in df.groupby(group_cols, dropna=False):
        sector_group, module, metric, region = key_values
        subgroup = group.sort_values(["scenario", "fuel", "year"], kind="mergesort")
        fig = go.Figure()
        for (scenario, fuel), series in subgroup.groupby(["scenario", "fuel"], dropna=False):
            line_name = f"{scenario} | {fuel}"
            fig.add_trace(
                go.Scatter(
                    x=series["year"].astype(int).tolist(),
                    y=series["value"].tolist(),
                    mode="lines+markers",
                    name=line_name,
                )
            )
        fig.update_layout(
            template="plotly_white",
            title=f"{module} - {metric}",
            xaxis_title="Year",
            yaxis_title=str(subgroup["unit"].dropna().astype(str).iloc[0] if subgroup["unit"].notna().any() else "Value"),
            legend_title="Scenario | Fuel",
        )
        safe_name = "__".join(
            _safe_filename_token(part)
            for part in [str(sector_group), str(module), str(metric), str(region)]
            if str(part).strip()
        )
        chart_path = charts_dir / f"{safe_name}.html"
        fig.write_html(chart_path, include_plotlyjs="cdn")
        created.append((str(sector_group), str(module), str(metric), str(region), chart_path))

    index_lines = [
        "<html><body>",
        "<h1>Losses / Own-Use Inspection Charts</h1>",
        "<p>These charts come directly from the normalized extraction outputs. No ESTO/9th mapping is applied here.</p>",
        "<p><a href=\"all_charts.html\">Open all charts in one page</a></p>",
        "<ul>",
    ]
    for sector_group, module, metric, region, chart_path in created:
        rel = chart_path.relative_to(dashboards_dir.parent).as_posix()
        index_lines.append(
            f"<li><a href=\"../{rel}\">{module} | {metric} | {region}</a> <span>({sector_group})</span></li>"
        )
    index_lines.extend(["</ul>", "</body></html>"])
    index_path = dashboards_dir / "index.html"
    index_path.write_text("\n".join(index_lines), encoding="utf-8")

    combined_lines = [
        "<html><head><title>Losses / Own-Use Inspection Charts</title>",
        "<style>",
        "body { font-family: Arial, sans-serif; margin: 24px; background: #f6f7f9; color: #1f2937; }",
        "h1 { margin-bottom: 8px; }",
        "p { margin-top: 0; }",
        ".chart-card { background: #fff; border: 1px solid #d1d5db; border-radius: 10px; padding: 16px; margin: 0 0 18px 0; }",
        ".chart-title { font-size: 18px; font-weight: 600; margin: 0 0 8px 0; }",
        ".chart-meta { font-size: 13px; color: #4b5563; margin: 0 0 10px 0; }",
        "iframe { width: 100%; height: 560px; border: 0; background: #fff; }",
        "a { color: #0f766e; }",
        "</style></head><body>",
        "<h1>Losses / Own-Use Inspection Charts</h1>",
        "<p>Combined view of all inspection charts. No ESTO/9th mapping is applied here.</p>",
        "<p><a href=\"index.html\">Back to chart index</a></p>",
    ]
    for sector_group, module, metric, region, chart_path in created:
        rel = chart_path.relative_to(dashboards_dir.parent).as_posix()
        combined_lines.extend(
            [
                "<section class=\"chart-card\">",
                f"<div class=\"chart-title\">{module} | {metric}</div>",
                f"<div class=\"chart-meta\">Sector group: {sector_group} | Region: {region}</div>",
                f"<iframe src=\"../{rel}\" loading=\"lazy\"></iframe>",
                "</section>",
            ]
        )
    combined_lines.append("</body></html>")
    combined_path = dashboards_dir / "all_charts.html"
    combined_path.write_text("\n".join(combined_lines), encoding="utf-8")
    return {
        "dashboard_index": str(index_path),
        "dashboard_all_charts": str(combined_path),
        "chart_count": str(len(created)),
    }


def _safe_filename_token(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return "blank"
    out = []
    for char in text:
        if char.isalnum():
            out.append(char)
        else:
            out.append("_")
    return "".join(out).strip("_") or "blank"
