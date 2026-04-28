from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re

import pandas as pd

from codebase.utilities.master_config import config_table_exists, read_config_table
from codebase.utilities.leap_results_dashboard_utils import (
    build_sector_to_esto_flow_lookup,
    parse_template_sheet,
    pull_base_year_value,
    pull_projection_series,
)

REPO_ROOT = Path(__file__).resolve().parents[3]
DUMMY_OWN_USE_SHARE_PATH = REPO_ROOT / "config" / "transformation_auxiliary_own_use_dummy_shares.csv"
PROCESSED_TABLES_SUBDIR = "processed tables"


@dataclass(frozen=True)
class DerivedTransformationSpec:
    sheet_name: str
    sector_code_9th: str
    sector_name: str
    input_sheet: str
    output_sheet: str
    output_feed_sheet: str
    measure: str
    display_label: str
    viability: str
    rationale: str


@dataclass(frozen=True)
class _SplitMeasure:
    key: str
    measure: str


MEASURE_LOSSES = _SplitMeasure(key="losses", measure="Losses (PJ)")
MEASURE_OWN_USE = _SplitMeasure(key="own_use", measure="Own use (PJ)")


SPEC_OVERRIDES: dict[str, dict[str, str]] = {
    "refining": {
        "viability": "strong",
        "rationale": (
            "Inputs and product outputs are both fully mapped in the official dashboard. "
            "The derived gap is defensible as a combined losses-and-own-use indicator."
        ),
    },
    "hydrogen": {
        "viability": "partial",
        "rationale": (
            "The dashboard already maps hydrogen inputs and product outputs, but some comparator rows "
            "depend on explicit overrides and estimated series. Treat the derived gap as informative, not final."
        ),
    },
    "coke_ovens": {
        "viability": "control",
        "rationale": (
            "The official component mappings are clean, but the current derived gap is zero in the sampled output. "
            "Keep this as a control case rather than a high-priority dashboard signal."
        ),
    },
}


LEAP_LONG_COLUMNS = [
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


def _processed_tables_dir(data_dir: Path | None) -> Path | None:
    if data_dir is None:
        return None
    return Path(data_dir) / PROCESSED_TABLES_SUBDIR


def _resolve_aux_own_use_path(data_dir: Path | None, economy_token: str | None) -> Path | None:
    if data_dir is None or not str(economy_token or "").strip():
        return None
    token = str(economy_token).strip().upper()
    file_name = f"transformation_auxiliary_own_use_20_{token}.csv"
    processed_dir = _processed_tables_dir(data_dir)
    processed_path = (processed_dir / file_name) if processed_dir is not None else None
    legacy_path = Path(data_dir) / file_name
    if processed_path is not None and processed_path.exists():
        return processed_path
    if legacy_path.exists():
        return legacy_path
    return processed_path if processed_path is not None else legacy_path


def derive_transformation_specs(sheet_map: pd.DataFrame) -> tuple[DerivedTransformationSpec, ...]:
    if sheet_map.empty:
        return ()
    rows_by_name = {
        str(row["sheet_name"]).strip(): row
        for _, row in sheet_map.iterrows()
        if str(row.get("sheet_name") or "").strip()
    }
    specs: list[DerivedTransformationSpec] = []
    for sheet_name, row in rows_by_name.items():
        if not sheet_name.endswith("_inputs"):
            continue
        prefix = sheet_name[: -len("_inputs")]
        output_sheet = f"{prefix}_out_fuel"
        if output_sheet not in rows_by_name:
            continue
        output_feed_sheet = f"{prefix}_out_feed"
        input_sector_code = str(row.get("sector_code_9th") or "").strip()
        if not input_sector_code:
            continue
        if not (
            input_sector_code.startswith("09_")
            or input_sector_code.startswith("10_02")
            or input_sector_code.startswith("08_")
        ):
            continue
        loss_row = rows_by_name.get(f"{prefix}_loss_own_use_total")
        loss_sector_code = ""
        if loss_row is not None:
            loss_sector_code = str(loss_row.get("sector_code_9th") or "").strip()
        sector_code = loss_sector_code or input_sector_code
        if sheet_name == "upstream_liquids_inputs" or prefix == "refinery_blending":
            continue
        sector_name = str(row.get("sector_name") or prefix.replace("_", " ").title()).strip()
        override = SPEC_OVERRIDES.get(prefix, {})
        specs.append(
            DerivedTransformationSpec(
                sheet_name=f"{prefix}_loss_own_use_total",
                sector_code_9th=sector_code,
                sector_name=sector_name,
                input_sheet=sheet_name,
                output_sheet=output_sheet,
                output_feed_sheet=output_feed_sheet if output_feed_sheet in rows_by_name else "",
                measure="Losses (PJ)",
                display_label=f"{sector_name} losses & own use",
                viability=override.get("viability", "screening"),
                rationale=override.get(
                    "rationale",
                    (
                        "Calculated from the official dashboard component sheets as inputs minus product outputs. "
                        "Use this as a combined transformation-gap screen rather than a clean split between "
                        "conversion losses and own use."
                    ),
                ),
            )
        )
    return tuple(sorted(specs, key=lambda item: item.sheet_name))


def _derive_gap_frame(
    frame: pd.DataFrame,
    *,
    input_col: str,
    output_col: str,
    group_cols: list[str],
    value_col: str,
    required_sources: set[str] | None = None,
) -> pd.DataFrame:
    if frame.empty:
        return pd.DataFrame(columns=group_cols + [input_col, output_col, value_col])
    sub = frame.copy()
    if required_sources is not None and "source" in sub.columns:
        sub = sub[sub["source"].astype(str).isin(required_sources)].copy()
    sub[value_col] = pd.to_numeric(sub[value_col], errors="coerce")
    grouped = sub.groupby(group_cols + ["sheet_name"], as_index=False)[value_col].sum(min_count=1)
    pivot = (
        grouped.pivot_table(
            index=group_cols,
            columns="sheet_name",
            values=value_col,
            aggfunc="first",
        )
        .reset_index()
    )
    return pivot


def _build_input_fuel_shares(
    frame: pd.DataFrame,
    *,
    sheet_col: str,
    value_col: str,
    input_sheet: str,
    group_cols: list[str],
    fuel_col: str = "fuel_label",
) -> pd.DataFrame:
    columns = group_cols + [fuel_col, "input_fuel_value", "input_fuel_share", "allocation_method"]
    if frame.empty:
        return pd.DataFrame(columns=columns)
    subset = frame[frame[sheet_col].astype(str).eq(input_sheet)].copy()
    if subset.empty:
        return pd.DataFrame(columns=columns)
    subset[value_col] = pd.to_numeric(subset[value_col], errors="coerce")
    subset[fuel_col] = subset[fuel_col].astype(str).str.strip()
    subset = subset[subset[fuel_col].ne("")].copy()
    non_total = subset[~subset[fuel_col].str.lower().eq("total")].copy()
    if not non_total.empty:
        subset = non_total
    grouped = (
        subset.groupby(group_cols + [fuel_col], as_index=False)[value_col]
        .sum(min_count=1)
        .rename(columns={value_col: "input_fuel_value"})
    )
    if grouped.empty:
        return pd.DataFrame(columns=columns)
    grouped["input_fuel_value"] = pd.to_numeric(grouped["input_fuel_value"], errors="coerce").fillna(0.0)
    grouped["input_weight"] = grouped["input_fuel_value"].abs()
    weight_sum = grouped.groupby(group_cols, as_index=False)["input_weight"].sum().rename(columns={"input_weight": "weight_sum"})
    fuel_count = grouped.groupby(group_cols, as_index=False)[fuel_col].size().rename(columns={"size": "fuel_count"})
    grouped = grouped.merge(weight_sum, on=group_cols, how="left")
    grouped = grouped.merge(fuel_count, on=group_cols, how="left")
    grouped["input_fuel_share"] = pd.NA
    weight_mask = pd.to_numeric(grouped["weight_sum"], errors="coerce").gt(0)
    grouped.loc[weight_mask, "input_fuel_share"] = (
        pd.to_numeric(grouped.loc[weight_mask, "input_weight"], errors="coerce")
        / pd.to_numeric(grouped.loc[weight_mask, "weight_sum"], errors="coerce")
    )
    grouped.loc[~weight_mask, "input_fuel_share"] = 1.0 / pd.to_numeric(grouped.loc[~weight_mask, "fuel_count"], errors="coerce")
    grouped["allocation_method"] = "proportional_abs_input"
    grouped.loc[~weight_mask, "allocation_method"] = "equal_split_zero_input"
    grouped["input_fuel_share"] = pd.to_numeric(grouped["input_fuel_share"], errors="coerce")
    return grouped.drop(columns=["input_weight", "weight_sum", "fuel_count"], errors="ignore")


def _load_dummy_own_use_shares() -> dict[str, float]:
    defaults = {"*": 0.25}
    if not config_table_exists(DUMMY_OWN_USE_SHARE_PATH):
        return defaults
    try:
        df = read_config_table(DUMMY_OWN_USE_SHARE_PATH)
    except Exception:
        return defaults
    if df.empty or "sheet_name" not in df.columns or "own_use_share" not in df.columns:
        return defaults
    out: dict[str, float] = {}
    for row in df.itertuples(index=False):
        sheet = str(getattr(row, "sheet_name", "") or "").strip()
        if not sheet:
            continue
        share = pd.to_numeric(getattr(row, "own_use_share", pd.NA), errors="coerce")
        if pd.isna(share):
            continue
        out[sheet] = float(min(max(float(share), 0.0), 1.0))
    if "*" not in out:
        out["*"] = defaults["*"]
    return out


def _load_aux_own_use_totals(
    *,
    data_dir: Path | None,
    economy_token: str | None,
    expected_sheet_names: set[str],
    group_cols: list[str],
) -> pd.DataFrame:
    if data_dir is None or not str(economy_token or "").strip():
        return pd.DataFrame(columns=group_cols + ["sheet_name", "own_use_total"])
    path = _resolve_aux_own_use_path(data_dir, economy_token)
    if path is None or not path.exists():
        return pd.DataFrame(columns=group_cols + ["sheet_name", "own_use_total"])
    try:
        df = pd.read_csv(path)
    except Exception:
        return pd.DataFrame(columns=group_cols + ["sheet_name", "own_use_total"])
    if df.empty:
        return pd.DataFrame(columns=group_cols + ["sheet_name", "own_use_total"])
    working = df.copy()
    if "sheet_name" not in working.columns:
        return pd.DataFrame(columns=group_cols + ["sheet_name", "own_use_total"])
    working["sheet_name"] = working["sheet_name"].astype(str).str.strip()
    working = working[working["sheet_name"].isin(expected_sheet_names)].copy()
    if working.empty:
        return pd.DataFrame(columns=group_cols + ["sheet_name", "own_use_total"])
    if "own_use_value" in working.columns:
        value_col = "own_use_value"
    elif "leap_value" in working.columns:
        value_col = "leap_value"
    elif "value" in working.columns:
        value_col = "value"
    else:
        return pd.DataFrame(columns=group_cols + ["sheet_name", "own_use_total"])
    for col in group_cols:
        if col not in working.columns:
            working[col] = pd.NA
    working[value_col] = pd.to_numeric(working[value_col], errors="coerce")
    grouped = (
        working.groupby(group_cols + ["sheet_name"], as_index=False)[value_col]
        .sum(min_count=1)
        .rename(columns={value_col: "own_use_total"})
    )
    grouped["own_use_total"] = pd.to_numeric(grouped["own_use_total"], errors="coerce").fillna(0.0)
    return grouped


def _load_aux_own_use_by_fuel(
    *,
    data_dir: Path | None,
    economy_token: str | None,
    expected_sheet_names: set[str],
    group_cols: list[str],
    fuel_col: str = "fuel_label",
) -> pd.DataFrame:
    columns = group_cols + ["sheet_name", fuel_col, "aux_own_use_fuel"]
    if data_dir is None or not str(economy_token or "").strip():
        return pd.DataFrame(columns=columns)
    path = _resolve_aux_own_use_path(data_dir, economy_token)
    if path is None or not path.exists():
        return pd.DataFrame(columns=columns)
    try:
        df = pd.read_csv(path)
    except Exception:
        return pd.DataFrame(columns=columns)
    if df.empty:
        return pd.DataFrame(columns=columns)
    if "sheet_name" not in df.columns or fuel_col not in df.columns:
        return pd.DataFrame(columns=columns)
    if "own_use_value" in df.columns:
        value_col = "own_use_value"
    elif "leap_value" in df.columns:
        value_col = "leap_value"
    elif "value" in df.columns:
        value_col = "value"
    else:
        return pd.DataFrame(columns=columns)
    working = df.copy()
    working["sheet_name"] = working["sheet_name"].astype(str).str.strip()
    working = working[working["sheet_name"].isin(expected_sheet_names)].copy()
    if working.empty:
        return pd.DataFrame(columns=columns)
    working[fuel_col] = working[fuel_col].astype(str).str.strip()
    working = working[working[fuel_col].ne("") & ~working[fuel_col].str.lower().eq("total")].copy()
    if working.empty:
        return pd.DataFrame(columns=columns)
    for col in group_cols:
        if col not in working.columns:
            working[col] = pd.NA
    working[value_col] = pd.to_numeric(working[value_col], errors="coerce")
    grouped = (
        working.groupby(group_cols + ["sheet_name", fuel_col], as_index=False)[value_col]
        .sum(min_count=1)
        .rename(columns={value_col: "aux_own_use_fuel"})
    )
    grouped["aux_own_use_fuel"] = pd.to_numeric(grouped["aux_own_use_fuel"], errors="coerce")
    return grouped


def _apply_own_use_fuel_allocation(
    *,
    allocated: pd.DataFrame,
    group_cols: list[str],
    data_dir: Path | None,
    economy_token: str | None,
) -> pd.DataFrame:
    if allocated.empty:
        return allocated
    out = allocated.copy()
    out["derived_value"] = pd.to_numeric(out["derived_value"], errors="coerce")
    out["input_fuel_share"] = pd.to_numeric(out["input_fuel_share"], errors="coerce").fillna(0.0)
    out["total_allocated"] = out["derived_value"] * out["input_fuel_share"]
    out["own_use_total"] = pd.to_numeric(out["own_use_total"], errors="coerce")
    out["own_use_allocated"] = out["own_use_total"] * out["input_fuel_share"]

    aux_by_fuel = _load_aux_own_use_by_fuel(
        data_dir=data_dir,
        economy_token=economy_token,
        expected_sheet_names={str(v).strip() for v in out["sheet_name"].dropna().astype(str).tolist()},
        group_cols=group_cols,
        fuel_col="fuel_label",
    )
    if not aux_by_fuel.empty:
        merge_keys = group_cols + ["sheet_name", "fuel_label"]
        out = out.merge(aux_by_fuel, on=merge_keys, how="left")
        out["aux_own_use_fuel"] = pd.to_numeric(out["aux_own_use_fuel"], errors="coerce")
        out["has_aux_own_use_fuel"] = out["aux_own_use_fuel"].notna()
        group_keys = group_cols + ["sheet_name"]
        out["group_has_aux_own_use_fuel"] = out.groupby(group_keys, dropna=False)["has_aux_own_use_fuel"].transform("any")
        out["group_aux_own_use_provided"] = (
            out["aux_own_use_fuel"].fillna(0.0).groupby([out[col] for col in group_keys], dropna=False).transform("sum")
        )
        out["missing_share_component"] = out["input_fuel_share"].where(~out["has_aux_own_use_fuel"], 0.0)
        out["group_missing_share_sum"] = (
            out["missing_share_component"].fillna(0.0).groupby([out[col] for col in group_keys], dropna=False).transform("sum")
        )
        out["group_own_use_remainder"] = out["own_use_total"] - out["group_aux_own_use_provided"]

        group_mask = out["group_has_aux_own_use_fuel"].fillna(False)
        provided_mask = group_mask & out["has_aux_own_use_fuel"].fillna(False)
        missing_mask = group_mask & ~out["has_aux_own_use_fuel"].fillna(False)
        out.loc[provided_mask, "own_use_allocated"] = pd.to_numeric(out.loc[provided_mask, "aux_own_use_fuel"], errors="coerce").fillna(0.0)
        remainder_mask = missing_mask & pd.to_numeric(out["group_missing_share_sum"], errors="coerce").gt(0)
        out.loc[remainder_mask, "own_use_allocated"] = (
            pd.to_numeric(out.loc[remainder_mask, "group_own_use_remainder"], errors="coerce")
            * pd.to_numeric(out.loc[remainder_mask, "input_fuel_share"], errors="coerce")
            / pd.to_numeric(out.loc[remainder_mask, "group_missing_share_sum"], errors="coerce")
        )
        out.loc[missing_mask & ~remainder_mask, "own_use_allocated"] = 0.0

        out.loc[provided_mask, "split_source"] = "aux_own_use_file_fuel"
        out.loc[missing_mask, "split_source"] = "aux_own_use_partial_fallback_input_share"
    else:
        out["aux_own_use_fuel"] = pd.NA
        out["has_aux_own_use_fuel"] = False
        out["group_has_aux_own_use_fuel"] = False
        out["group_aux_own_use_provided"] = 0.0
        out["group_missing_share_sum"] = out["input_fuel_share"]
        out["group_own_use_remainder"] = out["own_use_total"]

    out["losses_allocated"] = pd.to_numeric(out["total_allocated"], errors="coerce") - pd.to_numeric(
        out["own_use_allocated"],
        errors="coerce",
    )
    out["own_use_allocated"] = pd.to_numeric(out["own_use_allocated"], errors="coerce")
    out["losses_allocated"] = pd.to_numeric(out["losses_allocated"], errors="coerce")
    return out


def _build_split_totals(
    *,
    pivot: pd.DataFrame,
    spec: DerivedTransformationSpec,
    group_cols: list[str],
    total_col: str,
    data_dir: Path | None = None,
    economy_token: str | None = None,
) -> pd.DataFrame:
    if pivot.empty:
        return pd.DataFrame(columns=group_cols + ["sheet_name", "own_use_total", "losses_total", total_col, "own_use_share_effective", "split_source"])
    out = pivot[group_cols + [total_col]].copy()
    out["sheet_name"] = spec.sheet_name
    out[total_col] = pd.to_numeric(out[total_col], errors="coerce")
    aux_totals = _load_aux_own_use_totals(
        data_dir=data_dir,
        economy_token=economy_token,
        expected_sheet_names={spec.sheet_name},
        group_cols=group_cols,
    )
    if not aux_totals.empty:
        out = out.merge(aux_totals, on=group_cols + ["sheet_name"], how="left")
    else:
        out["own_use_total"] = pd.NA
    shares = _load_dummy_own_use_shares()
    own_use_share = shares.get(spec.sheet_name, shares.get("*", 0.25))
    missing_mask = pd.to_numeric(out["own_use_total"], errors="coerce").isna()
    out.loc[missing_mask, "own_use_total"] = pd.to_numeric(out.loc[missing_mask, total_col], errors="coerce") * float(own_use_share)
    out["own_use_total"] = pd.to_numeric(out["own_use_total"], errors="coerce")
    out["losses_total"] = pd.to_numeric(out[total_col], errors="coerce") - out["own_use_total"]
    out["own_use_share_effective"] = 0.0
    total_abs = pd.to_numeric(out[total_col], errors="coerce").abs()
    nonzero_mask = total_abs.gt(0) & pd.to_numeric(out[total_col], errors="coerce").notna()
    out.loc[nonzero_mask, "own_use_share_effective"] = (
        pd.to_numeric(out.loc[nonzero_mask, "own_use_total"], errors="coerce")
        / pd.to_numeric(out.loc[nonzero_mask, total_col], errors="coerce")
    )
    out["split_source"] = "dummy_share"
    out.loc[~missing_mask, "split_source"] = "aux_own_use_file"
    return out


def _choose_output_sheet_and_pivot(
    frame: pd.DataFrame,
    *,
    input_sheet: str,
    output_candidates: list[str],
    group_cols: list[str],
    value_col: str,
    source_col: str | None = None,
) -> tuple[str | None, pd.DataFrame]:
    best_sheet: str | None = None
    best_pivot = pd.DataFrame()
    best_score: tuple[int, int, int] | None = None
    preferred_rank = {name: idx for idx, name in enumerate(output_candidates)}
    for output_sheet in output_candidates:
        if not output_sheet:
            continue
        subset = frame[frame["sheet_name"].astype(str).isin([input_sheet, output_sheet])].copy()
        if subset.empty:
            continue
        pivot = _derive_gap_frame(
            subset,
            input_col=input_sheet,
            output_col=output_sheet,
            group_cols=group_cols,
            value_col=value_col,
        )
        if input_sheet not in pivot.columns or output_sheet not in pivot.columns:
            continue
        input_values = pd.to_numeric(pivot[input_sheet], errors="coerce")
        output_values = pd.to_numeric(pivot[output_sheet], errors="coerce")
        pair_mask = input_values.notna() & output_values.notna()
        non_leap_pairs = 0
        if source_col and source_col in pivot.columns:
            src = pivot[source_col].astype(str)
            non_leap_pairs = int((pair_mask & src.ne("leap")).sum())
        score = (
            non_leap_pairs,
            int(pair_mask.sum()),
            -preferred_rank.get(output_sheet, 999),
        )
        if best_score is None or score > best_score:
            best_score = score
            best_sheet = output_sheet
            pivot = pivot.copy()
            pivot["derived_value"] = input_values - output_values
            pivot["_pair_complete"] = pair_mask
            best_pivot = pivot
    return best_sheet, best_pivot


def _spec_prefix(spec: DerivedTransformationSpec) -> str:
    name = str(spec.sheet_name or "").strip()
    suffix = "_loss_own_use_total"
    if name.endswith(suffix):
        return name[: -len(suffix)]
    return name


def _aux_sheet_candidates(prefix: str) -> tuple[str, str]:
    return f"{prefix}_aux_outputs", f"{prefix}_aux_other_outputs"


def _discover_transformation_workbooks(data_dir: Path | None, economy_token: str | None) -> list[Path]:
    if data_dir is None or not str(economy_token or "").strip():
        return []
    token = str(economy_token).strip().upper()
    root = Path(data_dir)
    if not root.exists():
        return []
    out: list[Path] = []
    for path in sorted(root.glob(f"transformation_results_*_{token}_*.xls*")):
        if path.name.startswith("~$"):
            continue
        out.append(path)
    return out


def _read_aux_sheet_frame(workbook: Path, sheet_name: str, mapped_sheet_name: str) -> pd.DataFrame:
    xl = pd.ExcelFile(workbook)
    if sheet_name not in set(xl.sheet_names):
        return pd.DataFrame()
    parsed = parse_template_sheet(xl.parse(sheet_name, header=None))
    records = parsed.get("records", pd.DataFrame()).copy()
    if records.empty:
        return pd.DataFrame()
    records["fuel_label"] = records["fuel_label"].astype(str).str.strip()
    records = records[records["fuel_label"].ne("") & ~records["fuel_label"].str.lower().eq("total")].copy()
    if records.empty:
        return pd.DataFrame()
    meta = parsed.get("meta", {})
    scenario = str(meta.get("scenario") or "").strip()
    region = str(meta.get("region") or "").strip()
    economy = ""
    match = re.search(r"_([A-Z]{3})_", workbook.name)
    if match:
        economy = match.group(1)
    records["economy"] = economy or region
    records["scenario"] = scenario
    records["region"] = region
    records["sheet_name"] = mapped_sheet_name
    records["source_sheet_name"] = sheet_name
    records["leap_value"] = pd.to_numeric(records["leap_value"], errors="coerce")
    return records[["economy", "scenario", "region", "sheet_name", "source_sheet_name", "fuel_label", "year", "leap_value"]]


def _build_aux_direct_leap_long(
    *,
    specs: tuple[DerivedTransformationSpec, ...],
    data_dir: Path | None,
    economy_token: str | None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    workbooks = _discover_transformation_workbooks(data_dir, economy_token)
    if not workbooks:
        empty = pd.DataFrame(columns=LEAP_LONG_COLUMNS)
        return empty, pd.DataFrame(), pd.DataFrame()

    rows: list[dict[str, object]] = []
    audit_rows: list[dict[str, object]] = []
    summary_rows: list[dict[str, object]] = []
    for spec in specs:
        prefix = _spec_prefix(spec)
        aux_sheet, aux_other_sheet = _aux_sheet_candidates(prefix)
        component_frames: list[pd.DataFrame] = []
        for workbook in workbooks:
            for candidate in (aux_sheet, aux_other_sheet, f"{prefix}_aux_other_out"):
                component = _read_aux_sheet_frame(workbook, candidate, spec.sheet_name)
                if component.empty:
                    continue
                component["aux_component"] = candidate
                component_frames.append(component)
        if not component_frames:
            continue
        combined = pd.concat(component_frames, ignore_index=True, sort=False)
        grouped = (
            combined.groupby(["economy", "scenario", "region", "sheet_name", "fuel_label", "year"], as_index=False)["leap_value"]
            .sum(min_count=1)
            .rename(columns={"leap_value": "own_use_value"})
        )
        if grouped.empty:
            continue
        grouped["losses_value"] = 0.0
        for row in grouped.itertuples(index=False):
            rows.append(
                {
                    "economy": getattr(row, "economy", ""),
                    "scenario": getattr(row, "scenario", ""),
                    "region": getattr(row, "region", ""),
                    "sheet_name": spec.sheet_name,
                    "sector_code_9th": spec.sector_code_9th,
                    "sector_name": spec.sector_name,
                    "fuel_label": getattr(row, "fuel_label", ""),
                    "year": getattr(row, "year", pd.NA),
                    "leap_value": getattr(row, "own_use_value", pd.NA),
                    "leap_variable": "Transformation auxiliary fuels (direct)",
                    "leap_units": "Petajoules",
                    "measure": MEASURE_OWN_USE.measure,
                    "leap_scale_note": "Derived directly from LEAP transformation auxiliary output sheets.",
                }
            )
            rows.append(
                {
                    "economy": getattr(row, "economy", ""),
                    "scenario": getattr(row, "scenario", ""),
                    "region": getattr(row, "region", ""),
                    "sheet_name": spec.sheet_name,
                    "sector_code_9th": spec.sector_code_9th,
                    "sector_name": spec.sector_name,
                    "fuel_label": getattr(row, "fuel_label", ""),
                    "year": getattr(row, "year", pd.NA),
                    "leap_value": 0.0,
                    "leap_variable": "Transformation auxiliary fuels (direct)",
                    "leap_units": "Petajoules",
                    "measure": MEASURE_LOSSES.measure,
                    "leap_scale_note": "Losses set to zero in aux-direct mode; own-use from auxiliary output sheets.",
                }
            )
        source_totals = (
            combined.groupby(["economy", "scenario", "region", "year", "source_sheet_name"], as_index=False)["leap_value"]
            .sum(min_count=1)
        )
        for row in source_totals.itertuples(index=False):
            audit_rows.append(
                {
                    "sheet_name": spec.sheet_name,
                    "sector_name": spec.sector_name,
                    "scenario": getattr(row, "scenario", ""),
                    "region": getattr(row, "region", ""),
                    "year": getattr(row, "year", pd.NA),
                    "input_sheet": spec.input_sheet,
                    "output_sheet": getattr(row, "source_sheet_name", ""),
                    "derived_total_value": getattr(row, "leap_value", pd.NA),
                    "formula": f"{aux_sheet} + {aux_other_sheet}",
                    "viability": spec.viability,
                    "rationale": "Auxiliary own-use taken directly from LEAP auxiliary output sheets.",
                }
            )
        summary_rows.append(
            {
                "sheet_name": spec.sheet_name,
                "sector_name": spec.sector_name,
                "sector_code_9th": spec.sector_code_9th,
                "input_sheet": spec.input_sheet,
                "output_sheet": f"{aux_sheet}; {aux_other_sheet}",
                "formula": f"{aux_sheet} + {aux_other_sheet}",
                "measures": f"{MEASURE_OWN_USE.measure}; {MEASURE_LOSSES.measure}",
                "allocation": "Own use from auxiliary output sheets by fuel; losses fixed at zero in aux-direct mode.",
                "viability": spec.viability,
                "rationale": "Uses LEAP auxiliary output tables directly instead of input-output gap split.",
            }
        )
    return pd.DataFrame(rows, columns=LEAP_LONG_COLUMNS), pd.DataFrame(audit_rows), pd.DataFrame(summary_rows)


def build_derived_transformation_leap_long(
    leap_long: pd.DataFrame,
    *,
    sheet_map: pd.DataFrame,
    data_dir: Path | None = None,
    economy_token: str | None = None,
    derivation_mode: str = "aux_direct",
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if leap_long.empty:
        empty = pd.DataFrame(columns=LEAP_LONG_COLUMNS)
        return empty, pd.DataFrame(), pd.DataFrame()
    specs = derive_transformation_specs(sheet_map)
    if not specs:
        empty = pd.DataFrame(columns=LEAP_LONG_COLUMNS)
        return empty, pd.DataFrame(), pd.DataFrame()
    mode = str(derivation_mode or "").strip().lower()
    if mode == "aux_direct":
        aux_long, aux_audit, aux_summary = _build_aux_direct_leap_long(
            specs=specs,
            data_dir=data_dir,
            economy_token=economy_token,
        )
        if not aux_long.empty:
            return aux_long, aux_audit, aux_summary

    rows: list[dict[str, object]] = []
    audit_rows: list[dict[str, object]] = []
    summary_rows: list[dict[str, object]] = []
    for spec in specs:
        output_candidates = [spec.output_sheet]
        if spec.output_feed_sheet:
            output_candidates.append(spec.output_feed_sheet)
        subset = leap_long[
            leap_long["sheet_name"].astype(str).isin([spec.input_sheet] + output_candidates)
        ].copy()
        if subset.empty:
            continue
        output_sheet_used, pivot = _choose_output_sheet_and_pivot(
            subset.rename(columns={"sheet_name": "sheet_name", "leap_value": "leap_value"}),
            input_sheet=spec.input_sheet,
            output_candidates=output_candidates,
            group_cols=["economy", "scenario", "region", "year"],
            value_col="leap_value",
        )
        if output_sheet_used is None or pivot.empty:
            continue
        pivot["derived_value"] = pd.to_numeric(pivot["derived_value"], errors="coerce")
        pivot.loc[~pivot["_pair_complete"].fillna(False), "derived_value"] = pd.NA
        totals = _build_split_totals(
            pivot=pivot,
            spec=spec,
            group_cols=["economy", "scenario", "region", "year"],
            total_col="derived_value",
            data_dir=data_dir,
            economy_token=economy_token,
        )
        shares = _build_input_fuel_shares(
            subset,
            sheet_col="sheet_name",
            value_col="leap_value",
            input_sheet=spec.input_sheet,
            group_cols=["economy", "scenario", "region", "year"],
            fuel_col="fuel_label",
        )
        if shares.empty:
            continue
        allocated = totals.merge(shares, on=["economy", "scenario", "region", "year"], how="inner")
        allocated = _apply_own_use_fuel_allocation(
            allocated=allocated,
            group_cols=["economy", "scenario", "region", "year"],
            data_dir=data_dir,
            economy_token=economy_token,
        )
        allocated = allocated[allocated["derived_value"].notna()].copy()
        if allocated.empty:
            continue
        for row in allocated.itertuples(index=False):
            for split in (MEASURE_OWN_USE, MEASURE_LOSSES):
                value_field = "own_use_allocated" if split.key == "own_use" else "losses_allocated"
                rows.append(
                    {
                        "economy": getattr(row, "economy", ""),
                        "scenario": getattr(row, "scenario", ""),
                        "region": getattr(row, "region", ""),
                        "sheet_name": spec.sheet_name,
                        "sector_code_9th": spec.sector_code_9th,
                        "sector_name": spec.sector_name,
                        "fuel_label": getattr(row, "fuel_label", ""),
                        "year": getattr(row, "year", pd.NA),
                        "leap_value": getattr(row, value_field, pd.NA),
                        "leap_variable": "Calculated transformation metric",
                        "leap_units": "Petajoules",
                        "measure": split.measure,
                        "leap_scale_note": (
                            f"Base gap: {spec.input_sheet} minus {output_sheet_used}. "
                            f"Output component used: {output_sheet_used}. "
                            "Gap split into own use and losses using auxiliary-own-use values when provided, "
                            "otherwise dummy own-use shares; then allocated to input fuels by absolute input use."
                        ),
                    }
                )
            audit_rows.append(
                {
                    "sheet_name": spec.sheet_name,
                    "sector_name": spec.sector_name,
                    "scenario": getattr(row, "scenario", ""),
                    "region": getattr(row, "region", ""),
                    "year": getattr(row, "year", pd.NA),
                    "input_sheet": spec.input_sheet,
                    "output_sheet": output_sheet_used,
                    "input_value": getattr(row, spec.input_sheet, pd.NA),
                    "output_value": getattr(row, output_sheet_used, pd.NA),
                    "derived_value": getattr(row, "derived_value", pd.NA),
                    "input_fuel_label": getattr(row, "fuel_label", ""),
                    "input_fuel_value": getattr(row, "input_fuel_value", pd.NA),
                    "input_fuel_share": getattr(row, "input_fuel_share", pd.NA),
                    "allocation_method": getattr(row, "allocation_method", ""),
                        "split_source": getattr(row, "split_source", ""),
                        "aux_own_use_fuel": getattr(row, "aux_own_use_fuel", pd.NA),
                        "derived_total_value": getattr(row, "derived_value", pd.NA),
                    "own_use_total": getattr(row, "own_use_total", pd.NA),
                    "losses_total": getattr(row, "losses_total", pd.NA),
                    "own_use_share_effective": getattr(row, "own_use_share_effective", pd.NA),
                    "own_use_allocated": getattr(row, "own_use_allocated", pd.NA),
                    "losses_allocated": getattr(row, "losses_allocated", pd.NA),
                    "formula": f"{spec.input_sheet} - {spec.output_sheet}",
                    "viability": spec.viability,
                    "rationale": spec.rationale,
                }
            )
        summary_rows.append(
            {
                "sheet_name": spec.sheet_name,
                "sector_name": spec.sector_name,
                "sector_code_9th": spec.sector_code_9th,
                "input_sheet": spec.input_sheet,
                "output_sheet": output_sheet_used,
                "formula": f"{spec.input_sheet} - {output_sheet_used}",
                "measures": f"{MEASURE_OWN_USE.measure}; {MEASURE_LOSSES.measure}",
                "allocation": "Split to own use/losses then allocated across input fuels by absolute input use share per year",
                "viability": spec.viability,
                "rationale": spec.rationale,
            }
        )

    derived_leap_long = pd.DataFrame(rows, columns=LEAP_LONG_COLUMNS)
    derived_audit = pd.DataFrame(audit_rows)
    derived_summary = pd.DataFrame(summary_rows)
    return derived_leap_long, derived_audit, derived_summary


def build_derived_transformation_comparison_rows(
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    *,
    sheet_map: pd.DataFrame,
    data_dir: Path | None = None,
    economy_token: str | None = None,
    base_df: pd.DataFrame | None = None,
    ninth_df: pd.DataFrame | None = None,
    base_year: int | None = None,
    projection_years: tuple[int, ...] | None = None,
    derivation_mode: str = "aux_direct",
    derived_leap_long: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if comparison_long.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    specs = derive_transformation_specs(sheet_map)
    if not specs:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    mode = str(derivation_mode or "").strip().lower()

    comparison_sources = {
        "leap",
        "base",
        "base_estimated",
        "base_mixed",
        "projection",
        "projection_estimated",
        "projection_mixed",
    }
    rows: list[dict[str, object]] = []
    status_rows: list[dict[str, object]] = []
    audit_rows: list[dict[str, object]] = []
    sector_to_flow = build_sector_to_esto_flow_lookup() if base_df is not None else {}
    base_cache: dict[tuple[str, str], float] = {}
    projection_cache: dict[tuple[str, str, str], pd.Series] = {}
    leap_direct = (
        derived_leap_long.copy()
        if isinstance(derived_leap_long, pd.DataFrame) and not derived_leap_long.empty
        else pd.DataFrame()
    )

    for spec in specs:
        output_candidates = [spec.output_sheet]
        if spec.output_feed_sheet:
            output_candidates.append(spec.output_feed_sheet)
        subset = comparison_long[
            comparison_long["sheet"].astype(str).isin([spec.input_sheet] + output_candidates)
            & comparison_long["source"].astype(str).isin(comparison_sources)
        ].copy()
        if subset.empty:
            continue
        subset_norm = subset.rename(columns={"sheet": "sheet_name", "value": "series_value"})
        output_sheet_used = spec.output_sheet
        if mode == "aux_direct":
            totals = pd.DataFrame(columns=["economy", "scenario", "source", "year", "sheet_name", "derived_value", "own_use_total", "losses_total", "own_use_share_effective", "split_source"])
        else:
            output_sheet_used, pivot = _choose_output_sheet_and_pivot(
                subset_norm,
                input_sheet=spec.input_sheet,
                output_candidates=output_candidates,
                group_cols=["economy", "scenario", "source", "year"],
                value_col="series_value",
            )
            if output_sheet_used is None or pivot.empty:
                continue
            pivot["derived_value"] = pd.to_numeric(pivot["derived_value"], errors="coerce")
            pivot.loc[~pivot["_pair_complete"].fillna(False), "derived_value"] = pd.NA
            totals = _build_split_totals(
                pivot=pivot,
                spec=spec,
                group_cols=["economy", "scenario", "source", "year"],
                total_col="derived_value",
                data_dir=data_dir,
                economy_token=economy_token,
            )
        if base_df is not None or ninth_df is not None:
            totals = totals.copy()
            if mode == "aux_direct":
                keys = subset_norm[subset_norm["sheet_name"].astype(str).eq(spec.input_sheet)][
                    ["economy", "scenario", "source", "year"]
                ].drop_duplicates()
                if not keys.empty:
                    keys["sheet_name"] = spec.sheet_name
                    keys["derived_value"] = pd.NA
                    keys["own_use_total"] = pd.NA
                    keys["losses_total"] = 0.0
                    keys["own_use_share_effective"] = pd.NA
                    keys["split_source"] = ""
                    totals = keys
            for idx, row in totals.iterrows():
                source = str(row.get("source") or "").strip().lower()
                economy = str(row.get("economy") or "").strip()
                scenario = str(row.get("scenario") or "").strip()
                scenario_token = scenario.lower()
                year = row.get("year")
                if source in {"base", "base_estimated", "base_mixed"} and base_df is not None:
                    if base_year is None:
                        continue
                    flow = sector_to_flow.get(str(spec.sector_code_9th or "").strip().lower(), "")
                    if not flow:
                        continue
                    cache_key = (economy, flow)
                    if cache_key not in base_cache:
                        base_cache[cache_key] = pull_base_year_value(
                            base_df,
                            base_year=base_year,
                            economy_code=economy,
                            esto_flow=flow,
                            esto_product="",
                        )
                    value = base_cache[cache_key]
                elif source in {"projection", "projection_estimated", "projection_mixed"} and ninth_df is not None:
                    if projection_years is None:
                        continue
                    sector_code = str(spec.sector_code_9th or "").strip()
                    if not sector_code:
                        continue
                    cache_key = (economy, scenario_token, sector_code)
                    if cache_key not in projection_cache:
                        projection_cache[cache_key] = pull_projection_series(
                            ninth_df,
                            sector_code=sector_code,
                            fuel_code="",
                            economy_code=economy,
                            scenario=scenario_token,
                            projection_years=projection_years,
                        )
                    series = projection_cache[cache_key]
                    value = series.get(int(year)) if year is not None else pd.NA
                else:
                    continue
                if pd.isna(value):
                    continue
                totals.at[idx, "derived_value"] = value
                totals.at[idx, "own_use_total"] = value
                totals.at[idx, "losses_total"] = 0.0
                totals.at[idx, "own_use_share_effective"] = 1.0
                totals.at[idx, "split_source"] = "reference_own_use"
        shares = _build_input_fuel_shares(
            subset_norm,
            sheet_col="sheet_name",
            value_col="series_value",
            input_sheet=spec.input_sheet,
            group_cols=["economy", "scenario", "source", "year"],
            fuel_col="fuel_label",
        )
        if shares.empty:
            continue
        allocated = totals.merge(shares, on=["economy", "scenario", "source", "year"], how="inner")
        if mode == "aux_direct":
            allocated["derived_value"] = pd.to_numeric(allocated["derived_value"], errors="coerce")
            allocated["own_use_total"] = pd.to_numeric(allocated["own_use_total"], errors="coerce")
            allocated["own_use_allocated"] = allocated["own_use_total"] * pd.to_numeric(allocated["input_fuel_share"], errors="coerce").fillna(0.0)
            allocated["losses_allocated"] = 0.0
            allocated["split_source"] = "reference_own_use"
            allocated["aux_own_use_fuel"] = pd.NA
        else:
            allocated = _apply_own_use_fuel_allocation(
                allocated=allocated,
                group_cols=["economy", "scenario", "source", "year"],
                data_dir=data_dir,
                economy_token=economy_token,
            )
        allocated = allocated[allocated["derived_value"].notna()].copy()
        if not allocated.empty:
            for row in allocated.itertuples(index=False):
                for split in (MEASURE_OWN_USE, MEASURE_LOSSES):
                    value_field = "own_use_allocated" if split.key == "own_use" else "losses_allocated"
                    rows.append(
                        {
                            "economy": getattr(row, "economy", ""),
                            "scenario": getattr(row, "scenario", ""),
                            "sheet": spec.sheet_name,
                            "measure": split.measure,
                            "fuel_label": getattr(row, "fuel_label", ""),
                            "source": getattr(row, "source", ""),
                            "year": getattr(row, "year", pd.NA),
                            "value": getattr(row, value_field, pd.NA),
                        }
                    )
                audit_rows.append(
                    {
                        "sheet_name": spec.sheet_name,
                        "sector_name": spec.sector_name,
                        "scenario": getattr(row, "scenario", ""),
                        "source": getattr(row, "source", ""),
                        "year": getattr(row, "year", pd.NA),
                        "input_sheet": spec.input_sheet,
                        "output_sheet": output_sheet_used,
                        "input_value": getattr(row, spec.input_sheet, pd.NA),
                        "output_value": getattr(row, output_sheet_used, pd.NA),
                        "derived_value": getattr(row, "derived_value", pd.NA),
                        "input_fuel_label": getattr(row, "fuel_label", ""),
                        "input_fuel_value": getattr(row, "input_fuel_value", pd.NA),
                        "input_fuel_share": getattr(row, "input_fuel_share", pd.NA),
                        "allocation_method": getattr(row, "allocation_method", ""),
                        "split_source": getattr(row, "split_source", ""),
                        "aux_own_use_fuel": getattr(row, "aux_own_use_fuel", pd.NA),
                        "derived_total_value": getattr(row, "derived_value", pd.NA),
                        "own_use_total": getattr(row, "own_use_total", pd.NA),
                        "losses_total": getattr(row, "losses_total", pd.NA),
                        "own_use_share_effective": getattr(row, "own_use_share_effective", pd.NA),
                        "own_use_allocated": getattr(row, "own_use_allocated", pd.NA),
                        "losses_allocated": getattr(row, "losses_allocated", pd.NA),
                        "formula": (
                            f"{_spec_prefix(spec)}_aux_outputs + {_spec_prefix(spec)}_aux_other_outputs"
                            if mode == "aux_direct"
                            else f"{spec.input_sheet} - {output_sheet_used}"
                        ),
                        "viability": spec.viability,
                        "rationale": spec.rationale,
                    }
                )

        leap_subset = pd.DataFrame()
        if mode == "aux_direct" and not leap_direct.empty:
            leap_subset = leap_direct[
                leap_direct["sheet_name"].astype(str).eq(spec.sheet_name)
            ].copy()
            if not leap_subset.empty:
                leap_subset["leap_value"] = pd.to_numeric(leap_subset["leap_value"], errors="coerce")
                leap_rows = (
                    leap_subset.groupby(
                        ["economy", "scenario", "sheet_name", "measure", "fuel_label", "year"],
                        as_index=False,
                    )["leap_value"]
                    .sum(min_count=1)
                )
                for row in leap_rows.itertuples(index=False):
                    rows.append(
                        {
                            "economy": getattr(row, "economy", ""),
                            "scenario": getattr(row, "scenario", ""),
                            "sheet": spec.sheet_name,
                            "measure": getattr(row, "measure", ""),
                            "fuel_label": getattr(row, "fuel_label", ""),
                            "source": "leap",
                            "year": getattr(row, "year", pd.NA),
                            "value": getattr(row, "leap_value", pd.NA),
                        }
                    )

        spec_rows = allocated.copy()
        source_values = set(spec_rows["source"].dropna().astype(str))
        base_complete = bool(source_values & {"base", "base_estimated", "base_mixed"})
        projection_complete = bool(source_values & {"projection", "projection_estimated", "projection_mixed"})
        has_any_mapping = base_complete or projection_complete
        fuels = {str(v).strip() for v in spec_rows["fuel_label"].dropna().tolist() if str(v).strip()}
        if not leap_subset.empty:
            fuels |= {str(v).strip() for v in leap_subset["fuel_label"].dropna().tolist() if str(v).strip()}
        fuels = sorted(fuels)
        for fuel in fuels:
            for split in (MEASURE_OWN_USE, MEASURE_LOSSES):
                status_rows.append(
                    {
                        "sheet": spec.sheet_name,
                        "fuel_label": fuel,
                        "measure": split.measure,
                        "sector_code_9th": spec.sector_code_9th,
                        "ninth_fuel_code": "",
                        "esto_flow": "",
                        "esto_product": "",
                        "has_any_mapping": has_any_mapping,
                        "base_mapping_complete": base_complete,
                        "projection_mapping_complete": projection_complete,
                        "partially_mapped": has_any_mapping and not (base_complete and projection_complete),
                        "mapped": base_complete and projection_complete,
                        "mapping_source": "derived_from_component_sheets",
                        "flow_source": "derived",
                        "fuel_source": "derived_from_input_share",
                        "sector_match_method": "component_sheet_gap",
                    "mapping_note": (
                        "Aux-direct mode: LEAP own-use rows come directly from transformation auxiliary output sheets; "
                        "base/projection comparator totals use the 9th/ESTO own-use sector and are allocated across "
                        "input fuels by absolute input shares (losses fixed at 0)."
                        if mode == "aux_direct"
                        else (
                            f"Calculated from mapped component sheets: {spec.input_sheet} minus {output_sheet_used}, "
                            "split to own use/losses using auxiliary-own-use values when available (else dummy share), "
                            "then allocated to input fuels using absolute input shares. Comparator totals for base/projection "
                            "sources use the 9th/ESTO own-use sector when available (losses set to 0)."
                        )
                    ),
                        "projection_fuel_filter": "",
                        "projection_fuel_codes_detail": "",
                        "projection_targets_detail": "",
                        "base_targets_detail": "",
                        "projection_parent_fallback": False,
                        "projection_parent_sector_code": "",
                        "comparator_scope": "child",
                        "base_mapping_optional": False,
                    }
                )

    derived_comparison = pd.DataFrame(rows)
    derived_status = pd.DataFrame(status_rows)
    derived_audit = pd.DataFrame(audit_rows)
    if not derived_comparison.empty:
        derived_comparison = derived_comparison.sort_values(
            ["sheet", "measure", "fuel_label", "scenario", "source", "year"],
            kind="mergesort",
        ).reset_index(drop=True)
    if not derived_status.empty and not mapping_status.empty:
        for col in mapping_status.columns:
            if col not in derived_status.columns:
                derived_status[col] = pd.NA
        derived_status = derived_status[mapping_status.columns]
    return derived_comparison, derived_status, derived_audit


def write_derived_transformation_artifacts(
    *,
    data_dir: Path,
    out_dir: Path,
    economy_token: str,
    derived_leap_long: pd.DataFrame,
    leap_audit: pd.DataFrame,
    comparison_audit: pd.DataFrame,
    summary: pd.DataFrame,
) -> dict[str, str | None]:
    out_dir.mkdir(parents=True, exist_ok=True)
    data_dir.mkdir(parents=True, exist_ok=True)
    processed_dir = _processed_tables_dir(data_dir)
    if processed_dir is not None:
        processed_dir.mkdir(parents=True, exist_ok=True)
    data_path = (
        processed_dir / f"transformation_derived_metrics_20_{economy_token.upper()}.csv"
        if processed_dir is not None
        else data_dir / f"transformation_derived_metrics_20_{economy_token.upper()}.csv"
    )
    fallback_data_path = out_dir / f"transformation_derived_metrics_20_{economy_token.upper()}.csv"
    leap_audit_path = out_dir / "derived_transformation_leap_audit.csv"
    comparison_audit_path = out_dir / "derived_transformation_comparison_audit.csv"
    summary_path = out_dir / "derived_transformation_metric_assessment.csv"

    written_data_path = data_path
    try:
        if derived_leap_long.empty:
            data_path.write_text("", encoding="utf-8")
        else:
            derived_leap_long.to_csv(data_path, index=False)
    except PermissionError:
        written_data_path = fallback_data_path
        if derived_leap_long.empty:
            fallback_data_path.write_text("", encoding="utf-8")
        else:
            derived_leap_long.to_csv(fallback_data_path, index=False)
    if leap_audit.empty:
        leap_audit_path.write_text("", encoding="utf-8")
    else:
        leap_audit.to_csv(leap_audit_path, index=False)
    if comparison_audit.empty:
        comparison_audit_path.write_text("", encoding="utf-8")
    else:
        comparison_audit.to_csv(comparison_audit_path, index=False)
    if summary.empty:
        summary_path.write_text("", encoding="utf-8")
    else:
        summary.to_csv(summary_path, index=False)

    return {
        "derived_transformation_leap_long": str(written_data_path) if written_data_path.exists() else None,
        "derived_transformation_data_dir": str(processed_dir) if processed_dir is not None else str(data_dir),
        "derived_transformation_leap_audit": str(leap_audit_path) if leap_audit_path.exists() else None,
        "derived_transformation_comparison_audit": (
            str(comparison_audit_path) if comparison_audit_path.exists() else None
        ),
        "derived_transformation_metric_assessment": str(summary_path) if summary_path.exists() else None,
    }
