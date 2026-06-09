# %%
"""Replica workflow for LEAP results-template extraction with strict sheet controls."""
from __future__ import annotations

from pathlib import Path
import sys
import difflib
import time
import re
import shutil

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.utilities.leap_results_extraction_replica import run_replica_extraction
from codebase.utilities.leap_results_extraction_replica import apply_strict_meta
from codebase.utilities.leap_results_extraction_replica import is_effectively_empty_results_table
from codebase import leap_results_workflow


def _resolve(path: Path | str) -> Path:
    raw = str(path).replace("\\", "/")
    candidate = Path(raw)
    return candidate if candidate.is_absolute() else (REPO_ROOT / candidate)


# Replica scope: start with transformation tables only.
TEMPLATE_PATHS = [
    Path("data/leap results tables/transformation_results_20_USA_Target.xlsx"),
    Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
]

# Modes:
# - replay_csv: rebuild sheets from previously exported CSVs in outputs/leap_results/_tmp_template_exports
# - live: call LEAP API and export fresh CSVs for each sheet
MODE = "replay_csv"

# Optional: provide a golden workbook path to compare sheet-by-sheet.
# Set to your target workbook for automatic comparison report.
GOLDEN_WORKBOOK_PATH: Path | None = Path(
    "data/leap results tables/archive/transformation_results_20_USA_Reference - golden.xlsx"
)

OUTPUT_DIR = Path("outputs/leap_results_replica")
SUPPORTING_DIR = OUTPUT_DIR / "supporting_files"


def _try_value_rs(variable_obj, *, region: str, scenario: str, year: int, filter_str: str) -> tuple[bool, float | None, str]:
    try:
        value = variable_obj.ValueRS(region, scenario, int(year), "", filter_str)
    except Exception as exc:  # noqa: BLE001
        return False, None, str(exc)
    numeric = leap_results_workflow.pd.to_numeric(value, errors="coerce")
    if leap_results_workflow.pd.isna(numeric):
        return True, None, ""
    return True, float(numeric), ""


def run_single_input_type_probe(
    *,
    template_path: Path | str = Path("data/leap results tables/transformation_results_20_USA_Target.xlsx"),
    sheet_name: str = "elecgen_inputs",
    probe_year: int = 2022,
    sample_fuel: str = "Natural gas",
) -> None:
    """
    Probe whether LEAP needs an extra Input Type filter (e.g., All Input Types)
    for transformation input extraction.
    """
    template_path = _resolve(template_path)
    wb = leap_results_workflow.load_workbook(template_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found in {template_path}")
    ws = wb[sheet_name]
    template_meta = leap_results_workflow.parse_template_worksheet(ws)
    meta = apply_strict_meta(template_path, sheet_name, template_meta)

    app = leap_results_workflow.connect_leap()
    leap_results_workflow.ensure_calculated(app, force=False)
    leap_results_workflow._show_results_view_table(app)
    leap_results_workflow.set_axes(app, x_axis="Years", legend=str(meta.get("legend_label") or "Fuel"))
    leap_results_workflow.set_context(
        app,
        scenario=meta.get("scenario"),
        region=meta.get("region"),
        branch_path=meta.get("branch"),
    )
    variable_obj = leap_results_workflow._resolve_branch_variable(
        app,
        str(meta.get("branch") or ""),
        meta.get("variable"),
        allow_substitution=False,
    )

    print("\n" + "=" * 90)
    print("Single Input-Type Probe")
    print("=" * 90)
    print(f"Template: {template_path.name} / {sheet_name}")
    print(f"Scenario: {meta.get('scenario')}, Region: {meta.get('region')}, Year: {probe_year}")
    print(f"Branch: {meta.get('branch')}")
    print(f"Variable: {getattr(variable_obj, 'Name', meta.get('variable'))}")
    print(f"Legend used: {meta.get('legend_label')}")

    dims = leap_results_workflow.list_dimensions(app)
    print(f"\nDimensions ({len(dims)}): {dims}")
    input_dims = [name for name in dims if "input" in str(name).strip().lower()]
    print(f"Input-like dimensions: {input_dims}")

    probe_filters = [
        f"Fuel={sample_fuel}",
        f"Input Type=All Input Types;Fuel={sample_fuel}",
        f"Input Type=All Input Types|Fuel={sample_fuel}",
        f"Input Type=All Input Types,Fuel={sample_fuel}",
        f"Inputs Type=All Input Types;Fuel={sample_fuel}",
        f"Input Types=All Input Types;Fuel={sample_fuel}",
        "Input Type=All Input Types",
        "Input Types=All Input Types",
        "Inputs Type=All Input Types",
        "",
    ]

    print("\nValueRS filter probes:")
    for filter_str in probe_filters:
        ok, value, err = _try_value_rs(
            variable_obj,
            region=str(meta.get("region") or ""),
            scenario=str(meta.get("scenario") or ""),
            year=int(probe_year),
            filter_str=filter_str,
        )
        status = "OK" if ok else "ERR"
        print(f"  [{status}] filter='{filter_str}' -> value={value} error={err}")

    # Discover actual Input Type members from LEAP and probe those exact names.
    input_dim_name, input_members = leap_results_workflow.discover_legend_members_from_api(
        app,
        "Input Type",
        preferred_dimension_names=["Input Type"],
    )
    print(f"\nResolved Input Type dimension: {input_dim_name}")
    print(f"Input Type members ({len(input_members)}): {input_members}")

    if input_members:
        print("\nValueRS probes using discovered Input Type members:")
        for member in input_members[:25]:
            # Probe member alone
            filter_only = f"Input Type={member}"
            ok1, value1, err1 = _try_value_rs(
                variable_obj,
                region=str(meta.get("region") or ""),
                scenario=str(meta.get("scenario") or ""),
                year=int(probe_year),
                filter_str=filter_only,
            )
            status1 = "OK" if ok1 else "ERR"
            print(f"  [{status1}] {filter_only!r} -> value={value1} error={err1}")

            # Probe member + fuel
            filter_with_fuel = f"Input Type={member};Fuel={sample_fuel}"
            ok2, value2, err2 = _try_value_rs(
                variable_obj,
                region=str(meta.get("region") or ""),
                scenario=str(meta.get("scenario") or ""),
                year=int(probe_year),
                filter_str=filter_with_fuel,
            )
            status2 = "OK" if ok2 else "ERR"
            print(f"  [{status2}] {filter_with_fuel!r} -> value={value2} error={err2}")

    export_csv_path = leap_results_workflow._sheet_tmp_export_path(template_path, sheet_name, 1)
    leap_results_workflow._export_results_csv_file(app, export_csv_path.resolve())
    df = leap_results_workflow.parse_exported_results_csv(export_csv_path.resolve())
    labels = df.iloc[:, 0].dropna().astype(str).str.strip().tolist() if not df.empty else []
    print(f"\nCSV export path: {export_csv_path}")
    print(f"CSV rows={len(df)}, cols={len(df.columns)}")
    print(f"First labels: {labels[:15]}")
    print("=" * 90 + "\n")


def run_single_empty_clear_probe(
    *,
    template_path: Path | str = Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
    sheet_name: str = "lng_regas_out_feed",
    probe_year: int = 2022,
) -> None:
    """
    Single-sheet probe: verify empty extraction clears stale existing values.
    """
    template_path = _resolve(template_path)
    probe_dir = _resolve(OUTPUT_DIR) / "probes"
    probe_dir.mkdir(parents=True, exist_ok=True)
    probe_workbook = probe_dir / f"{template_path.stem}.empty_clear_probe.xlsx"
    probe_workbook.write_bytes(template_path.read_bytes())

    wb = leap_results_workflow.load_workbook(probe_workbook)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found in {probe_workbook}")
    ws = wb[sheet_name]

    # Seed obviously stale values into the first two data rows so we can confirm
    # they are wiped when extraction returns no meaningful data.
    ws.cell(row=7, column=1).value = "STALE_FAKE_FUEL_A"
    ws.cell(row=7, column=2).value = 999999.0
    ws.cell(row=8, column=1).value = "STALE_FAKE_FUEL_B"
    ws.cell(row=8, column=2).value = 888888.0
    wb.save(probe_workbook)

    wb = leap_results_workflow.load_workbook(probe_workbook)
    ws = wb[sheet_name]
    template_meta = leap_results_workflow.parse_template_worksheet(ws)
    strict_meta = apply_strict_meta(probe_workbook, sheet_name, template_meta)

    app = leap_results_workflow.connect_leap()
    leap_results_workflow.ensure_calculated(app, force=False)
    export_csv_path = leap_results_workflow._sheet_tmp_export_path(probe_workbook, sheet_name, 1)
    fresh_df = leap_results_workflow.build_fresh_table_from_export(
        app,
        strict_meta,
        export_csv_path=export_csv_path,
        context=f"empty_clear_probe/{probe_workbook.name}/{sheet_name}",
        allow_variable_substitution=False,
    )

    effective_empty = is_effectively_empty_results_table(fresh_df)
    if effective_empty:
        fresh_df = leap_results_workflow._build_no_results_table(strict_meta)

    leap_results_workflow.write_table_values_preserve_format(ws, fresh_df)
    wb.save(probe_workbook)

    post = leap_results_workflow.pd.read_excel(probe_workbook, sheet_name=sheet_name, header=None)
    post_labels = post.iloc[6:, 0].fillna("").astype(str).str.strip().tolist()
    post_labels = [label for label in post_labels if label]

    print("\n" + "=" * 90)
    print("Single Empty-Clear Probe")
    print("=" * 90)
    print(f"Workbook: {probe_workbook}")
    print(f"Sheet: {sheet_name}")
    print(f"Scenario: {strict_meta.get('scenario')}, Region: {strict_meta.get('region')}, Year: {probe_year}")
    print(f"Branch: {strict_meta.get('branch')}")
    print(f"Variable: {strict_meta.get('variable')}")
    print(f"CSV source: {export_csv_path}")
    print(f"Detected effective empty extraction: {effective_empty}")
    print(f"Post-write labels: {post_labels[:15]}")
    print("Expected for clear behavior: only header (no stale fake labels).")
    print("=" * 90 + "\n")


def run_handle_search_probe(
    *,
    template_path: Path | str = Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
    sheet_name: str = "elecgen_inputs",
    keyword: str = "input type",
) -> None:
    """
    Inspect accessible LEAP COM handles and fuzzy-near names for a keyword.
    """
    template_path = _resolve(template_path)
    wb = leap_results_workflow.load_workbook(template_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found in {template_path}")
    ws = wb[sheet_name]
    template_meta = leap_results_workflow.parse_template_worksheet(ws)
    meta = apply_strict_meta(template_path, sheet_name, template_meta)

    app = leap_results_workflow.connect_leap()
    leap_results_workflow.ensure_calculated(app, force=False)
    leap_results_workflow._show_results_view_table(app)
    leap_results_workflow.set_axes(app, x_axis="Years", legend=str(meta.get("legend_label") or "Fuel"))
    leap_results_workflow.set_context(
        app,
        scenario=meta.get("scenario"),
        region=meta.get("region"),
        branch_path=meta.get("branch"),
    )
    variable_obj = leap_results_workflow._resolve_branch_variable(
        app,
        str(meta.get("branch") or ""),
        meta.get("variable"),
        allow_substitution=False,
    )

    kw = str(keyword or "").strip().lower()

    def _matching_names(names: list[str]) -> list[str]:
        clean = [str(n).strip() for n in names if str(n).strip()]
        by_contains = [n for n in clean if kw in n.lower()]
        fuzzy = difflib.get_close_matches(str(keyword), clean, n=20, cutoff=0.35)
        out: list[str] = []
        seen = set()
        for item in by_contains + fuzzy:
            key = item.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(item)
        return out

    # App-level attribute/method names
    app_names = [name for name in dir(app) if not name.startswith("_")]
    app_hits = _matching_names(app_names)

    # Variable-level attribute/method names
    var_names = [name for name in dir(variable_obj) if not name.startswith("_")]
    var_hits = _matching_names(var_names)

    # Live dimensions in this Results context
    dims = leap_results_workflow.list_dimensions(app)
    dim_hits = _matching_names(dims)

    # Try to inspect dimension members for close matches as well.
    dim_member_hits: dict[str, list[str]] = {}
    try:
        dim_collection = app.Dimensions
        dim_count = int(getattr(dim_collection, "Count", 0))
    except Exception:
        dim_count = 0
    for idx in range(1, dim_count + 1):
        try:
            dim = dim_collection.Item(idx)
            dim_name = str(getattr(dim, "Name", "")).strip()
        except Exception:
            continue
        if not dim_name:
            continue
        members: list[str] = []
        try:
            dim_members = getattr(dim, "Members")
            m_count = int(getattr(dim_members, "Count", 0))
            for m_idx in range(1, m_count + 1):
                name = leap_results_workflow._extract_member_name(dim_members.Item(m_idx))
                if name:
                    members.append(str(name).strip())
        except Exception:
            pass
        hits = _matching_names(members)
        if hits:
            dim_member_hits[dim_name] = hits[:15]

    print("\n" + "=" * 90)
    print("Handle Search Probe")
    print("=" * 90)
    print(f"Template: {template_path.name} / {sheet_name}")
    print(f"Scenario: {meta.get('scenario')}, Region: {meta.get('region')}")
    print(f"Branch: {meta.get('branch')}")
    print(f"Variable: {getattr(variable_obj, 'Name', meta.get('variable'))}")
    print(f"Keyword: {keyword!r}")
    print("\nApp handle matches:")
    print(app_hits[:40] if app_hits else [])
    print("\nVariable handle matches:")
    print(var_hits[:40] if var_hits else [])
    print("\nDimension name matches:")
    print(dim_hits[:40] if dim_hits else [])
    print("\nDimension member matches:")
    if dim_member_hits:
        for dim_name, hits in dim_member_hits.items():
            print(f"  - {dim_name}: {hits}")
    else:
        print("  []")
    print("=" * 90 + "\n")


def run_input_type_legend_csv_probe(
    *,
    template_path: Path | str = Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
    sheet_name: str = "elecgen_inputs",
) -> None:
    """
    Force Results legend to 'Input Type', export CSV, and show whether member rows exist.
    """
    template_path = _resolve(template_path)
    wb = leap_results_workflow.load_workbook(template_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found in {template_path}")
    ws = wb[sheet_name]
    template_meta = leap_results_workflow.parse_template_worksheet(ws)
    meta = apply_strict_meta(template_path, sheet_name, template_meta)

    app = leap_results_workflow.connect_leap()
    leap_results_workflow.ensure_calculated(app, force=False)
    leap_results_workflow._show_results_view_table(app)
    leap_results_workflow.set_context(
        app,
        scenario=meta.get("scenario"),
        region=meta.get("region"),
        branch_path=meta.get("branch"),
    )

    variable_obj = leap_results_workflow._resolve_branch_variable(
        app,
        str(meta.get("branch") or ""),
        meta.get("variable"),
        allow_substitution=False,
    )
    leap_results_workflow._set_active_variable(app, variable_obj)

    # Probe with Input Type legend.
    leap_results_workflow.set_axes(app, x_axis="Years", legend="Input Type")
    leap_results_workflow._show_results_view_table(app)

    export_csv_path = leap_results_workflow._sheet_tmp_export_path(template_path, f"{sheet_name}_input_type_legend", 1)
    leap_results_workflow._export_results_csv_file(app, export_csv_path.resolve())
    df = leap_results_workflow.parse_exported_results_csv(export_csv_path.resolve())

    labels = []
    if not df.empty:
        labels = (
            df.iloc[:, 0]
            .dropna()
            .astype(str)
            .str.strip()
            .tolist()
        )
    labels_non_total = [label for label in labels if label and label.lower() != "total"]

    print("\n" + "=" * 90)
    print("Input-Type Legend CSV Probe")
    print("=" * 90)
    print(f"Template: {template_path.name} / {sheet_name}")
    print(f"Scenario: {meta.get('scenario')}, Region: {meta.get('region')}")
    print(f"Branch: {meta.get('branch')}")
    print(f"Variable: {getattr(variable_obj, 'Name', meta.get('variable'))}")
    print("Applied axes: x=Years, legend=Input Type")
    print(f"CSV path: {export_csv_path}")
    print(f"CSV shape: rows={len(df)}, cols={len(df.columns)}")
    print(f"Legend labels (first 25): {labels[:25]}")
    print(f"Non-total label count: {len(labels_non_total)}")
    print("=" * 90 + "\n")


def _extract_table_sum_for_year(df, *, year: int, exclude_total: bool = True) -> float:
    if df is None or df.empty:
        return 0.0
    cols = [str(c).strip() for c in df.columns]
    year_col = None
    for col in df.columns:
        try:
            if int(float(str(col).strip())) == int(year):
                year_col = col
                break
        except Exception:
            continue
    if year_col is None:
        return 0.0
    labels = df.iloc[:, 0].fillna("").astype(str).str.strip()
    values = leap_results_workflow.pd.to_numeric(df[year_col], errors="coerce").fillna(0.0)
    if exclude_total:
        mask = labels.str.lower() != "total"
        return float(values[mask].sum())
    return float(values.sum())


def _print_excel_like_table(df, *, title: str, max_rows: int = 40) -> None:
    """Print a DataFrame in a sheet-like layout (header + rows)."""
    print(f"\n{title}")
    if df is None or df.empty:
        print("<empty>")
        return
    shown = df.head(int(max_rows)).copy()
    shown.columns = [str(col) for col in shown.columns]
    print(shown.to_string(index=False))
    if len(df) > len(shown):
        print(f"... ({len(df) - len(shown)} more row(s))")


def _resolve_existing_sheet_name(wb, requested: str) -> str:
    """
    Resolve sheet name across old/new naming conventions.

    Supports migration from:
    - <prefix>_inputs
    to:
    - <prefix>_feed_inputs / <prefix>_aux
    """
    requested = str(requested or "").strip()
    if requested in wb.sheetnames:
        return requested
    if requested.endswith("_inputs"):
        prefix = requested[: -len("_inputs")]
        candidates = [f"{prefix}_feed_inputs", f"{prefix}_aux"]
        for name in candidates:
            if name in wb.sheetnames:
                return name
    # fuzzy fallback by prefix
    if "_" in requested:
        prefix = requested.split("_", 1)[0]
        prefixed = [name for name in wb.sheetnames if name.startswith(prefix + "_")]
        if prefixed:
            return prefixed[0]
    raise ValueError(
        f"Sheet '{requested}' not found. Available sheets include: {wb.sheetnames[:12]}"
    )


def run_dual_legend_probe(
    *,
    template_path: Path | str = Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
    sheet_name: str = "elecgen_inputs",
    probe_year: int = 2022,
    favorite_name: str | None = "FuelInputsClean",
) -> None:
    """
    Test dual-legend extraction:
    1) legend=Input Type (feed/aux totals)
    2) legend=Fuel (fuel breakdown)
    and compare year totals.
    """
    template_path = _resolve(template_path)
    wb = leap_results_workflow.load_workbook(template_path)
    sheet_name = _resolve_existing_sheet_name(wb, sheet_name)
    ws = wb[sheet_name]
    template_meta = leap_results_workflow.parse_template_worksheet(ws)
    meta = apply_strict_meta(template_path, sheet_name, template_meta)

    app = leap_results_workflow.connect_leap()
    leap_results_workflow.ensure_calculated(app, force=False)
    favorite_status = None
    if favorite_name:
        try:
            favorite_status = leap_results_workflow.activate_favorite(app, favorite_name)
        except Exception as exc:  # noqa: BLE001
            favorite_status = f"Favorite activate failed: {exc}"
    leap_results_workflow._show_results_view_table(app)
    leap_results_workflow.set_context(
        app,
        scenario=meta.get("scenario"),
        region=meta.get("region"),
        branch_path=meta.get("branch"),
    )
    variable_obj = leap_results_workflow._resolve_branch_variable(
        app,
        str(meta.get("branch") or ""),
        meta.get("variable"),
        allow_substitution=False,
    )
    leap_results_workflow._set_active_variable(app, variable_obj)

    def _active_legend_text() -> str:
        try:
            return str(getattr(app, "ResultsLegend", "") or "").strip()
        except Exception:
            return ""

    # Pass 1: Input Type legend
    leap_results_workflow.set_axes(app, x_axis="Years", legend="Input Type")
    leap_results_workflow.set_context(
        app,
        scenario=meta.get("scenario"),
        region=meta.get("region"),
        branch_path=meta.get("branch"),
    )
    leap_results_workflow._show_results_view_table(app)
    csv_input_type = leap_results_workflow._sheet_tmp_export_path(template_path, f"{sheet_name}_dual_input_type", 1)
    leap_results_workflow._export_results_csv_file(app, csv_input_type.resolve())
    df_input_type = leap_results_workflow.parse_exported_results_csv(csv_input_type.resolve())
    legend_after_input_type = _active_legend_text()

    # Pass 2: safe mode - only switch once to Fuel (no legend candidate loop).
    # Repeated/invalid legend switching can crash LEAP's report option writer.
    fuel_legend_used = "Fuel"
    csv_fuel = leap_results_workflow._sheet_tmp_export_path(template_path, f"{sheet_name}_dual_fuel", 1)
    leap_results_workflow.set_axes(app, x_axis="Years", legend=fuel_legend_used)
    leap_results_workflow.set_context(
        app,
        scenario=meta.get("scenario"),
        region=meta.get("region"),
        branch_path=meta.get("branch"),
    )
    # Let LEAP settle the view state before exporting.
    time.sleep(0.25)
    leap_results_workflow._show_results_view_table(app)
    time.sleep(0.25)
    leap_results_workflow._export_results_csv_file(app, csv_fuel.resolve())
    df_fuel = leap_results_workflow.parse_exported_results_csv(csv_fuel.resolve())
    fuel_labels = (
        df_fuel.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        if not df_fuel.empty
        else []
    )
    fuel_legend_after_set = _active_legend_text()

    labels_input_type = (
        df_input_type.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        if not df_input_type.empty
        else []
    )
    labels_fuel = fuel_labels

    input_type_sum = _extract_table_sum_for_year(df_input_type, year=probe_year, exclude_total=True)
    fuel_sum = _extract_table_sum_for_year(df_fuel, year=probe_year, exclude_total=True)
    total_gap = fuel_sum - input_type_sum

    # Extra check: "All fuels" style filters in ValueRS.
    fuel_filter_trials = [
        "",
        "Fuel=All Fuels",
        "Fuel=All fuels",
        "Fuel=Total",
        "Input Type=Feedstock Fuels",
        "Input Type=Auxiliary Fuels From Outputs",
        "Input Type=Auxiliary Fuels From Other Modules",
        "Input Type=Feedstock Fuels;Fuel=Natural gas",
        "Input Type=Auxiliary Fuels From Outputs;Fuel=Natural gas",
    ]
    fuel_filter_results = []
    for trial in fuel_filter_trials:
        ok, value, err = _try_value_rs(
            variable_obj,
            region=str(meta.get("region") or ""),
            scenario=str(meta.get("scenario") or ""),
            year=int(probe_year),
            filter_str=trial,
        )
        fuel_filter_results.append((trial, ok, value, err))

    print("\n" + "=" * 90)
    print("Dual Legend Probe (Input Type -> Fuel)")
    print("=" * 90)
    print(f"Template: {template_path.name} / {sheet_name}")
    print(f"Scenario: {meta.get('scenario')}, Region: {meta.get('region')}, Year: {probe_year}")
    print(f"Branch: {meta.get('branch')}")
    print(f"Variable: {getattr(variable_obj, 'Name', meta.get('variable'))}")
    print(f"Favorite activation: {favorite_status}")
    print(f"Legend after Input-Type set: {legend_after_input_type}")
    print(f"Fuel legend requested/used: {fuel_legend_used}")
    print(f"Legend after fuel set: {fuel_legend_after_set}")
    print(f"Input-Type CSV: {csv_input_type} | shape={df_input_type.shape}")
    print(f"Fuel CSV: {csv_fuel} | shape={df_fuel.shape}")
    print(f"Input-Type labels: {labels_input_type[:20]}")
    print(f"Fuel labels: {labels_fuel[:20]}")
    print(f"Year {probe_year} sum(Input-Type components) = {input_type_sum}")
    print(f"Year {probe_year} sum(Fuel components) = {fuel_sum}")
    print(f"Gap (Fuel - InputType) = {total_gap}")
    print("\nValueRS fuel filter trials:")
    for trial, ok, value, err in fuel_filter_results:
        status = "OK" if ok else "ERR"
        print(f"  [{status}] filter={trial!r} -> value={value} error={err}")
    _print_excel_like_table(df_input_type, title="Input-Type Table (Excel-like)")
    _print_excel_like_table(df_fuel, title="Fuel Table (Excel-like)")
    print("=" * 90 + "\n")


def run_input_type_filter_then_fuel_probe(
    *,
    template_path: Path | str = Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
    sheet_name: str = "elecgen_inputs",
    probe_year: int = 2022,
    favorite_name: str | None = "FuelInputsClean",
) -> None:
    """
    Experiment requested:
    - set legend to Fuel
    - attempt Input Type filters (Feedstock/Auxiliary)
    - check if LEAP exposes fuel rows under those filters
    """
    template_path = _resolve(template_path)
    wb = leap_results_workflow.load_workbook(template_path)
    sheet_name = _resolve_existing_sheet_name(wb, sheet_name)
    ws = wb[sheet_name]
    meta = apply_strict_meta(template_path, sheet_name, leap_results_workflow.parse_template_worksheet(ws))

    app = leap_results_workflow.connect_leap()
    leap_results_workflow.ensure_calculated(app, force=False)
    favorite_status = None
    if favorite_name:
        try:
            favorite_status = leap_results_workflow.activate_favorite(app, favorite_name)
        except Exception as exc:  # noqa: BLE001
            favorite_status = f"Favorite activate failed: {exc}"
    leap_results_workflow._show_results_view_table(app)
    leap_results_workflow.set_context(
        app,
        scenario=meta.get("scenario"),
        region=meta.get("region"),
        branch_path=meta.get("branch"),
    )
    variable_obj = leap_results_workflow._resolve_branch_variable(
        app,
        str(meta.get("branch") or ""),
        meta.get("variable"),
        allow_substitution=False,
    )
    leap_results_workflow._set_active_variable(app, variable_obj)

    leap_results_workflow.set_axes(app, x_axis="Years", legend="Fuel")
    time.sleep(0.25)
    leap_results_workflow._show_results_view_table(app)
    time.sleep(0.25)

    try:
        active_legend = str(getattr(app, "ResultsLegend", "") or "").strip()
    except Exception:
        active_legend = ""

    filters = [
        "",
        "Fuel=All Fuels",
        "Fuel=Total",
        "Input Type=Feedstock Fuels",
        "Input Type=Auxiliary Fuels From Outputs",
        "Input Type=Auxiliary Fuels From Other Modules",
        "Input Type=Feedstock Fuels;Fuel=Natural gas",
        "Input Type=Auxiliary Fuels From Outputs;Fuel=Natural gas",
    ]
    trial_rows: list[dict[str, object]] = []
    for f in filters:
        ok, value, err = _try_value_rs(
            variable_obj,
            region=str(meta.get("region") or ""),
            scenario=str(meta.get("scenario") or ""),
            year=int(probe_year),
            filter_str=f,
        )
        trial_rows.append({"filter": f, "ok": ok, "value": value, "error": err})

    csv_path = leap_results_workflow._sheet_tmp_export_path(template_path, f"{sheet_name}_fuel_legend_probe", 1)
    leap_results_workflow._export_results_csv_file(app, csv_path.resolve())
    df = leap_results_workflow.parse_exported_results_csv(csv_path.resolve())
    labels = df.iloc[:, 0].dropna().astype(str).str.strip().tolist() if not df.empty else []

    print("\n" + "=" * 90)
    print("InputType-Filter + Fuel-Legend Probe")
    print("=" * 90)
    print(f"Template: {template_path.name} / {sheet_name}")
    print(f"Scenario: {meta.get('scenario')}, Region: {meta.get('region')}, Year: {probe_year}")
    print(f"Branch: {meta.get('branch')}")
    print(f"Variable: {getattr(variable_obj, 'Name', meta.get('variable'))}")
    print(f"Favorite activation: {favorite_status}")
    print(f"Active ResultsLegend: {active_legend!r}")
    print(f"Fuel-legend CSV path: {csv_path}")
    print(f"CSV shape: {df.shape}")
    print(f"CSV first labels: {labels[:20]}")
    print("\nValueRS filter trials:")
    for row in trial_rows:
        status = "OK" if row["ok"] else "ERR"
        print(f"  [{status}] filter={row['filter']!r} -> value={row['value']} error={row['error']}")
    _print_excel_like_table(df, title="Fuel-Legend Export (Excel-like)")
    print("=" * 90 + "\n")


def _safe_token(text: str) -> str:
    token = re.sub(r"[^A-Za-z0-9._-]+", "_", str(text or "").strip())
    return token.strip("_") or "favorite"


def run_favorite_tables_probe(
    *,
    favorite_names: list[str],
    scenario: str = "Reference",
    region: str = "United States",
    probe_year: int = 2022,
    preview_rows: int = 12,
) -> None:
    """
    Activate each LEAP favorite, export Results CSV, and print table diagnostics.

    This is designed for your new favorite-per-sheet workflow where favorite names
    map directly to expected table outputs (e.g., elecgen_feed_inputs).
    """
    names = [str(name).strip() for name in favorite_names if str(name).strip()]
    if not names:
        raise ValueError("favorite_names is required")

    app = leap_results_workflow.connect_leap()
    leap_results_workflow.ensure_calculated(app, force=False)
    leap_results_workflow._show_results_view_table(app)
    leap_results_workflow.set_context(
        app,
        scenario=scenario,
        region=region,
    )

    print("\n" + "=" * 90)
    print("Favorite Tables Probe")
    print("=" * 90)
    print(f"Scenario: {scenario}, Region: {region}, Year: {probe_year}")
    print(f"Favorites count: {len(names)}")

    summary_rows: list[dict[str, object]] = []
    for idx, favorite in enumerate(names, start=1):
        status = ""
        csv_path = leap_results_workflow._sheet_tmp_export_path(
            Path("favorite_probe.xlsx"),
            f"{_safe_token(favorite)}",
            idx,
        )
        table_df = leap_results_workflow.pd.DataFrame()
        first_labels: list[str] = []
        total_value = None
        components_sum = None
        note = ""
        try:
            activation = leap_results_workflow.activate_favorite(app, favorite)
            status = str(activation or "")
            leap_results_workflow._show_results_view_table(app)
            leap_results_workflow._export_results_csv_file(app, csv_path.resolve())
            table_df = leap_results_workflow.parse_exported_results_csv(csv_path.resolve())
            if not table_df.empty:
                first_col = table_df.iloc[:, 0].dropna().astype(str).str.strip()
                first_labels = first_col.tolist()[:15]
                labels = first_col.str.lower()
                values = leap_results_workflow.pd.to_numeric(
                    table_df.get(str(probe_year), leap_results_workflow.pd.Series(dtype=float)),
                    errors="coerce",
                ).fillna(0.0)
                if len(values) == len(labels):
                    total_mask = labels.eq("total")
                    if total_mask.any():
                        total_value = float(values[total_mask].iloc[0])
                    components_sum = float(values[~total_mask].sum())
        except Exception as exc:  # noqa: BLE001
            note = str(exc)
            status = "ERROR"

        summary_rows.append(
            {
                "favorite": favorite,
                "status": status,
                "rows": int(len(table_df)),
                "cols": int(len(table_df.columns)) if not table_df.empty else 0,
                "year_total": total_value,
                "year_components_sum": components_sum,
                "labels_preview": " | ".join(first_labels),
                "csv_path": str(csv_path),
                "note": note,
            }
        )

        print("\n" + "-" * 90)
        print(f"Favorite: {favorite}")
        print(f"Status: {status}")
        if note:
            print(f"Note: {note}")
        if not table_df.empty:
            _print_excel_like_table(table_df, title="Exported Table (Excel-like)", max_rows=preview_rows)
        else:
            print("Exported Table (Excel-like)\n<empty>")

    summary_df = leap_results_workflow.pd.DataFrame(summary_rows)
    print("\n" + "=" * 90)
    print("Favorite Probe Summary")
    print("=" * 90)
    if not summary_df.empty:
        print(summary_df.to_string(index=False))
    else:
        print("<no results>")
    print("=" * 90 + "\n")


def _coerce_favorite_table_numeric(df):
    """Normalize a LEAP exported table to Fuel + numeric year columns."""
    if df is None or df.empty:
        return leap_results_workflow.pd.DataFrame(columns=["Fuel"])
    out = df.copy()
    first_col = str(out.columns[0])
    out = out.rename(columns={first_col: "Fuel"})
    out["Fuel"] = out["Fuel"].fillna("").astype(str).str.strip()
    out = out[out["Fuel"] != ""]
    for col in out.columns[1:]:
        out[col] = leap_results_workflow.pd.to_numeric(out[col], errors="coerce").fillna(0.0)
    return out


def _sum_favorite_tables_by_fuel(dfs: list) -> leap_results_workflow.pd.DataFrame:
    if not dfs:
        return leap_results_workflow.pd.DataFrame(columns=["Fuel"])
    combined = dfs[0].copy()
    year_cols = [c for c in combined.columns if c != "Fuel"]
    for nxt in dfs[1:]:
        all_cols = ["Fuel"] + sorted(set(year_cols + [c for c in nxt.columns if c != "Fuel"]), key=lambda x: str(x))
        combined = combined.reindex(columns=all_cols, fill_value=0.0)
        nxt2 = nxt.reindex(columns=all_cols, fill_value=0.0)
        combined = combined.merge(nxt2, on="Fuel", how="outer", suffixes=("_l", "_r")).fillna(0.0)
        merged_cols = [c for c in combined.columns if c != "Fuel"]
        base_names = sorted({c[:-2] for c in merged_cols if c.endswith("_l")} | {c[:-2] for c in merged_cols if c.endswith("_r")})
        rebuilt = leap_results_workflow.pd.DataFrame({"Fuel": combined["Fuel"]})
        for name in base_names:
            l = f"{name}_l"
            r = f"{name}_r"
            rebuilt[name] = combined.get(l, 0.0) + combined.get(r, 0.0)
        combined = rebuilt
        year_cols = [c for c in combined.columns if c != "Fuel"]
    # Collapse duplicate fuel labels if any.
    numeric_cols = [c for c in combined.columns if c != "Fuel"]
    if numeric_cols:
        combined = combined.groupby("Fuel", as_index=False)[numeric_cols].sum()
    return combined


def run_input_style_sum_probe(
    *,
    sector_prefixes: list[str],
    scenario: str = "Reference",
    region: str = "United States",
    probe_year: int = 2022,
    compare_to_suffix: str | None = "_inputs",
    preview_rows: int = 18,
) -> None:
    """
    For each sector prefix, export and sum:
      <prefix>_feed_inputs + <prefix>_aux + <prefix>_aux_other
    Optionally compare the summed table against <prefix><compare_to_suffix>.
    """
    prefixes = [str(p).strip() for p in sector_prefixes if str(p).strip()]
    if not prefixes:
        raise ValueError("sector_prefixes is required")

    app = leap_results_workflow.connect_leap()
    leap_results_workflow.ensure_calculated(app, force=False)
    leap_results_workflow._show_results_view_table(app)
    leap_results_workflow.set_context(app, scenario=scenario, region=region)

    print("\n" + "=" * 90)
    print("Input-Style Sum Probe")
    print("=" * 90)
    print(f"Scenario: {scenario}, Region: {region}, Year: {probe_year}")

    for idx, prefix in enumerate(prefixes, start=1):
        fav_feed = f"{prefix}_feed_inputs"
        fav_aux = f"{prefix}_aux"
        fav_aux_other = f"{prefix}_aux_other"
        components = [fav_feed, fav_aux, fav_aux_other]
        component_tables = []
        component_status = []

        for j, fav in enumerate(components, start=1):
            csv_path = leap_results_workflow._sheet_tmp_export_path(
                Path("input_style_sum_probe.xlsx"),
                f"{_safe_token(prefix)}_{_safe_token(fav)}",
                idx * 10 + j,
            )
            try:
                status = leap_results_workflow.activate_favorite(app, fav)
                leap_results_workflow._show_results_view_table(app)
                leap_results_workflow._export_results_csv_file(app, csv_path.resolve())
                raw_df = leap_results_workflow.parse_exported_results_csv(csv_path.resolve())
                norm_df = _coerce_favorite_table_numeric(raw_df)
                component_tables.append(norm_df)
                component_status.append((fav, str(status or ""), str(csv_path), ""))
            except Exception as exc:  # noqa: BLE001
                component_status.append((fav, "ERROR", str(csv_path), str(exc)))
                component_tables.append(leap_results_workflow.pd.DataFrame(columns=["Fuel"]))

        combined = _sum_favorite_tables_by_fuel(component_tables)
        if not combined.empty:
            numeric_cols = [c for c in combined.columns if c != "Fuel"]
            if numeric_cols:
                total_row = leap_results_workflow.pd.DataFrame(
                    [{"Fuel": "Total", **{c: float(combined[c].sum()) for c in numeric_cols}}]
                )
                combined_with_total = leap_results_workflow.pd.concat([combined, total_row], ignore_index=True)
            else:
                combined_with_total = combined
        else:
            combined_with_total = combined

        # Optional comparison against legacy/target favorite (e.g., <prefix>_inputs).
        compare_note = "not_compared"
        if compare_to_suffix:
            compare_fav = f"{prefix}{compare_to_suffix}"
            cmp_csv = leap_results_workflow._sheet_tmp_export_path(
                Path("input_style_sum_probe.xlsx"),
                f"{_safe_token(prefix)}_{_safe_token(compare_fav)}",
                idx * 10 + 9,
            )
            try:
                leap_results_workflow.activate_favorite(app, compare_fav)
                leap_results_workflow._show_results_view_table(app)
                leap_results_workflow._export_results_csv_file(app, cmp_csv.resolve())
                cmp_raw = leap_results_workflow.parse_exported_results_csv(cmp_csv.resolve())
                cmp_df = _coerce_favorite_table_numeric(cmp_raw)

                all_fuels = sorted(set(combined_with_total.get("Fuel", leap_results_workflow.pd.Series(dtype=str))) | set(cmp_df.get("Fuel", leap_results_workflow.pd.Series(dtype=str))))
                all_cols = sorted(set([c for c in combined_with_total.columns if c != "Fuel"] + [c for c in cmp_df.columns if c != "Fuel"]), key=lambda x: str(x))
                lhs = combined_with_total.set_index("Fuel").reindex(index=all_fuels, columns=all_cols, fill_value=0.0)
                rhs = cmp_df.set_index("Fuel").reindex(index=all_fuels, columns=all_cols, fill_value=0.0)
                diff = (lhs - rhs).abs()
                max_abs = float(diff.to_numpy().max()) if not diff.empty else 0.0
                year_col = str(probe_year)
                year_diff = float(diff[year_col].sum()) if year_col in diff.columns else 0.0
                compare_note = f"compare_favorite={compare_fav}, max_abs_diff={max_abs:.6f}, sum_abs_diff_{probe_year}={year_diff:.6f}"
            except Exception as exc:  # noqa: BLE001
                compare_note = f"compare_failed({compare_fav}): {exc}"

        print("\n" + "-" * 90)
        print(f"Sector prefix: {prefix}")
        for fav, status, csv_path, note in component_status:
            msg = f"  - {fav}: {status} | {csv_path}"
            if note:
                msg += f" | note={note}"
            print(msg)
        print(f"Comparison: {compare_note}")
        _print_excel_like_table(combined_with_total, title=f"Combined Input Table ({prefix})", max_rows=preview_rows)

    print("=" * 90 + "\n")


def build_inputs_by_fuel_from_favorites(
    *,
    sector_prefixes: list[str],
    scenario: str = "Reference",
    region: str = "United States",
    output_workbook: Path | str = Path("outputs/leap_results_replica/inputs_by_fuel_from_favorites.xlsx"),
) -> Path:
    """
    Build '<prefix>_inputs_by_fuel' tables by summing:
      <prefix>_feed_inputs + <prefix>_aux + <prefix>_aux_other
    and save them in one workbook (one sheet per prefix).
    """
    prefixes = [str(p).strip() for p in sector_prefixes if str(p).strip()]
    if not prefixes:
        raise ValueError("sector_prefixes is required")

    out_path = _resolve(output_workbook)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    app = leap_results_workflow.connect_leap()
    leap_results_workflow.ensure_calculated(app, force=False)
    leap_results_workflow._show_results_view_table(app)
    leap_results_workflow.set_context(app, scenario=scenario, region=region)

    built_tables: list[tuple[str, object]] = []
    with leap_results_workflow.pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for idx, prefix in enumerate(prefixes, start=1):
            favs = [
                f"{prefix}_feed_inputs",
                f"{prefix}_aux",
                f"{prefix}_aux_other",
            ]
            component_tables = []
            for j, fav in enumerate(favs, start=1):
                csv_path = leap_results_workflow._sheet_tmp_export_path(
                    Path("inputs_by_fuel_builder.xlsx"),
                    f"{_safe_token(prefix)}_{_safe_token(fav)}",
                    idx * 10 + j,
                )
                status = leap_results_workflow.activate_favorite(app, fav)
                leap_results_workflow._show_results_view_table(app)
                leap_results_workflow._export_results_csv_file(app, csv_path.resolve())
                raw_df = leap_results_workflow.parse_exported_results_csv(csv_path.resolve())
                component_tables.append(_coerce_favorite_table_numeric(raw_df))

            combined = _sum_favorite_tables_by_fuel(component_tables)
            if not combined.empty:
                numeric_cols = [c for c in combined.columns if c != "Fuel"]
                if numeric_cols:
                    total_row = leap_results_workflow.pd.DataFrame(
                        [{"Fuel": "Total", **{c: float(combined[c].sum()) for c in numeric_cols}}]
                    )
                    combined = leap_results_workflow.pd.concat([combined, total_row], ignore_index=True)

            sheet_name = f"{prefix}_inputs_by_fuel"
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]
            combined.to_excel(writer, sheet_name=sheet_name, index=False)
            built_tables.append((prefix, combined))

    print("\n" + "=" * 90)
    print("Inputs By Fuel Build Complete")
    print("=" * 90)
    print(f"Scenario: {scenario}")
    print(f"Region: {region}")
    print(f"Sectors: {len(prefixes)}")
    print(f"Workbook: {out_path}")
    print("-" * 90)
    for prefix, table in built_tables:
        year_total = None
        if not table.empty and "Fuel" in table.columns:
            total_mask = table["Fuel"].astype(str).str.strip().str.lower().eq("total")
            if total_mask.any() and "2022" in table.columns:
                year_total = float(leap_results_workflow.pd.to_numeric(table.loc[total_mask, "2022"], errors="coerce").fillna(0.0).iloc[0])
        print(f"\nSector: {prefix}")
        print(f"2022 total: {year_total}")
        _print_excel_like_table(table, title=f"{prefix}_inputs_by_fuel (Excel-like)", max_rows=20)
    print("=" * 90 + "\n")
    return out_path


def _list_favorite_names(app) -> list[str]:
    return leap_results_workflow._list_favorite_names(app)


def run_transformation_full_insert_test(
    *,
    source_template: Path | str = Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
    output_workbook: Path | str = Path("outputs/leap_results_replica/transformation_results_20_USA_Reference.full_insert_test.xlsx"),
    clear_on_effective_empty: bool = True,
) -> dict[str, object]:
    """
    End-to-end test:
    - copy source transformation workbook
    - verify/try all required input-style favorites
    - run full live extraction for all sheets into the copy
      (favorites for input-style sheets, normal extraction for others)
    """
    source = _resolve(source_template)
    output = _resolve(output_workbook)
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, output)
    leap_results_workflow.reset_rendered_csv_metadata_mismatch_events()

    wb = leap_results_workflow.load_workbook(output)
    naming_issues = leap_results_workflow.validate_transformation_sheet_naming(
        wb.sheetnames,
        context=output.name,
    )
    if naming_issues:
        raise RuntimeError("Sheet naming validation failed:\n - " + "\n - ".join(naming_issues))
    required_favorites = [
        s for s in wb.sheetnames
        if str(s).strip().lower().endswith(("_feed_inputs", "_aux", "_aux_other"))
    ]

    app = leap_results_workflow.connect_leap()
    leap_results_workflow.ensure_calculated(app, force=False)
    leap_results_workflow._show_results_view_table(app)
    available_favorites = _list_favorite_names(app)

    favorite_probe_rows: list[dict[str, object]] = []
    for fav in required_favorites:
        resolved_name, resolve_mode, status_text = leap_results_workflow.activate_favorite_for_sheet(app, fav)
        ok = bool(resolved_name) and "not found/activated" not in str(status_text).lower() and "error" not in str(status_text).lower()
        favorite_probe_rows.append(
            {
                "favorite": fav,
                "resolved_favorite": resolved_name,
                "resolve_mode": resolve_mode,
                "status": status_text,
                "ok": bool(ok),
                "exists_in_favorites_list": fav in available_favorites,
            }
        )

    temp_output_dir = output.parent / "_full_insert_temp"
    temp_output_dir.mkdir(parents=True, exist_ok=True)
    summary = run_replica_extraction(
        template_paths=[output],
        output_dir=temp_output_dir,
        mode="live",
        golden_workbook=None,
        clear_on_effective_empty=bool(clear_on_effective_empty),
    )
    mismatch_events = leap_results_workflow.get_rendered_csv_metadata_mismatch_events()

    replica_output = temp_output_dir / f"{output.stem}.replica.xlsx"
    if replica_output.exists():
        shutil.copyfile(replica_output, output)

    fav_df = leap_results_workflow.pd.DataFrame(favorite_probe_rows)
    fav_report = output.parent / f"{output.stem}.favorite_probe.csv"
    fav_df.to_csv(fav_report, index=False)
    mismatch_report = output.parent / f"{output.stem}.csv_metadata_mismatch_events.csv"
    leap_results_workflow.pd.DataFrame(mismatch_events).to_csv(mismatch_report, index=False)

    print("\n" + "=" * 90)
    print("Transformation Full Insert Test")
    print("=" * 90)
    print(f"Source template: {source}")
    print(f"Output workbook: {output}")
    print(f"Replica intermediate dir: {temp_output_dir}")
    print(f"Favorite probe report: {fav_report}")
    print(f"CSV metadata mismatch report: {mismatch_report}")
    print(f"Required input-style favorites: {len(required_favorites)}")
    print(f"Favorite checks passed: {int(fav_df.get('ok', leap_results_workflow.pd.Series(dtype=bool)).sum()) if not fav_df.empty else 0}")
    print(f"CSV metadata mismatch events: {len(mismatch_events)}")
    print("Replica summary:")
    for key, value in summary.items():
        print(f"  {key}: {value}")
    print("=" * 90 + "\n")

    return {
        "source_template": str(source),
        "output_workbook": str(output),
        "favorite_probe_report": str(fav_report),
        "csv_metadata_mismatch_report": str(mismatch_report),
        "csv_metadata_mismatch_events": mismatch_events,
        "required_favorites": required_favorites,
        "replica_summary": summary,
    }


def run_transformation_sector_probe(
    *,
    source_template: Path | str = Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
    sector_prefix: str = "transfers_unalloc",
    output_workbook: Path | str = Path("outputs/leap_results_replica/transfers_unalloc_sector_probe.xlsx"),
) -> dict[str, object]:
    """
    Probe one transformation sector only (all sheets starting with sector_prefix).
    Writes a workbook copy with only those sheets refilled and a summary CSV.
    """
    source = _resolve(source_template)
    output = _resolve(output_workbook)
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source, output)
    leap_results_workflow.reset_rendered_csv_metadata_mismatch_events()

    wb = leap_results_workflow.load_workbook(output)
    naming_issues = leap_results_workflow.validate_transformation_sheet_naming(
        wb.sheetnames,
        context=output.name,
    )
    if naming_issues:
        raise RuntimeError("Sheet naming validation failed:\n - " + "\n - ".join(naming_issues))
    target_sheets = [s for s in wb.sheetnames if str(s).startswith(f"{sector_prefix}_")]
    if not target_sheets:
        raise ValueError(f"No sheets found for sector_prefix='{sector_prefix}' in {source}")

    app = leap_results_workflow.connect_leap()
    leap_results_workflow.ensure_calculated(app, force=False)
    rows: list[dict[str, object]] = []

    for idx, sheet_name in enumerate(target_sheets, start=1):
        ws = wb[sheet_name]
        meta = apply_strict_meta(output, sheet_name, leap_results_workflow.parse_template_worksheet(ws))
        export_csv_path = leap_results_workflow._sheet_tmp_export_path(output, sheet_name, idx)
        status = "ok"
        note = ""
        try:
            input_style = str(sheet_name).lower().endswith(("_feed_inputs", "_aux", "_aux_other"))
            favorite_activated = False
            if input_style:
                resolved_name, resolve_mode, fav_status = leap_results_workflow.activate_favorite_for_sheet(app, sheet_name)
                favorite_activated = bool(resolved_name) and "not found/activated" not in str(fav_status).lower()
                note = f"favorite={fav_status}; resolve_mode={resolve_mode}; favorite_name={resolved_name}"
            fresh_df = leap_results_workflow.build_fresh_table_from_export(
                app,
                meta,
                export_csv_path=export_csv_path,
                context=f"sector_probe/{output.name}/{sheet_name}",
                allow_variable_substitution=False,
                trust_active_results_view=(input_style and favorite_activated),
                expected_input_type_qualifier=leap_results_workflow._expected_input_type_qualifier_for_sheet(sheet_name),
            )
            if is_effectively_empty_results_table(fresh_df):
                fresh_df = leap_results_workflow._build_no_results_table(meta)
                note = f"{note}; effective_empty_cleared".strip("; ")
            leap_results_workflow.write_table_values_preserve_format(ws, fresh_df)

            labels = (
                fresh_df.iloc[6:, 0].fillna("").astype(str).str.strip()
                if not fresh_df.empty and fresh_df.shape[0] >= 7
                else leap_results_workflow.pd.Series(dtype=str)
            )
            values = (
                fresh_df.iloc[6:, 1:].apply(leap_results_workflow.pd.to_numeric, errors="coerce").fillna(0.0)
                if not fresh_df.empty and fresh_df.shape[0] >= 7 and fresh_df.shape[1] > 1
                else leap_results_workflow.pd.DataFrame()
            )
            non_total_abs_sum = float(values[labels.str.lower() != "total"].abs().sum().sum()) if not values.empty else 0.0
            row1_variable = str(fresh_df.iloc[0, 0]) if not fresh_df.empty else ""
            label_preview = " | ".join(labels[labels != ""].head(10).tolist()) if not labels.empty else ""
        except Exception as exc:  # noqa: BLE001
            status = "error"
            note = str(exc)
            row1_variable = ""
            non_total_abs_sum = 0.0
            label_preview = ""

        rows.append(
            {
                "sheet": sheet_name,
                "status": status,
                "row1_variable": row1_variable,
                "non_total_abs_sum": non_total_abs_sum,
                "labels_preview": label_preview,
                "note": note,
            }
        )

    wb.save(output)
    summary_df = leap_results_workflow.pd.DataFrame(rows)
    summary_csv = output.parent / "supporting_files" / f"{output.stem}.summary.csv"
    summary_csv.parent.mkdir(parents=True, exist_ok=True)
    summary_df.to_csv(summary_csv, index=False)
    mismatch_events = leap_results_workflow.get_rendered_csv_metadata_mismatch_events()
    mismatch_csv = output.parent / "supporting_files" / f"{output.stem}.csv_metadata_mismatch_events.csv"
    leap_results_workflow.pd.DataFrame(mismatch_events).to_csv(mismatch_csv, index=False)

    print("\n" + "=" * 90)
    print("Transformation Sector Probe")
    print("=" * 90)
    print(f"Source: {source}")
    print(f"Sector prefix: {sector_prefix}")
    print(f"Output workbook: {output}")
    print(f"Summary CSV: {summary_csv}")
    print(f"Mismatch CSV: {mismatch_csv}")
    print(summary_df.to_string(index=False))
    print("=" * 90 + "\n")

    return {
        "output_workbook": str(output),
        "summary_csv": str(summary_csv),
        "mismatch_csv": str(mismatch_csv),
        "sheets": target_sheets,
    }


if __name__ == "__main__":
    summary = run_replica_extraction(
        template_paths=[_resolve(path) for path in TEMPLATE_PATHS],
        output_dir=_resolve(OUTPUT_DIR),
        mode=MODE,
        golden_workbook=_resolve(GOLDEN_WORKBOOK_PATH) if GOLDEN_WORKBOOK_PATH else None,
        clear_on_effective_empty=True,
    )
    print("Replica extraction complete:")
    for key, value in summary.items():
        print(f"  {key}: {value}")

    # Single-sheet live probe to test whether Input Type must be set explicitly.
    # Toggle to True when debugging LEAP live extraction behavior.
    RUN_SINGLE_INPUT_TYPE_PROBE = False
    if RUN_SINGLE_INPUT_TYPE_PROBE:
        run_single_input_type_probe(
            template_path=Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
            sheet_name="elecgen_inputs",
            probe_year=2022,
            sample_fuel="Natural gas",
        )

    # Single-sheet empty-result test (inject stale values, then verify they clear).
    RUN_SINGLE_EMPTY_CLEAR_PROBE = False
    if RUN_SINGLE_EMPTY_CLEAR_PROBE:
        run_single_empty_clear_probe(
            template_path=Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
            sheet_name="lng_regas_out_feed",
            probe_year=2022,
        )

    # Handle search around a keyword (e.g., "Input Type") in current LEAP context.
    RUN_HANDLE_SEARCH_PROBE = False
    if RUN_HANDLE_SEARCH_PROBE:
        run_handle_search_probe(
            template_path=Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
            sheet_name="elecgen_inputs",
            keyword="Input Type",
        )

    # Force legend to Input Type and inspect exported CSV rows.
    RUN_INPUT_TYPE_LEGEND_CSV_PROBE = False
    if RUN_INPUT_TYPE_LEGEND_CSV_PROBE:
        run_input_type_legend_csv_probe(
            template_path=Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
            sheet_name="elecgen_inputs",
        )

    # Two-pass legend test: Input Type pass then Fuel pass.
    RUN_DUAL_LEGEND_PROBE = False
    if RUN_DUAL_LEGEND_PROBE:
        run_dual_legend_probe(
            template_path=Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
            sheet_name="elecgen_inputs",
            probe_year=2022,
            favorite_name="FuelInputsClean",
        )

    RUN_INPUT_TYPE_FILTER_THEN_FUEL_PROBE = False
    if RUN_INPUT_TYPE_FILTER_THEN_FUEL_PROBE:
        run_input_type_filter_then_fuel_probe(
            template_path=Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
            sheet_name="elecgen_inputs",
            probe_year=2022,
            favorite_name="FuelInputsClean",
        )

    RUN_FAVORITE_TABLES_PROBE = False
    FAVORITE_TABLES_TO_TEST = [
        "elecgen_feed_inputs",
        "elecgen_aux",
        "elecgen_aux_other",
    ]
    if RUN_FAVORITE_TABLES_PROBE:
        run_favorite_tables_probe(
            favorite_names=FAVORITE_TABLES_TO_TEST,
            scenario="Reference",
            region="United States",
            probe_year=2022,
            preview_rows=12,
        )

    # Sum test:
    #   <prefix>_feed_inputs + <prefix>_aux + <prefix>_aux_other
    # and compare to optional <prefix>_inputs favorite if present.
    RUN_INPUT_STYLE_SUM_PROBE = False
    INPUT_STYLE_SUM_PREFIXES = [
        "elecgen",
        "refining",
    ]
    if RUN_INPUT_STYLE_SUM_PROBE:
        run_input_style_sum_probe(
            sector_prefixes=INPUT_STYLE_SUM_PREFIXES,
            scenario="Reference",
            region="United States",
            probe_year=2022,
            compare_to_suffix="_inputs",
            preview_rows=18,
        )

    # Final-builder utility:
    # create <prefix>_inputs_by_fuel = feed_inputs + aux_outputs + aux_other_outputs
    RUN_BUILD_INPUTS_BY_FUEL_FROM_FAVORITES = False
    BUILD_INPUTS_BY_FUEL_PREFIXES = [
        "elecgen",
        "refining",
    ]
    if RUN_BUILD_INPUTS_BY_FUEL_FROM_FAVORITES:
        build_inputs_by_fuel_from_favorites(
            sector_prefixes=BUILD_INPUTS_BY_FUEL_PREFIXES,
            scenario="Reference",
            region="United States",
            output_workbook=Path("outputs/leap_results_replica/inputs_by_fuel_from_favorites.xlsx"),
        )

    # Full workbook insertion test on a copy of the transformation reference template.
    RUN_TRANSFORMATION_FULL_INSERT_TEST = False
    if RUN_TRANSFORMATION_FULL_INSERT_TEST:
        run_transformation_full_insert_test(
            source_template=Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
            output_workbook=Path("outputs/leap_results_replica/transformation_results_20_USA_Reference.full_insert_test.xlsx"),
            clear_on_effective_empty=True,
        )

    # Sector-only debug probe (e.g., transfers_unalloc).
    RUN_TRANSFORMATION_SECTOR_PROBE = True
    if RUN_TRANSFORMATION_SECTOR_PROBE:
        run_transformation_sector_probe(
            source_template=Path("data/leap results tables/transformation_results_20_USA_Reference.xlsx"),
            sector_prefix="transfers_unalloc",
            output_workbook=Path("outputs/leap_results_replica/transfers_unalloc_sector_probe.xlsx"),
        )
#%%
