#%%
from __future__ import annotations

import base64
import os
import platform
import sys
from pathlib import Path
import pandas as pd
import math

# Allow repo root on sys.path so codebase imports resolve without install.
REPO_ROOT = Path(__file__).resolve().parents[1]
if REPO_ROOT.exists() and str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.functions.leap_series_comparison import (
    ComparisonArtifacts,
    TransportResultsComparisonConfig,
    run_transport_results_table_comparison,
)

#----------------------------------------------------------------------------
# Notebook-editable configuration
#----------------------------------------------------------------------------
LEAP_RESULTS_FILE = "../data/leap results tables/transport all results.xlsx"
ECONOMY = "00_APEC"
SCENARIO = "Target"
REGION = "United States of America"

BRANCH_SECTOR_MAPPING_CSV = "../config/leap_transport_branch_to_ninth_sector_map.csv"
FUEL_ALIASES_CSV = "../config/leap_transport_fuel_aliases.csv"
CODE_TO_NAME_PATH = "../config/sector_fuel_codes_to_names.xlsx"
CODE_TO_NAME_SHEET = "code_to_name"

ESTO_DATA_PATH = "../data/00APEC_2024_low.csv"
NINTH_DATA_PATH = "../data/merged_file_energy_ALL_20250814_pre_trump.csv"
SUBTOTAL_MAPPING_PATH = "../config/ESTO_subtotal_mapping.xlsx"
NINTH_TO_ESTO_MAPPING_PATH = "../config/ninth_pairs_to_esto_pairs.xlsx"

BASE_YEAR = 2022
PROJECTION_START_YEAR = 2023
PROJECTION_END_YEAR = 2061
SHARE_YEAR_OFFSET = 1
NINTH_SCENARIO = "target"

OUTPUT_DIR = "../outputs/transport_results_series_comparison/usa_target"
DISPLAY_CHARTS_IN_NOTEBOOK = False
DISPLAY_CHARTS_AS_DASHBOARD = True
GENERATE_SHEET_DASHBOARDS = True
OPEN_DASHBOARD_INDEX_IN_BROWSER = True
OPEN_CHARTS_IN_WINDOWS_VIEWER = False


def _resolve_repo_path(path_value: str | Path) -> Path:
    path = Path(path_value)
    if path.is_absolute():
        return path
    return (Path(__file__).resolve().parent / path).resolve()


def build_config(
    leap_results_file: str | Path = LEAP_RESULTS_FILE,
    economy: str = ECONOMY,
    scenario: str = SCENARIO,
    region: str = REGION,
    branch_sector_mapping_csv: str | Path = BRANCH_SECTOR_MAPPING_CSV,
    fuel_aliases_csv: str | Path = FUEL_ALIASES_CSV,
    code_to_name_path: str | Path = CODE_TO_NAME_PATH,
    code_to_name_sheet: str = CODE_TO_NAME_SHEET,
    esto_data_path: str | Path = ESTO_DATA_PATH,
    ninth_data_path: str | Path = NINTH_DATA_PATH,
    subtotal_mapping_path: str | Path = SUBTOTAL_MAPPING_PATH,
    ninth_to_esto_mapping_path: str | Path = NINTH_TO_ESTO_MAPPING_PATH,
    base_year: int = BASE_YEAR,
    projection_start_year: int = PROJECTION_START_YEAR,
    projection_end_year: int = PROJECTION_END_YEAR,
    share_year_offset: int = SHARE_YEAR_OFFSET,
    ninth_scenario: str = NINTH_SCENARIO,
    output_dir: str | Path = OUTPUT_DIR,
) -> TransportResultsComparisonConfig:
    return TransportResultsComparisonConfig(
        leap_results_file=_resolve_repo_path(leap_results_file),
        economy=economy,
        scenario=scenario,
        region=region,
        branch_sector_mapping_csv=_resolve_repo_path(branch_sector_mapping_csv),
        fuel_aliases_csv=_resolve_repo_path(fuel_aliases_csv),
        code_to_name_path=_resolve_repo_path(code_to_name_path),
        code_to_name_sheet=code_to_name_sheet,
        esto_data_path=_resolve_repo_path(esto_data_path),
        ninth_data_path=_resolve_repo_path(ninth_data_path),
        subtotal_mapping_path=_resolve_repo_path(subtotal_mapping_path),
        ninth_to_esto_mapping_path=_resolve_repo_path(ninth_to_esto_mapping_path),
        base_year=base_year,
        projection_start_year=projection_start_year,
        projection_end_year=projection_end_year,
        share_year_offset=share_year_offset,
        ninth_scenario=ninth_scenario,
        output_dir=_resolve_repo_path(output_dir),
    )


def run_with_config(config: TransportResultsComparisonConfig | None = None) -> ComparisonArtifacts:
    cfg = config or build_config()
    artifacts = run_transport_results_table_comparison(cfg)
    print("[OK] LEAP series analysis finished.")
    print(f"- comparison_long_csv: {artifacts.comparison_long_csv}")
    print(f"- comparison_wide_csv: {artifacts.comparison_wide_csv}")
    print(f"- comparison_summary_csv: {artifacts.comparison_summary_csv}")
    print(f"- mapping_status_csv: {artifacts.mapping_status_csv}")
    print(f"- unmatched_leap_rows_csv: {artifacts.unmatched_leap_rows_csv}")
    print(f"- charts_dir: {artifacts.charts_dir}")
    print(f"- sheet_inventory_csv: {Path(cfg.output_dir) / 'sheet_inventory.csv'}")
    print(f"- fuel_mapping_status_csv: {Path(cfg.output_dir) / 'fuel_mapping_status.csv'}")
    dashboard_index_path: Path | None = None
    if GENERATE_SHEET_DASHBOARDS:
        dashboard_index_path = _build_sheet_dashboards(
            output_dir=cfg.output_dir,
            comparison_long_csv=artifacts.comparison_long_csv,
            charts_dir=artifacts.charts_dir,
        )
    if OPEN_DASHBOARD_INDEX_IN_BROWSER and dashboard_index_path is not None:
        _open_dashboard_index(dashboard_index_path)
    elif OPEN_CHARTS_IN_WINDOWS_VIEWER:
        _open_charts_in_windows_viewer(artifacts.charts_dir)
    if DISPLAY_CHARTS_IN_NOTEBOOK:
        if DISPLAY_CHARTS_AS_DASHBOARD:
            _display_charts_dashboard(artifacts.charts_dir)
        else:
            _display_charts_inline(artifacts.charts_dir)
    return artifacts


def _safe_filename_token(value: object) -> str:
    text = str(value).strip() if value is not None else ""
    if not text:
        return "series"
    safe = "".join(ch if ch.isalnum() or ch in {"_", "-"} else "_" for ch in text)
    return safe.strip("_") or "series"


def _build_sheet_dashboards(
    output_dir: str | Path,
    comparison_long_csv: str | Path,
    charts_dir: str | Path,
) -> Path | None:
    comparison_path = Path(comparison_long_csv)
    chart_path = Path(charts_dir)
    if not comparison_path.exists():
        print(f"[INFO] comparison_long.csv not found: {comparison_path}")
        return None
    if not chart_path.exists():
        print(f"[INFO] Charts directory not found: {chart_path}")
        return None

    df = pd.read_csv(comparison_path)
    required_cols = {"branch_path", "fuel_label"}
    if not required_cols.issubset(df.columns):
        print("[INFO] comparison_long.csv missing branch/fuel columns; skipping dashboards.")
        return None

    dashboards_dir = Path(output_dir) / "dashboards"
    dashboards_dir.mkdir(parents=True, exist_ok=True)

    branch_to_entries: dict[str, list[tuple[str, Path]]] = {}
    unique_pairs = (
        df[["branch_path", "fuel_label"]]
        .dropna()
        .drop_duplicates()
        .sort_values(["branch_path", "fuel_label"])
    )
    for _, row in unique_pairs.iterrows():
        branch = str(row["branch_path"]).strip()
        fuel = str(row["fuel_label"]).strip()
        if not branch or not fuel:
            continue
        branch_slug = _safe_filename_token(branch.replace("\\", "_"))
        fuel_slug = _safe_filename_token(fuel)
        png_path = chart_path / f"{branch_slug}__{fuel_slug}.png"
        if not png_path.exists():
            continue
        branch_to_entries.setdefault(branch, []).append((fuel, png_path))

    if not branch_to_entries:
        print("[INFO] No matching chart files were found for branch dashboards.")
        return None

    branch_files: list[tuple[str, Path, int]] = []
    for branch, entries in sorted(branch_to_entries.items(), key=lambda item: item[0].lower()):
        branch_slug = _safe_filename_token(branch.replace("\\", "_"))
        dashboard_file = dashboards_dir / f"{branch_slug}.html"
        cols = 4 if len(entries) >= 4 else max(1, len(entries))
        rows = max(1, math.ceil(len(entries) / cols))
        visible_rows = 4 if rows >= 4 else rows
        cards = []
        for fuel, png_path in entries:
            rel_png = os.path.relpath(png_path, start=dashboards_dir).replace("\\", "/")
            cards.append(
                f"""
<section class="card">
  <h2>{fuel}</h2>
  <img src="{rel_png}" alt="{branch} | {fuel}" loading="lazy"/>
</section>
"""
            )
        html_doc = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>{branch} Dashboard</title>
  <style>
    :root {{
      color-scheme: light;
      --cols: {cols};
      --visible-rows: {visible_rows};
      --gap: 6px;
      --header-h: 38px;
      --tile-h: calc((100vh - var(--header-h) - ((var(--visible-rows) + 1) * var(--gap))) / var(--visible-rows));
    }}
    body {{ margin: 0; font-family: Segoe UI, Arial, sans-serif; background: #f4f6f8; color: #111; }}
    header {{ position: sticky; top: 0; background: #0b3d5c; color: #fff; padding: 8px 12px; z-index: 2; min-height: var(--header-h); box-sizing: border-box; }}
    header h1 {{ margin: 0; font-size: 14px; }}
    main {{
      box-sizing: border-box;
      height: calc(100vh - var(--header-h));
      overflow-y: auto;
      padding: var(--gap);
      display: grid;
      grid-template-columns: repeat(var(--cols), minmax(0, 1fr));
      grid-auto-rows: var(--tile-h);
      gap: var(--gap);
      align-items: stretch;
    }}
    .card {{
      background: #fff;
      border-radius: 6px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.08);
      padding: 4px;
      display: flex;
      flex-direction: column;
      min-height: 0;
      overflow: hidden;
    }}
    .card h2 {{
      margin: 0 0 3px 0;
      font-size: 12px;
      line-height: 1.2;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .card img {{
      width: 100%;
      height: 100%;
      min-height: 0;
      object-fit: contain;
      display: block;
      border: 1px solid #ddd;
      background: #fff;
      flex: 1 1 auto;
    }}
    @media (max-width: 1400px) {{ :root {{ --cols: 3; }} }}
    @media (max-width: 1000px) {{ :root {{ --cols: 2; }} }}
    @media (max-width: 700px) {{ :root {{ --cols: 1; --visible-rows: 2; }} }}
  </style>
</head>
<body>
  <header><h1>{branch} ({len(entries)} charts)</h1></header>
  <main>
    {''.join(cards)}
  </main>
</body>
</html>
"""
        dashboard_file.write_text(html_doc, encoding="utf-8")
        branch_files.append((branch, dashboard_file, len(entries)))

    links = []
    for branch, file_path, count in branch_files:
        rel = file_path.name
        links.append(f'<li><a href="{rel}">{branch}</a> ({count} charts)</li>')
    index_file = dashboards_dir / "index.html"
    index_html = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>LEAP Sheet Dashboards</title>
  <style>
    body {{ font-family: Segoe UI, Arial, sans-serif; margin: 24px; background: #f4f6f8; color: #111; }}
    h1 {{ margin-top: 0; }}
    ul {{ line-height: 1.8; }}
    a {{ color: #0b3d5c; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
  </style>
</head>
<body>
  <h1>LEAP Sheet Dashboards</h1>
  <p>{len(branch_files)} dashboards generated.</p>
  <ul>
    {''.join(links)}
  </ul>
</body>
</html>
"""
    index_file.write_text(index_html, encoding="utf-8")
    print(f"[INFO] Generated {len(branch_files)} branch dashboards in {dashboards_dir}")
    return index_file


def _open_dashboard_index(index_path: str | Path) -> None:
    path = Path(index_path)
    if not path.exists():
        print(f"[INFO] Dashboard index not found: {path}")
        return
    if platform.system().lower() != "windows":
        print(f"[INFO] Dashboard index ready: {path}")
        return
    try:
        os.startfile(str(path))  # type: ignore[attr-defined]
        print(f"[INFO] Opened dashboard index: {path}")
    except Exception as exc:
        print(f"[WARN] Failed to open dashboard index: {exc}")


def _display_charts_inline(charts_dir: str | Path) -> None:
    chart_path = Path(charts_dir)
    if not chart_path.exists():
        print(f"[INFO] Charts directory not found: {chart_path}")
        return
    png_files = sorted(chart_path.glob("*.png"))
    if not png_files:
        print(f"[INFO] No chart PNGs found in: {chart_path}")
        return
    try:
        from IPython.display import Image, display
    except Exception:
        print("[INFO] IPython display is unavailable; skipping inline chart display.")
        return
    print(f"[INFO] Displaying {len(png_files)} chart(s) inline from {chart_path}")
    for png in png_files:
        print(f"[CHART] {png.name}")
        display(Image(filename=str(png)))


def _display_charts_dashboard(charts_dir: str | Path) -> None:
    chart_path = Path(charts_dir)
    if not chart_path.exists():
        print(f"[INFO] Charts directory not found: {chart_path}")
        return
    png_files = sorted(chart_path.glob("*.png"))
    if not png_files:
        print(f"[INFO] No chart PNGs found in: {chart_path}")
        return
    try:
        from IPython.display import HTML, display
        import ipywidgets as widgets
    except Exception:
        print(
            "[INFO] ipywidgets/IPython display is unavailable; falling back to inline image list."
        )
        _display_charts_inline(charts_dir)
        return

    def _to_img_html(path: Path) -> str:
        encoded = base64.b64encode(path.read_bytes()).decode("ascii")
        return f"""
<div style="width:100vw;height:88vh;background:#111;display:flex;align-items:center;justify-content:center;overflow:hidden;">
  <img src="data:image/png;base64,{encoded}" style="max-width:100vw;max-height:88vh;object-fit:contain;" />
</div>
"""

    file_names = [path.name for path in png_files]
    file_lookup = {path.name: path for path in png_files}
    selector = widgets.Dropdown(
        options=file_names,
        value=file_names[0],
        description="Chart:",
        layout=widgets.Layout(width="50%"),
    )
    image_html = widgets.HTML(value=_to_img_html(file_lookup[file_names[0]]))

    def _set_chart(file_name: str) -> None:
        image_html.value = _to_img_html(file_lookup[file_name])

    def _on_selector_change(change: dict[str, object]) -> None:
        if change.get("name") == "value" and change.get("new"):
            _set_chart(str(change["new"]))

    selector.observe(_on_selector_change)

    prev_button = widgets.Button(description="Prev", icon="arrow-left")
    next_button = widgets.Button(description="Next", icon="arrow-right")

    def _shift(delta: int) -> None:
        current_idx = file_names.index(selector.value)
        selector.value = file_names[(current_idx + delta) % len(file_names)]

    prev_button.on_click(lambda _btn: _shift(-1))
    next_button.on_click(lambda _btn: _shift(1))

    header = widgets.HTML(
        value="<h3 style='margin:4px 0 8px 0;'>LEAP Series Comparison Dashboard</h3>"
    )
    controls = widgets.HBox([selector, prev_button, next_button])
    container = widgets.VBox([header, controls, image_html])
    display(container)
    display(
        HTML(
            "<style>.jupyter-widgets-output-area .output_scroll {height: auto !important;}</style>"
        )
    )
    print(f"[INFO] Dashboard loaded with {len(png_files)} chart(s).")


def _open_charts_in_windows_viewer(charts_dir: str | Path) -> None:
    chart_path = Path(charts_dir)
    if not chart_path.exists():
        print(f"[INFO] Charts directory not found: {chart_path}")
        return
    png_files = sorted(chart_path.glob("*.png"))
    if not png_files:
        print(f"[INFO] No chart PNGs found in: {chart_path}")
        return
    if platform.system().lower() != "windows":
        print("[INFO] External Windows image viewer is only available on Windows.")
        return
    try:
        os.startfile(str(png_files[0]))  # type: ignore[attr-defined]
        print(f"[INFO] Opened chart in Windows viewer: {png_files[0].name}")
        if len(png_files) > 1:
            print(
                f"[INFO] {len(png_files)} charts generated in {chart_path}. "
                "Use next/prev in viewer or open folder for more."
            )
    except Exception as exc:
        print(f"[WARN] Failed to open chart in Windows viewer: {exc}")


#%%
if __name__ == "__main__":
    run_with_config()
#%%
