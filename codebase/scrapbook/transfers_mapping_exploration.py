
#%%
"""
Scan ESTO transfer flows to flag possible transfer categories per economy.

Outputs
- transfer_category_flags.csv: economy-level flags + transfer row details.
- transfer_category_io_totals.csv: total inputs/outputs per category per economy.
- optional plots (category templates + process mapping facets).

Notes
- Uses TRANSFER_CATEGORY_TEMPLATES from codebase.transfers_workflow.
- Sums the last N years (default 5) of 00APEC_2024_low_with_subtotals.csv.
- Drops subtotal rows by default.
- Paste the printed TRANSFER_PROCESS_CONFIG into
  codebase/transfers_workflow.py after updating templates.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Iterable

import pandas as pd

# Allow the repository root to be importable regardless of the working directory.
REPO_ROOT = Path(__file__).resolve().parents[2]
CURRENT_DIR = Path.cwd()
if CURRENT_DIR != REPO_ROOT:
    os.chdir(REPO_ROOT)
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

from codebase.transfers_workflow import (
    TRANSFER_CATEGORY_TEMPLATES,
    TRANSFER_FLOW_CODES,
    TRANSFER_PROCESS_CONFIG,
    TRANSFER_ECONOMY_CONFIG_ALIASES,
    TRANSFER_COMBINED_FLOW_KEY,
)


DATA_PATH = Path("data/00APEC_2024_low_with_subtotals.csv")  # Source file with subtotals labeled.
OUTPUT_DIR = Path("outputs/transfer_category_scan")  # Root output directory for CSVs and plots.
PLOT_OUTPUT_DIR = OUTPUT_DIR / "plots"  # Plot output root.
ECONOMY_PLOT_DIR = PLOT_OUTPUT_DIR / "economies"  # Per-economy facet plot output directory.
STITCHED_PLOT_PATH = PLOT_OUTPUT_DIR / "all_economies_transfer_facets.png"  # Stitched grid plot.
PROCESS_ECONOMY_PLOT_DIR = PLOT_OUTPUT_DIR / "economies_process_config"  # Per-economy process mapping plots.
STITCHED_PROCESS_PLOT_PATH = (
    PLOT_OUTPUT_DIR / "all_economies_transfer_facets_process_config.png"
)  # Stitched grid for process mappings.
LAST_N_YEARS = 5  # Number of most recent years to sum for the scan.
DROP_SUBTOTALS = True  # Drop rows flagged as subtotals before analysis.
DROP_TOTAL_ENERGY_ROWS = True  # Drop total/renewables aggregate product rows.
MAKE_PLOTS = True  # Toggle plot generation.
MAKE_PROCESS_PLOTS = True  # Toggle plots based on TRANSFER_PROCESS_CONFIG.
INCLUDE_AGGREGATE_ECONOMY = True  # Append an aggregate economy rowset for config review.
AGGREGATE_ECONOMY_LABEL = "00_APEC"  # Aggregate economy label for config output.
AUTO_UPDATE_TRANSFER_PROCESS_CONFIG = False  # Toggle auto-updating TRANSFER_PROCESS_CONFIG in plan.
CATEGORY_PLOT_COLUMNS = 2  # Columns per economy facet grid (categories across subplots).
STITCH_PLOT_COLUMNS = 3  # Columns in the stitched grid of economy plots.
MAX_PLOT_PRODUCTS = 100  # Limit products displayed per economy (by abs total).
SHOW_ZERO_PROCESS_PRODUCTS = True  # Keep mapped products even if totals are zero.
PLOT_DPI = 160  # DPI for saved plots.
RATIO_WARNING_LOW = 0.0  # Highlight ratios below this threshold.
RATIO_WARNING_HIGH = float("inf")  # Highlight ratios above this threshold.
OTHERS_CATEGORY_NAME = "Others"
STITCH_TILE_PADDING = 8  # Pixels of padding around each stitched tile.
STITCH_BORDER_WIDTH = 4  # Border width (px) around each stitched economy tile.
STITCH_BORDER_COLOR = (0, 0, 0)  # Black border.
STITCH_BACKGROUND_COLOR = (255, 255, 255)  # White canvas background.


def _normalize_economy_codes(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize economy codes to include underscores (e.g., 01AUS -> 01_AUS)."""
    if "economy" not in df.columns:
        return df
    updated = df.copy()
    updated["economy"] = (
        updated["economy"]
        .astype(str)
        .str.replace(r"^(\d{2})([A-Z].+)$", r"\1_\2", regex=True)
    )
    return updated


def _add_aggregate_economy_total(
    df: pd.DataFrame, year_cols: list[int], economy_label: str
) -> pd.DataFrame:
    """Append an aggregate economy total row set to a dataset."""
    if "economy" not in df.columns or df.empty or not year_cols:
        return df
    if df["economy"].astype(str).eq(economy_label).any():
        return df
    group_cols = [col for col in df.columns if col not in year_cols and col != "economy"]
    totals = df.groupby(group_cols, dropna=False)[year_cols].sum().reset_index()
    totals["economy"] = economy_label
    totals = totals[df.columns.tolist()]
    return pd.concat([df, totals], ignore_index=True)


def _find_year_cols(df: pd.DataFrame) -> list[int]:
    """Return sorted year columns as ints."""
    year_cols = [int(col) for col in df.columns if str(col).isdigit()]
    return sorted(year_cols)


def _normalize_year_columns(df: pd.DataFrame) -> tuple[pd.DataFrame, list[int]]:
    """Convert digit-like year columns to int and return (df, year_cols)."""
    year_cols = _find_year_cols(df)
    if not year_cols:
        return df, []
    df = df.copy()
    df.columns = [int(col) if str(col).isdigit() else col for col in df.columns]
    return df, year_cols


def _coerce_years_numeric(df: pd.DataFrame, year_cols: list[int]) -> pd.DataFrame:
    """Ensure year columns are numeric for summing."""
    updated = df.copy()
    updated[year_cols] = updated[year_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    return updated


def _drop_total_energy_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Drop total/renewables summary rows from products."""
    total_labels = {
        "19 Total",
        "20 Total Renewables",
        "21 Modern renewables",
    }
    if "products" not in df.columns:
        return df
    return df[~df["products"].astype(str).isin(total_labels)].copy()


def _get_last_n_years(year_cols: list[int], n: int) -> list[int]:
    """Return last N years from the year columns."""
    if not year_cols:
        return []
    return year_cols[-n:] if len(year_cols) >= n else year_cols


def _sum_last_n_years(df: pd.DataFrame, year_cols: list[int], last_n: int) -> tuple[pd.DataFrame, list[int]]:
    """Attach last_n_sum column and return the used years."""
    last_years = _get_last_n_years(year_cols, last_n)
    if not last_years:
        df["last_n_sum"] = 0.0
        return df, last_years
    df = _coerce_years_numeric(df, last_years)
    df["last_n_sum"] = df[last_years].sum(axis=1)
    return df, last_years


def _build_long_df(df: pd.DataFrame, year_label: str) -> pd.DataFrame:
    """Return a long df with economy, flow, product, year, value."""
    base_cols = ["economy", "flows", "products", "last_n_sum"]
    for col in base_cols:
        if col not in df.columns:
            raise ValueError(f"Missing required column: {col}")
    long_df = (
        df.groupby(["economy", "flows", "products"], dropna=False)["last_n_sum"]
        .sum()
        .reset_index()
        .rename(columns={"flows": "flow", "products": "product", "last_n_sum": "value"})
    )
    long_df["year"] = year_label
    return long_df


def _update_transfer_plan_config(config_text: str) -> None:
    """Replace TRANSFER_PROCESS_CONFIG in transfers_workflow.py."""
    target_path = REPO_ROOT / "codebase" / "transfers_workflow.py"
    if not target_path.exists():
        raise FileNotFoundError(f"{target_path} not found")
    plan_text = target_path.read_text(encoding="utf-8")
    marker = "TRANSFER_PROCESS_CONFIG:"
    start = plan_text.find(marker)
    if start == -1:
        raise ValueError("TRANSFER_PROCESS_CONFIG not found in transfers_workflow.py")
    brace_start = plan_text.find("{", start)
    if brace_start == -1:
        raise ValueError("TRANSFER_PROCESS_CONFIG opening brace not found")
    level = 0
    end = None
    for idx, ch in enumerate(plan_text[brace_start:], start=brace_start):
        if ch == "{":
            level += 1
        elif ch == "}":
            level -= 1
            if level == 0:
                end = idx + 1
                break
    if end is None:
        raise ValueError("TRANSFER_PROCESS_CONFIG closing brace not found")
    line_start = plan_text.rfind("\n", 0, start) + 1
    new_block = f"TRANSFER_PROCESS_CONFIG: dict[str, dict[str, list[dict]]] = {config_text}"
    updated = plan_text[:line_start] + new_block + plan_text[end:]
    target_path.write_text(updated, encoding="utf-8")


def _build_economy_product_totals(long_df: pd.DataFrame) -> pd.DataFrame:
    """Return economy/product totals (last_n_sum) aggregated over transfer flows."""
    return (
        long_df.groupby(["economy", "product"], dropna=False)["value"]
        .sum()
        .reset_index()
    )


def _category_applicability(
    economy_product_totals: pd.DataFrame,
    templates: Iterable[dict],
) -> pd.DataFrame:
    """Return economy-level applicability flags for each category."""
    economies = economy_product_totals["economy"].dropna().unique().tolist()
    flags = []
    for economy in economies:
        econ_totals = economy_product_totals[economy_product_totals["economy"] == economy]
        econ_map = {
            row["product"]: row["value"] for _, row in econ_totals.iterrows()
        }
        row_flags = {"economy": economy}
        for template in templates:
            category = template["category"]
            if template.get("mode") == "others":
                known_products = set()
                for other in TRANSFER_CATEGORY_TEMPLATES:
                    if other.get("mode") == "others":
                        continue
                    known_products.update(_category_products(other))
                products = set(econ_map.keys()) - known_products
            else:
                products = _category_products(template)
            has_input = any(econ_map.get(label, 0.0) < 0 for label in products)
            has_output = any(econ_map.get(label, 0.0) > 0 for label in products)
            row_flags[category] = bool(has_input and has_output)
        flags.append(row_flags)
    return pd.DataFrame(flags)


def _category_io_totals(
    economy_product_totals: pd.DataFrame,
    templates: Iterable[dict],
) -> pd.DataFrame:
    """Return total input/output per category per economy using last_n_sum."""
    rows = []
    for economy in economy_product_totals["economy"].dropna().unique():
        econ_totals = economy_product_totals[economy_product_totals["economy"] == economy]
        econ_map = {
            row["product"]: row["value"] for _, row in econ_totals.iterrows()
        }
        for template in templates:
            category = template["category"]
            if template.get("mode") == "others":
                known_products = set()
                for other in TRANSFER_CATEGORY_TEMPLATES:
                    if other.get("mode") == "others":
                        continue
                    known_products.update(_category_products(other))
                products = set(econ_map.keys()) - known_products
            else:
                products = _category_products(template)
            inputs = [econ_map.get(label, 0.0) for label in products]
            outputs = [econ_map.get(label, 0.0) for label in products]
            input_total = sum(value for value in inputs if value < 0)
            output_total = sum(value for value in outputs if value > 0)
            rows.append(
                {
                    "economy": economy,
                    "category": category,
                    "input_total": input_total,
                    "output_total": output_total,
                    "net_total": input_total + output_total,
                    "applicable": bool(input_total < 0 and output_total > 0),
                }
            )
    return pd.DataFrame(rows)


def _build_process_config_rows(
    economy_product_totals: pd.DataFrame,
    templates: Iterable[dict],
    year_label: str,
    flow_label: str = "transfer_flows_combined",
) -> pd.DataFrame:
    """Build rows that can be pasted into TRANSFER_PROCESS_CONFIG."""
    rows = []
    for economy in economy_product_totals["economy"].dropna().unique():
        econ_totals = economy_product_totals[economy_product_totals["economy"] == economy]
        econ_map = {row["product"]: row["value"] for _, row in econ_totals.iterrows()}
        category_data = {}
        for template in templates:
            category = template["category"]
            if template.get("mode") == "others":
                known_products = set()
                for other in TRANSFER_CATEGORY_TEMPLATES:
                    if other.get("mode") == "others":
                        continue
                    known_products.update(_category_products(other))
                products = set(econ_map.keys()) - known_products
            else:
                products = _category_products(template)
            inputs = [label for label in products if econ_map.get(label, 0.0) < 0]
            outputs = [label for label in products if econ_map.get(label, 0.0) > 0]
            input_total = sum(econ_map.get(label, 0.0) for label in inputs)
            output_total = sum(econ_map.get(label, 0.0) for label in outputs)
            ratio = abs(input_total) / output_total if output_total else None
            category_data[category] = {
                "inputs": inputs,
                "outputs": outputs,
                "input_total": input_total,
                "output_total": output_total,
                "ratio": ratio,
            }
            if template.get("mode") == "others" and (inputs or outputs):
                raise ValueError(
                    f"Unmapped transfer fuels detected for {economy}: {sorted(products)}"
                )

        def _merge_categories(primary: str, secondary: str) -> dict:
            """Merge two categories into a combined process entry."""
            primary_data = category_data.get(primary, {})
            secondary_data = category_data.get(secondary, {})
            inputs = sorted(set(primary_data.get("inputs", [])) | set(secondary_data.get("inputs", [])))
            outputs = sorted(set(primary_data.get("outputs", [])) | set(secondary_data.get("outputs", [])))
            input_total = sum(econ_map.get(label, 0.0) for label in inputs)
            output_total = sum(econ_map.get(label, 0.0) for label in outputs)
            ratio = abs(input_total) / output_total if output_total else None
            return {
            "category": "Upstream & refinery transfers",
            "process": "Upstream & refinery transfers",
                "inputs": inputs,
                "outputs": outputs,
                "input_total": input_total,
                "output_total": output_total,
                "ratio": ratio,
            }

        merged_categories = set()
        for category, data in category_data.items():
            has_inputs = bool(data["inputs"])
            has_outputs = bool(data["outputs"])
            if has_inputs and has_outputs:
                continue
            if category in merged_categories:
                continue
            if category == OTHERS_CATEGORY_NAME:
                continue
            # Prefer merging with a category that already has both sides.
            target = None
            for other_category, other_data in category_data.items():
                if other_category == category or other_category in merged_categories:
                    continue
                if other_category == OTHERS_CATEGORY_NAME:
                    continue
                if other_data["inputs"] and other_data["outputs"]:
                    target = other_category
                    break
            # Fallback: merge with a category that has the missing side.
            if target is None:
                for other_category, other_data in category_data.items():
                    if other_category == category or other_category in merged_categories:
                        continue
                    if other_category == OTHERS_CATEGORY_NAME:
                        continue
                    if not other_data["inputs"] and not other_data["outputs"]:
                        continue
                    if has_inputs and other_data["outputs"]:
                        target = other_category
                        break
                    if has_outputs and other_data["inputs"]:
                        target = other_category
                        break
            if target:
                merged = _merge_categories(category, target)
                if merged["inputs"] and merged["outputs"]:
                    rows.append(
                        {
                            "economy": economy,
                            "flow": flow_label,
                            "category": merged["category"],
                            "process": merged["process"],
                            "inputs": "; ".join(merged["inputs"]),
                            "outputs": "; ".join(merged["outputs"]),
                            "input_total": merged["input_total"],
                            "output_total": merged["output_total"],
                            "ratio": merged["ratio"],
                            "year": year_label,
                        }
                    )
                    merged_categories.add(category)
                    merged_categories.add(target)

        for category, data in category_data.items():
            if category in merged_categories:
                continue
            if not data["inputs"] or not data["outputs"]:
                continue
            rows.append(
                {
                    "economy": economy,
                    "flow": flow_label,
                    "category": category,
                    "process": category,
                    "inputs": "; ".join(sorted(data["inputs"])),
                    "outputs": "; ".join(sorted(data["outputs"])),
                    "input_total": data["input_total"],
                    "output_total": data["output_total"],
                    "ratio": data["ratio"],
                    "year": year_label,
                }
            )
    return pd.DataFrame(rows)


def build_transfer_process_config(process_config_df: pd.DataFrame) -> dict:
    """Return TRANSFER_PROCESS_CONFIG-style dict from candidate rows."""
    config: dict[str, dict[str, list[dict]]] = {}
    for _, row in process_config_df.iterrows():
        economy = str(row.get("economy", "")).strip()
        if not economy:
            continue
        process_name = str(row.get("process", "")).strip()
        if not process_name:
            continue
        flow = str(row.get("flow", "")).strip() or "08 Transfers"
        inputs_raw = str(row.get("inputs", "")).strip()
        outputs_raw = str(row.get("outputs", "")).strip()
        inputs = [item.strip() for item in inputs_raw.split(";") if item.strip()]
        outputs = [item.strip() for item in outputs_raw.split(";") if item.strip()]
        if not inputs or not outputs:
            continue
        config.setdefault(economy, {}).setdefault(flow, []).append(
            {
                "process": process_name,
                "inputs": inputs,
                "outputs": outputs,
            }
        )
    return config


def format_transfer_process_config(config: dict) -> str:
    """Return a pretty-printed Python dict string."""
    import json

    return json.dumps(config, indent=4, ensure_ascii=True)


def _assign_product_colors(products: Iterable[str]) -> dict[str, str]:
    """Assign stable colors to products."""
    import matplotlib.pyplot as plt

    palette = plt.get_cmap("tab20")
    products = list(products)
    return {
        product: palette(idx % palette.N) for idx, product in enumerate(sorted(products))
    }


def _economies_with_transfers(economy_product_totals: pd.DataFrame) -> list[str]:
    """Return economies that have any nonzero transfer activity."""
    totals = (
        economy_product_totals.groupby("economy")["value"]
        .apply(lambda s: s.abs().sum())
    )
    economies = totals[totals > 0].index.tolist()
    return sorted(economies)


def _category_products(template: dict) -> set[str]:
    """Return the full product set for a category template."""
    if template.get("mode") == "others":
        return set()
    return set(template.get("inputs", [])) | set(template.get("outputs", []))


def _category_totals_for_economy(
    economy_product_totals: pd.DataFrame,
    economy: str,
    template: dict,
) -> tuple[float, float, float | None]:
    """Return (input_total, output_total, ratio) for a category."""
    econ_df = economy_product_totals[economy_product_totals["economy"] == economy]
    econ_map = {row["product"]: row["value"] for _, row in econ_df.iterrows()}
    if template.get("mode") == "others":
        known_products = set()
        for other in TRANSFER_CATEGORY_TEMPLATES:
            if other.get("mode") == "others":
                continue
            known_products.update(_category_products(other))
        products = set(econ_map.keys()) - known_products
    else:
        products = _category_products(template)
    inputs = [econ_map.get(label, 0.0) for label in products]
    outputs = [econ_map.get(label, 0.0) for label in products]
    input_total = sum(value for value in inputs if value < 0)
    output_total = sum(value for value in outputs if value > 0)
    if output_total <= 0 or input_total >= 0:
        return input_total, output_total, None
    ratio = abs(input_total) / output_total if output_total else None
    return input_total, output_total, ratio


def _resolve_economy_process_config(
    process_config: dict[str, dict[str, list[dict]]],
    economy: str,
) -> dict[str, list[dict]] | None:
    """Return the process config for an economy, applying alias fallback."""
    economy_config = process_config.get(economy)
    if economy_config:
        return economy_config
    alias = TRANSFER_ECONOMY_CONFIG_ALIASES.get(economy)
    if alias:
        return process_config.get(alias)
    return None


def _collect_process_labels(process_cfg: dict) -> list[str]:
    """Return unique labels for a process config in declared order."""
    label_keys = ("inputs", "outputs", "products", "fuels", "labels")
    labels: list[str] = []
    for key in label_keys:
        values = process_cfg.get(key, [])
        if not values:
            continue
        for value in values:
            label = str(value).strip()
            if label:
                labels.append(label)
    seen: set[str] = set()
    return [label for label in labels if not (label in seen or seen.add(label))]


def _split_labels_by_sign(
    labels: Iterable[str],
    totals_map: dict[str, float],
) -> tuple[list[str], list[str]]:
    """Split labels into inputs/outputs based on sign in totals."""
    inputs = [label for label in labels if totals_map.get(label, 0.0) < 0]
    outputs = [label for label in labels if totals_map.get(label, 0.0) > 0]
    return inputs, outputs


def _build_process_plot_entries(
    economy_product_totals: pd.DataFrame,
    economy: str,
    process_config: dict[str, dict[str, list[dict]]],
    max_products: int | None = None,
) -> tuple[list[dict], dict[str, float], list[str]]:
    """Return process entries + totals map + label list for plotting."""
    econ_df = economy_product_totals[economy_product_totals["economy"] == economy]
    totals_map = {row["product"]: row["value"] for _, row in econ_df.iterrows()}
    economy_config = _resolve_economy_process_config(process_config, economy)
    if not economy_config:
        return [], totals_map, []

    entries: list[dict] = []
    all_labels: list[str] = []
    for flow_code, processes in economy_config.items():
        for process_cfg in processes:
            labels = _collect_process_labels(process_cfg)
            if not labels:
                continue
            inputs, outputs = _split_labels_by_sign(labels, totals_map)
            input_total = sum(totals_map.get(label, 0.0) for label in inputs)
            output_total = sum(totals_map.get(label, 0.0) for label in outputs)
            ratio = (
                abs(input_total) / output_total
                if (output_total > 0 and input_total < 0)
                else None
            )
            process_name = (
                str(
                    process_cfg.get("process")
                    or process_cfg.get("category")
                    or flow_code
                ).strip()
                or flow_code
            )
            entries.append(
                {
                    "flow": flow_code,
                    "process": process_name,
                    "labels": labels,
                    "inputs": inputs,
                    "outputs": outputs,
                    "input_total": input_total,
                    "output_total": output_total,
                    "ratio": ratio,
                }
            )
            all_labels.extend(labels)

    if not entries:
        return [], totals_map, []

    seen: set[str] = set()
    unique_labels = [label for label in all_labels if not (label in seen or seen.add(label))]
    if max_products and len(unique_labels) > max_products:
        label_scores = {
            label: abs(totals_map.get(label, 0.0)) for label in unique_labels
        }
        top_labels = sorted(label_scores, key=label_scores.get, reverse=True)[:max_products]
        allowed = set(top_labels)
        unique_labels = [label for label in unique_labels if label in allowed]
        filtered_entries: list[dict] = []
        for entry in entries:
            labels = [label for label in entry["labels"] if label in allowed]
            if not labels:
                continue
            inputs = [label for label in entry["inputs"] if label in allowed]
            outputs = [label for label in entry["outputs"] if label in allowed]
            input_total = sum(totals_map.get(label, 0.0) for label in inputs)
            output_total = sum(totals_map.get(label, 0.0) for label in outputs)
            ratio = (
                abs(input_total) / output_total
                if (output_total > 0 and input_total < 0)
                else None
            )
            filtered_entry = dict(entry)
            filtered_entry.update(
                {
                    "labels": labels,
                    "inputs": inputs,
                    "outputs": outputs,
                    "input_total": input_total,
                    "output_total": output_total,
                    "ratio": ratio,
                }
            )
            filtered_entries.append(filtered_entry)
        entries = filtered_entries

    return entries, totals_map, unique_labels


def select_best_category_set(
    economy_product_totals: pd.DataFrame,
    economy: str,
    templates: Iterable[dict],
    require_applicable: bool = True,
) -> list[dict]:
    """Deprecated selection helper; kept for backward compatibility."""
    return []


def plot_economy_facets(
    economy_product_totals: pd.DataFrame,
    templates: Iterable[dict],
    year_label: str,
    output_dir: Path,
    max_products: int = 16,
    category_cols: int = 2,
    dpi: int = 160,
) -> list[Path]:
    """Create one plot per economy with category facets."""
    import math
    import matplotlib.pyplot as plt

    output_dir.mkdir(parents=True, exist_ok=True)
    plot_paths: list[Path] = []
    economies = _economies_with_transfers(economy_product_totals)
    for economy in economies:
        econ_df = economy_product_totals[economy_product_totals["economy"] == economy].copy()
        econ_df = econ_df[econ_df["value"] != 0]
        if econ_df.empty:
            continue
        products = (
            econ_df.groupby("product")["value"]
            .apply(lambda s: s.abs().sum())
            .sort_values(ascending=False)
            .index.tolist()[:max_products]
        )
        econ_df = econ_df[econ_df["product"].isin(products)]
        if econ_df.empty:
            continue
        product_colors = _assign_product_colors(products)

        categories = list(templates)
        n_categories = len(categories)
        cols = max(1, category_cols)
        rows = math.ceil(n_categories / cols)
        fig, axes = plt.subplots(
            rows, cols, figsize=(cols * 5, rows * 3.8), sharex=False
        )
        axes = axes.flatten() if isinstance(axes, Iterable) else [axes]

        for idx, template in enumerate(categories):
            ax = axes[idx]
            category = template["category"]
            econ_map = {
                row["product"]: row["value"] for _, row in econ_df.iterrows()
            }
            if template.get("mode") == "others":
                known_products = set()
                for other in TRANSFER_CATEGORY_TEMPLATES:
                    if other.get("mode") == "others":
                        continue
                    known_products.update(_category_products(other))
                cat_products = set(econ_df["product"].unique()) - known_products
            else:
                cat_products = _category_products(template)
            cat_df = econ_df[econ_df["product"].isin(cat_products)].copy()
            cat_df = cat_df[cat_df["value"] != 0]
            input_total, output_total, ratio = _category_totals_for_economy(
                economy_product_totals, economy, template
            )
            ratio_text = "ratio n/a" if ratio is None else f"ratio {ratio:.2f}"
            ax.set_title(f"{category}\n{ratio_text}")
            if cat_df.empty:
                ax.axis("off")
                continue
            cat_df = cat_df.sort_values(
                by="value", key=lambda s: s.abs(), ascending=True
            )
            ax.axvline(0, color="black", linewidth=0.6)
            ax.barh(
                cat_df["product"],
                cat_df["value"],
                color=[product_colors[prod] for prod in cat_df["product"]],
            )
            ax.set_xlabel(f"Last {LAST_N_YEARS}y sum ({year_label})")
            if ratio is not None and (ratio < RATIO_WARNING_LOW or ratio > RATIO_WARNING_HIGH):
                ax.patch.set_edgecolor("gold")
                ax.patch.set_linewidth(2.5)

        for ax in axes[n_categories:]:
            ax.axis("off")

        fig.suptitle(f"{economy} Transfers ({year_label})")
        fig.tight_layout()
        out_path = output_dir / f"{economy}_transfer_facets.png"
        fig.savefig(out_path, dpi=dpi)
        plt.close(fig)
        plot_paths.append(out_path)
    return plot_paths


def plot_economy_process_facets(
    economy_product_totals: pd.DataFrame,
    process_config: dict[str, dict[str, list[dict]]],
    year_label: str,
    output_dir: Path,
    max_products: int | None = None,
    category_cols: int = 2,
    dpi: int = 160,
) -> list[Path]:
    """Create one plot per economy based on TRANSFER_PROCESS_CONFIG."""
    import math
    import matplotlib.pyplot as plt

    output_dir.mkdir(parents=True, exist_ok=True)
    plot_paths: list[Path] = []
    economies = _economies_with_transfers(economy_product_totals)
    for economy in economies:
        entries, totals_map, unique_labels = _build_process_plot_entries(
            economy_product_totals,
            economy,
            process_config,
            max_products=max_products,
        )
        if not entries or not unique_labels:
            continue
        product_colors = _assign_product_colors(unique_labels)

        n_categories = len(entries)
        cols = max(1, category_cols)
        rows = math.ceil(n_categories / cols)
        fig, axes = plt.subplots(
            rows, cols, figsize=(cols * 5, rows * 3.8), sharex=False
        )
        axes = axes.flatten() if isinstance(axes, Iterable) else [axes]

        for idx, entry in enumerate(entries):
            ax = axes[idx]
            labels = entry["labels"]
            if not labels:
                ax.axis("off")
                continue
            values = [totals_map.get(label, 0.0) for label in labels]
            cat_df = pd.DataFrame({"product": labels, "value": values})
            if not SHOW_ZERO_PROCESS_PRODUCTS:
                cat_df = cat_df[cat_df["value"] != 0]
            ratio = entry["ratio"]
            ratio_text = "ratio n/a" if ratio is None else f"ratio {ratio:.2f}"
            title_lines = [entry["process"]]
            if entry["flow"] and entry["flow"] != TRANSFER_COMBINED_FLOW_KEY:
                title_lines.append(entry["flow"])
            title_lines.append(ratio_text)
            ax.set_title("\n".join(title_lines))
            if cat_df.empty:
                ax.axis("off")
                continue
            ax.axvline(0, color="black", linewidth=0.6)
            ax.barh(
                cat_df["product"],
                cat_df["value"],
                color=[product_colors.get(prod, "grey") for prod in cat_df["product"]],
            )
            ax.set_xlabel(f"Last {LAST_N_YEARS}y sum ({year_label})")
            if ratio is not None and (ratio < RATIO_WARNING_LOW or ratio > RATIO_WARNING_HIGH):
                ax.patch.set_edgecolor("gold")
                ax.patch.set_linewidth(2.5)

        for ax in axes[n_categories:]:
            ax.axis("off")

        fig.suptitle(f"{economy} Transfers (process mapping, {year_label})")
        fig.tight_layout()
        out_path = output_dir / f"{economy}_transfer_process_facets.png"
        fig.savefig(out_path, dpi=dpi)
        plt.close(fig)
        plot_paths.append(out_path)
    return plot_paths


def stitch_economy_plots(
    plot_paths: Iterable[Path],
    output_path: Path,
    columns: int = 3,
) -> None:
    """Stitch per-economy plots into a single image."""
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        _stitch_with_matplotlib(plot_paths, output_path, columns=columns)
        return

    plot_paths = list(plot_paths)
    if not plot_paths:
        print("No economy plots found; skipping stitched plot.")
        return

    images = [Image.open(path).convert("RGB") for path in plot_paths]
    max_width = max(img.width for img in images)
    max_height = max(img.height for img in images)
    tile_width = max_width + (2 * STITCH_TILE_PADDING)
    tile_height = max_height + (2 * STITCH_TILE_PADDING)
    cols = max(1, columns)
    rows = (len(images) + cols - 1) // cols
    stitched = Image.new(
        "RGB",
        (cols * tile_width, rows * tile_height),
        STITCH_BACKGROUND_COLOR,
    )

    for idx, img in enumerate(images):
        if img.width != max_width or img.height != max_height:
            padded = Image.new("RGB", (max_width, max_height), STITCH_BACKGROUND_COLOR)
            padded.paste(img, (0, 0))
            img = padded

        tile = Image.new("RGB", (tile_width, tile_height), STITCH_BACKGROUND_COLOR)
        tile.paste(img, (STITCH_TILE_PADDING, STITCH_TILE_PADDING))

        border_left = STITCH_TILE_PADDING
        border_top = STITCH_TILE_PADDING
        border_right = STITCH_TILE_PADDING + max_width - 1
        border_bottom = STITCH_TILE_PADDING + max_height - 1
        draw = ImageDraw.Draw(tile)
        draw.rectangle(
            [border_left, border_top, border_right, border_bottom],
            outline=STITCH_BORDER_COLOR,
            width=STITCH_BORDER_WIDTH,
        )

        x = (idx % cols) * tile_width
        y = (idx // cols) * tile_height
        stitched.paste(tile, (x, y))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    stitched.save(output_path)


def _stitch_with_matplotlib(
    plot_paths: Iterable[Path],
    output_path: Path,
    columns: int = 3,
) -> None:
    """Fallback stitcher using matplotlib when PIL is unavailable."""
    import math
    import matplotlib.pyplot as plt
    import matplotlib.image as mpimg
    import matplotlib.patches as patches

    plot_paths = list(plot_paths)
    if not plot_paths:
        print("No economy plots found; skipping stitched plot.")
        return
    cols = max(1, columns)
    rows = math.ceil(len(plot_paths) / cols)
    fig, axes = plt.subplots(rows, cols, figsize=(cols * 4, rows * 3))
    axes = axes.flatten() if isinstance(axes, Iterable) else [axes]
    for ax, path in zip(axes, plot_paths):
        img = mpimg.imread(path)
        ax.imshow(img)
        ax.add_patch(
            patches.Rectangle(
                (0, 0),
                1,
                1,
                transform=ax.transAxes,
                fill=False,
                edgecolor="black",
                linewidth=max(1.0, STITCH_BORDER_WIDTH / 2),
            )
        )
        ax.axis("off")
    for ax in axes[len(plot_paths):]:
        ax.axis("off")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.tight_layout()
    fig.savefig(output_path, dpi=PLOT_DPI)
    plt.close(fig)


def main() -> None:
    if not DATA_PATH.exists():
        raise FileNotFoundError(
            f"{DATA_PATH} not found. Run the subtotal labeling workflow to create it."
        )
    df = pd.read_csv(DATA_PATH)
    df = _normalize_economy_codes(df)
    df, year_cols = _normalize_year_columns(df)
    if DROP_SUBTOTALS and "is_subtotal" in df.columns:
        df = df[df["is_subtotal"] == False].copy()
    if DROP_TOTAL_ENERGY_ROWS:
        df = _drop_total_energy_rows(df)
    if INCLUDE_AGGREGATE_ECONOMY:
        df = _add_aggregate_economy_total(df, year_cols, AGGREGATE_ECONOMY_LABEL)
    if TRANSFER_FLOW_CODES:
        df = df[df["flows"].isin(TRANSFER_FLOW_CODES)].copy()

    df, last_years = _sum_last_n_years(df, year_cols, LAST_N_YEARS)
    year_label = (
        f"{min(last_years)}-{max(last_years)}" if last_years else "last_n_years"
    )

    long_df = _build_long_df(df, year_label)
    economy_product_totals = _build_economy_product_totals(long_df)
    flags_df = _category_applicability(economy_product_totals, TRANSFER_CATEGORY_TEMPLATES)
    long_df = long_df.merge(flags_df, on="economy", how="left")

    io_totals_df = _category_io_totals(
        economy_product_totals, TRANSFER_CATEGORY_TEMPLATES
    )
    process_config_df = _build_process_config_rows(
        economy_product_totals, TRANSFER_CATEGORY_TEMPLATES, year_label
    )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    long_df.to_csv(OUTPUT_DIR / "transfer_category_flags.csv", index=False)
    io_totals_df.to_csv(OUTPUT_DIR / "transfer_category_io_totals.csv", index=False)
    if not process_config_df.empty:
        process_config_path = OUTPUT_DIR / "transfer_process_config_candidates.csv"
        process_config_df.to_csv(process_config_path, index=False)
        config_dict = build_transfer_process_config(process_config_df)
        config_text = format_transfer_process_config(config_dict)
        config_path = OUTPUT_DIR / "transfer_process_config_candidates.py"
        with config_path.open("w", encoding="utf-8") as handle:
            handle.write("TRANSFER_PROCESS_CONFIG = ")
            handle.write(config_text)
            handle.write("\n")
        print("Generated TRANSFER_PROCESS_CONFIG:")
        print(config_text)
        if AUTO_UPDATE_TRANSFER_PROCESS_CONFIG:
            _update_transfer_plan_config(config_text)
            print("Updated codebase/transfers_workflow.py TRANSFER_PROCESS_CONFIG.")

    plot_paths = []
    plots_enabled = MAKE_PLOTS
    process_plot_paths = []
    process_plots_enabled = MAKE_PLOTS and MAKE_PROCESS_PLOTS
    if MAKE_PLOTS:
        # breakpoint()
        try:
            plot_paths = plot_economy_facets(
                economy_product_totals,
                TRANSFER_CATEGORY_TEMPLATES,
                year_label,
                ECONOMY_PLOT_DIR,
                max_products=MAX_PLOT_PRODUCTS,
                category_cols=CATEGORY_PLOT_COLUMNS,
                dpi=PLOT_DPI,
            )
            stitch_economy_plots(plot_paths, STITCHED_PLOT_PATH, columns=STITCH_PLOT_COLUMNS)
            if MAKE_PROCESS_PLOTS:
                process_plot_paths = plot_economy_process_facets(
                    economy_product_totals,
                    TRANSFER_PROCESS_CONFIG,
                    year_label,
                    PROCESS_ECONOMY_PLOT_DIR,
                    max_products=MAX_PLOT_PRODUCTS,
                    category_cols=CATEGORY_PLOT_COLUMNS,
                    dpi=PLOT_DPI,
                )
                stitch_economy_plots(
                    process_plot_paths,
                    STITCHED_PROCESS_PLOT_PATH,
                    columns=STITCH_PLOT_COLUMNS,
                )
        except ImportError as exc:
            plots_enabled = False
            process_plots_enabled = False
            print(f"[WARN] Plotting skipped (missing dependency): {exc}")

    print("Saved:")
    print(f"- {OUTPUT_DIR / 'transfer_category_flags.csv'}")
    print(f"- {OUTPUT_DIR / 'transfer_category_io_totals.csv'}")
    if not process_config_df.empty:
        print(f"- {OUTPUT_DIR / 'transfer_process_config_candidates.csv'}")
        print(f"- {OUTPUT_DIR / 'transfer_process_config_candidates.py'}")
    if plots_enabled:
        print(f"- {ECONOMY_PLOT_DIR}")
        print(f"- {STITCHED_PLOT_PATH}")
    if process_plots_enabled:
        print(f"- {PROCESS_ECONOMY_PLOT_DIR}")
        print(f"- {STITCHED_PROCESS_PLOT_PATH}")
    print("")
    print("Explanation:")
    print(
        "This run aggregates the last "
        f"{LAST_N_YEARS} years into a single value per economy/product across transfer flows."
    )
    print(
        "Categories are applied using TRANSFER_CATEGORY_TEMPLATES; 'Others' captures products "
        "not listed in the other categories."
    )
    print(
        "If a category has only inputs or only outputs, it is merged with a category that "
        "already has both sides (or with a category that supplies the missing side)."
    )
    print("Use these outputs:")
    print("  - transfer_category_flags.csv for per-economy applicability flags.")
    print("  - transfer_category_io_totals.csv for input/output totals by category.")
    print("  - transfer_process_config_candidates.csv for inputs/outputs per category.")
    print("  - transfer_process_config_candidates.py to paste into TRANSFER_PROCESS_CONFIG.")

#%%
if __name__ == "__main__":
    main()
#%%
