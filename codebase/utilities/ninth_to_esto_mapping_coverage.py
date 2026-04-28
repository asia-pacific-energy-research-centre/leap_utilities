from __future__ import annotations

from pathlib import Path
from typing import Sequence

import pandas as pd

from codebase.utilities.master_config import config_table_exists, read_config_table
from codebase.utilities.workflow_outputs import build_workflow_output_layout, write_output_manifest

from codebase.functions.ninth_projection_mapping import (
    add_ninth_pair_columns,
    filter_ninth_projection_rows,
)
from codebase.scrapbook.utilities import load_augmented_reference_tables


REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_SUBTOTAL_MAPPING_PATH = REPO_ROOT / "config" / "ESTO_subtotal_mapping.xlsx"
DEFAULT_SYNTHETIC_RULES_PATH = REPO_ROOT / "config" / "synthetic_reference_rows.csv"
DEFAULT_REFERENCE_CACHE_DIR = REPO_ROOT / "data" / ".cache" / "mapping_coverage_reference_tables"


def _resolve(path: str | Path) -> Path:
    return path if isinstance(path, Path) else Path(str(path))


def _read_table(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xls", ".xlsm"}:
        return read_config_table(path, dtype=str).fillna("")
    return read_config_table(path, dtype=str).fillna("")


def _coerce_bool_flag(series: pd.Series) -> pd.Series:
    return (
        series.fillna(False)
        .astype(str)
        .str.strip()
        .str.lower()
        .isin({"1", "true", "yes", "y", "t"})
    )


def load_mapping_pairs(mapping_path: str | Path) -> pd.DataFrame:
    path = _resolve(mapping_path)
    mapping_df = _read_table(path)
    required = ["9th_sector", "9th_fuel", "esto_flow", "esto_product"]
    missing = [col for col in required if col not in mapping_df.columns]
    if missing:
        raise ValueError(f"Mapping file missing required columns: {missing}")
    mapping_df = mapping_df[required].copy()
    for col in required:
        mapping_df[col] = mapping_df[col].astype(str).str.strip()
    return mapping_df.drop_duplicates().reset_index(drop=True)


def build_nonzero_esto_pairs(
    esto_df: pd.DataFrame,
    *,
    base_year: int,
) -> pd.DataFrame:
    year_col = str(base_year) if str(base_year) in esto_df.columns else base_year
    if year_col not in esto_df.columns:
        raise ValueError(f"ESTO data missing base year column {base_year}.")

    working = esto_df.copy()
    for subtotal_col in [
        "is_subtotal",
        "subtotal_2022_and_before",
        "subtotal_2023_and_after",
        "subtotal_layout",
        "subtotal_results",
    ]:
        if subtotal_col in working.columns:
            working = working[~_coerce_bool_flag(working[subtotal_col])].copy()

    working["flows"] = working["flows"].astype(str).str.strip()
    working["products"] = working["products"].astype(str).str.strip()
    working[year_col] = pd.to_numeric(working[year_col], errors="coerce").fillna(0.0)
    working = working[working["flows"].ne("") & working["products"].ne("")]
    working = working[working[year_col].abs() > 0]

    return (
        working[["flows", "products"]]
        .drop_duplicates()
        .rename(columns={"flows": "esto_flow", "products": "esto_product"})
        .sort_values(["esto_flow", "esto_product"])
        .reset_index(drop=True)
    )


def build_nonzero_ninth_pairs(
    ninth_df: pd.DataFrame,
    *,
    scenario: str,
    projection_years: Sequence[int],
) -> pd.DataFrame:
    working = filter_ninth_projection_rows(ninth_df, scenario=scenario)
    if working.empty:
        return pd.DataFrame(columns=["9th_sector", "9th_fuel"])

    for subtotal_col in ["subtotal_layout", "subtotal_results"]:
        if subtotal_col in working.columns:
            working = working[~_coerce_bool_flag(working[subtotal_col])].copy()

    working = add_ninth_pair_columns(working)
    year_cols: list[int | str] = []
    for year in projection_years:
        if year in working.columns:
            year_cols.append(year)
        elif str(year) in working.columns:
            year_cols.append(str(year))
    if not year_cols:
        raise ValueError("No requested projection year columns were found in 9th data.")

    numeric = working[year_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    working = working[
        working["9th_sector"].astype(str).str.strip().ne("")
        & working["9th_fuel"].astype(str).str.strip().ne("")
        & numeric.abs().sum(axis=1).gt(0)
    ].copy()

    return (
        working[["9th_sector", "9th_fuel"]]
        .drop_duplicates()
        .sort_values(["9th_sector", "9th_fuel"])
        .reset_index(drop=True)
    )


def run_mapping_coverage_check(
    *,
    mapping_path: str | Path,
    esto_data_path: str | Path,
    ninth_data_path: str | Path,
    output_dir: str | Path,
    base_year: int = 2022,
    projection_years: Sequence[int] = tuple(range(2023, 2071)),
    scenario: str = "reference",
    subtotal_mapping_path: str | Path | None = DEFAULT_SUBTOTAL_MAPPING_PATH,
    synthetic_rules_path: str | Path | None = DEFAULT_SYNTHETIC_RULES_PATH,
    cache_dir: str | Path | None = DEFAULT_REFERENCE_CACHE_DIR,
) -> dict[str, object]:
    mapping_path = _resolve(mapping_path)
    esto_data_path = _resolve(esto_data_path)
    ninth_data_path = _resolve(ninth_data_path)
    output_dir = _resolve(output_dir)
    subtotal_mapping_path = _resolve(subtotal_mapping_path) if subtotal_mapping_path else None
    synthetic_rules_path = _resolve(synthetic_rules_path) if synthetic_rules_path else None
    cache_dir = _resolve(cache_dir) if cache_dir else None

    if not config_table_exists(mapping_path):
        raise FileNotFoundError(f"Missing coverage-check input: {mapping_path}")
    for path in [esto_data_path, ninth_data_path]:
        if not path.exists():
            raise FileNotFoundError(f"Missing coverage-check input: {path}")

    layout = build_workflow_output_layout(output_dir)

    mapping_df = load_mapping_pairs(mapping_path)
    esto_df, ninth_df = load_augmented_reference_tables(
        esto_path=esto_data_path,
        ninth_path=ninth_data_path,
        subtotal_mapping_path=subtotal_mapping_path,
        synthetic_rules_path=synthetic_rules_path,
        cache_dir=cache_dir or DEFAULT_REFERENCE_CACHE_DIR,
        apply_esto_subtotal_map=True,
        filter_esto_subtotals_flag=True,
        filter_ninth_subtotals_flag=False,
    )

    nonzero_esto_pairs = build_nonzero_esto_pairs(esto_df, base_year=base_year)
    nonzero_ninth_pairs = build_nonzero_ninth_pairs(
        ninth_df,
        scenario=scenario,
        projection_years=projection_years,
    )

    mapped_esto_pairs = (
        mapping_df[["esto_flow", "esto_product"]]
        .drop_duplicates()
        .sort_values(["esto_flow", "esto_product"])
        .reset_index(drop=True)
    )
    mapped_ninth_pairs = (
        mapping_df[["9th_sector", "9th_fuel"]]
        .drop_duplicates()
        .sort_values(["9th_sector", "9th_fuel"])
        .reset_index(drop=True)
    )

    missing_esto_pairs = nonzero_esto_pairs.merge(
        mapped_esto_pairs,
        on=["esto_flow", "esto_product"],
        how="left",
        indicator=True,
    )
    missing_esto_pairs = (
        missing_esto_pairs[missing_esto_pairs["_merge"] == "left_only"]
        .drop(columns="_merge")
        .reset_index(drop=True)
    )

    missing_ninth_pairs = nonzero_ninth_pairs.merge(
        mapped_ninth_pairs,
        on=["9th_sector", "9th_fuel"],
        how="left",
        indicator=True,
    )
    missing_ninth_pairs = (
        missing_ninth_pairs[missing_ninth_pairs["_merge"] == "left_only"]
        .drop(columns="_merge")
        .reset_index(drop=True)
    )

    summary_df = pd.DataFrame(
        [
            {
                "mapping_path": str(mapping_path),
                "esto_data_path": str(esto_data_path),
                "ninth_data_path": str(ninth_data_path),
                "base_year": int(base_year),
                "projection_year_start": int(min(projection_years)),
                "projection_year_end": int(max(projection_years)),
                "scenario": str(scenario),
                "unique_mapping_esto_pairs": int(len(mapped_esto_pairs)),
                "unique_mapping_ninth_pairs": int(len(mapped_ninth_pairs)),
                "unique_nonzero_esto_pairs": int(len(nonzero_esto_pairs)),
                "unique_nonzero_ninth_pairs": int(len(nonzero_ninth_pairs)),
                "missing_esto_pairs": int(len(missing_esto_pairs)),
                "missing_ninth_pairs": int(len(missing_ninth_pairs)),
                "all_nonzero_pairs_mapped": bool(
                    missing_esto_pairs.empty and missing_ninth_pairs.empty
                ),
            }
        ]
    )

    summary_path = layout.root / "mapping_coverage_summary.csv"
    missing_esto_path = layout.checks / "missing_esto_flow_product_pairs.csv"
    missing_ninth_path = layout.checks / "missing_9th_sector_fuel_pairs.csv"

    summary_df.to_csv(summary_path, index=False)
    missing_esto_pairs.to_csv(missing_esto_path, index=False)
    missing_ninth_pairs.to_csv(missing_ninth_path, index=False)
    manifest_path = write_output_manifest(
        out_dir=layout.root,
        primary_outputs={"summary_csv": str(summary_path)},
        supporting_outputs={
            "missing_esto_csv": str(missing_esto_path),
            "missing_ninth_csv": str(missing_ninth_path),
        },
        primary_output_descriptions={
            "summary_csv": "Summary of mapped versus nonzero ESTO and 9th pair coverage.",
        },
        supporting_output_descriptions={
            "missing_esto_csv": "Nonzero ESTO flow/product pairs missing from the mapping workbook.",
            "missing_ninth_csv": "Nonzero 9th sector/fuel pairs missing from the mapping workbook.",
        },
        notes=[
            "Coverage summary stays at the workflow root.",
            "Missing-pair detail files are grouped under supporting_files/.",
        ],
    )

    return {
        "summary": summary_df.iloc[0].to_dict(),
        "summary_csv": str(summary_path),
        "missing_esto_csv": str(missing_esto_path),
        "missing_ninth_csv": str(missing_ninth_path),
        "output_manifest_json": str(manifest_path),
    }
