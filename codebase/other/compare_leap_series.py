#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from pathlib import Path

# Allow running the script directly without package install.
REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.functions.leap_series_comparison import (
    ComparisonRunConfig,
    run_leap_series_comparison,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Compare LEAP series against ESTO base-year values and allocated 9th projections."
        )
    )
    parser.add_argument("--leap-file", required=True, help="Path to LEAP export CSV/XLSX.")
    parser.add_argument(
        "--leap-sheet",
        default="LEAP",
        help="Sheet name for XLSX LEAP export files (default: LEAP).",
    )
    parser.add_argument(
        "--mapping-csv",
        required=True,
        help="Path to explicit mapping CSV (series -> ESTO flow/product).",
    )
    parser.add_argument("--economy", required=True, help="Economy code, e.g. 20_USA.")
    parser.add_argument("--scenario", required=True, help="Scenario label to filter in LEAP file.")
    parser.add_argument("--region", required=True, help="Region label to filter in LEAP file.")

    parser.add_argument(
        "--esto-data-path",
        default="data/00APEC_2024_low.csv",
        help="Path to ESTO base-year input table.",
    )
    parser.add_argument(
        "--ninth-data-path",
        default="data/merged_file_energy_ALL_20250814_pre_trump.csv",
        help="Path to 9th projection table.",
    )
    parser.add_argument(
        "--subtotal-mapping-path",
        default="config/ESTO_subtotal_mapping.xlsx",
        help="Path to ESTO subtotal mapping workbook.",
    )
    parser.add_argument(
        "--ninth-to-esto-mapping-path",
        default="config/ninth_pairs_to_esto_pairs.xlsx",
        help="Path to 9th<->ESTO mapping file used for projection allocation.",
    )

    parser.add_argument("--base-year", type=int, default=2022, help="Base year (default: 2022).")
    parser.add_argument(
        "--projection-start-year",
        type=int,
        default=2023,
        help="Projection start year (default: 2023).",
    )
    parser.add_argument(
        "--projection-end-year",
        type=int,
        default=2060,
        help="Projection end year (default: 2060).",
    )
    parser.add_argument(
        "--output-dir",
        default="outputs/series_comparison",
        help="Directory for CSV outputs and PNG charts.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    config = ComparisonRunConfig(
        leap_file=args.leap_file,
        leap_sheet=args.leap_sheet,
        mapping_csv=args.mapping_csv,
        economy=args.economy,
        scenario=args.scenario,
        region=args.region,
        esto_data_path=args.esto_data_path,
        ninth_data_path=args.ninth_data_path,
        subtotal_mapping_path=args.subtotal_mapping_path,
        ninth_to_esto_mapping_path=args.ninth_to_esto_mapping_path,
        base_year=args.base_year,
        projection_start_year=args.projection_start_year,
        projection_end_year=args.projection_end_year,
        output_dir=args.output_dir,
    )

    artifacts = run_leap_series_comparison(config)
    print("[OK] Series comparison complete.")
    print(f"- comparison_long_csv: {artifacts.comparison_long_csv}")
    print(f"- comparison_wide_csv: {artifacts.comparison_wide_csv}")
    print(f"- comparison_summary_csv: {artifacts.comparison_summary_csv}")
    print(f"- mapping_status_csv: {artifacts.mapping_status_csv}")
    print(f"- unmatched_leap_rows_csv: {artifacts.unmatched_leap_rows_csv}")
    print(f"- charts_dir: {artifacts.charts_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
