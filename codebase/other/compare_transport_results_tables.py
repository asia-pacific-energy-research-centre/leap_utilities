#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

# Allow running the script directly without package install.
REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.functions.leap_series_comparison import (  # noqa: E402
    TransportResultsComparisonConfig,
    run_transport_results_table_comparison,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Compare LEAP transport results tables against ESTO base year and 9th projections."
        )
    )
    parser.add_argument(
        "--leap-results-file",
        required=True,
        help="Path to LEAP transport results workbook (multi-sheet results tables).",
    )
    parser.add_argument("--economy", required=True, help="Economy code, e.g. 20_USA.")
    parser.add_argument("--scenario", required=True, help="Scenario label to match from A2 metadata.")
    parser.add_argument("--region", required=True, help="Region label to match from A2 metadata.")
    parser.add_argument(
        "--branch-sector-mapping-csv",
        default="config/leap_transport_branch_to_ninth_sector_map.csv",
        help="CSV mapping LEAP branch paths to 9th sector codes.",
    )
    parser.add_argument(
        "--fuel-aliases-csv",
        default="config/leap_transport_fuel_aliases.csv",
        help="CSV mapping LEAP fuel labels to codebook names/overrides.",
    )
    parser.add_argument(
        "--code-to-name-path",
        default="config/sector_fuel_codes_to_names.xlsx",
        help="Path to sector/fuel code-to-name workbook.",
    )
    parser.add_argument(
        "--code-to-name-sheet",
        default="code_to_name",
        help="Sheet name in code-to-name workbook (default: code_to_name).",
    )
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
        help="Path to 9th<->ESTO mapping used for flow/product relationships.",
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
        default=2061,
        help="Projection end year (default: 2061).",
    )
    parser.add_argument(
        "--share-year-offset",
        type=int,
        default=1,
        help="Offset from base year for parent->child share allocation (default: 1 => 2023).",
    )
    parser.add_argument(
        "--ninth-scenario",
        default="reference",
        help="9th scenario to filter (default: reference).",
    )
    parser.add_argument(
        "--output-dir",
        default="outputs/transport_results_series_comparison",
        help="Directory for CSV outputs and PNG charts.",
    )
    parser.add_argument(
        "--skip-filename-validation",
        action="store_true",
        help=(
            "Skip validation that the LEAP workbook filename contains both economy and scenario "
            "tokens (not recommended)."
        ),
    )
    return parser


def _normalize_token(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", value.lower())


def _validate_results_filename(path: Path, economy: str, scenario: str) -> None:
    stem = path.stem
    normalized_stem = _normalize_token(stem)
    normalized_economy = _normalize_token(economy)
    normalized_scenario = _normalize_token(scenario)

    economy_ok = normalized_economy in normalized_stem
    scenario_ok = normalized_scenario in normalized_stem
    if economy_ok and scenario_ok:
        return

    suggested_name = f"transport_results_{economy}_{scenario}{path.suffix}"
    raise ValueError(
        "LEAP results filename must include both economy and scenario tokens to avoid mismatched runs. "
        f"Got '{path.name}', expected tokens economy='{economy}' and scenario='{scenario}'. "
        f"Rename the file (for example): '{suggested_name}', or pass --skip-filename-validation "
        "to bypass this check."
    )


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    leap_results_file = Path(args.leap_results_file)
    if not args.skip_filename_validation:
        _validate_results_filename(leap_results_file, args.economy, args.scenario)

    config = TransportResultsComparisonConfig(
        leap_results_file=leap_results_file,
        economy=args.economy,
        scenario=args.scenario,
        region=args.region,
        branch_sector_mapping_csv=args.branch_sector_mapping_csv,
        fuel_aliases_csv=args.fuel_aliases_csv,
        code_to_name_path=args.code_to_name_path,
        code_to_name_sheet=args.code_to_name_sheet,
        esto_data_path=args.esto_data_path,
        ninth_data_path=args.ninth_data_path,
        subtotal_mapping_path=args.subtotal_mapping_path,
        ninth_to_esto_mapping_path=args.ninth_to_esto_mapping_path,
        base_year=args.base_year,
        projection_start_year=args.projection_start_year,
        projection_end_year=args.projection_end_year,
        share_year_offset=args.share_year_offset,
        ninth_scenario=args.ninth_scenario,
        output_dir=args.output_dir,
    )

    artifacts = run_transport_results_table_comparison(config)
    print("[OK] Transport results-table comparison complete.")
    print(f"- comparison_long_csv: {artifacts.comparison_long_csv}")
    print(f"- comparison_wide_csv: {artifacts.comparison_wide_csv}")
    print(f"- comparison_summary_csv: {artifacts.comparison_summary_csv}")
    print(f"- mapping_status_csv: {artifacts.mapping_status_csv}")
    print(f"- unmatched_leap_rows_csv: {artifacts.unmatched_leap_rows_csv}")
    print(f"- charts_dir: {artifacts.charts_dir}")
    print("- sheet_inventory_csv: " + str(Path(args.output_dir) / "sheet_inventory.csv"))
    print("- fuel_mapping_status_csv: " + str(Path(args.output_dir) / "fuel_mapping_status.csv"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
