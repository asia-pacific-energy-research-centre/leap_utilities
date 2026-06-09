from __future__ import annotations

import importlib.util
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd

from codebase.functions.leap_series_comparison import ComparisonRunConfig, run_leap_series_comparison


def _write_common_reference_inputs(tmp_path: Path) -> dict[str, Path]:
    esto_path = tmp_path / "esto.csv"
    ninth_path = tmp_path / "ninth.csv"
    subtotal_map_path = tmp_path / "subtotal_mapping.xlsx"
    ninth_to_esto_path = tmp_path / "ninth_to_esto.xlsx"

    esto_df = pd.DataFrame(
        [
            {
                "economy": "01AUS",
                "flows": "09.08.01 Coke ovens",
                "products": "07.01 Motor gasoline",
                "2022": 100.0,
            }
        ]
    )
    esto_df.to_csv(esto_path, index=False)

    ninth_df = pd.DataFrame(
        [
            {
                "scenarios": "reference",
                "economy": "01_AUS",
                "sectors": "09_08_coal_transformation",
                "sub1sectors": "x",
                "sub2sectors": "x",
                "sub3sectors": "x",
                "sub4sectors": "x",
                "fuels": "07_petroleum_products",
                "subfuels": "07_01_motor_gasoline",
                "subtotal_results": False,
                "2023": 110.0,
                "2024": 120.0,
            }
        ]
    )
    ninth_df.to_csv(ninth_path, index=False)

    subtotal_map_df = pd.DataFrame(
        [
            {
                "flow": "09.08.01 Coke ovens",
                "product": "07.01 Motor gasoline",
                "is_subtotal": False,
            }
        ]
    )
    subtotal_map_df.to_excel(subtotal_map_path, index=False)

    ninth_to_esto_df = pd.DataFrame(
        [
            {
                "9th_sector": "09_08_coal_transformation",
                "9th_fuel": "07_01_motor_gasoline",
                "esto_flow": "09.08.01 Coke ovens",
                "esto_product": "07.01 Motor gasoline",
            }
        ]
    )
    ninth_to_esto_df.to_excel(ninth_to_esto_path, index=False)

    return {
        "esto_path": esto_path,
        "ninth_path": ninth_path,
        "subtotal_map_path": subtotal_map_path,
        "ninth_to_esto_path": ninth_to_esto_path,
    }


def _write_mapping_csv(tmp_path: Path, rows: list[dict[str, object]]) -> Path:
    mapping_path = tmp_path / "mapping.csv"
    mapping_df = pd.DataFrame(rows)
    mapping_df.to_csv(mapping_path, index=False)
    return mapping_path


def _build_config(
    tmp_path: Path,
    leap_file: Path,
    mapping_csv: Path,
    reference_inputs: dict[str, Path],
    output_name: str,
) -> ComparisonRunConfig:
    return ComparisonRunConfig(
        leap_file=leap_file,
        leap_sheet="LEAP",
        mapping_csv=mapping_csv,
        economy="01_AUS",
        scenario="Reference",
        region="United States",
        esto_data_path=reference_inputs["esto_path"],
        ninth_data_path=reference_inputs["ninth_path"],
        subtotal_mapping_path=reference_inputs["subtotal_map_path"],
        ninth_to_esto_mapping_path=reference_inputs["ninth_to_esto_path"],
        base_year=2022,
        projection_start_year=2023,
        projection_end_year=2024,
        output_dir=tmp_path / output_name,
    )


class TestLeapSeriesComparison(unittest.TestCase):
    def test_wide_year_input_parsing(self) -> None:
        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            reference_inputs = _write_common_reference_inputs(tmp_path)

            leap_wide = pd.DataFrame(
                [
                    {
                        "Branch Path": (
                            "Transformation\\Coke ovens\\Processes\\Coke ovens"
                            "\\Output Fuels\\Motor gasoline"
                        ),
                        "Variable": "Output Energy",
                        "Scenario": "Reference",
                        "Region": "United States",
                        "2022": 90.0,
                        "2023": 111.0,
                        "2024": 119.0,
                    }
                ]
            )
            leap_path = tmp_path / "leap_wide.csv"
            leap_wide.to_csv(leap_path, index=False)

            mapping_path = _write_mapping_csv(
                tmp_path,
                [
                    {
                        "series_id": "series_1",
                        "sector_tag": "transformation",
                        "leap_variable": "Output Energy",
                        "leap_branch_contains": "Transformation\\Coke ovens\\Processes",
                        "leap_fuel_label": "Motor gasoline",
                        "esto_flow": "09.08.01 Coke ovens",
                        "esto_product": "07.01 Motor gasoline",
                        "ninth_sector_expected": "09_08_coal_transformation",
                        "ninth_fuel_expected": "07_01_motor_gasoline",
                        "active": True,
                        "notes": "",
                    }
                ],
            )

            config = _build_config(
                tmp_path, leap_path, mapping_path, reference_inputs, "out_wide"
            )
            artifacts = run_leap_series_comparison(config)
            long_df = pd.read_csv(artifacts.comparison_long_csv)

            years = set(long_df["year"].tolist())
            self.assertEqual(years, {2022, 2023, 2024})
            value_2022 = float(long_df.loc[long_df["year"] == 2022, "leap_value"].iloc[0])
            self.assertEqual(value_2022, 90.0)

    def test_expression_input_parsing(self) -> None:
        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            reference_inputs = _write_common_reference_inputs(tmp_path)

            leap_expr = pd.DataFrame(
                [
                    {
                        "Branch Path": (
                            "Transformation\\Coke ovens\\Processes\\Coke ovens"
                            "\\Output Fuels\\Motor gasoline"
                        ),
                        "Variable": "Output Energy",
                        "Scenario": "Reference",
                        "Region": "United States",
                        "Expression": "Data(2022,90, 2023,95, 2024,100)",
                    }
                ]
            )
            leap_path = tmp_path / "leap_expression.csv"
            leap_expr.to_csv(leap_path, index=False)

            mapping_path = _write_mapping_csv(
                tmp_path,
                [
                    {
                        "series_id": "series_expr",
                        "sector_tag": "transformation",
                        "leap_variable": "Output Energy",
                        "leap_branch_contains": "Transformation\\Coke ovens\\Processes",
                        "leap_fuel_label": "Motor gasoline",
                        "esto_flow": "09.08.01 Coke ovens",
                        "esto_product": "07.01 Motor gasoline",
                        "ninth_sector_expected": "09_08_coal_transformation",
                        "ninth_fuel_expected": "07_01_motor_gasoline",
                        "active": True,
                        "notes": "",
                    }
                ],
            )

            config = _build_config(
                tmp_path, leap_path, mapping_path, reference_inputs, "out_expr"
            )
            artifacts = run_leap_series_comparison(config)
            long_df = pd.read_csv(artifacts.comparison_long_csv)
            value_2023 = float(long_df.loc[long_df["year"] == 2023, "leap_value"].iloc[0])
            self.assertEqual(value_2023, 95.0)

    def test_mapping_match_and_aggregation(self) -> None:
        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            reference_inputs = _write_common_reference_inputs(tmp_path)

            leap_wide = pd.DataFrame(
                [
                    {
                        "Branch Path": (
                            "Transformation\\Coke ovens\\Processes\\Coke ovens"
                            "\\Output Fuels\\Motor gasoline"
                        ),
                        "Variable": "Output Energy",
                        "Scenario": "Reference",
                        "Region": "United States",
                        "2022": 40.0,
                        "2023": 50.0,
                        "2024": 60.0,
                    },
                    {
                        "Branch Path": (
                            "Transformation\\Coke ovens\\Processes\\Coke ovens"
                            "\\Output Fuels\\Motor gasoline"
                        ),
                        "Variable": "Output Energy",
                        "Scenario": "Reference",
                        "Region": "United States",
                        "2022": 10.0,
                        "2023": 20.0,
                        "2024": 30.0,
                    },
                ]
            )
            leap_path = tmp_path / "leap_agg.csv"
            leap_wide.to_csv(leap_path, index=False)

            mapping_path = _write_mapping_csv(
                tmp_path,
                [
                    {
                        "series_id": "series_agg",
                        "sector_tag": "transformation",
                        "leap_variable": "Output Energy",
                        "leap_branch_contains": "Transformation\\Coke ovens\\Processes",
                        "leap_fuel_label": "Motor gasoline",
                        "esto_flow": "09.08.01 Coke ovens",
                        "esto_product": "07.01 Motor gasoline",
                        "ninth_sector_expected": "09_08_coal_transformation",
                        "ninth_fuel_expected": "07_01_motor_gasoline",
                        "active": True,
                        "notes": "",
                    }
                ],
            )

            config = _build_config(
                tmp_path, leap_path, mapping_path, reference_inputs, "out_agg"
            )
            artifacts = run_leap_series_comparison(config)
            long_df = pd.read_csv(artifacts.comparison_long_csv)
            value_2022 = float(long_df.loc[long_df["year"] == 2022, "leap_value"].iloc[0])
            value_2023 = float(long_df.loc[long_df["year"] == 2023, "leap_value"].iloc[0])
            self.assertEqual(value_2022, 50.0)
            self.assertEqual(value_2023, 70.0)

    def test_reference_series_stitch(self) -> None:
        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            reference_inputs = _write_common_reference_inputs(tmp_path)

            leap_wide = pd.DataFrame(
                [
                    {
                        "Branch Path": (
                            "Transformation\\Coke ovens\\Processes\\Coke ovens"
                            "\\Output Fuels\\Motor gasoline"
                        ),
                        "Variable": "Output Energy",
                        "Scenario": "Reference",
                        "Region": "United States",
                        "2022": 100.0,
                        "2023": 100.0,
                        "2024": 100.0,
                    }
                ]
            )
            leap_path = tmp_path / "leap_stitch.csv"
            leap_wide.to_csv(leap_path, index=False)

            mapping_path = _write_mapping_csv(
                tmp_path,
                [
                    {
                        "series_id": "series_stitch",
                        "sector_tag": "transformation",
                        "leap_variable": "Output Energy",
                        "leap_branch_contains": "Transformation\\Coke ovens\\Processes",
                        "leap_fuel_label": "Motor gasoline",
                        "esto_flow": "09.08.01 Coke ovens",
                        "esto_product": "07.01 Motor gasoline",
                        "ninth_sector_expected": "09_08_coal_transformation",
                        "ninth_fuel_expected": "07_01_motor_gasoline",
                        "active": True,
                        "notes": "",
                    }
                ],
            )

            config = _build_config(
                tmp_path, leap_path, mapping_path, reference_inputs, "out_stitch"
            )
            artifacts = run_leap_series_comparison(config)
            long_df = pd.read_csv(artifacts.comparison_long_csv)

            source_2022 = long_df.loc[long_df["year"] == 2022, "reference_source"].iloc[0]
            source_2023 = long_df.loc[long_df["year"] == 2023, "reference_source"].iloc[0]
            ref_2022 = float(long_df.loc[long_df["year"] == 2022, "reference_value"].iloc[0])
            ref_2023 = float(long_df.loc[long_df["year"] == 2023, "reference_value"].iloc[0])

            self.assertEqual(source_2022, "esto_base_year")
            self.assertEqual(source_2023, "ninth_projection_allocated")
            self.assertEqual(ref_2022, 100.0)
            self.assertEqual(ref_2023, 110.0)

    def test_missing_mapping_and_missing_reference_flags(self) -> None:
        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            reference_inputs = _write_common_reference_inputs(tmp_path)

            leap_wide = pd.DataFrame(
                [
                    {
                        "Branch Path": (
                            "Transformation\\Coke ovens\\Processes\\Coke ovens"
                            "\\Output Fuels\\Motor gasoline"
                        ),
                        "Variable": "Output Energy",
                        "Scenario": "Reference",
                        "Region": "United States",
                        "2022": 90.0,
                        "2023": 111.0,
                        "2024": 119.0,
                    }
                ]
            )
            leap_path = tmp_path / "leap_missing.csv"
            leap_wide.to_csv(leap_path, index=False)

            mapping_path = _write_mapping_csv(
                tmp_path,
                [
                    {
                        "series_id": "series_valid",
                        "sector_tag": "transformation",
                        "leap_variable": "Output Energy",
                        "leap_branch_contains": "Transformation\\Coke ovens\\Processes",
                        "leap_fuel_label": "Motor gasoline",
                        "esto_flow": "09.08.01 Coke ovens",
                        "esto_product": "07.01 Motor gasoline",
                        "ninth_sector_expected": "09_08_coal_transformation",
                        "ninth_fuel_expected": "07_01_motor_gasoline",
                        "active": True,
                        "notes": "",
                    },
                    {
                        "series_id": "series_missing",
                        "sector_tag": "transformation",
                        "leap_variable": "Nonexistent Variable",
                        "leap_branch_contains": "missing_branch_token",
                        "leap_fuel_label": "Missing Fuel",
                        "esto_flow": "99.99 Missing flow",
                        "esto_product": "99.99 Missing product",
                        "ninth_sector_expected": "99_missing_sector",
                        "ninth_fuel_expected": "99_missing_fuel",
                        "active": True,
                        "notes": "",
                    },
                ],
            )

            config = _build_config(
                tmp_path, leap_path, mapping_path, reference_inputs, "out_missing"
            )
            artifacts = run_leap_series_comparison(config)
            status_df = pd.read_csv(artifacts.mapping_status_csv)

            missing_row = status_df[status_df["series_id"] == "series_missing"].iloc[0]
            self.assertFalse(bool(missing_row["has_leap_match"]))
            self.assertFalse(bool(missing_row["has_base_year_reference"]))
            self.assertGreater(int(missing_row["missing_reference_years_count"]), 0)

    @unittest.skipIf(
        importlib.util.find_spec("matplotlib") is None,
        "matplotlib is not installed in this environment.",
    )
    def test_chart_generation_smoke(self) -> None:
        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            reference_inputs = _write_common_reference_inputs(tmp_path)

            leap_wide = pd.DataFrame(
                [
                    {
                        "Branch Path": (
                            "Transformation\\Coke ovens\\Processes\\Coke ovens"
                            "\\Output Fuels\\Motor gasoline"
                        ),
                        "Variable": "Output Energy",
                        "Scenario": "Reference",
                        "Region": "United States",
                        "2022": 90.0,
                        "2023": 111.0,
                        "2024": 119.0,
                    }
                ]
            )
            leap_path = tmp_path / "leap_chart.csv"
            leap_wide.to_csv(leap_path, index=False)

            mapping_path = _write_mapping_csv(
                tmp_path,
                [
                    {
                        "series_id": "series_chart",
                        "sector_tag": "transformation",
                        "leap_variable": "Output Energy",
                        "leap_branch_contains": "Transformation\\Coke ovens\\Processes",
                        "leap_fuel_label": "Motor gasoline",
                        "esto_flow": "09.08.01 Coke ovens",
                        "esto_product": "07.01 Motor gasoline",
                        "ninth_sector_expected": "09_08_coal_transformation",
                        "ninth_fuel_expected": "07_01_motor_gasoline",
                        "active": True,
                        "notes": "",
                    }
                ],
            )

            config = _build_config(
                tmp_path, leap_path, mapping_path, reference_inputs, "out_chart"
            )
            artifacts = run_leap_series_comparison(config)

            png_files = list(Path(artifacts.charts_dir).glob("*.png"))
            self.assertGreaterEqual(len(png_files), 1)


if __name__ == "__main__":
    unittest.main()
