from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

import pandas as pd

from codebase.functions.industry_fuel_remap import remap_industry_export_fuels
from codebase.functions.buildings_fuel_remap import remap_buildings_export_fuels
from codebase.functions.leap_excel_io import read_export_sheet, write_export_sheet
from codebase.functions.leap_series_adapter import (
    collect_available_years,
    detect_series_format,
    extract_row_series,
    scale_row_series,
)


class TestSeriesAdapterAndRemaps(unittest.TestCase):
    def test_series_adapter_detects_expression_and_year_columns(self) -> None:
        df_expr = pd.DataFrame([{"Expression": "Data(2022, 1, 2023, 2)"}])
        self.assertEqual(detect_series_format(df_expr), "expression")
        self.assertEqual(collect_available_years(df_expr, "expression"), [2022, 2023])

        df_year = pd.DataFrame([{"2022": 1.0, "2023": 2.0}])
        self.assertEqual(detect_series_format(df_year), "year_columns")
        self.assertEqual(collect_available_years(df_year, "year_columns"), [2022, 2023])

        mixed = pd.DataFrame([{"Expression": "Data(2022,1)", "2022": 1.0}])
        with self.assertRaisesRegex(ValueError, "mixed series format"):
            detect_series_format(mixed)

    def test_series_adapter_extract_and_scale_year_columns(self) -> None:
        row = pd.Series({"2022": 10.0, "2023": 20.0})
        series = extract_row_series(row, "year_columns", year_cols=[2022, 2023])
        self.assertEqual(series, {2022: 10.0, 2023: 20.0})
        scaled = scale_row_series(
            row,
            scale_by={2022: 0.5, 2023: 0.25},
            base_year=2022,
            series_format="year_columns",
            year_cols=[2022, 2023],
        )
        self.assertEqual(float(scaled.get("2022", scaled.get(2022))), 5.0)
        self.assertEqual(float(scaled.get("2023", scaled.get(2023))), 5.0)

    def _write_export_workbook(
        self,
        path: Path,
        data: pd.DataFrame,
        columns: list[object],
    ) -> None:
        header_rows = pd.DataFrame(
            [
                {col: "" for col in columns},
                {col: pd.NA for col in columns},
            ]
        )
        write_export_sheet(
            path=path,
            sheet_name="Export",
            header_rows=header_rows,
            columns=columns,
            data=data,
        )

    def _write_industry_support_files(self, tmp: Path) -> tuple[Path, Path, Path, Path]:
        mapping_path = tmp / "industry_mapping.csv"
        mapping_df = pd.DataFrame(
            [
                {
                    "industry_fuel": "Gasoline",
                    "canonical_industry_fuel": "Gasoline",
                    "mapping_mode": "direct",
                    "target_fuels": "07.01 Motor gasoline",
                    "notes": "",
                }
            ]
        )
        mapping_df.to_csv(mapping_path, index=False)

        esto_path = tmp / "esto.csv"
        esto_df = pd.DataFrame(
            [
                {
                    "economy": "20USA",
                    "flows": "14 Industry sector",
                    "products": "07.01 Motor gasoline",
                    "2022": 100.0,
                    "2023": 100.0,
                }
            ]
        )
        esto_df.to_csv(esto_path, index=False)

        ninth_path = tmp / "ninth.csv"
        ninth_df = pd.DataFrame(
            [
                {
                    "economy": "20USA",
                    "scenarios": "reference",
                    "sectors": "14_industry_sector",
                    "subfuels": "16_x_hydrogen",
                    "2022": 0.0,
                    "2023": 0.0,
                }
            ]
        )
        ninth_df.to_csv(ninth_path, index=False)
        subtotal_mapping = Path(__file__).resolve().parents[1] / "config" / "ESTO_subtotal_mapping.xlsx"
        return mapping_path, esto_path, ninth_path, subtotal_mapping

    def test_industry_remap_supports_expression_and_year_columns(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            mapping_path, esto_path, ninth_path, subtotal_mapping = self._write_industry_support_files(tmp)

            common_cols = [
                "BranchID",
                "VariableID",
                "ScenarioID",
                "RegionID",
                "Branch Path",
                "Variable",
                "Scenario",
                "Region",
                "Scale",
                "Units",
                "Per...",
                "Level 1",
                "Level 2",
                "Level 3",
                "Level 4",
                "Level 5",
            ]

            expr_cols = [*common_cols[:11], "Expression", *common_cols[11:]]
            expr_data = pd.DataFrame(
                [
                    {
                        "BranchID": 1,
                        "VariableID": 1,
                        "ScenarioID": 1,
                        "RegionID": 1,
                        "Branch Path": r"Demand\Industry\Mining and quarrying\Gasoline",
                        "Variable": "Activity Level",
                        "Scenario": "Reference",
                        "Region": "United States",
                        "Scale": "",
                        "Units": "Petajoule",
                        "Per...": "",
                        "Expression": "Data(2022,10,2023,20)",
                        "Level 1": "Demand",
                        "Level 2": "Industry",
                        "Level 3": "Mining and quarrying",
                        "Level 4": "Gasoline",
                        "Level 5": "",
                    }
                ],
                columns=expr_cols,
            )
            expr_input = tmp / "industry_expr.xlsx"
            expr_output = tmp / "industry_expr_out.xlsx"
            self._write_export_workbook(expr_input, expr_data, expr_cols)

            remap_industry_export_fuels(
                input_path=expr_input,
                output_path=expr_output,
                mapping_csv_path=mapping_path,
                esto_data_path=esto_path,
                ninth_data_path=ninth_path,
                subtotal_mapping_path=subtotal_mapping,
                economy="20_USA",
                base_year=2022,
            )
            _, expr_out, _ = read_export_sheet(expr_output, "LEAP")
            self.assertIn("Expression", expr_out.columns)
            self.assertTrue(
                expr_out["Branch Path"].astype(str).str.endswith("Motor gasoline").any()
            )

            year_cols = [*common_cols[:11], "2022", "2023", *common_cols[11:]]
            year_data = pd.DataFrame(
                [
                    {
                        "BranchID": 1,
                        "VariableID": 1,
                        "ScenarioID": 1,
                        "RegionID": 1,
                        "Branch Path": r"Demand\Industry\Mining and quarrying\Gasoline",
                        "Variable": "Activity Level",
                        "Scenario": "Reference",
                        "Region": "United States",
                        "Scale": "",
                        "Units": "Petajoule",
                        "Per...": "",
                        "2022": 10.0,
                        "2023": 20.0,
                        "Level 1": "Demand",
                        "Level 2": "Industry",
                        "Level 3": "Mining and quarrying",
                        "Level 4": "Gasoline",
                        "Level 5": "",
                    }
                ],
                columns=year_cols,
            )
            year_input = tmp / "industry_year.xlsx"
            year_output = tmp / "industry_year_out.xlsx"
            self._write_export_workbook(year_input, year_data, year_cols)

            remap_industry_export_fuels(
                input_path=year_input,
                output_path=year_output,
                mapping_csv_path=mapping_path,
                esto_data_path=esto_path,
                ninth_data_path=ninth_path,
                subtotal_mapping_path=subtotal_mapping,
                economy="20_USA",
                base_year=2022,
            )
            _, year_out, _ = read_export_sheet(year_output, "LEAP")
            self.assertIn("2022", year_out.columns)
            self.assertNotIn("Expression", year_out.columns)
            self.assertTrue(
                year_out["Branch Path"].astype(str).str.endswith("Motor gasoline").any()
            )

    def test_buildings_remap_supports_expression_and_year_columns(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            tmp = Path(tmp_dir)
            mapping_path = tmp / "buildings_mapping.csv"
            mapping_df = pd.DataFrame(
                [
                    {
                        "sector_key": "RESIDENTIAL PER HOUSEHOLD",
                        "end_use": "Cooking",
                        "technology": "Electric Stove",
                        "canonical_technology": "Electric Stove",
                        "mapping_mode": "direct",
                        "esto_flow": "16.02 Residential",
                        "target_products": "17 Electricity",
                        "notes": "",
                    },
                    {
                        "sector_key": "RESIDENTIAL PER HOUSEHOLD",
                        "end_use": "Cooking",
                        "technology": "Others",
                        "canonical_technology": "Others",
                        "mapping_mode": "split_base_year",
                        "esto_flow": "16.02 Residential",
                        "target_products": "16 Others",
                        "notes": "",
                    },
                ]
            )
            mapping_df.to_csv(mapping_path, index=False)

            esto_path = tmp / "esto.csv"
            pd.DataFrame(
                [
                    {"economy": "20USA", "flows": "16.02 Residential", "products": "17 Electricity", "2022": 100.0},
                    {"economy": "20USA", "flows": "16.02 Residential", "products": "16 Others", "2022": 50.0},
                    {"economy": "20USA", "flows": "16.02 Residential", "products": "12 Solar", "2022": 50.0},
                ]
            ).to_csv(esto_path, index=False)
            subtotal_mapping = Path(__file__).resolve().parents[1] / "config" / "ESTO_subtotal_mapping.xlsx"

            base_cols = [
                "BranchID",
                "VariableID",
                "ScenarioID",
                "RegionID",
                "Branch Path",
                "Variable",
                "Scenario",
                "Region",
                "Scale",
                "Units",
                "Per...",
                "Level 1",
                "Level 2",
                "Level 3",
                "Level 4",
            ]

            year_cols = [*base_cols[:11], "2022", *base_cols[11:]]
            year_input_data = pd.DataFrame(
                [
                    {
                        "BranchID": 1,
                        "VariableID": 1,
                        "ScenarioID": 1,
                        "RegionID": 1,
                        "Branch Path": r"Demand\RESIDENTIAL PER HOUSEHOLD\Cooking\Electric Stove",
                        "Variable": "Activity Level",
                        "Scenario": "Reference",
                        "Region": "United States",
                        "Scale": "",
                        "Units": "Share",
                        "Per...": "",
                        "2022": 80.0,
                        "Level 1": "Demand",
                        "Level 2": "RESIDENTIAL PER HOUSEHOLD",
                        "Level 3": "Cooking",
                        "Level 4": "Electric Stove",
                    },
                    {
                        "BranchID": 2,
                        "VariableID": 2,
                        "ScenarioID": 1,
                        "RegionID": 1,
                        "Branch Path": r"Demand\RESIDENTIAL PER HOUSEHOLD\Cooking\Others",
                        "Variable": "Activity Level",
                        "Scenario": "Reference",
                        "Region": "United States",
                        "Scale": "",
                        "Units": "Share",
                        "Per...": "",
                        "2022": 20.0,
                        "Level 1": "Demand",
                        "Level 2": "RESIDENTIAL PER HOUSEHOLD",
                        "Level 3": "Cooking",
                        "Level 4": "Others",
                    },
                ],
                columns=year_cols,
            )
            year_input = tmp / "buildings_year.xlsx"
            year_output = tmp / "buildings_year_out.xlsx"
            self._write_export_workbook(year_input, year_input_data, year_cols)
            remap_buildings_export_fuels(
                input_path=year_input,
                output_path=year_output,
                mapping_csv_path=mapping_path,
                esto_data_path=esto_path,
                subtotal_mapping_path=subtotal_mapping,
                economy="20_USA",
                base_year=2022,
            )
            _, year_out, _ = read_export_sheet(year_output, "LEAP")
            self.assertIn("2022", year_out.columns)
            self.assertTrue(
                year_out["Branch Path"].astype(str).str.endswith("Electricity").any()
            )

            expr_cols = [*base_cols[:11], "Expression", *base_cols[11:]]
            expr_input_data = year_input_data.drop(columns=["2022"]).copy()
            expr_input_data["Expression"] = "Data(2022,10)"
            expr_input = tmp / "buildings_expr.xlsx"
            expr_output = tmp / "buildings_expr_out.xlsx"
            self._write_export_workbook(expr_input, expr_input_data, expr_cols)
            remap_buildings_export_fuels(
                input_path=expr_input,
                output_path=expr_output,
                mapping_csv_path=mapping_path,
                esto_data_path=esto_path,
                subtotal_mapping_path=subtotal_mapping,
                economy="20_USA",
                base_year=2022,
            )
            _, expr_out, _ = read_export_sheet(expr_output, "LEAP")
            self.assertIn("Expression", expr_out.columns)
            self.assertTrue(
                expr_out["Branch Path"].astype(str).str.endswith("Electricity").any()
            )


if __name__ == "__main__":
    unittest.main()
