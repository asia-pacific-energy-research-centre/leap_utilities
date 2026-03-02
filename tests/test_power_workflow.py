from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

import pandas as pd

from codebase.functions.leap_excel_io import read_export_sheet, write_export_sheet
from codebase import power_workflow


class TestPowerWorkflow(unittest.TestCase):
    def _build_minimal_export(self, path: Path) -> None:
        columns = [
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
            "Method",
            "2022",
            "2023",
            "Level 1",
            "Level 2",
            "Level 3",
            "Level 4",
            "Level 5",
        ]
        header_rows = pd.DataFrame(
            [
                {col: "" for col in columns},
                {col: pd.NA for col in columns},
            ]
        )
        data = pd.DataFrame(
            [
                {
                    "BranchID": 100,
                    "VariableID": 200,
                    "ScenarioID": 300,
                    "RegionID": 1,
                    "Branch Path": r"Transformation\Electricity Generation\Processes\Gas",
                    "Variable": "Fuel Cost",
                    "Scenario": "Optimization",
                    "Region": "United States of America",
                    "Scale": "",
                    "Units": "U.S. Dollar",
                    "Per...": "Gigajoule",
                    "Method": "",
                    "2022": 1.0,
                    "2023": 2.0,
                    "Level 1": "Transformation",
                    "Level 2": "Electricity Generation",
                    "Level 3": "Processes",
                    "Level 4": "Gas",
                    "Level 5": "",
                },
                {
                    "BranchID": 101,
                    "VariableID": 201,
                    "ScenarioID": 1,
                    "RegionID": 1,
                    "Branch Path": r"Transformation\Electricity Generation\Processes\Gas",
                    "Variable": "Fuel Cost",
                    "Scenario": "Current Accounts",
                    "Region": "United States of America",
                    "Scale": "",
                    "Units": "U.S. Dollar",
                    "Per...": "Gigajoule",
                    "Method": "",
                    "2022": 10.0,
                    "2023": 10.0,
                    "Level 1": "Transformation",
                    "Level 2": "Electricity Generation",
                    "Level 3": "Processes",
                    "Level 4": "Gas",
                    "Level 5": "",
                },
            ],
            columns=columns,
        )
        write_export_sheet(
            path=path,
            sheet_name="Export",
            header_rows=header_rows,
            columns=columns,
            data=data,
        )

    def test_scenario_expansion_from_optimization(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            export_path = Path(tmp_dir) / "power_test.xlsx"
            self._build_minimal_export(export_path)

            power_workflow._ensure_export_contains_scenarios(
                export_filename=export_path,
                export_sheet_name="Export",
                target_scenarios=["Reference", "Target"],
                source_scenario_for_missing={
                    "Reference": "Optimization",
                    "Target": "Optimization",
                },
            )

            _, output_df, _ = read_export_sheet(export_path, "Export")
            scenarios = sorted(output_df["Scenario"].dropna().astype(str).str.strip().unique())
            self.assertEqual(scenarios, ["Current Accounts", "Reference", "Target"])
            self.assertNotIn("Optimization", scenarios)

            copied = output_df[output_df["Scenario"].isin(["Reference", "Target"])]
            self.assertTrue(copied["ScenarioID"].isna().all())
            self.assertTrue((copied["2022"] == 1.0).all())
            self.assertTrue((copied["2023"] == 2.0).all())

    def test_cleaned_esto_fuel_validation_unmapped_set(self) -> None:
        repo_root = Path(__file__).resolve().parents[1]
        report_path = repo_root / "intermediate_data" / "_test_power_fuel_report.csv"
        report_df = power_workflow._validate_export_fuels_against_esto(
            export_path=repo_root / "data" / "power export.xlsx",
            sheet_name="Export",
            esto_data_path=repo_root / "data" / "00APEC_2024_low.csv",
            report_path=report_path,
        )
        unmapped = set(report_df.loc[report_df["status"] == "unmapped", "fuel_label"].tolist())
        self.assertEqual(unmapped, {"Biomass", "Coal Bituminous"})

    def test_preconfigured_skip_audit_rows(self) -> None:
        data = pd.DataFrame(
            [
                {
                    "Scenario": "Reference",
                    "Branch Path": r"Transformation\Electricity Generation\Processes\Gas",
                    "Variable": "Dispatchable",
                },
                {
                    "Scenario": "Current Accounts",
                    "Branch Path": r"Transformation\Electricity Generation\Processes\Coal",
                    "Variable": "Optimize",
                },
                {
                    "Scenario": "Target",
                    "Branch Path": r"Transformation\Electricity Generation\Processes\Solar",
                    "Variable": "Dispatchable",
                },
                {
                    "Scenario": "Reference",
                    "Branch Path": r"Transformation\Electricity Generation\Processes\Solar",
                    "Variable": "Fuel Cost",
                },
            ]
        )
        rows = power_workflow._build_preconfigured_skip_audit_rows(
            export_data=data,
            scenario_names=["Reference", "Current Accounts"],
            skip_variables={"dispatchable", "optimize"},
        )
        self.assertEqual(len(rows), 2)
        actions = {row["action"] for row in rows}
        scenarios = {row["scenario"] for row in rows}
        variables = {row["variable"] for row in rows}
        self.assertEqual(actions, {"preconfigured_skip"})
        self.assertEqual(scenarios, {"Reference", "Current Accounts"})
        self.assertEqual(variables, {"Dispatchable", "Optimize"})

    def test_validate_hardcoded_overrides(self) -> None:
        with self.assertRaisesRegex(ValueError, "missing required field 'expression'"):
            power_workflow._validate_hardcoded_overrides(
                [
                    {
                        "scenario": "Reference",
                        "branch_path": r"Transformation\Electricity Generation",
                        "variable": "Planning Reserve Margin",
                    }
                ]
            )

        validated = power_workflow._validate_hardcoded_overrides(
            [
                {
                    "scenario": " Reference ",
                    "branch_path": r" Transformation\Electricity Generation ",
                    "variable": " Planning Reserve Margin ",
                    "expression": " 25 ",
                    "units": " Percent ",
                    "reason": " Example ",
                }
            ]
        )
        self.assertEqual(
            validated,
            [
                {
                    "scenario": "Reference",
                    "branch_path": r"Transformation\Electricity Generation",
                    "variable": "Planning Reserve Margin",
                    "expression": "25",
                    "units": "Percent",
                    "reason": "Example",
                }
            ],
        )

    def test_excluded_branch_path_detection(self) -> None:
        excluded_path = (
            r"Transformation\Electricity Generation\Processes\Coal\Feedstock Fuels\Coal Bituminous\Carbon Dioxide"
        )
        included_path = (
            r"Transformation\Electricity Generation\Processes\Coal\Feedstock Fuels\Coal Bituminous"
        )
        self.assertTrue(power_workflow._is_branch_excluded_from_api(excluded_path))
        self.assertFalse(power_workflow._is_branch_excluded_from_api(included_path))

    def test_api_filtered_export_removes_excluded_rows(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            source_path = Path(tmp_dir) / "source.xlsx"
            filtered_path = Path(tmp_dir) / "filtered.xlsx"
            self._build_minimal_export(source_path)
            header_rows, data, columns = read_export_sheet(source_path, "Export")

            extra_row = data.iloc[0].copy()
            extra_row["Branch Path"] = (
                r"Transformation\Electricity Generation\Processes\Coal\Feedstock Fuels\Coal Bituminous\Carbon Dioxide"
            )
            extra_row["Variable"] = "Avg Environmental Loading"
            data2 = pd.concat([data, pd.DataFrame([extra_row])], ignore_index=True)
            write_export_sheet(
                path=source_path,
                sheet_name="Export",
                header_rows=header_rows,
                columns=columns,
                data=data2,
            )

            output_path, excluded_rows = power_workflow._build_api_filtered_export(
                source_export_path=source_path,
                sheet_name="Export",
                api_export_path=filtered_path,
            )
            self.assertEqual(output_path, filtered_path.resolve())
            self.assertEqual(len(excluded_rows), 1)

            _, filtered_data, _ = read_export_sheet(filtered_path, "Export")
            self.assertEqual(len(filtered_data), len(data2) - 1)
            self.assertFalse(
                filtered_data["Branch Path"]
                .astype(str)
                .str.endswith("Carbon Dioxide")
                .any()
            )


if __name__ == "__main__":
    unittest.main()
