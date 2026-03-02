from __future__ import annotations

import importlib.util
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd

from codebase.functions.leap_series_comparison import (
    TransportResultsComparisonConfig,
    _build_fuel_column_lookup,
    _load_code_to_name,
    _load_transport_fuel_aliases,
    _resolve_fuel_mapping,
    _sector_column_from_code,
    run_transport_results_table_comparison,
)


def _write_results_sheet(
    writer: pd.ExcelWriter,
    *,
    sheet_name: str,
    a1: str,
    a2: str,
    a3: str,
    a4: str,
    header_label: str,
    years: list[int],
    rows: list[list[object]],
) -> None:
    width = 1 + len(years)
    data: list[list[object]] = []
    data.append([a1] + [pd.NA] * (width - 1))
    data.append([a2] + [pd.NA] * (width - 1))
    data.append([a3] + [pd.NA] * (width - 1))
    data.append([a4] + [pd.NA] * (width - 1))
    data.append([pd.NA] * width)
    data.append([header_label] + years)
    data.extend(rows)
    pd.DataFrame(data).to_excel(writer, sheet_name=sheet_name, index=False, header=False)


def _build_transport_fixture(tmp_path: Path) -> TransportResultsComparisonConfig:
    leap_path = tmp_path / "transport_results.xlsx"
    with pd.ExcelWriter(leap_path, engine="openpyxl") as writer:
        _write_results_sheet(
            writer,
            sheet_name="RandomPassenger",
            a1="Final Energy Demand",
            a2="Scenario: Target, Region: United States of America",
            a3=r"Branch: Demand\Passenger road",
            a4="Units: Petajoules",
            header_label="Fuel",
            years=[2022, 2023],
            rows=[
                ["Motor gasoline", 70.0, 90.0],
                ["Unmapped fuel", 1.0, 1.0],
                ["Total", 71.0, 91.0],
            ],
        )
        _write_results_sheet(
            writer,
            sheet_name="RoadFreightSheet",
            a1="Final Energy Demand",
            a2="Scenario: Target, Region: United States of America",
            a3=r"Branch: Demand\Freight road",
            a4="Units: Thousand Petajoules",
            header_label="Fuel",
            years=[2022, 2023],
            rows=[
                ["Motor gasoline", 0.01, 0.02],
                ["Total", 0.01, 0.02],
            ],
        )
        _write_results_sheet(
            writer,
            sheet_name="INTL",
            a1="Final Energy Demand",
            a2="Scenario: Target, Region: United States of America",
            a3=r"Branch: Demand\International transport",
            a4="Units: Petajoules",
            header_label="Fuel",
            years=[2022, 2023],
            rows=[
                ["Motor gasoline", 50.0, 60.0],
                ["Total", 50.0, 60.0],
            ],
        )
        _write_results_sheet(
            writer,
            sheet_name="Pipe-View",
            a1="Final Energy Demand",
            a2="Scenario: Target, Region: United States of America, All Fuels",
            a3=r"Branch: Demand\Pipeline transport",
            a4="Units: Petajoules",
            header_label="Branch",
            years=[2022, 2023],
            rows=[
                ["Natural gas", 5.0, 6.0],
                ["Total", 5.0, 6.0],
            ],
        )
        _write_results_sheet(
            writer,
            sheet_name="DemandBreakdown",
            a1="Final Energy Demand",
            a2="Scenario: Target, Region: United States of America",
            a3=r"Branch: Demand",
            a4="Units: Thousand Petajoules",
            header_label="Branch",
            years=[2022, 2023],
            rows=[
                ["Passenger road", 71.0, 91.0],
                ["Freight road", 0.01, 0.02],
                ["International transport", 50.0, 60.0],
                ["Pipeline transport", 5.0, 6.0],
                ["Total", 126.01, 157.02],
            ],
        )

    branch_mapping_path = tmp_path / "branch_map.csv"
    pd.DataFrame(
        [
            {
                "branch_path": r"Demand\Passenger road",
                "ninth_sector_code": "15_02_01_passenger",
                "active": True,
                "include_in_demand_total": True,
                "notes": "",
            },
            {
                "branch_path": r"Demand\Freight road",
                "ninth_sector_code": "15_02_02_freight",
                "active": True,
                "include_in_demand_total": True,
                "notes": "",
            },
            {
                "branch_path": r"Demand\International transport",
                "ninth_sector_code": "04_international_marine_bunkers",
                "active": True,
                "include_in_demand_total": True,
                "notes": "",
            },
            {
                "branch_path": r"Demand\Pipeline transport",
                "ninth_sector_code": "15_05_pipeline_transport",
                "active": True,
                "include_in_demand_total": True,
                "notes": "",
            },
        ]
    ).to_csv(branch_mapping_path, index=False)

    alias_path = tmp_path / "fuel_aliases.csv"
    pd.DataFrame(
        [
            {
                "leap_fuel_label": "Gasoline",
                "codebook_name": "Motor gasoline",
                "ninth_fuel_override": "",
                "esto_product_override": "",
                "active": True,
                "notes": "",
            }
        ]
    ).to_csv(alias_path, index=False)

    code_to_name_path = tmp_path / "code_to_name.xlsx"
    pd.DataFrame(
        [
            {
                "9th_label": "15_02_road",
                "9th_column": "sub1sectors",
                "esto_label": "15.02 Road",
                "esto_column": "flows",
                "name": "Road",
            },
            {
                "9th_label": "15_02_01_passenger",
                "9th_column": "sub2sectors",
                "esto_label": "",
                "esto_column": "",
                "name": "Road passenger",
            },
            {
                "9th_label": "15_02_02_freight",
                "9th_column": "sub2sectors",
                "esto_label": "",
                "esto_column": "",
                "name": "Road freight",
            },
            {
                "9th_label": "04_international_marine_bunkers",
                "9th_column": "sectors",
                "esto_label": "04 International marine bunkers",
                "esto_column": "flows",
                "name": "International marine bunkers",
            },
            {
                "9th_label": "15_05_pipeline_transport",
                "9th_column": "sub1sectors",
                "esto_label": "15.05 Pipeline transport",
                "esto_column": "flows",
                "name": "Pipeline transport",
            },
            {
                "9th_label": "07_01_motor_gasoline",
                "9th_column": "subfuels",
                "esto_label": "07.01 Motor gasoline",
                "esto_column": "products",
                "name": "Motor gasoline",
            },
            {
                "9th_label": "08_01_natural_gas",
                "9th_column": "subfuels",
                "esto_label": "08.01 Natural gas",
                "esto_column": "products",
                "name": "Natural gas",
            },
        ]
    ).to_excel(code_to_name_path, sheet_name="code_to_name", index=False)

    esto_path = tmp_path / "esto.csv"
    pd.DataFrame(
        [
            {
                "economy": "01AUS",
                "flows": "15.02 Road",
                "products": "07.01 Motor gasoline",
                "2022": 100.0,
            },
            {
                "economy": "01AUS",
                "flows": "04 International marine bunkers",
                "products": "07.01 Motor gasoline",
                "2022": -50.0,
            },
            {
                "economy": "01AUS",
                "flows": "15.05 Pipeline transport",
                "products": "08.01 Natural gas",
                "2022": 5.0,
            },
        ]
    ).to_csv(esto_path, index=False)

    subtotal_mapping_path = tmp_path / "subtotal_mapping.xlsx"
    pd.DataFrame(
        [
            {"flow": "15.02 Road", "product": "07.01 Motor gasoline", "is_subtotal": False},
            {
                "flow": "04 International marine bunkers",
                "product": "07.01 Motor gasoline",
                "is_subtotal": False,
            },
            {
                "flow": "15.05 Pipeline transport",
                "product": "08.01 Natural gas",
                "is_subtotal": False,
            },
        ]
    ).to_excel(subtotal_mapping_path, index=False)

    ninth_path = tmp_path / "ninth.csv"
    pd.DataFrame(
        [
            {
                "scenarios": "reference",
                "economy": "01_AUS",
                "sectors": "15_transport_sector",
                "sub1sectors": "15_02_road",
                "sub2sectors": "15_02_01_passenger",
                "sub3sectors": "x",
                "sub4sectors": "x",
                "fuels": "07_petroleum_products",
                "subfuels": "07_01_motor_gasoline",
                "subtotal_layout": False,
                "subtotal_results": False,
                "2022": 60.0,
                "2023": 80.0,
            },
            {
                "scenarios": "reference",
                "economy": "01_AUS",
                "sectors": "15_transport_sector",
                "sub1sectors": "15_02_road",
                "sub2sectors": "15_02_02_freight",
                "sub3sectors": "x",
                "sub4sectors": "x",
                "fuels": "07_petroleum_products",
                "subfuels": "07_01_motor_gasoline",
                "subtotal_layout": False,
                "subtotal_results": False,
                "2022": 40.0,
                "2023": 20.0,
            },
            {
                "scenarios": "reference",
                "economy": "01_AUS",
                "sectors": "04_international_marine_bunkers",
                "sub1sectors": "x",
                "sub2sectors": "x",
                "sub3sectors": "x",
                "sub4sectors": "x",
                "fuels": "07_petroleum_products",
                "subfuels": "07_01_motor_gasoline",
                "subtotal_layout": False,
                "subtotal_results": False,
                "2022": -50.0,
                "2023": -60.0,
            },
            {
                "scenarios": "reference",
                "economy": "01_AUS",
                "sectors": "15_transport_sector",
                "sub1sectors": "15_05_pipeline_transport",
                "sub2sectors": "x",
                "sub3sectors": "x",
                "sub4sectors": "x",
                "fuels": "08_gas",
                "subfuels": "08_01_natural_gas",
                "subtotal_layout": False,
                "subtotal_results": False,
                "2022": 5.0,
                "2023": 6.0,
            },
        ]
    ).to_csv(ninth_path, index=False)

    mapping_pairs_path = tmp_path / "ninth_to_esto.xlsx"
    pd.DataFrame(
        [
            {
                "9th_sector": "15_02_road",
                "9th_fuel": "07_01_motor_gasoline",
                "esto_flow": "15.02 Road",
                "esto_product": "07.01 Motor gasoline",
            },
            {
                "9th_sector": "04_international_marine_bunkers",
                "9th_fuel": "07_01_motor_gasoline",
                "esto_flow": "04 International marine bunkers",
                "esto_product": "07.01 Motor gasoline",
            },
            {
                "9th_sector": "15_05_pipeline_transport",
                "9th_fuel": "08_01_natural_gas",
                "esto_flow": "15.05 Pipeline transport",
                "esto_product": "08.01 Natural gas",
            },
        ]
    ).to_excel(mapping_pairs_path, index=False)

    return TransportResultsComparisonConfig(
        leap_results_file=leap_path,
        economy="01_AUS",
        scenario="Target",
        region="United States of America",
        branch_sector_mapping_csv=branch_mapping_path,
        fuel_aliases_csv=alias_path,
        code_to_name_path=code_to_name_path,
        code_to_name_sheet="code_to_name",
        esto_data_path=esto_path,
        ninth_data_path=ninth_path,
        subtotal_mapping_path=subtotal_mapping_path,
        ninth_to_esto_mapping_path=mapping_pairs_path,
        base_year=2022,
        projection_start_year=2023,
        projection_end_year=2023,
        share_year_offset=1,
        ninth_scenario="reference",
        output_dir=tmp_path / "out",
    )


class TestTransportResultsTableComparison(unittest.TestCase):
    def test_sector_level_resolution_from_code_depth(self) -> None:
        self.assertEqual(_sector_column_from_code("04_international_marine_bunkers"), "sectors")
        self.assertEqual(_sector_column_from_code("15_05_pipeline_transport"), "sub1sectors")
        self.assertEqual(_sector_column_from_code("15_02_01_passenger"), "sub2sectors")

    def test_fuel_mapping_via_codebook_and_aliases(self) -> None:
        with TemporaryDirectory() as tmp:
            config = _build_transport_fixture(Path(tmp))
            aliases_df = _load_transport_fuel_aliases(Path(config.fuel_aliases_csv))
            codebook_df = _load_code_to_name(
                Path(config.code_to_name_path), config.code_to_name_sheet
            )
            fuel_lookup = _build_fuel_column_lookup(codebook_df)
            resolved = _resolve_fuel_mapping(
                leap_fuel_label="Gasoline",
                aliases_df=aliases_df,
                codebook_df=codebook_df,
                fuel_column_lookup=fuel_lookup,
            )
            self.assertEqual(resolved["ninth_fuel_code"], "07_01_motor_gasoline")
            self.assertEqual(resolved["esto_product"], "07.01 Motor gasoline")

    def test_sheet_identification_uses_a1_a4_not_sheet_name(self) -> None:
        with TemporaryDirectory() as tmp:
            config = _build_transport_fixture(Path(tmp))
            artifacts = run_transport_results_table_comparison(config)
            long_df = pd.read_csv(artifacts.comparison_long_csv)
            self.assertIn(r"Demand\Passenger road", set(long_df["branch_path"].tolist()))

            inventory_df = pd.read_csv(Path(config.output_dir) / "sheet_inventory.csv")
            accepted = inventory_df[inventory_df["status"] == "accepted"]
            self.assertIn("RandomPassenger", set(accepted["sheet_name"].tolist()))

    def test_parses_fuel_and_branch_first_column_variants(self) -> None:
        with TemporaryDirectory() as tmp:
            config = _build_transport_fixture(Path(tmp))
            artifacts = run_transport_results_table_comparison(config)
            long_df = pd.read_csv(artifacts.comparison_long_csv)
            pipeline = long_df[
                (long_df["branch_path"] == r"Demand\Pipeline transport")
                & (long_df["fuel_label"] == "Natural gas")
                & (long_df["year"] == 2022)
            ]
            self.assertEqual(len(pipeline), 1)
            self.assertAlmostEqual(float(pipeline["leap_value"].iloc[0]), 5.0, places=6)

    def test_unit_scaling_from_a4_petajoules_vs_thousand_petajoules(self) -> None:
        with TemporaryDirectory() as tmp:
            config = _build_transport_fixture(Path(tmp))
            artifacts = run_transport_results_table_comparison(config)
            long_df = pd.read_csv(artifacts.comparison_long_csv)
            freight = long_df[
                (long_df["branch_path"] == r"Demand\Freight road")
                & (long_df["fuel_label"] == "Motor gasoline")
                & (long_df["year"] == 2022)
            ]
            self.assertEqual(len(freight), 1)
            self.assertAlmostEqual(float(freight["leap_value"].iloc[0]), 10.0, places=6)

    def test_base_year_parent_allocation_uses_base_year_plus_one_shares(self) -> None:
        with TemporaryDirectory() as tmp:
            config = _build_transport_fixture(Path(tmp))
            artifacts = run_transport_results_table_comparison(config)
            long_df = pd.read_csv(artifacts.comparison_long_csv)

            passenger = long_df[
                (long_df["branch_path"] == r"Demand\Passenger road")
                & (long_df["fuel_label"] == "Motor gasoline")
                & (long_df["year"] == 2022)
            ].iloc[0]
            freight = long_df[
                (long_df["branch_path"] == r"Demand\Freight road")
                & (long_df["fuel_label"] == "Motor gasoline")
                & (long_df["year"] == 2022)
            ].iloc[0]

            self.assertAlmostEqual(float(passenger["reference_value"]), 80.0, places=6)
            self.assertAlmostEqual(float(freight["reference_value"]), 20.0, places=6)
            self.assertEqual(
                passenger["reference_source"], "esto_base_year_allocated_from_parent"
            )

    def test_bunker_reference_sign_flipped_positive(self) -> None:
        with TemporaryDirectory() as tmp:
            config = _build_transport_fixture(Path(tmp))
            artifacts = run_transport_results_table_comparison(config)
            long_df = pd.read_csv(artifacts.comparison_long_csv)
            intl_2022 = long_df[
                (long_df["branch_path"] == r"Demand\International transport")
                & (long_df["fuel_label"] == "Motor gasoline")
                & (long_df["year"] == 2022)
            ].iloc[0]
            intl_2023 = long_df[
                (long_df["branch_path"] == r"Demand\International transport")
                & (long_df["fuel_label"] == "Motor gasoline")
                & (long_df["year"] == 2023)
            ].iloc[0]
            self.assertAlmostEqual(float(intl_2022["reference_value"]), 50.0, places=6)
            self.assertAlmostEqual(float(intl_2023["reference_value"]), 60.0, places=6)

    def test_demand_total_includes_positive_international(self) -> None:
        with TemporaryDirectory() as tmp:
            config = _build_transport_fixture(Path(tmp))
            artifacts = run_transport_results_table_comparison(config)
            long_df = pd.read_csv(artifacts.comparison_long_csv)
            demand_2022 = long_df[
                (long_df["series_id"] == "Demand|__total__") & (long_df["year"] == 2022)
            ].iloc[0]
            self.assertAlmostEqual(float(demand_2022["leap_value"]), 136.0, places=6)
            self.assertAlmostEqual(float(demand_2022["reference_value"]), 155.0, places=6)

    def test_missing_mapping_and_missing_reference_flags(self) -> None:
        with TemporaryDirectory() as tmp:
            config = _build_transport_fixture(Path(tmp))
            artifacts = run_transport_results_table_comparison(config)
            status_df = pd.read_csv(artifacts.mapping_status_csv)
            unmapped = status_df[
                (status_df["branch_path"] == r"Demand\Passenger road")
                & (status_df["fuel_label"] == "Unmapped fuel")
            ].iloc[0]
            self.assertFalse(bool(unmapped["has_fuel_mapping"]))

            unmatched_df = pd.read_csv(artifacts.unmatched_leap_rows_csv)
            self.assertIn("Unmapped fuel", set(unmatched_df["row_label"].tolist()))
            reason = unmatched_df.loc[
                unmatched_df["row_label"] == "Unmapped fuel", "unmatched_reason"
            ].iloc[0]
            self.assertEqual(reason, "fuel_not_mapped")

    def test_00_apec_aggregates_all_economies_for_reference(self) -> None:
        with TemporaryDirectory() as tmp:
            config = _build_transport_fixture(Path(tmp))

            # Add a second economy to ESTO and 9th inputs.
            esto_df = pd.read_csv(config.esto_data_path)
            esto_df = pd.concat(
                [
                    esto_df,
                    pd.DataFrame(
                        [
                            {
                                "economy": "02BRA",
                                "flows": "15.02 Road",
                                "products": "07.01 Motor gasoline",
                                "2022": 50.0,
                            },
                            {
                                "economy": "02BRA",
                                "flows": "04 International marine bunkers",
                                "products": "07.01 Motor gasoline",
                                "2022": -10.0,
                            },
                        ]
                    ),
                ],
                ignore_index=True,
            )
            esto_df.to_csv(config.esto_data_path, index=False)

            ninth_df = pd.read_csv(config.ninth_data_path)
            ninth_df = pd.concat(
                [
                    ninth_df,
                    pd.DataFrame(
                        [
                            {
                                "scenarios": "reference",
                                "economy": "02_BRA",
                                "sectors": "15_transport_sector",
                                "sub1sectors": "15_02_road",
                                "sub2sectors": "15_02_01_passenger",
                                "sub3sectors": "x",
                                "sub4sectors": "x",
                                "fuels": "07_petroleum_products",
                                "subfuels": "07_01_motor_gasoline",
                                "subtotal_layout": False,
                                "subtotal_results": False,
                                "2022": 25.0,
                                "2023": 25.0,
                            },
                            {
                                "scenarios": "reference",
                                "economy": "02_BRA",
                                "sectors": "15_transport_sector",
                                "sub1sectors": "15_02_road",
                                "sub2sectors": "15_02_02_freight",
                                "sub3sectors": "x",
                                "sub4sectors": "x",
                                "fuels": "07_petroleum_products",
                                "subfuels": "07_01_motor_gasoline",
                                "subtotal_layout": False,
                                "subtotal_results": False,
                                "2022": 25.0,
                                "2023": 25.0,
                            },
                            {
                                "scenarios": "reference",
                                "economy": "02_BRA",
                                "sectors": "04_international_marine_bunkers",
                                "sub1sectors": "x",
                                "sub2sectors": "x",
                                "sub3sectors": "x",
                                "sub4sectors": "x",
                                "fuels": "07_petroleum_products",
                                "subfuels": "07_01_motor_gasoline",
                                "subtotal_layout": False,
                                "subtotal_results": False,
                                "2022": -10.0,
                                "2023": -12.0,
                            },
                        ]
                    ),
                ],
                ignore_index=True,
            )
            ninth_df.to_csv(config.ninth_data_path, index=False)

            config.economy = "00_APEC"
            artifacts = run_transport_results_table_comparison(config)
            long_df = pd.read_csv(artifacts.comparison_long_csv)

            passenger_2022 = long_df[
                (long_df["branch_path"] == r"Demand\Passenger road")
                & (long_df["fuel_label"] == "Motor gasoline")
                & (long_df["year"] == 2022)
            ].iloc[0]
            passenger_2023 = long_df[
                (long_df["branch_path"] == r"Demand\Passenger road")
                & (long_df["fuel_label"] == "Motor gasoline")
                & (long_df["year"] == 2023)
            ].iloc[0]
            intl_2022 = long_df[
                (long_df["branch_path"] == r"Demand\International transport")
                & (long_df["fuel_label"] == "Motor gasoline")
                & (long_df["year"] == 2022)
            ].iloc[0]
            intl_2023 = long_df[
                (long_df["branch_path"] == r"Demand\International transport")
                & (long_df["fuel_label"] == "Motor gasoline")
                & (long_df["year"] == 2023)
            ].iloc[0]

            # Parent ESTO (150) split by aggregated 2023 child shares (105/150).
            self.assertAlmostEqual(float(passenger_2022["reference_value"]), 105.0, places=6)
            # 9th projection should be summed across economies.
            self.assertAlmostEqual(float(passenger_2023["reference_value"]), 105.0, places=6)
            # Bunkers should be aggregated then sign-flipped positive.
            self.assertAlmostEqual(float(intl_2022["reference_value"]), 60.0, places=6)
            self.assertAlmostEqual(float(intl_2023["reference_value"]), 72.0, places=6)

    @unittest.skipIf(
        importlib.util.find_spec("matplotlib") is None,
        "matplotlib is not installed in this environment.",
    )
    def test_chart_generation_smoke(self) -> None:
        with TemporaryDirectory() as tmp:
            config = _build_transport_fixture(Path(tmp))
            artifacts = run_transport_results_table_comparison(config)
            png_files = list(Path(artifacts.charts_dir).glob("*.png"))
            self.assertGreaterEqual(len(png_files), 1)


if __name__ == "__main__":
    unittest.main()
