from __future__ import annotations

import unittest

import pandas as pd

from codebase import minor_demand_workflow as mdw


class TestMinorDemandWorkflow(unittest.TestCase):
    def _set_global(self, name: str, value):
        previous = getattr(mdw, name)
        setattr(mdw, name, value)
        self.addCleanup(lambda: setattr(mdw, name, previous))

    def test_sector_share_activity_and_sector_total_alignment(self) -> None:
        projection_years = [2023, 2024, 2025]
        projection_df = pd.DataFrame(
            [
                {
                    "economy_key": "20_USA",
                    "esto_flow": "16.03 Agriculture",
                    "esto_product": "Coal",
                    2023: 30.0,
                    2024: 40.0,
                    2025: 0.0,
                },
                {
                    "economy_key": "20_USA",
                    "esto_flow": "16.03 Agriculture",
                    "esto_product": "Oil",
                    2023: 70.0,
                    2024: 60.0,
                    2025: 0.0,
                },
            ]
        )

        sector_total = mdw.build_activity_series_from_allocated(
            projection_df=projection_df,
            flow="16.03 Agriculture",
            economy_key="20_USA",
            projection_years=projection_years,
        )
        self.assertEqual(sector_total, {2023: 100.0, 2024: 100.0, 2025: 0.0})

        coal_share = mdw.build_fuel_activity_series_sector_share_from_allocated(
            projection_df=projection_df,
            flow="16.03 Agriculture",
            esto_product="Coal",
            economy_key="20_USA",
            projection_years=projection_years,
        )
        oil_share = mdw.build_fuel_activity_series_sector_share_from_allocated(
            projection_df=projection_df,
            flow="16.03 Agriculture",
            esto_product="Oil",
            economy_key="20_USA",
            projection_years=projection_years,
        )
        self.assertEqual(coal_share, {2023: 0.3, 2024: 0.4, 2025: 0.0})
        self.assertEqual(oil_share, {2023: 0.7, 2024: 0.6, 2025: 0.0})

        for year in [2023, 2024]:
            self.assertAlmostEqual(coal_share[year] + oil_share[year], 1.0, places=9)

    def test_sector_share_mode_forces_intensity_to_one(self) -> None:
        self._set_global("FUEL_ACTIVITY_MODE", "sector_share")
        self._set_global("INTENSITY_MODE", "custom")
        self._set_global("DEFAULT_INTENSITY", 2.5)
        self._set_global("FUEL_INTENSITY_OVERRIDES", {("16.03 Agriculture", "Coal"): 4.0})

        intensity = mdw.build_intensity_series(
            flow="16.03 Agriculture",
            esto_product="Coal",
            projection_years=[2023, 2024],
            fuel_projection=pd.DataFrame(),
            sector_projection=pd.DataFrame(),
            flow_to_sectors={},
            economy_key="20_USA",
            mapping=pd.DataFrame(),
            allocated_projection=pd.DataFrame(),
        )
        self.assertEqual(intensity, {2023: 1.0, 2024: 1.0})

    def test_sector_share_mode_sets_fuel_activity_units_to_share(self) -> None:
        configured = {"units": "Petajoule", "scale": "kilo", "per": "year"}
        resolved = mdw._resolve_fuel_activity_units(configured, "sector_share")
        self.assertEqual(resolved, {"units": "Share", "scale": "", "per": ""})

    def test_energy_activity_mode_sets_fuel_activity_units_to_unspecified(self) -> None:
        configured = {"units": "Petajoule", "scale": "kilo", "per": "year"}
        resolved = mdw._resolve_fuel_activity_units(
            configured,
            "activity_as_energy_intensity_as_one",
        )
        self.assertEqual(resolved, {"units": "Unspecified Unit", "scale": "", "per": ""})

    def test_sector_activity_units_set_to_unspecified_for_sector_share_and_energy_modes(self) -> None:
        configured = {"units": "Petajoule", "scale": "kilo", "per": "year"}
        share_mode = mdw._resolve_sector_activity_units(configured, "sector_share")
        energy_mode = mdw._resolve_sector_activity_units(
            configured,
            "activity_as_energy_intensity_as_one",
        )
        self.assertEqual(share_mode, {"units": "Unspecified Unit", "scale": "", "per": ""})
        self.assertEqual(energy_mode, {"units": "Unspecified Unit", "scale": "", "per": ""})

    def test_energy_activity_mode_regression_activity_equals_energy_share_times_total(self) -> None:
        self._set_global("FUEL_ACTIVITY_MODE", "activity_as_energy_intensity_as_one")

        projection_df = pd.DataFrame(
            [
                {
                    "economy_key": "20_USA",
                    "esto_flow": "16.03 Agriculture",
                    "esto_product": "Coal",
                    2023: 30.0,
                    2024: 40.0,
                },
                {
                    "economy_key": "20_USA",
                    "esto_flow": "16.03 Agriculture",
                    "esto_product": "Oil",
                    2023: 70.0,
                    2024: 60.0,
                },
            ]
        )
        activity = mdw.build_fuel_activity_series(
            flow="16.03 Agriculture",
            esto_product="Coal",
            projection_years=[2023, 2024],
            fuel_projection=pd.DataFrame(),
            sector_projection=pd.DataFrame(),
            flow_to_sectors={},
            economy_key="20_USA",
            mapping=pd.DataFrame(),
            sector_activity={2023: 100.0, 2024: 200.0},
            allocated_projection=projection_df,
        )
        self.assertEqual(activity, {2023: 30.0, 2024: 80.0})

    def test_legacy_fuel_share_alias_matches_energy_activity_mode(self) -> None:
        kwargs = dict(
            flow="16.03 Agriculture",
            esto_product="Coal",
            projection_years=[2023, 2024],
            fuel_projection=pd.DataFrame(),
            sector_projection=pd.DataFrame(),
            flow_to_sectors={},
            economy_key="20_USA",
            mapping=pd.DataFrame(),
            sector_activity={2023: 100.0, 2024: 200.0},
            allocated_projection=pd.DataFrame(
                [
                    {
                        "economy_key": "20_USA",
                        "esto_flow": "16.03 Agriculture",
                        "esto_product": "Coal",
                        2023: 30.0,
                        2024: 40.0,
                    },
                    {
                        "economy_key": "20_USA",
                        "esto_flow": "16.03 Agriculture",
                        "esto_product": "Oil",
                        2023: 70.0,
                        2024: 60.0,
                    },
                ]
            ),
        )
        self._set_global("FUEL_ACTIVITY_MODE", "activity_as_energy_intensity_as_one")
        canonical = mdw.build_fuel_activity_series(**kwargs)
        self._set_global("FUEL_ACTIVITY_MODE", "fuel_share")
        legacy = mdw.build_fuel_activity_series(**kwargs)
        self.assertEqual(legacy, canonical)

    def test_sector_share_mode_falls_back_when_allocated_projection_missing(self) -> None:
        self._set_global("FUEL_ACTIVITY_MODE", "sector_share")

        flow = "16.03 Agriculture"
        mapping = pd.DataFrame(
            [
                {
                    "esto_flow": flow,
                    "esto_product": "Coal",
                    "9th_fuel": "01_coal",
                }
            ]
        )
        sector_projection = pd.DataFrame(
            [
                {"economy_key": "20_USA", "9th_sector": "s1", 2023: 100.0, 2024: 200.0}
            ]
        )
        fuel_projection = pd.DataFrame(
            [
                {
                    "economy_key": "20_USA",
                    "9th_sector": "s1",
                    "9th_fuel": "01_coal",
                    2023: 30.0,
                    2024: 50.0,
                }
            ]
        )

        share_activity = mdw.build_fuel_activity_series(
            flow=flow,
            esto_product="Coal",
            projection_years=[2023, 2024],
            fuel_projection=fuel_projection,
            sector_projection=sector_projection,
            flow_to_sectors={flow: ["s1"]},
            economy_key="20_USA",
            mapping=mapping,
            sector_activity={2023: 999.0, 2024: 999.0},
            allocated_projection=pd.DataFrame(),
        )
        self.assertEqual(share_activity, {2023: 0.3, 2024: 0.25})

    def test_build_allocated_projection_table_uses_passed_mapping_scope(self) -> None:
        ninth_data = pd.DataFrame(
            [
                {
                    "economy": "20_USA",
                    "economy_key": "20USA",
                    "scenarios": "reference",
                    "sectors": "s1",
                    "fuels": "f1",
                    2023: 100.0,
                }
            ]
        )
        esto_data = pd.DataFrame(
            [
                {
                    "economy": "20_USA",
                    "economy_key": "20USA",
                    "flows": "Flow A",
                    "products": "Fuel A",
                    2022: 60.0,
                },
                {
                    "economy": "20_USA",
                    "economy_key": "20USA",
                    "flows": "Flow B",
                    "products": "Fuel B",
                    2022: 40.0,
                },
            ]
        )
        mapping = pd.DataFrame(
            [
                {
                    "9th_sector": "s1",
                    "9th_fuel": "f1",
                    "esto_flow": "Flow A",
                    "esto_product": "Fuel A",
                }
            ]
        )

        allocated = mdw.build_allocated_projection_table(
            ninth_data=ninth_data,
            esto_data=esto_data,
            projection_years=[2023],
            scenario="reference",
            mapping=mapping,
        )

        self.assertEqual(float(allocated[2023].sum()), 100.0)
        self.assertEqual(allocated["esto_flow"].tolist(), ["Flow A"])

    def test_add_base_year_to_fuel_activity_uses_esto_share_for_sector_share_mode(self) -> None:
        esto_data = pd.DataFrame(
            [
                {
                    "economy_key": "20USA",
                    "flows": "16.03 Agriculture",
                    "products": "Coal",
                    2022: 30.0,
                },
                {
                    "economy_key": "20USA",
                    "flows": "16.03 Agriculture",
                    "products": "Oil",
                    2022: 70.0,
                },
            ]
        )

        updated = mdw.add_base_year_to_fuel_activity(
            fuel_activity_series={2023: 0.4},
            esto_data=esto_data,
            flow="16.03 Agriculture",
            esto_product="Coal",
            economy_key="20USA",
            base_year=2022,
            mode="sector_share",
        )

        self.assertEqual(updated, {2023: 0.4, 2022: 0.3})


if __name__ == "__main__":
    unittest.main()
