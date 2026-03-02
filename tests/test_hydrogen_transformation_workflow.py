import unittest

import pandas as pd

from codebase.functions import transformation_analysis_utils as core
from codebase import hydrogen_transformation_workflow as hw


class TestHydrogenTransformationWorkflow(unittest.TestCase):
    def test_modeled_outputs_match_ninth_source(self) -> None:
        checks, summary = hw.verify_hydrogen_output_reproduction(
            economies=list(core.ECONOMIES_TO_ANALYZE),
            abs_tolerance=1e-6,
            rel_tolerance=1e-6,
        )
        failed_rows = checks[~checks["passes_tolerance"]]
        with self.subTest(summary=summary):
            self.assertGreater(summary["row_count"], 0)
            self.assertEqual(
                summary["failed_rows"],
                0,
                msg=failed_rows.head(20).to_string(index=False),
            )

    def test_log_rows_are_unique_for_export_pivot(self) -> None:
        rows = hw.collect_hydrogen_rows(
            include_all_economies=True,
            aggregate_economy_label="ALL_ECONOMIES",
        )
        hw.deduplicate_hydrogen_output_targets(rows)
        mapping = hw._build_hydrogen_display_mapping()
        scenarios = list(core.SCENARIOS_TO_EXPORT)
        scenario_configs = {
            scenario: core.get_scenario_export_config(scenario)
            for scenario in scenarios
        }
        base_year, final_year = core.compute_combined_year_range(
            core.EXPORT_BASE_YEAR,
            core.EXPORT_FINAL_YEAR,
            scenario_configs,
        )
        log_rows = []
        for scenario in scenarios:
            log_rows.extend(
                core.build_transformation_log_rows(
                    rows,
                    scenario,
                    core.EXPORT_REGION,
                    base_year,
                    final_year,
                    mapping,
                    scenario_config=scenario_configs[scenario],
                )
            )
        log_df = pd.DataFrame(log_rows)
        duplicate_mask = log_df.duplicated(
            subset=[
                "Branch_Path",
                "Scenario",
                "Measure",
                "Units",
                "Scale",
                "Per...",
                "Date",
            ],
            keep=False,
        )
        duplicates = log_df.loc[duplicate_mask]
        self.assertFalse(
            duplicate_mask.any(),
            msg=duplicates.head(20).to_string(index=False),
        )

    def test_boolean_aggregate_label_uses_all_economies_name(self) -> None:
        rows = hw.collect_hydrogen_rows(
            include_all_economies=True,
            aggregate_economy_label=True,
        )
        economies = sorted({str(record.get("economy")) for record in rows})
        self.assertEqual(economies, ["ALL_ECONOMIES"])
