import unittest

from codebase.functions import transformation_analysis_utils as core
from codebase import transfers_workflow as tw


class TestTransferCategoryTemplatesCoverage(unittest.TestCase):
    def test_templates_cover_all_labels_if_used(self) -> None:
        data = core.esto_data
        year_cols = core.esto_year_cols
        start_year = core.YEAR_START_FOR_ANALYSIS

        economies = sorted(data["economy"].dropna().unique())
        for economy in economies:
            flow_codes = tw.select_transfer_flows(data, year_cols, economy)
            if not flow_codes:
                continue
            for flow_code in flow_codes:
                flow_rows = core.select_flow_rows(data, economy, flow_code)
                if flow_rows.empty:
                    continue
                totals, _ = core.summarize_fuel_totals(
                    flow_rows, year_cols, start_year, allow_all_years_fallback=True
                )
                processes = tw._build_template_processes(flow_rows, year_cols, start_year)
                if not processes:
                    # Template coverage is intentionally skipped (fallback will handle full coverage).
                    continue

                covered_inputs = set()
                covered_outputs = set()
                for process in processes:
                    covered_inputs.update(process.get("inputs", []))
                    covered_outputs.update(process.get("outputs", []))

                missing_inputs = sorted(
                    label for label, value in totals.items()
                    if value < 0 and label not in covered_inputs
                )
                missing_outputs = sorted(
                    label for label, value in totals.items()
                    if value > 0 and label not in covered_outputs
                )

                with self.subTest(economy=economy, flow=flow_code):
                    self.assertFalse(
                        missing_inputs or missing_outputs,
                        msg=(
                            "Template processes left out labels. "
                            f"Missing inputs={missing_inputs} "
                            f"Missing outputs={missing_outputs}"
                        ),
                    )

