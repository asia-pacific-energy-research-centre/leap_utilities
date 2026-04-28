#%%
from __future__ import annotations

import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.utilities.detailed_balance_from_esto import build_detailed_balance_from_esto
from codebase.utilities.workflow_common import archive_config_dir_once_per_day
from codebase.utilities.workflow_outputs import build_workflow_output_layout, write_output_manifest
from codebase.utilities.output_paths import BALANCE_TABLES_ROOT


def _resolve(path: str | Path) -> Path:
    if isinstance(path, Path):
        candidate = path
    else:
        normalized = str(path).replace("\\", "/")
        if len(normalized) >= 3 and normalized[1:3] == ":/":
            drive = normalized[0].lower()
            rest = normalized[3:]
            return Path(f"/mnt/{drive}/{rest}")
        candidate = Path(normalized)
    if candidate.is_absolute():
        return candidate
    return REPO_ROOT / candidate


#%%
TEMPLATE_WORKBOOK_PATH = _resolve(r"C:\Users\Work\github\leap_utilities\data\detailed balance table output example.xlsx")
ESTO_DATA_PATH = _resolve(r"C:\Users\Work\github\leap_utilities\data\00APEC_2024_low_with_subtotals.csv")
CODEBOOK_PATH = _resolve("config/sector_fuel_codes_to_names.xlsx")
OUTPUT_DIR = BALANCE_TABLES_ROOT / "detailed_balance_from_esto"
OUTPUT_WORKBOOK_PATH = OUTPUT_DIR / "detailed_balance_20USA_2022_from_esto.xlsx"

ECONOMY = "20USA"
YEAR = 2022
SCENARIO_LABEL = "Reference"
UNITS = "Billion Gigajoule"
AREA_LABEL = "USA (ESTO reconstructed)"


#%%
def main() -> dict[str, object]:
    archive_config_dir_once_per_day()
    layout = build_workflow_output_layout(OUTPUT_DIR)
    result = build_detailed_balance_from_esto(
        template_workbook_path=TEMPLATE_WORKBOOK_PATH,
        esto_data_path=ESTO_DATA_PATH,
        codebook_path=CODEBOOK_PATH,
        output_workbook_path=OUTPUT_WORKBOOK_PATH,
        output_dir=layout.supporting,
        sheet_name="Energy Balance",
        economy=ECONOMY,
        year=YEAR,
        scenario_label=SCENARIO_LABEL,
        units=UNITS,
        area_label=AREA_LABEL,
        include_subtotals=True,
    )
    manifest_path = write_output_manifest(
        out_dir=layout.root,
        primary_outputs={"output_workbook": result["output_workbook"]},
        supporting_outputs={
            "header_mapping_csv": result["header_mapping_csv"],
            "row_mapping_csv": result["row_mapping_csv"],
            "summary_csv": result["summary_csv"],
            "nonzero_pairs_csv": result["nonzero_pairs_csv"],
        },
        primary_output_descriptions={
            "output_workbook": "Primary reconstructed detailed balance workbook.",
        },
        supporting_output_descriptions={
            "header_mapping_csv": "Fuel-header mapping audit for the template columns.",
            "row_mapping_csv": "Row mapping audit for template flow resolution.",
            "summary_csv": "One-row summary of the detailed balance reconstruction.",
            "nonzero_pairs_csv": "Nonzero ESTO flow/product pairs used to populate the workbook.",
        },
        notes=[
            "The reconstructed workbook stays at the workflow root.",
            "Audit tables live under supporting_files/.",
        ],
    )
    result["output_manifest_json"] = str(manifest_path)
    return result


#%%
# Notebook run cell.
# Set to True when running this file through an interactive notebook-like runner.
RUN_WORKFLOW = False
WORKFLOW_RESULT: dict[str, object] | None = None
if RUN_WORKFLOW:
    WORKFLOW_RESULT = main()
    for key, value in WORKFLOW_RESULT.items():
        print(f"{key}: {value}")
#%%
