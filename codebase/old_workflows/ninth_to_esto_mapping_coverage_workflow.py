from __future__ import annotations

import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.utilities.ninth_to_esto_mapping_coverage import run_mapping_coverage_check
from codebase.utilities.workflow_common import archive_config_dir_once_per_day
from codebase.utilities.output_paths import MAPPINGS_ROOT


def _resolve(path: str | Path) -> Path:
    return path if isinstance(path, Path) else REPO_ROOT / Path(str(path))


MAPPING_PATH = _resolve("config/ninth_pairs_to_esto_pairs.xlsx")
ESTO_DATA_PATH = _resolve("data/00APEC_2025_low_with_subtotals.csv")
NINTH_DATA_PATH = _resolve("data/merged_file_energy_ALL_20251106.csv")
OUTPUT_DIR = MAPPINGS_ROOT / "ninth_to_esto_mapping_coverage"
BASE_YEAR = 2022
PROJECTION_YEARS = tuple(range(2023, 2071))
SCENARIO = "reference"


def main() -> dict[str, object]:
    archive_config_dir_once_per_day()
    return run_mapping_coverage_check(
        mapping_path=MAPPING_PATH,
        esto_data_path=ESTO_DATA_PATH,
        ninth_data_path=NINTH_DATA_PATH,
        output_dir=OUTPUT_DIR,
        base_year=BASE_YEAR,
        projection_years=PROJECTION_YEARS,
        scenario=SCENARIO,
    )


if __name__ == "__main__":
    result = main()
    print(f"summary_csv: {result['summary_csv']}")
    print(f"missing_esto_csv: {result['missing_esto_csv']}")
    print(f"missing_ninth_csv: {result['missing_ninth_csv']}")
    print(f"summary: {result['summary']}")


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
