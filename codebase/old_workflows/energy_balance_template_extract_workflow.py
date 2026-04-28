from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.utilities.energy_balance_template_extractor import run_template_balance_extraction
from codebase.utilities.workflow_common import archive_config_dir_once_per_day
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

WORKBOOK_PATH = _resolve(r"C:\Users\Work\OneDrive - APERC\testing blacanes.xlsx")
OUTPUT_DIR = BALANCE_TABLES_ROOT / "energy_balance_template_extract"
TEMPLATE_SHEET = "Targt Energy Balance 18"
CONVERT_UNITS_TO_PETAJOULE = True


def main() -> dict[str, object]:
    archive_config_dir_once_per_day()
    return run_template_balance_extraction(
        workbook_path=WORKBOOK_PATH,
        output_dir=OUTPUT_DIR,
        template_sheet=TEMPLATE_SHEET,
        mapping_pairs_path=_resolve("config/ninth_pairs_to_esto_pairs.xlsx"),
        codebook_path=_resolve("config/sector_fuel_codes_to_names.xlsx"),
        include_zero_values=True,
        convert_units_to_petajoule=CONVERT_UNITS_TO_PETAJOULE,
    )


if __name__ == "__main__":
    result = main()
    print(f"main_output_csv: {result['main_output_csv']}")
    print(f"raw_csv: {result['raw_csv']}")
    print(f"mapped_csv: {result['mapped_csv']}")
    print(f"coverage_csv: {result['coverage_csv']}")
    print(f"unit_diagnostics_csv: {result['unit_diagnostics_csv']}")
    print(f"diagnostics_csv: {result['diagnostics_csv']}")
    print(f"summary_csv: {result['summary_csv']}")
    print(f"summary: {result['summary']}")


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
