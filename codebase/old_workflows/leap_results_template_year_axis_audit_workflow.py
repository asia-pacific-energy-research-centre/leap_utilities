from __future__ import annotations

import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.leap_results_workflow import (  # noqa: E402
    TEMPLATE_PATHS,
    TEMPLATE_YEAR_AXIS_AUDIT_CSV,
    audit_template_year_axes,
)


def run_workflow() -> dict[str, object]:
    audit_df, output_csv_path = audit_template_year_axes(
        TEMPLATE_PATHS,
        output_csv_path=TEMPLATE_YEAR_AXIS_AUDIT_CSV,
    )
    normalized_count = int(audit_df.get("normalized", False).fillna(False).astype(bool).sum()) if not audit_df.empty else 0
    return {
        "rows": int(len(audit_df)),
        "normalized_rows": normalized_count,
        "output_csv": str(output_csv_path),
    }


if __name__ == "__main__":  # pragma: no cover
    result = run_workflow()
    print("[OK] LEAP template year-axis audit complete.")
    for key, value in result.items():
        print(f"- {key}: {value}")


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
