from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Iterable


REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_BALANCE_EXPORTS_ROOT = REPO_ROOT / "data" / "leap balances exports"


SCENARIO_CODE_ALIASES = {
    "ref": "REF",
    "reference": "REF",
    "tgt": "TGT",
    "target": "TGT",
}


@dataclass(frozen=True)
class BalanceExportWorkbook:
    path: Path
    economy: str
    scenario_code: str
    date_id: str
    parsed_date: date | None


def normalize_balance_scenario_code(scenario: str) -> str:
    """Return the balance-export filename scenario token."""
    text = str(scenario).strip()
    if not text:
        raise ValueError("Balance-export scenario cannot be blank.")
    return SCENARIO_CODE_ALIASES.get(text.lower(), text.upper())


def _parse_balance_export_date_id(date_id: str) -> date | None:
    """Parse compact workbook date ids such as 492026 or 4212026."""
    token = str(date_id).strip()
    if not token.isdigit():
        return None

    if len(token) == 8:
        for year, month, day in (
            (token[:4], token[4:6], token[6:8]),
            (token[4:8], token[:2], token[2:4]),
        ):
            try:
                return date(int(year), int(month), int(day))
            except ValueError:
                continue

    if len(token) in {6, 7}:
        year_text = token[-4:]
        month_day = token[:-4]
        month_day_splits: list[tuple[str, str]] = []
        if len(month_day) >= 3 and month_day[:2] in {"10", "11", "12"}:
            month_day_splits.append((month_day[:2], month_day[2:]))
        month_day_splits.append((month_day[:1], month_day[1:]))
        if len(month_day) == 4:
            month_day_splits.append((month_day[:2], month_day[2:]))
        for month_text, day_text in month_day_splits:
            try:
                return date(int(year_text), int(month_text), int(day_text))
            except ValueError:
                continue

    return None


def _iter_balance_export_workbooks(
    export_dir: Path,
    *,
    economy: str,
    scenario_code: str,
) -> Iterable[BalanceExportWorkbook]:
    pattern = re.compile(
        r"^full model output all years (?P<date_id>\d{5,8}) (?P<scenario>[A-Za-z]+)(?:\s[^.]*)?\.xlsx$",
        re.IGNORECASE,
    )
    if not export_dir.exists():
        return
    for path in export_dir.glob("*.xlsx"):
        if path.name.startswith("~$"):
            continue
        match = pattern.match(path.name)
        if not match:
            continue
        if normalize_balance_scenario_code(match.group("scenario")) != scenario_code:
            continue
        date_id = match.group("date_id")
        yield BalanceExportWorkbook(
            path=path,
            economy=economy,
            scenario_code=scenario_code,
            date_id=date_id,
            parsed_date=_parse_balance_export_date_id(date_id),
        )


def resolve_balance_export_workbook(
    *,
    economy: str,
    scenario: str,
    date_id: str | None = None,
    exports_root: Path | str = DEFAULT_BALANCE_EXPORTS_ROOT,
) -> Path:
    """Resolve a LEAP balance-export workbook by economy, scenario, and optional date id."""
    economy_text = str(economy).strip()
    if not economy_text:
        raise ValueError("Balance-export economy cannot be blank.")
    scenario_code = normalize_balance_scenario_code(scenario)
    export_dir = Path(exports_root) / economy_text
    candidates = list(
        _iter_balance_export_workbooks(
            export_dir,
            economy=economy_text,
            scenario_code=scenario_code,
        )
    )

    if date_id is not None:
        date_text = str(date_id).strip()
        candidates = [candidate for candidate in candidates if candidate.date_id == date_text]
        if not candidates:
            raise FileNotFoundError(
                "No LEAP balance-export workbook matched "
                f"economy={economy_text!r}, scenario={scenario_code!r}, date_id={date_text!r} "
                f"under {export_dir}."
            )
        if len(candidates) > 1:
            paths = "\n".join(
                f"- {candidate.path}"
                for candidate in sorted(candidates, key=lambda item: item.path.name)
            )
            raise ValueError(
                "Multiple LEAP balance-export workbooks matched "
                f"economy={economy_text!r}, scenario={scenario_code!r}, date_id={date_text!r}:\n{paths}"
            )
        return candidates[0].path

    if not candidates:
        raise FileNotFoundError(
            "No LEAP balance-export workbook matched "
            f"economy={economy_text!r}, scenario={scenario_code!r} under {export_dir}."
        )

    sortable = [
        candidate
        for candidate in candidates
        if candidate.parsed_date is not None
    ]
    if sortable:
        latest_date = max(candidate.parsed_date for candidate in sortable)
        latest = [candidate for candidate in sortable if candidate.parsed_date == latest_date]
    else:
        latest = candidates

    if len(latest) > 1:
        paths = "\n".join(
            f"- {candidate.path}"
            for candidate in sorted(latest, key=lambda item: item.path.name)
        )
        raise ValueError(
            "Multiple LEAP balance-export workbooks matched the latest date for "
            f"economy={economy_text!r}, scenario={scenario_code!r}. Set date_id explicitly.\n{paths}"
        )

    return latest[0].path
