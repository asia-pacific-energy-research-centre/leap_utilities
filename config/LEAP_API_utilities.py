#%%
"""Utilities for interacting with LEAP's API and extracting snapshots."""

from __future__ import annotations

import os
import re
import sys
from pathlib import Path

# NOTE: avoid shadowing LEAP Units collection with matplotlib.units

# ---------------------------------------------------------------------------
# Environment setup
# ---------------------------------------------------------------------------
REPO_ROOT = Path(__file__).resolve().parents[1]
CURRENT_DIR = Path.cwd()
if CURRENT_DIR != REPO_ROOT:
    os.chdir(REPO_ROOT)
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

# Also add the code folder to the path.
if str("code") not in [p for p in sys.path if p.endswith("code")]:
    sys.path.insert(0, str(REPO_ROOT / "code"))

from code.leap_core import connect_to_leap

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
_UNIT_CLASS_OPTIONS = [
    "Bad Unit Class [200]",
    "Bad Unit Class [201]",
    "Bad Unit Class [202]",
    "Bad Unit Class [203]",
    "Fuel Economy",
    "Fuel Share",
    "No data",
    "Unspecifed Unit",
    "TDLoss",
    "Environmental",
    "Transport",
    "Currency",
    "Power",
    "Area",
    "Length",
    "Volume",
    "Energy",
    "Mass",
    "Other",
    "Percent Saturation",
    "Percent",
    "Efficiency",
    "Share",
]
_UNIT_CLASS_TOKEN_OPTIONS = sorted(
    (option.split() for option in _UNIT_CLASS_OPTIONS),
    key=len,
    reverse=True,
)
_FLOAT_RE = re.compile(r"^-?\d+\.\d+$")


def _match_unit_class(tokens: list[str], line: str) -> tuple[str, list[str]]:
    for class_tokens in _UNIT_CLASS_TOKEN_OPTIONS:
        if len(class_tokens) <= len(tokens) and tokens[-len(class_tokens) :] == class_tokens:
            unit_class = " ".join(class_tokens)
            remainder = tokens[: -len(class_tokens)]
            return unit_class, remainder
    raise ValueError(f"Unable to match unit class in line: {line}")


def _find_abbreviation_and_name(
    remaining_tokens: list[str], ratio_str: str, line: str
) -> tuple[str, str]:
    slash_positions = [idx for idx, ch in enumerate(ratio_str) if ch == "/"]
    best_denom = None
    best_tokens = None
    for pos in slash_positions:
        denom = ratio_str[pos + 1 :].strip()
        denom_tokens = denom.split()
        if not denom_tokens:
            continue
        if len(remaining_tokens) >= len(denom_tokens) and remaining_tokens[-len(denom_tokens) :] == denom_tokens:
            if best_tokens is None or len(denom_tokens) > len(best_tokens):
                best_denom = denom
                best_tokens = denom_tokens
    if best_tokens is None:
        raise ValueError(f"Unable to match abbreviation in line: {line}")
    name_tokens = remaining_tokens[: -len(best_tokens)]
    name = " ".join(name_tokens).strip()
    if not name:
        raise ValueError(f"Empty name after parsing line: {line}")
    return name, best_denom


def _parse_units_dump(dump: str) -> tuple[list[dict[str, object]], dict[str, int]]:
    lines = [line.strip() for line in dump.splitlines() if line.strip()]
    header_count = None
    header_max_id = None
    start_idx = 0
    if lines:
        header_parts = lines[0].split()
        if len(header_parts) == 2 and all(part.isdigit() for part in header_parts):
            header_count = int(header_parts[0])
            header_max_id = int(header_parts[1])
            start_idx = 1

    units: list[dict[str, object]] = []
    for line in lines[start_idx:]:
        tokens = line.split()
        if len(tokens) < 6:
            raise ValueError(f"Too few tokens in line: {line}")
        index = int(tokens[0])
        float_idx = None
        for i in range(len(tokens) - 1, -1, -1):
            if _FLOAT_RE.match(tokens[i]):
                float_idx = i
                break
        if float_idx is None or float_idx < 2:
            raise ValueError(f"Conversion factor not found in line: {line}")
        unit_id = int(tokens[float_idx - 1])
        conversion_factor = " ".join(tokens[float_idx:])
        ratio_str = " ".join(tokens[float_idx + 1 :])
        mid_tokens = tokens[1 : float_idx - 1]
        unit_class, remaining_tokens = _match_unit_class(mid_tokens, line)
        name, abbreviation = _find_abbreviation_and_name(remaining_tokens, ratio_str, line)
        units.append(
            {
                "index": index,
                "name": name,
                "abbreviation": abbreviation,
                "unit_class": unit_class,
                "id": unit_id,
                "conversion_factor": conversion_factor,
            }
        )

    count = len(units)
    max_id = max((unit["id"] for unit in units), default=0)
    if header_count is not None and header_count != count:
        raise ValueError(f"Header count {header_count} does not match parsed {count}")
    if header_max_id is not None and header_max_id != max_id:
        raise ValueError(f"Header max id {header_max_id} does not match parsed {max_id}")
    return units, {"count": count, "max_id": max_id}


def _build_index(units: list[dict[str, object]], key: str) -> dict[object, object]:
    index: dict[object, object] = {}
    for unit in units:
        value = unit[key]
        if value in index:
            existing = index[value]
            if isinstance(existing, list):
                existing.append(unit)
            else:
                index[value] = [existing, unit]
        else:
            index[value] = unit
    return index


def _build_list_index(units: list[dict[str, object]], key: str) -> dict[object, list[dict[str, object]]]:
    index: dict[object, list[dict[str, object]]] = {}
    for unit in units:
        index.setdefault(unit[key], []).append(unit)
    return index

# ---------------------------------------------------------------------------
# Extracted data (fill _LEAP_UNITS_DUMP when refreshing from LEAP)
# ---------------------------------------------------------------------------
_LEAP_UNITS_DUMP = "".strip()


_LEAP_UNITS, _LEAP_UNITS_META = _parse_units_dump(_LEAP_UNITS_DUMP)

# ---------------------------------------------------------------------------
# Public data exports
# ---------------------------------------------------------------------------
LEAP_UNITS = _LEAP_UNITS
LEAP_UNITS_META = _LEAP_UNITS_META
LEAP_UNITS_BY_NAME = _build_index(LEAP_UNITS, "name")
LEAP_UNITS_BY_ABBREVIATION = _build_index(LEAP_UNITS, "abbreviation")
LEAP_UNITS_BY_ID = _build_index(LEAP_UNITS, "id")
LEAP_UNITS_BY_INDEX = _build_index(LEAP_UNITS, "index")
LEAP_UNITS_BY_CLASS = _build_list_index(LEAP_UNITS, "unit_class")

__all__ = [
    "LEAP_UNITS",
    "LEAP_UNITS_META",
    "LEAP_UNITS_BY_NAME",
    "LEAP_UNITS_BY_ABBREVIATION",
    "LEAP_UNITS_BY_ID",
    "LEAP_UNITS_BY_INDEX",
    "LEAP_UNITS_BY_CLASS",
    "extract_leap_units_DUMP",
]

# ---------------------------------------------------------------------------
# Extractors
# ---------------------------------------------------------------------------
L = connect_to_leap()


def extract_leap_units_DUMP() -> str:
    str_dump = ""
    leap_units = L.Units
    print(leap_units.Count, leap_units.MaxID)

    for i in range(1, leap_units.Count + 1):
        u = leap_units.Item(i)
        # print(i, u.Name, u.Abbreviation, u.UnitClass, u.ID, u.ConversionFactor)
        str_dump += f"{i} {u.Name} {u.Abbreviation} {u.UnitClass} {u.ID} {u.ConversionFactor}\n"
    return str_dump

# ---------------------------------------------------------------------------
# CLI entrypoints
# ---------------------------------------------------------------------------
_CLI_ACTIONS = {
    "dump_units": extract_leap_units_DUMP,
}


def _strip_jupyter_args(argv: list[str]) -> list[str]:
    """Drop Jupyter/kernel launcher args so notebooks can import/run safely."""
    cleaned: list[str] = []
    skip_next = False
    for arg in argv:
        if skip_next:
            skip_next = False
            continue
        if arg in {"-f", "--f"}:
            skip_next = True
            continue
        if arg.startswith("--f=") or arg.startswith("-f="):
            continue
        cleaned.append(arg)
    return cleaned


def main(argv: list[str] | None = None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    argv = _strip_jupyter_args(argv)
    if not argv:
        print("Available actions:")
        for action in sorted(_CLI_ACTIONS):
            print(f"- {action}")
        return 0
    action = argv[0]
    if action not in _CLI_ACTIONS:
        if action.startswith("-"):
            print("Available actions:")
            for action_name in sorted(_CLI_ACTIONS):
                print(f"- {action_name}")
            return 0
        print(f"Unknown action: {action}")
        print("Available actions:")
        for action_name in sorted(_CLI_ACTIONS):
            print(f"- {action_name}")
        return 2
    result = _CLI_ACTIONS[action]()
    if result is not None:
        print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
#%%
