from __future__ import annotations

import re


def clean_fuel_label_for_leap(raw: str | None) -> str:
    """
    Drop leading numeric codes from ESTO product labels (e.g. "01.01 Coking coal").

    LEAP fuel names typically do not include numeric prefixes, and the LEAP API
    rejects fuels like "01 01 Coking coal". We keep the first token that contains
    letters and everything after it.
    """
    if not raw:
        return ""
    text = str(raw).strip()
    match = re.match(r"^\d+_x_(.+)$", text, flags=re.IGNORECASE)
    if match:
        remainder = match.group(1).strip()
        key = remainder.lower().replace("_", "").replace(" ", "")
        overrides = {
            "ammonia": "Ammonia",
            "hydrogen": "Hydrogen",
            "efuel": "Efuel",
        }
        if key in overrides:
            return overrides[key]
        return " ".join(part.capitalize() for part in remainder.split("_") if part)
    tokens = re.split(r"\s+", text)
    for idx, token in enumerate(tokens):
        if any(char.isalpha() for char in token):
            return " ".join(tokens[idx:]).strip()
    return text


__all__ = ["clean_fuel_label_for_leap"]
