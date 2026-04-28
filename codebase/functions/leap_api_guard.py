from __future__ import annotations

LEAP_API_BLOCKED = True
LEAP_API_BLOCK_REASON = (
    "LEAP API usage is disabled in this repository due a known LEAP API bug. "
    "Use workbook/manual LEAP workflows instead of COM API automation."
)


def ensure_leap_api_allowed(context_label: str) -> None:
    """Raise when LEAP API calls are blocked by repository policy."""
    if LEAP_API_BLOCKED:
        raise RuntimeError(f"{LEAP_API_BLOCK_REASON} Blocked context: {context_label}")


def is_leap_api_allowed() -> bool:
    """Return True only when LEAP API calls are permitted."""
    return not LEAP_API_BLOCKED
