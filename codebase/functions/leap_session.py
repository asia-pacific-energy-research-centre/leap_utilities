from __future__ import annotations

from typing import Any

_PINNED_LEAP_APP: Any = None


def pin_leap_app(app: Any) -> Any:
    """Pin a LEAP COM application object for reuse in this Python process."""
    global _PINNED_LEAP_APP
    _PINNED_LEAP_APP = app
    return app


def clear_pinned_leap_app() -> None:
    """Clear the pinned LEAP COM application object."""
    global _PINNED_LEAP_APP
    _PINNED_LEAP_APP = None


def get_live_pinned_leap_app() -> Any | None:
    """
    Return the pinned LEAP app if still alive, otherwise clear and return None.

    Liveness check is intentionally lightweight and non-mutating.
    """
    app = _PINNED_LEAP_APP
    if app is None:
        return None
    try:
        # Access a common property to validate COM object liveness.
        _ = app.ActiveArea
        return app
    except Exception:
        clear_pinned_leap_app()
        return None

