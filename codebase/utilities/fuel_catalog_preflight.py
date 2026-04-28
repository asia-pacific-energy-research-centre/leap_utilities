from __future__ import annotations

import os
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

import pandas as pd

from codebase.functions.analysis_input_write_dispatcher import (
    ensure_analysis_view_api_read_allowed,
)
from codebase.utilities.output_paths import INTEGRATED_LEAP_EXPORTS_ROOT

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_FUEL_CATALOG_PATH = (
    INTEGRATED_LEAP_EXPORTS_ROOT
    / "supporting_files"
    / "checks"
    / "transformation_supply_fuel_branch_catalog.csv"
)
DEFAULT_FUEL_PROBE_PATH = (
    INTEGRATED_LEAP_EXPORTS_ROOT
    / "supporting_files"
    / "checks"
    / "transformation_supply_fuel_branch_catalog_probe.csv"
)
DEFAULT_FULL_MODEL_EXPORT_PATH = REPO_ROOT / "data" / "full model export.xlsx"
DEFAULT_FULL_MODEL_EXPORT_SHEET = "Export"
DEFAULT_STALE_DAYS = 7
MAX_SAMPLE_ITEMS = 8

_STALE_DECISIONS: dict[tuple[str, int], bool] = {}
_PREFLIGHT_RUN_CACHE: set[tuple[str, str, str]] = set()


def _resolve(path: Path | str) -> Path:
    raw = str(path).replace("\\", "/")
    candidate = Path(raw)
    return candidate if candidate.is_absolute() else (REPO_ROOT / candidate)


def _normalize_header_value(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _normalize_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def normalize_fuel_label(value: object) -> str:
    text = _normalize_text(value).lower()
    text = text.replace("&", " and ")
    text = text.replace("/", " ")
    text = " ".join(text.split())
    return text


def _parse_bool_env(name: str, default: bool = False) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return str(raw).strip().lower() in {"1", "true", "yes", "y", "on"}


def _parse_float_env(name: str, default: float) -> float:
    raw = os.getenv(name)
    if raw is None:
        return float(default)
    try:
        return float(str(raw).strip())
    except Exception:
        return float(default)


def _is_interactive() -> bool:
    stream = getattr(sys, "stdin", None)
    return bool(stream and hasattr(stream, "isatty") and stream.isatty())


def _mtime_unix(path: Path) -> int:
    return int(path.stat().st_mtime)


def _age_days(path: Path) -> float:
    modified = datetime.fromtimestamp(path.stat().st_mtime, tz=timezone.utc)
    now = datetime.now(timezone.utc)
    return max((now - modified).total_seconds() / 86400.0, 0.0)


def _sample_text(values: Iterable[str], limit: int = MAX_SAMPLE_ITEMS) -> str:
    cleaned = [str(item).strip() for item in values if str(item or "").strip()]
    if not cleaned:
        return ""
    if len(cleaned) <= limit:
        return ", ".join(cleaned)
    head = ", ".join(cleaned[:limit])
    return f"{head}, ... (+{len(cleaned) - limit} more)"


def _is_path_stale(path: Path | str, *, max_age_days: int = DEFAULT_STALE_DAYS) -> tuple[bool, float | None]:
    resolved = _resolve(path)
    if not resolved.exists():
        return True, None
    age_days = _age_days(resolved)
    return bool(age_days >= max(int(max_age_days), 0)), age_days


def _safe_leap_branch(app, path: str):
    """Return a LEAP branch object or None without raising."""
    branch_path = str(path or "").strip()
    if not branch_path:
        return None
    try:
        branches = app.Branches
        if not branches.Exists(branch_path):
            return None
        return branches.Item(branch_path)
    except Exception:
        return None


def _list_leap_child_branches(parent_branch) -> list[tuple[str, str]]:
    """List child branches as (name, full_path)."""
    rows: list[tuple[str, str]] = []
    if parent_branch is None:
        return rows
    try:
        children = parent_branch.Children
        count = int(children.Count)
    except Exception:
        return rows
    for idx in range(1, count + 1):
        try:
            child = children.Item(idx)
        except Exception:
            continue
        try:
            name = str(child.Name).strip()
        except Exception:
            name = ""
        try:
            full_name = str(child.FullName).strip()
        except Exception:
            full_name = ""
        if not name and full_name and "\\" in full_name:
            name = full_name.rsplit("\\", 1)[-1].strip()
        if name:
            rows.append((name, full_name or name))
    return rows


def _probe_branch_variable_expression(branch_obj, variable_candidates: Iterable[str]) -> tuple[str, str]:
    """Try candidate variables and read expression/value-like field to touch the branch."""
    for var_name in variable_candidates:
        candidate = str(var_name or "").strip()
        if not candidate:
            continue
        try:
            variable = branch_obj.Variable(candidate)
            if variable is None:
                continue
            try:
                _ = str(variable.Expression)
            except Exception:
                _ = ""
            return candidate, "ok"
        except Exception:
            continue
    return "", "variable_not_found"


def _probe_fuel_rows_from_leap(leap_app) -> list[dict[str, object]]:
    """Touch transformation/supply fuel branches in LEAP and return probe rows."""
    ensure_analysis_view_api_read_allowed(
        "fuel_catalog_preflight._probe_fuel_rows_from_leap"
    )
    rows: list[dict[str, object]] = []
    if leap_app is None:
        return rows

    try:
        active_scenario = str(getattr(leap_app, "ActiveScenario", "") or "")
    except Exception:
        active_scenario = ""

    transformation_root = _safe_leap_branch(leap_app, "Transformation")
    for module_name, module_full in _list_leap_child_branches(transformation_root):
        module_path = module_full or f"Transformation\\{module_name}"
        for fuel_group, probe_vars in (
            ("Output Fuels", ("Import Target", "Export Target", "Output Share", "Output")),
            ("Feedstock Fuels", ("Feedstock Fuel Share", "Inputs", "Output")),
            ("Auxiliary Fuels", ("Auxiliary Fuel Use", "Inputs", "Output")),
        ):
            group_path = f"{module_path}\\{fuel_group}"
            group_branch = _safe_leap_branch(leap_app, group_path)
            if group_branch is None:
                continue
            for fuel_name, fuel_full in _list_leap_child_branches(group_branch):
                fuel_path = fuel_full or f"{group_path}\\{fuel_name}"
                fuel_branch = _safe_leap_branch(leap_app, fuel_path)
                if fuel_branch is None:
                    continue
                variable_used, status = _probe_branch_variable_expression(fuel_branch, probe_vars)
                rows.append(
                    {
                        "catalog_type": "transformation",
                        "source_workbook": "__leap_probe__",
                        "scenario": active_scenario,
                        "module_or_root": module_name,
                        "fuel_group": fuel_group,
                        "fuel_name": fuel_name,
                        "branch_path": fuel_path,
                        "variable": variable_used,
                        "catalog_source": "leap_probe",
                        "probe_status": status,
                    }
                )

    for root_name in ("Primary", "Secondary"):
        root_path = f"Resources\\{root_name}"
        root_branch = _safe_leap_branch(leap_app, root_path)
        if root_branch is None:
            continue
        for fuel_name, fuel_full in _list_leap_child_branches(root_branch):
            fuel_path = fuel_full or f"{root_path}\\{fuel_name}"
            fuel_branch = _safe_leap_branch(leap_app, fuel_path)
            if fuel_branch is None:
                continue
            variable_used, status = _probe_branch_variable_expression(
                fuel_branch,
                ("Imports", "Exports", "Indigenous Production", "Unmet Requirements"),
            )
            rows.append(
                {
                    "catalog_type": "supply",
                    "source_workbook": "__leap_probe__",
                    "scenario": active_scenario,
                    "module_or_root": root_name,
                    "fuel_group": "",
                    "fuel_name": fuel_name,
                    "branch_path": fuel_path,
                    "variable": variable_used,
                    "catalog_source": "leap_probe",
                    "probe_status": status,
                }
            )
    return rows


def _rows_to_df(rows: list[dict[str, object]]) -> pd.DataFrame:
    cols = [
        "catalog_type",
        "source_workbook",
        "scenario",
        "module_or_root",
        "fuel_group",
        "fuel_name",
        "branch_path",
        "variable",
        "catalog_source",
        "probe_status",
    ]
    if not rows:
        return pd.DataFrame(columns=cols)
    out = pd.DataFrame(rows).copy()
    for col in cols:
        if col not in out.columns:
            out[col] = ""
    for col in cols:
        out[col] = out[col].map(_normalize_text)
    return out[cols]


def _canonical_fuel_set(df: pd.DataFrame) -> set[tuple[str, str, str, str]]:
    if df.empty:
        return set()
    fuel_group_series = df.get("fuel_group", pd.Series("", index=df.index)).astype(str).str.strip().str.lower()
    return {
        (
            str(row.catalog_type).strip().lower(),
            str(row.module_or_root).strip().lower(),
            str(fuel_group_series.iloc[i]).strip().lower(),
            normalize_fuel_label(row.fuel_name),
        )
        for i, row in enumerate(df.itertuples(index=False))
        if normalize_fuel_label(row.fuel_name)
    }


def _scope_key(catalog_type: str, module_or_root: str) -> tuple[str, str]:
    return str(catalog_type).strip().lower(), str(module_or_root).strip().lower()


def _validate_probe_vs_full_model(
    *,
    full_df: pd.DataFrame,
    probe_df: pd.DataFrame,
    max_probe_only_ratio: float = 0.25,
) -> dict[str, object]:
    """Validate probe-vs-full-model consistency and raise on hard mismatches."""
    if full_df.empty:
        raise RuntimeError(
            f"Fuel catalog validation failed: full model export source is empty at {DEFAULT_FULL_MODEL_EXPORT_PATH}."
        )
    if probe_df.empty:
        raise RuntimeError("Fuel catalog validation failed: LEAP probe returned no rows.")

    full_set = _canonical_fuel_set(full_df)
    probe_set = _canonical_fuel_set(probe_df)
    if not full_set:
        raise RuntimeError("Fuel catalog validation failed: no canonical full-model fuel rows parsed.")
    if not probe_set:
        raise RuntimeError("Fuel catalog validation failed: no canonical probe fuel rows parsed.")

    overlap = full_set & probe_set
    probe_only = probe_set - full_set
    full_only = full_set - probe_set

    if not overlap:
        raise RuntimeError(
            "Fuel catalog validation failed: zero overlap between full model export and LEAP probe."
        )

    probe_only_ratio = float(len(probe_only)) / float(len(probe_set))
    if probe_only_ratio > float(max_probe_only_ratio):
        sample_probe_only = _sample_text(
            [f"{t[0]}::{t[1]}::{t[2]}::{t[3]}" for t in sorted(probe_only)[:MAX_SAMPLE_ITEMS]]
        )
        raise RuntimeError(
            "Fuel catalog validation failed: too many probe-only fuels not present in full model export "
            f"(probe_only={len(probe_only)}, probe_total={len(probe_set)}, ratio={probe_only_ratio:.3f}, "
            f"threshold={max_probe_only_ratio:.3f}). Sample: {sample_probe_only}"
        )

    # Scope-level sanity: every probe scope should overlap full-model in that scope.
    full_by_scope: dict[tuple[str, str], set[str]] = {}
    probe_by_scope: dict[tuple[str, str], set[str]] = {}
    for catalog_type, module_or_root, _fuel_group, fuel_norm in full_set:
        full_by_scope.setdefault(_scope_key(catalog_type, module_or_root), set()).add(fuel_norm)
    for catalog_type, module_or_root, _fuel_group, fuel_norm in probe_set:
        probe_by_scope.setdefault(_scope_key(catalog_type, module_or_root), set()).add(fuel_norm)

    bad_scopes: list[str] = []
    for scope, probe_fuels in probe_by_scope.items():
        full_fuels = full_by_scope.get(scope, set())
        if probe_fuels and full_fuels and not (probe_fuels & full_fuels):
            bad_scopes.append(f"{scope[0]}::{scope[1]}")
    if bad_scopes:
        raise RuntimeError(
            "Fuel catalog validation failed: probe/full-model scope mismatch with zero shared fuels in "
            f"{len(bad_scopes)} scope(s): {_sample_text(sorted(bad_scopes), limit=6)}"
        )

    return {
        "full_total": int(len(full_set)),
        "probe_total": int(len(probe_set)),
        "overlap_total": int(len(overlap)),
        "probe_only_total": int(len(probe_only)),
        "full_only_total": int(len(full_only)),
        "probe_only_ratio": float(probe_only_ratio),
    }


def _resolve_or_connect_leap_app(leap_app=None):
    if leap_app is not None:
        return leap_app
    try:
        from codebase.functions import leap_api  # local import to avoid hard dependency loops at module import time
    except Exception as exc:
        raise RuntimeError(
            "Fuel catalog auto-refresh requires LEAP API, but leap_api import failed."
        ) from exc
    if not leap_api.is_available():
        raise RuntimeError("Fuel catalog auto-refresh requires LEAP API, but it is unavailable.")
    connected = leap_api.connect()
    if connected is None:
        raise RuntimeError("Fuel catalog auto-refresh failed: unable to connect to LEAP.")
    return connected


def refresh_fuel_catalog_from_sources(
    *,
    catalog_path: Path | str = DEFAULT_FUEL_CATALOG_PATH,
    probe_output_path: Path | str = DEFAULT_FUEL_PROBE_PATH,
    full_model_export_path: Path | str = DEFAULT_FULL_MODEL_EXPORT_PATH,
    full_model_sheet: str = DEFAULT_FULL_MODEL_EXPORT_SHEET,
    leap_app=None,
    max_probe_only_ratio: float = 0.25,
) -> dict[str, object]:
    """
    Refresh the shared fuel catalog by probing LEAP and cross-validating against full-model export.
    Raises on validation failures.
    """
    full_rows = _catalog_rows_from_full_model_export(
        source_path=full_model_export_path,
        sheet_name=full_model_sheet,
    )
    full_df = _rows_to_df(full_rows)
    if full_df.empty:
        raise RuntimeError(
            "Fuel catalog refresh failed: full model export yielded no catalog rows at "
            f"{_resolve(full_model_export_path)}."
        )

    app = _resolve_or_connect_leap_app(leap_app=leap_app)
    probe_rows = _probe_fuel_rows_from_leap(app)
    probe_df = _rows_to_df(probe_rows)
    if probe_df.empty:
        raise RuntimeError("Fuel catalog refresh failed: LEAP probe produced no rows.")

    max_probe_only_ratio = _parse_float_env(
        "LEAP_FUEL_CATALOG_MAX_PROBE_ONLY_RATIO",
        max_probe_only_ratio,
    )

    validation = _validate_probe_vs_full_model(
        full_df=full_df,
        probe_df=probe_df,
        max_probe_only_ratio=max_probe_only_ratio,
    )

    probe_path_resolved = _resolve(probe_output_path)
    probe_path_resolved.parent.mkdir(parents=True, exist_ok=True)
    probe_df = (
        probe_df.drop_duplicates(
            subset=["catalog_type", "module_or_root", "fuel_group", "fuel_name", "branch_path"]
        )
        .sort_values(["catalog_type", "module_or_root", "fuel_group", "fuel_name"])
        .reset_index(drop=True)
    )
    probe_df.to_csv(probe_path_resolved, index=False)

    merged_df = pd.concat([full_df, probe_df], axis=0, ignore_index=True)
    merged_df = (
        merged_df.drop_duplicates(
            subset=[
                "catalog_type",
                "source_workbook",
                "scenario",
                "module_or_root",
                "fuel_group",
                "fuel_name",
                "branch_path",
                "variable",
                "catalog_source",
                "probe_status",
            ]
        )
        .sort_values(
            by=[
                "catalog_type",
                "catalog_source",
                "module_or_root",
                "fuel_group",
                "fuel_name",
                "branch_path",
                "variable",
            ]
        )
        .reset_index(drop=True)
    )

    catalog_path_resolved = _resolve(catalog_path)
    catalog_path_resolved.parent.mkdir(parents=True, exist_ok=True)
    merged_df.to_csv(catalog_path_resolved, index=False)

    print(
        "[INFO] Refreshed fuel catalog from LEAP probe + full model export: "
        f"catalog={catalog_path_resolved}, probe={probe_path_resolved}, "
        f"full={validation['full_total']}, probe={validation['probe_total']}, "
        f"overlap={validation['overlap_total']}."
    )
    return {
        "catalog_path": str(catalog_path_resolved),
        "probe_path": str(probe_path_resolved),
        "validation": validation,
    }


def ensure_fuel_catalog_current(
    *,
    catalog_path: Path | str = DEFAULT_FUEL_CATALOG_PATH,
    max_age_days: int = DEFAULT_STALE_DAYS,
    leap_app=None,
    context: str = "",
    auto_refresh: bool = True,
    fail_on_refresh_error: bool = True,
    full_model_export_path: Path | str = DEFAULT_FULL_MODEL_EXPORT_PATH,
    full_model_sheet: str = DEFAULT_FULL_MODEL_EXPORT_SHEET,
    probe_output_path: Path | str = DEFAULT_FUEL_PROBE_PATH,
) -> dict[str, object]:
    """
    Ensure fuel catalog is fresh. If stale/missing and auto_refresh=True, refresh from LEAP probe + full model export.
    """
    catalog_resolved = _resolve(catalog_path)
    stale, age_days = _is_path_stale(catalog_resolved, max_age_days=max_age_days)
    if not stale:
        return {
            "catalog_path": str(catalog_resolved),
            "stale": False,
            "age_days": age_days,
            "refreshed": False,
        }

    reason = f" ({context})" if str(context).strip() else ""
    print(
        "[WARN] Fuel catalog is stale or missing"
        f"{reason}: {catalog_resolved} "
        f"(age={f'{age_days:.1f} days' if age_days is not None else 'missing'}, threshold={max_age_days} days)."
    )

    if not auto_refresh:
        ensure_recent_file_or_prompt(
            catalog_resolved,
            max_age_days=max_age_days,
            context=context or "fuel catalog",
            file_label="fuel catalog",
        )
        return {
            "catalog_path": str(catalog_resolved),
            "stale": True,
            "age_days": age_days,
            "refreshed": False,
        }

    try:
        refreshed = refresh_fuel_catalog_from_sources(
            catalog_path=catalog_resolved,
            probe_output_path=probe_output_path,
            full_model_export_path=full_model_export_path,
            full_model_sheet=full_model_sheet,
            leap_app=leap_app,
        )
        refreshed["stale"] = True
        refreshed["age_days"] = age_days
        refreshed["refreshed"] = True
        return refreshed
    except Exception:
        if fail_on_refresh_error:
            raise
        print(
            "[WARN] Fuel catalog auto-refresh failed, continuing without refresh because "
            "fail_on_refresh_error=False."
        )
        return {
            "catalog_path": str(catalog_resolved),
            "stale": True,
            "age_days": age_days,
            "refreshed": False,
            "refresh_failed": True,
        }


def ensure_recent_file_or_prompt(
    file_path: Path | str,
    *,
    max_age_days: int = DEFAULT_STALE_DAYS,
    context: str = "",
    file_label: str = "file",
) -> dict[str, object]:
    """
    Enforce stale-file confirmation (>= max_age_days) before continuing.

    Behavior:
    - Fresh file: pass silently.
    - Stale file in interactive terminal: prompt continue yes/no.
    - Stale file in non-interactive mode: raise unless LEAP_ALLOW_STALE_IMPORT_FILE=1.
    """
    path = _resolve(file_path)
    if not path.exists():
        return {
            "path": str(path),
            "exists": False,
            "stale": False,
            "age_days": None,
            "continued": True,
        }

    age_days = _age_days(path)
    stale = age_days >= max(int(max_age_days), 0)
    if not stale:
        return {
            "path": str(path),
            "exists": True,
            "stale": False,
            "age_days": age_days,
            "continued": True,
        }

    cache_key = (str(path.resolve()).lower(), _mtime_unix(path))
    cached = _STALE_DECISIONS.get(cache_key)
    if cached is not None:
        if not cached:
            raise RuntimeError(
                f"User declined to continue with stale {file_label}: {path}"
            )
        return {
            "path": str(path),
            "exists": True,
            "stale": True,
            "age_days": age_days,
            "continued": True,
            "cached": True,
        }

    reason = f" ({context})" if str(context).strip() else ""
    message = (
        f"[WARN] Stale {file_label}{reason}: {path} "
        f"(age={age_days:.1f} days, threshold={max_age_days} days)."
    )

    if _parse_bool_env("LEAP_ALLOW_STALE_IMPORT_FILE", default=False):
        print(f"{message} Continuing because LEAP_ALLOW_STALE_IMPORT_FILE=1.")
        _STALE_DECISIONS[cache_key] = True
        return {
            "path": str(path),
            "exists": True,
            "stale": True,
            "age_days": age_days,
            "continued": True,
            "env_override": True,
        }

    if _is_interactive():
        while True:
            choice = input(f"{message} Continue anyway? [y/N]: ").strip().lower()
            if choice in {"y", "yes"}:
                _STALE_DECISIONS[cache_key] = True
                return {
                    "path": str(path),
                    "exists": True,
                    "stale": True,
                    "age_days": age_days,
                    "continued": True,
                }
            if choice in {"", "n", "no"}:
                _STALE_DECISIONS[cache_key] = False
                raise RuntimeError(
                    f"Stopped because stale {file_label} was not approved: {path}"
                )
            print("Please enter 'y' or 'n'.")

    raise RuntimeError(
        f"{message} Non-interactive session cannot prompt. "
        "Set LEAP_ALLOW_STALE_IMPORT_FILE=1 to continue without prompt."
    )


def _read_branch_variable_rows(
    source_path: Path | str,
    *,
    sheet_name: str = "LEAP",
) -> pd.DataFrame:
    """Read workbook rows that contain Branch Path and Variable columns."""
    path = _resolve(source_path)
    if not path.exists():
        return pd.DataFrame()
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
        return df if {"Branch Path", "Variable"}.issubset(df.columns) else pd.DataFrame()

    for header in (0, 2):
        try:
            df = pd.read_excel(path, sheet_name=sheet_name, header=header)
        except Exception:
            continue
        if {"Branch Path", "Variable"}.issubset(df.columns):
            return df

    try:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    except Exception:
        return pd.DataFrame()
    header_row = None
    for idx in range(len(raw.index)):
        values = {
            _normalize_header_value(item).lower()
            for item in raw.iloc[idx].tolist()
        }
        if "branch path" in values and "variable" in values:
            header_row = int(idx)
            break
    if header_row is None:
        return pd.DataFrame()
    data = raw.iloc[header_row + 1 :].copy()
    data.columns = raw.iloc[header_row].tolist()
    return data if {"Branch Path", "Variable"}.issubset(data.columns) else pd.DataFrame()


def _catalog_rows_from_full_model_export(
    *,
    source_path: Path | str = DEFAULT_FULL_MODEL_EXPORT_PATH,
    sheet_name: str = DEFAULT_FULL_MODEL_EXPORT_SHEET,
) -> list[dict[str, object]]:
    """Fallback parser for full model export when catalog CSV is missing."""
    rows_df = _read_branch_variable_rows(source_path, sheet_name=sheet_name)
    if rows_df.empty:
        return []
    out: list[dict[str, object]] = []
    for _, row in rows_df.iterrows():
        branch_path = _normalize_text(row.get("Branch Path"))
        if not branch_path:
            continue
        variable = _normalize_text(row.get("Variable"))
        scenario = _normalize_text(row.get("Scenario"))
        parts = [part.strip() for part in branch_path.split("\\") if part and part.strip()]
        if len(parts) < 2:
            continue

        root = parts[0].lower()
        if root == "transformation":
            module_or_root = parts[1]
            fuel_group = ""
            fuel_name = ""
            for marker in ("Output Fuels", "Feedstock Fuels", "Auxiliary Fuels"):
                if marker in parts:
                    idx = parts.index(marker)
                    if idx + 1 < len(parts):
                        fuel_group = marker
                        fuel_name = parts[idx + 1]
                    break
            if not fuel_name:
                continue
            out.append(
                {
                    "catalog_type": "transformation",
                    "source_workbook": Path(source_path).name,
                    "scenario": scenario,
                    "module_or_root": module_or_root,
                    "fuel_group": fuel_group,
                    "fuel_name": fuel_name,
                    "branch_path": branch_path,
                    "variable": variable,
                    "catalog_source": "full_model_export_fallback",
                    "probe_status": "",
                }
            )
            continue

        if root == "resources":
            if len(parts) < 3:
                continue
            module_or_root = parts[1].title()
            if module_or_root.lower() not in {"primary", "secondary"}:
                continue
            out.append(
                {
                    "catalog_type": "supply",
                    "source_workbook": Path(source_path).name,
                    "scenario": scenario,
                    "module_or_root": module_or_root,
                    "fuel_group": "",
                    "fuel_name": parts[2],
                    "branch_path": branch_path,
                    "variable": variable,
                    "catalog_source": "full_model_export_fallback",
                    "probe_status": "",
                }
            )
            continue

        if root == "demand":
            if len(parts) < 3:
                continue
            variable_norm = variable.lower()
            if not (
                "intensity" in variable_norm
                or "share" in variable_norm
                or variable_norm in {"activity level", "energy demand", "final energy demand"}
            ):
                continue
            out.append(
                {
                    "catalog_type": "demand",
                    "source_workbook": Path(source_path).name,
                    "scenario": scenario,
                    "module_or_root": parts[1],
                    "fuel_group": "",
                    "fuel_name": parts[-1],
                    "branch_path": branch_path,
                    "variable": variable,
                    "catalog_source": "full_model_export_fallback",
                    "probe_status": "",
                }
            )
    return out


def load_fuel_catalog(
    catalog_path: Path | str = DEFAULT_FUEL_CATALOG_PATH,
    *,
    allow_full_model_fallback: bool = True,
    full_model_export_path: Path | str = DEFAULT_FULL_MODEL_EXPORT_PATH,
    full_model_sheet: str = DEFAULT_FULL_MODEL_EXPORT_SHEET,
) -> pd.DataFrame:
    """Load and normalize the shared fuel catalog."""
    path = _resolve(catalog_path)
    if path.exists():
        df = pd.read_csv(path)
    elif allow_full_model_fallback:
        rows = _catalog_rows_from_full_model_export(
            source_path=full_model_export_path,
            sheet_name=full_model_sheet,
        )
        df = pd.DataFrame(rows)
    else:
        return pd.DataFrame()

    required = {"catalog_type", "scenario", "module_or_root", "fuel_name"}
    if df.empty or not required.issubset(df.columns):
        return pd.DataFrame()

    out = df.copy()
    if "fuel_group" not in out.columns:
        out["fuel_group"] = ""
    if "branch_path" not in out.columns:
        out["branch_path"] = ""
    if "variable" not in out.columns:
        out["variable"] = ""

    for col in ("catalog_type", "scenario", "module_or_root", "fuel_group", "fuel_name", "branch_path", "variable"):
        out[col] = out[col].map(_normalize_text)
    out = out[out["fuel_name"] != ""].copy()
    if out.empty:
        return pd.DataFrame()

    out["scenario_norm"] = out["scenario"].str.lower()
    out["catalog_type_norm"] = out["catalog_type"].str.lower()
    out["module_or_root_norm"] = out["module_or_root"].str.lower()
    out["fuel_group_norm"] = out["fuel_group"].str.lower()
    out["fuel_name_norm"] = out["fuel_name"].map(normalize_fuel_label)
    out = out[out["fuel_name_norm"] != ""].copy()
    return out.reset_index(drop=True)


def _scope_rows_from_export_df(
    export_df: pd.DataFrame,
    *,
    scenario: str | None = None,
) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    scenario_norm = _normalize_text(scenario).lower()
    for _, row in export_df.iterrows():
        row_scenario = _normalize_text(row.get("Scenario"))
        if scenario_norm and row_scenario and row_scenario.lower() != scenario_norm:
            continue
        branch_path = _normalize_text(row.get("Branch Path"))
        if not branch_path:
            continue
        variable = _normalize_text(row.get("Variable"))
        parts = [part.strip() for part in branch_path.split("\\") if part and part.strip()]
        if len(parts) < 2:
            continue

        root = parts[0].lower()
        if root == "transformation":
            if len(parts) < 4:
                continue
            fuel_group = ""
            fuel_name = ""
            for marker in ("Output Fuels", "Feedstock Fuels", "Auxiliary Fuels"):
                if marker in parts:
                    idx = parts.index(marker)
                    if idx + 1 < len(parts):
                        fuel_group = marker
                        fuel_name = parts[idx + 1]
                    break
            if not fuel_name:
                continue
            rows.append(
                {
                    "catalog_type": "transformation",
                    "scenario": row_scenario,
                    "module_or_root": parts[1],
                    "fuel_group": fuel_group,
                    "fuel_name": fuel_name,
                    "branch_path": branch_path,
                    "variable": variable,
                }
            )
            continue

        if root == "resources":
            if len(parts) < 3:
                continue
            resource_root = parts[1].title()
            if resource_root.lower() not in {"primary", "secondary"}:
                continue
            rows.append(
                {
                    "catalog_type": "supply",
                    "scenario": row_scenario,
                    "module_or_root": resource_root,
                    "fuel_group": "",
                    "fuel_name": parts[2],
                    "branch_path": branch_path,
                    "variable": variable,
                }
            )
            continue

        if root == "demand":
            if len(parts) < 3:
                continue
            variable_norm = variable.lower()
            if not (
                "intensity" in variable_norm
                or "share" in variable_norm
                or variable_norm in {"activity level", "energy demand", "final energy demand"}
            ):
                continue
            rows.append(
                {
                    "catalog_type": "demand",
                    "scenario": row_scenario,
                    "module_or_root": parts[1],
                    "fuel_group": "",
                    "fuel_name": parts[-1],
                    "branch_path": branch_path,
                    "variable": variable,
                }
            )

    out = pd.DataFrame(rows)
    if out.empty:
        return out
    out["catalog_type_norm"] = out["catalog_type"].str.lower()
    out["module_or_root_norm"] = out["module_or_root"].str.lower()
    out["fuel_group_norm"] = out["fuel_group"].str.lower()
    out["scenario_norm"] = out["scenario"].str.lower()
    out["fuel_name_norm"] = out["fuel_name"].map(normalize_fuel_label)
    return out[out["fuel_name_norm"] != ""].reset_index(drop=True)


def _expected_fuels_for_scope(
    catalog_df: pd.DataFrame,
    *,
    catalog_type: str,
    module_or_root: str,
    scenario: str = "",
) -> set[str]:
    if catalog_df.empty:
        return set()
    scope = catalog_df[
        (catalog_df["catalog_type_norm"] == str(catalog_type).strip().lower())
        & (catalog_df["module_or_root_norm"] == str(module_or_root).strip().lower())
    ].copy()
    if scope.empty:
        return set()
    scenario_norm = str(scenario or "").strip().lower()
    if scenario_norm:
        scoped = scope[scope["scenario_norm"] == scenario_norm]
        if not scoped.empty:
            scope = scoped
    return {value for value in scope["fuel_name_norm"].tolist() if value}


def run_fuel_catalog_preflight(
    *,
    export_path: Path | str,
    sheet_name: str = "LEAP",
    scenario: str | None = None,
    context: str = "",
    strict_missing: bool = False,
    max_age_days: int = DEFAULT_STALE_DAYS,
    catalog_path: Path | str = DEFAULT_FUEL_CATALOG_PATH,
    leap_app=None,
    auto_refresh_stale_catalog: bool = True,
) -> dict[str, object]:
    """
    Shared preflight for LEAP import workbooks.

    - Enforces stale-file confirmation for the workbook itself.
    - Auto-refreshes stale fuel catalogs from LEAP probe + full-model export.
    - Compares workbook fuel scopes against the shared fuel catalog.
    """
    export_path_resolved = _resolve(export_path)
    cache_key = (
        str(export_path_resolved.resolve()).lower(),
        str(sheet_name).strip().lower(),
        str(scenario or "").strip().lower(),
    )
    if cache_key in _PREFLIGHT_RUN_CACHE:
        return {
            "export_path": str(export_path_resolved),
            "skipped": True,
            "reason": "cached",
            "missing_total": 0,
            "extra_total": 0,
        }
    ensure_recent_file_or_prompt(
        export_path_resolved,
        max_age_days=max_age_days,
        context=context or "LEAP import preflight",
        file_label="import workbook",
    )

    ensure_fuel_catalog_current(
        catalog_path=catalog_path,
        max_age_days=max_age_days,
        leap_app=leap_app,
        context=context or "LEAP import preflight",
        auto_refresh=auto_refresh_stale_catalog,
        fail_on_refresh_error=True,
    )

    if _parse_bool_env("LEAP_SKIP_FUEL_CATALOG_PREFLIGHT", default=False):
        _PREFLIGHT_RUN_CACHE.add(cache_key)
        return {
            "export_path": str(export_path_resolved),
            "skipped": True,
            "reason": "LEAP_SKIP_FUEL_CATALOG_PREFLIGHT=1",
            "missing_total": 0,
            "extra_total": 0,
        }

    catalog_df = load_fuel_catalog(catalog_path)
    if catalog_df.empty:
        print(
            "[INFO] Fuel catalog preflight skipped: no catalog rows available at "
            f"{_resolve(catalog_path)}"
        )
        _PREFLIGHT_RUN_CACHE.add(cache_key)
        return {
            "export_path": str(export_path_resolved),
            "skipped": True,
            "reason": "catalog_empty",
            "missing_total": 0,
            "extra_total": 0,
        }

    export_df = _read_branch_variable_rows(export_path_resolved, sheet_name=sheet_name)
    if export_df.empty:
        _PREFLIGHT_RUN_CACHE.add(cache_key)
        return {
            "export_path": str(export_path_resolved),
            "skipped": True,
            "reason": "export_rows_empty",
            "missing_total": 0,
            "extra_total": 0,
        }

    scoped_export = _scope_rows_from_export_df(export_df, scenario=scenario)
    if scoped_export.empty:
        _PREFLIGHT_RUN_CACHE.add(cache_key)
        return {
            "export_path": str(export_path_resolved),
            "skipped": True,
            "reason": "no_scoped_rows",
            "missing_total": 0,
            "extra_total": 0,
        }

    report_rows: list[dict[str, object]] = []
    missing_total = 0
    extra_total = 0
    for (catalog_type, module_or_root), group in scoped_export.groupby(
        ["catalog_type_norm", "module_or_root_norm"], dropna=False
    ):
        sample = group.iloc[0]
        scenario_value = _normalize_text(sample.get("scenario") or scenario or "")
        expected = _expected_fuels_for_scope(
            catalog_df,
            catalog_type=catalog_type,
            module_or_root=module_or_root,
            scenario=scenario_value,
        )
        actual = {item for item in group["fuel_name_norm"].tolist() if item}
        if not expected:
            status = "no_catalog_scope"
            missing: set[str] = set()
            extra = set()
        else:
            missing = expected - actual
            extra = actual - expected
            status = "ok" if not missing else "missing_expected"
        missing_total += len(missing)
        extra_total += len(extra)
        report_rows.append(
            {
                "catalog_type": _normalize_text(sample.get("catalog_type")),
                "module_or_root": _normalize_text(sample.get("module_or_root")),
                "scenario": scenario_value,
                "status": status,
                "expected_count": len(expected),
                "actual_count": len(actual),
                "missing_count": len(missing),
                "extra_count": len(extra),
                "missing_fuels": "; ".join(sorted(missing)),
                "extra_fuels": "; ".join(sorted(extra)),
            }
        )

    report_df = pd.DataFrame(report_rows).sort_values(
        ["catalog_type", "module_or_root"]
    ).reset_index(drop=True)
    reason = f" ({context})" if str(context).strip() else ""
    print(
        "[INFO] Fuel catalog preflight"
        f"{reason}: scopes={len(report_df)}, missing={missing_total}, extra={extra_total}"
    )
    if missing_total:
        missing_examples = report_df[report_df["missing_count"] > 0].copy()
        for _, row in missing_examples.head(5).iterrows():
            print(
                "[WARN] Missing catalog fuels for "
                f"{row['catalog_type']}::{row['module_or_root']} "
                f"(missing={int(row['missing_count'])}): "
                f"{_sample_text(str(row['missing_fuels']).split(';'))}"
            )

    strict = bool(strict_missing or _parse_bool_env("LEAP_PREFLIGHT_STRICT_CATALOG", default=False))
    if strict and missing_total > 0:
        raise RuntimeError(
            f"Fuel catalog preflight failed: {missing_total} expected fuel(s) were missing in "
            f"{export_path_resolved.name}."
        )
    _PREFLIGHT_RUN_CACHE.add(cache_key)

    return {
        "export_path": str(export_path_resolved),
        "skipped": False,
        "missing_total": int(missing_total),
        "extra_total": int(extra_total),
        "scope_count": int(len(report_df)),
        "report": report_df,
    }
