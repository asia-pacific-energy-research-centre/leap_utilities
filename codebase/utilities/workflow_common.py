from __future__ import annotations

import argparse
import re
import sys
import time
import zipfile
from datetime import date, datetime
from pathlib import Path
from typing import Callable, Iterable, Sequence

import pandas as pd

from codebase.functions.leap_core import (
    connect_to_leap,
    create_branches_from_export_file,
    fill_branches_from_export_file,
)
from codebase.functions.analysis_input_write_dispatcher import (
    dispatch_analysis_input_write,
)
from codebase.utilities import fuel_catalog_preflight

AGGREGATE_ECONOMY_LABELS = {"00_APEC", "ALL_ECONOMIES", "ALL"}
REPO_ROOT = Path(__file__).resolve().parents[2]


def format_duration(seconds: float) -> str:
    """Return a compact human-readable duration string."""
    total_seconds = max(float(seconds), 0.0)
    hours = int(total_seconds // 3600)
    minutes = int((total_seconds % 3600) // 60)
    secs = total_seconds - (hours * 3600) - (minutes * 60)
    return f"{hours}h {minutes}m {secs:.1f}s"


class WorkflowTimer:
    """Small stage timer for notebook-safe workflow scripts."""

    def __init__(
        self,
        workflow_name: str,
        *,
        enabled: bool = True,
        print_each: bool = True,
    ) -> None:
        self.workflow_name = str(workflow_name).strip() or "workflow"
        self.enabled = bool(enabled)
        self.print_each = bool(print_each)
        self.started_at = datetime.now()
        self._last_at = self.started_at
        self._start_perf = time.perf_counter()
        self._last_perf = self._start_perf
        self._records: list[dict[str, object]] = []

    @property
    def records(self) -> list[dict[str, object]]:
        return list(self._records)

    def lap(self, stage: str, *, status: str = "success") -> dict[str, object]:
        """Record elapsed time since the previous lap."""
        if not self.enabled:
            return {}
        now_perf = time.perf_counter()
        now = datetime.now()
        duration = now_perf - self._last_perf
        stage_started_at = self._last_at
        self._last_perf = now_perf
        self._last_at = now
        record = {
            "workflow": self.workflow_name,
            "stage_order": len(self._records) + 1,
            "stage": str(stage).strip() or "stage",
            "status": str(status).strip() or "success",
            "started_at": stage_started_at.isoformat(timespec="seconds"),
            "ended_at": now.isoformat(timespec="seconds"),
            "duration_seconds": round(duration, 3),
            "duration_formatted": format_duration(duration),
        }
        self._records.append(record)
        if self.print_each:
            print(
                "[TIMING] "
                f"{self.workflow_name} | {record['stage']} | "
                f"{record['duration_formatted']}"
            )
        return record

    def finish(self, *, status: str = "success") -> dict[str, object]:
        """Record total workflow runtime."""
        if not self.enabled:
            return {}
        now = datetime.now()
        duration = time.perf_counter() - self._start_perf
        record = {
            "workflow": self.workflow_name,
            "stage_order": len(self._records) + 1,
            "stage": "total",
            "status": str(status).strip() or "success",
            "started_at": self.started_at.isoformat(timespec="seconds"),
            "ended_at": now.isoformat(timespec="seconds"),
            "duration_seconds": round(duration, 3),
            "duration_formatted": format_duration(duration),
        }
        self._records.append(record)
        if self.print_each:
            print(
                "[TIMING] "
                f"{self.workflow_name} | total | {record['duration_formatted']}"
            )
        return record

    def write_csv(self, path: Path | str) -> Path | None:
        """Write timing records to CSV and return the path."""
        if not self.enabled or not self._records:
            return None
        output_path = Path(path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame(self._records).to_csv(output_path, index=False)
        if self.print_each:
            print(f"[TIMING] {self.workflow_name} timing written to {output_path}")
        return output_path


def emit_completion_beep(
    *,
    success: bool = True,
    style: str = "simple",
    enabled: bool = True,
    count: int = 1,
    frequency_hz: int = 880,
    duration_ms: int = 180,
    pause_seconds: float = 0.12,
) -> None:
    """Emit an audible completion signal (winsound, notebook audio, terminal bell)."""
    if not bool(enabled):
        return

    count = max(int(count), 1)
    frequency = max(int(frequency_hz), 37)
    duration = max(int(duration_ms), 50)
    pause_seconds = max(float(pause_seconds), 0.0)
    if not success:
        count = max(count, 2)
        frequency = max(frequency - 180, 37)
        if style == "chime":
            style = "error"

    if style == "chime":
        tone_plan = [(659, 90), (784, 90), (988, 140)]  # E5, G5, B5
        gap_ms = 40
    elif style == "error":
        tone_plan = [(440, 140), (330, 180)]  # A4 -> E4 (descending)
        gap_ms = 60
    else:
        tone_plan = [(frequency, duration)] * count
        gap_ms = int(pause_seconds * 1000)

    try:
        import winsound  # type: ignore

        for index, (freq_hz, tone_duration_ms) in enumerate(tone_plan):
            try:
                winsound.Beep(max(int(freq_hz), 37), max(int(tone_duration_ms), 50))
            except Exception:
                winsound.MessageBeep()
            if gap_ms > 0 and index < len(tone_plan) - 1:
                time.sleep(gap_ms / 1000.0)
        return
    except Exception:
        pass

    try:
        from IPython import get_ipython  # type: ignore
        from IPython.display import Javascript, display  # type: ignore

        ip = get_ipython()
        shell_name = type(ip).__name__ if ip is not None else ""
        if shell_name == "ZMQInteractiveShell":
            tones_js = ", ".join(
                f"{{freq: {max(int(freq_hz), 37)}, durMs: {max(int(tone_duration_ms), 50)}}}"
                for freq_hz, tone_duration_ms in tone_plan
            )
            js = f"""
            (() => {{
              const AudioCtx = window.AudioContext || window.webkitAudioContext;
              if (!AudioCtx) return;
              const tones = [{tones_js}];
              const gapMs = {int(gap_ms)};
              const playOne = (delayMs, freq, durMs) => {{
                setTimeout(() => {{
                  const ctx = new AudioCtx();
                  const osc = ctx.createOscillator();
                  const gain = ctx.createGain();
                  osc.type = "sine";
                  osc.frequency.value = freq;
                  gain.gain.value = 0.045;
                  osc.connect(gain);
                  gain.connect(ctx.destination);
                  osc.start();
                  osc.stop(ctx.currentTime + (durMs / 1000));
                  osc.onended = () => ctx.close();
                }}, delayMs);
              }};
              let cursor = 0;
              for (const tone of tones) {{
                playOne(cursor, tone.freq, tone.durMs);
                cursor += tone.durMs + gapMs;
              }}
            }})();
            """
            display(Javascript(js))
            return
    except Exception:
        pass

    for index, _ in enumerate(tone_plan):
        print("\a", end="", flush=True)
        if gap_ms > 0 and index < len(tone_plan) - 1:
            time.sleep(gap_ms / 1000.0)
    print("", flush=True)


def archive_config_dir_once_per_day(
    config_dir: Path | None = None,
    archive_root: Path | None = None,
    *,
    today: date | None = None,
) -> Path | None:
    """Archive the config folder once per day, skipping if already archived."""
    config_dir = (config_dir or (REPO_ROOT / "config")).resolve()
    archive_root = (archive_root or (config_dir / ".archive")).resolve()
    date_token = (today or date.today()).strftime("%Y%m%d")
    daily_dir = archive_root / date_token
    daily_dir.mkdir(parents=True, exist_ok=True)

    existing = sorted(daily_dir.glob("config_*.zip"))
    if existing:
        return existing[0]

    archive_path = daily_dir / f"config_{date_token}.zip"
    base_dir = config_dir
    skip_dir = archive_root
    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in base_dir.rglob("*"):
            if path.is_dir():
                continue
            if skip_dir in path.parents:
                continue
            # Excel lock files (for open workbooks) are temporary and frequently unreadable.
            if path.name.startswith("~$"):
                continue
            rel_path = path.relative_to(base_dir)
            try:
                zf.write(path, arcname=str(rel_path))
            except PermissionError:
                print(f"[WARN] Skipping unreadable config file during archive: {path}")
                continue
    return archive_path


def parse_notebook_safe_args(
    parser: argparse.ArgumentParser,
    argv: Sequence[str] | None = None,
) -> argparse.Namespace:
    """Parse CLI args while tolerating Jupyter kernel connection-file flags."""
    if argv is None:
        argv = sys.argv[1:]

    filtered_args: list[str] = []
    skip_next = False
    for token in argv:
        if skip_next:
            skip_next = False
            continue
        if token == "-f":
            skip_next = True
            continue
        if str(token).startswith("--f="):
            continue
        filtered_args.append(str(token))
    return parser.parse_args(filtered_args)


def normalize_economies(economies: str | Iterable[str] | None) -> list[str]:
    """Return a normalized list of economy labels."""
    if economies is None:
        return []
    if isinstance(economies, str):
        text = economies.strip()
        return [text] if text else []
    return [str(value).strip() for value in economies if str(value).strip()]


def resolve_aggregate_economy(
    economies: str | Iterable[str] | None,
    aggregate_label: str | None = None,
    *,
    aggregate_labels: set[str] | None = None,
) -> tuple[bool, str, list[str]]:
    """Return (should_aggregate, aggregate_label, normalized_economies)."""
    normalized = normalize_economies(economies)
    labels = aggregate_labels or AGGREGATE_ECONOMY_LABELS
    if len(normalized) == 1 and normalized[0] in labels:
        return True, normalized[0], normalized
    resolved_label = aggregate_label or "ALL_ECONOMIES"
    return False, resolved_label, normalized

def format_filename_segment(value: str | None) -> str:
    """Return a file-safe string for economy or scenario labels."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    sanitized = re.sub(r"[^A-Za-z0-9_-]+", "_", text)
    return sanitized.strip("_") or text


def normalize_scenarios(scenarios: str | Iterable[str] | None) -> list[str]:
    """Return a list of scenario labels."""
    if scenarios is None:
        return []
    if isinstance(scenarios, str):
        return [scenarios]
    return list(scenarios)


def normalize_workflow_scenarios(
    scenarios: str | Iterable[str] | None,
    default_scenarios: Sequence[str],
) -> list[str]:
    """Return cleaned scenario names for export/import workflow operations."""
    if scenarios is None:
        scenario_values = list(default_scenarios)
    elif isinstance(scenarios, str):
        scenario_values = [scenarios]
    else:
        scenario_values = list(scenarios)
    cleaned: list[str] = []
    seen: set[str] = set()
    for value in scenario_values:
        scenario_name = str(value).strip()
        if not scenario_name or scenario_name in seen:
            continue
        seen.add(scenario_name)
        cleaned.append(scenario_name)
    return cleaned or list(default_scenarios)


def resolve_import_scenarios(
    scenario_list: Sequence[str],
    import_scenario: str | Sequence[str] | None,
    *,
    current_accounts_labels: set[str] | None = None,
) -> list[str]:
    """Return ordered scenario names to import, excluding current-accounts labels."""
    account_labels = current_accounts_labels or {"current accounts", "current account"}
    available_by_lower = {str(name).strip().lower(): str(name) for name in scenario_list}
    default_scenarios = [
        scenario
        for scenario in scenario_list
        if str(scenario).strip().lower() not in account_labels
    ]
    if import_scenario is None:
        if not default_scenarios:
            raise ValueError(
                f"No non-'Current Accounts' scenarios available for import in {list(scenario_list)}."
            )
        return list(default_scenarios)

    if isinstance(import_scenario, str):
        requested_values = [import_scenario]
    else:
        requested_values = list(import_scenario)

    resolved: list[str] = []
    for value in requested_values:
        scenario_name = str(value).strip()
        if not scenario_name:
            continue
        scenario_key = scenario_name.lower()
        if scenario_key in account_labels:
            continue
        if scenario_key not in available_by_lower:
            raise ValueError(
                f"Import scenario '{scenario_name}' is not in exported scenarios: {list(scenario_list)}"
            )
        matched = available_by_lower[scenario_key]
        if matched not in resolved:
            resolved.append(matched)
    if not resolved:
        if not default_scenarios:
            raise ValueError(
                f"No non-'Current Accounts' scenarios available for import in {list(scenario_list)}."
            )
        return list(default_scenarios)
    return resolved


def _format_scenario_segment(
    scenarios: Sequence[str],
    format_segment_fn: Callable[[str], str],
) -> str:
    tokens = [format_segment_fn(segment) for segment in scenarios if segment]
    sanitized = "_".join(token for token in tokens if token)
    return sanitized or "scenarios"


def format_export_filename(
    economy_label: str,
    scenarios: Sequence[str],
    template: str,
    format_segment_fn: Callable[[str], str],
    fallback_template: str | None = None,
) -> str:
    """Return a safe filename for export workbooks."""
    scenario_segment = _format_scenario_segment(scenarios, format_segment_fn)
    economy_segment = format_segment_fn(economy_label)
    try:
        return template.format(economy=economy_segment, scenario=scenario_segment)
    except Exception as exc:
        print(f"Failed to format export filename: {exc}")
        fallback = fallback_template or template
        try:
            return fallback.format(economy=economy_segment, scenario=scenario_segment)
        except Exception:
            return fallback


def build_workflow_export_filename(
    economy_label: str,
    scenarios: str | Iterable[str] | None,
    template: str,
    format_segment_fn: Callable[[str], str] = format_filename_segment,
    fallback_template: str | None = None,
) -> str:
    """Return a filename that includes economy and scenario(s)."""
    scenario_list = normalize_scenarios(scenarios)
    return format_export_filename(
        economy_label,
        scenario_list,
        template,
        format_segment_fn,
        fallback_template=fallback_template,
    )


def read_export_column_values(
    export_path: Path,
    sheet_name: str,
    column: str,
) -> list[str]:
    """Return unique values in a column while preserving order."""
    for header in (2, 0):
        try:
            df = pd.read_excel(
                export_path, sheet_name=sheet_name, header=header, usecols=[column]
            )
        except Exception:
            continue
        if column not in df.columns:
            continue
        seen: list[str] = []
        for value in df[column].dropna().astype(str):
            if value not in seen:
                seen.append(value)
        if seen:
            return seen
    return []


def list_export_scenarios(export_path: Path, sheet_name: str) -> list[str]:
    """Return the Scenario column values in declaration order."""
    return read_export_column_values(export_path, sheet_name, "Scenario")


def validate_export_region(export_path: Path, sheet_name: str, region: str) -> None:
    """Ensure the workbook contains the requested region."""
    regions = read_export_column_values(export_path, sheet_name, "Region")
    if not regions:
        print(f"Warning: 'Region' column missing from {export_path.name}; skipping region check.")
        return
    if region not in regions:
        raise ValueError(
            f"Requested region '{region}' not present in {export_path.name}; available: {regions}"
        )


def find_latest_export_workbook(
    directory: Path | str,
    prefix: str,
    filename: str | None = None,
) -> Path:
    """Locate a workbook by explicit name or latest matching prefix."""
    directory_path = Path(directory)
    if filename:
        candidate = directory_path / filename
        if candidate.exists():
            return candidate
        raise FileNotFoundError(f"Specified export missing: {candidate}")
    matches = sorted(directory_path.glob(f"{prefix}*.xlsx"))
    if not matches:
        raise FileNotFoundError(f"No exports detected in {directory_path}")
    return matches[-1]


def import_workbook_to_leap(
    export_path: Path,
    sheet_name: str,
    scenario: str | None,
    region: str | None,
    create_branches: bool = True,
    fill_branches: bool = True,
    include_current_accounts: bool = True,
    default_branch_type: tuple | None = None,
    branch_type_mapping: dict | None = None,
    branch_root: str | None = None,
    branch_path_col: str | None = None,
    raise_on_missing_branch: bool = False,
) -> Path:
    """Connect to LEAP, validate the workbook, and fill branches."""
    available = list_export_scenarios(export_path, sheet_name)
    scenario_choice = scenario or (available[0] if available else None)
    available_lower = {str(name).strip().lower() for name in available}
    current_accounts_available = any(
        label in available_lower for label in {"current accounts", "current account"}
    )
    if scenario_choice and scenario_choice not in available:
        raise ValueError(
            f"Scenario '{scenario_choice}' not found in {export_path.name}; options {available}"
        )
    if region:
        validate_export_region(export_path, sheet_name, region)
    if include_current_accounts and not current_accounts_available:
        print(
            "[INFO] Skipping Current Accounts import for "
            f"{export_path.name}: workbook scenarios are {available}."
        )
        include_current_accounts = False

    def _run_api_write() -> Path:
        leap_conn = connect_to_leap()
        if leap_conn is None:
            raise RuntimeError("Unable to connect to LEAP.")

        fuel_catalog_preflight.run_fuel_catalog_preflight(
            export_path=export_path,
            sheet_name=sheet_name,
            scenario=scenario_choice,
            context="workflow_common.import_workbook_to_leap",
            leap_app=leap_conn,
        )
        if create_branches:
            create_kwargs = {
                "sheet_name": sheet_name,
                "branch_root": branch_root,
                "branch_type_mapping": branch_type_mapping,
                "default_branch_type": default_branch_type,
                "RAISE_ERROR_ON_FAILED_BRANCH_CREATION": raise_on_missing_branch,
            }
            if branch_path_col is not None:
                create_kwargs["branch_path_col"] = branch_path_col
            create_branches_from_export_file(
                leap_conn,
                export_path,
                **create_kwargs,
            )
        if fill_branches:
            fill_branches_from_export_file(
                leap_conn,
                export_path,
                sheet_name=sheet_name,
                scenario=scenario_choice,
                region=region,
                RAISE_ERROR_ON_FAILED_SET=raise_on_missing_branch,
                HANDLE_CURRENT_ACCOUNTS_TOO=include_current_accounts,
                RUN_FUEL_CATALOG_PREFLIGHT=False,
            )
        return export_path

    dispatch_result = dispatch_analysis_input_write(
        export_path=export_path,
        sheet_name=sheet_name,
        scenario=scenario_choice,
        region=region,
        context_label="workflow_common.import_workbook_to_leap",
        run_api_write=_run_api_write,
    )
    if dispatch_result.get("mode") == "api":
        result_path = dispatch_result.get("api_result")
        if isinstance(result_path, Path):
            return result_path
    return export_path
