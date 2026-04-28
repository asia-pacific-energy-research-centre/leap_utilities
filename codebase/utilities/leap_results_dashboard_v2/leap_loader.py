from __future__ import annotations

from pathlib import Path

import pandas as pd

from codebase.utilities.leap_results_dashboard_utils import load_leap_workbook


def discover_workbooks(root: Path, economy_token: str, scenarios: tuple[str, ...]) -> list[Path]:
    if not root.exists():
        raise FileNotFoundError(f"LEAP results directory not found: {root}")
    economy = economy_token.lower().strip()
    scen = [s.lower().strip() for s in scenarios]
    out: list[Path] = []
    for path in sorted(root.glob("*.xls*")):
        name = path.name.lower()
        if path.name.startswith("~$"):
            continue
        if economy in name and any(s in name for s in scen):
            out.append(path)
    if not out:
        raise FileNotFoundError(
            f"No LEAP workbooks found in {root} matching economy '{economy_token}' and scenarios {scenarios}."
        )
    # Always ignore legacy combined transformation+supply workbooks.
    # The workflow now relies on dedicated transformation_results_* and supply_results_* files.
    out = [p for p in out if "transformation_and_supply_results_" not in p.name.lower()]
    return out


def _read_additional_leap_long_csvs(paths: list[Path] | None) -> list[pd.DataFrame]:
    frames: list[pd.DataFrame] = []
    for path in paths or []:
        candidate = Path(path)
        if not candidate.exists():
            continue
        frame = pd.read_csv(candidate)
        if frame.empty:
            continue
        rename_map = {"sheet": "sheet_name", "value": "leap_value", "unit": "leap_units"}
        frame = frame.rename(columns={k: v for k, v in rename_map.items() if k in frame.columns})
        for col in [
            "economy",
            "scenario",
            "region",
            "sheet_name",
            "sector_code_9th",
            "sector_name",
            "fuel_label",
            "year",
            "leap_value",
            "leap_variable",
            "leap_units",
            "measure",
            "leap_scale_note",
        ]:
            if col not in frame.columns:
                frame[col] = ""
        frames.append(frame)
    return frames


def load_leap_long(
    workbooks: list[Path],
    sheet_map: pd.DataFrame,
    additional_long_paths: list[Path] | None = None,
) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []

    def _scenario_from_workbook_name(path: Path) -> str | None:
        name = path.name.lower()
        if "reference" in name:
            return "Reference"
        if "target" in name:
            return "Target"
        return None

    for wb in workbooks:
        expected_scenario = _scenario_from_workbook_name(wb)
        frame = load_leap_workbook(wb, sheet_map=sheet_map, expected_scenario=expected_scenario)
        if frame.empty:
            continue
        frame = frame.copy()
        frame["_source_workbook"] = wb.name
        frames.append(frame)
    extra_frames = _read_additional_leap_long_csvs(additional_long_paths)
    for idx, frame in enumerate(extra_frames, start=1):
        frame = frame.copy()
        frame["_source_workbook"] = f"additional_long_{idx}"
        frames.append(frame)
    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames, ignore_index=True)

    key_cols = ["economy", "scenario", "sheet_name", "fuel_label", "year"]
    out["leap_value"] = pd.to_numeric(out["leap_value"], errors="coerce")
    overlap = (
        out.groupby(key_cols, dropna=False)
        .agg(
            row_count=("leap_value", "size"),
            unique_non_null=("leap_value", lambda s: pd.to_numeric(s, errors="coerce").dropna().nunique()),
        )
        .reset_index()
    )
    conflicting = overlap[(overlap["row_count"] > 1) & (overlap["unique_non_null"] > 1)].copy()
    if not conflicting.empty:
        merged = out.merge(conflicting[key_cols], on=key_cols, how="inner")
        sample = (
            merged[
                key_cols
                + ["leap_value", "_source_workbook"]
            ]
            .sort_values(key_cols + ["_source_workbook"])
            .head(20)
            .to_dict("records")
        )
        raise RuntimeError(
            "Conflicting LEAP values detected across overlapping workbooks for the same "
            f"economy/scenario/sheet/fuel/year key. Conflicts: {len(conflicting)}. Examples: {sample}"
        )

    out = out.sort_values(key_cols + ["_source_workbook"], kind="mergesort")
    out = out.drop_duplicates(subset=key_cols, keep="first").drop(columns=["_source_workbook"], errors="ignore")
    return out.reset_index(drop=True)
