#%%
from __future__ import annotations

import json
from pathlib import Path
from typing import Callable

import pandas as pd

from codebase.configuration import workflow_config as workflow_cfg
from codebase.functions.leap_api_guard import ensure_leap_api_allowed
from codebase.utilities.master_config import config_table_exists, read_config_table

DEFAULT_ANALYSIS_INPUT_WRITE_MODE = "api"
VALID_ANALYSIS_INPUT_WRITE_MODES = {"api", "workbook"}
REQUIRED_WORKBOOK_COLUMNS = [
    "Branch Path",
    "Variable",
    "Scenario",
    "Region",
    "Scale",
    "Units",
    "Per...",
]
REQUIRED_KEY_COLUMNS = [
    "Branch Path",
    "Variable",
    "Scenario",
    "Region",
]
MAPPING_FIELD_COLUMNS = {
    "units": "Units",
    "scale": "Scale",
    "per": "Per...",
}


def _clean_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def _normalize_header_value(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and float(value).is_integer():
        return str(int(value))
    return str(value)


def _as_bool(value: object) -> bool:
    text = _clean_text(value).lower()
    return text in {"1", "true", "yes", "y", "on"}


def _year_columns(columns: list[str]) -> list[str]:
    return [
        str(col)
        for col in columns
        if str(col).strip().isdigit() and len(str(col).strip()) == 4
    ]


def get_analysis_input_write_mode() -> str:
    configured = getattr(
        workflow_cfg,
        "ANALYSIS_INPUT_WRITE_MODE",
        DEFAULT_ANALYSIS_INPUT_WRITE_MODE,
    )
    mode = _clean_text(configured).lower() or DEFAULT_ANALYSIS_INPUT_WRITE_MODE
    if mode not in VALID_ANALYSIS_INPUT_WRITE_MODES:
        raise ValueError(
            "Invalid ANALYSIS_INPUT_WRITE_MODE="
            f"{configured!r}. Valid values: {sorted(VALID_ANALYSIS_INPUT_WRITE_MODES)}"
        )
    return mode


def is_workbook_mode() -> bool:
    return get_analysis_input_write_mode() == "workbook"


def ensure_api_write_allowed(context_label: str) -> None:
    if is_workbook_mode():
        raise RuntimeError(
            f"Analysis-view API writes are disabled in workbook mode. "
            f"Blocked write context: {context_label}. "
            "Generate/validate workbook and import manually into LEAP, then recalculate."
        )
    ensure_leap_api_allowed(context_label)


def ensure_analysis_view_api_read_allowed(context_label: str) -> None:
    if is_workbook_mode():
        raise RuntimeError(
            f"Analysis-view API reads are disabled in workbook mode. "
            f"Blocked read context: {context_label}. "
            "Use workbook generation + manual import for Analysis updates. "
            "LEAP API is results-extraction-only in workbook mode."
        )
    ensure_leap_api_allowed(context_label)


def ensure_analysis_view_api_access_allowed(context_label: str, *, access_kind: str = "read/write") -> None:
    if is_workbook_mode():
        raise RuntimeError(
            f"Analysis-view API {access_kind} is disabled in workbook mode. "
            f"Blocked context: {context_label}. "
            "Use workbook generation + manual import for Analysis updates. "
            "LEAP API is results-extraction-only in workbook mode."
        )
    ensure_leap_api_allowed(context_label)


def _find_header_row(raw: pd.DataFrame) -> int:
    for idx in range(len(raw.index)):
        tokens = {_clean_text(value).lower() for value in raw.iloc[idx].tolist()}
        if "branch path" in tokens and "variable" in tokens:
            return int(idx)
    raise ValueError("Could not locate workbook header row containing 'Branch Path' and 'Variable'.")


def _read_workbook_sheet(path: Path, sheet_name: str) -> tuple[pd.DataFrame, int, list[str], pd.DataFrame]:
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
    header_row = _find_header_row(raw)
    columns = [_normalize_header_value(value) for value in raw.iloc[header_row].tolist()]
    data = raw.iloc[header_row + 1 :].copy()
    data.columns = columns
    data = data.reset_index(drop=True)
    return raw, header_row, columns, data


def _write_workbook_sheet(
    path: Path,
    sheet_name: str,
    original_raw: pd.DataFrame,
    header_row: int,
    columns: list[str],
    data: pd.DataFrame,
) -> None:
    width = len(columns)
    preamble = original_raw.iloc[:header_row, :width].copy()
    preamble.columns = list(range(width))
    header_df = pd.DataFrame([columns], columns=list(range(width)))
    body = data.reindex(columns=columns).copy()
    body.columns = list(range(width))
    output = pd.concat([preamble, header_df, body], ignore_index=True)
    writer_kwargs = {"engine": "openpyxl", "mode": "a", "if_sheet_exists": "replace"}
    if not path.exists():
        writer_kwargs = {"engine": "openpyxl", "mode": "w"}
    with pd.ExcelWriter(path, **writer_kwargs) as writer:
        output.to_excel(writer, sheet_name=sheet_name, index=False, header=False)


def _canonical_source_paths() -> list[Path]:
    configured = list(
        getattr(
            workflow_cfg,
            "ANALYSIS_INPUT_CANONICAL_TEMPLATE_PATHS",
            [],
        )
    )
    resolved: list[Path] = []
    for item in configured:
        path = Path(str(item)).expanduser().resolve()
        if path.exists():
            resolved.append(path)
    return resolved


def _iter_canonical_workbooks() -> list[Path]:
    workbooks: list[Path] = []
    for source in _canonical_source_paths():
        if source.is_dir():
            workbooks.extend(sorted(source.glob("*.xlsx")))
        elif source.suffix.lower() == ".xlsx":
            workbooks.append(source)
    deduped: list[Path] = []
    seen: set[str] = set()
    for item in workbooks:
        token = str(item.resolve())
        if token in seen:
            continue
        seen.add(token)
        deduped.append(item)
    return deduped


def _choose_lookup_value(values: set[str]) -> str | None:
    if not values:
        return None
    normalized = {_clean_text(value) for value in values}
    non_empty = sorted(value for value in normalized if value)
    if len(non_empty) == 1:
        return non_empty[0]
    if len(non_empty) > 1:
        return None
    return ""


def _register_lookup(table: dict[tuple[str, str, str], set[str]], key: tuple[str, str, str], value: str) -> None:
    table.setdefault(key, set()).add(value)


def _build_template_lookup_catalog() -> dict[str, dict[tuple[str, str, str], set[str]]]:
    branch_variable: dict[tuple[str, str, str], set[str]] = {}
    variable: dict[tuple[str, str, str], set[str]] = {}
    branch: dict[tuple[str, str, str], set[str]] = {}
    for workbook in _iter_canonical_workbooks():
        sheet_candidates = ["LEAP", "Export"]
        try:
            available_sheets = set(pd.ExcelFile(workbook).sheet_names)
        except Exception:
            continue
        for sheet_name in sheet_candidates:
            if sheet_name not in available_sheets:
                continue
            try:
                _, _, _, frame = _read_workbook_sheet(workbook, sheet_name)
            except Exception:
                continue
            required = {"Branch Path", "Variable"}
            if not required.issubset(set(frame.columns)):
                continue
            for _, row in frame.iterrows():
                branch_path = _clean_text(row.get("Branch Path"))
                variable_name = _clean_text(row.get("Variable"))
                if not branch_path or not variable_name:
                    continue
                for field_key, field_col in MAPPING_FIELD_COLUMNS.items():
                    value = _clean_text(row.get(field_col))
                    _register_lookup(
                        branch_variable,
                        (branch_path, variable_name, field_key),
                        value,
                    )
                    _register_lookup(
                        variable,
                        (variable_name, "", field_key),
                        value,
                    )
                    _register_lookup(
                        branch,
                        (branch_path, "", field_key),
                        value,
                    )
    return {
        "branch_variable": branch_variable,
        "variable": variable,
        "branch": branch,
    }


def _load_canonical_structures() -> list[dict[str, object]]:
    structures: list[dict[str, object]] = []
    for workbook in _iter_canonical_workbooks():
        try:
            sheet_names = set(pd.ExcelFile(workbook).sheet_names)
        except Exception:
            continue
        for sheet_name in ("LEAP", "Export"):
            if sheet_name not in sheet_names:
                continue
            try:
                _, header_row, columns, _ = _read_workbook_sheet(workbook, sheet_name)
            except Exception:
                continue
            structures.append(
                {
                    "workbook": str(workbook),
                    "sheet_name": sheet_name,
                    "header_row_index": int(header_row),
                    "columns": list(columns),
                }
            )
    return structures


def _validate_canonical_structure_available(canonical_structures: list[dict[str, object]]) -> None:
    if not canonical_structures:
        raise ValueError(
            "No readable canonical template workbook/sheet found. "
            "Check ANALYSIS_INPUT_CANONICAL_TEMPLATE_PATHS."
        )
    if not any(
        set(REQUIRED_WORKBOOK_COLUMNS).issubset(set(item.get("columns", [])))
        for item in canonical_structures
    ):
        raise ValueError(
            "Canonical template sources do not expose required LEAP columns "
            f"{REQUIRED_WORKBOOK_COLUMNS}."
        )


def _validate_workbook_structure_against_canonical(
    *,
    workbook_columns: list[str],
    canonical_structures: list[dict[str, object]],
    workbook_path: Path,
    sheet_name: str,
) -> None:
    _validate_canonical_structure_available(canonical_structures)
    missing_required = [
        col for col in REQUIRED_WORKBOOK_COLUMNS if col not in set(workbook_columns)
    ]
    if missing_required:
        raise ValueError(
            f"Workbook '{workbook_path.name}' sheet '{sheet_name}' missing required columns: "
            f"{missing_required}"
        )
    missing_key = [col for col in REQUIRED_KEY_COLUMNS if col not in set(workbook_columns)]
    if missing_key:
        raise ValueError(
            f"Workbook '{workbook_path.name}' sheet '{sheet_name}' missing required key columns: "
            f"{missing_key}"
        )


def _mapping_file_path() -> Path:
    configured = getattr(
        workflow_cfg,
        "ANALYSIS_INPUT_FIELD_MAPPING_PATH",
        Path("config/leap_export_workbook_mappings.xlsx"),
    )
    return Path(str(configured)).expanduser().resolve()


def _mapping_sheet_name() -> str:
    configured = getattr(
        workflow_cfg,
        "ANALYSIS_INPUT_FIELD_MAPPING_SHEET",
        "field_mappings",
    )
    return _clean_text(configured) or "field_mappings"


def _load_mapping_table(mapping_path: Path | None = None, mapping_sheet: str | None = None) -> pd.DataFrame:
    path = mapping_path or _mapping_file_path()
    sheet_name = mapping_sheet or _mapping_sheet_name()
    if not config_table_exists(path, sheet_name):
        return pd.DataFrame(
            columns=[
                "enabled",
                "match_scope",
                "branch_path",
                "variable",
                "units",
                "scale",
                "per",
                "confidence",
                "notes",
            ]
        )
    table = read_config_table(path, sheet_name=sheet_name)
    table.columns = [str(col).strip().lower() for col in table.columns]
    required_cols = {
        "enabled",
        "match_scope",
        "branch_path",
        "variable",
        "units",
        "scale",
        "per",
        "confidence",
        "notes",
    }
    missing = sorted(required_cols.difference(table.columns))
    if missing:
        raise ValueError(
            f"Mapping file '{path}' sheet '{sheet_name}' is missing required columns: {missing}"
        )
    for col in required_cols:
        if col not in table.columns:
            table[col] = ""
    return table


def _build_mapping_indexes(mapping_table: pd.DataFrame) -> dict[str, dict[tuple[str, str], dict[str, str]]]:
    indexes: dict[str, dict[tuple[str, str], dict[str, str]]] = {
        "branch_variable": {},
        "variable": {},
        "branch": {},
    }
    for _, row in mapping_table.iterrows():
        if not _as_bool(row.get("enabled")):
            continue
        scope = _clean_text(row.get("match_scope")).lower()
        if scope not in indexes:
            continue
        branch_path = _clean_text(row.get("branch_path"))
        variable_name = _clean_text(row.get("variable"))
        if scope == "branch_variable":
            key = (branch_path, variable_name)
            if not key[0] or not key[1]:
                continue
        elif scope == "variable":
            key = (variable_name, "")
            if not key[0]:
                continue
        else:
            key = (branch_path, "")
            if not key[0]:
                continue
        payload = {
            "units": _clean_text(row.get("units")),
            "scale": _clean_text(row.get("scale")),
            "per": _clean_text(row.get("per")),
            "confidence": _clean_text(row.get("confidence")).lower() or "known",
            "notes": _clean_text(row.get("notes")),
            "scope": scope,
        }
        if key in indexes[scope]:
            raise ValueError(
                f"Duplicate enabled mapping rows detected for scope='{scope}' key={key}."
            )
        indexes[scope][key] = payload
    return indexes


def _match_mapping_entry(
    indexes: dict[str, dict[tuple[str, str], dict[str, str]]],
    branch_path: str,
    variable_name: str,
) -> dict[str, str] | None:
    branch_variable = indexes["branch_variable"].get((branch_path, variable_name))
    if branch_variable is not None:
        return branch_variable
    variable = indexes["variable"].get((variable_name, ""))
    if variable is not None:
        return variable
    branch = indexes["branch"].get((branch_path, ""))
    if branch is not None:
        return branch
    return None


def _lookup_template_value(
    lookups: dict[str, dict[tuple[str, str, str], set[str]]],
    branch_path: str,
    variable_name: str,
    field_key: str,
) -> tuple[str | None, str]:
    candidate = _choose_lookup_value(
        lookups["branch_variable"].get((branch_path, variable_name, field_key), set())
    )
    if candidate is not None:
        return candidate, "template_branch_variable"
    candidate = _choose_lookup_value(
        lookups["variable"].get((variable_name, "", field_key), set())
    )
    if candidate is not None:
        return candidate, "template_variable"
    candidate = _choose_lookup_value(
        lookups["branch"].get((branch_path, "", field_key), set())
    )
    if candidate is not None:
        return candidate, "template_branch"
    return None, ""


def validate_workbook_for_manual_import(
    export_path: Path | str,
    *,
    sheet_name: str = "LEAP",
    scenario: str | None = None,
    region: str | None = None,
    mapping_path: Path | str | None = None,
    mapping_sheet: str | None = None,
) -> dict[str, object]:
    path = Path(export_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Workbook not found: {path}")

    raw, header_row, columns, data = _read_workbook_sheet(path, sheet_name)

    canonical_structures = _load_canonical_structures()
    _validate_workbook_structure_against_canonical(
        workbook_columns=columns,
        canonical_structures=canonical_structures,
        workbook_path=path,
        sheet_name=sheet_name,
    )

    year_cols = _year_columns(columns)
    has_expression = "Expression" in columns
    if not has_expression and not year_cols:
        raise ValueError(
            f"Workbook '{path.name}' sheet '{sheet_name}' must include 'Expression' or year columns."
        )

    filtered = data.copy()
    if scenario is not None and "Scenario" in filtered.columns:
        filtered = filtered[
            filtered["Scenario"].astype(str).str.strip().str.lower()
            == _clean_text(scenario).lower()
        ].copy()
    if region is not None and "Region" in filtered.columns:
        filtered = filtered[
            filtered["Region"].astype(str).str.strip().str.lower()
            == _clean_text(region).lower()
        ].copy()

    key_dupes = filtered.duplicated(subset=REQUIRED_KEY_COLUMNS, keep=False)
    if key_dupes.any():
        duplicate_count = int(key_dupes.sum())
        raise ValueError(
            f"Workbook '{path.name}' has duplicate key rows (Branch Path, Variable, Scenario, Region): {duplicate_count}"
        )

    mapping_table = _load_mapping_table(
        mapping_path=(Path(mapping_path).expanduser().resolve() if mapping_path else None),
        mapping_sheet=mapping_sheet,
    )
    mapping_indexes = _build_mapping_indexes(mapping_table)
    template_lookups = _build_template_lookup_catalog()

    modified = False
    defaults_applied: list[dict[str, str]] = []
    confidently_known: list[dict[str, str]] = []
    inferred: list[dict[str, str]] = []
    needs_confirmation: list[dict[str, str]] = []

    for idx, row in data.iterrows():
        branch_path = _clean_text(row.get("Branch Path"))
        variable_name = _clean_text(row.get("Variable"))
        if not branch_path or not variable_name:
            continue
        mapping_entry = _match_mapping_entry(mapping_indexes, branch_path, variable_name)
        for field_key, field_col in MAPPING_FIELD_COLUMNS.items():
            current_value = _clean_text(row.get(field_col))
            resolved_value = current_value
            source = "existing"
            confidence = "known"
            if mapping_entry is not None:
                mapped_value = _clean_text(mapping_entry.get(field_key))
                source = f"mapping_{_clean_text(mapping_entry.get('scope'))}"
                confidence = _clean_text(mapping_entry.get("confidence")).lower() or "known"
                if mapped_value:
                    resolved_value = mapped_value
                elif field_key in {"scale", "per"}:
                    resolved_value = ""
                else:
                    needs_confirmation.append(
                        {
                            "branch_path": branch_path,
                            "variable": variable_name,
                            "field": field_col,
                            "reason": "Mapping row matched but required field value is blank.",
                        }
                    )
                    continue
            elif current_value:
                resolved_value = current_value
                source = "existing"
                confidence = "known"
            else:
                template_value, template_source = _lookup_template_value(
                    template_lookups,
                    branch_path,
                    variable_name,
                    field_key,
                )
                if template_value is not None:
                    resolved_value = template_value
                    source = template_source
                    confidence = "inferred"
                elif field_key in {"scale", "per"}:
                    resolved_value = ""
                    source = "default_blank"
                    confidence = "known"
                    defaults_applied.append(
                        {
                            "branch_path": branch_path,
                            "variable": variable_name,
                            "field": field_col,
                            "value": "",
                        }
                    )
                else:
                    needs_confirmation.append(
                        {
                            "branch_path": branch_path,
                            "variable": variable_name,
                            "field": field_col,
                            "reason": "No mapping, existing value, or canonical template fallback.",
                        }
                    )
                    continue

            if confidence == "needs_confirmation":
                needs_confirmation.append(
                    {
                        "branch_path": branch_path,
                        "variable": variable_name,
                        "field": field_col,
                        "reason": "Mapping row marked needs_confirmation.",
                    }
                )
                continue

            if resolved_value != current_value:
                data.at[idx, field_col] = resolved_value
                modified = True

            record = {
                "branch_path": branch_path,
                "variable": variable_name,
                "field": field_col,
                "value": resolved_value,
                "source": source,
            }
            if confidence == "known":
                confidently_known.append(record)
            else:
                inferred.append(record)

    if needs_confirmation:
        sample = needs_confirmation[:8]
        raise ValueError(
            "Workbook-mode metadata resolution failed. "
            f"Fields needing confirmation: {sample}"
        )

    if modified:
        _write_workbook_sheet(
            path=path,
            sheet_name=sheet_name,
            original_raw=raw,
            header_row=header_row,
            columns=columns,
            data=data,
        )

    required_columns_populated = [col for col in REQUIRED_WORKBOOK_COLUMNS if col in columns]
    extra_columns_populated = [
        col
        for col in columns
        if col not in REQUIRED_WORKBOOK_COLUMNS
        and col in data.columns
        and data[col].astype(str).str.strip().ne("").any()
    ]
    units_used = sorted(
        {
            _clean_text(value)
            for value in data.get("Units", pd.Series(dtype=object)).tolist()
            if _clean_text(value)
        }
    )

    return {
        "mode": "workbook",
        "workbook_path": str(path),
        "sheet_name": sheet_name,
        "scenario": _clean_text(scenario),
        "region": _clean_text(region),
        "header_row_index": int(header_row),
        "row_count": int(len(data)),
        "required_columns_populated": required_columns_populated,
        "extra_columns_populated": extra_columns_populated,
        "units_used": units_used,
        "defaults_applied": defaults_applied,
        "confidently_known_fields": confidently_known,
        "inferred_fields": inferred,
        "fields_needing_confirmation": needs_confirmation,
        "mapping_file_used": str((mapping_path or _mapping_file_path())),
        "mapping_sheet_used": mapping_sheet or _mapping_sheet_name(),
        "canonical_sources_used": [str(path) for path in _iter_canonical_workbooks()],
        "canonical_structures": canonical_structures,
        "modified_workbook": bool(modified),
        "has_expression_column": bool(has_expression),
        "year_columns": year_cols,
    }


def emit_workbook_mode_summary(
    summary: dict[str, object],
    *,
    summary_path: Path | str | None = None,
) -> Path:
    workbook_path = Path(str(summary.get("workbook_path", ""))).resolve()
    output_path = (
        Path(summary_path).resolve()
        if summary_path is not None
        else Path(str(workbook_path) + ".analysis_input_write_summary.json")
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as handle:
        json.dump(summary, handle, indent=2, ensure_ascii=False)
    return output_path


def _print_visible_workbook_mode_messages(
    context_label: str,
    workbook_path: Path,
    summary_path: Path,
) -> None:
    print("\n" + "=" * 92)
    print("[WORKBOOK MODE] Analysis-view API writes are disabled.")
    print(f"[WORKBOOK MODE] Context: {context_label}")
    print(f"[WORKBOOK MODE] Workbook path: {workbook_path}")
    print("[WORKBOOK MODE] Workbook generation/validation succeeded.")
    print(
        "[WORKBOOK MODE] Manual LEAP action required: "
        "Import this workbook into LEAP (Analysis view), then recalculate."
    )
    print(
        "[WORKBOOK MODE][WARNING] LEAP results will NOT reflect these changes "
        "until manual import + recalculation are completed."
    )
    print(f"[WORKBOOK MODE] Summary JSON: {summary_path}")
    print("=" * 92 + "\n")


def _print_structured_summary(summary: dict[str, object]) -> None:
    print("[WORKBOOK MODE] Structured summary:")
    print(f"  - units_used: {summary.get('units_used', [])}")
    print(
        "  - required_columns_populated: "
        f"{summary.get('required_columns_populated', [])}"
    )
    print(
        "  - extra_columns_populated: "
        f"{summary.get('extra_columns_populated', [])}"
    )
    print(f"  - defaults_applied: {len(summary.get('defaults_applied', []))}")
    print(
        "  - confidently_known_fields: "
        f"{len(summary.get('confidently_known_fields', []))}"
    )
    print(f"  - inferred_fields: {len(summary.get('inferred_fields', []))}")
    print(
        "  - fields_needing_confirmation: "
        f"{len(summary.get('fields_needing_confirmation', []))}"
    )


def dispatch_analysis_input_write(
    *,
    export_path: Path | str,
    sheet_name: str = "LEAP",
    scenario: str | None = None,
    region: str | None = None,
    context_label: str,
    run_api_write: Callable[[], object] | None = None,
) -> dict[str, object]:
    mode = get_analysis_input_write_mode()
    workbook_path = Path(export_path).expanduser().resolve()
    if mode == "api":
        result = run_api_write() if callable(run_api_write) else None
        return {
            "mode": "api",
            "workbook_path": str(workbook_path),
            "api_result": result,
            "summary_path": "",
            "summary": None,
        }

    summary = validate_workbook_for_manual_import(
        workbook_path,
        sheet_name=sheet_name,
        scenario=scenario,
        region=region,
    )
    summary_path = emit_workbook_mode_summary(summary)
    _print_visible_workbook_mode_messages(context_label, workbook_path, summary_path)
    _print_structured_summary(summary)
    return {
        "mode": "workbook",
        "workbook_path": str(workbook_path),
        "api_result": None,
        "summary_path": str(summary_path),
        "summary": summary,
    }
