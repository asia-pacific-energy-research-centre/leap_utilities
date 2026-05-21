#%%
"""
Find ESTO flow/product rows with nonzero data that do not have a matching LEAP
branch in a full model export.

The check uses config/leap_mappings.xlsx > leap_combined_esto as the crosswalk
between ESTO rows and LEAP sector/fuel names, then searches the full model
export for branches ending in each mapped LEAP fuel name. It writes both a
complete mapping-row audit and a shorter modeller-facing missing-fuels report.
"""

from __future__ import annotations

import os
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd


#%%
REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


def _resolve(path: Path | str) -> Path:
    raw = str(path).replace("\\", "/")
    drive_match = re.match(r"^([a-zA-Z]):/(.*)$", raw)
    if drive_match:
        drive = drive_match.group(1).lower()
        rest = drive_match.group(2)
        if os.name == "nt":
            return Path(f"{drive.upper()}:/{rest}")
        return Path(f"/mnt/{drive}/{rest}")
    candidate = Path(raw)
    return candidate if candidate.is_absolute() else (REPO_ROOT / candidate)


#%%
FULL_MODEL_EXPORT_PATH = _resolve("data/full model export.xlsx")
FULL_MODEL_EXPORT_SHEET = "Export"
LEAP_TO_ESTO_MAPPING_PATH = _resolve("config/leap_mappings.xlsx")
LEAP_TO_ESTO_MAPPING_SHEET = "leap_combined_esto"
ESTO_DATA_PATH = _resolve("data/00APEC_2025_low_with_subtotals.csv")
OUTPUT_DIR = _resolve("outputs/esto_missing_leap_fuels_from_full_export")

ESTO_ECONOMIES: tuple[str, ...] | None = None  # None means all economies.
ESTO_YEARS: tuple[int, ...] | None = None  # None means all year columns.
NONZERO_ABS_TOLERANCE = 1e-9
EXCLUDE_ESTO_SUBTOTALS = True
EXCLUDE_MAPPING_REMOVE_ROWS_FROM_MISSING_REPORT = True
EXCLUDE_MAPPING_SUBTOTAL_ROWS_FROM_MISSING_REPORT = True
WRITE_DETAILED_AUDIT_OUTPUTS = False
EXCLUDED_SIMPLE_OUTPUT_ESTO_FLOW_PREFIXES = ("01", "02", "03", "06", "07")


#%%
def _clean_text(value: object) -> str:
    text = str(value or "").strip()
    return "" if text.lower() in {"", "nan", "none", "<na>"} else text


def _truthy(value: object) -> bool:
    return _clean_text(value).lower() in {"1", "true", "yes", "y", "on", "t"}


def _normalize_label(value: object) -> str:
    text = _clean_text(value).lower()
    text = text.replace("&", " and ")
    text = text.replace("/", " ")
    text = text.replace("\\", " ")
    text = text.replace("-", " ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def _path_tokens(value: object) -> list[str]:
    text = _clean_text(value).replace("\\", "/")
    return [_normalize_label(part) for part in text.split("/") if _normalize_label(part)]


def _subsequence_in_order(needles: list[str], haystack: list[str]) -> bool:
    """Return True when all needle tokens appear in haystack in order."""
    if not needles:
        return True
    pos = 0
    for token in haystack:
        if token == needles[pos]:
            pos += 1
            if pos == len(needles):
                return True
    return False


def _collapse_consecutive_duplicates(values: list[str]) -> list[str]:
    collapsed: list[str] = []
    for value in values:
        if not collapsed or collapsed[-1] != value:
            collapsed.append(value)
    return collapsed


def _join_unique(values: pd.Series) -> str:
    clean_values = [_clean_text(value) for value in values if _clean_text(value)]
    return "|".join(sorted(set(clean_values)))


def _join_limited(values: pd.Series, limit: int = 12) -> str:
    clean_values = sorted(set(_clean_text(value) for value in values if _clean_text(value)))
    shown = clean_values[:limit]
    if len(clean_values) > limit:
        shown.append(f"... plus {len(clean_values) - limit} more")
    return "|".join(shown)


def _value_sign(value: object) -> str:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number) or abs(float(number)) <= NONZERO_ABS_TOLERANCE:
        return ""
    return "positive" if float(number) > 0 else "negative"


def _mapping_sector_token_options(sector_tokens: list[str]) -> list[list[str]]:
    """Return path token options for model-export matching."""
    options: list[list[str]] = []
    for candidate in [sector_tokens, _collapse_consecutive_duplicates(sector_tokens)]:
        if candidate not in options:
            options.append(candidate)

    # Supply rows in the mapping workbook are ESTO balance concepts, but the
    # full LEAP export stores the underlying fuel branches under Resources.
    if sector_tokens and sector_tokens[0] in {"production", "imports", "exports"}:
        resource_tokens = ["resources", "primary"]
        if resource_tokens not in options:
            options.append(resource_tokens)
    return options


def _required_variable_norm_for_mapping(sector_tokens: list[str]) -> str:
    if not sector_tokens:
        return ""
    if sector_tokens[0] == "production":
        return "maximum production"
    if sector_tokens[0] == "imports":
        return "imports"
    if sector_tokens[0] == "exports":
        return "exports"
    return ""


#%%
def load_leap_to_esto_mapping() -> pd.DataFrame:
    mapping = pd.read_excel(
        LEAP_TO_ESTO_MAPPING_PATH,
        sheet_name=LEAP_TO_ESTO_MAPPING_SHEET,
        dtype=str,
    ).fillna("")
    required_cols = [
        "leap_sector_name_full_path",
        "raw_leap_fuel_name",
        "esto_flow",
        "esto_product",
    ]
    missing_cols = [col for col in required_cols if col not in mapping.columns]
    if missing_cols:
        raise ValueError(f"Mapping sheet is missing required columns: {missing_cols}")

    for col in required_cols:
        mapping[col] = mapping[col].map(_clean_text)
    if "remove_row" not in mapping.columns:
        mapping["remove_row"] = ""
    if "leap_is_subtotal" not in mapping.columns:
        mapping["leap_is_subtotal"] = ""
    if "esto_pair_is_subtotal" not in mapping.columns:
        mapping["esto_pair_is_subtotal"] = ""

    mapping["mapping_row_id"] = range(1, len(mapping) + 1)
    mapping["mapping_remove_row"] = mapping["remove_row"].map(_truthy)
    mapping["mapping_leap_is_subtotal"] = mapping["leap_is_subtotal"].map(_truthy)
    mapping["mapping_esto_pair_is_subtotal"] = mapping["esto_pair_is_subtotal"].map(_truthy)
    mapping["mapping_has_required_fields"] = (
        mapping["leap_sector_name_full_path"].ne("")
        & mapping["raw_leap_fuel_name"].ne("")
        & mapping["esto_flow"].ne("")
        & mapping["esto_product"].ne("")
    )
    return mapping


def load_full_model_export_branches() -> pd.DataFrame:
    export = pd.read_excel(
        FULL_MODEL_EXPORT_PATH,
        sheet_name=FULL_MODEL_EXPORT_SHEET,
        header=2,
        dtype=str,
    ).fillna("")
    if "Branch Path" not in export.columns:
        raise ValueError("Full model export is missing the 'Branch Path' column.")

    branch_cols = ["Branch Path", "Variable", "Scenario", "Region"]
    for col in branch_cols:
        if col not in export.columns:
            export[col] = ""
        export[col] = export[col].map(_clean_text)

    branches = (
        export[export["Branch Path"].ne("")]
        .groupby("Branch Path", as_index=False)
        .agg(
            matched_variables=("Variable", _join_unique),
            matched_scenarios=("Scenario", _join_unique),
            matched_regions=("Region", _join_unique),
        )
    )
    branches["branch_tokens"] = branches["Branch Path"].map(_path_tokens)
    branches["branch_leaf_norm"] = branches["branch_tokens"].map(lambda tokens: tokens[-1] if tokens else "")
    return branches


def load_nonzero_esto_pairs() -> pd.DataFrame:
    esto = pd.read_csv(ESTO_DATA_PATH, dtype={"economy": str, "flows": str, "products": str})
    for col in ["economy", "flows", "products"]:
        if col not in esto.columns:
            raise ValueError(f"ESTO data is missing required column: {col}")
        esto[col] = esto[col].map(_clean_text)

    if ESTO_ECONOMIES is not None:
        economies = {_clean_text(economy) for economy in ESTO_ECONOMIES if _clean_text(economy)}
        esto = esto[esto["economy"].isin(economies)].copy()
    if EXCLUDE_ESTO_SUBTOTALS and "is_subtotal" in esto.columns:
        esto = esto[~esto["is_subtotal"].map(_truthy)].copy()

    available_year_cols = [col for col in esto.columns if str(col).isdigit()]
    if ESTO_YEARS is None:
        year_cols = available_year_cols
    else:
        wanted_years = {str(year) for year in ESTO_YEARS}
        year_cols = [col for col in available_year_cols if str(col) in wanted_years]
    if not year_cols:
        raise ValueError("No ESTO year columns were available after applying ESTO_YEARS.")

    numeric_years = esto[year_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    esto["_row_abs_sum"] = numeric_years.abs().sum(axis=1)
    esto["_row_abs_max"] = numeric_years.abs().max(axis=1)
    nonzero = esto[esto["_row_abs_max"].gt(NONZERO_ABS_TOLERANCE)].copy()

    if nonzero.empty:
        return pd.DataFrame(
            columns=[
                "esto_flow",
                "esto_product",
                "esto_value_sign",
                "esto_value_sum",
                "esto_nonzero_economies",
                "esto_nonzero_years",
            ]
        )

    melted = nonzero.melt(
        id_vars=["economy", "flows", "products"],
        value_vars=year_cols,
        var_name="year",
        value_name="value",
    )
    melted["value"] = pd.to_numeric(melted["value"], errors="coerce").fillna(0.0)
    melted = melted[melted["value"].abs().gt(NONZERO_ABS_TOLERANCE)].copy()
    melted["esto_value_sign"] = melted["value"].map(_value_sign)
    summary = (
        melted.groupby(["flows", "products", "esto_value_sign"], as_index=False)
        .agg(
            esto_value_sum=("value", lambda s: float(s.sum())),
            esto_nonzero_economies=("economy", _join_limited),
            esto_nonzero_years=("year", _join_limited),
        )
        .rename(columns={"flows": "esto_flow", "products": "esto_product"})
    )
    return summary


#%%
def attach_model_branch_presence(mapping: pd.DataFrame, branches: pd.DataFrame) -> pd.DataFrame:
    branch_groups = {
        fuel_norm: group.copy()
        for fuel_norm, group in branches.groupby("branch_leaf_norm", dropna=False)
    }

    rows: list[dict[str, object]] = []
    for row in mapping.to_dict("records"):
        fuel_norm = _normalize_label(row.get("raw_leap_fuel_name"))
        sector_tokens = _path_tokens(row.get("leap_sector_name_full_path"))
        if sector_tokens and fuel_norm and sector_tokens[-1] == fuel_norm:
            sector_tokens = sector_tokens[:-1]
        sector_token_options = _mapping_sector_token_options(sector_tokens)
        required_variable_norm = _required_variable_norm_for_mapping(sector_tokens)

        candidates = branch_groups.get(fuel_norm, pd.DataFrame())
        matched = pd.DataFrame()
        if not candidates.empty:
            keep_mask = candidates["branch_tokens"].map(
                lambda tokens: any(
                    _subsequence_in_order(option, list(tokens[:-1]))
                    for option in sector_token_options
                )
            )
            matched = candidates[keep_mask].copy()
            if required_variable_norm and not matched.empty:
                matched = matched[
                    matched["matched_variables"].map(
                        lambda value: required_variable_norm in {
                            _normalize_label(part) for part in str(value or "").split("|")
                        }
                    )
                ].copy()

        row["model_branch_found"] = not matched.empty
        row["matched_branch_count"] = int(len(matched))
        row["matched_branch_paths"] = _join_limited(matched.get("Branch Path", pd.Series(dtype=str)), limit=8)
        row["matched_variables"] = _join_limited(matched.get("matched_variables", pd.Series(dtype=str)), limit=8)
        row["matched_scenarios"] = _join_limited(matched.get("matched_scenarios", pd.Series(dtype=str)), limit=8)
        row["fuel_name_normalized"] = fuel_norm
        row["sector_path_normalized"] = "/".join(sector_tokens)
        row["matched_sector_path_options"] = "|".join("/".join(option) for option in sector_token_options)
        row["required_variable_for_match"] = required_variable_norm
        rows.append(row)

    return pd.DataFrame(rows)


def build_missing_fuel_audit() -> dict[str, pd.DataFrame]:
    mapping = load_leap_to_esto_mapping()
    branches = load_full_model_export_branches()
    esto_nonzero = load_nonzero_esto_pairs()

    row_audit = attach_model_branch_presence(mapping, branches)
    row_audit = row_audit.merge(esto_nonzero, on=["esto_flow", "esto_product"], how="left")
    row_audit["esto_pair_has_nonzero_data"] = row_audit["esto_value_sum"].notna() & row_audit[
        "esto_value_sum"
    ].abs().gt(NONZERO_ABS_TOLERANCE)

    row_audit["included_in_missing_report"] = row_audit["mapping_has_required_fields"] & row_audit[
        "esto_pair_has_nonzero_data"
    ]
    if EXCLUDE_MAPPING_REMOVE_ROWS_FROM_MISSING_REPORT:
        row_audit["included_in_missing_report"] &= ~row_audit["mapping_remove_row"]
    if EXCLUDE_MAPPING_SUBTOTAL_ROWS_FROM_MISSING_REPORT:
        row_audit["included_in_missing_report"] &= ~row_audit["mapping_leap_is_subtotal"]
        row_audit["included_in_missing_report"] &= ~row_audit["mapping_esto_pair_is_subtotal"]

    missing_rows = row_audit[
        row_audit["included_in_missing_report"] & ~row_audit["model_branch_found"]
    ].copy()
    present_rows = row_audit[
        row_audit["included_in_missing_report"] & row_audit["model_branch_found"]
    ].copy()

    pair_audit = (
        row_audit[row_audit["included_in_missing_report"]]
        .groupby(["esto_flow", "esto_product", "esto_value_sign"], as_index=False)
        .agg(
            mapped_leap_row_count=("mapping_row_id", "count"),
            matched_leap_row_count=("model_branch_found", "sum"),
            missing_leap_row_count=("model_branch_found", lambda s: int((~s).sum())),
            leap_sector_paths=("leap_sector_name_full_path", _join_limited),
            leap_fuels=("raw_leap_fuel_name", _join_limited),
            matched_branch_paths=("matched_branch_paths", _join_limited),
            esto_value_sum=("esto_value_sum", "first"),
            esto_nonzero_economies=("esto_nonzero_economies", _join_unique),
            esto_nonzero_years=("esto_nonzero_years", _join_unique),
        )
    )
    if not pair_audit.empty:
        pair_audit["any_model_branch_found_for_esto_pair"] = pair_audit["matched_leap_row_count"].gt(0)
        pair_audit["all_mapped_leap_rows_found_for_esto_pair"] = pair_audit["missing_leap_row_count"].eq(0)

    missing_pairs = pair_audit[~pair_audit["any_model_branch_found_for_esto_pair"]].copy()
    simple_missing_esto_pairs = build_simple_missing_esto_pairs(missing_pairs)

    return {
        "mapping_row_audit": row_audit,
        "missing_mapping_rows": missing_rows,
        "present_mapping_rows": present_rows,
        "esto_pair_audit": pair_audit,
        "missing_esto_pairs": missing_pairs,
        "simple_missing_esto_pairs": simple_missing_esto_pairs,
        "nonzero_esto_pairs": esto_nonzero,
        "model_export_branches": branches.drop(columns=["branch_tokens"], errors="ignore"),
    }


def build_simple_missing_esto_pairs(missing_pairs: pd.DataFrame) -> pd.DataFrame:
    """Return a narrow modeller-facing list of ESTO pairs absent from LEAP."""
    columns = [
        "esto_flow",
        "esto_product",
        "esto_value_sign",
        "esto_value_sum",
        "esto_nonzero_economies",
        "expected_leap_locations",
    ]
    if missing_pairs.empty:
        return pd.DataFrame(columns=columns)

    simple = missing_pairs.copy()
    simple = simple[
        ~simple["esto_flow"].map(
            lambda value: _esto_flow_code(value) in set(EXCLUDED_SIMPLE_OUTPUT_ESTO_FLOW_PREFIXES)
        )
    ].copy()
    if simple.empty:
        return pd.DataFrame(columns=columns)

    simple["expected_leap_locations"] = simple.apply(
        lambda row: _format_expected_leap_locations(
            row.get("leap_sector_paths", ""),
            row.get("leap_fuels", ""),
        ),
        axis=1,
    )
    simple = simple[columns].drop_duplicates().sort_values(columns[:2], kind="mergesort")
    return simple.reset_index(drop=True)


def _esto_flow_code(value: object) -> str:
    text = _clean_text(value)
    match = re.match(r"^(\d+(?:\.\d+)*)\b", text)
    return match.group(1) if match else ""


def _format_expected_leap_locations(leap_sector_paths: object, leap_fuels: object) -> str:
    sector_values = [_clean_text(value) for value in str(leap_sector_paths or "").split("|") if _clean_text(value)]
    fuel_values = [_clean_text(value) for value in str(leap_fuels or "").split("|") if _clean_text(value)]
    if not sector_values and not fuel_values:
        return ""
    if len(sector_values) == 1 and len(fuel_values) == 1:
        return f"{sector_values[0]}/{fuel_values[0]}"
    if len(sector_values) == len(fuel_values):
        return "|".join(f"{sector}/{fuel}" for sector, fuel in zip(sector_values, fuel_values))
    if len(sector_values) == 1:
        return "|".join(f"{sector_values[0]}/{fuel}" for fuel in fuel_values)
    if len(fuel_values) == 1:
        return "|".join(f"{sector}/{fuel_values[0]}" for sector in sector_values)
    return "|".join(sector_values + fuel_values)


def write_missing_fuel_outputs(tables: dict[str, pd.DataFrame]) -> dict[str, Path]:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    paths = {
        "simple_missing_esto_pairs_csv": OUTPUT_DIR / "simple_missing_esto_pairs.csv",
        "simple_missing_esto_pairs_xlsx": OUTPUT_DIR / "simple_missing_esto_pairs.xlsx",
        "missing_mapping_rows_csv": OUTPUT_DIR / "missing_mapping_rows.csv",
        "present_mapping_rows_csv": OUTPUT_DIR / "present_mapping_rows.csv",
        "mapping_row_audit_csv": OUTPUT_DIR / "mapping_row_audit.csv",
        "esto_pair_audit_csv": OUTPUT_DIR / "esto_pair_audit.csv",
        "missing_esto_pairs_csv": OUTPUT_DIR / "missing_esto_pairs.csv",
        "nonzero_esto_pairs_csv": OUTPUT_DIR / "nonzero_esto_pairs.csv",
        "workbook": OUTPUT_DIR / "missing_leap_fuels_audit.xlsx",
    }

    paths["simple_missing_esto_pairs_csv"] = _write_csv_with_locked_file_fallback(
        tables["simple_missing_esto_pairs"],
        paths["simple_missing_esto_pairs_csv"],
    )
    paths["simple_missing_esto_pairs_xlsx"] = _write_xlsx_with_locked_file_fallback(
        tables["simple_missing_esto_pairs"],
        paths["simple_missing_esto_pairs_xlsx"],
        sheet_name="missing_esto_pairs",
    )

    if not WRITE_DETAILED_AUDIT_OUTPUTS:
        return {
            "simple_missing_esto_pairs_csv": paths["simple_missing_esto_pairs_csv"],
            "simple_missing_esto_pairs_xlsx": paths["simple_missing_esto_pairs_xlsx"],
        }

    for key, table_key in [
        ("missing_mapping_rows_csv", "missing_mapping_rows"),
        ("present_mapping_rows_csv", "present_mapping_rows"),
        ("mapping_row_audit_csv", "mapping_row_audit"),
        ("esto_pair_audit_csv", "esto_pair_audit"),
        ("missing_esto_pairs_csv", "missing_esto_pairs"),
        ("nonzero_esto_pairs_csv", "nonzero_esto_pairs"),
    ]:
        tables[table_key].to_csv(paths[key], index=False)

    with pd.ExcelWriter(paths["workbook"], engine="openpyxl") as writer:
        for sheet_name, table_key in [
            ("simple_missing_esto_pairs", "simple_missing_esto_pairs"),
            ("missing_mapping_rows", "missing_mapping_rows"),
            ("missing_esto_pairs", "missing_esto_pairs"),
            ("present_mapping_rows", "present_mapping_rows"),
            ("esto_pair_audit", "esto_pair_audit"),
            ("mapping_row_audit", "mapping_row_audit"),
            ("nonzero_esto_pairs", "nonzero_esto_pairs"),
            ("model_export_branches", "model_export_branches"),
        ]:
            tables[table_key].to_excel(writer, sheet_name=sheet_name[:31], index=False)

    return paths


def _fallback_output_path(path: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return path.with_name(f"{path.stem}_{timestamp}{path.suffix}")


def _write_csv_with_locked_file_fallback(frame: pd.DataFrame, path: Path) -> Path:
    try:
        frame.to_csv(path, index=False)
        return path
    except PermissionError:
        fallback_path = _fallback_output_path(path)
        frame.to_csv(fallback_path, index=False)
        print(f"[WARN] Output file was locked, wrote fallback CSV: {fallback_path}")
        return fallback_path


def _write_xlsx_with_locked_file_fallback(frame: pd.DataFrame, path: Path, *, sheet_name: str) -> Path:
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        return path
    except PermissionError:
        fallback_path = _fallback_output_path(path)
        with pd.ExcelWriter(fallback_path, engine="openpyxl") as writer:
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        print(f"[WARN] Output file was locked, wrote fallback XLSX: {fallback_path}")
        return fallback_path


def run_workflow() -> dict[str, object]:
    tables = build_missing_fuel_audit()
    paths = write_missing_fuel_outputs(tables)
    return {
        "simple_missing_esto_pairs": int(len(tables["simple_missing_esto_pairs"])),
        "missing_mapping_rows": int(len(tables["missing_mapping_rows"])),
        "missing_esto_pairs": int(len(tables["missing_esto_pairs"])),
        "present_mapping_rows": int(len(tables["present_mapping_rows"])),
        "mapping_row_audit_rows": int(len(tables["mapping_row_audit"])),
        **{key: str(value) for key, value in paths.items()},
    }


#%%
RUN_WORKFLOW = True
WORKFLOW_RESULT: dict[str, object] | None = None
if RUN_WORKFLOW:
    WORKFLOW_RESULT = run_workflow()
    print("[OK] LEAP full-export missing ESTO fuel audit complete.")
    for key, value in WORKFLOW_RESULT.items():
        print(f"- {key}: {value}")
#%%
