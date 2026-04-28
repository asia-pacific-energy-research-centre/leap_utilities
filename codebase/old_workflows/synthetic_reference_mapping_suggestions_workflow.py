from __future__ import annotations

import sys
from pathlib import Path
from datetime import datetime
import shutil

import pandas as pd

from codebase.utilities.master_config import read_config_table

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.mappings.canonical_mapping import load_canonical_pairs  # noqa: E402
from codebase.utilities.leap_results_dashboard_v2.reference_loader import (  # noqa: E402
    _expand_resolved_rule_targets,
    _resolve_rule_targets,
    load_synthetic_reference_rows_config,
)
from codebase.utilities.workflow_common import archive_config_dir_once_per_day  # noqa: E402


def _resolve(path_str: str) -> Path:
    normalized = str(path_str).replace("\\", "/")
    path = Path(normalized)
    if path.is_absolute():
        return path
    return (REPO_ROOT / path).resolve()


ESTO_TABLE_PATH = _resolve("data/00APEC_2025_low_with_subtotals.csv")
NINTH_TABLE_PATH = _resolve("data/merged_file_energy_ALL_20251106.csv")
SYNTHETIC_RULES_PATH = _resolve("config/synthetic_reference_rows.csv")
EXPLICIT_MAPPINGS_PATH = _resolve("config/leap_results_explicit_mappings.csv")
PAIR_PATH = _resolve("config/ninth_pairs_to_esto_pairs.xlsx")
CODEBOOK_PATH = _resolve("config/sector_fuel_codes_to_names.xlsx")
OUTPUT_DIR = _resolve("outputs/dashboards/leap_results_dashboard_v2/USA")
GENERATED_SYNTHETIC_DIR = _resolve("config/computer_generated_config/synthetic_data")
ARCHIVE_DIR = GENERATED_SYNTHETIC_DIR / "archive"
SYNTHETIC_OUTPUT_PATH = GENERATED_SYNTHETIC_DIR / "synthetic_reference_rows.xlsx"
MAPPING_OUTPUT_PATH = GENERATED_SYNTHETIC_DIR / "synthetic_mapping_suggestions.xlsx"

TARGET_FUEL_TO_ESTO_PRODUCT = {
    "16_x_ammonia": "16.10 Ammonia",
    "16_x_efuel": "16.11 E-fuel",
    "16_x_hydrogen": "16.12 Hydrogen",
}
MANUAL_CODEBOOK_SUGGESTIONS = [
    {
        "suggested_action": "add_or_update",
        "9th_label": "10_01_19_hydrogen_transformation",
        "9th_column": "sub2sectors",
        "esto_label": "10.01.19 Hydrogen transformation",
        "esto_column": "flows",
        "name": "Hydrogen transformation",
        "note": "Needed so the synthetic hydrogen own-use sector rows have a sector codebook link.",
    },
]
MANUAL_PAIR_SUGGESTIONS = [
    {
        "sheet_group": "pair_core_synthetic",
        "9th_sector": "10_01_19_hydrogen_transformation",
        "9th_fuel": "17_electricity",
        "esto_flow": "10.01.19 Hydrogen transformation",
        "esto_product": "17 Electricity",
        "suggested_action": "add_row",
        "sector_match_method": "manual review",
        "fuel_match_method": "manual review",
        "esto_base_year_nonzero": "",
        "ninth_pair_exists": False,
        "mapping_note": "Needed so the synthetic hydrogen own-use electricity rows can map through the standard 9th-ESTO comparator path.",
        "orig_9th_sector": "10_01_19_hydrogen_transformation",
        "orig_9th_fuel": "17_electricity",
        "faulty mapping note": "",
    },
]
SECTOR_COLUMNS = ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]
ESTO_KEY_COLUMNS = ["economy", "flows", "products"]
NINTH_KEY_COLUMNS = [
    "economy",
    "scenarios",
    "subtotal_layout",
    "subtotal_results",
    "sectors",
    "sub1sectors",
    "sub2sectors",
    "sub3sectors",
    "sub4sectors",
    "fuels",
    "subfuels",
]


def _clean(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _truthy(value: object) -> bool:
    return _clean(value).lower() in {"true", "1", "yes"}


def _archive_existing_file(path: Path) -> Path | None:
    if not path.exists():
        return None
    ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    archived = ARCHIVE_DIR / f"{path.stem}_{timestamp}{path.suffix}"
    shutil.copy2(path, archived)
    return archived


def _csv_columns(path: Path) -> list[str]:
    return list(pd.read_csv(path, nrows=0).columns)


def _year_columns(columns: list[str]) -> list[str]:
    return [col for col in columns if str(col).isdigit()]


def _load_minimal_esto_df() -> pd.DataFrame:
    columns = _csv_columns(ESTO_TABLE_PATH)
    usecols = [col for col in columns if col in ESTO_KEY_COLUMNS or col in _year_columns(columns)]
    return pd.read_csv(ESTO_TABLE_PATH, usecols=usecols)


def _load_minimal_ninth_df() -> pd.DataFrame:
    columns = _csv_columns(NINTH_TABLE_PATH)
    usecols = [col for col in columns if col in NINTH_KEY_COLUMNS or col in _year_columns(columns)]
    return pd.read_csv(NINTH_TABLE_PATH, usecols=usecols, low_memory=False)


def _add_deepest_sector_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["9th_sector"] = "x"
    out["9th_sector_column"] = "sectors"
    unresolved = pd.Series(True, index=out.index)
    for col in reversed(SECTOR_COLUMNS):
        values = out[col].fillna("").astype(str).str.strip()
        valid = values.ne("") & values.ne("x") & values.str.lower().ne("nan")
        mask = unresolved & valid
        out.loc[mask, "9th_sector"] = values.loc[mask]
        out.loc[mask, "9th_sector_column"] = col
        unresolved &= ~mask
    return out


def _choose_suggested_flow(pairs: pd.DataFrame, sector_code: str, fuel_code: str) -> str:
    exact = pairs[
        pairs["9th_sector"].fillna("").astype(str).str.strip().eq(sector_code)
        & pairs["9th_fuel"].fillna("").astype(str).str.strip().eq(fuel_code)
    ].copy()
    exact_flows = [token for token in exact.get("esto_flow", pd.Series(dtype=str)).map(_clean) if token]
    if exact_flows:
        return exact_flows[0]

    same_sector = pairs[pairs["9th_sector"].fillna("").astype(str).str.strip().eq(sector_code)].copy()
    same_sector["esto_flow"] = same_sector.get("esto_flow", pd.Series(dtype=str)).map(_clean)
    same_sector = same_sector[same_sector["esto_flow"] != ""]
    if same_sector.empty:
        return ""
    counts = same_sector["esto_flow"].value_counts(dropna=False)
    if counts.empty:
        return ""
    return str(counts.index[0]).strip()


def _build_synthetic_rows_workbook() -> dict[str, int]:
    archive_config_dir_once_per_day()
    print("Loading minimal ESTO/Ninth tables for synthetic rows...", flush=True)
    esto_df = _load_minimal_esto_df().drop_duplicates().reset_index(drop=True)
    ninth_df = _load_minimal_ninth_df().drop_duplicates().reset_index(drop=True)
    print("Loading explicit mappings and synthetic rules...", flush=True)
    explicit_mappings = read_config_table(EXPLICIT_MAPPINGS_PATH)
    canonical_pairs, _ = load_canonical_pairs(PAIR_PATH, strict=False)
    rules = load_synthetic_reference_rows_config(SYNTHETIC_RULES_PATH)

    print("Creating synthetic ESTO/9th rows...", flush=True)
    esto_year_cols = _year_columns(list(esto_df.columns))
    ninth_year_cols = _year_columns(list(ninth_df.columns))
    esto_templates = esto_df[[col for col in ["economy"] if col in esto_df.columns]].drop_duplicates().to_dict("records")
    ninth_templates = ninth_df[
        [col for col in ["economy", "scenarios", "subtotal_layout", "subtotal_results"] if col in ninth_df.columns]
    ].drop_duplicates().to_dict("records")
    if not esto_templates:
        esto_templates = [{}]
    if not ninth_templates:
        ninth_templates = [{}]

    existing_esto_keys = {
        tuple(_clean(row.get(col)) for col in ESTO_KEY_COLUMNS if col in esto_df.columns)
        for row in esto_df[ESTO_KEY_COLUMNS].to_dict("records")
    }
    existing_ninth_keys = {
        tuple(_clean(row.get(col)) for col in NINTH_KEY_COLUMNS if col in ninth_df.columns)
        for row in ninth_df[NINTH_KEY_COLUMNS].to_dict("records")
    }
    new_esto_keys: set[tuple[str, ...]] = set()
    new_ninth_keys: set[tuple[str, ...]] = set()
    new_esto_rows: list[dict[str, object]] = []
    new_ninth_rows: list[dict[str, object]] = []

    for _, rule in rules.iterrows():
        resolved_rule = _resolve_rule_targets(
            rule=rule,
            ninth_df=ninth_df,
            explicit_mappings=explicit_mappings,
        )
        expanded_rules = _expand_resolved_rule_targets(
            resolved_rule=resolved_rule,
            canonical_pairs=canonical_pairs,
        )
        for expanded_rule in expanded_rules:
            rule_name = _clean(expanded_rule.get("rule_name")) or "unnamed_rule"
            source_flow = _clean(expanded_rule.get("source_esto_flow"))
            source_product = _clean(expanded_rule.get("source_esto_product"))

            if bool(expanded_rule.get("create_esto")):
                for template in esto_templates:
                    candidate = {col: pd.NA for col in esto_df.columns}
                    for col, value in template.items():
                        candidate[col] = value
                    if "flows" in candidate:
                        candidate["flows"] = _clean(expanded_rule.get("target_esto_flow")) or source_flow
                    if "products" in candidate:
                        candidate["products"] = _clean(expanded_rule.get("target_esto_product")) or source_product or "x"
                    for year_col in esto_year_cols:
                        candidate[year_col] = 0
                    key = tuple(_clean(candidate.get(col)) for col in ESTO_KEY_COLUMNS if col in esto_df.columns)
                    if key in existing_esto_keys or key in new_esto_keys:
                        continue
                    candidate["_synthetic_esto_row"] = True
                    candidate["_synthetic_rule_name"] = rule_name
                    new_esto_rows.append(candidate)
                    new_esto_keys.add(key)

            if bool(expanded_rule.get("create_ninth")):
                for template in ninth_templates:
                    candidate = {col: pd.NA for col in ninth_df.columns}
                    for col, value in template.items():
                        candidate[col] = value
                    for col, rule_key in [
                        ("sectors", "target_sectors"),
                        ("sub1sectors", "target_sub1sectors"),
                        ("sub2sectors", "target_sub2sectors"),
                        ("sub3sectors", "target_sub3sectors"),
                        ("sub4sectors", "target_sub4sectors"),
                        ("fuels", "target_fuels"),
                        ("subfuels", "target_subfuels"),
                    ]:
                        if col in candidate:
                            candidate[col] = _clean(expanded_rule.get(rule_key)) or "x"
                    if "subtotal_layout" in candidate:
                        candidate["subtotal_layout"] = False
                    if "subtotal_results" in candidate:
                        candidate["subtotal_results"] = False
                    for year_col in ninth_year_cols:
                        candidate[year_col] = 0
                    key = tuple(_clean(candidate.get(col)) for col in NINTH_KEY_COLUMNS if col in ninth_df.columns)
                    if key in existing_ninth_keys or key in new_ninth_keys:
                        continue
                    candidate["_synthetic_ninth_row"] = True
                    candidate["_synthetic_rule_name"] = rule_name
                    new_ninth_rows.append(candidate)
                    new_ninth_keys.add(key)

    esto_synth = pd.DataFrame(new_esto_rows, columns=list(esto_df.columns) + ["_synthetic_esto_row", "_synthetic_rule_name"])
    ninth_synth = pd.DataFrame(new_ninth_rows, columns=list(ninth_df.columns) + ["_synthetic_ninth_row", "_synthetic_rule_name"])
    if not esto_synth.empty:
        esto_synth = esto_synth.sort_values(["economy", "flows", "products"], kind="stable")
    if not ninth_synth.empty:
        ninth_synth = ninth_synth.sort_values(
            ["economy", "scenarios", "sectors", "sub1sectors", "sub2sectors", "fuels", "subfuels"],
            kind="stable",
        )

    print("Writing synthetic rows workbook...", flush=True)
    GENERATED_SYNTHETIC_DIR.mkdir(parents=True, exist_ok=True)
    _archive_existing_file(SYNTHETIC_OUTPUT_PATH)
    with pd.ExcelWriter(SYNTHETIC_OUTPUT_PATH, engine="openpyxl") as writer:
        esto_synth.to_excel(writer, sheet_name="ESTO Synthetic Rows", index=False)
        ninth_synth.to_excel(writer, sheet_name="9th Synthetic Rows", index=False)
    return {
        "esto_rows": int(len(esto_synth)),
        "ninth_rows": int(len(ninth_synth)),
    }


def _build_pair_suggestions(
    *,
    ninth_df: pd.DataFrame,
    pairs: pd.DataFrame,
) -> pd.DataFrame:
    year_cols = [col for col in ninth_df.columns if str(col).isdigit()]
    mask = ninth_df["subfuels"].fillna("").astype(str).isin(TARGET_FUEL_TO_ESTO_PRODUCT)
    working = ninth_df.loc[mask].copy()
    values = working[year_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
    working = working.loc[values.ne(0).any(axis=1)].copy()
    working = _add_deepest_sector_columns(working)

    group_cols = ["9th_sector", "9th_sector_column", "fuels", "subfuels"]
    grouped = (
        working.groupby(group_cols, dropna=False)
        .agg(
            nonzero_row_count=("subfuels", "size"),
            economies=("economy", lambda s: ",".join(sorted({str(v).strip() for v in s if _clean(v)}))),
            scenarios=("scenarios", lambda s: ",".join(sorted({str(v).strip() for v in s if _clean(v)}))),
        )
        .reset_index()
        .rename(columns={"subfuels": "9th_fuel", "fuels": "9th_fuel_group"})
        .sort_values(["9th_sector", "9th_fuel"], kind="stable")
    )

    suggestion_rows: list[dict[str, object]] = []
    pairs = pairs.copy()
    pairs["9th_sector"] = pairs["9th_sector"].map(_clean)
    pairs["9th_fuel"] = pairs["9th_fuel"].map(_clean)
    pairs["esto_flow"] = pairs["esto_flow"].map(_clean)
    pairs["esto_product"] = pairs["esto_product"].map(_clean)

    for row in grouped.to_dict("records"):
        sector_code = _clean(row["9th_sector"])
        fuel_code = _clean(row["9th_fuel"])
        suggested_product = TARGET_FUEL_TO_ESTO_PRODUCT.get(fuel_code, "")
        exact = pairs[(pairs["9th_sector"] == sector_code) & (pairs["9th_fuel"] == fuel_code)].copy()
        exact_nonblank = exact[(exact["esto_flow"] != "") | (exact["esto_product"] != "")].copy()
        suggested_flow = _choose_suggested_flow(pairs, sector_code, fuel_code)
        exact_has_target = bool(
            (
                exact["esto_flow"].eq(suggested_flow)
                & exact["esto_product"].eq(suggested_product)
            ).any()
        ) if not exact.empty and suggested_flow else False
        exact_has_any = not exact_nonblank.empty
        exact_faulty = exact.get("faulty mapping", pd.Series(dtype=object)).map(_truthy).any() if not exact.empty else False

        if exact_has_target and not exact_faulty:
            action = "existing_ok"
        elif exact_has_any:
            action = "update_existing"
        else:
            action = "add_row"

        exact_first = exact.reset_index(drop=True).iloc[0] if not exact.empty else pd.Series(dtype=object)
        faulty_note = ""
        if exact_faulty:
            faulty_note = "Existing row for this 9th pair is blank, faulty, or points to a different ESTO target."

        suggestion_rows.append(
            {
                "sheet_group": "pair_core_synthetic",
                "suggested_action": action,
                "9th_sector": sector_code,
                "9th_fuel": fuel_code,
                "esto_flow": suggested_flow,
                "esto_product": suggested_product,
                "sector_match_method": _clean(exact_first.get("sector_match_method")) or "manual review",
                "fuel_match_method": _clean(exact_first.get("fuel_match_method")) or "manual review",
                "esto_base_year_nonzero": _clean(exact_first.get("esto_base_year_nonzero")),
                "ninth_pair_exists": bool(exact_has_any),
                "mapping_note": "Nonzero 9th sector-fuel combination for new 16_x fuel; review ESTO flow and insert/update canonical pair manually.",
                "orig_9th_sector": sector_code,
                "orig_9th_fuel": fuel_code,
                "faulty mapping note": faulty_note,
                "nonzero_row_count": int(row["nonzero_row_count"]),
                "economies": row["economies"],
                "scenarios": row["scenarios"],
                "suggested_action": action,
            }
        )
    return pd.DataFrame(suggestion_rows)


def _build_codebook_suggestions(
    *,
    codebook: pd.DataFrame,
) -> pd.DataFrame:
    codebook = codebook.copy()
    codebook["9th_label"] = codebook["9th_label"].map(_clean)
    codebook["9th_column"] = codebook["9th_column"].map(_clean)
    codebook["esto_label"] = codebook["esto_label"].map(_clean)
    codebook["esto_column"] = codebook["esto_column"].map(_clean)
    codebook["name"] = codebook["name"].map(_clean)

    suggestions: list[dict[str, object]] = []
    seen_keys: set[tuple[str, str, str, str]] = set()

    for fuel_code, esto_product in TARGET_FUEL_TO_ESTO_PRODUCT.items():
        exact_existing = codebook[
            codebook["9th_label"].eq(fuel_code)
            & codebook["9th_column"].eq("subfuels")
            & codebook["esto_label"].eq(esto_product)
            & codebook["esto_column"].eq("products")
        ]
        if exact_existing.empty:
            partial_existing = codebook[
                codebook["9th_label"].eq(fuel_code)
                | codebook["esto_label"].eq(esto_product)
            ]
            action = "update_existing" if not partial_existing.empty else "add_row"
            key = (fuel_code, "subfuels", esto_product, "products")
            if key not in seen_keys:
                seen_keys.add(key)
                suggestions.append(
                    {
                        "suggested_action": action,
                        "9th_label": fuel_code,
                        "9th_column": "subfuels",
                        "esto_label": esto_product,
                        "esto_column": "products",
                        "name": codebook.loc[codebook["9th_label"].eq(fuel_code), "name"].map(_clean).head(1).squeeze() or fuel_code,
                        "note": "Link new 16_x fuel code to its ESTO product code.",
                    }
                )

    for row in MANUAL_CODEBOOK_SUGGESTIONS:
        exact_existing = codebook[
            codebook["9th_label"].eq(_clean(row["9th_label"]))
            & codebook["9th_column"].eq(_clean(row["9th_column"]))
            & codebook["esto_label"].eq(_clean(row["esto_label"]))
            & codebook["esto_column"].eq(_clean(row["esto_column"]))
        ]
        if not exact_existing.empty:
            continue
        partial_existing = codebook[
            codebook["9th_label"].eq(_clean(row["9th_label"]))
            | codebook["esto_label"].eq(_clean(row["esto_label"]))
        ]
        suggestion = dict(row)
        suggestion["suggested_action"] = "update_existing" if not partial_existing.empty else "add_row"
        suggestions.append(suggestion)
    out = pd.DataFrame(suggestions).drop_duplicates(
        subset=["9th_label", "9th_column", "esto_label", "esto_column"],
        keep="first",
    )
    return out.sort_values(["9th_column", "9th_label", "esto_column", "esto_label"], kind="stable")


def _write_mapping_suggestions_workbook() -> dict[str, int]:
    print("Loading minimal Ninth table and mapping workbooks...", flush=True)
    ninth_df = _load_minimal_ninth_df()
    pairs = read_config_table(PAIR_PATH)
    codebook = read_config_table(CODEBOOK_PATH, sheet_name="code_to_name")

    print("Building pair suggestions...", flush=True)
    pair_suggestions = _build_pair_suggestions(ninth_df=ninth_df, pairs=pairs)
    pair_suggestions = pd.concat(
        [pair_suggestions, pd.DataFrame(MANUAL_PAIR_SUGGESTIONS)],
        ignore_index=True,
        sort=False,
    ).drop_duplicates(
        subset=["sheet_group", "9th_sector", "9th_fuel", "esto_flow", "esto_product"],
        keep="first",
    )
    print("Building code_to_name suggestions...", flush=True)
    codebook_suggestions = _build_codebook_suggestions(
        codebook=codebook,
    )

    print("Writing mapping suggestions workbook...", flush=True)
    GENERATED_SYNTHETIC_DIR.mkdir(parents=True, exist_ok=True)
    _archive_existing_file(MAPPING_OUTPUT_PATH)
    core_pair_suggestions = pair_suggestions[pair_suggestions["sheet_group"].eq("pair_core_synthetic")].copy()
    pair_cols = [
        "suggested_action",
        "9th_sector",
        "9th_fuel",
        "esto_flow",
        "esto_product",
        "sector_match_method",
        "fuel_match_method",
        "esto_base_year_nonzero",
        "ninth_pair_exists",
        "mapping_note",
        "orig_9th_sector",
        "orig_9th_fuel",
        "faulty mapping note",
    ]
    with pd.ExcelWriter(MAPPING_OUTPUT_PATH, engine="openpyxl") as writer:
        codebook_suggestions.to_excel(writer, sheet_name="code_to_name", index=False)
        core_pair_suggestions[pair_cols].to_excel(writer, sheet_name="pair_core_synthetic", index=False)
    return {
        "code_to_name_rows": int(len(codebook_suggestions)),
        "pair_rows": int(len(core_pair_suggestions)),
    }


def main() -> None:
    print("Starting synthetic reference outputs...", flush=True)
    synthetic_counts = _build_synthetic_rows_workbook()
    print("Starting mapping suggestion outputs...", flush=True)
    mapping_counts = _write_mapping_suggestions_workbook()
    print(f"Wrote {SYNTHETIC_OUTPUT_PATH}")
    print(f"  ESTO synthetic rows: {synthetic_counts['esto_rows']}")
    print(f"  9th synthetic rows: {synthetic_counts['ninth_rows']}")
    print(f"Wrote {MAPPING_OUTPUT_PATH}")
    print(f"  code_to_name suggestions: {mapping_counts['code_to_name_rows']}")
    print(f"  pair suggestions: {mapping_counts['pair_rows']}")


if __name__ == "__main__":
    main()


try:
    from codebase.utilities.workflow_common import emit_completion_beep as _emit_completion_beep
except Exception:  # pragma: no cover
    def _emit_completion_beep(*, success: bool = True) -> None:  # noqa: ARG001
        return


if __name__ == "__main__":  # pragma: no cover
    _emit_completion_beep(success=True, style="chime")
