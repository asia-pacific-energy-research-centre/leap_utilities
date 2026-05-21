#%%
"""
Utility helpers for the LEAP Results dashboard workflow.

Functions are designed for notebook-first usage with clear toggles and small,
composable pieces. The helpers:
- load sheet/sector mappings
- normalize fuel labels using canonical codebooks (with optional backup overrides)
- parse LEAP result workbooks (template-style sheets)
- pull reference (ESTO) and projection (9th) series
- assemble comparison DataFrames and lightweight status diagnostics
- generate simple charts and HTML dashboards (style reused from leap_transport)
"""
from __future__ import annotations

import math
import os
import re
import sys
import json
from collections import Counter, defaultdict
from html import escape
from pathlib import Path
from typing import Optional, Sequence

import pandas as pd

from codebase.utilities.master_config import config_table_exists, read_config_table

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.functions.ninth_projection_mapping import normalize_economy_key
from codebase.mappings.canonical_mapping import (
    build_code_match_method as shared_build_code_match_method,
    build_sector_to_esto_flow_lookup as shared_build_sector_to_esto_flow_lookup,
    is_reverse_independent_match as shared_is_reverse_independent_match,
    clean_token as shared_clean_token,
    load_canonical_pairs as shared_load_canonical_pairs,
    load_fuel_aliases as shared_load_fuel_aliases,
    load_sheet_map as shared_load_sheet_map,
    map_fuel_label as shared_map_fuel_label,
    normalize_match_method as shared_normalize_match_method,
    normalize_label as shared_normalize_label,
    split_sector_codes as shared_split_sector_codes,
)

# Stable paths (overridable by caller)
DEFAULT_SHEET_MAP = Path("config/leap_results_sheet_map.csv")
DEFAULT_BACKUP_LEAP_MAPPINGS = Path("config/backup_leap_mappings.xlsx")
DEFAULT_EXPLICIT_LEAP_MAPPINGS = Path("config/leap_results_explicit_mappings.csv")
DEFAULT_EXPLICIT_LEAP_REASSIGNMENTS = Path("config/leap_results_explicit_reassignments.csv")
DEFAULT_CODEBOOK = Path("config/sector_fuel_codes_to_names.xlsx")
DEFAULT_NINTH_FUEL_PAIRS = Path("config/ninth_sector_fuel_pairs.csv")
DEFAULT_NINTH_TO_ESTO = Path("config/ninth_pairs_to_esto_pairs.xlsx")

def _safe_read_codebook_sheet(codebook_path: Path, sheet_name: str) -> pd.DataFrame:
    try:
        return read_config_table(codebook_path, sheet_name=sheet_name)
    except (FileNotFoundError, ValueError) as exc:
        print(f"[WARN] Codebook read failed for sheet {sheet_name!r} in {codebook_path}: {exc}")
        return pd.DataFrame()

TRANSFER_SHEET_DISPLAY_OVERRIDES = {
    "transfers_inputs": "Transformation transfer inputs",
    "transfers_out_feed": "Transformation transfer outputs by feedstock",
    "transfers_out_fuel": "Transformation transfer outputs by product",
    "transfers_unallocated_inputs": "Transfers unallocated inputs",
    "refinery_blending_inputs": "Refinery blending inputs",
    "upstream_liquids_inputs": "Upstream liquids inputs",
    "transfers_unallocated_out_feed": "Transfers unallocated outputs by feedstock",
    "refinery_blending_out_feed": "Refinery blending outputs by feedstock",
    "upstream_liquids_out_feed": "Upstream liquids outputs by feedstock",
    "transfers_unallocated_out_fuel": "Transfers unallocated outputs by product",
    "refinery_blending_out_fuel": "Refinery blending outputs by product",
    "upstream_liquids_out_fuel": "Upstream liquids outputs by product",
}

_LABEL_TO_NINTH_FUEL_FALLBACK = {
    "ammonia": "16_x_ammonia",
    "bitumen": "07_x_other_petroleum_products",
    "efuel": "16_x_efuel",
    "hydrogen": "16_x_hydrogen",
    "lubricants": "07_x_other_petroleum_products",
    "other products": "07_x_other_petroleum_products",
    "other sources": "16_09_other_sources",
    "paraffin waxes": "07_x_other_petroleum_products",
    "petprod nonspecified": "07_x_other_petroleum_products",
    "petroleum coke": "07_x_other_petroleum_products",
    "white spirit sbp": "07_x_other_petroleum_products",
}


# -----------------------------------------------------------------------------
# Helpers: path / repo handling
# -----------------------------------------------------------------------------
REPO_ROOT = Path(__file__).resolve().parents[2]


def ensure_repo_root() -> None:
    """Force cwd to repo root so relative paths work in notebooks."""
    cwd = Path.cwd()
    if cwd != REPO_ROOT:
        os.chdir(REPO_ROOT)


def _infer_fuel_from_label_fallback(fuel_label: str) -> str:
    """
    Fill a small set of known LEAP labels that map cleanly to one 9th fuel.

    This covers labels where the codebook provides the ESTO product, but the
    canonical pairs table may not yield a sector-specific 9th match for
    aggregate sectors such as ``14_03_manufacturing``.
    """
    return _LABEL_TO_NINTH_FUEL_FALLBACK.get(_normalize_label(fuel_label), "")


def _merge_mapping_note(existing: str, addition: str) -> str:
    existing = _clean_token(existing)
    addition = _clean_token(addition)
    if not addition:
        return existing
    if not existing:
        return addition
    existing_parts = {part.strip().lower() for part in re.split(r"\s*;\s*", existing) if part.strip()}
    if addition.lower() in existing_parts:
        return existing
    return f"{existing}; {addition}"


def _global_numeric_sector_prefix(code: object) -> str:
    text = _clean_token(code).lower()
    if not text:
        return ""
    parts: list[str] = []
    for token in text.split("_"):
        tok = token.strip()
        if tok.isdigit():
            parts.append(tok)
            continue
        break
    return "_".join(parts)


def _hierarchical_parent_levels(child_code: object, parent_code: object) -> int | None:
    child_prefix = _global_numeric_sector_prefix(child_code)
    parent_prefix = _global_numeric_sector_prefix(parent_code)
    if not child_prefix or not parent_prefix:
        return None
    child_tokens = child_prefix.split("_")
    parent_tokens = parent_prefix.split("_")
    if len(parent_tokens) >= len(child_tokens):
        return None
    if child_tokens[: len(parent_tokens)] != parent_tokens:
        return None
    return len(child_tokens) - len(parent_tokens)


def _derive_parent_flow_levels(
    sector_codes: list[str],
    esto_flow_norm: str,
    flow_to_parent_sector_code: dict[str, str],
) -> int | None:
    flow_parent_code = _clean_token(flow_to_parent_sector_code.get(str(esto_flow_norm or "").strip().lower(), ""))
    if not flow_parent_code:
        return None
    matches = [
        level
        for sector_code in sector_codes
        for level in [_hierarchical_parent_levels(sector_code, flow_parent_code)]
        if level is not None
    ]
    if not matches:
        return None
    return min(matches)


def _global_numeric_sector_depth(code: object) -> int:
    prefix = _global_numeric_sector_prefix(code)
    return len(prefix.split("_")) if prefix else 0


# -----------------------------------------------------------------------------
# Mapping loaders
# -----------------------------------------------------------------------------
def load_sheet_map(path: Path = DEFAULT_SHEET_MAP) -> pd.DataFrame:
    """Read sheet→sector map and return active rows with normalized names."""
    return shared_load_sheet_map(path)


def _build_sheet_display_labels(sheet_map: pd.DataFrame | None = None) -> dict[str, str]:
    """
    Build human-readable sheet labels for chart/dashboard display.

    Prefer the explicit notes column because transformation/power sheets need
    the variable role called out (for example inputs vs outputs by product).
    Fall back to sector names and finally the raw sheet name.
    """
    if sheet_map is None or sheet_map.empty:
        try:
            sheet_map = load_sheet_map(DEFAULT_SHEET_MAP)
        except Exception:
            return {}

    labels: dict[str, str] = {}
    for _, row in sheet_map.iterrows():
        sheet_name = str(row.get("sheet_name") or "").strip()
        if not sheet_name:
            continue
        notes = _clean_token(row.get("notes"))
        final_category_name = _clean_token(row.get("final_category_name"))
        sector_name = _clean_token(row.get("sector_name"))
        labels[sheet_name] = TRANSFER_SHEET_DISPLAY_OVERRIDES.get(
            sheet_name,
            notes or final_category_name or sector_name or sheet_name,
        )
    labels.update(TRANSFER_SHEET_DISPLAY_OVERRIDES)
    return labels


def _format_loss_own_use_display_label(sheet_name: str, notes: str = "", sector_name: str = "") -> str:
    sheet_text = str(sheet_name or "").strip()
    note_text = str(notes or "").strip()
    sector_text = re.sub(r"\s*\(own-use\)\s*$", "", str(sector_name or "").strip(), flags=re.IGNORECASE)
    if not sheet_text.endswith("_loss_own_use_total"):
        return ""
    if note_text.lower().startswith("derived "):
        compact = re.sub(r"^Derived\s+", "", note_text, flags=re.IGNORECASE).strip()
        compact = re.sub(
            r"\s+total\s+from\s+inputs\s+minus\s+product\s+outputs.*$",
            "",
            compact,
            flags=re.IGNORECASE,
        ).strip(" ;")
        compact = re.sub(r"\blosses\s*/\s*own use\b", "losses & own use", compact, flags=re.IGNORECASE)
        if compact:
            return compact[:1].upper() + compact[1:] + " (derived)"
    if sector_text:
        return f"{sector_text} losses & own use (derived)"
    compact_sheet = sheet_text[: -len("_loss_own_use_total")].replace("_", " ").strip()
    if compact_sheet:
        return compact_sheet[:1].upper() + compact_sheet[1:] + " losses & own use (derived)"
    return sheet_text


def _build_sheet_display_metadata(sheet_map: pd.DataFrame | None = None) -> dict[str, dict[str, str]]:
    """Return display metadata keyed by sheet name."""
    if sheet_map is None or sheet_map.empty:
        try:
            sheet_map = load_sheet_map(DEFAULT_SHEET_MAP)
        except Exception:
            return {}

    transfer_sheet_overrides = _build_sheet_display_labels(sheet_map)

    metadata: dict[str, dict[str, str]] = {}
    for _, row in sheet_map.iterrows():
        sheet_name = str(row.get("sheet_name") or "").strip()
        if not sheet_name:
            continue
        notes = _clean_token(row.get("notes"))
        sector_name = _clean_token(row.get("sector_name"))
        category_type = _clean_token(row.get("category_type")).lower()
        metadata[sheet_name] = {
            "label": (
                _format_loss_own_use_display_label(sheet_name, notes, sector_name)
                or transfer_sheet_overrides.get(sheet_name, notes or sector_name or sheet_name)
            ),
            "notes": notes,
            "sector_name": sector_name,
            "category_type": category_type,
        }
    for sheet_name, label in TRANSFER_SHEET_DISPLAY_OVERRIDES.items():
        metadata.setdefault(
            sheet_name,
            {
                "label": label,
                "notes": "",
                "sector_name": "Transfers",
                "category_type": "fuel",
            },
        )
    return metadata


def _normalize_measure_label(text: object, *, default_units: str = "PJ") -> str:
    cleaned = _clean_token(text)
    if not cleaned:
        return ""
    lowered = cleaned.lower()
    if any(unit in lowered for unit in ["(pj", "(gwh", "(mw", "(gw", " petajoule", " gwh", " mw", " gw"]):
        return cleaned
    return f"{cleaned} ({default_units})"


def _infer_sheet_measure_from_row(row: pd.Series | dict[str, object]) -> str:
    explicit_measure = _clean_token(row.get("measure"))
    if explicit_measure:
        return _normalize_measure_label(explicit_measure)

    category_type = _clean_token(row.get("category_type")).lower()
    sector_code = _clean_token(row.get("sector_code_9th")).lower()
    sector_name = _clean_token(row.get("sector_name"))
    notes = _clean_token(row.get("notes"))

    if category_type == "sector":
        return "Demand (PJ)"
    if sector_name in {"Production", "Imports", "Exports"}:
        return f"{sector_name} (PJ)"
    if sector_code.startswith(("01_", "02_", "03_", "07_")):
        return "Supply flow (PJ)"
    if notes:
        return _normalize_measure_label(notes)
    if sector_code.startswith(("09_", "10_", "18_")):
        return "Transformation inputs and outputs (PJ)"
    if sector_code.startswith("15_"):
        return "Demand (PJ)"
    return "Energy (PJ)"


def _measure_is_input_only(measure: object) -> bool:
    text = _clean_token(measure).lower()
    if not text:
        return False
    return "input" in text and "output" not in text


def _measure_is_output_only(measure: object) -> bool:
    text = _clean_token(measure).lower()
    if not text:
        return False
    return "output" in text and "input" not in text


def _transformation_sign_role_from_measure(measure: object, sheet: object = "") -> str:
    """
    Infer directional sign role for transformation lookups.

    Feedstock-output sheets (except power feedstock MAP sheets) represent
    feed inputs allocated to output families, so they should pull input-signed
    values from ESTO/9th transformation tables.
    """
    sheet_text = _clean_token(sheet).lower()
    if sheet_text.endswith("_out_feed") and sheet_text not in {"elecgen_out_feed", "heat_out_feed"}:
        return "input"
    if _measure_is_input_only(measure):
        return "input"
    if _measure_is_output_only(measure):
        return "output"
    return ""


def _sheet_is_export_flow(sheet: object, measure: object = "") -> bool:
    sheet_text = _clean_token(sheet).lower()
    measure_text = _clean_token(measure).lower()
    return sheet_text.startswith("exports") or "export" in measure_text


def _power_feedstock_output_override(sheet: object) -> tuple[list[str], str]:
    sheet_text = _clean_token(sheet)
    if sheet_text == "elecgen_out_feed":
        return ["18_01_electricity_plants"], "18.01 MAP electricity plants"
    if sheet_text == "heat_out_feed":
        return ["18_02_chp_plants"], "18.02 MAP CHP plants"
    return [], ""


def _suppress_feedstock_output_chart(sheet: object) -> bool:
    # Keep feedstock-output charts visible so transformation branches can show
    # consistent inputs/outputs coverage.
    return False


def _build_sheet_measure_lookup(sheet_map: pd.DataFrame | None = None) -> dict[str, str]:
    if sheet_map is None or sheet_map.empty:
        try:
            sheet_map = load_sheet_map(DEFAULT_SHEET_MAP)
        except Exception:
            return {}

    lookup: dict[str, str] = {}
    for _, row in sheet_map.iterrows():
        sheet_name = str(row.get("sheet_name") or "").strip()
        if not sheet_name:
            continue
        lookup[sheet_name] = _infer_sheet_measure_from_row(row)
    return lookup


def load_explicit_sector_fuel_mappings(path: Path = DEFAULT_EXPLICIT_LEAP_MAPPINGS) -> pd.DataFrame:
    """
    Load exact sheet/fuel(/sector) overrides.

    These mappings are intentionally literal. If a row matches, it bypasses the
    generic canonical/codebook inference logic in ``build_comparisons``.
    Multiple rows with the same ``sheet_name``/``fuel_label``/``sector_code_9th``
    are allowed and are treated as explicit aggregate components.
    """
    if not config_table_exists(path):
        return pd.DataFrame(
            columns=[
                "sheet_name",
                "fuel_label",
                "sector_code_9th",
                "projection_sector_code",
                "ninth_fuel_code",
                "esto_flow",
                "esto_product",
                "mapping_note",
            ]
        )

    try:
        df = read_config_table(path, encoding="utf-8")
    except UnicodeDecodeError:
        # Some maintained mapping files are saved as cp1252/latin-1 from Excel.
        # Fall back so explicit overrides still load instead of aborting workflow.
        df = read_config_table(path, encoding="latin1")
    df.columns = [str(c).strip().lower() for c in df.columns]
    if "active" in df.columns:
        df = df[df["active"].astype(str).str.strip().str.lower().isin({"true", "1", "yes"})]

    for col in [
        "sheet_name",
        "fuel_label",
        "sector_code_9th",
        "projection_sector_code",
        "ninth_fuel_code",
        "esto_flow",
        "esto_product",
        "mapping_note",
    ]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].map(_clean_token)

    return df[
        [
            "sheet_name",
            "fuel_label",
            "sector_code_9th",
            "projection_sector_code",
            "ninth_fuel_code",
            "esto_flow",
            "esto_product",
            "mapping_note",
        ]
    ].reset_index(drop=True)


def _extract_allow_missing_base_flag(mapping_note: object) -> tuple[str, bool]:
    """Parse optional explicit note tokens that relax base-coverage requirements."""
    text = _clean_token(mapping_note)
    if not text:
        return "", False
    allow_missing_base = "[allow_missing_base]" in text.lower()
    if allow_missing_base:
        text = re.sub(r"\s*\[allow_missing_base\]\s*", " ", text, flags=re.IGNORECASE)
        text = re.sub(r"\s{2,}", " ", text).strip()
    return text, allow_missing_base


def load_explicit_sector_reassignments(path: Path = DEFAULT_EXPLICIT_LEAP_REASSIGNMENTS) -> pd.DataFrame:
    """
    Load exact source->target row reassignments for ESTO base and 9th projection data.

    These rules are literal. Each populated source column is matched exactly, and
    the corresponding target columns are written exactly.
    """
    columns = [
        "rule_name",
        "source_sectors",
        "source_sub1sectors",
        "source_sub2sectors",
        "source_sub3sectors",
        "source_sub4sectors",
        "source_fuels",
        "source_subfuels",
        "source_esto_flow",
        "source_esto_product",
        "target_sectors",
        "target_sub1sectors",
        "target_sub2sectors",
        "target_sub3sectors",
        "target_sub4sectors",
        "target_fuels",
        "target_subfuels",
        "target_esto_flow",
        "target_esto_product",
        "notes",
    ]
    if not config_table_exists(path):
        return pd.DataFrame(columns=columns)

    df = read_config_table(path)
    df.columns = [str(c).strip().lower() for c in df.columns]
    if "active" in df.columns:
        df = df[df["active"].astype(str).str.strip().str.lower().isin({"true", "1", "yes"})]
    for col in columns:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].map(_clean_token)
    return df[columns].reset_index(drop=True)


def apply_explicit_sector_reassignments(
    base_df: pd.DataFrame,
    ninth_df: pd.DataFrame,
    reassignments: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Apply exact source->target reassignments to ESTO base and 9th projection rows.

    Returns ``(base_df_adjusted, ninth_df_adjusted, reassignment_status)``.
    """
    rules = reassignments.copy() if reassignments is not None else load_explicit_sector_reassignments()
    if rules.empty:
        empty_status = pd.DataFrame(
            columns=["rule_name", "dataset", "matched_rows", "target_esto_flow", "target_esto_product", "notes"]
        )
        return base_df.copy(), ninth_df.copy(), empty_status

    base_out = base_df.copy()
    ninth_out = ninth_df.copy()
    status_rows: list[dict[str, object]] = []

    def _exact_mask(df: pd.DataFrame, criteria: dict[str, str]) -> pd.Series:
        if df.empty:
            return pd.Series(False, index=df.index)
        mask = pd.Series(True, index=df.index)
        for col, expected in criteria.items():
            if col not in df.columns or expected == "":
                continue
            values = df[col].fillna("").astype(str).str.strip().str.lower()
            mask &= values.eq(expected.lower())
        return mask

    for _, rule in rules.iterrows():
        rule_name = _clean_token(rule.get("rule_name")) or "unnamed_rule"
        notes = _clean_token(rule.get("notes"))

        base_mask = _exact_mask(
            base_out,
            {
                "flows": _clean_token(rule.get("source_esto_flow")),
                "products": _clean_token(rule.get("source_esto_product")),
            },
        )
        base_match_count = int(base_mask.sum())
        if base_match_count:
            if "flows" in base_out.columns and _clean_token(rule.get("target_esto_flow")):
                base_out.loc[base_mask, "flows"] = _clean_token(rule.get("target_esto_flow"))
            if "products" in base_out.columns and _clean_token(rule.get("target_esto_product")):
                base_out.loc[base_mask, "products"] = _clean_token(rule.get("target_esto_product"))
        status_rows.append(
            {
                "rule_name": rule_name,
                "dataset": "base_df",
                "matched_rows": base_match_count,
                "target_esto_flow": _clean_token(rule.get("target_esto_flow")),
                "target_esto_product": _clean_token(rule.get("target_esto_product")),
                "notes": notes,
            }
        )

        ninth_mask = _exact_mask(
            ninth_out,
            {
                "sectors": _clean_token(rule.get("source_sectors")),
                "sub1sectors": _clean_token(rule.get("source_sub1sectors")),
                "sub2sectors": _clean_token(rule.get("source_sub2sectors")),
                "sub3sectors": _clean_token(rule.get("source_sub3sectors")),
                "sub4sectors": _clean_token(rule.get("source_sub4sectors")),
                "fuels": _clean_token(rule.get("source_fuels")),
                "subfuels": _clean_token(rule.get("source_subfuels")),
            },
        )
        ninth_match_count = int(ninth_mask.sum())
        if ninth_match_count:
            for col in ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]:
                target_key = f"target_{col}"
                if col in ninth_out.columns and _clean_token(rule.get(target_key)):
                    ninth_out.loc[ninth_mask, col] = _clean_token(rule.get(target_key))
            for col in ["fuels", "subfuels"]:
                target_key = f"target_{col}"
                if col in ninth_out.columns and _clean_token(rule.get(target_key)):
                    ninth_out.loc[ninth_mask, col] = _clean_token(rule.get(target_key))
        status_rows.append(
            {
                "rule_name": rule_name,
                "dataset": "ninth_df",
                "matched_rows": ninth_match_count,
                "target_esto_flow": _clean_token(rule.get("target_esto_flow")),
                "target_esto_product": _clean_token(rule.get("target_esto_product")),
                "notes": notes,
            }
        )

    status_df = pd.DataFrame(status_rows)
    return base_out, ninth_out, status_df


def _split_sector_codes(raw_value: object) -> list[str]:
    """
    Parse one-to-many sector mapping tokens from a sheet-map cell.
    Accepted separators: ',', ';', '|', 'AND'.
    """
    return shared_split_sector_codes(raw_value)


def _build_codebook_lookup(codebook_path: Path) -> dict[str, str]:
    """Map human-readable names to 9th fuel codes."""
    df = _safe_read_codebook_sheet(codebook_path, "code_to_name")
    lookup: dict[str, str] = {}
    if df.empty:
        return lookup
    for _, row in df.iterrows():
        name = str(row.get("name") or "").strip().lower()
        code = str(row.get("9th_label") or "").strip()
        if name and code:
            lookup[name] = code
    return lookup


def _clean_token(value: object) -> str:
    return shared_clean_token(value)


def _build_name_to_esto_product(codebook_path: Path) -> dict[str, str]:
    """Map human-readable names to ESTO product codes."""
    df = _safe_read_codebook_sheet(codebook_path, "code_to_name")
    lookup: dict[str, str] = {}
    if df.empty:
        return lookup
    for _, row in df.iterrows():
        name = str(row.get("name") or "").strip().lower()
        est = str(row.get("esto_label") or "").strip()
        if name and est:
            lookup[name] = est
    return lookup


def build_sector_to_esto_flow_lookup(codebook_path: Path = DEFAULT_CODEBOOK) -> dict[str, str]:
    """Map 9th sector codes to ESTO flow labels from code_to_name."""
    return shared_build_sector_to_esto_flow_lookup(codebook_path)


def _build_leap_esto_lookup(codebook_path: Path) -> dict[str, str]:
    """
    Build mapping from LEAP fuel label (as used in LEAP exports) to 9th fuel code
    by chaining through ESTO_LEAP_names (LEAP label -> ESTO label) and code_to_name
    (ESTO label -> 9th label).
    """
    df_leap = _safe_read_codebook_sheet(codebook_path, "ESTO_LEAP_names")
    df_code = _safe_read_codebook_sheet(codebook_path, "code_to_name")
    if df_leap.empty or df_code.empty:
        return {}
    esto_to_9th = {
        str(r["esto_label"]).strip().lower(): str(r["9th_label"]).strip()
        for _, r in df_code.iterrows()
        if pd.notna(r.get("esto_label")) and pd.notna(r.get("9th_label"))
    }
    lookup: dict[str, str] = {}
    for _, row in df_leap.iterrows():
        if str(row.get("category")).strip().lower() != "products":
            continue
        leap_label = str(row.get("leap_name") or "").strip().lower()
        esto_label = str(row.get("original_label") or "").strip().lower()
        if not leap_label or not esto_label:
            continue
        ninth = esto_to_9th.get(esto_label, "")
        if ninth:
            lookup[leap_label] = ninth
    return lookup


def _normalize_label(value: object) -> str:
    """Lowercase and collapse whitespace for robust text joins."""
    return shared_normalize_label(value)


def load_canonical_pairs(
    path: Path = DEFAULT_NINTH_TO_ESTO,
    *,
    strict: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Load canonical 9th->ESTO pairs and return (clean_pairs, conflicts).
    Hard conflict: same (9th_sector, 9th_fuel) maps to inconsistent
    (esto_flow, esto_product).
    """
    return shared_load_canonical_pairs(path=path, strict=strict)


def _extract_esto_code(esto_label: str) -> str:
    """Extract ESTO code prefix from labels like '07.12 White spirit SBP'."""
    match = re.match(r"^\s*(\d{2}(?:\.\d{2})?)\b", str(esto_label or ""))
    if not match:
        return ""
    return match.group(1).replace(".", "_")


def _build_ninth_fuel_lookup(ninth_fuel_pairs_path: Path) -> tuple[dict[str, str], dict[str, str]]:
    """
    Build lookups from ESTO-like numeric prefixes to valid 9th fuel codes.
    Returns:
    - exact_lookup: 07_01 -> 07_01_motor_gasoline (if available in pairs)
    - group_lookup: 07 -> 07_petroleum_products
    """
    pairs = read_config_table(ninth_fuel_pairs_path)
    exact_lookup: dict[str, str] = {}
    group_lookup: dict[str, str] = {}

    # Group mapping (e.g., 07 -> 07_petroleum_products) comes only from fuels.
    if "fuels" in pairs.columns:
        fuel_codes = (
            pairs["fuels"]
            .astype(str)
            .str.strip()
            .replace({"": pd.NA, "x": pd.NA})
            .dropna()
            .unique()
            .tolist()
        )
        for code in sorted(fuel_codes):
            m2 = re.match(r"^(\d{2})_", code)
            if m2:
                group_lookup.setdefault(m2.group(1), code)

    # Exact mapping (e.g., 07_01 -> 07_01_motor_gasoline) from detailed fuel codes.
    cols = [c for c in ["fuel_pair", "subfuels"] if c in pairs.columns]
    candidates = (
        pd.concat([pairs[c].astype(str) for c in cols], ignore_index=True)
        .str.strip()
        .replace({"": pd.NA, "x": pd.NA})
        .dropna()
        .unique()
        .tolist()
    )
    for code in sorted(candidates):
        m = re.match(r"^(\d{2}(?:_\d{2})?)_", code)
        if m:
            exact_lookup.setdefault(m.group(1), code)
    return exact_lookup, group_lookup


def load_fuel_aliases(
    alias_path: Path | str | None = DEFAULT_BACKUP_LEAP_MAPPINGS,
    codebook_path: Path = DEFAULT_CODEBOOK,
) -> dict[str, dict[str, str]]:
    """
    Build a mapping from LEAP fuel labels to ESTO products (+ optional explicit overrides).
    Priority:
    1) codebook-driven mapping (ESTO_LEAP_names + code_to_name name harmonization)
    2) explicit backup overrides (optional, wins)
    Returns dict keyed by normalized leap fuel label.
    """
    return shared_load_fuel_aliases(alias_path=alias_path, codebook_path=codebook_path)


def map_fuel_label(
    fuel_label: str,
    fuel_mapping: dict[str, dict[str, str]],
    fallback_codebook: dict[str, str] | None = None,
) -> dict[str, str]:
    """
    Return mapping hints for a LEAP fuel label.
    """
    return shared_map_fuel_label(
        fuel_label=fuel_label,
        fuel_mapping=fuel_mapping,
        fallback_codebook=fallback_codebook,
    )


# -----------------------------------------------------------------------------
# LEAP workbook parsing
# -----------------------------------------------------------------------------
def parse_template_sheet(sheet: pd.DataFrame) -> dict:
    """
    Extract metadata and series from a template-style LEAP results sheet.
    Expected structure (matching leap_results_workflow template refills):
    row0: variable
    row1: "Scenario: X, Region: Y"
    row2: "Branch: ..."
    row3: "Units: ..."
    row5: header with legend label + years
    row6+: legend members with values per year
    """
    meta: dict[str, object] = {}
    meta["variable"] = str(sheet.iloc[0, 0]).strip()
    scenario_region = str(sheet.iloc[1, 0])
    for part in scenario_region.split(","):
        if "Scenario:" in part:
            meta["scenario"] = part.split(":", 1)[1].strip()
        if "Region:" in part:
            meta["region"] = part.split(":", 1)[1].strip()
    meta["branch"] = str(sheet.iloc[2, 0]).split(":", 1)[-1].strip()
    meta["units"] = str(sheet.iloc[3, 0]).split(":", 1)[-1].strip()
    meta["legend_label"] = str(sheet.iloc[5, 0]).strip()
    def _parse_year_token(value: object) -> int | None:
        if pd.isna(value):
            return None
        # Direct numeric parse first.
        try:
            num = float(value)
            if math.isfinite(num):
                direct = int(round(num))
                if 1900 <= direct <= 2200:
                    return direct
                # Some LEAP exports encode years as tiny scientific notation, e.g. 2.022e-12.
                for factor in (1e3, 1e6, 1e9, 1e12, 1e15):
                    scaled = int(round(num * factor))
                    if 1900 <= scaled <= 2200:
                        return scaled
        except Exception:
            pass

        # String fallback: extract a 4-digit year if present.
        text = str(value).strip()
        m = re.search(r"\b(19\d{2}|20\d{2}|21\d{2})\b", text)
        if m:
            return int(m.group(1))
        return None

    year_cols: list[tuple[int, int]] = []
    for col_idx, val in enumerate(sheet.iloc[5, 1:], start=1):
        year_int = _parse_year_token(val)
        if year_int is None:
            # Skip non-year tokens such as "Total"
            continue
        year_cols.append((col_idx, year_int))
    years = [y for _, y in year_cols]
    meta["years"] = years
    records: list[dict] = []
    data_rows: list[tuple[str, pd.Series]] = []
    for _, row in sheet.iloc[6:, :].iterrows():
        fuel = str(row.iloc[0]).strip()
        if not fuel or pd.isna(fuel):
            break
        data_rows.append((fuel, row))

    default_skip_labels = {
        "total",
        "demand total",
        "international transport",
        "freight road",
        "freight non road",
        "nonspecified transport",
    }
    has_non_skipped_rows = any(str(fuel).strip().lower() not in default_skip_labels for fuel, _ in data_rows)
    allow_total_row = not has_non_skipped_rows

    for fuel, row in data_rows:
        fuel_lower = fuel.lower()
        if fuel_lower in default_skip_labels:
            if fuel_lower == "total" and allow_total_row:
                pass
            else:
                continue
        for col_idx, year in year_cols:
            val = row.iloc[col_idx]
            try:
                num = float(val)
            except Exception:
                num = float("nan")
            records.append(
                {
                    "fuel_label": fuel,
                    "year": year,
                    "leap_value": num,
                }
            )
    return {"meta": meta, "records": pd.DataFrame(records)}


_RESULTS_UNIT_PREFIX_SCALE = {
    "thousand": 1e3,
    "million": 1e6,
    "billion": 1e9,
}

_RESULTS_UNIT_BASE_SCALE = {
    "joule": 1.0,
    "joules": 1.0,
    "gigajoule": 1e9,
    "gigajoules": 1e9,
    "petajoule": 1e15,
    "petajoules": 1e15,
    "watt": 1.0,
    "watts": 1.0,
    "kilowatt": 1e3,
    "kilowatts": 1e3,
    "megawatt": 1e6,
    "megawatts": 1e6,
    "gigawatt": 1e9,
    "gigawatts": 1e9,
}


def _normalize_results_unit_text(unit_text: object) -> str:
    return " ".join(str(unit_text or "").strip().lower().split())


def _results_unit_to_si_scale(unit_text: object) -> float | None:
    normalized = _normalize_results_unit_text(unit_text)
    if not normalized:
        return None
    prefix_scale = 1.0
    base_unit = normalized
    parts = normalized.split(" ", 1)
    if len(parts) == 2 and parts[0] in _RESULTS_UNIT_PREFIX_SCALE:
        prefix_scale = _RESULTS_UNIT_PREFIX_SCALE[parts[0]]
        base_unit = parts[1]
    base_scale = _RESULTS_UNIT_BASE_SCALE.get(base_unit)
    if base_scale is None:
        return None
    return prefix_scale * base_scale


def _conversion_factor_between_results_units(from_unit: object, to_unit: object) -> float | None:
    from_norm = _normalize_results_unit_text(from_unit)
    to_norm = _normalize_results_unit_text(to_unit)
    if not from_norm or not to_norm:
        return None
    if from_norm == to_norm:
        return 1.0
    from_scale = _results_unit_to_si_scale(from_norm)
    to_scale = _results_unit_to_si_scale(to_norm)
    if from_scale is None or to_scale is None:
        return None
    return float(from_scale) / float(to_scale)


def _leap_results_target_unit(workbook: Path, meta: dict[str, object]) -> str:
    """Return the dashboard's desired display unit for a LEAP results sheet."""
    variable = str(meta.get("variable") or "").strip()
    workbook_name = workbook.name.lower()
    energy_variables = {
        "Final Energy Demand",
        "Inputs",
        "Outputs by Feedstock Fuel",
        "Outputs by Output Fuel",
        "Indigenous Production",
        "Imports",
        "Exports",
    }
    if variable in energy_variables:
        return "Petajoules"
    if variable.lower() == "capacity" or "capacity" in variable.lower():
        return "Thousand Megawatts"
    if workbook_name.startswith(("supply_results_", "transformation_results_")) and variable in energy_variables:
        return "Petajoules"
    return ""


def _leap_workbook_energy_scale(workbook: Path, meta: dict[str, object], records: pd.DataFrame) -> tuple[float, str]:
    """
    Infer a scale factor for dashboard workbook imports.

    Prefer explicit unit conversion from the sheet's declared LEAP Results unit
    to the dashboard target unit. Keep the old magnitude-based fallback only
    for legacy workbooks that mislabeled raw values as Petajoules.
    """
    if records.empty:
        return 1.0, ""

    variable = str(meta.get("variable") or "").strip()
    units = str(meta.get("units") or "").strip()
    workbook_name = workbook.name.lower()
    target_unit = _leap_results_target_unit(workbook, meta)
    if target_unit:
        explicit_factor = _conversion_factor_between_results_units(units, target_unit)
        if explicit_factor is not None:
            if float(explicit_factor) == 1.0:
                return 1.0, ""
            return float(explicit_factor), f"converted from {units} to {target_unit}"
    raw_energy_variables = {
        "Inputs",
        "Outputs by Feedstock Fuel",
        "Outputs by Output Fuel",
        "Indigenous Production",
        "Imports",
        "Exports",
    }
    is_supply_or_transformation = (
        workbook_name.startswith("supply_results_")
        or workbook_name.startswith("transformation_results_")
    )
    if not (
        is_supply_or_transformation
        and _normalize_results_unit_text(units) == "petajoules"
        and variable in raw_energy_variables
    ):
        return 1.0, ""

    values = pd.to_numeric(records.get("leap_value"), errors="coerce")
    max_abs = values.abs().max()
    if pd.isna(max_abs) or float(max_abs) < 1e6:
        return 1.0, ""
    return 1e-6, "supply/transformation raw LEAP energy values converted to Petajoules"


def load_leap_workbook(
    workbook: Path,
    sheet_map: pd.DataFrame,
    expected_scenario: Optional[str] = None,
) -> pd.DataFrame:
    """
    Load all mapped sheets from a LEAP results workbook into long form.
    Columns: economy, scenario, sheet_name, sector_code_9th, fuel_label, year, leap_value
    """
    xl = pd.ExcelFile(workbook)

    def _sheet_name_candidates(mapped_sheet_name: str) -> list[str]:
        base = str(mapped_sheet_name or "").strip()
        if not base:
            return []
        candidates = [base]
        if base.endswith("_inputs"):
            candidates.append(base[: -len("_inputs")] + "_feed_inputs")
        if base.startswith("transfers_unallocated_"):
            shortened = base.replace("transfers_unallocated_", "transfers_unalloc_", 1)
            candidates.append(shortened)
            if shortened.endswith("_inputs"):
                candidates.append(shortened[: -len("_inputs")] + "_feed_inputs")
        out: list[str] = []
        seen: set[str] = set()
        for name in candidates:
            if name and name not in seen:
                out.append(name)
                seen.add(name)
        return out

    mapped_sheet_names = set(sheet_map["sheet_name"].astype(str).str.strip().tolist())
    mapped_or_alias_sheet_names = set()
    for mapped_name in mapped_sheet_names:
        for candidate in _sheet_name_candidates(mapped_name):
            mapped_or_alias_sheet_names.add(candidate)
    unmapped_sheets = [sheet for sheet in xl.sheet_names if str(sheet).strip() not in mapped_or_alias_sheet_names]
    if unmapped_sheets:
        print(
            f"[WARN] Workbook {workbook.name} contains {len(unmapped_sheets)} sheet(s) not in "
            f"config/leap_results_sheet_map.csv; they will be skipped: {unmapped_sheets}"
        )
    rows: list[pd.DataFrame] = []
    workbook_sheets = set(xl.sheet_names)

    for _, mapping in sheet_map.iterrows():
        sheet_name = mapping["sheet_name"]
        candidate_names = _sheet_name_candidates(sheet_name)
        actual_sheet = next((candidate for candidate in candidate_names if candidate in workbook_sheets), "")
        if not actual_sheet:
            continue
        sheet_df = xl.parse(actual_sheet, header=None)
        parsed = parse_template_sheet(sheet_df)
        meta = parsed["meta"]
        # Prefer the explicitly expected scenario from the workbook context
        # (filename/workflow selection), because some sheets can carry stale or
        # inconsistent embedded scenario metadata.
        scenario = str(expected_scenario or meta.get("scenario") or "").strip()
        region = str(meta.get("region") or "").strip()
        df = parsed["records"].copy()
        scale_factor, scale_note = _leap_workbook_energy_scale(workbook, meta, df)
        if scale_factor != 1.0 and not df.empty:
            df["leap_value"] = pd.to_numeric(df["leap_value"], errors="coerce") * float(scale_factor)
        df["sheet_name"] = sheet_name
        df["sector_code_9th"] = mapping["sector_code_9th"]
        df["sector_name"] = mapping.get("sector_name", "")
        df["scenario"] = scenario
        df["region"] = region
        df["leap_variable"] = str(meta.get("variable") or "").strip()
        df["leap_units"] = str(meta.get("units") or "").strip()
        df["measure"] = _normalize_measure_label(
            df["leap_variable"].iloc[0] if not df.empty else str(meta.get("variable") or "").strip(),
            default_units=str(meta.get("units") or "PJ").strip() or "PJ",
        )
        df["leap_scale_note"] = scale_note
        # Infer economy code from filename tokens if region missing
        economy_token = None
        m = re.search(r"_([A-Z]{3})_", workbook.name)
        if m:
            economy_token = m.group(1)
        df["economy"] = economy_token or region or ""
        rows.append(df)
    if not rows:
        return pd.DataFrame(
            columns=[
                "economy",
                "scenario",
                "sheet_name",
                "sector_code_9th",
                "sector_name",
                "fuel_label",
                "year",
                "leap_value",
                "leap_variable",
                "leap_units",
                "leap_scale_note",
            ]
        )
    return pd.concat(rows, ignore_index=True)


# -----------------------------------------------------------------------------
# Reference/projection data handlers
# -----------------------------------------------------------------------------
SECTOR_COLUMNS = ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]
FUEL_COLUMNS = ["fuels", "subfuels"]


def _filter_ninth_by_sector_fuel(
    ninth_df: pd.DataFrame,
    sector_code: str,
    fuel_code: str,
    economy_code: str,
    scenario: str,
    value_sign_role: str = "",
) -> pd.DataFrame:
    """Return rows for matching economy, scenario, sector, and fuel."""
    prepared = "__projection_prepared" in ninth_df.columns
    working = ninth_df
    economy_token = str(economy_code or "").strip()
    scenario_token = str(scenario or "").strip().lower()
    if prepared:
        if "__economy_norm" in working.columns and economy_token:
            working = working[working["__economy_norm"] == economy_token]
        if "__scenario_norm" in working.columns and scenario_token:
            working = working[working["__scenario_norm"] == scenario_token]
    else:
        working = working[(working["economy"] == economy_code) & (working["scenarios"] == scenario)]
    if working.empty:
        return working
    # OR-match the sector code across all sector columns.
    sector_code_lower = sector_code.lower()
    sector_mask = pd.Series(False, index=working.index)
    for col in SECTOR_COLUMNS:
        norm_col = f"__sector_norm__{col}"
        if prepared and norm_col in working.columns:
            sector_mask |= working[norm_col] == sector_code_lower
        elif col in working.columns:
            sector_mask |= working[col].astype(str).str.lower() == sector_code_lower
    working = working[sector_mask]
    if working.empty:
        return working
    # Prefer detailed rows when a sector contains a mix of detailed and subtotal
    # fuel rows. Keeping subtotals first can drop valid detailed fuel matches
    # before the fuel filter is applied.
    if prepared and "__subtotal_results" in working.columns:
        detail_rows = working[~working["__subtotal_results"]]
        if not detail_rows.empty:
            working = detail_rows
        elif "__subtotal_layout" in working.columns:
            detail_layout_rows = working[~working["__subtotal_layout"]]
            if not detail_layout_rows.empty:
                working = detail_layout_rows
    else:
        for flag_col in ("subtotal_results", "subtotal_layout"):
            if flag_col in working.columns:
                flag_values = working[flag_col].fillna(False).astype(bool)
                detail_rows = working[~flag_values]
                if not detail_rows.empty:
                    working = detail_rows
                    break
    if working.empty:
        return working
    if fuel_code:
        fuel_mask = False
        for col in FUEL_COLUMNS:
            norm_col = f"__fuel_norm__{col}"
            if prepared and norm_col in working.columns:
                fuel_mask |= working[norm_col] == fuel_code.lower()
            elif col in working.columns:
                fuel_mask |= working[col].astype(str).str.lower() == fuel_code.lower()
        working = working[fuel_mask]
    if working.empty:
        return working
    sign_role = str(value_sign_role or "").strip().lower()
    is_directional_balance_sector = str(sector_code or "").strip().lower().startswith(("08_", "09_"))
    if sign_role in {"input", "output"} and is_directional_balance_sector:
        if prepared and {"__has_negative", "__has_positive"}.issubset(working.columns):
            if sign_role == "input":
                working = working[working["__has_negative"]]
            else:
                working = working[working["__has_positive"]]
        else:
            year_cols = [c for c in working.columns if str(c).isdigit()]
            if year_cols:
                values = working[year_cols].apply(pd.to_numeric, errors="coerce")
                if sign_role == "input":
                    working = working[values.lt(0).any(axis=1)]
                else:
                    working = working[values.gt(0).any(axis=1)]
    return working


def _prepare_ninth_projection_frame(
    ninth_df: pd.DataFrame,
    *,
    economy_code: str,
    scenario_values: set[str],
) -> pd.DataFrame:
    """Pre-filter and normalize the 9th table once for repeated projection lookups."""
    working = ninth_df.copy()
    if "economy" in working.columns:
        working["__economy_norm"] = working["economy"].astype(str).str.strip()
        working = working[working["__economy_norm"] == str(economy_code or "").strip()]
    if "scenarios" in working.columns:
        working["__scenario_norm"] = working["scenarios"].astype(str).str.strip().str.lower()
        if scenario_values:
            working = working[working["__scenario_norm"].isin(scenario_values)]
    for col in SECTOR_COLUMNS:
        if col in working.columns:
            working[f"__sector_norm__{col}"] = working[col].astype(str).str.strip().str.lower()
    for col in FUEL_COLUMNS:
        if col in working.columns:
            working[f"__fuel_norm__{col}"] = working[col].astype(str).str.strip().str.lower()
    for flag_col, prepared_col in (("subtotal_results", "__subtotal_results"), ("subtotal_layout", "__subtotal_layout")):
        if flag_col in working.columns:
            working[prepared_col] = working[flag_col].fillna(False).astype(bool)
        else:
            working[prepared_col] = False
    year_cols = [c for c in working.columns if str(c).isdigit()]
    if year_cols:
        year_values = working[year_cols].apply(pd.to_numeric, errors="coerce")
        working["__has_negative"] = year_values.lt(0).any(axis=1)
        working["__has_positive"] = year_values.gt(0).any(axis=1)
    else:
        working["__has_negative"] = False
        working["__has_positive"] = False
    working["__projection_prepared"] = True
    return working


def _extract_year_series(df: pd.DataFrame, years: Sequence[int]) -> pd.Series:
    """Return a series indexed by year with numeric values (NaN preserved)."""
    if df.empty:
        return pd.Series(dtype="float64", index=years)
    year_col_map: dict[int, object] = {}
    for year in years:
        if year in df.columns:
            year_col_map[int(year)] = year
            continue
        year_str = str(int(year))
        if year_str in df.columns:
            year_col_map[int(year)] = year_str
    if not year_col_map:
        return pd.Series(dtype="float64", index=years)
    # sum across matching rows (common convention in 9th data)
    summed = df[list(year_col_map.values())].apply(pd.to_numeric, errors="coerce").sum()
    summed.index = [int(str(col)) for col in summed.index]
    # Reindex to requested years to keep shape stable.
    return summed.reindex([int(y) for y in years])


def _projection_series_scale_to_pj(sector_code: object) -> float:
    """
    Return the projection-unit scale needed for dashboard comparison values.

    The 9th edition stores the `18_*` electricity-output branches in GWh.
    Feedstock-output dashboard sheets reuse those branches to preserve
    feedstock lineage, so convert them back to PJ before charting.
    """
    sector_text = _clean_token(sector_code).lower()
    if sector_text.startswith("18_") or sector_text == "18_electricity_output_in_gwh":
        return 0.0036
    return 1.0


def _base_value_scale_to_pj(esto_flow: object) -> float:
    """Return the ESTO flow scale needed for dashboard comparison values."""
    flow_text = _clean_token(esto_flow).lower()
    if flow_text.startswith("18"):
        return 0.0036
    return 1.0


def pull_projection_series(
    ninth_df: pd.DataFrame,
    sector_code: str,
    fuel_code: str,
    economy_code: str,
    scenario: str,
    projection_years: Sequence[int],
    value_sign_role: str = "",
) -> pd.Series:
    filtered = _filter_ninth_by_sector_fuel(
        ninth_df,
        sector_code,
        fuel_code,
        economy_code,
        scenario,
        value_sign_role=value_sign_role,
    )
    series = _extract_year_series(filtered, projection_years)
    sign_role = str(value_sign_role or "").strip().lower()
    # Directional sign filtering is valid for signed transfer/transformation
    # balance sectors (`08_*` and `09_*`). Power-output sectors (`18_*`) can
    # carry positive series even when the dashboard sheet is conceptually
    # "inputs" after remapping.
    is_directional_balance_sector = str(sector_code or "").strip().lower().startswith(("08_", "09_"))
    if is_directional_balance_sector:
        if sign_role == "input":
            # Keep structural zeros for directional transformation charts;
            # dropping them to NaN creates false "missing comparator" holes.
            series = series.where(series.le(0))
        elif sign_role == "output":
            series = series.where(series.ge(0))
    scale_to_pj = _projection_series_scale_to_pj(sector_code)
    if scale_to_pj != 1.0:
        series = series.astype("float64") * scale_to_pj
    return series


def pull_projection_series_from_descendants(
    ninth_df: pd.DataFrame,
    sector_code: str,
    fuel_code: str,
    economy_code: str,
    scenario: str,
    projection_years: Sequence[int],
    value_sign_role: str = "",
) -> tuple[pd.Series, list[str]]:
    """
    Aggregate descendant 9th sector rows for a parent sector/fuel when there is
    no direct parent series.

    This keeps a comparison group on a consistent child level rather than
    mixing one child-resolved comparator with an ancestor fallback.
    """
    sector_token = _clean_token(sector_code).lower()
    fuel_token = _clean_token(fuel_code).lower()
    if not sector_token or not fuel_token or ninth_df.empty:
        return pd.Series(dtype="float64", index=projection_years), []

    prefix = _global_numeric_sector_prefix(sector_token)
    if not prefix:
        return pd.Series(dtype="float64", index=projection_years), []

    working = ninth_df.copy()
    prepared = "__projection_prepared" in working.columns

    fuel_mask = False
    for col in FUEL_COLUMNS:
        norm_col = f"__fuel_norm__{col}"
        if prepared and norm_col in working.columns:
            fuel_mask |= working[norm_col] == fuel_token
        elif col in working.columns:
            fuel_mask |= working[col].astype(str).str.strip().str.lower() == fuel_token
    working = working[fuel_mask]
    if working.empty:
        return pd.Series(dtype="float64", index=projection_years), []

    descendant_codes: set[str] = set()
    for col in SECTOR_COLUMNS:
        norm_col = f"__sector_norm__{col}"
        if prepared and norm_col in working.columns:
            values = working[norm_col].astype(str)
        elif col in working.columns:
            values = working[col].astype(str).str.strip().str.lower()
        else:
            continue
        for value in values.unique().tolist():
            code = str(value).strip().lower()
            if not code or code == "x":
                continue
            code_prefix = _global_numeric_sector_prefix(code)
            if not code_prefix or code_prefix == prefix:
                continue
            parent_parts = prefix.split("_")
            code_parts = code_prefix.split("_")
            if len(code_parts) > len(parent_parts) and code_parts[: len(parent_parts)] == parent_parts:
                descendant_codes.add(code)

    if not descendant_codes:
        return pd.Series(dtype="float64", index=projection_years), []

    deepest = max(_global_numeric_sector_depth(code) for code in descendant_codes)
    selected_codes = sorted(code for code in descendant_codes if _global_numeric_sector_depth(code) == deepest)
    if not selected_codes:
        return pd.Series(dtype="float64", index=projection_years), []

    parts: list[pd.Series] = []
    for child_code in selected_codes:
        part = pull_projection_series(
            ninth_df,
            sector_code=child_code,
            fuel_code=fuel_code,
            economy_code=economy_code,
            scenario=scenario,
            projection_years=projection_years,
            value_sign_role=value_sign_role,
        )
        if part.notna().any():
            parts.append(part.reindex(projection_years))
    if not parts:
        return pd.Series(dtype="float64", index=projection_years), []
    return pd.concat(parts, axis=1).sum(axis=1, min_count=1), selected_codes


def pull_base_year_value(
    esto_df: pd.DataFrame,
    base_year: int,
    economy_code: str,
    esto_flow: str,
    esto_product: str,
    value_sign_role: str = "",
) -> float:
    prepared = "__base_prepared" in esto_df.columns
    working = esto_df
    if prepared and "__economy_norm" in working.columns:
        working = working[working["__economy_norm"] == str(economy_code or "").strip()]
    else:
        working = working[(working["economy"] == economy_code)]
    if esto_flow:
        if prepared and "__flow_norm" in working.columns:
            working = working[working["__flow_norm"] == esto_flow.lower()]
        else:
            working = working[working["flows"].astype(str).str.lower() == esto_flow.lower()]
    if eso_product := esto_product:
        if prepared and "__product_norm" in working.columns:
            working = working[working["__product_norm"] == eso_product.lower()]
        else:
            working = working[working["products"].astype(str).str.lower() == eso_product.lower()]
    # If parent-flow exact match is unavailable, fallback to summing child flows under that parent code (e.g., 14.03.*).
    if working.empty and esto_flow and eso_product:
        parent = str(esto_flow).strip().lower()
        parent_code_match = re.match(r"^(\d+(?:\.\d+)*)", parent)
        parent_code = parent_code_match.group(1) if parent_code_match else ""
        fallback = esto_df
        if prepared and "__economy_norm" in fallback.columns:
            fallback = fallback[fallback["__economy_norm"] == str(economy_code or "").strip()]
        else:
            fallback = fallback[(fallback["economy"] == economy_code)]
        if prepared and "__product_norm" in fallback.columns:
            fallback = fallback[fallback["__product_norm"] == eso_product.lower()]
        else:
            fallback = fallback[fallback["products"].astype(str).str.lower() == eso_product.lower()]
        if parent_code:
            if prepared and "__flow_code" in fallback.columns:
                flow_codes = fallback["__flow_code"].fillna("")
            else:
                flow_codes = fallback["flows"].astype(str).str.extract(r"^(\d+(?:\.\d+)*)", expand=False).fillna("")
            fallback = fallback[flow_codes.str.startswith(parent_code + ".")]
        else:
            if prepared and "__flow_norm" in fallback.columns:
                fallback = fallback[fallback["__flow_norm"].str.startswith(parent + ".")]
            else:
                fallback = fallback[fallback["flows"].astype(str).str.lower().str.startswith(parent + ".")]
        working = fallback
    sign_role = str(value_sign_role or "").strip().lower()
    is_directional_balance_flow = str(esto_flow or "").strip().lower().startswith(("08", "09"))
    if sign_role in {"input", "output"} and is_directional_balance_flow:
        if prepared and "__base_value_num" in working.columns:
            base_values = working["__base_value_num"]
        elif str(base_year) in working.columns:
            base_values = pd.to_numeric(working[str(base_year)], errors="coerce")
        else:
            base_values = pd.Series(dtype="float64")
        if not base_values.empty:
            if sign_role == "input":
                working = working[base_values.lt(0)]
            else:
                working = working[base_values.gt(0)]
    try:
        if prepared and "__base_value_num" in working.columns:
            value = float(pd.to_numeric(working["__base_value_num"], errors="coerce").sum())
        else:
            value = float(pd.to_numeric(working[str(base_year)], errors="coerce").sum())
        value *= _base_value_scale_to_pj(esto_flow)
        if sign_role == "input" and value > 0:
            return float("nan")
        if sign_role == "output" and value < 0:
            return float("nan")
        return value
    except Exception:
        return float("nan")


def _prepare_base_lookup_frame(
    esto_df: pd.DataFrame,
    *,
    economy_code: str,
    base_year: int,
) -> pd.DataFrame:
    """Pre-filter and normalize the ESTO table once for repeated base lookups."""
    working = esto_df.copy()
    if "economy" in working.columns:
        working["__economy_norm"] = working["economy"].astype(str).str.strip()
        working = working[working["__economy_norm"] == str(economy_code or "").strip()]
    if "flows" in working.columns:
        working["__flow_norm"] = working["flows"].astype(str).str.strip().str.lower()
        working["__flow_code"] = working["flows"].astype(str).str.extract(r"^(\d+(?:\.\d+)*)", expand=False)
    if "products" in working.columns:
        working["__product_norm"] = working["products"].astype(str).str.strip().str.lower()
    if str(base_year) in working.columns:
        working["__base_value_num"] = pd.to_numeric(working[str(base_year)], errors="coerce")
    else:
        working["__base_value_num"] = pd.NA
    working["__base_prepared"] = True
    return working


def aggregate_esto_by_ninth_pairs(
    esto_df: pd.DataFrame,
    ninth_pairs: pd.DataFrame,
    base_year: int,
    economy_code: str,
) -> pd.DataFrame:
    """
    Attach 9th sector/fuel codes to ESTO flows/products using pairs file, then aggregate.
    Returns long DataFrame: economy, scenario, sheet, fuel_label, source, year, value, ninth_sector, ninth_fuel.
    """
    working = esto_df.copy()
    working["flows_norm"] = working["flows"].astype(str).str.strip().str.lower()
    working["products_norm"] = working["products"].astype(str).str.strip().str.lower()
    pairs = ninth_pairs.copy()
    pairs["esto_flow_norm"] = pairs["esto_flow"].astype(str).str.strip().str.lower()
    pairs["esto_product_norm"] = pairs["esto_product"].astype(str).str.strip().str.lower()

    merged = working.merge(
        pairs[["esto_flow_norm", "esto_product_norm", "9th_sector", "9th_fuel"]],
        left_on=["flows_norm", "products_norm"],
        right_on=["esto_flow_norm", "esto_product_norm"],
        how="inner",
    )
    if merged.empty:
        return pd.DataFrame(columns=["economy", "ninth_sector", "ninth_fuel", "year", "value"])

    year_cols = [str(base_year)] + [c for c in merged.columns if c.isdigit()]
    value_cols = [c for c in year_cols if c in merged.columns]
    melted = merged.melt(
        id_vars=["economy", "9th_sector", "9th_fuel"],
        value_vars=value_cols,
        var_name="year",
        value_name="value",
    )
    melted["year"] = pd.to_numeric(melted["year"], errors="coerce")
    melted = melted[(melted["economy"] == economy_code) & melted["year"].notna()]
    melted["value"] = pd.to_numeric(melted["value"], errors="coerce")
    agg = (
        melted.groupby(["economy", "9th_sector", "9th_fuel", "year"], as_index=False)["value"]
        .sum(min_count=1)
    )
    return agg


# -----------------------------------------------------------------------------
# Comparison assembly
# -----------------------------------------------------------------------------
def build_comparisons(
    leap_long: pd.DataFrame,
    sheet_map: pd.DataFrame,
    fuel_mapping: dict[str, dict[str, str]],
    sector_flow_mapping: dict[str, str],
    ninth_pairs: pd.DataFrame,
    base_df: pd.DataFrame,
    ninth_df: pd.DataFrame,
    explicit_mappings: pd.DataFrame | None = None,
    *,
    base_year: int,
    base_economy: str,
    projection_economy: str,
    projection_years: Sequence[int],
    scenario_map: dict[str, str],
    use_esto_agg_only: bool = False,
    sibling_comparator_mode: str = "none",
    include_sibling_parent_totals: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Return comparison_long, comparison_wide, mapping_status.
    - comparison_long: economy, scenario, sheet, fuel, source, year, value
    - comparison_wide: economy, scenario, sheet, fuel, year, leap_value, base_value, projection_value
    - mapping_status: per fuel mapping diagnostics
    """
    sibling_mode = str(sibling_comparator_mode or "").strip().lower()
    status_rows: list[dict] = []
    long_rows: list[dict] = []
    pairs, _ = load_canonical_pairs(DEFAULT_NINTH_TO_ESTO, strict=False) if ninth_pairs.empty else (ninth_pairs.copy(), pd.DataFrame())
    for col in ["9th_sector", "9th_fuel", "esto_flow", "esto_product"]:
        pairs[col] = pairs[col].map(_clean_token)
    for col in ["sector_match_method", "fuel_match_method", "mapping_note"]:
        if col not in pairs.columns:
            pairs[col] = ""
        pairs[col] = pairs[col].map(_clean_token)
    for col in ["sector_match_method", "fuel_match_method"]:
        pairs[col] = pairs[col].map(shared_normalize_match_method)
    pairs = pairs[(pairs["9th_sector"] != "") & (pairs["9th_fuel"] != "")]
    pairs = pairs[(pairs["esto_flow"] != "") & (pairs["esto_product"] != "")]
    pairs["sector_norm"] = pairs["9th_sector"].str.lower()
    pairs["fuel_norm"] = pairs["9th_fuel"].str.lower()
    pairs["esto_flow_norm"] = pairs["esto_flow"].str.lower()
    pairs["esto_product_norm"] = pairs["esto_product"].str.lower()
    projection_cache: dict[tuple[str, str, str], pd.Series] = {}
    base_cache: dict[tuple[str, str], float] = {}
    scenario_values = {str(v).strip().lower() for v in scenario_map.values()}
    sheet_measure_lookup = _build_sheet_measure_lookup(sheet_map)
    exact_esto_product_lookup = _build_name_to_esto_product(DEFAULT_CODEBOOK)
    explicit_map = explicit_mappings.copy() if explicit_mappings is not None else load_explicit_sector_fuel_mappings()
    if explicit_map.empty:
        explicit_lookup: dict[tuple[str, str, str], list[dict[str, str]]] = {}
    else:
        for col in [
            "sheet_name",
            "fuel_label",
            "sector_code_9th",
            "projection_sector_code",
            "ninth_fuel_code",
            "esto_flow",
            "esto_product",
            "mapping_note",
        ]:
            if col not in explicit_map.columns:
                explicit_map[col] = ""
            explicit_map[col] = explicit_map[col].map(_clean_token)
        explicit_lookup = {}
        for _, row in explicit_map.iterrows():
            key = (
                _normalize_label(row.get("sheet_name", "")),
                _normalize_label(row.get("fuel_label", "")),
                _normalize_label(row.get("sector_code_9th", "")),
            )
            explicit_lookup.setdefault(key, []).append(
                {
                    "projection_sector_code": _clean_token(row.get("projection_sector_code", "")),
                    "ninth_fuel_code": _clean_token(row.get("ninth_fuel_code", "")),
                    "esto_flow": _clean_token(row.get("esto_flow", "")),
                    "esto_product": _clean_token(row.get("esto_product", "")),
                    "mapping_note": _clean_token(row.get("mapping_note", "")),
                }
            )
    ninth_projection_df = _prepare_ninth_projection_frame(
        ninth_df,
        economy_code=projection_economy,
        scenario_values=scenario_values,
    )
    prepared_base_df = _prepare_base_lookup_frame(
        base_df,
        economy_code=base_economy,
        base_year=base_year,
    )
    def _sector_numeric_prefix(code: str) -> str:
        token = str(code or "").strip().lower()
        m = re.match(r"^(\d{2}(?:_\d{2})*)", token)
        return m.group(1) if m else ""

    available_sector_codes: set[str] = set()
    for sector_col in SECTOR_COLUMNS:
        if sector_col not in ninth_projection_df.columns:
            continue
        values = (
            ninth_projection_df[sector_col]
            .dropna()
            .astype(str)
            .str.strip()
            .str.lower()
            .replace("", pd.NA)
            .dropna()
            .tolist()
        )
        for value in values:
            if value != "x":
                available_sector_codes.add(value)

    sector_prefix_to_codes: dict[str, list[str]] = {}
    for sector_code in available_sector_codes:
        prefix = _sector_numeric_prefix(sector_code)
        if not prefix:
            continue
        sector_prefix_to_codes.setdefault(prefix, []).append(sector_code)
    for prefix, codes in sector_prefix_to_codes.items():
        sector_prefix_to_codes[prefix] = sorted(set(codes), key=lambda code: (len(code), code))

    def _nearest_parent_sector_code(sector_code: str) -> str:
        prefix = _sector_numeric_prefix(sector_code)
        if not prefix:
            return ""
        parts = prefix.split("_")
        token = str(sector_code or "").strip().lower()
        for width in range(len(parts) - 1, 0, -1):
            parent_prefix = "_".join(parts[:width])
            candidates = sector_prefix_to_codes.get(parent_prefix, [])
            for candidate in candidates:
                if candidate != token:
                    return candidate
        return ""

    sheet_map_local = sheet_map.copy()
    if "category_type" not in sheet_map_local.columns:
        sheet_map_local["category_type"] = "fuel"
    sheet_map_local["category_type"] = (
        sheet_map_local["category_type"]
        .fillna("fuel")
        .astype(str)
        .str.strip()
        .str.lower()
        .replace({"": "fuel"})
    )
    sheet_map_local["sheet_name_norm"] = sheet_map_local["sheet_name"].map(_normalize_label)
    sheet_rows_by_name = {
        str(row["sheet_name"]).strip(): row
        for _, row in sheet_map_local.drop_duplicates(subset=["sheet_name"], keep="first").iterrows()
    }

    def _normalize_category_key(value: object) -> str:
        text = str(value or "").strip().lower()
        return re.sub(r"[^a-z0-9]+", "", text)

    sheet_map_local["sector_name_norm"] = sheet_map_local.get("sector_name", "").map(_normalize_label)
    sheet_map_local["sector_name_key"] = sheet_map_local.get("sector_name", "").map(_normalize_category_key)
    sheet_map_local["sheet_name_key"] = sheet_map_local["sheet_name"].map(_normalize_category_key)
    sheet_to_sector_names: dict[str, set[str]] = {}
    for _, mapping_row in sheet_map_local.iterrows():
        category_type = str(mapping_row.get("category_type") or "").strip().lower()
        if category_type and category_type != "sector":
            continue
        sheet_label = str(mapping_row.get("sheet_name") or "").strip()
        sector_label = str(mapping_row.get("sector_name") or "").strip()
        if not sheet_label or not sector_label:
            continue
        sheet_to_sector_names.setdefault(sheet_label.lower(), set()).add(sector_label.lower())

    def _is_sheet_parent_alias(sheet_label: object, parent_label: object) -> bool:
        """
        Return True when two labels should be treated as the same logical category.

        This catches cosmetic naming differences such as:
        - LEAP sheet name: "Chemicals"
        - mapped sector label: "Chemical (incl. petrochemical)"
        """
        sheet_text = str(sheet_label or "").strip()
        parent_text = str(parent_label or "").strip()
        if not sheet_text or not parent_text:
            return False

        sheet_norm = _normalize_label(sheet_text)
        parent_norm = _normalize_label(parent_text)
        if sheet_norm == parent_norm:
            return True

        sheet_key = _normalize_category_key(sheet_text)
        parent_key = _normalize_category_key(parent_text)
        if sheet_key and sheet_key == parent_key:
            return True

        mapped_sector_names = sheet_to_sector_names.get(sheet_text.lower(), set())
        return parent_norm in mapped_sector_names

    def _resolve_sector_flow(sector_codes: list[str]) -> str:
        for sector_code in sector_codes:
            key = str(sector_code or "").strip().lower()
            if not key:
                continue
            if key in sector_flow_mapping:
                return sector_flow_mapping[key]
            # Fallback: detailed sector -> nearest mapped parent (e.g., 15_02_01_* -> 15_02_road).
            prefix_match = re.match(r"^(\d{2}_\d{2})_", key)
            if prefix_match:
                prefix = prefix_match.group(1) + "_"
                candidates = [k for k in sector_flow_mapping if k.startswith(prefix)]
                if candidates:
                    best = min(candidates, key=len)
                    return sector_flow_mapping.get(best, "")
        return ""

    def _resolve_exact_sector_flow(sector_codes: list[str]) -> str:
        for sector_code in sector_codes:
            key = str(sector_code or "").strip().lower()
            if key and key in sector_flow_mapping:
                return sector_flow_mapping[key]
        return ""

    def _resolve_sector_category_targets(category_label: str) -> tuple[list[str], list[str], str]:
        if not category_label:
            return [], [], ""

        exact = sheet_map_local[sheet_map_local["sheet_name_norm"] == _normalize_label(category_label)]
        if exact.empty:
            exact = sheet_map_local[sheet_map_local["sheet_name_key"] == _normalize_category_key(category_label)]

        candidate_rows = exact
        match_method = "sheet_name_lookup"
        if candidate_rows.empty:
            candidate_rows = sheet_map_local[sheet_map_local["sector_name_key"] == _normalize_category_key(category_label)]
            match_method = "sector_name_lookup"
        if candidate_rows.empty:
            return [], [], ""

        sector_codes: list[str] = []
        flow_values: list[str] = []
        for _, candidate in candidate_rows.iterrows():
            codes = _split_sector_codes(candidate.get("sector_code_9th"))
            if not codes:
                single = _clean_token(candidate.get("sector_code_9th"))
                codes = [single] if single else []
            for code in codes:
                token = _clean_token(code)
                if token and token not in sector_codes:
                    sector_codes.append(token)
            flow = _clean_token(candidate.get("esto_flow_override"))
            if not flow:
                for code in codes:
                    flow = _resolve_sector_flow([code])
                    if flow:
                        break
            if flow and flow not in flow_values:
                flow_values.append(flow)
        return sector_codes, flow_values, match_method

    def _canonical_targets(df: pd.DataFrame) -> list[tuple[str, str]]:
        if df.empty:
            return []
        rows = (
            df[["esto_flow", "esto_product"]]
            .drop_duplicates()
            .sort_values(["esto_flow", "esto_product"])
            .itertuples(index=False)
        )
        return [(_clean_token(row.esto_flow), _clean_token(row.esto_product)) for row in rows]

    def _display_target_value(values: list[str], *, suffix: str) -> str:
        uniq = [val for val in values if val]
        if not uniq:
            return ""
        if len(uniq) == 1:
            return uniq[0]
        return f"{len(uniq)} {suffix} (aggregated)"

    def _dedupe_tokens(values: Sequence[object]) -> list[str]:
        seen: set[str] = set()
        out: list[str] = []
        for value in values:
            token = _clean_token(value)
            if not token:
                continue
            key = token.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(token)
        return out

    def _dedupe_targets(values: Sequence[tuple[object, object]]) -> list[tuple[str, str]]:
        seen: set[tuple[str, str]] = set()
        out: list[tuple[str, str]] = []
        for raw_flow, raw_product in values:
            flow = _clean_token(raw_flow)
            product = _clean_token(raw_product)
            if not flow and not product:
                continue
            key = (flow.lower(), product.lower())
            if key in seen:
                continue
            seen.add(key)
            out.append((flow, product))
        return out

    def _dedupe_projection_targets(values: Sequence[tuple[object, object]]) -> list[tuple[str, str]]:
        seen: set[tuple[str, str]] = set()
        out: list[tuple[str, str]] = []
        for raw_sector, raw_fuel in values:
            sector = _clean_token(raw_sector)
            fuel = _clean_token(raw_fuel)
            if not sector and not fuel:
                continue
            key = (sector.lower(), fuel.lower())
            if key in seen:
                continue
            seen.add(key)
            out.append((sector, fuel))
        return out

    def _restrict_targets_to_exact_flow(
        targets: Sequence[tuple[object, object]],
        exact_flow: object,
    ) -> list[tuple[str, str]]:
        flow_token = _clean_token(exact_flow)
        normalized = _dedupe_targets(targets)
        if not flow_token or not normalized:
            return normalized
        exact_matches = [(flow, product) for flow, product in normalized if _clean_token(flow) == flow_token]
        return exact_matches if exact_matches else normalized

    def _sum_base_targets(targets: list[tuple[str, str]]) -> float:
        if not targets:
            return float("nan")
        total = 0.0
        saw_value = False
        sign_role = _transformation_sign_role_from_measure(measure_label, sheet)
        for flow_value, product_value in targets:
            cache_key = (flow_value.strip().lower(), product_value.strip().lower(), sign_role)
            if cache_key not in base_cache:
                base_cache[cache_key] = pull_base_year_value(
                    prepared_base_df,
                    base_year=base_year,
                    economy_code=base_economy,
                    esto_flow=flow_value,
                    esto_product=product_value,
                    value_sign_role=sign_role,
                )
            val = base_cache[cache_key]
            if not pd.isna(val):
                total += float(val)
                saw_value = True
        return total if saw_value else float("nan")

    def _base_target_exists(flow_value: object, product_value: object) -> bool:
        flow_token = _clean_token(flow_value).lower()
        product_token = _clean_token(product_value).lower()
        if not flow_token or not product_token:
            return False
        working = prepared_base_df
        if "__flow_norm" in working.columns:
            working = working[working["__flow_norm"] == flow_token]
        else:
            working = working[working["flows"].astype(str).str.strip().str.lower() == flow_token]
        if working.empty:
            return False
        if "__product_norm" in working.columns:
            working = working[working["__product_norm"] == product_token]
        else:
            working = working[working["products"].astype(str).str.strip().str.lower() == product_token]
        return not working.empty

    def _format_scenario_label(value: object) -> str:
        raw = str(value or "").strip()
        if not raw:
            return raw
        lowered = raw.lower()
        if lowered == "reference":
            return "Reference"
        if lowered == "target":
            return "Target"
        return raw

    def _sector_root_group(raw_sector_codes: object) -> str:
        """
        Build a coarse sector-family key so sibling branches (e.g., 15_02_01_*)
        can share one parent comparator without mixing unrelated families.
        """
        codes = _split_sector_codes(raw_sector_codes)
        if not codes:
            return ""
        roots: list[str] = []
        for code in codes:
            token = str(code or "").strip().lower()
            if not token:
                continue
            m3 = re.match(r"^(\d{2}_\d{2}_\d{2})", token)
            if m3:
                roots.append(m3.group(1))
                continue
            m2 = re.match(r"^(\d{2}_\d{2})", token)
            if m2:
                roots.append(m2.group(1))
                continue
            roots.append(token)
        if not roots:
            return ""
        return " | ".join(sorted(set(roots)))

    def _preserve_signed_values(esto_flow: str) -> bool:
        """
        Keep signed values only for flows where input/output direction matters,
        notably TPES/supply and transformation/own-use style balances.
        """
        flow = str(esto_flow or "").strip().lower()
        if not flow:
            return False
        if "tpes" in flow or "total primary energy supply" in flow:
            return True
        # ESTO numbering conventions: 07.* supply, 08.* transfers,
        # 09.* transformation, 10.* own use.
        if flow.startswith(("07", "08", "09", "10")):
            return True
        if "transformation" in flow:
            return True
        return False

    def _numeric_sector_prefix(code: object) -> str:
        token = str(code or "").strip().lower()
        m = re.match(r"^(\d{2}(?:_\d{2})*)", token)
        return m.group(1) if m else ""

    def _common_sector_prefix(codes: Sequence[object]) -> str:
        seqs = [prefix.split("_") for raw in codes if (prefix := _numeric_sector_prefix(raw))]
        if not seqs:
            return ""
        common = list(seqs[0])
        for seq in seqs[1:]:
            keep = 0
            for a, b in zip(common, seq):
                if a != b:
                    break
                keep += 1
            common = common[:keep]
            if not common:
                break
        return "_".join(common)

    def _normalize_sector_codes_list(value: object) -> list[str]:
        if isinstance(value, (list, tuple, set)):
            items = value
        else:
            items = _split_sector_codes(value)
        out: list[str] = []
        seen: set[str] = set()
        for raw in items:
            token = str(raw or "").strip().lower()
            if token and token not in seen:
                seen.add(token)
                out.append(token)
        return out

    def _effective_comparator_key_from_sector_codes(codes: Sequence[object]) -> str:
        normalized = _normalize_sector_codes_list(codes)
        if not normalized:
            return ""
        if len(normalized) == 1:
            return normalized[0]
        return "|".join(sorted(normalized))

    def _merge_sector_codes_lists(series: pd.Series) -> list[str]:
        merged: list[str] = []
        seen: set[str] = set()
        for value in series:
            for code in _normalize_sector_codes_list(value):
                if code not in seen:
                    seen.add(code)
                    merged.append(code)
        return merged

    def _numeric_sector_depth(code: object) -> int:
        prefix = _numeric_sector_prefix(code)
        return len(prefix.split("_")) if prefix else 0

    def _is_sector_ancestor_code(parent_code: object, child_code: object) -> bool:
        parent_prefix = _numeric_sector_prefix(parent_code)
        child_prefix = _numeric_sector_prefix(child_code)
        if not parent_prefix or not child_prefix or parent_prefix == child_prefix:
            return False
        parent_parts = parent_prefix.split("_")
        child_parts = child_prefix.split("_")
        return len(parent_parts) < len(child_parts) and child_parts[: len(parent_parts)] == parent_parts

    def _values_close_enough(left: object, right: object, *, rel_tol: float = 0.02, abs_tol: float = 1e-9) -> bool:
        try:
            lval = float(left)
            rval = float(right)
        except Exception:
            return False
        if pd.isna(lval) or pd.isna(rval):
            return False
        scale = max(abs(lval), abs(rval), 1.0)
        return abs(lval - rval) <= max(abs_tol, rel_tol * scale)

    def _coalesce_duplicate_comparison_rows(comp_frame: pd.DataFrame) -> pd.DataFrame:
        if comp_frame.empty:
            return comp_frame

        frame = comp_frame.copy()
        frame["sheet"] = frame["sheet"].astype(str)
        if "measure" not in frame.columns:
            frame["measure"] = ""
        frame["measure"] = frame["measure"].fillna("").astype(str)
        frame["fuel_label"] = frame["fuel_label"].astype(str)
        frame["scenario"] = frame["scenario"].astype(str)
        frame["source"] = frame["source"].astype(str).str.strip()
        frame["value"] = pd.to_numeric(frame["value"], errors="coerce")
        if "effective_comparator_sector_code" not in frame.columns:
            frame["effective_comparator_sector_code"] = ""
        frame["effective_comparator_sector_code"] = (
            frame["effective_comparator_sector_code"].fillna("").astype(str).str.strip().str.lower()
        )

        key_cols = ["sheet", "measure", "fuel_label", "scenario", "source", "year"]
        duplicate_mask = frame.duplicated(subset=key_cols, keep=False)
        if not duplicate_mask.any():
            return frame

        merged_rows: list[dict[str, object]] = []
        duplicate_groups = frame.loc[duplicate_mask].copy()
        for _, group in duplicate_groups.groupby(key_cols, dropna=False):
            if len(group) == 1:
                merged_rows.append(group.iloc[0].to_dict())
                continue

            template = group.iloc[0].to_dict()
            source_token = str(template.get("source", "") or "").strip().lower()
            non_null_values = pd.to_numeric(group["value"], errors="coerce").dropna()

            if source_token in {"base", "base_estimated", "base_mixed"} and len(non_null_values) and non_null_values.nunique() == 1:
                template["value"] = float(non_null_values.iloc[0])
                template["effective_comparator_sector_code"] = ""
                template["sector_codes_list"] = _normalize_sector_codes_list(template.get("sector_codes_list", []))
                merged_rows.append(template)
                continue

            if source_token == "leap":
                template["value"] = pd.to_numeric(group["value"], errors="coerce").sum(min_count=1)
                template["effective_comparator_sector_code"] = ""
                template["sector_codes_list"] = _normalize_sector_codes_list(template.get("sector_codes_list", []))
                merged_rows.append(template)
                continue

            identified = group[group["effective_comparator_sector_code"].ne("")].copy()
            unidentified = group[group["effective_comparator_sector_code"].eq("")].copy()
            total_value = 0.0
            has_value = False

            if not identified.empty:
                by_identity = (
                    identified.groupby("effective_comparator_sector_code", as_index=False)
                    .agg(
                        sum_value=("value", lambda s: s.sum(min_count=1)),
                        first_non_null=("value", lambda s: s.dropna().iloc[0] if len(s.dropna()) else float("nan")),
                        unique_non_null=("value", lambda s: s.dropna().nunique()),
                        contributing_rows=("value", "size"),
                        sector_codes_list=("sector_codes_list", _merge_sector_codes_lists),
                    )
                )
                collapse_mask = (
                    by_identity["unique_non_null"].le(1)
                    & pd.to_numeric(by_identity["contributing_rows"], errors="coerce").fillna(1).gt(1)
                )
                by_identity["value"] = by_identity["sum_value"]
                by_identity.loc[collapse_mask, "value"] = by_identity.loc[collapse_mask, "first_non_null"]

                keep_mask = pd.Series(True, index=by_identity.index)
                group_slice = by_identity.copy()
                group_slice["__depth"] = group_slice["effective_comparator_sector_code"].map(_numeric_sector_depth)
                group_slice = group_slice.sort_values(["__depth", "effective_comparator_sector_code"])
                codes = group_slice["effective_comparator_sector_code"].fillna("").astype(str).tolist()
                values_by_code = {
                    str(row["effective_comparator_sector_code"]): float(row["value"])
                    for _, row in group_slice.iterrows()
                }
                sector_lists_by_code = {
                    str(row["effective_comparator_sector_code"]): _normalize_sector_codes_list(row.get("sector_codes_list", []))
                    for _, row in group_slice.iterrows()
                }
                group_keep = {code: True for code in codes if code}

                for code in codes:
                    if not code or not group_keep.get(code, True):
                        continue
                    desc_codes = [
                        other
                        for other in codes
                        if (
                            other
                            and other != code
                            and group_keep.get(other, True)
                            and (
                                _is_sector_ancestor_code(code, other)
                                or (
                                    sector_lists_by_code.get(other)
                                    and all(_is_sector_ancestor_code(code, member) for member in sector_lists_by_code.get(other, []))
                                )
                            )
                        )
                    ]
                    if not desc_codes:
                        continue
                    ancestor_value = values_by_code.get(code)
                    descendant_sum = sum(values_by_code.get(other, 0.0) for other in desc_codes)
                    if _values_close_enough(descendant_sum, ancestor_value):
                        group_keep[code] = False
                    else:
                        for other in desc_codes:
                            group_keep[other] = False

                drop_codes = [code for code, keep in group_keep.items() if not keep]
                if drop_codes:
                    keep_mask.loc[by_identity["effective_comparator_sector_code"].isin(drop_codes)] = False
                by_identity = by_identity[keep_mask].copy()
                identified_value = pd.to_numeric(by_identity["value"], errors="coerce").sum(min_count=1)
                if pd.notna(identified_value):
                    total_value += float(identified_value)
                    has_value = True

            if not unidentified.empty:
                unidentified_values = pd.to_numeric(unidentified["value"], errors="coerce").dropna()
                if len(unidentified_values):
                    if unidentified_values.nunique() == 1:
                        total_value += float(unidentified_values.iloc[0])
                    else:
                        total_value += float(unidentified_values.sum(min_count=1))
                    has_value = True

            template["value"] = total_value if has_value else float("nan")
            template["effective_comparator_sector_code"] = ""
            template["sector_codes_list"] = _normalize_sector_codes_list(template.get("sector_codes_list", []))
            merged_rows.append(template)

        merged = pd.DataFrame(merged_rows)
        passthrough = frame.loc[~duplicate_mask].copy()
        return pd.concat([passthrough, merged], ignore_index=True, sort=False)

    def _allows_descendant_canonical_match(sector_code: object) -> bool:
        token = str(sector_code or "").strip().lower()
        if not token:
            return False
        # Limit descendant expansion to high-level parent sectors such as
        # 09_01_electricity_plants and 18_01_electricity_plants. Applying the
        # same rule to detailed process leaves (for example 09_08_04_*) causes
        # unrelated canonical rows to bleed across transformation branches.
        return _numeric_sector_depth(token) <= 2

    def _descendant_sector_subset(df: pd.DataFrame, sector_code: str) -> pd.DataFrame:
        key = str(sector_code or "").strip().lower()
        if not key or df.empty:
            return df.iloc[0:0].copy()
        if not _allows_descendant_canonical_match(key):
            return df.iloc[0:0].copy()

        key_prefix = _numeric_sector_prefix(key)
        if not key_prefix:
            return df.iloc[0:0].copy()

        working = df.copy()
        if "sector_prefix_norm" not in working.columns:
            working["sector_prefix_norm"] = working["sector_norm"].map(_numeric_sector_prefix)
        subset = working[
            working["sector_prefix_norm"].str.startswith(key_prefix + "_")
        ].copy()
        if subset.empty:
            return subset

        subset["prefix_depth"] = subset["sector_prefix_norm"].map(_numeric_sector_depth)
        deepest = int(subset["prefix_depth"].max())
        subset = subset[subset["prefix_depth"] == deepest].copy()
        subset["derived_sector_match_method"] = shared_build_code_match_method(
            max(deepest - _numeric_sector_depth(key_prefix), 0)
        )
        return subset.drop(columns=["prefix_depth"], errors="ignore")

    def _sector_match_subset(df: pd.DataFrame, sector_codes: list[str]) -> pd.DataFrame:
        if df.empty:
            return df
        scored: list[pd.DataFrame] = []
        for sector in sector_codes:
            key = str(sector or "").strip().lower()
            if not key:
                continue
            exact = df[df["sector_norm"] == key].copy()
            if not exact.empty:
                exact["match_priority"] = 0
                scored.append(exact)
            child = df[df["sector_norm"].str.startswith(key + "_")].copy()
            if not child.empty:
                child["match_priority"] = 1
                scored.append(child)
            descendant = _descendant_sector_subset(df, key)
            if not descendant.empty:
                descendant["match_priority"] = 1
                scored.append(descendant)
        if not scored:
            return df.iloc[0:0].copy()
        merged = pd.concat(scored, ignore_index=True).drop_duplicates()
        best = int(merged["match_priority"].min())
        merged = merged[merged["match_priority"] == best].drop(columns=["match_priority"], errors="ignore")
        if "derived_sector_match_method" in merged.columns:
            if "sector_match_method" not in merged.columns:
                merged["sector_match_method"] = ""
            merged["sector_match_method"] = merged["sector_match_method"].where(
                merged["sector_match_method"].astype(str).str.strip().ne(""),
                merged["derived_sector_match_method"],
            )
            merged = merged.drop(columns=["derived_sector_match_method"], errors="ignore")
        return merged

    def _canonical_by_sector_and_fuel(sector_codes: list[str], fuel_code: str) -> pd.DataFrame:
        if not fuel_code:
            return pd.DataFrame(columns=pairs.columns)
        subset = pairs[pairs["fuel_norm"] == fuel_code.strip().lower()]
        return _sector_match_subset(subset, sector_codes)

    def _canonical_by_sector_and_fuel_with_parent_fallback(
        sector_codes: list[str], fuel_code: str
    ) -> tuple[pd.DataFrame, list[str]]:
        direct = _canonical_by_sector_and_fuel(sector_codes, fuel_code)
        if not direct.empty:
            return direct, []

        parent_codes: list[str] = []
        match_frames: list[pd.DataFrame] = []
        for sector_code in sector_codes:
            seen: set[str] = set()
            parent_code = _nearest_parent_sector_code(sector_code)
            while parent_code and parent_code not in seen:
                seen.add(parent_code)
                parent_match = _canonical_by_sector_and_fuel([parent_code], fuel_code)
                if not parent_match.empty:
                    match_frames.append(parent_match.copy())
                    parent_codes.append(parent_code)
                    break
                parent_code = _nearest_parent_sector_code(parent_code)

        if not match_frames:
            return pd.DataFrame(columns=pairs.columns), []
        merged = pd.concat(match_frames, ignore_index=True).drop_duplicates()
        return merged, sorted(set(parent_codes))

    def _canonical_by_sector_and_product(sector_codes: list[str], esto_product: str) -> pd.DataFrame:
        if not esto_product:
            return pd.DataFrame(columns=pairs.columns)
        subset = pairs[pairs["esto_product_norm"] == esto_product.strip().lower()]
        return _sector_match_subset(subset, sector_codes)

    def _choose_single_candidate(df: pd.DataFrame) -> tuple[str, str, str, str]:
        if df.empty:
            return "", "", "", ""
        unique_pairs = df[["9th_fuel", "esto_flow", "esto_product"]].drop_duplicates().sort_values(
            ["9th_fuel", "esto_flow", "esto_product"]
        )
        if len(unique_pairs) != 1:
            return "", "", "", ""
        row = unique_pairs.iloc[0]
        match = df[
            (df["9th_fuel"] == row["9th_fuel"])
            & (df["esto_flow"] == row["esto_flow"])
            & (df["esto_product"] == row["esto_product"])
        ].copy()
        if not match.empty and "sector_match_method" in match.columns:
            match = match.sort_values(["sector_match_method", "fuel_match_method", "mapping_note"], na_position="last")
            method = _clean_token(match.iloc[0].get("sector_match_method"))
        else:
            method = ""
        return _clean_token(row["9th_fuel"]), _clean_token(row["esto_flow"]), _clean_token(row["esto_product"]), method

    def _infer_fuel_from_flow_product(sector_codes: list[str], esto_flow: str, esto_product: str) -> str:
        if not esto_flow or not esto_product:
            return ""
        subset = pairs[
            (pairs["esto_flow_norm"] == esto_flow.strip().lower())
            & (pairs["esto_product_norm"] == esto_product.strip().lower())
        ]
        match = _sector_match_subset(subset, sector_codes)
        vals = sorted(match["9th_fuel"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())
        if len(vals) == 1:
            return vals[0]
        # Fallback 1: global flow+product unique fuel across canonical pairs.
        vals_global = sorted(
            subset["9th_fuel"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist()
        )
        if len(vals_global) == 1:
            return vals_global[0]
        # Fallback 2: product-only unique fuel (independent of flow/sector).
        prod_subset = pairs[pairs["esto_product_norm"] == esto_product.strip().lower()]
        vals_prod = sorted(
            prod_subset["9th_fuel"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist()
        )
        if len(vals_prod) == 1:
            return vals_prod[0]
        return ""

    def _derive_explicit_projection_targets(
        entries: list[dict[str, str]],
        sector_codes: list[str],
    ) -> list[tuple[str, str]]:
        derived: list[tuple[str, str]] = []
        for row in entries:
            fuel_code = _clean_token(row.get("ninth_fuel_code", ""))
            explicit_sector = _clean_token(row.get("projection_sector_code", ""))
            if explicit_sector and fuel_code:
                derived.append((explicit_sector, fuel_code))
                continue
            if not fuel_code or not sector_codes:
                continue
            candidates, _ = _canonical_by_sector_and_fuel_with_parent_fallback(sector_codes, fuel_code)
            if candidates.empty:
                continue
            flow = _clean_token(row.get("esto_flow", ""))
            product = _clean_token(row.get("esto_product", ""))
            narrowed = candidates
            if flow:
                narrowed = narrowed[narrowed["esto_flow_norm"] == flow.lower()]
            if product:
                narrowed_by_product = narrowed[narrowed["esto_product_norm"] == product.lower()]
                if not narrowed_by_product.empty:
                    narrowed = narrowed_by_product
            elif flow and narrowed.empty:
                narrowed = candidates[candidates["esto_flow_norm"] == flow.lower()]
            if narrowed.empty:
                narrowed = candidates
            sector_values = _dedupe_tokens(narrowed.get("9th_sector", pd.Series(dtype=str)).tolist())
            derived.extend((sector_value, fuel_code) for sector_value in sector_values)
        return _dedupe_projection_targets(derived)

    def _build_explicit_mapping_summary(
        entries: list[dict[str, str]],
        *,
        sector_codes: list[str],
    ) -> dict[str, object]:
        if not entries:
            return {}
        projection_fuel_codes = _dedupe_tokens([row.get("ninth_fuel_code", "") for row in entries])
        projection_targets = _derive_explicit_projection_targets(entries, sector_codes)
        base_targets = _dedupe_targets(
            [(row.get("esto_flow", ""), row.get("esto_product", "")) for row in entries]
        )
        flow_values = _dedupe_tokens([flow for flow, _ in base_targets])
        product_values = _dedupe_tokens([product for _, product in base_targets])
        note_values: list[str] = []
        allow_missing_base = False
        for row in entries:
            cleaned_note, row_allow_missing_base = _extract_allow_missing_base_flag(row.get("mapping_note", ""))
            allow_missing_base = allow_missing_base or row_allow_missing_base
            if cleaned_note:
                note_values.append(cleaned_note)
        note_values = _dedupe_tokens(note_values)
        mapping_note = "; ".join(note_values)
        if len(entries) > 1 or len(projection_fuel_codes) > 1 or len(base_targets) > 1:
            agg_note = "aggregated explicit targets"
            mapping_note = f"{mapping_note}; {agg_note}" if mapping_note else agg_note
        return {
            "ninth_fuel_code": _display_target_value(projection_fuel_codes, suffix="fuels"),
            "esto_flow": _display_target_value(flow_values, suffix="flows"),
            "esto_product": _display_target_value(product_values, suffix="products"),
            "mapping_note": mapping_note,
            "projection_fuel_codes": projection_fuel_codes,
            "projection_targets": projection_targets,
            "base_targets": base_targets,
            "allow_missing_base": allow_missing_base,
        }

    def _get_explicit_mapping(
        sheet_name: str,
        fuel_label: str,
        sector_code_text: str,
        sector_codes: list[str],
    ) -> dict[str, object]:
        key = (
            _normalize_label(sheet_name),
            _normalize_label(fuel_label),
            _normalize_label(sector_code_text),
        )
        if key in explicit_lookup:
            return _build_explicit_mapping_summary(explicit_lookup[key], sector_codes=sector_codes)
        wildcard_key = (
            _normalize_label(sheet_name),
            _normalize_label(fuel_label),
            "",
        )
        return _build_explicit_mapping_summary(explicit_lookup.get(wildcard_key, []), sector_codes=sector_codes)

    def _split_fuel_filters(raw_value: object) -> list[str]:
        text = str(raw_value or "").strip()
        if not text or text.lower() == "nan":
            return []
        parts = re.split(r"\s*(?:,|;|\||\band\b)\s*", text, flags=re.IGNORECASE)
        seen: set[str] = set()
        out: list[str] = []
        for part in parts:
            token = _clean_token(part)
            if not token:
                continue
            key = token.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(token)
        return out

    def _canonical_targets_for_fuel_filters(sector_codes: list[str], fuel_filters: list[str]) -> list[tuple[str, str]]:
        if not sector_codes or not fuel_filters:
            return []
        frames: list[pd.DataFrame] = []
        for fuel_code in fuel_filters:
            c_by_fuel, _ = _canonical_by_sector_and_fuel_with_parent_fallback(sector_codes, fuel_code)
            if not c_by_fuel.empty:
                frames.append(c_by_fuel)
        if not frames:
            return []
        merged = pd.concat(frames, ignore_index=True).drop_duplicates()
        return _canonical_targets(merged)

    for (sheet, fuel), sub in leap_long.groupby(["sheet_name", "fuel_label"], dropna=False):
        sheet_key = str(sheet or "").strip()
        sheet_row = sheet_rows_by_name.get(sheet_key)
        if sheet_row is None:
            print(f"[WARN] Skipping unmapped LEAP sheet during comparison build: {sheet_key!r}")
            continue
        sheet_sector_codes = _split_sector_codes(sheet_row.get("sector_code_9th"))
        if not sheet_sector_codes:
            single = _clean_token(sheet_row.get("sector_code_9th"))
            sheet_sector_codes = [single] if single else []
        comparison_sector_codes = list(sheet_sector_codes)
        projection_fuel_filters = _split_fuel_filters(sheet_row.get("projection_fuel_filter", ""))
        explicit_sector_code_text = " | ".join(sheet_sector_codes)
        category_type = _clean_token(sheet_row.get("category_type")) or "fuel"
        sector_code_text = " | ".join(comparison_sector_codes)
        sheet_flow_override = ""
        if "esto_flow_override" in sheet_map.columns:
            sheet_flow_override = _clean_token(sheet_row.get("esto_flow_override"))
        feedstock_sector_override, feedstock_flow_override = _power_feedstock_output_override(sheet)
        if feedstock_sector_override:
            comparison_sector_codes = list(feedstock_sector_override)
            sector_code_text = " | ".join(comparison_sector_codes)
            explicit_sector_code_text = sector_code_text
        if feedstock_flow_override and not sheet_flow_override:
            sheet_flow_override = feedstock_flow_override

        explicit_entry = _get_explicit_mapping(sheet, fuel, explicit_sector_code_text, comparison_sector_codes)
        explicit_projection_fuel_codes: list[str] = []
        explicit_projection_targets: list[tuple[str, str]] = []
        base_targets: list[tuple[str, str]] = []
        allow_missing_base = False
        if explicit_entry:
            explicit_projection_fuel_codes = list(explicit_entry.get("projection_fuel_codes") or [])
            explicit_projection_targets = list(explicit_entry.get("projection_targets") or [])
            base_targets = list(explicit_entry.get("base_targets") or [])
            allow_missing_base = bool(explicit_entry.get("allow_missing_base"))
            ninth_fuel = _clean_token(explicit_entry.get("ninth_fuel_code"))
            esto_flow = _clean_token(explicit_entry.get("esto_flow"))
            esto_product = _clean_token(explicit_entry.get("esto_product"))
            esto_product_hint = esto_product
            esto_flow_hint = esto_flow
            mapping_source = "explicit"
            flow_source = "explicit" if esto_flow else ""
            fuel_source = "explicit" if ninth_fuel else ""
            sector_match_method = "manual_override"
            mapping_note = _clean_token(explicit_entry.get("mapping_note"))
        elif category_type == "sector":
            category_sector_codes, category_flows, category_match_method = _resolve_sector_category_targets(fuel)
            if category_sector_codes:
                comparison_sector_codes = category_sector_codes
                sector_code_text = " | ".join(comparison_sector_codes)
            ninth_fuel = ""
            esto_flow = _display_target_value(category_flows, suffix="flows")
            esto_product = ""
            esto_product_hint = ""
            esto_flow_hint = ""
            mapping_source = "category_sector" if comparison_sector_codes else ""
            flow_source = "category_sector" if category_flows else ""
            fuel_source = ""
            sector_match_method = category_match_method if comparison_sector_codes else ""
            mapping_note = "category labels treated as sectors" if comparison_sector_codes else ""
            base_targets = [(flow_value, "") for flow_value in category_flows]
        else:
            mapped_hint = map_fuel_label(fuel, fuel_mapping)
            ninth_fuel = _clean_token(mapped_hint.get("ninth_fuel"))
            esto_product_hint = _clean_token(mapped_hint.get("esto_product"))
            exact_esto_product_hint = _clean_token(exact_esto_product_lookup.get(_normalize_label(fuel), ""))
            if exact_esto_product_hint:
                esto_product_hint = exact_esto_product_hint
            esto_flow_hint = _clean_token(mapped_hint.get("esto_flow"))
            mapping_source = _clean_token(mapped_hint.get("mapping_source")) or "canonical"
            flow_source = _clean_token(mapped_hint.get("flow_source"))
            fuel_source = _clean_token(mapped_hint.get("fuel_source"))
            sector_match_method = ""
            mapping_note = ""
            if not ninth_fuel:
                label_fallback = _infer_fuel_from_label_fallback(fuel)
                if label_fallback:
                    ninth_fuel = label_fallback
                    if fuel_source != "override":
                        fuel_source = "inferred"
            exact_sector_flow = sheet_flow_override or _resolve_exact_sector_flow(comparison_sector_codes)
            exact_sheet_level_target: list[tuple[str, str]] = []
            if exact_sector_flow and esto_product_hint and _base_target_exists(exact_sector_flow, esto_product_hint):
                exact_sheet_level_target = [(_clean_token(exact_sector_flow), _clean_token(esto_product_hint))]
            c_by_fuel, parent_fallback_codes = _canonical_by_sector_and_fuel_with_parent_fallback(
                comparison_sector_codes, ninth_fuel
            )
            fuel_match_count = len(c_by_fuel[["9th_fuel", "esto_flow", "esto_product"]].drop_duplicates()) if not c_by_fuel.empty else 0
            c_targets = _canonical_targets(c_by_fuel)
            c_targets = _restrict_targets_to_exact_flow(c_targets, exact_sector_flow)
            c_ninth, c_flow, c_prod, c_sector_method = _choose_single_candidate(c_by_fuel)
            c_by_product = _canonical_by_sector_and_product(comparison_sector_codes, esto_product_hint)
            prod_match_count = (
                len(c_by_product[["9th_fuel", "esto_flow", "esto_product"]].drop_duplicates()) if not c_by_product.empty else 0
            )
            p_ninth, p_flow, p_prod, p_sector_method = _choose_single_candidate(c_by_product)

            if c_ninth and c_flow and c_prod:
                ninth_fuel = c_ninth
                sector_match_method = c_sector_method
                mapping_source = "override" if mapping_source == "override" else "canonical"
                if fuel_source != "override":
                    fuel_source = "canonical"
                if parent_fallback_codes:
                    fallback_note = "canonical parent fallback via " + ", ".join(parent_fallback_codes)
                    mapping_note = f"{mapping_note}; {fallback_note}" if mapping_note else fallback_note
                if exact_sheet_level_target and exact_sheet_level_target[0] != (c_flow, c_prod):
                    esto_flow, esto_product = exact_sheet_level_target[0]
                    if flow_source != "override":
                        flow_source = "sheet_exact_level"
                    mapping_note = _merge_mapping_note(
                        mapping_note,
                        "exact sheet-level ESTO target retained for base while projection fuel resolves via canonical comparator mapping",
                    )
                    base_targets = exact_sheet_level_target
                else:
                    esto_flow, esto_product = c_flow, c_prod
                    if flow_source != "override":
                        flow_source = "canonical"
                    base_targets = [(esto_flow, esto_product)]
            elif p_ninth and p_flow and p_prod:
                # Prefer unique sector+product resolution before aggregating broad fuel buckets.
                ninth_fuel = ninth_fuel or p_ninth
                esto_flow = _clean_token(esto_flow_hint) or p_flow
                esto_product = _clean_token(esto_product_hint) or p_prod
                sector_match_method = p_sector_method
                if mapping_source != "override":
                    mapping_source = "codebook_fallback"
                if flow_source != "override":
                    flow_source = "canonical"
                fuel_source = "override" if fuel_source == "override" else "inferred"
                base_targets = [(esto_flow, esto_product)]
            elif fuel_match_count > 1 and ninth_fuel and c_targets:
                if exact_sheet_level_target:
                    esto_flow, esto_product = exact_sheet_level_target[0]
                    sector_match_method = "exact_sheet_level_target"
                    mapping_source = "override" if mapping_source == "override" else "canonical"
                    if flow_source != "override":
                        flow_source = "sheet_exact_level"
                    if fuel_source != "override":
                        fuel_source = "canonical"
                    mapping_note = (
                        "exact sheet-level ESTO target used instead of aggregated parent fallback"
                        + (f"; canonical parent fallback via {', '.join(parent_fallback_codes)}" if parent_fallback_codes else "")
                    )
                    base_targets = exact_sheet_level_target
                else:
                    flow_values = sorted({flow_value for flow_value, _ in c_targets if flow_value})
                    product_values = sorted({product_value for _, product_value in c_targets if product_value})
                    esto_flow = _display_target_value(flow_values, suffix="flows")
                    esto_product = _display_target_value(product_values, suffix="products")
                    sector_match_method = "aggregated_canonical_targets"
                    mapping_source = "override" if mapping_source == "override" else "canonical_aggregated"
                    if flow_source != "override":
                        flow_source = "canonical_aggregated"
                    if fuel_source != "override":
                        fuel_source = "canonical"
                    mapping_note = "aggregated canonical targets for sector+fuel conflict"
                    if parent_fallback_codes:
                        mapping_note = f"{mapping_note}; canonical parent fallback via {', '.join(parent_fallback_codes)}"
                    base_targets = c_targets
            else:
                if fuel_match_count > 1 and ninth_fuel:
                    mapping_note = "ambiguous canonical matches for sector+fuel"
                if prod_match_count > 1 and not c_by_product.empty:
                    product_fuels = (
                        c_by_product["9th_fuel"]
                        .astype(str)
                        .map(_clean_token)
                        .replace("", pd.NA)
                        .dropna()
                        .drop_duplicates()
                        .tolist()
                    )
                    product_targets = _canonical_targets(c_by_product)
                    product_targets = _restrict_targets_to_exact_flow(product_targets, exact_sector_flow)
                    if len(product_fuels) == 1 and product_targets:
                        if exact_sheet_level_target:
                            ninth_fuel = product_fuels[0]
                            esto_flow, esto_product = exact_sheet_level_target[0]
                            sector_match_method = "exact_sheet_level_target"
                            mapping_source = "override" if mapping_source == "override" else "canonical"
                            if flow_source != "override":
                                flow_source = "sheet_exact_level"
                            if fuel_source != "override":
                                fuel_source = "canonical"
                            mapping_note = "exact sheet-level ESTO target used instead of aggregated parent fallback"
                            base_targets = exact_sheet_level_target
                        else:
                        # Deterministic aggregate path:
                        # same sector+product maps to multiple flows but one 9th fuel.
                            ninth_fuel = product_fuels[0]
                            flow_values = sorted({flow_value for flow_value, _ in product_targets if flow_value})
                            product_values = sorted({product_value for _, product_value in product_targets if product_value})
                            esto_flow = _display_target_value(flow_values, suffix="flows")
                            esto_product = _display_target_value(product_values, suffix="products")
                            sector_match_method = "aggregated_canonical_targets"
                            mapping_source = "override" if mapping_source == "override" else "canonical_aggregated"
                            if flow_source != "override":
                                flow_source = "canonical_aggregated"
                            if fuel_source != "override":
                                fuel_source = "canonical"
                            mapping_note = "aggregated canonical targets for sector+esto_product conflict"
                            base_targets = product_targets
                else:
                    if prod_match_count > 1 and esto_product_hint:
                        mapping_note = "ambiguous canonical matches for sector+esto_product"
                    esto_flow = _clean_token(esto_flow_hint)
                    esto_product = _clean_token(esto_product_hint)
                    if exact_sheet_level_target:
                        esto_flow, esto_product = exact_sheet_level_target[0]
                        sector_match_method = "exact_sheet_level_target"
                        if mapping_source != "override":
                            mapping_source = "codebook_fallback"
                        if flow_source != "override":
                            flow_source = "sheet_exact_level"
                        if fuel_source != "override" and ninth_fuel:
                            fuel_source = "canonical" if ninth_fuel else fuel_source
                        mapping_note = (
                            "exact sheet-level ESTO target used instead of aggregated parent fallback"
                            + (f"; canonical parent fallback via {', '.join(parent_fallback_codes)}" if parent_fallback_codes else "")
                        )
                        base_targets = exact_sheet_level_target
                    if sheet_flow_override:
                        esto_flow = sheet_flow_override
                        flow_source = "sheet_override"
                    elif not esto_flow:
                        fallback_flow = _resolve_sector_flow(comparison_sector_codes)
                        if fallback_flow:
                            esto_flow = _clean_token(fallback_flow)
                            flow_source = "sector_fallback"
                    if not ninth_fuel:
                        inferred = _infer_fuel_from_flow_product(comparison_sector_codes, esto_flow, esto_product)
                        if inferred:
                            ninth_fuel = inferred
                            fuel_source = "inferred"
                    if ninth_fuel and not base_targets:
                        late_c_by_fuel, late_parent_fallback_codes = _canonical_by_sector_and_fuel_with_parent_fallback(
                            comparison_sector_codes, ninth_fuel
                        )
                        late_targets = _canonical_targets(late_c_by_fuel)
                        late_targets = _restrict_targets_to_exact_flow(late_targets, exact_sector_flow)
                        late_ninth, late_flow, late_prod, late_sector_method = _choose_single_candidate(late_c_by_fuel)
                        if late_ninth and late_flow and late_prod:
                            ninth_fuel = late_ninth
                            sector_match_method = late_sector_method
                            mapping_source = "override" if mapping_source == "override" else "canonical"
                            if fuel_source != "override":
                                fuel_source = "canonical"
                            if late_parent_fallback_codes:
                                note = "canonical parent fallback via " + ", ".join(late_parent_fallback_codes)
                                mapping_note = f"{mapping_note}; {note}" if mapping_note else note
                            if exact_sheet_level_target and exact_sheet_level_target[0] != (late_flow, late_prod):
                                esto_flow, esto_product = exact_sheet_level_target[0]
                                if flow_source != "override":
                                    flow_source = "sheet_exact_level"
                                mapping_note = _merge_mapping_note(
                                    mapping_note,
                                    "exact sheet-level ESTO target retained for base while projection fuel resolves via canonical comparator mapping",
                                )
                                base_targets = exact_sheet_level_target
                            else:
                                esto_flow = late_flow
                                esto_product = late_prod
                                if flow_source != "override":
                                    flow_source = "canonical"
                                base_targets = [(esto_flow, esto_product)]
                        elif len(late_targets) > 1:
                            flow_values = sorted({flow_value for flow_value, _ in late_targets if flow_value})
                            product_values = sorted({product_value for _, product_value in late_targets if product_value})
                            esto_flow = _display_target_value(flow_values, suffix="flows")
                            esto_product = _display_target_value(product_values, suffix="products")
                            sector_match_method = "aggregated_canonical_targets"
                            mapping_source = "override" if mapping_source == "override" else "canonical_aggregated"
                            if flow_source != "override":
                                flow_source = "canonical_aggregated"
                            if fuel_source != "override":
                                fuel_source = "canonical"
                            mapping_note = "aggregated canonical targets for sector+fuel conflict"
                            if late_parent_fallback_codes:
                                mapping_note = f"{mapping_note}; canonical parent fallback via {', '.join(late_parent_fallback_codes)}"
                            base_targets = late_targets
                    if not mapping_source:
                        mapping_source = "override" if fuel_source == "override" else "codebook_fallback"
                    # For fuel mappings, do not create product-only base targets.
                    # Product-only pulls can sum across all ESTO flows and inflate values.
                    if esto_flow and esto_product:
                        if not base_targets:
                            base_targets = [(esto_flow, esto_product)]

        if projection_fuel_filters and not explicit_entry:
            filtered_targets = _canonical_targets_for_fuel_filters(comparison_sector_codes, projection_fuel_filters)
            if filtered_targets:
                base_targets = filtered_targets
                if mapping_source != "explicit":
                    mapping_source = "canonical_aggregated"
                if flow_source != "override":
                    flow_source = "canonical_aggregated"
                note = f"projection_fuel_filter applied ({'|'.join(projection_fuel_filters)})"
                mapping_note = f"{mapping_note}; {note}" if mapping_note else note

        if not flow_source and esto_flow:
            flow_source = "canonical"
        if not fuel_source and ninth_fuel:
            fuel_source = "canonical"
        if category_type == "sector" and not explicit_entry:
            has_any_mapping = bool(comparison_sector_codes or base_targets)
            base_mapping_complete = bool(base_targets)
            projection_mapping_complete = bool(comparison_sector_codes)
        else:
            has_any_mapping = bool(ninth_fuel or esto_flow or esto_product or allow_missing_base)
            base_mapping_complete = bool(base_targets) or allow_missing_base
            projection_mapping_complete = bool(ninth_fuel or projection_fuel_filters)
        mapped_flag = bool(base_mapping_complete and projection_mapping_complete)
        partially_mapped = bool(has_any_mapping and not mapped_flag)
        if base_targets:
            preserve_sign = any(_preserve_signed_values(flow_value) for flow_value, _ in base_targets)
        else:
            preserve_sign = _preserve_signed_values(esto_flow)
        base_targets_detail = (
            json.dumps(
                [
                    {"esto_flow": flow_value, "esto_product": product_value}
                    for flow_value, product_value in base_targets
                    if flow_value or product_value
                ],
                ensure_ascii=True,
            )
            if base_targets
            else ""
        )
        projection_fuel_codes_detail = (
            json.dumps([code for code in explicit_projection_fuel_codes if code], ensure_ascii=True)
            if explicit_projection_fuel_codes
            else ""
        )
        projection_targets_detail = (
            json.dumps(
                [
                    {"sector_code_9th": sector_code, "ninth_fuel_code": fuel_code}
                    for sector_code, fuel_code in explicit_projection_targets
                    if sector_code or fuel_code
                ],
                ensure_ascii=True,
            )
            if explicit_projection_targets
            else ""
        )
        source_measure_values = (
            sub.get("measure", pd.Series(dtype=str))
            .fillna("")
            .astype(str)
            .str.strip()
        )
        source_measure_values = [value for value in source_measure_values.unique().tolist() if value]
        sheet_measure_override = _clean_token(sheet_row.get("measure"))
        measure_label = sheet_measure_override or (source_measure_values[0] if source_measure_values else "") or sheet_measure_lookup.get(sheet, "")

        status_row = {
            "sheet": sheet,
            "fuel_label": fuel,
            "measure": measure_label,
            "sector_code_9th": sector_code_text,
            "ninth_fuel_code": ninth_fuel,
            "esto_flow": esto_flow,
            "esto_product": esto_product,
            "has_any_mapping": has_any_mapping,
            "base_mapping_complete": base_mapping_complete,
            "projection_mapping_complete": projection_mapping_complete,
            "partially_mapped": partially_mapped,
            "mapped": mapped_flag,
            "mapping_source": mapping_source or "",
            "flow_source": flow_source or "",
            "fuel_source": fuel_source or "",
            "sector_match_method": sector_match_method or "",
            "mapping_note": mapping_note,
            "projection_fuel_filter": " | ".join(projection_fuel_filters),
            "projection_fuel_codes_detail": projection_fuel_codes_detail,
            "projection_targets_detail": projection_targets_detail,
            "base_targets_detail": base_targets_detail,
            "projection_parent_fallback": False,
            "projection_parent_sector_code": "",
            "comparator_scope": "child",
            "base_mapping_optional": allow_missing_base,
        }
        status_rows.append(status_row)
        status_idx = len(status_rows) - 1
        transformation_sign_role = _transformation_sign_role_from_measure(measure_label, sheet)
        display_inputs_as_positive = _measure_is_input_only(measure_label) or (
            transformation_sign_role == "input" and _measure_is_output_only(measure_label)
        )
        display_exports_as_positive = _sheet_is_export_flow(sheet, measure_label)

        sub_with_scenario = sub.copy()
        sub_with_scenario["scenario_display"] = sub_with_scenario["scenario"].map(_format_scenario_label)
        leap_scenarios = [
            str(value).strip()
            for value in sub_with_scenario["scenario_display"].dropna().astype(str).tolist()
            if str(value).strip()
        ]
        mapped_scenarios = [
            _format_scenario_label(str(raw or "").strip())
            for raw in list((scenario_map or {}).keys())
            if str(raw or "").strip()
        ]
        scenario_order: list[str] = []
        for scen in leap_scenarios + mapped_scenarios:
            if scen and scen not in scenario_order:
                scenario_order.append(scen)

        for scenario_display in scenario_order:
            projection_scenario = scenario_map.get(str(scenario_display).lower(), "reference")
            sub_scenario = sub_with_scenario[sub_with_scenario["scenario_display"] == scenario_display]

            # LEAP series (only where LEAP rows exist for this scenario)
            if not sub_scenario.empty:
                leap_economy_raw = normalize_economy_key(sub_scenario["economy"].iloc[0])
                leap_economy = leap_economy_raw
                raw_token = str(leap_economy_raw or "").strip().upper()
                proj_token = str(projection_economy or "").strip().upper()
                if raw_token and proj_token.endswith(raw_token):
                    leap_economy = projection_economy
                for _, row in sub_scenario.iterrows():
                    long_rows.append(
                        {
                            "economy": leap_economy,
                            "scenario": scenario_display,
                            "sheet": sheet,
                            "measure": measure_label,
                            "fuel_label": fuel,
                            "source": "leap",
                            "year": int(row["year"]),
                            "value": float(row["leap_value"]),
                        }
                    )

            if use_esto_agg_only:
                # ESTO-aggregated reference only (no projection)
                base_value = _sum_base_targets(base_targets)
                if (display_inputs_as_positive or display_exports_as_positive or not preserve_sign) and not pd.isna(base_value):
                    base_value = abs(float(base_value))
                long_rows.append(
                    {
                        "economy": base_economy,
                        "scenario": scenario_display,
                        "sheet": sheet,
                        "measure": measure_label,
                        "fuel_label": fuel,
                        "source": "esto_aggregated",
                        "year": base_year,
                        "value": base_value,
                    }
                )
            else:
                # Base year
                base_value = _sum_base_targets(base_targets)
                if (display_inputs_as_positive or display_exports_as_positive or not preserve_sign) and not pd.isna(base_value):
                    base_value = abs(float(base_value))
                long_rows.append(
                    {
                        "economy": base_economy,
                        "scenario": scenario_display,
                        "sheet": sheet,
                        "measure": measure_label,
                        "fuel_label": fuel,
                        "source": "base",
                        "year": base_year,
                        "value": base_value,
                    }
                )

                # Projection series
                proj_parts: list[pd.Series] = []
                allow_sector_total_projection = category_type == "sector" and not explicit_entry
                projection_fuel_codes: list[str] = []
                projection_targets = list(explicit_projection_targets)
                if explicit_projection_targets:
                    projection_fuel_codes = [fuel_code for _, fuel_code in explicit_projection_targets if fuel_code]
                elif explicit_projection_fuel_codes:
                    projection_fuel_codes = list(explicit_projection_fuel_codes)
                elif projection_fuel_filters:
                    if ninth_fuel:
                        if any(ninth_fuel.strip().lower() == code.strip().lower() for code in projection_fuel_filters):
                            projection_fuel_codes = [ninth_fuel]
                        else:
                            projection_fuel_codes = []
                    else:
                        projection_fuel_codes = list(projection_fuel_filters)
                elif ninth_fuel:
                    projection_fuel_codes = [ninth_fuel]

                allow_projection_parent_fallback = not bool(explicit_projection_targets)
                if projection_targets:
                    projection_pairs = projection_targets
                else:
                    if projection_fuel_codes:
                        projection_pairs = [
                            (sector_code, proj_fuel_code)
                            for sector_code in comparison_sector_codes
                            for proj_fuel_code in projection_fuel_codes
                        ]
                    elif allow_sector_total_projection:
                        # Only sector-summary rows may intentionally pull an
                        # all-fuels projection total (blank fuel code).
                        projection_pairs = [
                            (sector_code, "")
                            for sector_code in comparison_sector_codes
                        ]
                    else:
                        # Fuel-level rows with no resolved projection fuel
                        # must remain missing rather than repeating a
                        # sector-wide total for each fuel label.
                        projection_pairs = []

                if projection_pairs or allow_sector_total_projection:
                    scenario_token = str(projection_scenario or "").strip().lower()
                    seen_projection_identities: set[tuple[str, str, str, str]] = set()
                    for sector_code, proj_fuel_code in projection_pairs or [("", "")]:
                        sector_token = str(sector_code or "").strip().lower()
                        fuel_token = str(proj_fuel_code or "").strip().lower()
                        cache_key = (sector_token, fuel_token, scenario_token, transformation_sign_role)
                        if cache_key not in projection_cache:
                            projection_cache[cache_key] = pull_projection_series(
                                ninth_projection_df,
                                sector_code=sector_token,
                                fuel_code=proj_fuel_code,
                                economy_code=projection_economy,
                                scenario=projection_scenario,
                                projection_years=projection_years,
                                value_sign_role=transformation_sign_role,
                            ).reindex(projection_years)
                        proj_part = projection_cache[cache_key]
                        effective_sector_token = sector_token
                        if allow_projection_parent_fallback and not proj_part.notna().any():
                            descendant_part, descendant_codes = pull_projection_series_from_descendants(
                                ninth_projection_df,
                                sector_code=sector_token,
                                fuel_code=proj_fuel_code,
                                economy_code=projection_economy,
                                scenario=projection_scenario,
                                projection_years=projection_years,
                                value_sign_role=transformation_sign_role,
                            )
                            if descendant_codes and descendant_part.notna().any():
                                proj_part = descendant_part
                                effective_sector_token = sector_token + "|__descendants__"
                                status_rows[status_idx]["comparator_scope"] = "child"
                                status_rows[status_idx]["projection_parent_fallback"] = False
                                status_rows[status_idx]["projection_parent_sector_code"] = ""
                                existing_note = str(status_rows[status_idx].get("mapping_note", "") or "")
                                status_rows[status_idx]["mapping_note"] = _merge_mapping_note(
                                    existing_note,
                                    "projection aggregated from descendant sectors",
                                )
                            else:
                                seen_sector_tokens = {sector_token}
                                parent_sector_token = _nearest_parent_sector_code(sector_token)
                                while parent_sector_token and parent_sector_token not in seen_sector_tokens:
                                    seen_sector_tokens.add(parent_sector_token)
                                    parent_cache_key = (
                                        parent_sector_token,
                                        fuel_token,
                                        scenario_token,
                                        transformation_sign_role,
                                    )
                                    if parent_cache_key not in projection_cache:
                                        projection_cache[parent_cache_key] = pull_projection_series(
                                            ninth_projection_df,
                                            sector_code=parent_sector_token,
                                            fuel_code=proj_fuel_code,
                                            economy_code=projection_economy,
                                            scenario=projection_scenario,
                                            projection_years=projection_years,
                                            value_sign_role=transformation_sign_role,
                                        ).reindex(projection_years)
                                    parent_part = projection_cache[parent_cache_key]
                                    if parent_part.notna().any():
                                        proj_part = parent_part
                                        effective_sector_token = parent_sector_token
                                        status_rows[status_idx]["projection_parent_fallback"] = True
                                        if not status_rows[status_idx].get("projection_parent_sector_code"):
                                            status_rows[status_idx]["projection_parent_sector_code"] = parent_sector_token
                                        break
                                    parent_sector_token = _nearest_parent_sector_code(parent_sector_token)
                        projection_identity = (
                            effective_sector_token,
                            fuel_token,
                            scenario_token,
                            transformation_sign_role,
                        )
                        if projection_identity in seen_projection_identities:
                            continue
                        seen_projection_identities.add(projection_identity)
                        proj_parts.append(proj_part.reindex(projection_years))
                if proj_parts:
                    proj_series = pd.concat(proj_parts, axis=1).sum(axis=1, min_count=1)
                else:
                    proj_series = pd.Series(dtype="float64", index=projection_years)
                if display_inputs_as_positive or display_exports_as_positive or not preserve_sign:
                    proj_series = proj_series.abs()
                for year, val in proj_series.items():
                    long_rows.append(
                        {
                            "economy": projection_economy,
                            "scenario": scenario_display,
                            "sheet": sheet,
                            "measure": measure_label,
                            "fuel_label": fuel,
                            "source": "projection",
                            "year": int(year),
                            "value": float(val) if not pd.isna(val) else float("nan"),
                        }
                    )

    status_df = pd.DataFrame(status_rows)
    if not status_df.empty:
        flow_depth_lookup: dict[str, int] = {}
        flow_name_lookup: dict[str, str] = {}
        flow_to_parent_sector_code: dict[str, str] = {}
        sector_name_lookup: dict[str, str] = {}
        code_df = _safe_read_codebook_sheet(DEFAULT_CODEBOOK, "code_to_name")
        if not code_df.empty:
            code_df["esto_label_norm"] = code_df.get("esto_label", "").map(_clean_token).str.lower()
            code_df["esto_column_norm"] = code_df.get("esto_column", "").map(_clean_token).str.lower()
            code_df["ninth_label_norm"] = code_df.get("9th_label", "").map(_clean_token).str.lower()
            code_df["name_clean"] = code_df.get("name", "").map(_clean_token)
            for _, row in code_df.iterrows():
                ninth_label = str(row.get("ninth_label_norm") or "").strip()
                name_clean = str(row.get("name_clean") or "").strip()
                if ninth_label and name_clean and ninth_label not in sector_name_lookup:
                    sector_name_lookup[ninth_label] = name_clean
                if row["esto_column_norm"] != "flows":
                    continue
                flow = str(row["esto_label_norm"] or "").strip()
                if flow and name_clean and flow not in flow_name_lookup:
                    flow_name_lookup[flow] = name_clean
                if flow and ninth_label and flow not in flow_to_parent_sector_code:
                    flow_to_parent_sector_code[flow] = ninth_label
            if flow and ninth_label:
                depth = _code_depth(ninth_label)
                if flow not in flow_depth_lookup or depth < flow_depth_lookup[flow]:
                    flow_depth_lookup[flow] = depth
        status_df = status_df.copy()
        status_df["esto_flow_norm"] = status_df["esto_flow"].fillna("").astype(str).str.strip().str.lower()
        status_df["flow_source"] = status_df["flow_source"].fillna("").astype(str).str.strip().str.lower()
        status_df["sector_match_method"] = status_df["sector_match_method"].map(shared_normalize_match_method)
        fallback_parent_flag = (
            status_df["projection_parent_fallback"].fillna(False).astype(bool)
            if "projection_parent_fallback" in status_df.columns
            else pd.Series(False, index=status_df.index)
        )
        status_df["projection_parent_sector_code_norm"] = (
            status_df.get("projection_parent_sector_code", "")
            .fillna("")
            .astype(str)
            .str.strip()
            .str.lower()
        )
        status_df["projection_parent_name"] = status_df["projection_parent_sector_code_norm"].map(sector_name_lookup).fillna("")
        status_df["esto_flow_depth"] = status_df["esto_flow_norm"].map(flow_depth_lookup)
        status_df["sector_codes_list"] = status_df["sector_code_9th"].map(_split_sector_codes)
        status_df["min_sector_depth"] = status_df["sector_codes_list"].map(
            lambda xs: min((_code_depth(x) for x in xs), default=9999)
        )
        status_df["derived_parent_flow_levels"] = status_df.apply(
            lambda row: _derive_parent_flow_levels(
                row.get("sector_codes_list", []),
                row.get("esto_flow_norm", ""),
                flow_to_parent_sector_code,
            ),
            axis=1,
        )
        derived_parent_flag = pd.to_numeric(
            status_df["derived_parent_flow_levels"], errors="coerce"
        ).fillna(0).gt(0)
        status_df["uses_parent_flow"] = (
            (
                pd.to_numeric(status_df["esto_flow_depth"], errors="coerce").fillna(9999)
                < pd.to_numeric(status_df["min_sector_depth"], errors="coerce").fillna(9999)
            )
            | derived_parent_flag
            | fallback_parent_flag
        )
        status_df["allow_parent_estimate"] = (
            (status_df["flow_source"] == "sector_fallback")
            | (status_df["flow_source"] == "canonical_aggregated")
            | (status_df["flow_source"] == "category_sector")
            | status_df["sector_match_method"].eq("aggregated_canonical_targets")
            | status_df["sector_match_method"].isin({"sheet_name_lookup", "sector_name_lookup"})
            | derived_parent_flag
            | status_df["sector_match_method"].map(shared_is_reverse_independent_match)
            | fallback_parent_flag
        )
        status_df["effective_parent_name"] = status_df["esto_flow_norm"].map(flow_name_lookup).fillna("")
        status_df.loc[
            fallback_parent_flag & status_df["projection_parent_name"].ne(""),
            "effective_parent_name",
        ] = status_df.loc[
            fallback_parent_flag & status_df["projection_parent_name"].ne(""),
            "projection_parent_name",
        ]
        if sibling_mode == "aggregate_to_parent":
            status_df["comparator_scope"] = "child"
            status_df.loc[
                status_df["uses_parent_flow"] & status_df["allow_parent_estimate"],
                "comparator_scope",
            ] = "parent"
        strict_parent_flow = status_df[
            status_df["uses_parent_flow"]
            & status_df["esto_flow_norm"].ne("")
            & ~status_df["allow_parent_estimate"]
        ].copy()
        if not strict_parent_flow.empty:
            examples = (
                strict_parent_flow[
                    ["sheet", "fuel_label", "sector_code_9th", "esto_flow", "flow_source", "sector_match_method"]
                ]
                .drop_duplicates()
                .head(10)
                .to_dict("records")
            )
            raise RuntimeError(
                "Parent-flow canonical mappings detected. These rows map to a shallower ESTO flow and must be fixed "
                f"in the mapping files instead of being estimated. Total rows: {len(strict_parent_flow)}. "
                f"Examples: {examples}"
            )
    comparison_long = pd.DataFrame(long_rows)
    if not comparison_long.empty and not status_df.empty:
        def _esto_product_code(product_value: object) -> str:
            token = _clean_token(product_value)
            return token.split(" ", 1)[0].strip() if token else ""

        def _is_parent_product_code(product_value: object) -> bool:
            code = _esto_product_code(product_value)
            return bool(code) and "." not in code

        exact_base_status = status_df.copy()
        exact_base_status["sheet"] = exact_base_status["sheet"].astype(str)
        exact_base_status["fuel_label"] = exact_base_status["fuel_label"].astype(str)
        exact_base_status["esto_flow"] = exact_base_status["esto_flow"].fillna("").astype(str).str.strip()
        exact_base_status["esto_product"] = exact_base_status["esto_product"].fillna("").astype(str).str.strip()
        exact_base_status = exact_base_status[
            exact_base_status["esto_flow"].ne("")
            & exact_base_status["esto_product"].ne("")
        ].copy()
        exact_base_status = exact_base_status[
            ~exact_base_status["esto_flow"].str.contains(r"\(aggregated\)", case=False, na=False)
            & ~exact_base_status["esto_product"].str.contains(r"\(aggregated\)", case=False, na=False)
        ].copy()
        exact_base_status["product_code"] = exact_base_status["esto_product"].map(_esto_product_code)
        exact_base_status["is_parent_product"] = exact_base_status["esto_product"].map(_is_parent_product_code)

        if not exact_base_status.empty:
            comp_base_parent = comparison_long.copy()
            comp_base_parent["sheet"] = comp_base_parent["sheet"].astype(str)
            comp_base_parent["fuel_label"] = comp_base_parent["fuel_label"].astype(str)
            comp_base_parent["scenario"] = comp_base_parent["scenario"].astype(str)
            comp_base_parent["value"] = pd.to_numeric(comp_base_parent["value"], errors="coerce")
            comp_base_parent = comp_base_parent[
                comp_base_parent["source"].isin(["base", "base_estimated"]) & comp_base_parent["year"].eq(base_year)
            ].copy()
            if not comp_base_parent.empty:
                comp_base_parent = comp_base_parent.merge(
                    exact_base_status[
                        ["sheet", "fuel_label", "esto_flow", "esto_product", "product_code", "is_parent_product"]
                    ],
                    on=["sheet", "fuel_label"],
                    how="inner",
                )
                parent_candidates = comp_base_parent[comp_base_parent["is_parent_product"]].copy()
                if not parent_candidates.empty:
                    child_status = exact_base_status[
                        ~exact_base_status["is_parent_product"]
                    ][["sheet", "fuel_label", "esto_flow", "esto_product", "product_code"]].drop_duplicates()
                    adjusted_parent_keys: list[tuple[str, str]] = []
                    for parent in parent_candidates[
                        ["sheet", "fuel_label", "esto_flow", "esto_product", "product_code"]
                    ].drop_duplicates().itertuples(index=False):
                        child_matches = child_status[
                            (child_status["sheet"] == parent.sheet)
                            & (child_status["esto_flow"] == parent.esto_flow)
                            & (child_status["product_code"].str.startswith(parent.product_code + "."))
                        ].copy()
                        if child_matches.empty:
                            continue
                        child_labels = child_matches["fuel_label"].drop_duplicates().tolist()
                        if not child_labels:
                            continue
                        parent_mask = (
                            (comparison_long["sheet"].astype(str) == parent.sheet)
                            & (comparison_long["fuel_label"].astype(str) == parent.fuel_label)
                            & (comparison_long["source"].isin(["base", "base_estimated"]))
                            & (comparison_long["year"].eq(base_year))
                        )
                        if not parent_mask.any():
                            continue
                        for scenario_value in (
                            comparison_long.loc[parent_mask, "scenario"].astype(str).drop_duplicates().tolist()
                        ):
                            scenario_parent_mask = parent_mask & (
                                comparison_long["scenario"].astype(str) == scenario_value
                            )
                            parent_value_series = pd.to_numeric(
                                comparison_long.loc[scenario_parent_mask, "value"],
                                errors="coerce",
                            )
                            if parent_value_series.dropna().empty:
                                continue
                            parent_value = float(parent_value_series.dropna().iloc[0])
                            child_mask = (
                                (comparison_long["sheet"].astype(str) == parent.sheet)
                                & (comparison_long["fuel_label"].astype(str).isin(child_labels))
                                & (comparison_long["source"].isin(["base", "base_estimated"]))
                                & (comparison_long["year"].eq(base_year))
                                & (comparison_long["scenario"].astype(str) == scenario_value)
                            )
                            child_value = float(
                                pd.to_numeric(comparison_long.loc[child_mask, "value"], errors="coerce").sum(min_count=1)
                            ) if child_mask.any() else 0.0
                            residual = parent_value - child_value
                            if pd.isna(residual):
                                continue
                            if abs(residual) < 1e-9:
                                residual = 0.0
                            elif residual < 0:
                                residual = 0.0
                            comparison_long.loc[scenario_parent_mask, "value"] = residual
                        adjusted_parent_keys.append((parent.sheet, parent.fuel_label))
                    if adjusted_parent_keys:
                        adjusted_parent_keys = sorted(set(adjusted_parent_keys))
                        adjusted_df = pd.DataFrame(adjusted_parent_keys, columns=["sheet", "fuel_label"])
                        status_df = status_df.merge(
                            adjusted_df.assign(_base_parent_residual=True),
                            on=["sheet", "fuel_label"],
                            how="left",
                        )
                        adjusted_mask = status_df["_base_parent_residual"].fillna(False).astype(bool)
                        residual_note = "base parent product reduced by mapped child products"
                        status_df.loc[adjusted_mask, "mapping_note"] = status_df.loc[
                            adjusted_mask, "mapping_note"
                        ].fillna("").astype(str).map(
                            lambda note: residual_note if not note else f"{note}; {residual_note}"
                        )
                        status_df = status_df.drop(columns=["_base_parent_residual"], errors="ignore")

        alloc_status = status_df.copy()
        alloc_status["sheet"] = alloc_status["sheet"].astype(str)
        alloc_status["fuel_label"] = alloc_status["fuel_label"].astype(str)
        alloc_status["ninth_fuel_code"] = alloc_status["ninth_fuel_code"].fillna("").astype(str).str.strip()
        # Allocate projection values for any duplicated mapped fuel bucket
        # within a sheet (not just *_x_* aggregate codes).
        alloc_status = alloc_status[alloc_status["ninth_fuel_code"].ne("")].copy()

        duplicate_aggregate_groups = (
            alloc_status.groupby(["sheet", "ninth_fuel_code"], as_index=False)["fuel_label"]
            .nunique()
            .rename(columns={"fuel_label": "label_count"})
        )
        duplicate_aggregate_groups = duplicate_aggregate_groups[duplicate_aggregate_groups["label_count"] > 1].copy()

        if not duplicate_aggregate_groups.empty:
            alloc_status = alloc_status.merge(
                duplicate_aggregate_groups[["sheet", "ninth_fuel_code"]],
                on=["sheet", "ninth_fuel_code"],
                how="inner",
            )

            comp_alloc = comparison_long.copy()
            comp_alloc["sheet"] = comp_alloc["sheet"].astype(str)
            comp_alloc["fuel_label"] = comp_alloc["fuel_label"].astype(str)
            comp_alloc["scenario"] = comp_alloc["scenario"].astype(str)
            comp_alloc["value"] = pd.to_numeric(comp_alloc["value"], errors="coerce")
            comp_alloc = comp_alloc.merge(
                alloc_status[["sheet", "fuel_label", "ninth_fuel_code"]],
                on=["sheet", "fuel_label"],
                how="inner",
            )

            base_rows = comp_alloc[
                comp_alloc["source"].isin(["base", "base_estimated"]) & comp_alloc["year"].eq(base_year)
            ].copy()
            if not base_rows.empty:
                base_weights = (
                    base_rows.groupby(["sheet", "scenario", "fuel_label", "ninth_fuel_code"], as_index=False)["value"]
                    .agg(weight=lambda s: pd.to_numeric(s, errors="coerce").abs().sum(min_count=1))
                )
            else:
                base_weights = pd.DataFrame(columns=["sheet", "scenario", "fuel_label", "ninth_fuel_code", "weight"])

            label_members = (
                alloc_status.groupby(["sheet", "ninth_fuel_code"], as_index=False)["fuel_label"]
                .nunique()
                .rename(columns={"fuel_label": "label_count"})
            )

            projection_rows = comp_alloc[comp_alloc["source"] == "projection"].copy()
            if not projection_rows.empty:
                projection_parent = (
                    projection_rows.groupby(["sheet", "scenario", "year", "ninth_fuel_code"], as_index=False)["value"]
                    .agg(
                        first_non_null=lambda s: s.dropna().iloc[0] if len(s.dropna()) else float("nan"),
                        unique_non_null=lambda s: s.dropna().nunique(),
                    )
                )
                projection_parent = projection_parent[projection_parent["unique_non_null"] <= 1].copy()

                weight_table = alloc_status[["sheet", "fuel_label", "ninth_fuel_code"]].drop_duplicates()
                scenarios = (
                    projection_rows[["sheet", "scenario", "ninth_fuel_code"]]
                    .drop_duplicates()
                )
                weight_table = scenarios.merge(weight_table, on=["sheet", "ninth_fuel_code"], how="left")
                weight_table = weight_table.merge(
                    base_weights,
                    on=["sheet", "scenario", "fuel_label", "ninth_fuel_code"],
                    how="left",
                )
                weight_table = weight_table.merge(
                    label_members,
                    on=["sheet", "ninth_fuel_code"],
                    how="left",
                )
                weight_table["weight"] = pd.to_numeric(weight_table["weight"], errors="coerce")
                weight_table["weight_total"] = weight_table.groupby(
                    ["sheet", "scenario", "ninth_fuel_code"]
                )["weight"].transform(lambda s: s.sum(min_count=1))
                weight_table["alloc_share"] = weight_table["weight"] / weight_table["weight_total"]
                invalid_share = ~weight_table["alloc_share"].replace([float("inf"), float("-inf")], pd.NA).notna()
                weight_table.loc[invalid_share, "alloc_share"] = (
                    1.0 / pd.to_numeric(weight_table["label_count"], errors="coerce")
                )

                alloc_projection_rows = projection_parent.merge(
                    weight_table[
                        ["sheet", "scenario", "fuel_label", "ninth_fuel_code", "alloc_share"]
                    ],
                    on=["sheet", "scenario", "ninth_fuel_code"],
                    how="inner",
                )
                if not alloc_projection_rows.empty:
                    alloc_projection_rows["value"] = (
                        alloc_projection_rows["first_non_null"] * alloc_projection_rows["alloc_share"]
                    )
                    # Shared 9th fuel buckets split across multiple display labels
                    # are estimated using base-year comparator shares.
                    alloc_projection_rows["source"] = "projection_estimated"
                    keep_cols = ["economy", "scenario", "sheet", "fuel_label", "source", "year", "value"]
                    econ_ref = (
                        projection_rows.groupby(
                            ["sheet", "scenario", "fuel_label", "ninth_fuel_code"], as_index=False
                        )["economy"].first()
                    )
                    alloc_projection_rows = alloc_projection_rows.merge(
                        econ_ref,
                        on=["sheet", "scenario", "fuel_label", "ninth_fuel_code"],
                        how="left",
                    )

                    replace_keys = alloc_projection_rows[
                        ["sheet", "scenario", "fuel_label", "year"]
                    ].drop_duplicates()
                    replace_keys["source"] = "projection"
                    comparison_keyed = comparison_long.merge(
                        replace_keys.assign(_replace=True),
                        on=["sheet", "scenario", "fuel_label", "source", "year"],
                        how="left",
                    )
                    comparison_long = comparison_keyed[comparison_keyed["_replace"] != True].drop(
                        columns=["_replace"],
                        errors="ignore",
                    )
                    comparison_long = pd.concat(
                        [comparison_long, alloc_projection_rows[keep_cols]],
                        ignore_index=True,
                        sort=False,
                    )

        duplicate_base_target_status = status_df.copy()
        duplicate_base_target_status["sheet"] = duplicate_base_target_status["sheet"].astype(str)
        duplicate_base_target_status["fuel_label"] = duplicate_base_target_status["fuel_label"].astype(str)
        duplicate_base_target_status["esto_flow"] = duplicate_base_target_status["esto_flow"].fillna("").astype(str).str.strip()
        duplicate_base_target_status["esto_product"] = duplicate_base_target_status["esto_product"].fillna("").astype(str).str.strip()
        duplicate_base_target_status = duplicate_base_target_status[
            duplicate_base_target_status["esto_flow"].ne("")
            & duplicate_base_target_status["esto_product"].ne("")
        ].copy()
        duplicate_base_target_status = duplicate_base_target_status[
            ~duplicate_base_target_status["esto_flow"].str.contains(r"\(aggregated\)", case=False, na=False)
            & ~duplicate_base_target_status["esto_product"].str.contains(r"\(aggregated\)", case=False, na=False)
        ].copy()

        duplicate_base_target_groups = (
            duplicate_base_target_status.groupby(["sheet", "esto_flow", "esto_product"], as_index=False)["fuel_label"]
            .nunique()
            .rename(columns={"fuel_label": "label_count"})
        )
        duplicate_base_target_groups = duplicate_base_target_groups[
            duplicate_base_target_groups["label_count"] > 1
        ].copy()

        if not duplicate_base_target_groups.empty:
            duplicate_base_target_status = duplicate_base_target_status.merge(
                duplicate_base_target_groups[["sheet", "esto_flow", "esto_product", "label_count"]],
                on=["sheet", "esto_flow", "esto_product"],
                how="inner",
            )

            comp_base_alloc = comparison_long.copy()
            comp_base_alloc["sheet"] = comp_base_alloc["sheet"].astype(str)
            comp_base_alloc["fuel_label"] = comp_base_alloc["fuel_label"].astype(str)
            comp_base_alloc["scenario"] = comp_base_alloc["scenario"].astype(str)
            comp_base_alloc["value"] = pd.to_numeric(comp_base_alloc["value"], errors="coerce")
            comp_base_alloc = comp_base_alloc.merge(
                duplicate_base_target_status[["sheet", "fuel_label", "esto_flow", "esto_product", "label_count"]],
                on=["sheet", "fuel_label"],
                how="inner",
            )

            base_rows = comp_base_alloc[
                comp_base_alloc["source"].eq("base") & comp_base_alloc["year"].eq(base_year)
            ].copy()
            if not base_rows.empty:
                base_parent = (
                    base_rows.groupby(["sheet", "scenario", "year", "esto_flow", "esto_product"], as_index=False)["value"]
                    .agg(
                        first_non_null=lambda s: s.dropna().iloc[0] if len(s.dropna()) else float("nan"),
                        unique_non_null=lambda s: s.dropna().nunique(),
                    )
                )
                base_parent = base_parent[base_parent["unique_non_null"] <= 1].copy()
            else:
                base_parent = pd.DataFrame(
                    columns=["sheet", "scenario", "year", "esto_flow", "esto_product", "first_non_null", "unique_non_null"]
                )

            if not base_parent.empty:
                member_rows = duplicate_base_target_status[
                    ["sheet", "fuel_label", "esto_flow", "esto_product", "label_count"]
                ].drop_duplicates()
                scenarios = base_parent[["sheet", "scenario", "esto_flow", "esto_product"]].drop_duplicates()
                weight_table = scenarios.merge(
                    member_rows,
                    on=["sheet", "esto_flow", "esto_product"],
                    how="left",
                )

                projection_weight_rows = comp_base_alloc[
                    comp_base_alloc["source"].isin(["projection", "projection_estimated"])
                ].copy()
                if not projection_weight_rows.empty:
                    projection_weight_rows = projection_weight_rows.sort_values(
                        ["sheet", "scenario", "fuel_label", "esto_flow", "esto_product", "year"]
                    )
                    projection_weights = (
                        projection_weight_rows.groupby(
                            ["sheet", "scenario", "fuel_label", "esto_flow", "esto_product"], as_index=False
                        )["value"]
                        .agg(weight=lambda s: pd.to_numeric(s, errors="coerce").abs().dropna().iloc[0] if len(pd.to_numeric(s, errors="coerce").abs().dropna()) else float("nan"))
                    )
                    projection_weights_any_scenario = (
                        projection_weight_rows.groupby(
                            ["sheet", "fuel_label", "esto_flow", "esto_product"], as_index=False
                        )["value"]
                        .agg(weight_any_scenario=lambda s: pd.to_numeric(s, errors="coerce").abs().dropna().iloc[0] if len(pd.to_numeric(s, errors="coerce").abs().dropna()) else float("nan"))
                    )
                else:
                    projection_weights = pd.DataFrame(
                        columns=["sheet", "scenario", "fuel_label", "esto_flow", "esto_product", "weight"]
                    )
                    projection_weights_any_scenario = pd.DataFrame(
                        columns=["sheet", "fuel_label", "esto_flow", "esto_product", "weight_any_scenario"]
                    )

                leap_weight_rows = comp_base_alloc[
                    comp_base_alloc["source"].eq("leap") & comp_base_alloc["year"].isin(list(projection_years))
                ].copy()
                if not leap_weight_rows.empty:
                    leap_weight_rows = leap_weight_rows.sort_values(
                        ["sheet", "scenario", "fuel_label", "esto_flow", "esto_product", "year"]
                    )
                    leap_weights = (
                        leap_weight_rows.groupby(
                            ["sheet", "scenario", "fuel_label", "esto_flow", "esto_product"], as_index=False
                        )["value"]
                        .agg(leap_weight=lambda s: pd.to_numeric(s, errors="coerce").abs().dropna().iloc[0] if len(pd.to_numeric(s, errors="coerce").abs().dropna()) else float("nan"))
                    )
                    leap_weights_any_scenario = (
                        leap_weight_rows.groupby(
                            ["sheet", "fuel_label", "esto_flow", "esto_product"], as_index=False
                        )["value"]
                        .agg(leap_weight_any_scenario=lambda s: pd.to_numeric(s, errors="coerce").abs().dropna().iloc[0] if len(pd.to_numeric(s, errors="coerce").abs().dropna()) else float("nan"))
                    )
                else:
                    leap_weights = pd.DataFrame(
                        columns=["sheet", "scenario", "fuel_label", "esto_flow", "esto_product", "leap_weight"]
                    )
                    leap_weights_any_scenario = pd.DataFrame(
                        columns=["sheet", "fuel_label", "esto_flow", "esto_product", "leap_weight_any_scenario"]
                    )

                weight_table = weight_table.merge(
                    projection_weights,
                    on=["sheet", "scenario", "fuel_label", "esto_flow", "esto_product"],
                    how="left",
                )
                weight_table = weight_table.merge(
                    projection_weights_any_scenario,
                    on=["sheet", "fuel_label", "esto_flow", "esto_product"],
                    how="left",
                )
                weight_table = weight_table.merge(
                    leap_weights,
                    on=["sheet", "scenario", "fuel_label", "esto_flow", "esto_product"],
                    how="left",
                )
                weight_table = weight_table.merge(
                    leap_weights_any_scenario,
                    on=["sheet", "fuel_label", "esto_flow", "esto_product"],
                    how="left",
                )
                weight_table["weight"] = pd.to_numeric(weight_table.get("weight"), errors="coerce")
                weight_table["weight_any_scenario"] = pd.to_numeric(
                    weight_table.get("weight_any_scenario"), errors="coerce"
                )
                weight_table["leap_weight"] = pd.to_numeric(weight_table.get("leap_weight"), errors="coerce")
                weight_table["leap_weight_any_scenario"] = pd.to_numeric(
                    weight_table.get("leap_weight_any_scenario"), errors="coerce"
                )
                weight_table["alloc_weight"] = weight_table["weight"]
                weight_table["alloc_weight"] = weight_table["alloc_weight"].fillna(weight_table["weight_any_scenario"])
                weight_table["alloc_weight"] = weight_table["alloc_weight"].fillna(weight_table["leap_weight"])
                weight_table["alloc_weight"] = weight_table["alloc_weight"].fillna(
                    weight_table["leap_weight_any_scenario"]
                )
                weight_table["weight_total"] = weight_table.groupby(
                    ["sheet", "scenario", "esto_flow", "esto_product"]
                )["alloc_weight"].transform(lambda s: s.sum(min_count=1))
                weight_table["alloc_share"] = weight_table["alloc_weight"] / weight_table["weight_total"]
                invalid_share = ~weight_table["alloc_share"].replace([float("inf"), float("-inf")], pd.NA).notna()
                weight_table.loc[invalid_share, "alloc_share"] = (
                    1.0 / pd.to_numeric(weight_table["label_count"], errors="coerce")
                )

                alloc_base_rows = base_parent.merge(
                    weight_table[
                        ["sheet", "scenario", "fuel_label", "esto_flow", "esto_product", "alloc_share"]
                    ],
                    on=["sheet", "scenario", "esto_flow", "esto_product"],
                    how="inner",
                )
                if not alloc_base_rows.empty:
                    alloc_base_rows["value"] = alloc_base_rows["first_non_null"] * alloc_base_rows["alloc_share"]
                    alloc_base_rows["source"] = "base_estimated"
                    keep_cols = ["economy", "scenario", "sheet", "fuel_label", "source", "year", "value"]
                    econ_ref = (
                        base_rows.groupby(["sheet", "scenario", "fuel_label", "esto_flow", "esto_product"], as_index=False)[
                            "economy"
                        ].first()
                    )
                    alloc_base_rows = alloc_base_rows.merge(
                        econ_ref,
                        on=["sheet", "scenario", "fuel_label", "esto_flow", "esto_product"],
                        how="left",
                    )
                    replace_keys = alloc_base_rows[["sheet", "scenario", "fuel_label", "year"]].drop_duplicates()
                    replace_keys["source"] = "base"
                    comparison_keyed = comparison_long.merge(
                        replace_keys.assign(_replace=True),
                        on=["sheet", "scenario", "fuel_label", "source", "year"],
                        how="left",
                    )
                    comparison_long = comparison_keyed[comparison_keyed["_replace"] != True].drop(
                        columns=["_replace"],
                        errors="ignore",
                    )
                    comparison_long = pd.concat(
                        [comparison_long, alloc_base_rows[keep_cols]],
                        ignore_index=True,
                        sort=False,
                    )
    mapping_status = status_df.copy() if not status_df.empty else pd.DataFrame(status_rows)

    if not comparison_long.empty and sibling_mode in {"allocate_by_leap_share", "aggregate_to_parent"}:
        if not status_df.empty:
            flow_name_lookup: dict[str, str] = {}
            flow_depth_lookup: dict[str, int] = {}
            flow_to_parent_sector_code: dict[str, str] = {}
            sector_name_lookup: dict[str, str] = {}
            code_df = _safe_read_codebook_sheet(DEFAULT_CODEBOOK, "code_to_name")
            if not code_df.empty:
                code_df["esto_label_norm"] = code_df.get("esto_label", "").map(_clean_token).str.lower()
                code_df["esto_column_norm"] = code_df.get("esto_column", "").map(_clean_token).str.lower()
                code_df["name_clean"] = code_df.get("name", "").map(_clean_token)
                code_df["ninth_label_norm"] = code_df.get("9th_label", "").map(_clean_token).str.lower()
                for _, row in code_df.iterrows():
                    ninth_label = str(row.get("ninth_label_norm") or "").strip()
                    name = str(row["name_clean"] or "").strip()
                    if ninth_label and name and ninth_label not in sector_name_lookup:
                        sector_name_lookup[ninth_label] = name
                    if row["esto_column_norm"] != "flows":
                        continue
                    flow = str(row["esto_label_norm"] or "").strip()
                    if flow and name and flow not in flow_name_lookup:
                        flow_name_lookup[flow] = name
                    if flow and ninth_label and flow not in flow_to_parent_sector_code:
                        flow_to_parent_sector_code[flow] = ninth_label
                    if flow and ninth_label:
                        depth = _code_depth(ninth_label)
                        if flow not in flow_depth_lookup or depth < flow_depth_lookup[flow]:
                            flow_depth_lookup[flow] = depth

            status_df = status_df.copy()
            status_df["sheet"] = status_df["sheet"].astype(str)
            status_df["fuel_label"] = status_df["fuel_label"].astype(str)
            status_df["esto_flow_norm"] = status_df["esto_flow"].fillna("").astype(str).str.strip().str.lower()
            status_df["flow_source"] = status_df["flow_source"].fillna("").astype(str).str.strip().str.lower()
            status_df["sector_match_method"] = status_df["sector_match_method"].map(shared_normalize_match_method)
            status_df["mapping_note"] = status_df.get("mapping_note", "").fillna("").astype(str)
            fallback_parent_flag = (
                status_df["projection_parent_fallback"].fillna(False).astype(bool)
                if "projection_parent_fallback" in status_df.columns
                else pd.Series(False, index=status_df.index)
            )
            status_df["projection_parent_sector_code_norm"] = (
                status_df.get("projection_parent_sector_code", "")
                .fillna("")
                .astype(str)
                .str.strip()
                .str.lower()
            )
            fallback_parent_flag = fallback_parent_flag | status_df["projection_parent_sector_code_norm"].ne("")
            status_df["projection_parent_name"] = status_df["projection_parent_sector_code_norm"].map(sector_name_lookup).fillna("")
            status_df["effective_parent_name"] = status_df["esto_flow_norm"].map(flow_name_lookup).fillna("")
            status_df.loc[
                fallback_parent_flag & status_df["projection_parent_name"].ne(""),
                "effective_parent_name",
            ] = status_df.loc[
                fallback_parent_flag & status_df["projection_parent_name"].ne(""),
                "projection_parent_name",
            ]
            status_df["esto_flow_depth"] = status_df["esto_flow_norm"].map(flow_depth_lookup)
            status_df["sector_codes_list"] = status_df["sector_code_9th"].map(_split_sector_codes)
            status_df["effective_comparator_sector_code"] = status_df["projection_parent_sector_code_norm"]
            missing_effective_code = status_df["effective_comparator_sector_code"].eq("")
            status_df.loc[missing_effective_code, "effective_comparator_sector_code"] = status_df.loc[
                missing_effective_code, "sector_codes_list"
            ].map(_effective_comparator_key_from_sector_codes)
            status_df["min_sector_depth"] = status_df["sector_codes_list"].map(
                lambda xs: min((_code_depth(x) for x in xs), default=9999)
            )
            status_df["derived_parent_flow_levels"] = status_df.apply(
                lambda row: _derive_parent_flow_levels(
                    row.get("sector_codes_list", []),
                    row.get("esto_flow_norm", ""),
                    flow_to_parent_sector_code,
                ),
                axis=1,
            )
            derived_parent_flag = pd.to_numeric(
                status_df["derived_parent_flow_levels"], errors="coerce"
            ).fillna(0).gt(0)
            status_df["uses_parent_flow"] = (
                (
                    pd.to_numeric(status_df["esto_flow_depth"], errors="coerce").fillna(9999)
                    < pd.to_numeric(status_df["min_sector_depth"], errors="coerce").fillna(9999)
                )
                | derived_parent_flag
                | fallback_parent_flag
            )
            status_df["allow_parent_estimate"] = (
                (status_df["flow_source"] == "sector_fallback")
                | (status_df["flow_source"] == "canonical_aggregated")
                | (status_df["flow_source"] == "category_sector")
                | status_df["sector_match_method"].eq("aggregated_canonical_targets")
                | status_df["sector_match_method"].isin({"sheet_name_lookup", "sector_name_lookup"})
                | derived_parent_flag
                | status_df["sector_match_method"].map(shared_is_reverse_independent_match)
                | fallback_parent_flag
            )
            status_df = (
                status_df.sort_values(["sheet", "fuel_label"])
                .drop_duplicates(subset=["sheet", "fuel_label"], keep="first")
                [[
                    "sheet",
                    "fuel_label",
                    "esto_flow_norm",
                    "flow_source",
                    "effective_parent_name",
                    "effective_comparator_sector_code",
                    "sector_codes_list",
                    "min_sector_depth",
                    "uses_parent_flow",
                    "allow_parent_estimate",
                ]]
            )

            comp = comparison_long.copy()
            comp["sheet"] = comp["sheet"].astype(str)
            comp["fuel_label"] = comp["fuel_label"].astype(str)
            comp["value"] = pd.to_numeric(comp["value"], errors="coerce")
            comp = comp.merge(status_df, on=["sheet", "fuel_label"], how="left")
            comp["esto_flow_norm"] = comp["esto_flow_norm"].fillna("")
            comp["flow_source"] = comp["flow_source"].fillna("")
            comp["effective_parent_name"] = comp["effective_parent_name"].fillna("")
            if "effective_comparator_sector_code" in comp.columns:
                comp["effective_comparator_sector_code"] = (
                    comp["effective_comparator_sector_code"].fillna("").astype(str).str.strip().str.lower()
                )
            comp["min_sector_depth"] = pd.to_numeric(comp["min_sector_depth"], errors="coerce").fillna(9999)
            comp["uses_parent_flow"] = comp["uses_parent_flow"].fillna(False).astype(bool)
            comp["allow_parent_estimate"] = comp["allow_parent_estimate"].fillna(False).astype(bool)

            fallback_mask = comp["uses_parent_flow"] & comp["allow_parent_estimate"] & (comp["effective_parent_name"] != "")
            group_cols = ["scenario", "fuel_label", "year", "effective_parent_name"]

            if fallback_mask.any():
                fallback_rows = comp[fallback_mask].copy()

                if sibling_mode == "allocate_by_leap_share":
                    # Allocate parent-level comparator series to sibling detail sheets using LEAP shares.
                    share_rows = pd.DataFrame()
                    leap_rows = fallback_rows[fallback_rows["source"] == "leap"].copy()
                    if not leap_rows.empty:
                        min_depth = (
                            leap_rows.groupby(group_cols, as_index=False)["min_sector_depth"]
                            .min()
                            .rename(columns={"min_sector_depth": "group_min_depth"})
                        )
                        parent_rows = leap_rows.merge(min_depth, on=group_cols, how="left")
                        parent_rows = parent_rows[parent_rows["min_sector_depth"] == parent_rows["group_min_depth"]].copy()
                        totals = (
                            parent_rows.groupby(group_cols, as_index=False)["value"]
                            .sum(min_count=1)
                            .rename(columns={"value": "parent_total"})
                        )
                        shares = leap_rows.merge(totals, on=group_cols, how="left")
                        shares["detail_share"] = shares["value"] / shares["parent_total"]
                        shares.loc[
                            ~shares["detail_share"].replace([float("inf"), float("-inf")], pd.NA).notna(),
                            "detail_share",
                        ] = pd.NA
                        share_rows = shares[
                            ["sheet"] + group_cols + ["esto_flow_norm", "flow_source", "min_sector_depth", "economy", "detail_share"]
                        ].copy()
                    if not share_rows.empty:
                        share_rows = share_rows.dropna(subset=["detail_share"]).drop_duplicates(
                            subset=["sheet"] + group_cols,
                            keep="first",
                        )
                        keep_cols = [
                            "economy",
                            "scenario",
                            "sheet",
                            "fuel_label",
                            "source",
                            "year",
                            "value",
                            "esto_flow_norm",
                            "flow_source",
                            "effective_parent_name",
                            "min_sector_depth",
                        ]
                        base_rows = fallback_rows[fallback_rows["source"] == "base"].copy()
                        if not base_rows.empty:
                            base_parent = (
                                base_rows.groupby(group_cols, as_index=False)["value"]
                                .agg(
                                    first_non_null=lambda s: s.dropna().iloc[0] if len(s.dropna()) else float("nan"),
                                    unique_non_null=lambda s: s.dropna().nunique(),
                                )
                            )
                            base_parent = base_parent[base_parent["unique_non_null"] <= 1].copy()
                            est_rows = share_rows.merge(
                                base_parent[group_cols + ["first_non_null"]],
                                on=group_cols,
                                how="inner",
                            )
                            if not est_rows.empty:
                                est_rows["value"] = est_rows["detail_share"] * est_rows["first_non_null"]
                                existing_base = comp[
                                    comp["source"].eq("base")
                                ][["sheet", "scenario", "fuel_label", "year", "value"]].rename(
                                    columns={"value": "existing_value"}
                                )
                                est_rows = est_rows.merge(
                                    existing_base,
                                    on=["sheet", "scenario", "fuel_label", "year"],
                                    how="left",
                                )
                                existing_base_num = pd.to_numeric(est_rows["existing_value"], errors="coerce")
                                est_rows = est_rows[existing_base_num.isna() | existing_base_num.ne(0)].copy()
                                base_replace_keys = est_rows[
                                    ["sheet", "scenario", "fuel_label", "year"]
                                ].drop_duplicates()
                                base_replace_keys["source"] = "base"
                                est_rows["source"] = "base_estimated"
                                comp = comp.merge(
                                    base_replace_keys.assign(_replace=True),
                                    on=["sheet", "scenario", "fuel_label", "source", "year"],
                                    how="left",
                                )
                                comp = comp[comp["_replace"] != True].drop(
                                    columns=["_replace"],
                                    errors="ignore",
                                )
                                comp = pd.concat([comp, est_rows[keep_cols]], ignore_index=True, sort=False)

                        projection_rows = fallback_rows[fallback_rows["source"] == "projection"].copy()
                        if not projection_rows.empty:
                            share_key_cols = ["sheet", "scenario", "fuel_label", "effective_parent_name"]
                            projection_years = sorted(
                                pd.to_numeric(projection_rows["year"], errors="coerce").dropna().astype(int).unique().tolist()
                            )
                            share_rows_for_projection: list[pd.DataFrame] = []
                            for _, share_group in share_rows.groupby(share_key_cols, dropna=False):
                                share_group = share_group.sort_values("year").drop_duplicates(subset=["year"], keep="last")
                                detail_share = share_group.set_index("year")["detail_share"]
                                expanded_share = detail_share.reindex(
                                    sorted(set(detail_share.index.tolist()) | set(projection_years))
                                ).sort_index()
                                expanded_share = expanded_share.ffill().bfill().reindex(projection_years)
                                if expanded_share.isna().all():
                                    continue
                                template = share_group.iloc[0]
                                expanded_df = pd.DataFrame(
                                    {
                                        "year": projection_years,
                                        "detail_share": expanded_share.values,
                                    }
                                )
                                for col in [
                                    "sheet",
                                    "scenario",
                                    "fuel_label",
                                    "effective_parent_name",
                                    "esto_flow_norm",
                                    "flow_source",
                                    "min_sector_depth",
                                    "economy",
                                ]:
                                    expanded_df[col] = template[col]
                                share_rows_for_projection.append(expanded_df)

                            if share_rows_for_projection:
                                projection_share_rows = pd.concat(
                                    share_rows_for_projection,
                                    ignore_index=True,
                                    sort=False,
                                )
                                projection_parent = (
                                    projection_rows.groupby(group_cols, as_index=False)["value"]
                                    .agg(
                                        first_non_null=lambda s: s.dropna().iloc[0] if len(s.dropna()) else float("nan"),
                                        unique_non_null=lambda s: s.dropna().nunique(),
                                    )
                                )
                                projection_parent = projection_parent[projection_parent["unique_non_null"] <= 1].copy()
                                alloc_projection_rows = projection_share_rows.merge(
                                    projection_parent[group_cols + ["first_non_null"]],
                                    on=["scenario", "fuel_label", "year", "effective_parent_name"],
                                    how="inner",
                                )
                                if not alloc_projection_rows.empty:
                                    alloc_projection_rows["value"] = (
                                        alloc_projection_rows["detail_share"] * alloc_projection_rows["first_non_null"]
                                    )
                                    existing_projection = comp[
                                        comp["source"].eq("projection")
                                    ][["sheet", "scenario", "fuel_label", "year", "value"]].rename(
                                        columns={"value": "existing_value"}
                                    )
                                    alloc_projection_rows = alloc_projection_rows.merge(
                                        existing_projection,
                                        on=["sheet", "scenario", "fuel_label", "year"],
                                        how="left",
                                    )
                                    existing_projection_num = pd.to_numeric(
                                        alloc_projection_rows["existing_value"], errors="coerce"
                                    )
                                    alloc_projection_rows = alloc_projection_rows[
                                        existing_projection_num.isna() | existing_projection_num.ne(0)
                                    ].copy()
                                    alloc_projection_rows["source"] = "projection_estimated"
                                    projection_replace_keys = alloc_projection_rows[
                                        ["sheet", "scenario", "fuel_label", "source", "year"]
                                    ].drop_duplicates()
                                    projection_replace_keys["source"] = "projection"
                                    comp = comp.merge(
                                        projection_replace_keys.assign(_replace=True),
                                        on=["sheet", "scenario", "fuel_label", "source", "year"],
                                        how="left",
                                    )
                                    comp = comp[comp["_replace"] != True].drop(
                                        columns=["_replace"],
                                        errors="ignore",
                                    )
                                    comp = pd.concat([comp, alloc_projection_rows[keep_cols]], ignore_index=True, sort=False)
                else:
                    # Keep parent comparators at the parent chart level and remove
                    # repeated parent-only rows from each child sheet.
                    fallback_measure_preserve_mask = (
                        fallback_rows["measure"]
                        .fillna("")
                        .astype(str)
                        .str.strip()
                        .str.lower()
                        .map(
                            lambda text: any(
                                token in text
                                for token in ["inputs", "outputs by feedstock", "outputs by product"]
                            )
                        )
                    )
                    child_parent_rows = fallback_rows[
                        fallback_rows["source"].isin(["base", "projection"])
                        & ~fallback_measure_preserve_mask
                    ][["sheet", "scenario", "fuel_label", "source", "year"]].drop_duplicates()
                    if not child_parent_rows.empty:
                        comp = comp.merge(
                            child_parent_rows.assign(_replace=True),
                            on=["sheet", "scenario", "fuel_label", "source", "year"],
                            how="left",
                        )
                        comp = comp[comp["_replace"] != True].drop(
                            columns=["_replace"],
                            errors="ignore",
                        )

            # Promote true parent-level charts named by the effective parent category (e.g. Road).
            if include_sibling_parent_totals:
                if sibling_mode == "aggregate_to_parent" and fallback_mask.any():
                    parent_rows = fallback_rows[fallback_rows["source"].isin(["base", "projection"])].copy()
                    if not parent_rows.empty:
                        parent_rows["effective_comparator_sector_code"] = (
                            parent_rows.get("effective_comparator_sector_code", "")
                            .fillna("")
                            .astype(str)
                            .str.strip()
                            .str.lower()
                        )
                        identified_parent_rows = parent_rows[
                            parent_rows["effective_comparator_sector_code"].ne("")
                        ].copy()
                        unidentified_parent_rows = parent_rows[
                            parent_rows["effective_comparator_sector_code"].eq("")
                        ].copy()

                        parent_parts: list[pd.DataFrame] = []
                        if not identified_parent_rows.empty:
                            identified_grouped = (
                                identified_parent_rows.groupby(
                                    group_cols + ["source", "effective_comparator_sector_code"], as_index=False
                                )
                                .agg(
                                    sum_value=("value", lambda s: s.sum(min_count=1)),
                                    first_non_null=("value", lambda s: s.dropna().iloc[0] if len(s.dropna()) else float("nan")),
                                    unique_non_null=("value", lambda s: s.dropna().nunique()),
                                    contributing_sheets=("sheet", "nunique"),
                                    economy=("economy", "first"),
                                    esto_flow_norm=("esto_flow_norm", "first"),
                                    flow_source=("flow_source", "first"),
                                    effective_parent_name=("effective_parent_name", "first"),
                                    sector_codes_list=("sector_codes_list", _merge_sector_codes_lists),
                                    min_sector_depth=("min_sector_depth", "min"),
                                )
                            )
                            identified_grouped["value"] = identified_grouped["sum_value"]
                            collapse_mask = (
                                identified_grouped["unique_non_null"].le(1)
                                & pd.to_numeric(identified_grouped["contributing_sheets"], errors="coerce").fillna(1).gt(1)
                            )
                            identified_grouped.loc[collapse_mask, "value"] = (
                                identified_grouped.loc[collapse_mask, "first_non_null"]
                            )
                            parent_parts.append(identified_grouped)

                        if not unidentified_parent_rows.empty:
                            unidentified_grouped = (
                                unidentified_parent_rows.groupby(group_cols + ["source"], as_index=False)
                                .agg(
                                    sum_value=("value", lambda s: s.sum(min_count=1)),
                                    first_non_null=("value", lambda s: s.dropna().iloc[0] if len(s.dropna()) else float("nan")),
                                    unique_non_null=("value", lambda s: s.dropna().nunique()),
                                    contributing_sheets=("sheet", "nunique"),
                                    economy=("economy", "first"),
                                    esto_flow_norm=("esto_flow_norm", "first"),
                                    flow_source=("flow_source", "first"),
                                    effective_parent_name=("effective_parent_name", "first"),
                                    sector_codes_list=("sector_codes_list", _merge_sector_codes_lists),
                                    min_sector_depth=("min_sector_depth", "min"),
                                )
                            )
                            unidentified_grouped["effective_comparator_sector_code"] = ""
                            unidentified_grouped["value"] = unidentified_grouped["sum_value"]
                            collapse_mask = (
                                unidentified_grouped["unique_non_null"].le(1)
                                & pd.to_numeric(unidentified_grouped["contributing_sheets"], errors="coerce").fillna(1).gt(1)
                            )
                            unidentified_grouped.loc[collapse_mask, "value"] = (
                                unidentified_grouped.loc[collapse_mask, "first_non_null"]
                            )
                            parent_parts.append(unidentified_grouped)

                        if parent_parts:
                            parent_grouped = pd.concat(parent_parts, ignore_index=True, sort=False)
                            parent_grouped["sheet"] = parent_grouped["effective_parent_name"]
                            keep_cols = [
                                "economy",
                                "scenario",
                                "sheet",
                                "fuel_label",
                                "source",
                                "year",
                                "value",
                                "esto_flow_norm",
                                "flow_source",
                                "effective_parent_name",
                                "effective_comparator_sector_code",
                                "sector_codes_list",
                                "min_sector_depth",
                            ]
                            replace_keys = parent_grouped[
                                ["sheet", "scenario", "fuel_label", "source", "year"]
                            ].drop_duplicates()
                            comp = comp.merge(
                                replace_keys.assign(_replace=True),
                                on=["sheet", "scenario", "fuel_label", "source", "year"],
                                how="left",
                            )
                            comp = comp[comp["_replace"] != True].drop(
                                columns=["_replace"],
                                errors="ignore",
                            )
                            comp = pd.concat([comp, parent_grouped[keep_cols]], ignore_index=True, sort=False)
                promoted = comp[
                    (comp["effective_parent_name"] != "")
                    & comp["source"].isin(["leap", "projection", "projection_estimated", "base", "base_estimated"])
                ].copy()
                if not promoted.empty:
                    promoted_measure_preserve_mask = (
                        promoted["measure"]
                        .fillna("")
                        .astype(str)
                        .str.strip()
                        .str.lower()
                        .map(
                            lambda text: any(
                                token in text
                                for token in ["inputs", "outputs by feedstock", "outputs by product"]
                            )
                        )
                    )
                    promoted = promoted[~promoted_measure_preserve_mask].copy()
                if not promoted.empty:
                    promoted["original_sheet"] = promoted["sheet"].astype(str).str.strip()
                    min_depth = (
                        promoted.groupby(group_cols, as_index=False)["min_sector_depth"]
                        .min()
                        .rename(columns={"min_sector_depth": "group_min_depth"})
                    )
                    promoted = promoted.merge(min_depth, on=group_cols, how="left")
                    promoted = promoted[promoted["min_sector_depth"] == promoted["group_min_depth"]].copy()
                    promoted = promoted[
                        ~promoted.apply(
                            lambda row: _is_sheet_parent_alias(
                                row.get("original_sheet"),
                                row.get("effective_parent_name"),
                            ),
                            axis=1,
                        )
                    ].copy()
                    if not promoted.empty:
                        contributor_counts = (
                            promoted.groupby(group_cols + ["source"], as_index=False)["original_sheet"]
                            .nunique()
                            .rename(columns={"original_sheet": "contributing_sheets"})
                        )
                        promoted = promoted.merge(
                            contributor_counts,
                            on=group_cols + ["source"],
                            how="left",
                        )
                        # Keep single-contributor fuels as well. Parent charts need the
                        # full child sum, even when a fuel appears in only one child sheet.
                if not promoted.empty:
                    parent_like_base = promoted[
                        promoted["source"].eq("base")
                        & promoted["uses_parent_flow"]
                        & promoted["allow_parent_estimate"]
                    ].copy()
                    direct_rows = promoted.drop(parent_like_base.index, errors="ignore")

                    promoted_parts: list[pd.DataFrame] = []
                    if not direct_rows.empty:
                        leap_direct_rows = direct_rows[direct_rows["source"].eq("leap")].copy()
                        non_leap_direct_rows = direct_rows[~direct_rows["source"].eq("leap")].copy()

                        if not leap_direct_rows.empty:
                            leap_direct_grouped = (
                                leap_direct_rows.groupby(group_cols + ["source"], as_index=False)
                                .agg(
                                    sum_value=("value", lambda s: s.sum(min_count=1)),
                                    first_non_null=("value", lambda s: s.dropna().iloc[0] if len(s.dropna()) else float("nan")),
                                    unique_non_null=("value", lambda s: s.dropna().nunique()),
                                    contributing_sheets=("contributing_sheets", "max"),
                                )
                            )
                            leap_direct_grouped["value"] = leap_direct_grouped["sum_value"]
                            leap_direct_grouped["effective_comparator_sector_code"] = ""
                            leap_direct_grouped["sector_codes_list"] = [[] for _ in range(len(leap_direct_grouped))]
                            promoted_parts.append(
                                leap_direct_grouped[
                                    group_cols + ["source", "effective_comparator_sector_code", "sector_codes_list", "value"]
                                ]
                            )

                        if not non_leap_direct_rows.empty:
                            comparator_group_cols = group_cols + ["source", "effective_comparator_sector_code"]
                            non_leap_identified = non_leap_direct_rows[
                                non_leap_direct_rows["effective_comparator_sector_code"].astype(str).str.strip().ne("")
                            ].copy()
                            non_leap_unidentified = non_leap_direct_rows[
                                non_leap_direct_rows["effective_comparator_sector_code"].astype(str).str.strip().eq("")
                            ].copy()

                            if not non_leap_identified.empty:
                                by_identity = (
                                    non_leap_identified.groupby(comparator_group_cols, as_index=False)
                                    .agg(
                                        sum_value=("value", lambda s: s.sum(min_count=1)),
                                        first_non_null=("value", lambda s: s.dropna().iloc[0] if len(s.dropna()) else float("nan")),
                                        unique_non_null=("value", lambda s: s.dropna().nunique()),
                                        contributing_sheets=("contributing_sheets", "max"),
                                        sector_codes_list=("sector_codes_list", _merge_sector_codes_lists),
                                    )
                                )
                                collapse_mask = (
                                    by_identity["unique_non_null"].le(1)
                                    & pd.to_numeric(by_identity["contributing_sheets"], errors="coerce").fillna(1).gt(1)
                                )
                                by_identity["value"] = by_identity["sum_value"]
                                by_identity.loc[collapse_mask, "value"] = by_identity.loc[collapse_mask, "first_non_null"]

                                keep_mask = pd.Series(True, index=by_identity.index)
                                for _, idx_group in by_identity.groupby(group_cols + ["source"], dropna=False).groups.items():
                                    idx_list = list(idx_group)
                                    group_slice = by_identity.loc[idx_list].copy()
                                    group_slice["__depth"] = group_slice["effective_comparator_sector_code"].map(_numeric_sector_depth)
                                    group_slice = group_slice.sort_values(["__depth", "effective_comparator_sector_code"])
                                    codes = group_slice["effective_comparator_sector_code"].fillna("").astype(str).tolist()
                                    values_by_code = {
                                        str(row["effective_comparator_sector_code"]): float(row["value"])
                                        for _, row in group_slice.iterrows()
                                    }
                                    group_keep = {code: True for code in codes if code}

                                    for code in codes:
                                        if not code or not group_keep.get(code, True):
                                            continue
                                        desc_codes = [
                                            other
                                            for other in codes
                                            if other and other != code and group_keep.get(other, True) and _is_sector_ancestor_code(code, other)
                                        ]
                                        if not desc_codes:
                                            continue
                                        ancestor_value = values_by_code.get(code)
                                        descendant_sum = sum(values_by_code.get(other, 0.0) for other in desc_codes)
                                        if _values_close_enough(descendant_sum, ancestor_value):
                                            group_keep[code] = False
                                        else:
                                            for other in desc_codes:
                                                group_keep[other] = False

                                    drop_codes = [code for code, keep in group_keep.items() if not keep]
                                    if drop_codes:
                                        keep_mask.loc[
                                            by_identity.index.isin(idx_list)
                                            & by_identity["effective_comparator_sector_code"].isin(drop_codes)
                                        ] = False
                                by_identity = by_identity[keep_mask].copy()
                                if not by_identity.empty:
                                    promoted_parts.append(
                                        by_identity[
                                            group_cols
                                            + ["source", "effective_comparator_sector_code", "sector_codes_list", "value"]
                                        ]
                                    )

                            if not non_leap_unidentified.empty:
                                unidentified_grouped = (
                                    non_leap_unidentified.groupby(group_cols + ["source"], as_index=False)
                                    .agg(
                                        sum_value=("value", lambda s: s.sum(min_count=1)),
                                        first_non_null=("value", lambda s: s.dropna().iloc[0] if len(s.dropna()) else float("nan")),
                                        unique_non_null=("value", lambda s: s.dropna().nunique()),
                                        contributing_sheets=("contributing_sheets", "max"),
                                    )
                                )
                                collapse_mask = (
                                    unidentified_grouped["unique_non_null"].le(1)
                                    & pd.to_numeric(unidentified_grouped["contributing_sheets"], errors="coerce").fillna(1).gt(1)
                                )
                                unidentified_grouped["value"] = unidentified_grouped["sum_value"]
                                unidentified_grouped.loc[collapse_mask, "value"] = (
                                    unidentified_grouped.loc[collapse_mask, "first_non_null"]
                                )
                                unidentified_grouped["effective_comparator_sector_code"] = ""
                                unidentified_grouped["sector_codes_list"] = [[] for _ in range(len(unidentified_grouped))]
                                promoted_parts.append(
                                    unidentified_grouped[
                                        group_cols + ["source", "effective_comparator_sector_code", "sector_codes_list", "value"]
                                    ]
                                )

                    if not parent_like_base.empty:
                        parent_grouped = (
                            parent_like_base.groupby(group_cols + ["source"], as_index=False)["value"]
                            .agg(
                                first_non_null=lambda s: s.dropna().iloc[0] if len(s.dropna()) else float("nan"),
                                unique_non_null=lambda s: s.dropna().nunique(),
                            )
                        )
                        parent_grouped = parent_grouped[parent_grouped["unique_non_null"] <= 1].copy()
                        if not parent_grouped.empty:
                            parent_grouped["value"] = parent_grouped["first_non_null"]
                            parent_grouped["effective_comparator_sector_code"] = ""
                            parent_grouped["sector_codes_list"] = [[] for _ in range(len(parent_grouped))]
                            promoted_parts.append(
                                parent_grouped[
                                    group_cols + ["source", "effective_comparator_sector_code", "sector_codes_list", "value"]
                                ]
                            )

                    if promoted_parts:
                        promoted = pd.concat(promoted_parts, ignore_index=True, sort=False)
                        # preserve economy code per source where possible
                        econ_ref = (
                            comp[
                                (comp["effective_parent_name"] != "")
                                & comp["source"].isin(["leap", "projection", "projection_estimated", "base", "base_estimated"])
                            ]
                            .groupby(group_cols + ["source"], as_index=False)["economy"]
                            .first()
                        )
                        promoted = promoted.merge(econ_ref, on=group_cols + ["source"], how="left")
                        promoted["sheet"] = promoted["effective_parent_name"]
                        keep_cols = [
                            "economy",
                            "scenario",
                            "sheet",
                            "fuel_label",
                            "source",
                            "year",
                            "value",
                            "effective_comparator_sector_code",
                            "sector_codes_list",
                        ]
                        replace_keys = promoted[
                                ["sheet", "scenario", "fuel_label", "source", "year"]
                        ].drop_duplicates()
                        comp = comp.merge(
                            replace_keys.assign(_replace_parent=True),
                            on=["sheet", "scenario", "fuel_label", "source", "year"],
                            how="left",
                        )
                        comp = comp[comp["_replace_parent"] != True].drop(
                            columns=["_replace_parent"],
                            errors="ignore",
                        )
                        comp = pd.concat([comp, promoted[keep_cols]], ignore_index=True, sort=False)

                if sibling_mode == "aggregate_to_parent":
                    scoped = comp[comp["effective_parent_name"] != ""].copy()
                    if not scoped.empty:
                        scoped["sheet_norm"] = scoped["sheet"].astype(str).str.strip().str.lower()
                        scoped["parent_norm"] = scoped["effective_parent_name"].astype(str).str.strip().str.lower()
                        scoped["is_parent_sheet"] = scoped["sheet_norm"] == scoped["parent_norm"]
                        scoped["is_non_leap"] = scoped["source"].astype(str).str.strip().str.lower() != "leap"

                        child_measure_preserve_mask = (
                            scoped["measure"]
                            .fillna("")
                            .astype(str)
                            .str.strip()
                            .str.lower()
                            .map(
                                lambda text: any(
                                    token in text
                                    for token in ["inputs", "outputs by feedstock", "outputs by product"]
                                )
                            )
                        )

                        parent_non_leap = (
                            scoped[scoped["is_parent_sheet"] & scoped["is_non_leap"]]
                            [["scenario", "fuel_label", "effective_parent_name"]]
                            .drop_duplicates()
                        )
                        if not parent_non_leap.empty:
                            child_stats = (
                                scoped[~scoped["is_parent_sheet"] & ~child_measure_preserve_mask]
                                .groupby(["sheet", "scenario", "fuel_label", "effective_parent_name"], as_index=False)
                                .agg(has_non_leap=("is_non_leap", "any"))
                            )
                            leap_only_children = child_stats[
                                ~child_stats["has_non_leap"]
                            ][["sheet", "scenario", "fuel_label", "effective_parent_name"]]
                            if not leap_only_children.empty:
                                drop_keys = leap_only_children.merge(
                                    parent_non_leap,
                                    on=["scenario", "fuel_label", "effective_parent_name"],
                                    how="inner",
                                )
                                if not drop_keys.empty:
                                    comp = comp.merge(
                                        drop_keys.assign(_drop_child=True),
                                        on=["sheet", "scenario", "fuel_label", "effective_parent_name"],
                                        how="left",
                                    )
                                    # Keep LEAP rows on child sheets so parent totals can still
                                    # include LEAP, while removing duplicated non-LEAP comparators.
                                    child_drop_mask = comp["_drop_child"].fillna(False).astype(bool)
                                    is_non_leap_source = (
                                        comp["source"].astype(str).str.strip().str.lower() != "leap"
                                    )
                                    comp = comp[~(child_drop_mask & is_non_leap_source)].drop(
                                        columns=["_drop_child"],
                                        errors="ignore",
                                    )

            comp = _coalesce_duplicate_comparison_rows(comp)
            comparison_long = comp.drop(
                columns=[
                    "esto_flow_norm",
                    "flow_source",
                    "effective_parent_name",
                    "min_sector_depth",
                    "uses_parent_flow",
                    "allow_parent_estimate",
                ],
                errors="ignore",
            )

    comparison_long = _coalesce_exact_comparator_source_overlaps(comparison_long)
    _validate_comparison_long(comparison_long, mapping_status)

    if comparison_long.empty:
        comparison_wide = pd.DataFrame()
    else:
        comparison_wide = (
            comparison_long.pivot_table(
                index=["economy", "scenario", "sheet", "fuel_label", "year"],
                columns="source",
                values="value",
                aggfunc="first",
            )
            .reset_index()
        )
    if use_esto_agg_only and not mapping_status.empty:
        mapping_status["projection_available"] = False
    return comparison_long, comparison_wide, mapping_status


# -----------------------------------------------------------------------------
# Charting and dashboards
# -----------------------------------------------------------------------------
def _safe_token(value: object) -> str:
    text = str(value).strip() if value is not None else ""
    if not text:
        return "item"
    safe = "".join(ch if ch.isalnum() or ch in {"_", "-"} else "_" for ch in text)
    return safe.strip("_") or "item"


def _code_depth(value: object) -> int:
    parts = [part for part in str(_clean_token(value) or "").strip().split("_") if part]
    depth = 0
    for part in parts:
        if not part.isdigit():
            break
        depth += 1
    return depth


def _validate_comparison_long(
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame | None = None,
) -> None:
    """
    Fail fast on known invalid comparison states before chart/dashboard output.

    The comparison builder intentionally supports fallback sources such as
    ``base_estimated``. These validations ensure the fallback logic has replaced
    rows cleanly rather than duplicating or dropping them.
    """
    if comparison_long.empty:
        return

    comp = comparison_long.copy()
    comp["sheet"] = comp["sheet"].astype(str)
    if "measure" not in comp.columns:
        comp["measure"] = ""
    comp["measure"] = comp["measure"].fillna("").astype(str)
    comp["fuel_label"] = comp["fuel_label"].astype(str)
    comp["scenario"] = comp["scenario"].astype(str)
    comp["source"] = comp["source"].astype(str).str.strip()
    comp["year"] = pd.to_numeric(comp["year"], errors="coerce").astype("Int64")

    key_cols = ["sheet", "measure", "fuel_label", "scenario", "year"]

    def _collect_overlap(left_source: str, right_source: str) -> pd.DataFrame:
        overlap = (
            comp[comp["source"].isin([left_source, right_source])]
            .groupby(key_cols, dropna=False)["source"]
            .agg(lambda s: sorted(set(str(v).strip() for v in s if str(v).strip())))
            .reset_index(name="sources")
        )
        return overlap[overlap["sources"].map(lambda vals: left_source in vals and right_source in vals)].copy()

    base_overlap = _collect_overlap("base", "base_estimated")
    if not base_overlap.empty:
        sample = base_overlap.head(12).to_dict("records")
        raise RuntimeError(
            "Invalid comparison output: found both 'base' and 'base_estimated' "
            f"for the same sheet/fuel/scenario/year. Total conflicts: {len(base_overlap)}. "
            f"Examples: {sample}"
        )

    projection_overlap = _collect_overlap("projection", "projection_estimated")
    if not projection_overlap.empty:
        sample = projection_overlap.head(12).to_dict("records")
        raise RuntimeError(
            "Invalid comparison output: found both 'projection' and 'projection_estimated' "
            f"for the same sheet/fuel/scenario/year. Total conflicts: {len(projection_overlap)}. "
            f"Examples: {sample}"
        )

    duplicate_rows = (
        comp.groupby(key_cols + ["source"], dropna=False)
        .size()
        .reset_index(name="row_count")
    )
    duplicate_rows = duplicate_rows[duplicate_rows["row_count"] > 1].copy()
    if not duplicate_rows.empty:
        sample = duplicate_rows.head(12).to_dict("records")
        raise RuntimeError(
            "Invalid comparison output: duplicate comparison rows detected for the same "
            f"sheet/fuel/scenario/year/source. Total duplicates: {len(duplicate_rows)}. "
            f"Examples: {sample}"
        )

    if mapping_status is None or mapping_status.empty:
        return

    status = mapping_status.copy()
    status["sheet"] = status["sheet"].astype(str)
    status["fuel_label"] = status["fuel_label"].astype(str)
    status["projection_mapping_complete"] = status["projection_mapping_complete"].fillna(False).astype(bool)
    if "comparator_scope" not in status.columns:
        status["comparator_scope"] = "child"
    status["comparator_scope"] = status["comparator_scope"].fillna("child").astype(str).str.strip().str.lower()
    projection_expected = status[
        status["projection_mapping_complete"] & status["comparator_scope"].ne("parent")
    ][["sheet", "fuel_label"]].drop_duplicates()
    if projection_expected.empty:
        return

    projection_present = (
        comp[comp["source"].isin(["projection", "projection_estimated"])][["sheet", "fuel_label"]]
        .drop_duplicates()
        .assign(has_projection=True)
    )
    missing_projection = projection_expected.merge(
        projection_present,
        on=["sheet", "fuel_label"],
        how="left",
    )
    missing_projection = missing_projection[missing_projection["has_projection"] != True].copy()
    if not missing_projection.empty:
        sample = missing_projection.head(12).to_dict("records")
        raise RuntimeError(
            "Invalid comparison output: rows marked projection-mappable have no 'projection' series. "
            f"Total affected rows: {len(missing_projection)}. Examples: {sample}"
        )


def _coalesce_exact_comparator_source_overlaps(comparison_long: pd.DataFrame) -> pd.DataFrame:
    """
    Drop weaker fallback comparator rows when a direct comparator exists for the exact same key.

    This keeps the final comparison output deterministic even if earlier fallback
    allocation logic leaves both the direct and estimated row behind.
    """
    if comparison_long.empty:
        return comparison_long

    comp = comparison_long.copy()
    if "measure" not in comp.columns:
        comp["measure"] = ""
    comp["measure"] = comp["measure"].fillna("").astype(str)
    comp["sheet"] = comp["sheet"].astype(str)
    comp["fuel_label"] = comp["fuel_label"].astype(str)
    comp["scenario"] = comp["scenario"].astype(str)
    comp["source"] = comp["source"].astype(str).str.strip()
    comp["year"] = pd.to_numeric(comp["year"], errors="coerce").astype("Int64")

    key_cols = ["sheet", "measure", "fuel_label", "scenario", "year"]

    for direct_source, estimated_source in [
        ("base", "base_estimated"),
        ("projection", "projection_estimated"),
    ]:
        direct_keys = comp[comp["source"].eq(direct_source)][key_cols].drop_duplicates()
        if direct_keys.empty:
            continue
        comp = comp.merge(
            direct_keys.assign(_drop_estimated=True),
            on=key_cols,
            how="left",
        )
        drop_mask = comp["source"].eq(estimated_source) & comp["_drop_estimated"].fillna(False).astype(bool)
        comp = comp[~drop_mask].drop(columns=["_drop_estimated"], errors="ignore")

    return comp


def _collapse_chart_value(values: pd.Series, *, tol: float = 1e-9) -> float:
    numeric = pd.to_numeric(values, errors="coerce").dropna().tolist()
    if not numeric:
        return float("nan")
    first = float(numeric[0])
    if all(abs(float(val) - first) <= tol for val in numeric[1:]):
        return first
    non_zero = [float(val) for val in numeric if abs(float(val)) > tol]
    if not non_zero:
        return 0.0
    ref = non_zero[0]
    if all(abs(val - ref) <= tol for val in non_zero[1:]):
        return ref
    return max(non_zero, key=lambda val: abs(val))


def _pick_base_family_output_source(values: Sequence[object]) -> str:
    normalized = {str(value).strip() for value in values if str(value).strip()}
    if not normalized:
        return "base"
    if normalized == {"base"}:
        return "base"
    if normalized == {"base_estimated"}:
        return "base_estimated"
    return "base_mixed"


def _pick_projection_family_output_source(values: Sequence[object]) -> str:
    normalized = {str(value).strip() for value in values if str(value).strip()}
    if not normalized:
        return "projection"
    if normalized == {"projection"}:
        return "projection"
    if normalized == {"projection_estimated"}:
        return "projection_estimated"
    return "projection_mixed"


def _collapse_base_family_rows_for_display(frame: pd.DataFrame) -> pd.DataFrame:
    """
    Collapse base-family rows to one scenario-neutral row per chart line.

    Base-year comparator values are shared across REF/TGT in the dashboard.
    When REF and TGT both carry the same sector/fuel bucket, keep a single
    rendered base row keyed only by sheet/measure/fuel/year. If provenance is
    mixed across direct and estimated sources, preserve that via
    ``source=base_mixed`` while still emitting one chart point.
    """
    if frame.empty or "source" not in frame.columns:
        return frame.copy()

    out = frame.copy()
    if "measure" not in out.columns:
        out["measure"] = ""
    out["measure"] = out["measure"].fillna("").astype(str)
    if "scenario" not in out.columns:
        out["scenario"] = ""
    out["scenario"] = out["scenario"].fillna("").astype(str)
    out["source"] = out["source"].fillna("").astype(str).str.strip()
    out["value"] = pd.to_numeric(out["value"], errors="coerce")

    base_sources = {"base", "base_estimated", "base_mixed"}
    base_mask = out["source"].isin(base_sources)
    if not base_mask.any():
        return out

    base_rows = out.loc[base_mask].copy()
    other_rows = out.loc[~base_mask].copy()
    group_keys = [
        col
        for col in ["economy", "sheet", "structural_component_sheet", "measure", "fuel_label", "year"]
        if col in base_rows.columns
    ]

    # Prefer direct base rows when present, but fall back to mixed/estimated
    # rows if the higher-priority source only provides null/zero placeholders.
    source_preference = {"base": 0, "base_mixed": 1, "base_estimated": 2}
    base_rows["_source_rank"] = base_rows["source"].map(source_preference).fillna(9).astype(int)
    base_rows["_scenario_blank"] = base_rows["scenario"].eq("")
    base_rows = base_rows.sort_values(group_keys + ["_source_rank", "_scenario_blank", "scenario"], kind="stable")
    representative = (
        base_rows.groupby(group_keys, dropna=False, as_index=False)
        .first()
        .drop(columns=["_source_rank", "_scenario_blank"], errors="ignore")
    )

    provenance = (
        base_rows.groupby(group_keys, dropna=False)["source"]
        .agg(lambda s: _pick_base_family_output_source(list(s)))
        .reset_index(name="_base_source")
    )

    def _pick_group_value(group: pd.DataFrame) -> float:
        zero_candidate: float | None = None
        for source_name in ["base", "base_mixed", "base_estimated"]:
            source_rows = group[group["source"].eq(source_name)]
            if source_rows.empty:
                continue
            value = _collapse_chart_value(source_rows["value"])
            if pd.isna(value):
                continue
            if abs(float(value)) > 1e-9:
                return float(value)
            if zero_candidate is None:
                zero_candidate = float(value)
        fallback = _collapse_chart_value(group["value"])
        if pd.notna(fallback):
            return float(fallback)
        if zero_candidate is not None:
            return zero_candidate
        return float("nan")

    chosen_values = (
        base_rows.groupby(group_keys, dropna=False)
        .apply(_pick_group_value, include_groups=False)
        .reset_index(name="_base_value")
    )

    bool_cols = [col for col in ["force_show_chart", "mapped", "partially_mapped", "has_any_mapping"] if col in base_rows.columns]
    bool_meta = None
    if bool_cols:
        bool_meta = (
            base_rows.groupby(group_keys, dropna=False)[bool_cols]
            .max()
            .reset_index()
        )

    collapsed = representative.merge(provenance, on=group_keys, how="left")
    collapsed = collapsed.merge(chosen_values, on=group_keys, how="left")
    if bool_meta is not None:
        collapsed = collapsed.drop(columns=bool_cols, errors="ignore").merge(bool_meta, on=group_keys, how="left")
    collapsed["source"] = collapsed["_base_source"].fillna("base")
    collapsed["value"] = pd.to_numeric(collapsed["_base_value"], errors="coerce")
    collapsed["scenario"] = ""
    collapsed = collapsed.drop(columns=["_base_source", "_base_value"], errors="ignore")

    return pd.concat([other_rows, collapsed], ignore_index=True, sort=False)


def _collapse_projection_family_rows_for_display(frame: pd.DataFrame) -> pd.DataFrame:
    """
    Collapse projection-family rows to one scenario-specific row per chart line.

    Unlike the base-family collapse, projection rows remain scenario-specific.
    When both direct and estimated comparator totals contribute to the same
    sheet/measure/fuel/scenario/year point, emit one ``projection_mixed`` row
    with the summed value.
    """
    if frame.empty or "source" not in frame.columns:
        return frame.copy()

    out = frame.copy()
    if "measure" not in out.columns:
        out["measure"] = ""
    out["measure"] = out["measure"].fillna("").astype(str)
    if "scenario" not in out.columns:
        out["scenario"] = ""
    out["scenario"] = out["scenario"].fillna("").astype(str)
    out["source"] = out["source"].fillna("").astype(str).str.strip()
    out["value"] = pd.to_numeric(out["value"], errors="coerce")

    projection_sources = {"projection", "projection_estimated", "projection_mixed"}
    proj_mask = out["source"].isin(projection_sources)
    if not proj_mask.any():
        return out

    proj_rows = out.loc[proj_mask].copy()
    other_rows = out.loc[~proj_mask].copy()
    group_keys = [
        col
        for col in ["economy", "sheet", "structural_component_sheet", "measure", "fuel_label", "scenario", "year"]
        if col in proj_rows.columns
    ]
    if not group_keys:
        return out

    representative = (
        proj_rows.groupby(group_keys, dropna=False, as_index=False)
        .first()
    )
    provenance = (
        proj_rows.groupby(group_keys, dropna=False)["source"]
        .agg(lambda s: _pick_projection_family_output_source(list(s)))
        .reset_index(name="_projection_source")
    )
    grouped_values = (
        proj_rows.groupby(group_keys, dropna=False, as_index=False)["value"]
        .sum(min_count=1)
        .rename(columns={"value": "_projection_value"})
    )
    bool_cols = [col for col in ["force_show_chart", "mapped", "partially_mapped", "has_any_mapping"] if col in proj_rows.columns]
    bool_meta = None
    if bool_cols:
        bool_meta = (
            proj_rows.groupby(group_keys, dropna=False)[bool_cols]
            .max()
            .reset_index()
        )

    collapsed = representative.merge(provenance, on=group_keys, how="left")
    collapsed = collapsed.merge(grouped_values, on=group_keys, how="left")
    if bool_meta is not None:
        collapsed = collapsed.drop(columns=bool_cols, errors="ignore").merge(bool_meta, on=group_keys, how="left")
    collapsed["source"] = collapsed["_projection_source"].fillna("projection")
    collapsed["value"] = pd.to_numeric(collapsed["_projection_value"], errors="coerce")
    collapsed = collapsed.drop(columns=["_projection_source", "_projection_value"], errors="ignore")

    return pd.concat([other_rows, collapsed], ignore_index=True, sort=False)


def _aggregate_display_rows_to_total(
    frame: pd.DataFrame,
    *,
    title: str,
    measure_value: str,
    collapse_base_family: bool = False,
    collapse_projection_family: bool = False,
) -> pd.DataFrame:
    """
    Sum already-resolved display rows into a chartable ``Total`` series.

    This keeps dashboard/node totals aligned with the V2 rule of aggregating
    the rows actually shown on the child charts, instead of rerunning the
    legacy total builder on summary pages.
    """
    if frame.empty:
        return pd.DataFrame()

    total = frame.copy()
    total["value"] = pd.to_numeric(total["value"], errors="coerce")
    if "measure" not in total.columns:
        total["measure"] = ""
    total["measure"] = str(measure_value).strip()
    total["sheet"] = str(title).strip()
    total["fuel_label"] = "Total"

    group_cols = ["sheet", "measure", "fuel_label", "scenario", "source", "year"]
    if "economy" in total.columns:
        group_cols = ["economy"] + group_cols

    total = total.groupby(group_cols, as_index=False)["value"].sum(min_count=1)
    if collapse_base_family:
        total = _collapse_base_family_rows_for_display(total)
    if collapse_projection_family:
        total = _collapse_projection_family_rows_for_display(total)
    return total[
        (total["sheet"].astype(str) == str(title).strip())
        & (total["fuel_label"].astype(str) == "Total")
        & (total["measure"].astype(str) == str(measure_value).strip())
    ].copy()


def _derive_total_rows_from_display_rows(
    comparison_long: pd.DataFrame,
    *,
    collapse_base_family: bool = True,
    collapse_projection_family: bool = False,
) -> pd.DataFrame:
    """
    Derive `Total` rows from existing display rows using the V2 chart policy.

    This is the fallback for callers that provide chart rows without any
    precomputed `Total` rows. By default it keeps projection families separate
    and only collapses the base family, matching V2 chart totals.
    """
    if comparison_long.empty:
        return comparison_long.copy()

    frame = comparison_long.copy()
    if "measure" not in frame.columns:
        frame["measure"] = ""
    frame["measure"] = frame["measure"].fillna("").astype(str)
    if "fuel_label" not in frame.columns:
        frame["fuel_label"] = ""
    frame["fuel_label"] = frame["fuel_label"].fillna("").astype(str)

    non_total = frame[frame["fuel_label"].ne("Total")].copy()
    if non_total.empty:
        return frame

    totals_parts: list[pd.DataFrame] = []
    for (sheet, measure), subset in non_total.groupby(["sheet", "measure"], dropna=False):
        total = _aggregate_display_rows_to_total(
            subset,
            title=str(sheet).strip(),
            measure_value=str(measure).strip(),
            collapse_base_family=collapse_base_family,
            collapse_projection_family=collapse_projection_family,
        )
        if not total.empty:
            totals_parts.append(total)

    if not totals_parts:
        return frame

    totals = pd.concat(totals_parts, ignore_index=True, sort=False)
    return pd.concat([non_total, totals], ignore_index=True, sort=False)


def make_chart(
    sheet: str,
    fuel: str,
    subset: pd.DataFrame,
    output_dir: Path,
    backend: str = "plotly",
    display_sheet: str | None = None,
    file_sheet: str | None = None,
) -> Path | None:
    """Generate a simple comparison chart for one sheet/fuel."""
    output_dir.mkdir(parents=True, exist_ok=True)
    display_name = display_sheet or sheet
    escaped_display_name = escape(display_name)
    escaped_fuel = escape(fuel)
    sheet_slug = _safe_token((file_sheet or sheet).replace("\\", "_"))
    fuel_slug = _safe_token(fuel)
    out_png = output_dir / f"{sheet_slug}__{fuel_slug}.png"
    out_html = output_dir / f"{sheet_slug}__{fuel_slug}.html"
    y_axis_label = "Energy (PJ)"

    def _format_scenario_label(value: object) -> str:
        raw = str(value or "").strip()
        if not raw:
            return "Scenario"
        lowered = raw.lower()
        if lowered == "reference":
            return "Reference"
        if lowered == "target":
            return "Target"
        return raw

    def _compact_scenario_label(value: object) -> str:
        raw = str(value or "").strip()
        if not raw:
            return ""
        lowered = raw.lower()
        if lowered == "reference":
            return "REF"
        if lowered == "target":
            return "TGT"
        return raw

    def _trace_label(base: str, scen: object = "", *, show_scenario: bool = False) -> str:
        label = str(base or "").strip()
        if not show_scenario:
            return label
        compact_scenario = _compact_scenario_label(scen)
        if not compact_scenario:
            return label
        return f"{label} {compact_scenario}"

    def _series_by_scenario(source_name: str) -> dict[str, pd.Series]:
        out: dict[str, pd.Series] = {}
        src = subset[subset["source"] == source_name].copy()
        if src.empty:
            return out
        for scenario, g in src.groupby("scenario", dropna=False):
            scen_label = _format_scenario_label(scenario)
            s = (
                g.sort_values("year")
                .groupby("year", as_index=True)["value"]
                .first()
            )
            out[scen_label] = pd.to_numeric(s, errors="coerce")
        return out

    def _series_not_used_in_total(source_name: str, scenario_label: str) -> bool:
        if "used_in_total" not in subset.columns:
            return False
        src = subset[subset["source"] == source_name].copy()
        if src.empty:
            return False
        src["scenario_label"] = src["scenario"].map(_format_scenario_label)
        src = src[src["scenario_label"] == str(scenario_label)]
        if src.empty:
            return False
        text = src["used_in_total"].fillna(True).astype(str).str.strip().str.lower()
        true_tokens = {"1", "true", "yes", "y", "t"}
        false_tokens = {"0", "false", "no", "n", "f"}
        flags = pd.Series(True, index=src.index, dtype="bool")
        flags.loc[text.isin(true_tokens)] = True
        flags.loc[text.isin(false_tokens)] = False
        return bool(len(flags) and not flags.any())

    comparator_sources = {"base", "base_estimated", "base_mixed", "projection", "projection_estimated", "projection_mixed"}
    chart_has_all_comparators_not_used = False
    if "used_in_total" in subset.columns and "source" in subset.columns:
        comp_scope = subset[subset["source"].astype(str).isin(comparator_sources)].copy()
        if not comp_scope.empty:
            text = comp_scope["used_in_total"].fillna(True).astype(str).str.strip().str.lower()
            true_tokens = {"1", "true", "yes", "y", "t"}
            false_tokens = {"0", "false", "no", "n", "f"}
            flags = pd.Series(True, index=comp_scope.index, dtype="bool")
            flags.loc[text.isin(true_tokens)] = True
            flags.loc[text.isin(false_tokens)] = False
            chart_has_all_comparators_not_used = bool(not flags.any())

    def _trace_label_with_total_scope(
        base: str,
        source_name: str,
        scen: object = "",
        *,
        show_scenario: bool = False,
    ) -> str:
        label = _trace_label(base, scen, show_scenario=show_scenario)
        if source_name in comparator_sources and (
            chart_has_all_comparators_not_used or _series_not_used_in_total(source_name, str(scen))
        ):
            return f"{label} (not used in total)"
        return label

    def _plotly_xy(series: pd.Series) -> tuple[list[object], list[float | None]]:
        years = pd.Index(series.index).tolist()
        numeric = pd.to_numeric(series, errors="coerce")
        values = [None if pd.isna(value) else float(value) for value in numeric.tolist()]
        return years, values

    def _series_maps_match(left: pd.Series, right: pd.Series, *, tol: float = 1e-9) -> bool:
        years = left.index.union(right.index)
        lvals = pd.to_numeric(left.reindex(years), errors="coerce")
        rvals = pd.to_numeric(right.reindex(years), errors="coerce")
        both_null = lvals.isna() & rvals.isna()
        close = (lvals - rvals).abs().le(tol)
        return bool((both_null | close).fillna(False).all())

    def _collapse_identical_scenarios(series_map: dict[str, pd.Series]) -> dict[str, pd.Series]:
        if len(series_map) <= 1:
            return series_map
        ordered_keys = sorted(
            series_map.keys(),
            key=lambda key: (str(key).lower() not in {"reference", "target"}, str(key).lower()),
        )
        first_key = ordered_keys[0]
        first_series = series_map[first_key]
        if all(_series_maps_match(first_series, series_map[key]) for key in ordered_keys[1:]):
            return {first_key: first_series}
        return series_map

    leap_by_scenario = _series_by_scenario("leap")
    base_by_scenario = _series_by_scenario("base")
    base_est_by_scenario = _series_by_scenario("base_estimated")
    base_mixed_by_scenario = _series_by_scenario("base_mixed")
    proj_by_scenario = _series_by_scenario("projection")
    proj_est_by_scenario = _series_by_scenario("projection_estimated")
    proj_mixed_by_scenario = _series_by_scenario("projection_mixed")
    base_by_scenario = _collapse_identical_scenarios(base_by_scenario)
    base_est_by_scenario = _collapse_identical_scenarios(base_est_by_scenario)
    base_mixed_by_scenario = _collapse_identical_scenarios(base_mixed_by_scenario)
    projection_family_scenarios = set(proj_by_scenario) | set(proj_est_by_scenario) | set(proj_mixed_by_scenario)
    base_family_scenarios = set(base_by_scenario) | set(base_est_by_scenario) | set(base_mixed_by_scenario)

    def _scenario_key(scen: str) -> str:
        normalized = str(scen or "").strip().lower()
        if normalized == "reference":
            return "reference"
        if normalized == "target":
            return "target"
        return normalized or "other"

    def _projection_trace_label(scen: str) -> str:
        return _trace_label("9th projection", scen, show_scenario=True)

    def _trace_color(trace_kind: str, scen: str) -> str:
        # Unique color per legend item (trace family + scenario).
        keyed_palette = {
            ("leap", "reference"): "#1f77b4",
            ("leap", "target"): "#17becf",
            ("projection", "reference"): "#2ca02c",
            ("projection", "target"): "#bcbd22",
            ("projection_mixed", "reference"): "#008b8b",
            ("projection_mixed", "target"): "#20b2aa",
            ("projection_estimated", "reference"): "#ff7f0e",
            ("projection_estimated", "target"): "#ffbb78",
            ("base", "reference"): "#d62728",
            ("base", "target"): "#e377c2",
            ("base_mixed", "reference"): "#8c564b",
            ("base_mixed", "target"): "#c49c94",
            ("base_estimated", "reference"): "#9467bd",
            ("base_estimated", "target"): "#7f7f7f",
        }
        scen_key = _scenario_key(scen)
        if (trace_kind, scen_key) in keyed_palette:
            return keyed_palette[(trace_kind, scen_key)]

        # Fallback for non-standard scenario names.
        fallback_palette = {
            "leap": "#1f77b4",
            "projection": "#2ca02c",
            "projection_mixed": "#008b8b",
            "projection_estimated": "#ff7f0e",
            "base": "#d62728",
            "base_mixed": "#8c564b",
            "base_estimated": "#9467bd",
        }
        return fallback_palette.get(trace_kind, "#111111")

    def _mpl_line_style(trace_kind: str) -> str:
        return {
            "leap": "-",
            "projection": "-",
            "projection_mixed": "--",
            "projection_estimated": ":",
        }.get(trace_kind, "-")

    def _plotly_line_dash(trace_kind: str) -> str:
        return {
            "leap": "solid",
            "projection": "solid",
            "projection_mixed": "dash",
            "projection_estimated": "dot",
        }.get(trace_kind, "solid")

    def _mpl_marker(trace_kind: str) -> str:
        return {
            "leap": "o",
            "projection": "s",
            "projection_mixed": "s",
            "projection_estimated": "s",
            "base": "D",
            "base_mixed": "d",
            "base_estimated": "P",
        }.get(trace_kind, "o")

    def _plotly_marker_symbol(trace_kind: str) -> str:
        return {
            "leap": "circle",
            "projection": "square",
            "projection_mixed": "square-open",
            "projection_estimated": "square-dot",
            "base": "diamond",
            "base_mixed": "diamond-open",
            "base_estimated": "diamond-cross",
        }.get(trace_kind, "circle")

    def _render_static() -> Path | None:
        try:
            import matplotlib.pyplot as plt

            plt.figure(figsize=(7, 4))
            for scen, leap_s in sorted(leap_by_scenario.items()):
                if not leap_s.empty:
                    label = _trace_label_with_total_scope(
                        "LEAP",
                        "leap",
                        scen,
                        show_scenario=len(leap_by_scenario) > 1,
                    )
                    plt.plot(
                        leap_s.index,
                        leap_s.values,
                        label=label,
                        marker=_mpl_marker("leap"),
                        linestyle=_mpl_line_style("leap"),
                        color=_trace_color("leap", scen),
                    )
            for scen, proj_s in sorted(proj_by_scenario.items()):
                if not proj_s.empty:
                    label = _trace_label_with_total_scope(
                        "9th projection",
                        "projection",
                        scen,
                        show_scenario=True,
                    )
                    plt.plot(
                        proj_s.index,
                        proj_s.values,
                        label=label,
                        marker=_mpl_marker("projection"),
                        linestyle=_mpl_line_style("projection"),
                        color=_trace_color("projection", scen),
                    )
            for scen, proj_s in sorted(proj_mixed_by_scenario.items()):
                if not proj_s.empty:
                    label = _trace_label_with_total_scope(
                        "9th projection est/real",
                        "projection_mixed",
                        scen,
                        show_scenario=len(projection_family_scenarios) > 1,
                    )
                    plt.plot(
                        proj_s.index,
                        proj_s.values,
                        label=label,
                        marker=_mpl_marker("projection_mixed"),
                        linestyle=_mpl_line_style("projection_mixed"),
                        color=_trace_color("projection_mixed", scen),
                    )
            for scen, proj_s in sorted(proj_est_by_scenario.items()):
                if not proj_s.empty:
                    label = _trace_label_with_total_scope(
                        "9th projection est",
                        "projection_estimated",
                        scen,
                        show_scenario=len(projection_family_scenarios) > 1,
                    )
                    plt.plot(
                        proj_s.index,
                        proj_s.values,
                        label=label,
                        marker=_mpl_marker("projection_estimated"),
                        linestyle=_mpl_line_style("projection_estimated"),
                        color=_trace_color("projection_estimated", scen),
                    )
            for scen, base_s in sorted(base_by_scenario.items()):
                if not base_s.empty:
                    label = _trace_label_with_total_scope(
                        "Base 2022",
                        "base",
                        scen,
                        show_scenario=len(base_family_scenarios) > 1,
                    )
                    plt.scatter(
                        base_s.index,
                        base_s.values,
                        label=label,
                        marker=_mpl_marker("base"),
                        color=_trace_color("base", scen),
                    )
            for scen, base_s in sorted(base_mixed_by_scenario.items()):
                if not base_s.empty:
                    label = _trace_label_with_total_scope(
                        "Base est/real",
                        "base_mixed",
                        scen,
                        show_scenario=len(base_family_scenarios) > 1,
                    )
                    plt.scatter(
                        base_s.index,
                        base_s.values,
                        label=label,
                        marker=_mpl_marker("base_mixed"),
                        color=_trace_color("base_mixed", scen),
                    )
            for scen, base_s in sorted(base_est_by_scenario.items()):
                if not base_s.empty:
                    label = _trace_label_with_total_scope(
                        "Base est",
                        "base_estimated",
                        scen,
                        show_scenario=len(base_family_scenarios) > 1,
                    )
                    plt.scatter(
                        base_s.index,
                        base_s.values,
                        label=label,
                        marker=_mpl_marker("base_estimated"),
                        color=_trace_color("base_estimated", scen),
                    )
            plt.xlabel("Year")
            plt.ylabel(y_axis_label)
            plt.legend()
            plt.tight_layout()
            plt.savefig(out_png, dpi=150)
            plt.close()
            return out_png
        except Exception as exc:  # noqa: BLE001
            print(f"[WARN] Failed to render static chart for {display_name}/{fuel}: {exc}")
            return None

    try:
        if backend == "plotly":
            import plotly.graph_objects as go
            import plotly.io as pio

            fig = go.Figure()
            for scen, leap_s in sorted(leap_by_scenario.items()):
                if leap_s.empty:
                    continue
                name = _trace_label_with_total_scope(
                    "LEAP",
                    "leap",
                    scen,
                    show_scenario=len(leap_by_scenario) > 1,
                )
                xs, ys = _plotly_xy(leap_s)
                fig.add_trace(
                    go.Scatter(
                        x=xs,
                        y=ys,
                        mode="lines+markers",
                        line=dict(color=_trace_color("leap", scen), dash=_plotly_line_dash("leap")),
                        marker=dict(symbol=_plotly_marker_symbol("leap"), color=_trace_color("leap", scen)),
                        name=name,
                    )
                )
            for scen, proj_s in sorted(proj_by_scenario.items()):
                if proj_s.empty:
                    continue
                name = _trace_label_with_total_scope(
                    "9th projection",
                    "projection",
                    scen,
                    show_scenario=True,
                )
                xs, ys = _plotly_xy(proj_s)
                fig.add_trace(
                    go.Scatter(
                        x=xs,
                        y=ys,
                        mode="lines+markers",
                        line=dict(color=_trace_color("projection", scen), dash=_plotly_line_dash("projection")),
                        marker=dict(symbol=_plotly_marker_symbol("projection"), color=_trace_color("projection", scen)),
                        name=name,
                    )
                )
            for scen, proj_s in sorted(proj_mixed_by_scenario.items()):
                if proj_s.empty:
                    continue
                name = _trace_label_with_total_scope(
                    "9th projection est/real",
                    "projection_mixed",
                    scen,
                    show_scenario=len(projection_family_scenarios) > 1,
                )
                xs, ys = _plotly_xy(proj_s)
                fig.add_trace(
                    go.Scatter(
                        x=xs,
                        y=ys,
                        mode="lines+markers",
                        line=dict(
                            color=_trace_color("projection_mixed", scen),
                            dash=_plotly_line_dash("projection_mixed"),
                        ),
                        marker=dict(
                            symbol=_plotly_marker_symbol("projection_mixed"),
                            color=_trace_color("projection_mixed", scen),
                        ),
                        name=name,
                    )
                )
            for scen, proj_s in sorted(proj_est_by_scenario.items()):
                if proj_s.empty:
                    continue
                name = _trace_label_with_total_scope(
                    "9th projection est",
                    "projection_estimated",
                    scen,
                    show_scenario=len(projection_family_scenarios) > 1,
                )
                xs, ys = _plotly_xy(proj_s)
                fig.add_trace(
                    go.Scatter(
                        x=xs,
                        y=ys,
                        mode="lines+markers",
                        line=dict(
                            color=_trace_color("projection_estimated", scen),
                            dash=_plotly_line_dash("projection_estimated"),
                        ),
                        marker=dict(
                            symbol=_plotly_marker_symbol("projection_estimated"),
                            color=_trace_color("projection_estimated", scen),
                        ),
                        name=name,
                    )
                )
            for scen, base_s in sorted(base_by_scenario.items()):
                if base_s.empty:
                    continue
                name = _trace_label_with_total_scope(
                    "Base 2022",
                    "base",
                    scen,
                    show_scenario=len(base_family_scenarios) > 1,
                )
                xs, ys = _plotly_xy(base_s)
                fig.add_trace(
                    go.Scatter(
                        x=xs,
                        y=ys,
                        mode="markers",
                        marker=dict(
                            size=10,
                            symbol=_plotly_marker_symbol("base"),
                            color=_trace_color("base", scen),
                        ),
                        name=name,
                    )
                )
            for scen, base_s in sorted(base_mixed_by_scenario.items()):
                if base_s.empty:
                    continue
                name = _trace_label_with_total_scope(
                    "Base est/real",
                    "base_mixed",
                    scen,
                    show_scenario=len(base_family_scenarios) > 1,
                )
                xs, ys = _plotly_xy(base_s)
                fig.add_trace(
                    go.Scatter(
                        x=xs,
                        y=ys,
                        mode="markers",
                        marker=dict(
                            size=10,
                            symbol=_plotly_marker_symbol("base_mixed"),
                            color=_trace_color("base_mixed", scen),
                        ),
                        name=name,
                    )
                )
            for scen, base_s in sorted(base_est_by_scenario.items()):
                if base_s.empty:
                    continue
                name = _trace_label_with_total_scope(
                    "Base est",
                    "base_estimated",
                    scen,
                    show_scenario=len(base_family_scenarios) > 1,
                )
                xs, ys = _plotly_xy(base_s)
                fig.add_trace(
                    go.Scatter(
                        x=xs,
                        y=ys,
                        mode="markers",
                        marker=dict(
                            size=10,
                            symbol=_plotly_marker_symbol("base_estimated"),
                            color=_trace_color("base_estimated", scen),
                        ),
                        name=name,
                    )
                )
            ninth_pairs_label = ""
            if "ninth_pairs_label" in subset.columns:
                labels = subset["ninth_pairs_label"].dropna().astype(str).str.strip()
                labels = labels[labels.ne("")]
                if not labels.empty:
                    ninth_pairs_label = labels.iloc[0]
            layout_annotations = []
            layout_margin = dict(l=64, r=28, t=56, b=56)
            if ninth_pairs_label:
                layout_margin = dict(l=64, r=28, t=56, b=72)
                layout_annotations.append(
                    dict(
                        text=f"9th mapping: {escape(ninth_pairs_label)}",
                        xref="paper", yref="paper",
                        x=0, y=-0.13,
                        xanchor="left", yanchor="top",
                        showarrow=False,
                        font=dict(size=9, color="#888"),
                    )
                )
            fig.update_layout(
                autosize=True,
                xaxis_title="Year",
                yaxis_title=y_axis_label,
                template="plotly_white",
                margin=layout_margin,
                annotations=layout_annotations,
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="left",
                    x=0,
                ),
            )
            plot_html = pio.to_html(
                fig,
                include_plotlyjs="cdn",
                full_html=False,
                config={"responsive": True},
            )
            chart_html = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>{escaped_display_name} - {escaped_fuel}</title>
  <style>
    :root {{
      color-scheme: light;
      font-size: clamp(14px, 0.35vw + 13px, 18px);
    }}
    html, body {{
      margin: 0;
      padding: 0;
      width: 100%;
      min-width: 0;
      background: #ffffff;
      color: #111827;
      overflow: hidden;
      font-family: "Segoe UI", Arial, sans-serif;
    }}
    body {{
      min-height: 100vh;
      padding: 0;
      box-sizing: border-box;
    }}
    .chart-shell {{
      width: 100%;
      min-height: 100vh;
    }}
    .plotly-graph-div {{
      width: 100% !important;
      min-height: 100%;
    }}
  </style>
</head>
<body>
  <div class="chart-shell">
    {plot_html}
  </div>
  <script>
    (function() {{
      const chart = document.querySelector('.plotly-graph-div');
      if (!chart || !window.Plotly) return;
      const chartShell = chart.closest('.chart-shell') || chart.parentElement;
      let lastPixelRatio = window.devicePixelRatio || 1;
      const applyViewportTypography = () => {{
        const viewportWidth = Math.max(
          document.documentElement.clientWidth || 0,
          window.innerWidth || 0,
        );
        const viewportHeight = Math.max(
          document.documentElement.clientHeight || 0,
          window.innerHeight || 0,
        );
        const containerWidth = Math.max(
          (chartShell && chartShell.clientWidth) || 0,
          viewportWidth - 20,
        );
        const targetHeight = Math.max(
          360,
          Math.round(viewportHeight * 0.82),
        );
        if (chartShell) {{
          chartShell.style.minHeight = `${{targetHeight}}px`;
        }}
        const scale = Math.max(1, Math.min(1.45, viewportWidth / 1100));
        window.Plotly.relayout(chart, {{
          width: containerWidth,
          height: targetHeight,
          font: {{ size: Math.round(13 * scale) }},
          legend: {{ font: {{ size: Math.round(12 * scale) }} }},
          xaxis: {{
            title: {{ font: {{ size: Math.round(14 * scale) }} }},
            tickfont: {{ size: Math.round(12 * scale) }},
          }},
          yaxis: {{
            title: {{ font: {{ size: Math.round(14 * scale) }} }},
            tickfont: {{ size: Math.round(12 * scale) }},
          }},
        }});
        window.Plotly.Plots.resize(chart);
      }};
      let resizeTimer = null;
      const queueApplyViewportTypography = () => {{
        window.clearTimeout(resizeTimer);
        resizeTimer = window.setTimeout(applyViewportTypography, 80);
      }};
      applyViewportTypography();
      window.addEventListener('resize', queueApplyViewportTypography);
      if (window.visualViewport) {{
        window.visualViewport.addEventListener('resize', queueApplyViewportTypography);
      }}
      if ('ResizeObserver' in window && chartShell) {{
        const observer = new ResizeObserver(queueApplyViewportTypography);
        observer.observe(document.documentElement);
        observer.observe(chartShell);
      }}
      window.setInterval(() => {{
        const pixelRatio = window.devicePixelRatio || 1;
        if (pixelRatio !== lastPixelRatio) {{
          lastPixelRatio = pixelRatio;
          queueApplyViewportTypography();
        }}
      }}, 250);
    }})();
  </script>
</body>
</html>
"""
            out_html.write_text(chart_html, encoding="utf-8")
            return out_html
        return _render_static()
    except Exception as exc:  # noqa: BLE001
        if backend == "plotly":
            print(f"[WARN] Plotly chart failed for {display_name}/{fuel}; falling back to static: {exc}")
            return _render_static()
        print(f"[WARN] Failed to render chart for {display_name}/{fuel}: {exc}")
        return None


def _prepare_render_long(comparison_long: pd.DataFrame) -> pd.DataFrame:
    """
    Use precomputed chart totals when available; otherwise derive them lazily.

    V2 chart inputs already contain audited Total rows built from the resolved
    display set. Recomputing totals here can reintroduce legacy mixed-source
    rollups that diverge from the chart ledgers and rendered child lines.
    """
    render_long = comparison_long.copy()
    if "measure" not in render_long.columns:
        render_long["measure"] = ""
    render_long["measure"] = render_long["measure"].fillna("").astype(str)
    if "fuel_label" not in render_long.columns:
        render_long["fuel_label"] = ""
    render_long["fuel_label"] = render_long["fuel_label"].fillna("").astype(str)

    has_total_rows = render_long["fuel_label"].eq("Total").any()
    if not has_total_rows:
        render_long = _derive_total_rows_from_display_rows(
            render_long,
            collapse_base_family=True,
            collapse_projection_family=False,
        )
        render_long["measure"] = render_long["measure"].fillna("").astype(str)
    return render_long


def build_charts(
    comparison_long: pd.DataFrame,
    charts_dir: Path,
    backend: str = "plotly",
    *,
    hide_leap_only_charts: bool = False,
) -> list[Path]:
    written: list[Path] = []
    if comparison_long.empty:
        return written
    charts_dir.mkdir(parents=True, exist_ok=True)
    render_long = _prepare_render_long(comparison_long)
    sheet_display_metadata = _build_sheet_display_metadata()
    sheet_display_labels = {sheet: meta.get("label", sheet) for sheet, meta in sheet_display_metadata.items()}
    sheet_display_labels.update(
        {
            "MAP electricity plants": "Electricity plants (electricity output)",
            "MAP CHP plants (electricity)": "CHP plants (electricity output)",
        }
    )

    def _chart_sheet_slug(sheet: object, measure: object) -> str:
        measure_text = str(measure or "").strip()
        if measure_text:
            return _safe_token(f"{sheet}__{measure_text}".replace("\\", "_"))
        return _safe_token(str(sheet).replace("\\", "_"))

    def _remove_stale_chart_files(sheet: object, measure: object, fuel: object) -> None:
        sheet_slug = _chart_sheet_slug(sheet, measure)
        fuel_slug = _safe_token(str(fuel))
        for stale_path in [
            charts_dir / f"{sheet_slug}__{fuel_slug}.html",
            charts_dir / f"{sheet_slug}__{fuel_slug}.png",
        ]:
            if stale_path.exists():
                try:
                    stale_path.unlink()
                except OSError:
                    pass

    for (sheet, measure, fuel), sub in render_long.groupby(["sheet", "measure", "fuel_label"]):
        values = pd.to_numeric(sub["value"], errors="coerce").fillna(0.0)
        force_show_chart = (
            bool(sub["force_show_chart"].fillna(False).astype(bool).any())
            if "force_show_chart" in sub.columns
            else False
        )
        if not values.ne(0).any():
            _remove_stale_chart_files(sheet, measure, fuel)
            continue
        if hide_leap_only_charts:
            non_leap_sources = {
                str(src).strip()
                for src in sub["source"].dropna().astype(str)
                if str(src).strip() and str(src).strip() != "leap"
            }
            if not non_leap_sources and not force_show_chart:
                _remove_stale_chart_files(sheet, measure, fuel)
                continue
        display_sheet = sheet_display_labels.get(str(sheet), str(sheet))
        chart_title = display_sheet
        if str(measure).strip() and str(measure).strip().lower() not in display_sheet.lower():
            chart_title = f"{display_sheet} [{measure}]"
        file_sheet = f"{sheet}__{measure}" if str(measure).strip() else str(sheet)
        out = make_chart(sheet, fuel, sub, charts_dir, backend=backend, display_sheet=chart_title, file_sheet=file_sheet)
        if out:
            written.append(out)
    return written


def build_dashboards(
    output_dir: Path,
    comparison_long: pd.DataFrame,
    charts_dir: Path,
    mapping_status: pd.DataFrame | None = None,
    return_sheet_paths: bool = False,
) -> Path | None:
    """
    Reuse the lightweight dashboard style from leap_transport (_build_sheet_dashboards).
    """
    if comparison_long.empty or (not charts_dir.exists() and not return_sheet_paths):
        print("[INFO] No charts available for dashboard rendering.")
        return None

    render_long = _prepare_render_long(comparison_long)
    sheet_display_metadata = _build_sheet_display_metadata()
    sheet_display_labels = {sheet: meta.get("label", sheet) for sheet, meta in sheet_display_metadata.items()}
    sheet_display_labels.update(
        {
            "MAP electricity plants": "Electricity plants (electricity output)",
            "MAP CHP plants (electricity)": "CHP plants (electricity output)",
        }
    )
    for sheet in render_long["sheet"].astype(str).drop_duplicates():
        if not sheet.endswith("_loss_own_use_total"):
            continue
        if sheet_display_labels.get(sheet, sheet) == sheet:
            meta = sheet_display_metadata.get(sheet, {})
            sheet_display_labels[sheet] = _format_loss_own_use_display_label(
                sheet,
                str(meta.get("notes") or ""),
                str(meta.get("sector_name") or ""),
            ) or sheet

    dashboard_page_by_sheet: dict[str, str] = {}
    dashboard_note_override_by_sheet: dict[str, str] = {}
    dashboard_level_by_sheet: dict[str, int] = {}
    final_category_name_by_sheet: dict[str, str] = {}
    try:
        dashboard_sheet_map = read_config_table(DEFAULT_SHEET_MAP)
        for col in ["sheet_name", "dashboard_page", "dashboard_note_override", "dashboard_level"]:
            if col not in dashboard_sheet_map.columns:
                dashboard_sheet_map[col] = ""
            dashboard_sheet_map[col] = dashboard_sheet_map[col].fillna("").astype(str).str.strip()
        dashboard_page_by_sheet = {}
        for row in dashboard_sheet_map.itertuples(index=False):
            sheet_name = str(row.sheet_name).strip()
            if not sheet_name:
                continue
            override_value = str(getattr(row, "dashboard_page", "")).strip()
            if override_value:
                dashboard_page_by_sheet[sheet_name] = override_value
            final_category = str(getattr(row, "final_category_name", "")).strip()
            if final_category:
                final_category_name_by_sheet[sheet_name] = final_category
            level_raw = str(getattr(row, "dashboard_level", "")).strip()
            if level_raw:
                try:
                    level = int(float(level_raw))
                except ValueError:
                    continue
                if level >= 1:
                    dashboard_level_by_sheet[sheet_name] = level
        dashboard_note_override_by_sheet = {
            str(row.sheet_name): str(row.dashboard_note_override)
            for row in dashboard_sheet_map.itertuples(index=False)
            if str(row.sheet_name).strip() and str(row.dashboard_note_override).strip()
        }
    except Exception:
        dashboard_page_by_sheet = {}
        dashboard_note_override_by_sheet = {}

    # Temporary fallback until leap_results_sheet_map.csv can be updated.
    dashboard_page_by_sheet.setdefault("Datacentres", "Commercial and public services")
    dashboard_note_override_by_sheet.setdefault(
        "Datacentres",
        "Datacentres is shown under Commercial and public services here to match the LEAP structure, even though it is a separate sibling sector in the 9th data.",
    )

    dashboards_dir = output_dir / "dashboards"
    dashboards_dir.mkdir(parents=True, exist_ok=True)
    page_shell_css = """
    :root {
      color-scheme: light;
      --page-padding-x: clamp(12px, 1.8vw, 24px);
      --page-padding-y: clamp(14px, 1.8vw, 24px);
      --body-font-size: clamp(15px, 0.22vw + 14px, 18px);
      --title-font-size: clamp(24px, 0.75vw + 18px, 34px);
      --section-title-size: clamp(18px, 0.45vw + 14px, 24px);
      --sticky-top: 0px;
    }
    html { background: #f4f6f8; }
    body {
      font-family: Segoe UI, Arial, sans-serif;
      margin: 0;
      background: #f4f6f8;
      color: #111;
      font-size: var(--body-font-size);
      line-height: 1.45;
      min-width: 320px;
      -webkit-text-size-adjust: 100%;
    }
    h1 { margin: 0; font-size: var(--title-font-size); }
    h2 { font-size: var(--section-title-size); }
    a { color: #0b3d5c; text-decoration: none; }
    a:hover { text-decoration: underline; }
    .page-shell {
      width: 100%;
      max-width: none;
      margin: 0 auto;
      padding: 0 var(--page-padding-x) 32px;
      box-sizing: border-box;
    }
    .page-header {
      position: sticky;
      top: var(--sticky-top);
      z-index: 100;
      margin: 0 0 18px 0;
      padding: var(--page-padding-y) 0 10px 0;
      background: rgba(244, 246, 248, 0.96);
      backdrop-filter: blur(8px);
      border-bottom: 1px solid #d8dee4;
      box-shadow: 0 1px 0 rgba(0, 0, 0, 0.04);
    }
    .page-body {
      max-width: 100%;
    }
    .header-main-row { display:flex;flex-wrap:wrap;justify-content:space-between;align-items:flex-start;gap:10px 16px; }
    .header-side-controls {
      display:flex;flex-wrap:wrap;align-items:center;justify-content:flex-end;
      gap:8px 10px;flex:1 1 460px;
    }
    .header-inline-controls {
      display:flex;align-items:center;gap:8px;justify-content:flex-end;
      flex:0 0 auto;flex-wrap:nowrap;margin-left:auto;
    }
    .header-toggle {
      width:30px; height:30px; border:1px solid #c5ccd3; border-radius:999px;
      background:#fff; color:#0b3d5c; cursor:pointer; font-size:16px; line-height:1;
    }
    .header-toggle:hover { background:#eef2f5; }
    .header-toggle-row {
      display:flex;
      justify-content:flex-end;
      margin-top:8px;
    }
    .page-header.is-collapsed {
      padding-bottom:0;
      background:transparent;
      backdrop-filter:none;
      border-bottom-color:transparent;
      box-shadow:none;
    }
    .page-header.is-collapsed .header-collapsible { display:none; }
    .page-header.is-collapsed .header-toggle-row { margin-top:0; }
    .dashboard-grid {
      display:grid;
      gap:12px;
      grid-template-columns:repeat(4, minmax(0, 1fr));
      align-items:start;
    }
    .dashboard-grid.expand-1 { grid-template-columns:minmax(0, 1fr); }
    .dashboard-grid.expand-2 { grid-template-columns:repeat(2, minmax(0, 1fr)); }
    .dashboard-grid.expand-3 { grid-template-columns:repeat(3, minmax(0, 1fr)); }
    @media (max-width: 1600px) {
      .dashboard-grid { grid-template-columns:repeat(3, minmax(0, 1fr)); }
    }
    @media (max-width: 1200px) {
      .dashboard-grid { grid-template-columns:repeat(2, minmax(0, 1fr)); }
    }
    .jump-nav {
      margin-top:8px;
      padding-top:8px;
      border-top:1px solid #d8dee4;
      display:flex;
      flex-wrap:wrap;
      gap:8px 10px;
      align-items:flex-start;
    }
    .jump-nav-label {
      font-weight:600;
      color:#4b5563;
      font-size:12px;
      white-space:nowrap;
      padding-top:4px;
    }
    .jump-nav-groups {
      display:flex;
      flex-direction:column;
      gap:6px;
      min-width:0;
      flex:1 1 640px;
    }
    .jump-nav-row {
      display:flex;
      flex-wrap:wrap;
      gap:6px 8px;
      align-items:center;
      min-width:0;
    }
    .jump-nav-row[data-level="1"] { padding-left:18px; }
    .jump-nav-row[data-level="2"] { padding-left:36px; }
    .jump-nav-row[data-level="3"] { padding-left:54px; }
    .jump-chip {
      position:relative;
      display:inline-flex;
      align-items:center;
      gap:6px;
      padding:4px 9px;
      border:1px solid #c5ccd3;
      border-radius:999px;
      background:#fff;
      color:#0b3d5c;
      text-decoration:none;
      font-size:12px;
      line-height:1.25;
      box-shadow:0 1px 1px rgba(15, 23, 42, 0.04);
    }
    .jump-chip::before {
      content:"";
      display:block;
      width:8px;
      height:8px;
      border-radius:999px;
      flex:0 0 auto;
      background:#94a3b8;
    }
    .jump-chip[data-kind="structural_total"] {
      background:#f8fafc;
      border-color:#cbd5e1;
      color:#334155;
      font-weight:600;
    }
    .jump-chip[data-kind="structural_total"]::before { background:#64748b; }
    .jump-chip[data-level="1"] {
      background:#eefbf0;
      border-color:#8ad29a;
      color:#14532d;
    }
    .jump-chip[data-level="1"]::before { background:#22c55e; }
    .jump-chip[data-level="2"] {
      background:#eff8ff;
      border-color:#8ecdf6;
      color:#0c4a6e;
    }
    .jump-chip[data-level="2"]::before { background:#38bdf8; }
    .jump-chip[data-level="3"] {
      background:#f8fafc;
      border-color:#cbd5e1;
      color:#334155;
    }
    .jump-chip[data-level="3"]::before { background:#94a3b8; }
    .lazy-chart-frame {
      width: 100%;
      height: clamp(380px, 62vh, 1100px);
      border: 1px solid #d0d7de;
      border-radius: 6px;
      background: #fff;
      display: block;
      box-sizing: border-box;
    }
    @media (max-width: 720px) {
      .page-shell { width: 100%; }
      .header-inline-controls { width: 100%; justify-content:space-between; margin-left:0; }
      .dashboard-grid { grid-template-columns:minmax(0, 1fr); }
    }
    """
    header_toggle_script = """
    (function() {
      const pageHeader = document.getElementById('page-header');
      const headerToggle = document.getElementById('header-toggle');
      if (!pageHeader || !headerToggle) return;
      const storageKey = 'dashboard-header-collapsed';
      const applyCollapsed = (collapsed) => {
        pageHeader.classList.toggle('is-collapsed', collapsed);
        headerToggle.textContent = collapsed ? '▾' : '▴';
        headerToggle.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
        headerToggle.setAttribute('aria-label', collapsed ? 'Expand header' : 'Collapse header');
      };
      let collapsed = false;
      try {
        collapsed = window.localStorage.getItem(storageKey) === 'true';
      } catch (err) {}
      applyCollapsed(collapsed);
      headerToggle.addEventListener('click', () => {
        collapsed = !pageHeader.classList.contains('is-collapsed');
        applyCollapsed(collapsed);
        try {
          window.localStorage.setItem(storageKey, collapsed ? 'true' : 'false');
        } catch (err) {}
      });
    })();
    """

    # Flag sheet/fuel pairs with a meaningful base-year mismatch so they stand out in the UI.
    base_issue_lookup: dict[tuple[str, str, str], dict[str, float | str]] = {}
    base_probe_year = 2022
    base_diag = comparison_long.copy()
    base_diag["value"] = pd.to_numeric(base_diag["value"], errors="coerce")
    base_diag = base_diag[
        (pd.to_numeric(base_diag["year"], errors="coerce") == base_probe_year)
        & (base_diag["source"].isin(["leap", "base", "base_estimated", "base_mixed"]))
    ].copy()
    if not base_diag.empty:
        base_diag["base_compare"] = base_diag["source"].replace({"base_estimated": "base", "base_mixed": "base"})
        base_wide = (
            base_diag.pivot_table(
                index=["sheet", "measure", "fuel_label"],
                columns="base_compare",
                values="value",
                aggfunc="first",
            )
            .reset_index()
        )
        if "leap" not in base_wide.columns:
            base_wide["leap"] = pd.NA
        if "base" not in base_wide.columns:
            base_wide["base"] = pd.NA
        base_wide["abs_gap"] = (pd.to_numeric(base_wide["leap"], errors="coerce") - pd.to_numeric(base_wide["base"], errors="coerce")).abs()
        base_wide["magnitude"] = (
            pd.concat(
                [
                    pd.to_numeric(base_wide["leap"], errors="coerce").abs(),
                    pd.to_numeric(base_wide["base"], errors="coerce").abs(),
                ],
                axis=1,
            )
            .max(axis=1)
        )
        denom = pd.to_numeric(base_wide["leap"], errors="coerce").abs().where(lambda s: s > 1e-9)
        alt = pd.to_numeric(base_wide["base"], errors="coerce").abs().where(lambda s: s > 1e-9)
        denom = denom.fillna(alt).fillna(1.0)
        base_wide["gap_ratio"] = base_wide["abs_gap"] / denom
        magnitude_p90 = float(base_wide["magnitude"].dropna().quantile(0.9)) if base_wide["magnitude"].notna().any() else 0.0
        magnitude_p50 = float(base_wide["magnitude"].dropna().quantile(0.5)) if base_wide["magnitude"].notna().any() else 0.0
        sig = base_wide[
            pd.to_numeric(base_wide["leap"], errors="coerce").notna()
            & pd.to_numeric(base_wide["base"], errors="coerce").notna()
            & (base_wide["gap_ratio"] >= 0.10)
        ].copy()
        for _, row in sig.iterrows():
            pct = float(row["gap_ratio"]) * 100.0
            if pct >= 200.0:
                severity = "Extreme"
            elif pct >= 50.0:
                severity = "High"
            else:
                severity = "Moderate"
            magnitude = float(row["magnitude"]) if pd.notna(row["magnitude"]) else 0.0
            if magnitude_p90 > 0 and magnitude >= magnitude_p90:
                impact = "major"
            elif magnitude_p50 > 0 and magnitude >= magnitude_p50:
                impact = "medium"
            else:
                impact = "minor"
            base_issue_lookup[(str(row["sheet"]), str(row.get("measure", "")), str(row["fuel_label"]))] = {
                "pct": pct,
                "severity": severity,
                "impact": impact,
                "label": f"Base-year gap: {severity}",
            }

    magnitude_lookup = (
        render_long.assign(value_abs=pd.to_numeric(render_long["value"], errors="coerce").abs())
        .groupby(["sheet", "measure", "fuel_label"], dropna=False)["value_abs"]
        .max()
        .to_dict()
    )
    sheet_fuels_lookup = {
        (str(sheet), str(measure)): {
            str(fuel)
            for fuel in group["fuel_label"].dropna().astype(str)
            if str(fuel) != "Total"
        }
        for (sheet, measure), group in render_long.groupby(["sheet", "measure"], dropna=False)
    }
    sheet_fuels_any_lookup = {
        str(sheet): {
            str(fuel).strip()
            for fuel in group["fuel_label"].dropna().astype(str).tolist()
            if str(fuel).strip() and str(fuel).strip().lower() != "total"
        }
        for sheet, group in render_long.groupby(render_long["sheet"].astype(str), dropna=False)
    }
    derived_loss_sheet_fuels: dict[str, set[str]] = {}
    unique_sheets = {str(sheet) for sheet in render_long["sheet"].dropna().astype(str).tolist()}
    for sheet in sorted(unique_sheets):
        if not sheet.endswith("_loss_own_use_total"):
            continue
        prefix = sheet[: -len("_loss_own_use_total")]
        component_sheets = (
            f"{prefix}_inputs",
            f"{prefix}_out_fuel",
            f"{prefix}_out_feed",
        )
        fuel_values: set[str] = set()
        for component in component_sheets:
            if component in unique_sheets:
                component_rows = render_long[render_long["sheet"].astype(str).eq(component)]
                fuel_values.update(
                    str(fuel).strip()
                    for fuel in component_rows["fuel_label"].dropna().astype(str).tolist()
                    if str(fuel).strip() and str(fuel).strip().lower() != "total"
                )
        derived_loss_sheet_fuels[sheet] = fuel_values

    code_df = _safe_read_codebook_sheet(DEFAULT_CODEBOOK, "code_to_name")
    if code_df.empty:
        return pd.DataFrame()
    code_df["9th_label_clean"] = code_df.get("9th_label", "").map(_clean_token)
    code_df["9th_column_clean"] = code_df.get("9th_column", "").map(_clean_token).str.lower()
    code_df["name_clean"] = code_df.get("name", "").map(_clean_token)
    sector_rows = code_df[code_df["9th_column_clean"].isin({"sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"})].copy()

    def _numeric_seq(code: str) -> tuple[int, ...]:
        parts: list[int] = []
        for token in str(code or "").strip().split("_"):
            if token.isdigit():
                parts.append(int(token))
            else:
                break
        return tuple(parts)

    sector_rows["num_seq"] = sector_rows["9th_label_clean"].map(_numeric_seq)
    seq_to_name = {
        tuple(row["num_seq"]): str(row["name_clean"]).strip()
        for _, row in sector_rows.iterrows()
        if tuple(row["num_seq"]) and str(row["name_clean"]).strip()
    }
    name_to_seq = {
        str(row["name_clean"]).strip().lower(): tuple(row["num_seq"])
        for _, row in sector_rows.iterrows()
        if str(row["name_clean"]).strip() and tuple(row["num_seq"])
    }
    all_sector_seqs = sorted({tuple(seq) for seq in sector_rows["num_seq"] if tuple(seq)})
    leaf_sector_seqs = {
        seq
        for seq in all_sector_seqs
        if not any(other != seq and other[: len(seq)] == seq for other in all_sector_seqs)
    }

    status_by_sheet: dict[str, list[str]] = {}
    sector_category_sheets: set[str] = set()
    try:
        dashboard_sheet_map = read_config_table(DEFAULT_SHEET_MAP)
        if "sheet_name" in dashboard_sheet_map.columns and "category_type" in dashboard_sheet_map.columns:
            dashboard_sheet_map["sheet_name"] = dashboard_sheet_map["sheet_name"].fillna("").astype(str).str.strip()
            dashboard_sheet_map["category_type"] = dashboard_sheet_map["category_type"].fillna("").astype(str).str.strip().str.lower()
            sector_category_sheets.update(
                dashboard_sheet_map.loc[
                    dashboard_sheet_map["category_type"].eq("sector"),
                    "sheet_name",
                ]
            )
    except Exception:
        pass
    if mapping_status is None or mapping_status.empty:
        raise RuntimeError(
            "Dashboard rendering requires non-empty mapping_status to build the grouped hierarchy. "
            "Refusing to fall back to raw sheet-name navigation."
        )
    ms = mapping_status.copy()
    ms["sheet"] = ms["sheet"].astype(str)
    if "mapping_source" in ms.columns:
        mapping_source = ms["mapping_source"].fillna("").astype(str).str.strip().str.lower()
    else:
        mapping_source = pd.Series("", index=ms.index, dtype="object")
    if "mapping_note" in ms.columns:
        mapping_note = ms["mapping_note"].fillna("").astype(str)
    else:
        mapping_note = pd.Series("", index=ms.index, dtype="object")
    sector_category_mask = mapping_source.eq("category_sector") | mapping_note.str.contains(
        "category labels treated as sectors",
        case=False,
        regex=False,
    )
    if sector_category_mask.any():
        sector_category_sheets.update(ms.loc[sector_category_mask, "sheet"].dropna().astype(str))
    for sheet, g in ms.groupby("sheet", dropna=False):
        codes: list[str] = []
        for raw in g["sector_code_9th"].dropna().astype(str):
            codes.extend(_split_sector_codes(raw))
        if codes:
            status_by_sheet[str(sheet)] = codes

    def _expand_to_leaf_sector_seqs(codes: list[str]) -> set[tuple[int, ...]]:
        leaves: set[tuple[int, ...]] = set()
        for code in codes:
            seq = _numeric_seq(code)
            if not seq:
                continue
            matched = {leaf for leaf in leaf_sector_seqs if leaf[: len(seq)] == seq}
            if matched:
                leaves.update(matched)
            else:
                leaves.add(seq)
        return leaves

    sheet_leaf_sector_seqs: dict[str, set[tuple[int, ...]]] = {
        sheet: _expand_to_leaf_sector_seqs(codes)
        for sheet, codes in status_by_sheet.items()
    }
    sheet_numeric_seqs: dict[str, set[tuple[int, ...]]] = {
        sheet: {
            seq
            for code in codes
            if (seq := _numeric_seq(code))
        }
        for sheet, codes in status_by_sheet.items()
    }
    exact_sector_sheet_counts: Counter[tuple[int, ...]] = Counter(
        next(iter(seqs))
        for seqs in sheet_numeric_seqs.values()
        if len(seqs) == 1
    )

    def _common_prefix(seqs: list[tuple[int, ...]]) -> tuple[int, ...]:
        if not seqs:
            return ()
        prefix = list(seqs[0])
        for seq in seqs[1:]:
            keep = 0
            for a, b in zip(prefix, seq):
                if a != b:
                    break
                keep += 1
            prefix = prefix[:keep]
            if not prefix:
                break
        return tuple(prefix)

    def _normalize_sheet_path(names: list[str]) -> list[str]:
        if not names:
            return names
        top = str(names[0]).strip().lower()
        top_title = str(names[0]).strip()
        lowered = [str(name).strip().lower() for name in names]

        # Group supply-side leaves under one section.
        if top in {"production", "imports", "exports"}:
            return ["Supply", top_title]
        if top == "transfers":
            return ["Other transformation", top_title]

        # Pull refining out of the generic transformation bucket.
        if top in {"oil refineries", "refining"} or any("refiner" in token or "refining" in token for token in lowered):
            return ["Refining"] + names

        # Build a dedicated power section that also contains heat.
        if any(
            token in {"electricity plants", "chp plants"}
            for token in lowered
        ):
            leaf = str(names[-1]).strip()
            return ["Power", leaf]

        # Feedstock-output mappings use 18.* canonical sectors so they can pull
        # the right comparator rows, but they should still live under Power in
        # the dashboard rather than surfacing a separate top-level electricity
        # output section.
        if top == "electricity output in gwh":
            if len(names) >= 2:
                return ["Power"] + names[1:]
            return ["Power"]

        if any(
            token in {"heat output in pj", "heat plants", "heat plants and chp"}
            for token in lowered
        ):
            leaf = str(names[-1]).strip()
            return ["Power", "Heat", leaf]

        # Rename and retain non-power/non-refining transformation branches.
        if top == "total transformation sector":
            return ["Other transformation"] + names[1:]
        # Transmission & distribution losses should live under Losses & own use,
        # not the generic "Other transformation" page.
        if top == "transmission and distribution losses":
            return ["Losses & own use"] + names[1:]

        if top in {"international aviation bunkers", "international marine bunkers"}:
            return ["Bunkers"] + names
        if len(names) >= 2:
            second = str(names[1]).strip().lower()
            if top in {"other sector", "other sectors"} and second == "buildings":
                # Promote buildings out of the generic "other sector" bucket so
                # they render as their own top-level dashboard category.
                return names[1:]
        return names

    def _sheet_path(sheet: str) -> list[str]:
        normalized_sheet = str(sheet).strip().lower()
        if normalized_sheet == "trans_dist_loss_own_use_total":
            return ["Power", "Transmission and distribution losses"]
        if normalized_sheet.endswith("_loss_own_use_total"):
            display_sheet = sheet_display_labels.get(sheet, sheet)
            return ["Losses & own use", display_sheet]
        map_power_paths = {
            "map electricity plants": ["Power", "Electricity plants", "Electricity plants (electricity output)"],
            "map chp plants (electricity)": ["Power", "CHP plants", "CHP plants (electricity output)"],
            "elecgen_inputs": ["Power", "Electricity plants", "Electricity generation inputs"],
            "elecgen_out_fuel": ["Power", "Electricity plants", "Electricity generation outputs by product"],
            "elecgen_out_feed": ["Power", "Electricity plants", "Electricity generation outputs by feedstock"],
            "heat_inputs": ["Power", "CHP plants", "Transformation heat inputs"],
            "heat_out_fuel": ["Power", "CHP plants", "Transformation heat output by product"],
            "heat_out_feed": ["Power", "CHP plants", "Transformation heat output by feedstock"],
        }
        if normalized_sheet in map_power_paths:
            return map_power_paths[normalized_sheet]
        codes = status_by_sheet.get(sheet, [])
        seqs = [_numeric_seq(code) for code in codes if _numeric_seq(code)]
        if seqs:
            seq = _common_prefix(seqs)
        else:
            seq = name_to_seq.get(str(sheet).strip().lower(), ())
        if not seq:
            names = _normalize_sheet_path([sheet])
        else:
            names = []
            for i in range(1, len(seq) + 1):
                prefix = seq[:i]
                name = seq_to_name.get(prefix)
                if name:
                    names.append(name)
            if not names:
                names = _normalize_sheet_path([sheet])
            else:
                names = _normalize_sheet_path(names)
        meta = sheet_display_metadata.get(sheet, {})
        meta_note_text = str(meta.get("notes") or "").strip().lower()
        meta_measure_text = str(meta.get("measure") or "").strip().lower()
        sheet_fuels_any = sheet_fuels_any_lookup.get(str(sheet), set())
        output_only_fuels = bool(sheet_fuels_any) and sheet_fuels_any.issubset({"Electricity", "Heat"})
        is_component_style_sheet = (
            normalized_sheet.endswith(("_inputs", "_out_feed", "_out_fuel"))
            or any(
                token in (meta_note_text + " " + meta_measure_text)
                for token in [
                    "input",
                    "inputs",
                    "output by",
                    "outputs by",
                    "feedstock",
                    "output fuel",
                ]
            )
            or output_only_fuels
        )
        # Systematic routing rule:
        # Only dedicated loss/own-use sheets should appear under the
        # "Losses & own use" top group. Component transformation sheets
        # (inputs/output-by-*) stay in the transformation hierarchy.
        def _apply_routing_rules(path_names: list[str]) -> list[str]:
            # Centralized routing exceptions. Toggle by commenting out the
            # rule entries below to see the dashboard without a rule.
            def _rule_supply_group(names: list[str]) -> list[str]:
                if not names:
                    return names
                top = str(names[0]).strip().lower()
                top_title = str(names[0]).strip()
                if top in {"production", "imports", "exports"}:
                    return ["Supply", top_title]
                if top == "transfers":
                    return ["Other transformation", top_title]
                return names

            def _rule_refining_group(names: list[str]) -> list[str]:
                if not names:
                    return names
                top = str(names[0]).strip().lower()
                lowered = [str(name).strip().lower() for name in names]
                if top in {"oil refineries", "refining"} or any("refiner" in token or "refining" in token for token in lowered):
                    return ["Refining"] + names
                return names

            def _rule_power_group(names: list[str]) -> list[str]:
                if not names:
                    return names
                lowered = [str(name).strip().lower() for name in names]
                if any(token in {"electricity plants", "chp plants"} for token in lowered):
                    leaf = str(names[-1]).strip()
                    return ["Power", leaf]
                top = str(names[0]).strip().lower()
                if top == "electricity output in gwh":
                    if len(names) >= 2:
                        return ["Power"] + names[1:]
                    return ["Power"]
                if any(token in {"heat output in pj", "heat plants", "heat plants and chp"} for token in lowered):
                    leaf = str(names[-1]).strip()
                    return ["Power", "Heat", leaf]
                return names

            def _rule_transform_bucket(names: list[str]) -> list[str]:
                if not names:
                    return names
                top = str(names[0]).strip().lower()
                if top == "total transformation sector":
                    return ["Other transformation"] + names[1:]
                if top == "transmission and distribution losses":
                    return ["Losses & own use"] + names[1:]
                return names

            def _rule_bunkers(names: list[str]) -> list[str]:
                if not names:
                    return names
                top = str(names[0]).strip().lower()
                if top in {"international aviation bunkers", "international marine bunkers"}:
                    return ["Bunkers"] + names
                return names

            def _rule_buildings(names: list[str]) -> list[str]:
                if len(names) >= 2:
                    top = str(names[0]).strip().lower()
                    second = str(names[1]).strip().lower()
                    if top in {"other sector", "other sectors"} and second == "buildings":
                        return names[1:]
                return names

            def _rule_loss_own_use_component(names: list[str]) -> list[str]:
                if (
                    names
                    and str(names[0]).strip().lower() == "losses & own use"
                    and not normalized_sheet.endswith("_loss_own_use_total")
                    and is_component_style_sheet
                    and not normalized_sheet.startswith("trans_dist_")
                ):
                    return ["Other transformation"] + list(names[1:])
                return names

            routing_rules = [
                ("supply_group", _rule_supply_group),
                ("refining_group", _rule_refining_group),
                ("power_group", _rule_power_group),
                ("transform_bucket", _rule_transform_bucket),
                ("bunkers", _rule_bunkers),
                ("buildings", _rule_buildings),
                ("loss_own_use_component", _rule_loss_own_use_component),
            ]
            for _, rule in routing_rules:
                path_names = rule(path_names)
            return path_names

        names = _apply_routing_rules(list(names))
        display_sheet = sheet_display_labels.get(sheet, sheet)
        forced_display_leaf = False
        skip_raw_sheet = False
        parent_override = str(dashboard_page_by_sheet.get(str(sheet), "")).strip()
        if parent_override and names:
            top_name = str(names[0]).strip()
            top_token = _clean_token(top_name).lower()
            override_token = _clean_token(parent_override).lower()
            if top_token and override_token.endswith(" sector"):
                equivalent_tokens = {top_token}
                if top_token.endswith(" sector"):
                    equivalent_tokens.add(top_token[: -len(" sector")].strip())
                else:
                    equivalent_tokens.add(f"{top_token} sector")
                if override_token in equivalent_tokens:
                    parent_override = ""
        if parent_override:
            top = str(names[0]).strip() if names else ""
            if top:
                if parent_override == top:
                    pass
                elif display_sheet == parent_override:
                    names = [top, parent_override]
                    forced_display_leaf = True
                else:
                    names = [top, parent_override, display_sheet]
                    forced_display_leaf = True
        level_override = dashboard_level_by_sheet.get(str(sheet))
        if level_override is not None and level_override >= 1:
            desired_len = level_override + 1
            if len(names) >= desired_len:
                names = names[:desired_len]
                names[-1] = display_sheet
                forced_display_leaf = True
                skip_raw_sheet = True
        exact_sheet_seqs = sheet_numeric_seqs.get(sheet, set())
        preserve_supply_flow_branch = (
            len(names) >= 2
            and str(names[0]).strip().lower() == "supply"
            and str(names[-1]).strip().lower() in {"production", "imports", "exports"}
        )
        replace_leaf_with_display = (
            names[-1] != display_sheet
            and len(exact_sheet_seqs) == 1
            and exact_sector_sheet_counts.get(next(iter(exact_sheet_seqs)), 0) == 1
            and not preserve_supply_flow_branch
            and not forced_display_leaf
        )
        if replace_leaf_with_display:
            names[-1] = display_sheet
        elif not forced_display_leaf and names[-1] != display_sheet:
            names.append(display_sheet)
        raw_sheet = str(sheet).strip()
        if raw_sheet and raw_sheet != names[-1] and not skip_raw_sheet:
            names.append(raw_sheet)
        return names

    sheet_paths: dict[str, list[str]] = {}
    for sheet in render_long["sheet"].astype(str).drop_duplicates():
        sheet_paths[sheet] = _sheet_path(sheet)
    if return_sheet_paths:
        return sheet_paths

    def _sheet_sort_key(sheet: str) -> tuple:
        path = sheet_paths.get(sheet, [sheet])
        return tuple([p.lower() for p in path] + [sheet.lower()])

    top_group_display_labels = {
        "Buildings": "Buildings",
        "Bunkers": "Bunkers",
        "Industry sector": "Industry",
        "Transport sector": "Transport",
        "Other sector": "Other",
        "Power": "Power",
        "Refining": "Refining",
        "Other transformation": "Other transformation",
        "Losses & own use": "Losses & own use",
        "Supply": "Supply",
    }
    nav_group_blocks = [
        ("Demand", ["Buildings", "Bunkers", "Industry sector", "Transport sector", "Other sector"]),
        ("Transformation", ["Power", "Refining", "Other transformation", "Losses & own use"]),
        ("Supply", ["Supply"]),
    ]
    demand_top_groups = {"Buildings", "Bunkers", "Industry sector", "Other sector", "Transport sector"}
    supply_top_groups = {"Supply"}
    transformation_top_groups = {"Power", "Other transformation", "Refining", "Losses & own use"}
    supply_overview_measure = "Supply overview (PJ)"
    supply_tpes_measure = "TPES excl. bunkers (PJ)"
    demand_measure_label = "Final Energy Demand (PJ)"
    supply_overview_note = (
        "TPES is calculated here as production + imports - exports. "
        "International aviation and marine bunkers are excluded. "
        "Stock changes are not available in the current LEAP supply export. "
        "The other charts show total production, imports, and exports."
    )

    def _normalize_measure_label(text: object, *, default_units: str = "PJ") -> str:
        cleaned = _clean_token(text)
        if not cleaned:
            return ""
        lowered = cleaned.lower()
        if any(unit in lowered for unit in ["(pj", "(gwh", "(mw", "(gw", " petajoule", " gwh", " mw", " gw"]):
            return cleaned
        return f"{cleaned} ({default_units})"

    def _display_top_group(name: object) -> str:
        text = str(name or "").strip()
        return top_group_display_labels.get(text, text)

    def _infer_sheet_measure(sheet: str, measure: str = "") -> str:
        path = sheet_paths.get(sheet, [sheet])
        top_group = path[0] if path else ""
        measure_text = str(measure).strip()
        compact_measure = re.sub(r"\s*\([^)]*\)\s*", "", measure_text).strip().lower()
        if measure_text and measure_text.lower() != "nan":
            if top_group in demand_top_groups and compact_measure in {"energy", "demand", "final energy demand"}:
                return demand_measure_label
            return measure_text
        meta = sheet_display_metadata.get(sheet, {})
        notes = str(meta.get("notes") or "").strip()
        category_type = str(meta.get("category_type") or "").strip().lower()
        fuels = sheet_fuels_lookup.get((sheet, str(measure).strip()), set())
        output_fuels = {"Electricity", "Heat"}

        if notes:
            note_label = _normalize_measure_label(notes)
            if sheet.endswith("_loss_own_use_total"):
                loss_fuels = sorted(derived_loss_sheet_fuels.get(sheet, set()))
                if loss_fuels:
                    preview = ", ".join(loss_fuels[:8])
                    if len(loss_fuels) > 8:
                        preview = f"{preview}, +{len(loss_fuels) - 8} more"
                    return f"{note_label}; fuels: {preview}"
            return note_label
        if category_type == "sector" or top_group in demand_top_groups:
            return demand_measure_label
        if top_group in supply_top_groups:
            return "Supply flow (PJ)"
        if top_group == "Losses & own use":
            return "Losses & own use (PJ)"
        if top_group in transformation_top_groups:
            if fuels and fuels.issubset(output_fuels):
                return "Output by product (PJ)"
            if fuels & output_fuels:
                return "Transformation inputs and outputs (PJ)"
            if top_group == "Power":
                return "Power transformation (PJ)"
            return "Transformation inputs and outputs (PJ)"
        return "Energy (PJ)"

    def _infer_page_measure(path: tuple[str, ...], own_sheet: str | None) -> str:
        top_group = path[0] if path else ""
        if own_sheet:
            return _infer_sheet_measure(own_sheet)
        if top_group in demand_top_groups:
            return demand_measure_label
        if top_group in supply_top_groups:
            if len(path) == 1:
                return supply_overview_measure
            supply_branch = str(path[1] if len(path) > 1 else "").strip().lower()
            if supply_branch == "production":
                return "Production (PJ)"
            if supply_branch == "imports":
                return "Imports (PJ)"
            if supply_branch == "exports":
                return "Exports (PJ)"
            if supply_branch == "transfers":
                return "Transfers (PJ)"
            return "Supply flow summary (PJ)"
        if top_group == "Losses & own use":
            return "Losses & own use summary (PJ)"
        if top_group == "Power":
            return "Power totals by measure (PJ)"
        if top_group == "Other transformation":
            return "Transformation totals by measure (PJ)"
        return "Energy summary (PJ)"

    render_long["measure"] = render_long.apply(
        lambda row: _infer_sheet_measure(str(row.get("sheet", "")), str(row.get("measure", ""))),
        axis=1,
    )

    INPUTS_ONLY_LABEL = "Inputs only (PJ)"
    INPUTS_PLUS_LOSSES_LABEL = "Inputs + losses & own use (PJ)"

    def _structural_total_measure_bucket(measure: str, path: tuple[str, ...]) -> str:
        measure_text = _normalize_measure_label(measure).strip()
        compact = re.sub(r"\s*\([^)]*\)\s*", "", measure_text).strip().lower()
        top_group = path[0] if path else ""
        if top_group in demand_top_groups and compact in {"energy", "demand", "final energy demand"}:
            return demand_measure_label
        if top_group in transformation_top_groups:
            if "input" in compact and "output" not in compact:
                if "losses" in compact and "own use" in compact:
                    return INPUTS_PLUS_LOSSES_LABEL
                return INPUTS_ONLY_LABEL
            if "output" in compact and "feedstock" in compact:
                return "Outputs by feedstock (PJ)"
            if "output" in compact and ("product" in compact or "fuel" in compact):
                return "Outputs by product (PJ)"
        return measure_text

    def _with_measure_badge(label: str, measure: str, *, linked_file: str | None = None) -> str:
        base_label = f'<a href="{linked_file}">{label}</a>' if linked_file else label
        if not measure:
            return base_label
        badge = (
            '<span style="display:inline-block;margin-left:8px;padding:2px 8px;'
            'border:1px solid #c5ccd3;border-radius:999px;background:#fff;'
            'color:#4b5563;font-size:12px;font-weight:500;vertical-align:middle;">'
            f'{measure}</span>'
        )
        return f"{base_label}{badge}"

    def _compact_jump_label(label: str, measure: str = "") -> str:
        measure_text = str(measure or "").strip()
        if not measure_text:
            return label
        compact = re.sub(r"\s*\([^)]*\)\s*", "", measure_text).strip().lower()
        compact = compact.replace("aggregate of displayed ", "").replace(" charts", "")
        replacements = {
            "power": "summary",
            "transformation inputs and outputs": "transformation",
            "supply flow": "supply",
            "supply flow summary": "supply",
            "demand": "demand",
            "losses / own use": "losses",
            "losses / own use summary": "losses",
            "losses & own use": "losses",
            "losses & own use summary": "losses",
            "electricity generation inputs": "inputs",
            "electricity generation outputs by feedstock": "feedstock",
            "electricity generation outputs by product": "product",
            "heat generation inputs": "inputs",
            "heat generation outputs by feedstock": "feedstock",
            "heat generation outputs by product": "product",
            "chp inputs": "inputs",
            "chp outputs by feedstock": "feedstock",
            "chp outputs by product": "product",
            "output by product": "output",
            "output by feedstock": "feedstock",
        }
        compact = replacements.get(compact, compact)
        compact = compact.strip(" -:")
        if not compact or compact == label.lower():
            return label
        return f"{label}: {compact}"

    def _compact_section_jump_label(label: str) -> str:
        compact = str(label or "").strip()
        if not compact:
            return compact
        replacements = {
            "Electricity generation inputs": "Electricity inputs",
            "Electricity generation outputs by feedstock": "Electricity feedstock",
            "Electricity generation outputs by product": "Electricity product",
            "Heat generation inputs": "Heat inputs",
            "Heat generation outputs by feedstock": "Heat feedstock",
            "Heat generation outputs by product": "Heat product",
            "Transformation heat inputs": "Heat inputs",
            "Transformation heat output by feedstock": "Heat feedstock",
            "Transformation heat output by product": "Heat product",
            "Natural gas blending inputs": "Gas blend inputs",
            "Natural gas blending outputs by feedstock": "Gas blend feedstock",
            "Natural gas liquefaction inputs": "Gas liq inputs",
            "Gas processing outputs by feedstock": "Gas processing feedstock",
            "Non-specified transformation inputs": "Other trans. inputs",
            "Non-specified transformation outputs by feedstock": "Other trans. feedstock",
        }
        return replacements.get(compact, compact)

    def _compact_measure_name(measure: str) -> str:
        compact = re.sub(r"\s*\([^)]*\)\s*", "", str(measure or "").strip()).strip()
        replacements = {
            "Aggregate of displayed power charts": "Summary",
            "Aggregate of displayed transformation charts": "Summary",
            "TPES excl. bunkers": "TPES",
            "Summary": "Summary",
            "Inputs": "Inputs",
            "Outputs by feedstock": "Feedstock outputs",
            "Outputs by product": "Product outputs",
            "Electricity generation inputs": "Inputs",
            "Electricity generation outputs by feedstock": "Feedstock",
            "Electricity generation outputs by product": "Product",
            "Transformation inputs and outputs": "Transformation",
            "Supply flow summary": "Supply",
            "Losses / own use summary": "Losses",
            "Losses & own use summary": "Losses",
        }
        return replacements.get(compact, compact)

    hidden_measure_sections = {
        "transformation inputs and outputs",
    }
    explicit_input_output_measure_buckets = {
        INPUTS_ONLY_LABEL,
        INPUTS_PLUS_LOSSES_LABEL,
        "Outputs by feedstock (PJ)",
        "Outputs by product (PJ)",
    }

    def _is_hidden_measure_section(measure: object) -> bool:
        token = re.sub(r"\s*\([^)]*\)\s*", "", str(measure or "").strip()).strip().lower()
        return token in hidden_measure_sections

    def _filter_hidden_measure_entries(entries: list[tuple[str, str, Path]]) -> list[tuple[str, str, Path]]:
        return [entry for entry in entries if not _is_hidden_measure_section(entry[0])]

    def _path_depth_for_jump(path: tuple[str, ...] | None) -> int:
        if not path:
            return 0
        tokens = [str(part).strip() for part in path if str(part).strip()]
        if len(tokens) >= 2:
            def _token_key(value: str) -> str:
                return re.sub(r"[^a-z0-9]+", "", _clean_token(value).lower())

            while len(tokens) >= 2 and _token_key(tokens[-1]) == _token_key(tokens[-2]):
                tokens.pop()
        return len(tokens)

    def _section_jump_level(current_path: tuple[str, ...], target_path: tuple[str, ...] | None) -> int:
        if not target_path:
            return 0
        current_depth = _path_depth_for_jump(current_path)
        target_depth = _path_depth_for_jump(target_path)
        if target_depth <= current_depth:
            return 1
        return max(1, min(3, target_depth - current_depth))

    def _looks_measure_like_label(label: object) -> bool:
        text = str(label or "").strip().lower()
        if not text:
            return False
        measure_tokens = [
            "input",
            "inputs",
            "output",
            "outputs",
            "feedstock",
            "product",
            "losses / own use",
            "losses & own use",
            "losses and own use",
            "loss / own use",
            "derived",
            "total",
            "totals",
        ]
        return any(token in text for token in measure_tokens)

    def _looks_raw_sheet_token(label: object) -> bool:
        text = str(label or "").strip()
        if not text:
            return False
        # Raw sheet ids are typically snake_case tokens without spaces.
        if " " in text:
            return False
        if "_" not in text:
            return False
        lowered = text.lower()
        return lowered == text

    def _sector_label_from_target_path(current_path: tuple[str, ...], target_path: tuple[str, ...] | None) -> str:
        if not target_path:
            return ""
        relative = list(target_path[len(current_path):])
        if not relative:
            return ""
        non_measure = [
            token
            for token in relative
            if token and not _looks_measure_like_label(token) and not _looks_raw_sheet_token(token)
        ]
        if non_measure:
            return str(non_measure[-1]).strip()
        return str(relative[0]).strip()

    def _coalesce_sector_jump_links_for_complex_pages(
        section_jump_links: list[dict[str, object]],
        current_path: tuple[str, ...],
    ) -> list[dict[str, object]]:
        if not section_jump_links:
            return section_jump_links
        if not current_path or current_path[0] not in {"Other transformation", "Power", "Losses & own use"}:
            return section_jump_links

        out: list[dict[str, object]] = []
        seen_sector: set[str] = set()
        for item in section_jump_links:
            kind = str(item.get("kind") or "")
            if kind == "structural_total":
                out.append(item)
                continue
            sector_label = _sector_label_from_target_path(
                current_path,
                item.get("target_path"),  # type: ignore[arg-type]
            )
            if not sector_label:
                continue
            sector_key = sector_label.strip().lower()
            if not sector_key or sector_key in seen_sector:
                continue
            seen_sector.add(sector_key)
            out.append(
                {
                    "label": sector_label,
                    "section_id": item.get("section_id"),
                    "level": 1,
                    "kind": "sector",
                    "target_path": item.get("target_path"),
                }
            )
        return out

    def _inject_intermediate_parent_jump_links(
        section_jump_links: list[dict[str, object]],
        current_path: tuple[str, ...],
    ) -> list[dict[str, object]]:
        if not section_jump_links or not current_path:
            return section_jump_links

        current_tokens = [str(token).strip() for token in current_path if str(token).strip()]
        if not current_tokens:
            return section_jump_links
        current_depth = len(current_tokens)
        top_token = _clean_token(current_tokens[0]).lower()
        redundant_top_tokens = {top_token, f"{top_token} sector"}

        existing_labels = {
            _clean_token(str(item.get("label", ""))).lower()
            for item in section_jump_links
            if _clean_token(str(item.get("label", ""))).strip()
        }
        parent_links: dict[str, dict[str, object]] = {}
        for item in section_jump_links:
            target_path_raw = item.get("target_path")
            if not isinstance(target_path_raw, tuple):
                continue
            target_tokens = [str(token).strip() for token in target_path_raw if str(token).strip()]
            if len(target_tokens) <= current_depth + 1:
                continue
            parent_label = target_tokens[current_depth]
            parent_key = _clean_token(parent_label).lower()
            if not parent_key or parent_key in redundant_top_tokens:
                continue
            if parent_key in existing_labels or parent_key in parent_links:
                continue
            parent_links[parent_key] = {
                "label": parent_label,
                "section_id": item.get("section_id"),
                "level": 1,
                "kind": "parent",
                "target_path": tuple(target_tokens[: current_depth + 1]),
            }

        if not parent_links:
            return section_jump_links
        return [*section_jump_links, *parent_links.values()]

    def _drop_redundant_top_sector_jump_links(
        section_jump_links: list[dict[str, object]],
        current_path: tuple[str, ...],
    ) -> list[dict[str, object]]:
        if not section_jump_links or not current_path:
            return section_jump_links
        top_token = _clean_token(current_path[0]).lower()
        redundant_tokens = {top_token, f"{top_token} sector"}
        out: list[dict[str, object]] = []
        for item in section_jump_links:
            label_token = _clean_token(str(item.get("label", ""))).lower()
            if label_token in redundant_tokens and str(item.get("kind", "")) != "structural_total":
                continue
            out.append(item)
        return out

    def _prepare_section_jump_links(
        section_jump_links: list[dict[str, object]],
        current_path: tuple[str, ...],
    ) -> list[dict[str, object]]:
        links = _coalesce_sector_jump_links_for_complex_pages(section_jump_links, current_path)
        links = _inject_intermediate_parent_jump_links(links, current_path)
        links = _drop_redundant_top_sector_jump_links(links, current_path)
        return links

    def _render_section_jump_row(section_jump_links: list[dict[str, object]]) -> str:
        if not section_jump_links:
            return ""
        grouped_links: dict[int, list[dict[str, object]]] = {}
        for item in section_jump_links:
            raw_level = item.get("level", 1)
            level = 1 if raw_level is None else int(raw_level)
            grouped_links.setdefault(level, []).append(item)
        group_html = "".join(
            '<div class="jump-nav-row" data-level="{level}">{chips}</div>'.format(
                level=level,
                chips="".join(
                    f'<a href="#{entry["section_id"]}" class="jump-chip" '
                    f'data-level="{level}" data-kind="{entry["kind"]}">{entry["label"]}</a>'
                    for entry in grouped_links[level]
                ),
            )
            for level in sorted(grouped_links)
        )
        return (
            '<div class="jump-nav">'
            '<span class="jump-nav-label">Jump to:</span>'
            f'<div class="jump-nav-groups">{group_html}</div>'
            '</div>'
        )

    def _supply_component_sign(sheet: str) -> int:
        path = tuple(sheet_paths.get(sheet, [sheet]))
        if not path or path[0] != "Supply":
            return 0
        branch = str(path[1] if len(path) > 1 else "").strip().lower()
        if branch in {"production", "imports"}:
            return 1
        if branch == "exports":
            return -1
        return 0

    def _build_supply_tpes_total(node_total_subset: pd.DataFrame, title: str) -> pd.DataFrame:
        signed_subset = node_total_subset.copy()
        signed_subset["supply_component_sign"] = signed_subset["sheet"].map(_supply_component_sign)
        signed_subset = signed_subset[signed_subset["supply_component_sign"] != 0].copy()
        if signed_subset.empty:
            return pd.DataFrame()
        signed_subset["value"] = (
            pd.to_numeric(signed_subset["value"], errors="coerce")
            * pd.to_numeric(signed_subset["supply_component_sign"], errors="coerce")
        )
        return _aggregate_display_rows_to_total(
            signed_subset,
            title=title,
            measure_value=supply_tpes_measure,
            collapse_base_family=True,
            collapse_projection_family=True,
        )

    def _build_supply_overview_component_total(section_sheet: str) -> pd.DataFrame:
        section_total = render_long[
            (render_long["sheet"].astype(str) == str(section_sheet).strip())
            & (render_long["fuel_label"].astype(str) == "Total")
        ].copy()
        if _has_nonzero_chart_values(section_total):
            return section_total

        section_rows = render_long[
            (render_long["sheet"].astype(str) == str(section_sheet).strip())
            & (render_long["fuel_label"].astype(str) != "Total")
        ].copy()
        if section_rows.empty:
            return pd.DataFrame()

        measure_value = _infer_page_measure(("Supply", str(section_sheet).strip()), str(section_sheet).strip())
        return _aggregate_display_rows_to_total(
            section_rows,
            title=str(section_sheet).strip(),
            measure_value=measure_value,
            collapse_base_family=True,
            collapse_projection_family=True,
        )

    def _find_total_chart_path(section_sheet: str) -> Path | None:
        for _, fuel, chart_path in sheet_entries_lookup.get(section_sheet, []):
            if str(fuel) == "Total":
                return chart_path
        return None

    def _render_measure_total_cards(section_sheet: str, entries: list[tuple[str, str, Path]]) -> str:
        entries = _filter_hidden_measure_entries(entries)
        if not entries:
            return ""
        def _grid_expand_class(item_count: int) -> str:
            if item_count == 1:
                return " expand-1"
            if item_count == 2:
                return " expand-2"
            if item_count == 3:
                return " expand-3"
            return ""
        cards = []
        for measure, fuel, png_path in entries:
            issue = base_issue_lookup.get((section_sheet, str(measure).strip(), fuel))
            rel_chart = os.path.relpath(png_path, start=dashboards_dir).replace("\\", "/")
            if png_path.suffix.lower() == ".html":
                chart_markup = (
                    f'<iframe data-src="{rel_chart}" '
                    f'title="{section_sheet} – {measure}" '
                    'class="lazy-chart-frame" '
                    'loading="lazy"></iframe>'
                )
            else:
                chart_markup = (
                    f'<img data-src="{rel_chart}" alt="{section_sheet} – {measure}" '
                    'class="lazy-chart-image" '
                    'style="max-width:100%;height:auto;display:block;min-height:160px;background:#f8fafc;" '
                    'loading="lazy" />'
                )
            card_style = "margin:8px;padding:8px;border:1px solid #d0d7de;border-radius:8px;background:#fff;box-shadow:0 1px 2px rgba(0,0,0,0.05);"
            issue_badge = ""
            if issue:
                severity = str(issue.get("severity", "Moderate"))
                impact = str(issue.get("impact", "minor"))
                palette = {
                    "Moderate": {"border": "rgba(217,119,6,{a})", "bg": "rgba(245,158,11,{b})", "text": "#92400e"},
                    "High": {"border": "rgba(234,88,12,{a})", "bg": "rgba(249,115,22,{b})", "text": "#9a3412"},
                    "Extreme": {"border": "rgba(220,38,38,{a})", "bg": "rgba(220,38,38,{b})", "text": "#991b1b"},
                }.get(severity, {"border": "rgba(217,119,6,{a})", "bg": "rgba(245,158,11,{b})", "text": "#92400e"})
                alpha = {"minor": "0.30", "medium": "0.42", "major": "0.58"}.get(impact, "0.30")
                bg_alpha = {"minor": "0.04", "medium": "0.07", "major": "0.11"}.get(impact, "0.04")
                card_style = (
                    "margin:8px;padding:8px;border:2px solid "
                    + palette["border"].format(a=alpha)
                    + ";border-radius:8px;background:"
                    + palette["bg"].format(b=bg_alpha)
                    + ";box-shadow:0 1px 2px rgba(0,0,0,0.05);"
                )
                issue_badge = (
                    f'<div style="margin-top:4px;color:{palette["text"]};font-size:12px;font-weight:600;">'
                    f'{issue["label"]}</div>'
                )
            cards.append(
                f"""
<figure style="{card_style}">
  {"<div style=\"margin-bottom:4px;color:#4b5563;font-size:12px;font-weight:600;\">" + str(measure) + "</div>" if str(measure).strip() else ""}
  {issue_badge}
  {chart_markup}
</figure>
"""
            )
        return (
            f'<div class="dashboard-grid{_grid_expand_class(len(cards))}">'
            + "".join(cards)
            + "</div>"
        )

    def _has_nonzero_chart_values(frame: pd.DataFrame) -> bool:
        if frame.empty or "value" not in frame.columns:
            return False
        values = pd.to_numeric(frame["value"], errors="coerce").dropna()
        if values.empty:
            return False
        return bool((values.abs() > 1e-12).any())

    def _menu_label(sheet: str) -> str:
        path = sheet_paths.get(sheet, [sheet])
        depth = max(0, len(path) - 1)
        return f"{'  ' * depth}{sheet}"

    sheet_to_entries: dict[str, list[tuple[str, str, Path]]] = {}
    chart_files: dict[str, Path] = {}
    for p in charts_dir.glob("*"):
        if p.suffix.lower() not in {".png", ".svg", ".html"}:
            continue
        stem = p.stem
        existing = chart_files.get(stem)
        if existing is None or p.stat().st_mtime > existing.stat().st_mtime:
            chart_files[stem] = p

    def _resolve_chart_file(sheet: object, measure: object, fuel: object) -> Path | None:
        sheet_text = str(sheet or "")
        measure_text = str(measure or "").strip()
        fuel_slug = _safe_token(fuel)
        sheet_slug = _safe_token((f"{sheet_text}__{measure_text}" if measure_text else sheet_text).replace("\\", "_"))
        exact_stem = f"{sheet_slug}__{fuel_slug}"
        if exact_stem in chart_files:
            return chart_files[exact_stem]

        raw_sheet_slug = _safe_token(sheet_text.replace("\\", "_"))
        blank_stem = f"{raw_sheet_slug}__{fuel_slug}"
        candidates: list[tuple[str, Path]] = []
        if blank_stem in chart_files:
            candidates.append((blank_stem, chart_files[blank_stem]))

        suffix = f"__{fuel_slug}"
        prefix = f"{raw_sheet_slug}__"
        for stem, path in chart_files.items():
            if stem == exact_stem or stem == blank_stem:
                continue
            if stem.startswith(prefix) and stem.endswith(suffix):
                candidates.append((stem, path))

        if not candidates:
            return None

        measure_slug = _safe_token(measure_text) if measure_text else ""
        if measure_slug:
            for stem, path in candidates:
                if f"__{measure_slug}__" in stem:
                    return path

        candidates.sort(key=lambda item: (item[0].count("__"), len(item[0])), reverse=True)
        return candidates[0][1]

    for (sheet, measure, fuel), sub in render_long.groupby(["sheet", "measure", "fuel_label"]):
        if str(sheet) in sector_category_sheets:
            continue
        chart_path = _resolve_chart_file(sheet, measure, fuel)
        if chart_path is None:
            values = pd.to_numeric(sub["value"], errors="coerce").fillna(0.0)
            if values.ne(0).any():
                display_sheet = sheet_display_labels.get(str(sheet), str(sheet))
                chart_title = display_sheet
                if str(measure).strip() and str(measure).strip().lower() not in display_sheet.lower():
                    chart_title = f"{display_sheet} [{measure}]"
                file_sheet = f"{sheet}__{measure}" if str(measure).strip() else str(sheet)
                chart_path = make_chart(
                    str(sheet),
                    str(fuel),
                    sub.copy(),
                    charts_dir,
                    backend="plotly",
                    display_sheet=chart_title,
                    file_sheet=file_sheet,
                )
                if chart_path is not None:
                    chart_files[chart_path.stem] = chart_path
        if chart_path is not None:
            sheet_to_entries.setdefault(sheet, []).append((str(measure).strip(), fuel, chart_path))

    if not sheet_to_entries:
        print("[INFO] No matching chart files were found for dashboards.")
        return None

    ordered_sheets = sorted(sheet_to_entries.items(), key=lambda item: _sheet_sort_key(item[0]))
    sheet_entries_lookup = {sheet: entries for sheet, entries in ordered_sheets}
    sheet_path_tuples = {sheet: tuple(sheet_paths.get(sheet, [sheet])) for sheet in sheet_to_entries}
    sheet_abs_sum_lookup: dict[str, float] = (
        render_long.assign(value_abs=pd.to_numeric(render_long["value"], errors="coerce").abs().fillna(0.0))
        .groupby(render_long["sheet"].astype(str), dropna=False)["value_abs"]
        .sum()
        .to_dict()
    )
    comparator_sources = {"base", "base_estimated", "base_mixed", "projection", "projection_estimated", "projection_mixed"}
    sheet_source_lookup = {
        str(sheet): set(group["source"].fillna("").astype(str).str.strip())
        for sheet, group in render_long.groupby(render_long["sheet"].astype(str), dropna=False)
    }

    node_paths: set[tuple[str, ...]] = set()
    for path in sheet_path_tuples.values():
        for i in range(1, len(path) + 1):
            node_paths.add(path[:i])

    node_to_sheet: dict[tuple[str, ...], str] = {path: sheet for sheet, path in sheet_path_tuples.items()}
    node_abs_sum_lookup: dict[tuple[str, ...], float] = defaultdict(float)
    for sheet, path in sheet_path_tuples.items():
        sheet_abs = float(sheet_abs_sum_lookup.get(str(sheet), 0.0) or 0.0)
        if sheet_abs == 0.0 or not path:
            continue
        for i in range(1, len(path) + 1):
            node_abs_sum_lookup[path[:i]] += sheet_abs

    children_by_parent: dict[tuple[str, ...], list[tuple[str, ...]]] = defaultdict(list)
    for path in node_paths:
        children_by_parent[path[:-1]].append(path)

    def _path_display_name(path: tuple[str, ...]) -> str:
        if not path:
            return ""
        if len(path) == 1:
            return _display_top_group(path[0])
        return str(path[-1]).strip()

    def _ordered_children(parent: tuple[str, ...]) -> list[tuple[str, ...]]:
        return sorted(
            children_by_parent.get(parent, []),
            key=lambda child: (
                -float(node_abs_sum_lookup.get(child, 0.0) or 0.0),
                _path_display_name(child).lower(),
                tuple(part.lower() for part in child),
            ),
        )

    ordered_node_paths: list[tuple[str, ...]] = []

    def _append_ordered_descendants(parent: tuple[str, ...]) -> None:
        for child in _ordered_children(parent):
            ordered_node_paths.append(child)
            _append_ordered_descendants(child)

    _append_ordered_descendants(())
    path_order_index = {path: idx for idx, path in enumerate(ordered_node_paths)}
    fallback_path_rank = len(path_order_index) + 1

    def _dashboard_filename(path: tuple[str, ...]) -> str:
        if path in node_to_sheet:
            sheet = node_to_sheet[path]
            return f"{_safe_token(sheet.replace('\\', '_'))}.html"
        slug = "__".join(_safe_token(part.replace("\\", "_")) for part in path)
        return f"node__{slug}.html"

    path_to_file = {path: dashboards_dir / _dashboard_filename(path) for path in ordered_node_paths}
    expected_dashboard_files = set(path_to_file.values()) | {dashboards_dir / "index.html"}
    for existing_html in dashboards_dir.glob("*.html"):
        if existing_html not in expected_dashboard_files:
            try:
                existing_html.unlink()
            except OSError:
                pass

    def _descendant_sheets(path: tuple[str, ...]) -> list[str]:
        out = [sheet for sheet, sheet_path in sheet_path_tuples.items() if sheet_path[: len(path)] == path]
        return sorted(
            out,
            key=lambda sheet: (
                path_order_index.get(tuple(sheet_path_tuples.get(sheet, (sheet,))), fallback_path_rank),
                -float(sheet_abs_sum_lookup.get(str(sheet), 0.0) or 0.0),
                _sheet_sort_key(sheet),
            ),
        )

    def _leaf_descendant_sheets(path: tuple[str, ...]) -> list[str]:
        descendants = _descendant_sheets(path)
        leafs: list[str] = []
        for sheet in descendants:
            sheet_path = sheet_path_tuples[sheet]
            has_child = any(
                other != sheet and sheet_path == sheet_path_tuples[other][: len(sheet_path)]
                for other in descendants
            )
            if not has_child:
                leafs.append(sheet)
        return sorted(leafs, key=_sheet_sort_key)

    def _direct_child_descendant_sheets(path: tuple[str, ...]) -> list[str]:
        descendants = _descendant_sheets(path)
        if not descendants:
            return []
        explicit_measure_roles = set(explicit_input_output_measure_buckets)

        def _sheet_measure_role(sheet: str) -> str:
            return _structural_total_measure_bucket(
                _infer_sheet_measure(sheet),
                tuple(sheet_paths.get(sheet, [sheet])),
            )

        by_child_segment: dict[str, list[str]] = {}
        path_len = len(path)
        for sheet in descendants:
            sheet_path = tuple(sheet_path_tuples.get(sheet, ()))
            if len(sheet_path) <= path_len:
                continue
            child_segment = str(sheet_path[path_len]).strip()
            if not child_segment:
                continue
            by_child_segment.setdefault(child_segment, []).append(sheet)

        chosen: list[str] = []
        for child_segment, sheets in by_child_segment.items():
            exact_depth = [
                sheet
                for sheet in sheets
                if len(tuple(sheet_path_tuples.get(sheet, ()))) == path_len + 1
            ]
            candidate_pool = exact_depth or sheets
            by_role: dict[str, list[str]] = {}
            for candidate in candidate_pool:
                by_role.setdefault(_sheet_measure_role(candidate), []).append(candidate)

            # If an exact-depth summary node exists (for example "Electricity plants"),
            # keep it but also pull explicit transformation role descendants so
            # child-sum totals can expose inputs/feedstock/product separately.
            if exact_depth:
                roles_in_exact = set(by_role.keys())
                roles_in_child_descendants = {_sheet_measure_role(candidate) for candidate in sheets}
                missing_explicit_roles = (
                    roles_in_child_descendants & explicit_measure_roles
                ) - roles_in_exact
                for missing_role in sorted(missing_explicit_roles):
                    role_candidates = [
                        candidate for candidate in sheets if _sheet_measure_role(candidate) == missing_role
                    ]
                    if not role_candidates:
                        continue
                    preferred = min(
                        role_candidates,
                        key=lambda sheet: (
                            len(tuple(sheet_path_tuples.get(sheet, ()))),
                            _sheet_sort_key(sheet),
                        ),
                    )
                    by_role.setdefault(missing_role, []).append(preferred)
            for role_sheets in by_role.values():
                chosen.append(
                    min(
                        role_sheets,
                        key=lambda sheet: (
                            len(tuple(sheet_path_tuples.get(sheet, ()))),
                            _sheet_sort_key(sheet),
                        ),
                    )
                )
        chosen = sorted(set(chosen), key=_sheet_sort_key)

        # Some nodes expose both a canonical category page and a one-to-one
        # alias page for the same underlying sector footprint (for example
        # bunkers parent labels and transport-style aliases). Keep a single
        # representative per equivalent direct-child comparator footprint so
        # node totals do not double count the same comparator rows.
        deduped: list[str] = []
        for sheet in chosen:
            sheet_leaves = sheet_leaf_sector_seqs.get(sheet, set())
            if not sheet_leaves:
                deduped.append(sheet)
                continue
            sheet_role = _structural_total_measure_bucket(
                _infer_sheet_measure(sheet),
                tuple(sheet_paths.get(sheet, [sheet])),
            )
            equivalent_idx: int | None = None
            for idx, other in enumerate(deduped):
                other_leaves = sheet_leaf_sector_seqs.get(other, set())
                if sheet_leaves != other_leaves:
                    continue
                other_role = _structural_total_measure_bucket(
                    _infer_sheet_measure(other),
                    tuple(sheet_paths.get(other, [other])),
                )
                if sheet_role != other_role:
                    continue
                equivalent_idx = idx
                break
            if equivalent_idx is None:
                deduped.append(sheet)
                continue

            incumbent = deduped[equivalent_idx]
            incumbent_path = tuple(sheet_paths.get(incumbent, [incumbent]))
            sheet_path = tuple(sheet_paths.get(sheet, [sheet]))
            preferred = min(
                [incumbent, sheet],
                key=lambda candidate: (
                    len(tuple(sheet_paths.get(candidate, [candidate]))),
                    _sheet_sort_key(candidate),
                ),
            )
            deduped[equivalent_idx] = preferred
        return sorted(set(deduped), key=_sheet_sort_key)

    def _non_overlapping_leaf_descendant_sheets(path: tuple[str, ...]) -> list[str]:
        leafs = _leaf_descendant_sheets(path)
        if len(leafs) <= 1:
            return leafs

        excluded: set[str] = set()
        all_numeric_by_sheet = {
            sheet: sheet_numeric_seqs.get(sheet, set())
            for sheet in leafs
        }
        for sheet in leafs:
            sheet_seqs = all_numeric_by_sheet.get(sheet, set())
            if not sheet_seqs:
                continue
            other_seqs = {
                seq
                for other in leafs
                if other != sheet
                for seq in all_numeric_by_sheet.get(other, set())
            }
            if not other_seqs:
                continue
            if all(
                any(
                    len(other_seq) > len(sheet_seq)
                    and other_seq[: len(sheet_seq)] == sheet_seq
                    for other_seq in other_seqs
                )
                for sheet_seq in sheet_seqs
            ):
                excluded.add(sheet)
                continue

        for sheet in leafs:
            if sheet in excluded:
                continue
            sheet_leaves = sheet_leaf_sector_seqs.get(sheet, set())
            if not sheet_leaves:
                continue
            sheet_path = tuple(sheet_paths.get(sheet, [sheet]))
            for other in leafs:
                if other == sheet or other in excluded:
                    continue
                other_leaves = sheet_leaf_sector_seqs.get(other, set())
                if not other_leaves:
                    continue
                other_path = tuple(sheet_paths.get(other, [other]))
                hierarchical_overlap = (
                    len(sheet_path) < len(other_path) and other_path[: len(sheet_path)] == sheet_path
                ) or (
                    len(other_path) < len(sheet_path) and sheet_path[: len(other_path)] == other_path
                )
                if not hierarchical_overlap:
                    # Distinct sibling leaves can legitimately share the same
                    # sector-code footprint (for example buildings end-uses all
                    # map to one residential/services sector). Do not dedupe
                    # those siblings here; only prune true parent/child-style
                    # overlaps within the dashboard hierarchy.
                    continue
                if sheet_leaves == other_leaves:
                    sheet_role = _structural_total_measure_bucket(
                        _infer_sheet_measure(sheet),
                        tuple(sheet_paths.get(sheet, [sheet])),
                    )
                    other_role = _structural_total_measure_bucket(
                        _infer_sheet_measure(other),
                        tuple(sheet_paths.get(other, [other])),
                    )
                    if sheet_role != other_role:
                        continue
                    if _sheet_sort_key(sheet) > _sheet_sort_key(other):
                        excluded.add(sheet)
                        break
                elif sheet_leaves > other_leaves:
                    sheet_role = _structural_total_measure_bucket(
                        _infer_sheet_measure(sheet),
                        tuple(sheet_paths.get(sheet, [sheet])),
                    )
                    other_role = _structural_total_measure_bucket(
                        _infer_sheet_measure(other),
                        tuple(sheet_paths.get(other, [other])),
                    )
                    if sheet_role != other_role:
                        continue
                    excluded.add(sheet)
                    break

        filtered_leafs = [sheet for sheet in leafs if sheet not in excluded]
        return filtered_leafs or leafs

    def _summary_component_sheets(path: tuple[str, ...]) -> list[str]:
        return _non_overlapping_leaf_descendant_sheets(path)

    def _sheet_has_explicit_measure_descendants(sheet: str) -> bool:
        sheet_path = sheet_path_tuples.get(sheet, ())
        if not sheet_path:
            return False
        own_measure_bucket = _structural_total_measure_bucket(
            _infer_sheet_measure(sheet),
            sheet_path,
        )
        if own_measure_bucket != "Summary (PJ)":
            return False
        descendants = [
            other
            for other, other_path in sheet_path_tuples.items()
            if other != sheet and other_path[: len(sheet_path)] == sheet_path
        ]
        if not descendants:
            return False
        descendant_measure_buckets = {
            _structural_total_measure_bucket(_infer_sheet_measure(other), sheet_path_tuples.get(other, (other,)))
            for other in descendants
        }
        explicit_measure_buckets = set(explicit_input_output_measure_buckets)
        return bool(descendant_measure_buckets & explicit_measure_buckets)

    def _should_render_sheet_section(sheet: str) -> bool:
        sheet_path = sheet_path_tuples.get(sheet, ())
        if not sheet_path:
            return True
        measure_bucket = _structural_total_measure_bucket(
            _infer_sheet_measure(sheet),
            sheet_path,
        )
        if sheet_path[0] in transformation_top_groups and measure_bucket == "Summary (PJ)":
            # Some grouped sheets (for example "Oil refineries") carry explicit
            # input/output measure rows in render_long even though sheet-level
            # inference returns a summary label. Keep those sections visible.
            explicit_buckets = set(explicit_input_output_measure_buckets)
            sheet_rows = render_long[render_long["sheet"].astype(str).eq(sheet)]
            if not sheet_rows.empty:
                row_buckets = {
                    _structural_total_measure_bucket(str(measure), sheet_path)
                    for measure in sheet_rows["measure"].dropna().astype(str).tolist()
                }
                if row_buckets & explicit_buckets:
                    return True
            return False
        if _sheet_has_explicit_measure_descendants(sheet):
            return False
        return True

    def _render_cards(section_sheet: str, entries: list[tuple[str, str, Path]]) -> str:
        if not entries:
            return ""
        def _grid_expand_class(item_count: int) -> str:
            if item_count == 1:
                return " expand-1"
            if item_count == 2:
                return " expand-2"
            if item_count == 3:
                return " expand-3"
            return ""
        grouped_entries: dict[str, list[tuple[str, str, Path]]] = {}
        for measure, fuel, png_path in entries:
            measure_key = str(measure).strip() or "Summary (PJ)"
            if _is_hidden_measure_section(measure_key):
                continue
            grouped_entries.setdefault(measure_key, []).append((str(measure).strip(), str(fuel), png_path))
        if not grouped_entries:
            return ""

        measure_order = sorted(grouped_entries.keys(), key=lambda item: _compact_measure_name(item).lower())
        measure_sections: list[str] = []
        for measure_key in measure_order:
            measure_entries = sorted(
                grouped_entries[measure_key],
                key=lambda item: (
                    0 if str(item[1]) == "Total" else 1,
                    -float(magnitude_lookup.get((section_sheet, str(item[0]).strip(), item[1]), 0.0) or 0.0),
                    str(item[1]).lower(),
                ),
            )
            cards: list[str] = []
            for measure, fuel, png_path in measure_entries:
                issue = base_issue_lookup.get((section_sheet, str(measure).strip(), fuel))
                rel_chart = os.path.relpath(png_path, start=dashboards_dir).replace("\\", "/")
                if png_path.suffix.lower() == ".html":
                    chart_markup = (
                        f'<iframe data-src="{rel_chart}" '
                        f'title="{section_sheet} – {fuel}" '
                        'class="lazy-chart-frame" '
                        'loading="lazy"></iframe>'
                    )
                else:
                    chart_markup = (
                        f'<img data-src="{rel_chart}" alt="{section_sheet} – {fuel}" '
                        'class="lazy-chart-image" '
                        'style="max-width:100%;height:auto;display:block;min-height:160px;background:#f8fafc;" '
                        'loading="lazy" />'
                    )
                card_style = "margin:0;padding:10px;border:1px solid #d0d7de;border-radius:10px;background:#fff;box-shadow:0 1px 2px rgba(0,0,0,0.05);"
                issue_badge = ""
                if issue:
                    severity = str(issue.get("severity", "Moderate"))
                    impact = str(issue.get("impact", "minor"))
                    palette = {
                        "Moderate": {"border": "rgba(217,119,6,{a})", "bg": "rgba(245,158,11,{b})", "text": "#92400e"},
                        "High": {"border": "rgba(234,88,12,{a})", "bg": "rgba(249,115,22,{b})", "text": "#9a3412"},
                        "Extreme": {"border": "rgba(220,38,38,{a})", "bg": "rgba(220,38,38,{b})", "text": "#991b1b"},
                    }.get(severity, {"border": "rgba(217,119,6,{a})", "bg": "rgba(245,158,11,{b})", "text": "#92400e"})
                    alpha = {"minor": "0.30", "medium": "0.42", "major": "0.58"}.get(impact, "0.30")
                    bg_alpha = {"minor": "0.04", "medium": "0.07", "major": "0.11"}.get(impact, "0.04")
                    card_style = (
                        "margin:0;padding:10px;border:2px solid "
                        + palette["border"].format(a=alpha)
                        + ";border-radius:8px;background:"
                        + palette["bg"].format(b=bg_alpha)
                        + ";box-shadow:0 1px 2px rgba(0,0,0,0.05);"
                    )
                    issue_badge = (
                        f'<div style="margin-top:4px;color:{palette["text"]};font-size:12px;font-weight:600;">'
                        f'{issue["label"]}</div>'
                    )
                cards.append(
                    f"""
<figure style="{card_style}">
  <figcaption style="font-weight:600;margin-bottom:4px;">{fuel}</figcaption>
  {issue_badge}
  {chart_markup}
</figure>
"""
                )
            measure_sections.append(
                f'<section class="measure-group" style="margin:6px 0 14px 0;">'
                f'<h3 style="margin:0 0 8px 0;font-size:14px;font-weight:600;color:#5b6470;">{measure_key}</h3>'
                f'<div class="dashboard-grid{_grid_expand_class(len(cards))}">{"".join(cards)}</div>'
                f'</section>'
            )
        return "".join(measure_sections)

    def _render_group_cards(group_label: str, entries: list[tuple[str, str, str, Path]]) -> str:
        if not entries:
            return ""
        def _grid_expand_class(item_count: int) -> str:
            if item_count == 1:
                return " expand-1"
            if item_count == 2:
                return " expand-2"
            if item_count == 3:
                return " expand-3"
            return ""
        grouped_entries: dict[str, list[tuple[str, str, str, Path]]] = {}
        for sheet, measure, fuel, png_path in entries:
            measure_key = str(measure).strip() or "Summary (PJ)"
            if _is_hidden_measure_section(measure_key):
                continue
            grouped_entries.setdefault(measure_key, []).append((str(sheet), str(measure).strip(), str(fuel), png_path))
        if not grouped_entries:
            return ""

        measure_order = sorted(grouped_entries.keys(), key=lambda item: _compact_measure_name(item).lower())
        measure_sections: list[str] = []
        sheets_in_group = {sheet for sheet, _, _, _ in entries}
        show_sheet_label = len(sheets_in_group) > 1
        for measure_key in measure_order:
            measure_entries = sorted(
                grouped_entries[measure_key],
                key=lambda item: (
                    0 if str(item[2]) == "Total" else 1,
                    -float(magnitude_lookup.get((item[0], str(item[1]).strip(), item[2]), 0.0) or 0.0),
                    str(item[2]).lower(),
                ),
            )
            cards: list[str] = []
            for sheet, measure, fuel, png_path in measure_entries:
                issue = base_issue_lookup.get((sheet, str(measure).strip(), fuel))
                rel_chart = os.path.relpath(png_path, start=dashboards_dir).replace("\\", "/")
                chart_title = f"{group_label} – {fuel}"
                if png_path.suffix.lower() == ".html":
                    chart_markup = (
                        f'<iframe data-src="{rel_chart}" '
                        f'title="{chart_title}" '
                        'class="lazy-chart-frame" '
                        'loading="lazy"></iframe>'
                    )
                else:
                    chart_markup = (
                        f'<img data-src="{rel_chart}" alt="{chart_title}" '
                        'class="lazy-chart-image" '
                        'style="max-width:100%;height:auto;display:block;min-height:160px;background:#f8fafc;" '
                        'loading="lazy" />'
                    )
                card_style = "margin:0;padding:10px;border:1px solid #d0d7de;border-radius:10px;background:#fff;box-shadow:0 1px 2px rgba(0,0,0,0.05);"
                issue_badge = ""
                if issue:
                    severity = str(issue.get("severity", "Moderate"))
                    impact = str(issue.get("impact", "minor"))
                    palette = {
                        "Moderate": {"border": "rgba(217,119,6,{a})", "bg": "rgba(245,158,11,{b})", "text": "#92400e"},
                        "High": {"border": "rgba(234,88,12,{a})", "bg": "rgba(249,115,22,{b})", "text": "#9a3412"},
                        "Extreme": {"border": "rgba(220,38,38,{a})", "bg": "rgba(220,38,38,{b})", "text": "#991b1b"},
                    }.get(severity, {"border": "rgba(217,119,6,{a})", "bg": "rgba(245,158,11,{b})", "text": "#92400e"})
                    alpha = {"minor": "0.30", "medium": "0.42", "major": "0.58"}.get(impact, "0.30")
                    bg_alpha = {"minor": "0.04", "medium": "0.07", "major": "0.11"}.get(impact, "0.04")
                    card_style = (
                        "margin:0;padding:10px;border:2px solid "
                        + palette["border"].format(a=alpha)
                        + ";border-radius:8px;background:"
                        + palette["bg"].format(b=bg_alpha)
                        + ";box-shadow:0 1px 2px rgba(0,0,0,0.05);"
                    )
                    issue_badge = (
                        f'<div style="margin-top:4px;color:{palette["text"]};font-size:12px;font-weight:600;">'
                        f'{issue["label"]}</div>'
                    )
                if show_sheet_label:
                    sheet_label = sheet_display_labels.get(sheet, sheet)
                    caption = f"{sheet_label}: {fuel}"
                else:
                    caption = fuel
                cards.append(
                    f"""
<figure style="{card_style}">
  <figcaption style="font-weight:600;margin-bottom:4px;">{caption}</figcaption>
  {issue_badge}
  {chart_markup}
</figure>
"""
                )
            measure_sections.append(
                f'<section class="measure-group" style="margin:6px 0 14px 0;">'
                f'<h3 style="margin:0 0 8px 0;font-size:14px;font-weight:600;color:#5b6470;">{measure_key}</h3>'
                f'<div class="dashboard-grid{_grid_expand_class(len(cards))}">{"".join(cards)}</div>'
                f'</section>'
            )
        return "".join(measure_sections)

    page_chart_counts: dict[tuple[str, ...], int] = {}
    for path in ordered_node_paths:
        desc = _descendant_sheets(path)
        page_chart_counts[path] = sum(len(sheet_entries_lookup.get(sheet, [])) for sheet in desc)

    nav_groups: dict[str, list[tuple[tuple[str, ...], str, str]]] = {}
    for path in ordered_node_paths:
        top_group = path[0] if path else "Other"
        if top_group not in nav_groups:
            nav_groups[top_group] = []
        selected_placeholder = "__SELECTED__"
        depth = max(0, len(path) - 1)
        leaf_label = _display_top_group(path[-1]) if depth == 0 else path[-1]
        label = f"{'  ' * depth}{leaf_label}"
        nav_groups[top_group].append((path, label, selected_placeholder))

    nav_group_order: list[str] = []
    for _, group_names in nav_group_blocks:
        present_group_names = [group_name for group_name in group_names if group_name in nav_groups]
        present_group_names = sorted(
            present_group_names,
            key=lambda group_name: (
                -float(node_abs_sum_lookup.get((group_name,), 0.0) or 0.0),
                _display_top_group(group_name).lower(),
            ),
        )
        for group_name in present_group_names:
            if group_name in nav_groups and group_name not in nav_group_order:
                nav_group_order.append(group_name)
    remaining_groups = [group_name for group_name in nav_groups if group_name not in nav_group_order]
    remaining_groups = sorted(
        remaining_groups,
        key=lambda group_name: (
            -float(node_abs_sum_lookup.get((group_name,), 0.0) or 0.0),
            _display_top_group(group_name).lower(),
        ),
    )
    for group_name in remaining_groups:
        if group_name not in nav_group_order:
            nav_group_order.append(group_name)

    ordered_nav_blocks: list[tuple[str, list[str]]] = []
    for block_label, group_names in nav_group_blocks:
        block_present = [group_name for group_name in group_names if group_name in nav_group_order]
        if block_present:
            ordered_nav_blocks.append((block_label, block_present))

    page_files: list[tuple[tuple[str, ...], Path, int]] = []
    for path in ordered_node_paths:
        dashboard_file = path_to_file[path]
        title = _display_top_group(path[-1]) if len(path) == 1 else path[-1]
        desc_sheets = _descendant_sheets(path)
        own_sheet = node_to_sheet.get(path)
        page_measure = _infer_page_measure(path, own_sheet)
        section_label_counts = Counter(
            sheet_display_labels.get(sheet, sheet)
            for sheet in desc_sheets
        )

        sections: list[str] = []
        section_jump_links: list[dict[str, object]] = []
        section_ids_seen: set[str] = set()

        def _new_section_id(label: str) -> str:
            base = _safe_token(str(label).replace("\\", "_"))
            if not base:
                base = "section"
            section_id = f"sec-{base}"
            suffix = 2
            while section_id in section_ids_seen:
                section_id = f"sec-{base}-{suffix}"
                suffix += 1
            section_ids_seen.add(section_id)
            return section_id

        def _append_section(
            *,
            jump_label: str,
            heading_html: str,
            cards_html: str,
            heading_margin: str,
            target_path: tuple[str, ...] | None = None,
            kind: str = "section",
        ) -> None:
            section_id = _new_section_id(jump_label)
            section_jump_links.append(
                {
                    "label": jump_label,
                    "section_id": section_id,
                    "level": 0 if kind == "structural_total" else _section_jump_level(path, target_path),
                    "kind": kind,
                    "target_path": target_path,
                }
            )
            sections.append(
                f'<section id="{section_id}" style="scroll-margin-top:140px;">'
                f'<h2 style="{heading_margin}">{heading_html}</h2>'
                f'{cards_html}</section>'
            )

        leaf_sheets = _summary_component_sheets(path)
        total_component_sheets = _direct_child_descendant_sheets(path)
        # Industry parent mapping can include an intermediate "Manufacturing"
        # node that is not a full rollup row; prefer leaf sheets there only.
        if path and str(path[-1]).strip().lower() == "industry sector":
            if leaf_sheets:
                total_component_sheets = leaf_sheets
        elif len(total_component_sheets) <= 1:
            total_component_sheets = leaf_sheets
        should_render_structural_total = (not own_sheet) or len(total_component_sheets) > 1
        direct_parent_conflict = bool(own_sheet) and len(total_component_sheets) > 1
        if should_render_structural_total and total_component_sheets:
            node_subset = render_long[
                render_long["sheet"].astype(str).isin(total_component_sheets)
                & (render_long["fuel_label"].astype(str) != "Total")
            ].copy()
            if not node_subset.empty:
                if path == ("Supply",):
                    overview_entries: list[tuple[str, str, Path]] = []
                    for child_sheet in ["Exports", "Imports", "Production"]:
                        section_total = _build_supply_overview_component_total(child_sheet)
                        if not _has_nonzero_chart_values(section_total):
                            continue
                        total_chart_path = make_chart(
                            child_sheet,
                            "Total",
                            section_total,
                            charts_dir,
                            backend="plotly",
                            display_sheet=f"{title} - {_compact_measure_name(_infer_page_measure(('Supply', child_sheet), child_sheet))}",
                            file_sheet=f"node__{'__'.join(_safe_token(part.replace('\\', '_')) for part in path)}__overview__{child_sheet}",
                        )
                        if total_chart_path:
                            overview_entries.append((child_sheet, "Total", total_chart_path))
                    if overview_entries:
                        _append_section(
                            jump_label="Supply overview",
                            heading_html="Supply overview",
                            cards_html=(
                                '<p style="margin:0 0 10px 0;color:#4b5563;font-size:13px;">'
                                f'{supply_overview_note}'
                                '</p>'
                                f'{_render_measure_total_cards(title, overview_entries)}'
                            ),
                            heading_margin="margin:18px 0 8px 0;",
                            target_path=path,
                            kind="structural_total",
                        )
                else:
                    total_entries: list[tuple[str, str, Path]] = []
                    node_subset["structural_total_measure"] = node_subset.apply(
                        lambda row: _structural_total_measure_bucket(
                            str(row.get("measure", "")),
                            path,
                        ),
                        axis=1,
                    )
                    for measure_value, node_subset_measure in node_subset.groupby("structural_total_measure", dropna=False):
                        node_total = _aggregate_display_rows_to_total(
                            node_subset_measure,
                            title=title,
                            measure_value=str(measure_value).strip(),
                            collapse_base_family=True,
                            collapse_projection_family=True,
                        )
                        if not _has_nonzero_chart_values(node_total):
                            continue
                        node_chart = make_chart(
                            title,
                            "Total",
                            node_total,
                            charts_dir,
                            backend="plotly",
                            display_sheet=f"{title} - {_compact_measure_name(str(measure_value).strip())}",
                            file_sheet=f"node__{'__'.join(_safe_token(part.replace('\\', '_')) for part in path)}__{measure_value}",
                        )
                        if node_chart:
                            normalized_measure = str(measure_value).strip()
                            if not normalized_measure or normalized_measure.lower() == "nan":
                                normalized_measure = _infer_page_measure(path, None)
                            total_entries.append((normalized_measure, "Total", node_chart))
                    if total_entries:
                        total_entries = _filter_hidden_measure_entries(total_entries)
                    if total_entries:
                        total_entries = sorted(total_entries, key=lambda item: str(item[0]).lower())
                        specific_total_measures = set(explicit_input_output_measure_buckets)
                        if any(measure in specific_total_measures for measure, _, _ in total_entries):
                            total_entries = [entry for entry in total_entries if entry[0] != "Summary (PJ)"]
                        if len(total_entries) == 1:
                            section_measure = total_entries[0][0] or _infer_page_measure(path, None)
                            structural_total_label = title
                            total_note = "Aggregated from displayed categories at the resolved comparison level."
                            override_notes = sorted(
                                {
                                    str(dashboard_note_override_by_sheet.get(sheet, "")).strip()
                                    for sheet in total_component_sheets
                                    if str(dashboard_note_override_by_sheet.get(sheet, "")).strip()
                                }
                            )
                            if override_notes:
                                total_note = total_note + " " + " ".join(override_notes)
                            _append_section(
                                jump_label=_compact_section_jump_label(title),
                                heading_html=_with_measure_badge(
                                    structural_total_label,
                                    section_measure,
                                ),
                                cards_html=(
                                    '<p style="margin:0 0 10px 0;color:#4b5563;font-size:13px;">'
                                    f'{total_note}'
                                    '</p>'
                                    f'{_render_cards(title, total_entries)}'
                                ),
                                heading_margin="margin:18px 0 8px 0;",
                                target_path=path,
                                kind="structural_total",
                            )
                        else:
                            totals_heading = title
                            totals_note = (
                                "Aggregated from displayed categories at the resolved comparison level. "
                                "Separate charts are shown for each measure."
                            )
                            override_notes = sorted(
                                {
                                    str(dashboard_note_override_by_sheet.get(sheet, "")).strip()
                                    for sheet in total_component_sheets
                                    if str(dashboard_note_override_by_sheet.get(sheet, "")).strip()
                                }
                            )
                            if override_notes:
                                totals_note = totals_note + " " + " ".join(override_notes)
                            _append_section(
                                jump_label=totals_heading,
                                heading_html=totals_heading,
                                cards_html=(
                                    '<p style="margin:0 0 10px 0;color:#4b5563;font-size:13px;">'
                                    f'{totals_note}'
                                    '</p>'
                                    f'{_render_measure_total_cards(title, total_entries)}'
                                ),
                                heading_margin="margin:18px 0 8px 0;",
                                target_path=path,
                                kind="structural_total",
                            )
        if path == ("Supply",) and not any(str(item.get("label", "")).strip().lower() == "supply overview" for item in section_jump_links):
            overview_entries: list[tuple[str, str, Path]] = []
            for child_sheet in ["Exports", "Imports", "Production"]:
                section_total = _build_supply_overview_component_total(child_sheet)
                if not _has_nonzero_chart_values(section_total):
                    continue
                total_chart_path = make_chart(
                    child_sheet,
                    "Total",
                    section_total,
                    charts_dir,
                    backend="plotly",
                    display_sheet=f"{title} - {_compact_measure_name(_infer_page_measure(('Supply', child_sheet), child_sheet))}",
                    file_sheet=f"node__{'__'.join(_safe_token(part.replace('\\', '_')) for part in path)}__overview__{child_sheet}",
                )
                if total_chart_path:
                    overview_entries.append((child_sheet, "Total", total_chart_path))
            if overview_entries:
                _append_section(
                    jump_label="Supply overview",
                    heading_html="Supply overview",
                    cards_html=(
                        '<p style="margin:0 0 10px 0;color:#4b5563;font-size:13px;">'
                        f'{supply_overview_note}'
                        '</p>'
                        f'{_render_measure_total_cards(title, overview_entries)}'
                    ),
                    heading_margin="margin:18px 0 8px 0;",
                    target_path=path,
                    kind="sheet",
                )
        render_own_sheet = (
            own_sheet
            and own_sheet in sheet_entries_lookup
            and _should_render_sheet_section(own_sheet)
            and not (
                should_render_structural_total
                and sheet_display_labels.get(own_sheet, own_sheet) == title
            )
        )
        if render_own_sheet:
            own_cards = _render_cards(own_sheet, sheet_entries_lookup[own_sheet])
            if own_cards:
                own_label = sheet_display_labels.get(own_sheet, own_sheet)
                own_token = _clean_token(own_label).lower()
                top_token = _clean_token(path[0]).lower() if path else ""
                if own_token in {top_token, f"{top_token} sector"}:
                    own_cards = ""
            if own_cards:
                if section_label_counts.get(own_label, 0) > 1:
                    own_label = own_sheet
                _append_section(
                    jump_label=_compact_section_jump_label(own_label),
                    heading_html=own_label,
                    cards_html=own_cards,
                    heading_margin="margin:18px 0 8px 0;",
                    target_path=sheet_path_tuples.get(own_sheet),
                    kind="sheet",
                )
        grouped_children: dict[str, list[str]] = {}
        group_order: list[str] = []
        for child_sheet in desc_sheets:
            if child_sheet == own_sheet:
                continue
            if not _should_render_sheet_section(child_sheet):
                continue
            if _clean_token(child_sheet).lower() in {
                _clean_token(title).lower(),
                _clean_token(path[0]).lower() if path else "",
            }:
                # Avoid duplicating the parent name as its own child section.
                continue
            group_label = final_category_name_by_sheet.get(child_sheet, "") or sheet_display_labels.get(child_sheet, child_sheet)
            group_token = _clean_token(group_label).lower()
            top_token = _clean_token(path[0]).lower() if path else ""
            if group_token in {top_token, f"{top_token} sector"}:
                continue
            if group_label not in grouped_children:
                grouped_children[group_label] = []
                group_order.append(group_label)
            grouped_children[group_label].append(child_sheet)

        group_order = sorted(
            group_order,
            key=lambda label: (
                -sum(float(sheet_abs_sum_lookup.get(str(sheet), 0.0) or 0.0) for sheet in grouped_children.get(label, [])),
                str(label).lower(),
            ),
        )
        for group_label in group_order:
            child_sheets = grouped_children.get(group_label, [])
            child_sheets = sorted(
                child_sheets,
                key=lambda sheet: (
                    path_order_index.get(tuple(sheet_path_tuples.get(sheet, (sheet,))), fallback_path_rank),
                    -float(sheet_abs_sum_lookup.get(str(sheet), 0.0) or 0.0),
                    _sheet_sort_key(sheet),
                ),
            )
            group_entries: list[tuple[str, str, str, Path]] = []
            for child_sheet in child_sheets:
                for measure, fuel, chart_path in sheet_entries_lookup.get(child_sheet, []):
                    group_entries.append((child_sheet, measure, fuel, chart_path))
            cards = _render_group_cards(group_label, group_entries)
            if not cards:
                continue
            group_paths = [sheet_path_tuples.get(child_sheet, ()) for child_sheet in child_sheets]
            group_paths = [path for path in group_paths if path]
            if group_paths:
                child_path = min(
                    group_paths,
                    key=lambda path: (
                        path_order_index.get(path, fallback_path_rank),
                        len(path),
                        tuple(part.lower() for part in path),
                    ),
                )
            else:
                child_path = path
            _append_section(
                jump_label=_compact_section_jump_label(group_label),
                heading_html=group_label,
                cards_html=cards,
                heading_margin="margin:22px 0 8px 0;",
                target_path=child_path,
                kind="sheet",
            )
        body_content = "".join(sections) if sections else '<p>No charts available for this category.</p>'

        nav_options = []
        for block_label, top_groups in ordered_nav_blocks:
            opts = []
            for top_group in top_groups:
                for other_path, label, _ in nav_groups[top_group]:
                    selected = " selected" if other_path == path else ""
                    opts.append(f'<option value="{path_to_file[other_path].name}"{selected}>{label}</option>')
            nav_options.append(f'<optgroup label="{block_label}">{"".join(opts)}</optgroup>')
        if mapping_status is not None and not mapping_status.empty:
            nav_options.append(
                '<optgroup label="Audit">'
                '<option value="mappings_audit.html">Mappings Audit</option>'
                '</optgroup>'
            )

        breadcrumb_parts = []
        for i in range(1, len(path) + 1):
            prefix = path[:i]
            file_path = path_to_file.get(prefix)
            name = _display_top_group(prefix[-1]) if len(prefix) == 1 else prefix[-1]
            if file_path:
                breadcrumb_parts.append(f'<a href="{file_path.name}">{name}</a>')
            else:
                breadcrumb_parts.append(name)
        breadcrumb = " > ".join(breadcrumb_parts)

        path_idx = path_order_index[path]
        prev_path = ordered_node_paths[path_idx - 1] if path_idx > 0 else None
        next_path = ordered_node_paths[path_idx + 1] if path_idx < len(ordered_node_paths) - 1 else None
        def _nav_button(target_path: tuple[str, ...] | None, label: str) -> str:
            if not target_path:
                return (
                    f'<span style="padding:6px 10px;border:1px solid #d8dee4;border-radius:6px;'
                    'color:#9ca3af;background:#eef2f5;">'
                    f'{label}</span>'
                )
            target_file = path_to_file[target_path].name
            return (
                f'<a href="{target_file}" style="padding:6px 10px;border:1px solid #c5ccd3;'
                'border-radius:6px;background:#fff;color:#0b3d5c;text-decoration:none;">'
                f'{label}</a>'
            )

        nav_button_row = (
            '<div style="display:flex;flex-wrap:wrap;gap:8px;">'
            + _nav_button(prev_path, 'Prev page')
            + _nav_button(next_path, 'Next page')
            + (
                '<a href="mappings_audit.html" style="padding:6px 10px;border:1px solid #c5ccd3;'
                'border-radius:6px;background:#fff;color:#0b3d5c;text-decoration:none;">Mappings Audit</a>'
                if mapping_status is not None and not mapping_status.empty
                else ""
            )
            + '</div>'
        )
        major_nav_buttons = [
            '<a href="index.html" style="padding:6px 10px;border:1px solid #c5ccd3;'
            'border-radius:6px;background:#fff;color:#0b3d5c;text-decoration:none;">Home</a>'
        ]
        for block_idx, (_, top_groups) in enumerate(ordered_nav_blocks):
            if block_idx > 0:
                major_nav_buttons.append(
                    '<span style="align-self:center;color:#94a3b8;font-weight:600;padding:0 2px;">|</span>'
                )
            for top_group in top_groups:
                top_path = (top_group,)
                if top_path not in path_to_file:
                    continue
                major_nav_buttons.append(
                    f'<a href="{path_to_file[top_path].name}" style="padding:6px 10px;border:1px solid #c5ccd3;'
                    'border-radius:6px;background:#fff;color:#0b3d5c;text-decoration:none;">'
                    f'{_display_top_group(top_group)}</a>'
                )
        major_nav_row = '<div style="display:flex;flex-wrap:wrap;gap:8px;">' + "".join(major_nav_buttons) + '</div>'
        section_jump_row = _render_section_jump_row(
            _prepare_section_jump_links(section_jump_links, path)
        )

        html_doc = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>{title} – LEAP Results Dashboard</title>
  <style>
    {page_shell_css}
  </style>
</head>
<body>
  <div class="page-shell">
    <div class="page-header" id="page-header">
      <div class="header-collapsible">
      <div class="header-main-row">
        <div style="min-width:220px;flex:1 1 320px;">
          <h1 style="font-size:24px;line-height:1.15;">{title}</h1>
          <div style="margin:4px 0 0 0;color:#4b5563;font-size:13px;line-height:1.3;">{breadcrumb}</div>
          <div style="margin:6px 0 0 0;color:#4b5563;font-size:12px;line-height:1.3;">Measure: {page_measure}</div>
        </div>
        <div class="header-side-controls">
          <label for="dashboard-picker" style="font-weight:600;white-space:nowrap;">View:</label>
          <select id="dashboard-picker" onchange="if (this.value) window.location.href=this.value;" style="min-width:220px;max-width:320px;padding:6px 10px;border:1px solid #c5ccd3;border-radius:6px;background:#fff;flex:1 1 220px;">
            <option value="index.html">Home</option>
            {''.join(nav_options)}
          </select>
          <div class="header-inline-controls">
            {nav_button_row}
          </div>
        </div>
      </div>
      <div style="margin-top:8px;">{major_nav_row}</div>
      {section_jump_row}
      </div>
      <div class="header-toggle-row">
        <button type="button" class="header-toggle" id="header-toggle" aria-expanded="true" aria-label="Collapse header">▴</button>
      </div>
    </div>
    <main class="page-body">
      {body_content}
    </main>
  </div>
  <script>
    {header_toggle_script}
    (function() {{
      const lazyNodes = Array.from(document.querySelectorAll('iframe[data-src], img[data-src]'));
      if (!lazyNodes.length) return;
      const activate = (node) => {{
        const src = node.getAttribute('data-src');
        if (!src) return;
        node.setAttribute('src', src);
        node.removeAttribute('data-src');
      }};
      if (!('IntersectionObserver' in window)) {{
        lazyNodes.forEach(activate);
        return;
      }}
      const observer = new IntersectionObserver((entries, obs) => {{
        entries.forEach((entry) => {{
          if (!entry.isIntersecting) return;
          activate(entry.target);
          obs.unobserve(entry.target);
        }});
      }}, {{
        rootMargin: '300px 0px'
      }});
      lazyNodes.forEach((node) => observer.observe(node));
    }})();
    (function() {{
      const chartFrames = Array.from(document.querySelectorAll('iframe.lazy-chart-frame'));
      if (!chartFrames.length) return;
      let lastPixelRatio = window.devicePixelRatio || 1;
      let resizeTimer = null;
      const applyFrameSizing = () => {{
        const viewportHeight = Math.max(
          document.documentElement.clientHeight || 0,
          window.innerHeight || 0,
        );
        chartFrames.forEach((frame) => {{
          const parentWidth = frame.parentElement ? frame.parentElement.clientWidth : 0;
          const frameWidth = frame.clientWidth || parentWidth || 0;
          const targetHeight = Math.max(
            380,
            Math.min(
              Math.round(viewportHeight * 0.64),
              Math.round(frameWidth * 0.72),
              1100,
            ),
          );
          frame.style.height = `${{targetHeight}}px`;
        }});
      }};
      const queueApplyFrameSizing = () => {{
        window.clearTimeout(resizeTimer);
        resizeTimer = window.setTimeout(applyFrameSizing, 80);
      }};
      applyFrameSizing();
      window.addEventListener('resize', queueApplyFrameSizing);
      if (window.visualViewport) {{
        window.visualViewport.addEventListener('resize', queueApplyFrameSizing);
      }}
      if ('ResizeObserver' in window) {{
        const observer = new ResizeObserver(queueApplyFrameSizing);
        observer.observe(document.documentElement);
        chartFrames.forEach((frame) => {{
          if (frame.parentElement) {{
            observer.observe(frame.parentElement);
          }}
        }});
      }}
      window.setInterval(() => {{
        const pixelRatio = window.devicePixelRatio || 1;
        if (pixelRatio !== lastPixelRatio) {{
          lastPixelRatio = pixelRatio;
          queueApplyFrameSizing();
        }}
      }}, 250);
    }})();
  </script>
</body>
</html>
"""
        dashboard_file.write_text(html_doc, encoding="utf-8")
        page_files.append((path, dashboard_file, page_chart_counts.get(path, 0)))

    def _linked_sheet_label(sheet: object) -> str:
        sheet_name = str(sheet or "").strip()
        if not sheet_name:
            return ""
        sheet_path = sheet_path_tuples.get(sheet_name)
        sheet_file = path_to_file.get(sheet_path) if sheet_path else None
        sheet_label = sheet_display_labels.get(sheet_name, sheet_name)
        if sheet_file:
            return f'<a href="{sheet_file.name}">{escape(sheet_label)}</a>'
        return escape(sheet_label)

    def _humanize_mapping_note_fragment(fragment: str) -> str:
        note = str(fragment or "").strip().rstrip(".")
        if not note:
            return ""
        if note == "aggregated canonical targets for sector+fuel conflict":
            return (
                "Multiple LEAP fuel labels were combined into one comparison bucket because "
                "there was not a clean one-to-one sector-and-fuel match."
            )
        if note == "aggregated canonical targets for sector+esto_product conflict":
            return (
                "Multiple canonical matches shared the same ESTO product, so they were combined "
                "into one comparison bucket."
            )
        if note == "aggregated explicit targets":
            return "Several manual mapping rules were grouped into one comparison bucket."
        if note == "ambiguous canonical matches for sector+fuel":
            return "The automatic mapping found more than one possible sector-and-fuel match."
        if note == "ambiguous canonical matches for sector+esto_product":
            return "The automatic mapping found more than one possible sector-and-product match."
        if note == "category labels treated as sectors":
            return "This label was treated as a sector name rather than a fuel name."
        if note.startswith("canonical parent fallback via "):
            parent_codes = note.removeprefix("canonical parent fallback via ").strip()
            return (
                "No exact child-sector match was available, so the mapping fell back to the parent "
                f"sector: {parent_codes}."
            )
        if note.startswith("projection_fuel_filter applied (") and note.endswith(")"):
            fuel_filters = note[len("projection_fuel_filter applied ("):-1].replace("|", ", ")
            return f"This sheet uses a projection fuel filter: {fuel_filters}."
        explicit_override_patterns = [
            (
                re.compile(
                    r"^explicit aggregate override for power (.+) outputs-by-feedstock across (.+) generation mappings$",
                    re.IGNORECASE,
                ),
                lambda m: (
                    f"Manual override: {m.group(1).capitalize()} generation outputs are intentionally grouped "
                    "into one output-by-feedstock comparison bucket."
                ),
            ),
            (
                re.compile(
                    r"^explicit aggregate override for primary (.+) production across (.+) product mappings$",
                    re.IGNORECASE,
                ),
                lambda m: (
                    f"Manual override: primary {m.group(1)} production is intentionally grouped into one "
                    "comparison bucket."
                ),
            ),
            (
                re.compile(
                    r"^projection-only explicit override for (.+) hydrogen outputs; esto has no auditable base analogue$",
                    re.IGNORECASE,
                ),
                lambda m: (
                    f"Manual override for {m.group(1)} hydrogen output. There is no directly auditable ESTO "
                    "base-year equivalent for this line."
                ),
            ),
        ]
        for pattern, formatter in explicit_override_patterns:
            match = pattern.match(note)
            if match:
                return formatter(match)
        return note[0].upper() + note[1:] + "." if note else ""

    def _humanize_mapping_notes(value: object, *, joiner: str = "; ") -> str:
        text = str(value or "").strip()
        if not text:
            return ""
        parts = re.split(r"\s*\|\s*|\s*;\s*", text)
        seen: set[str] = set()
        readable_parts: list[str] = []
        for part in parts:
            readable = _humanize_mapping_note_fragment(part)
            key = readable.lower()
            if not readable or key in seen:
                continue
            seen.add(key)
            readable_parts.append(readable)
        return joiner.join(readable_parts)

    def _render_audit_table(frame: pd.DataFrame, columns: list[tuple[str, str]], *, empty_message: str) -> str:
        if frame.empty:
            return (
                '<p style="margin:0;padding:14px 16px;border:1px solid #d8dee4;border-radius:10px;'
                'background:#fff;color:#4b5563;">'
                f'{escape(empty_message)}</p>'
            )
        header_html = "".join(
            f'<th style="position:sticky;top:0;background:#eff3f6;padding:8px 10px;border-bottom:1px solid #d8dee4;'
            f'text-align:left;white-space:nowrap;">{escape(label)}</th>'
            for _, label in columns
        )
        body_rows: list[str] = []
        for _, row in frame.iterrows():
            cells: list[str] = []
            for column, _ in columns:
                raw_value = row.get(column, "")
                if pd.isna(raw_value):
                    raw_value = ""
                text = str(raw_value).strip()
                if column == "sheet_link":
                    cell_html = text
                else:
                    cell_html = escape(text)
                cells.append(
                    '<td style="padding:8px 10px;border-bottom:1px solid #eef2f5;vertical-align:top;">'
                    f'{cell_html or "&nbsp;"}'
                    '</td>'
                )
            body_rows.append("<tr>" + "".join(cells) + "</tr>")
        return (
            '<div style="overflow:auto;border:1px solid #d8dee4;border-radius:10px;background:#fff;">'
            '<table style="width:100%;border-collapse:collapse;font-size:13px;line-height:1.35;">'
            f'<thead><tr>{header_html}</tr></thead>'
            f'<tbody>{"".join(body_rows)}</tbody>'
            '</table></div>'
        )

    audit_pages: list[tuple[str, str, str, str]] = []
    if mapping_status is not None and not mapping_status.empty:
        audit_status = mapping_status.copy()
        if "measure" not in audit_status.columns:
            audit_status["measure"] = ""
        audit_status["sheet"] = audit_status.get("sheet", "").fillna("").astype(str).str.strip()
        audit_status["measure"] = audit_status["measure"].fillna("").astype(str).str.strip()
        audit_status["fuel_label"] = audit_status.get("fuel_label", "").fillna("").astype(str).str.strip()
        for col in ["mapped", "base_mapping_complete", "projection_mapping_complete", "projection_parent_fallback"]:
            if col not in audit_status.columns:
                audit_status[col] = False
            audit_status[col] = audit_status[col].fillna(False).astype(bool)
        for col in [
            "sector_code_9th",
            "ninth_fuel_code",
            "esto_flow",
            "esto_product",
            "mapping_source",
            "flow_source",
            "fuel_source",
            "sector_match_method",
            "mapping_note",
            "comparator_scope",
        ]:
            if col not in audit_status.columns:
                audit_status[col] = ""
            audit_status[col] = audit_status[col].fillna("").astype(str).str.strip()
        audit_status["mapping_note"] = audit_status["mapping_note"].map(_humanize_mapping_notes)

        audit_status["sheet_link"] = audit_status["sheet"].map(_linked_sheet_label)

        unmapped = audit_status[
            (~audit_status["mapped"])
            | (~audit_status["base_mapping_complete"])
            | (~audit_status["projection_mapping_complete"])
        ].copy()
        unmapped["issue"] = ""
        unmapped.loc[~unmapped["mapped"], "issue"] = "Unmapped"
        unmapped.loc[unmapped["issue"].eq("") & ~unmapped["base_mapping_complete"], "issue"] = "Missing base mapping"
        unmapped.loc[
            unmapped["issue"].eq("") & ~unmapped["projection_mapping_complete"],
            "issue",
        ] = "Missing projection mapping"
        unmapped = unmapped.sort_values(["sheet", "measure", "fuel_label"], kind="mergesort")
        unmapped_table = _render_audit_table(
            unmapped,
            [
                ("sheet_link", "Sheet"),
                ("measure", "Measure"),
                ("fuel_label", "Fuel label"),
                ("issue", "Issue"),
                ("sector_code_9th", "9th sector"),
                ("ninth_fuel_code", "9th fuel"),
                ("esto_flow", "ESTO flow"),
                ("esto_product", "ESTO product"),
                ("mapping_source", "Mapping source"),
                ("flow_source", "Flow source"),
                ("fuel_source", "Fuel source"),
                ("mapping_note", "Mapping note"),
            ],
            empty_message="No unmapped or incomplete rows were detected.",
        )
        unmapped_body = (
            '<section style="margin:18px 0 22px 0;">'
            '<p style="margin:0 0 12px 0;color:#4b5563;font-size:13px;">'
            'Rows shown here are missing a complete base or projection mapping and should be treated as real mapping gaps.'
            '</p>'
            f'{unmapped_table}'
            '</section>'
        )
        audit_pages.append(
            (
                "unmapped_mappings.html",
                "Unmapped",
                f"Unmapped or incomplete rows: {len(unmapped)}",
                unmapped_body,
            )
        )

        mapping_rundown = (
            audit_status[
                [
                    "sheet",
                    "sheet_link",
                    "measure",
                    "fuel_label",
                    "mapped",
                    "sector_code_9th",
                    "ninth_fuel_code",
                    "esto_flow",
                    "esto_product",
                    "mapping_source",
                    "flow_source",
                    "fuel_source",
                    "sector_match_method",
                    "comparator_scope",
                    "projection_parent_fallback",
                    "mapping_note",
                ]
            ]
            .drop_duplicates()
            .sort_values(["sheet", "measure", "fuel_label"], kind="mergesort")
            .copy()
        )
        for col in ["mapped", "projection_parent_fallback"]:
            mapping_rundown[col] = mapping_rundown[col].map(lambda v: "Yes" if bool(v) else "")
        rundown_table = _render_audit_table(
            mapping_rundown,
            [
                ("sheet_link", "Sheet"),
                ("measure", "Measure"),
                ("fuel_label", "Fuel label"),
                ("mapped", "Mapped"),
                ("sector_code_9th", "9th sector"),
                ("ninth_fuel_code", "9th fuel"),
                ("esto_flow", "ESTO flow"),
                ("esto_product", "ESTO product"),
                ("mapping_source", "Mapping source"),
                ("flow_source", "Flow source"),
                ("fuel_source", "Fuel source"),
                ("sector_match_method", "Sector match"),
                ("comparator_scope", "Comparator scope"),
                ("projection_parent_fallback", "Parent fallback"),
                ("mapping_note", "Mapping note"),
            ],
            empty_message="No chart mapping rows were available.",
        )
        rundown_body = (
            '<section style="margin:18px 0 22px 0;">'
            '<p style="margin:0 0 12px 0;color:#4b5563;font-size:13px;">'
            'Each row below shows the resolved mapping used for a displayed chart line, so many-to-one mappings can be reviewed directly.'
            '</p>'
            f'{rundown_table}'
            '</section>'
        )
        audit_pages.append(
            (
                "chart_mapping_rundown.html",
                "Chart Mapping Rundown",
                f"Displayed chart-line mappings: {len(mapping_rundown)}",
                rundown_body,
            )
        )

        aggregate_groups = audit_status[audit_status["mapped"] & audit_status["ninth_fuel_code"].ne("")].copy()
        aggregate_groups = (
            aggregate_groups.groupby(
                ["sheet", "measure", "ninth_fuel_code", "sector_code_9th", "esto_flow", "esto_product"],
                as_index=False,
            )
            .agg(
                member_count=("fuel_label", "nunique"),
                member_fuels=("fuel_label", lambda s: ", ".join(sorted({str(v).strip() for v in s if str(v).strip()}))),
                mapping_sources=("mapping_source", lambda s: ", ".join(sorted({str(v).strip() for v in s if str(v).strip()}))),
                comparator_scopes=("comparator_scope", lambda s: ", ".join(sorted({str(v).strip() for v in s if str(v).strip()}))),
                notes=("mapping_note", lambda s: _humanize_mapping_notes(" | ".join([str(v).strip() for v in s if str(v).strip()]), joiner=" ")),
            )
        )
        aggregate_groups = aggregate_groups[aggregate_groups["member_count"] > 1].copy()
        aggregate_groups["sheet_link"] = aggregate_groups["sheet"].map(_linked_sheet_label)
        aggregate_groups = aggregate_groups.sort_values(
            ["member_count", "sheet", "measure", "ninth_fuel_code"],
            ascending=[False, True, True, True],
            kind="mergesort",
        )
        aggregate_table = _render_audit_table(
            aggregate_groups,
            [
                ("sheet_link", "Sheet"),
                ("measure", "Measure"),
                ("ninth_fuel_code", "Resolved 9th fuel"),
                ("sector_code_9th", "9th sector"),
                ("esto_flow", "ESTO flow"),
                ("esto_product", "ESTO product"),
                ("member_count", "Member count"),
                ("member_fuels", "Member fuel labels"),
                ("mapping_sources", "Mapping sources"),
                ("comparator_scopes", "Comparator scopes"),
                ("notes", "Mapping notes"),
            ],
            empty_message="No aggregate mapping groups were detected.",
        )
        aggregate_body = (
            '<section style="margin:18px 0 22px 0;">'
            '<p style="margin:0 0 12px 0;color:#4b5563;font-size:13px;">'
            'Rows shown here are the many-to-one groupings where multiple LEAP labels resolve to the same comparator fuel bucket.'
            '</p>'
            f'{aggregate_table}'
            '</section>'
        )
        audit_pages.append(
            (
                "aggregate_mapping_groups.html",
                "Aggregate Mapping Groups",
                f"Shared comparator buckets: {len(aggregate_groups)}",
                aggregate_body,
            )
        )

    audit_page_files: list[tuple[str, Path, str]] = []
    if audit_pages:
        audit_jump_links = "".join(
            f'<a href="#{_safe_token(title)}" '
            'style="padding:4px 9px;border:1px solid #c5ccd3;border-radius:999px;'
            'background:#fff;color:#0b3d5c;text-decoration:none;font-size:12px;line-height:1.25;">'
            f'{escape(title)}</a>'
            for _, title, _, _ in audit_pages
        )
        audit_sections = "".join(
            f'<section id="{_safe_token(title)}" style="scroll-margin-top:140px;">'
            f'<h2 style="margin:18px 0 8px 0;">{escape(title)}</h2>'
            f'<div style="margin:0 0 10px 0;color:#4b5563;font-size:12px;line-height:1.35;">{escape(subtitle)}</div>'
            f'{body_html}</section>'
            for _, title, subtitle, body_html in audit_pages
        )
        mappings_audit_html = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Mappings Audit – LEAP Results Dashboard</title>
  <style>
    {page_shell_css}
  </style>
</head>
<body>
  <div class="page-shell">
    <div class="page-header" id="page-header">
      <div class="header-collapsible">
      <div class="header-main-row">
        <div style="min-width:220px;flex:1 1 320px;">
          <h1 style="font-size:24px;line-height:1.15;">Mappings Audit</h1>
          <div style="margin:4px 0 0 0;color:#4b5563;font-size:13px;line-height:1.3;">Mappings Audit</div>
          <div style="margin:6px 0 0 0;color:#4b5563;font-size:12px;line-height:1.3;">Measure: Audit</div>
        </div>
        <div class="header-side-controls">
          <label for="dashboard-picker" style="font-weight:600;white-space:nowrap;">View:</label>
          <select id="dashboard-picker" onchange="if (this.value) window.location.href=this.value;" style="min-width:220px;max-width:320px;padding:6px 10px;border:1px solid #c5ccd3;border-radius:6px;background:#fff;flex:1 1 220px;">
            <option value="index.html">Home</option>
            {''.join(
                f'<optgroup label="{block_label}">' + "".join(
                    f'<option value="{path_to_file[other_path].name}">{label}</option>'
                    for top_group in top_groups
                    for other_path, label, _ in nav_groups[top_group]
                ) + '</optgroup>'
                for block_label, top_groups in ordered_nav_blocks
            )}
            <optgroup label="Audit"><option value="mappings_audit.html" selected>Mappings Audit</option></optgroup>
          </select>
          <div class="header-inline-controls">
            <div style="display:flex;flex-wrap:wrap;gap:8px;">
              <span style="padding:6px 10px;border:1px solid #d8dee4;border-radius:6px;color:#9ca3af;background:#eef2f5;">Prev page</span>
              <span style="padding:6px 10px;border:1px solid #d8dee4;border-radius:6px;color:#9ca3af;background:#eef2f5;">Next page</span>
            </div>
          </div>
        </div>
      </div>
      <div style="margin-top:8px;display:flex;flex-wrap:wrap;gap:8px;">
        <a href="index.html" style="padding:6px 10px;border:1px solid #c5ccd3;border-radius:6px;background:#fff;color:#0b3d5c;text-decoration:none;">Home</a>
        {''.join(
            (
                ('<span style="align-self:center;color:#94a3b8;font-weight:600;padding:0 2px;">|</span>' if block_idx > 0 else '')
                + ''.join(
                    f'<a href="{path_to_file[(top_group,)].name}" style="padding:6px 10px;border:1px solid #c5ccd3;border-radius:6px;background:#fff;color:#0b3d5c;text-decoration:none;">{_display_top_group(top_group)}</a>'
                    for top_group in top_groups
                    if (top_group,) in path_to_file
                )
            )
            for block_idx, (_, top_groups) in enumerate(ordered_nav_blocks)
        )}
      </div>
      <div style="margin-top:8px;padding-top:8px;border-top:1px solid #d8dee4;display:flex;flex-wrap:wrap;gap:6px 8px;align-items:center;">
        <span style="font-weight:600;color:#4b5563;font-size:12px;white-space:nowrap;">Jump to:</span>
        {audit_jump_links}
      </div>
      </div>
      <div class="header-toggle-row">
        <button type="button" class="header-toggle" id="header-toggle" aria-expanded="true" aria-label="Collapse header">▴</button>
      </div>
    </div>
    <main class="page-body">
      <section style="margin:18px 0 22px 0;">
        <p style="margin:0 0 12px 0;color:#4b5563;font-size:13px;">Use this page to inspect missing mappings, resolved chart-line mappings, and many-to-one aggregate fuel buckets in one place.</p>
      </section>
      {audit_sections}
    </main>
  </div>
  <script>
    {header_toggle_script}
  </script>
</body>
</html>
"""
        mappings_audit_path = dashboards_dir / "mappings_audit.html"
        mappings_audit_path.write_text(mappings_audit_html, encoding="utf-8")
        audit_page_files.append(("Mappings Audit", mappings_audit_path, "Audit landing page"))

        audit_links = "".join(
            f'<a href="{escape(filename)}" style="padding:6px 10px;border:1px solid #c5ccd3;'
            'border-radius:6px;background:#fff;color:#0b3d5c;text-decoration:none;">'
            f'{escape(title)}</a>'
            for filename, title, _, _ in audit_pages
        )
        for idx, (filename, title, subtitle, body_html) in enumerate(audit_pages):
            prev_file = audit_pages[idx - 1][0] if idx > 0 else ""
            next_file = audit_pages[idx + 1][0] if idx < len(audit_pages) - 1 else ""
            nav_buttons = [
                '<a href="index.html" style="padding:6px 10px;border:1px solid #c5ccd3;'
                'border-radius:6px;background:#fff;color:#0b3d5c;text-decoration:none;">Home</a>'
            ]
            if prev_file:
                nav_buttons.append(
                    f'<a href="{escape(prev_file)}" style="padding:6px 10px;border:1px solid #c5ccd3;'
                    'border-radius:6px;background:#fff;color:#0b3d5c;text-decoration:none;">Prev audit</a>'
                )
            if next_file:
                nav_buttons.append(
                    f'<a href="{escape(next_file)}" style="padding:6px 10px;border:1px solid #c5ccd3;'
                    'border-radius:6px;background:#fff;color:#0b3d5c;text-decoration:none;">Next audit</a>'
                )
            audit_html = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>{escape(title)} – LEAP Results Dashboard</title>
  <style>
    {page_shell_css}
  </style>
</head>
<body>
  <div class="page-shell">
    <div class="page-header">
      <div style="display:flex;flex-wrap:wrap;justify-content:space-between;align-items:flex-start;gap:10px 16px;">
        <div style="min-width:220px;flex:1 1 320px;">
          <h1>{escape(title)}</h1>
          <div style="margin:4px 0 0 0;color:#4b5563;font-size:13px;line-height:1.3;">Audit pages</div>
          <div style="margin:6px 0 0 0;color:#4b5563;font-size:12px;line-height:1.3;">{escape(subtitle)}</div>
        </div>
        <div style="display:flex;flex-wrap:wrap;gap:8px;align-items:center;justify-content:flex-end;">
          {''.join(nav_buttons)}
        </div>
      </div>
      <div style="margin-top:8px;display:flex;flex-wrap:wrap;gap:8px;">{audit_links}</div>
    </div>
    <main class="page-body">
      {body_html}
    </main>
  </div>
</body>
</html>
"""
            audit_path = dashboards_dir / filename
            audit_path.write_text(audit_html, encoding="utf-8")
            audit_page_files.append((title, audit_path, subtitle))

    tree: dict[str, dict] = {}
    for path in ordered_node_paths:
        node = tree
        for part in path:
            node = node.setdefault(part, {})

    def _render_tree(node: dict, depth: int = 0, path_prefix: tuple[str, ...] = ()) -> str:
        items: list[str] = []
        for name in sorted(node.keys(), key=lambda item: item.lower()):
            child = node[name]
            current_path = path_prefix + (name,)
            file_path = path_to_file.get(current_path)
            count = page_chart_counts.get(current_path, 0)
            label_html = f'<a href="{file_path.name}">{name}</a> <span style="color:#4b5563;">({count} charts)</span>' if file_path else name
            child_html = _render_tree(child, depth + 1, current_path)
            section_style = (
                "margin:10px 0 6px 0;padding:8px 10px;border-left:3px solid #c5ccd3;background:#fff;"
                if depth == 0
                else "margin:6px 0 6px 14px;padding-left:10px;border-left:1px solid #d8dee4;"
            )
            items.append(f'<li style="{section_style}">{label_html}{child_html}</li>')
        if not items:
            return ""
        return f'<ul style="list-style:none;margin:{6 if depth else 0}px 0 0 0;padding:0;">{"".join(items)}</ul>'

    links_html = _render_tree(tree)
    audit_links_html = ""
    if audit_page_files:
        audit_items = "".join(
            '<li style="margin:8px 0;padding:8px 10px;border-left:3px solid #c5ccd3;background:#fff;">'
            f'<a href="{file_path.name}">{escape(title)}</a>'
            f' <span style="color:#4b5563;">({escape(subtitle)})</span>'
            '</li>'
            for title, file_path, subtitle in audit_page_files
        )
        audit_links_html = (
            '<section style="margin:18px 0 24px 0;">'
            '<h2 style="margin:0 0 10px 0;color:#0f172a;font-size:22px;">Audit Pages</h2>'
            '<ul style="list-style:none;margin:0;padding:0;">'
            + audit_items
            + '</ul></section>'
        )
    issue_groups: dict[str, list[str]] = {}
    severity_rank = {"Extreme": 3, "High": 2, "Moderate": 1}
    impact_rank = {"major": 3, "medium": 2, "minor": 1}
    for (sheet, measure, fuel), issue in sorted(
        base_issue_lookup.items(),
        key=lambda item: (
            -severity_rank.get(str(item[1].get("severity", "")), 0),
            -impact_rank.get(str(item[1].get("impact", "")), 0),
            -float(item[1]["pct"]),
            item[0][0].lower(),
            item[0][2].lower(),
        ),
    ):
        sheet_path = sheet_path_tuples.get(sheet)
        file_path = path_to_file.get(sheet_path) if sheet_path else None
        sheet_label = sheet_display_labels.get(sheet, sheet)
        sheet_link = f'<a href="{file_path.name}">{sheet_label}</a>' if file_path else sheet_label
        top_group = sheet_path[0] if sheet_path else "Other"
        measure_text = str(measure).strip()
        measure_suffix = f" [{measure_text}]" if measure_text else ""
        issue_groups.setdefault(top_group, []).append(
            f'<li style="break-inside:avoid-column;margin-bottom:6px;"><span style="color:#b91c1c;font-weight:600;">'
            f'{issue["label"]} ({str(issue.get("impact", "minor")).capitalize()} impact)</span>: {sheet_link}{measure_suffix} – {fuel}</li>'
        )
    issue_count = sum(len(rows) for rows in issue_groups.values())
    issue_sections: list[str] = []
    shown_issue_count = 0
    for top_group in nav_group_order + sorted(k for k in issue_groups if k not in nav_group_order):
        rows = issue_groups.get(top_group, [])
        if not rows:
            continue
        remaining = max(0, 100 - shown_issue_count)
        if remaining <= 0:
            break
        shown_rows = rows[:remaining]
        shown_issue_count += len(shown_rows)
        issue_sections.append(
            '<section style="padding:10px 12px;border:1px solid rgba(185,28,28,0.14);'
            'border-radius:8px;background:rgba(255,255,255,0.72);">'
            f'<h3 style="margin:0 0 8px 0;color:#7f1d1d;font-size:16px;">{top_group}</h3>'
            '<ul style="margin:0;padding-left:18px;line-height:1.5;columns:2;column-gap:20px;">'
            + "".join(shown_rows)
            + '</ul></section>'
        )
    issues_html = (
        '<section style="margin:18px 0 24px 0;padding:12px 14px;border:1px solid rgba(220,38,38,0.25);'
        'background:rgba(220,38,38,0.04);border-radius:10px;">'
        '<h2 style="margin:0 0 8px 0;color:#991b1b;font-size:20px;">Significant Base-Year Differences</h2>'
        '<div style="display:grid;grid-template-columns:repeat(auto-fit, minmax(360px, 1fr));gap:12px;">'
        + "".join(issue_sections)
        + '</div>'
        + (f'<p style="margin:8px 0 0 0;color:#7f1d1d;">Showing first 100 issues.</p>' if issue_count > 100 else "")
        + '</section>'
        if issue_count
        else ""
    )
    index_file = dashboards_dir / "index.html"
    index_file.write_text(
        f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>LEAP Results Dashboards</title>
  <style>
    {page_shell_css}
  </style>
</head>
<body>
  <div class="page-shell">
    <div class="page-header" id="page-header">
      <div class="header-collapsible">
      <div class="header-main-row">
        <div style="min-width:220px;flex:1 1 320px;">
          <h1 style="font-size:24px;line-height:1.15;">LEAP Results Dashboards</h1>
          <div style="margin:4px 0 0 0;color:#4b5563;font-size:13px;line-height:1.3;">Home</div>
          <div style="margin:6px 0 0 0;color:#4b5563;font-size:12px;line-height:1.3;">Measure: Overview</div>
        </div>
        <div class="header-side-controls">
          <label for="dashboard-picker" style="font-weight:600;white-space:nowrap;">View:</label>
          <select id="dashboard-picker" onchange="if (this.value) window.location.href=this.value;" style="min-width:220px;max-width:320px;padding:6px 10px;border:1px solid #c5ccd3;border-radius:6px;background:#fff;flex:1 1 220px;">
            <option value="index.html" selected>Home</option>
            {''.join(
                f'<optgroup label="{block_label}">' + "".join(
                    f'<option value="{path_to_file[other_path].name}">{label}</option>'
                    for top_group in top_groups
                    for other_path, label, _ in nav_groups[top_group]
                ) + '</optgroup>'
                for block_label, top_groups in ordered_nav_blocks
            )}
            {(
                '<optgroup label="Audit">'
                '<option value="mappings_audit.html">Mappings Audit</option>'
                '</optgroup>'
            ) if audit_page_files else ''}
          </select>
          <div class="header-inline-controls">
            <div style="display:flex;flex-wrap:wrap;gap:8px;">
              <span style="padding:6px 10px;border:1px solid #d8dee4;border-radius:6px;color:#9ca3af;background:#eef2f5;">Prev page</span>
              <span style="padding:6px 10px;border:1px solid #d8dee4;border-radius:6px;color:#9ca3af;background:#eef2f5;">Next page</span>
              {(
                  '<a href="mappings_audit.html" style="padding:6px 10px;border:1px solid #c5ccd3;'
                  'border-radius:6px;background:#fff;color:#0b3d5c;text-decoration:none;">Mappings Audit</a>'
              ) if audit_page_files else ''}
            </div>
          </div>
        </div>
      </div>
      <div style="margin-top:8px;display:flex;flex-wrap:wrap;gap:8px;">
        <a href="index.html" style="padding:6px 10px;border:1px solid #c5ccd3;border-radius:6px;background:#fff;color:#0b3d5c;text-decoration:none;">Home</a>
        {''.join(
            (
                ('<span style="align-self:center;color:#94a3b8;font-weight:600;padding:0 2px;">|</span>' if block_idx > 0 else '')
                + ''.join(
                    f'<a href="{path_to_file[(top_group,)].name}" style="padding:6px 10px;border:1px solid #c5ccd3;border-radius:6px;background:#fff;color:#0b3d5c;text-decoration:none;">{_display_top_group(top_group)}</a>'
                    for top_group in top_groups
                    if (top_group,) in path_to_file
                )
            )
            for block_idx, (_, top_groups) in enumerate(ordered_nav_blocks)
        )}
      </div>
      </div>
      <div class="header-toggle-row">
        <button type="button" class="header-toggle" id="header-toggle" aria-expanded="true" aria-label="Collapse header">▴</button>
      </div>
    </div>
    <main class="page-body">
      {issues_html}
      {audit_links_html}
      {links_html}
    </main>
  </div>
  <script>
    {header_toggle_script}
  </script>
</body>
</html>
""",
        encoding="utf-8",
    )
    print(f"[INFO] Generated dashboards: {index_file}")
    return index_file


# -----------------------------------------------------------------------------
# Diagnostics / checks
# -----------------------------------------------------------------------------
def basic_checks(
    sheet_map: pd.DataFrame,
    fuel_mapping: dict[str, dict[str, str]],
    comparison_long: pd.DataFrame,
    mapping_status: pd.DataFrame,
    *,
    fuel_coverage_threshold: float = 0.95,
) -> dict[str, object]:
    """Return summary diagnostics."""
    coverage = 0.0
    if not mapping_status.empty:
        coverage = mapping_status["mapped"].mean()

    has_all_sheets = sheet_map["sheet_name"].isin(comparison_long["sheet"].unique()).mean() == 1.0 if not comparison_long.empty else False

    issues = []
    if coverage < fuel_coverage_threshold:
        issues.append(f"Fuel mapping coverage {coverage:.2%} below target {fuel_coverage_threshold:.0%}.")
    if not has_all_sheets:
        issues.append("Some mapped sheets missing in comparison output.")

    return {
        "fuel_mapping_coverage": coverage,
        "all_sheets_present": has_all_sheets,
        "issues": issues,
    }


#%%
