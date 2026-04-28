from __future__ import annotations

import re
from pathlib import Path
from typing import TypeAlias

import pandas as pd

from codebase.utilities.master_config import config_table_exists, read_config_table


def _safe_read_codebook_sheet(codebook_path: Path, sheet_name: str) -> pd.DataFrame:
    try:
        return read_config_table(codebook_path, sheet_name=sheet_name)
    except (FileNotFoundError, ValueError) as exc:
        print(f"[WARN] Codebook read failed for sheet {sheet_name!r} in {codebook_path}: {exc}")
        return pd.DataFrame()

DEFAULT_SHEET_MAP = Path("config/leap_results_sheet_map.csv")
DEFAULT_BACKUP_LEAP_MAPPINGS = Path("config/backup_leap_mappings.xlsx")
DEFAULT_CODEBOOK = Path("config/sector_fuel_codes_to_names.xlsx")
DEFAULT_NINTH_TO_ESTO = Path("config/ninth_pairs_to_esto_pairs.xlsx")

SECTOR_COLUMNS = ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]
PARENT_CODE_MATCH_PREFIX = "parent_code_match_"
ConfigTableRef: TypeAlias = Path | str | tuple[Path | str, str]


def split_config_table_ref(ref: ConfigTableRef) -> tuple[Path | str, str | None]:
    if isinstance(ref, tuple):
        if len(ref) != 2:
            raise ValueError("Config table references must be (path, sheet_name).")
        path, sheet_name = ref
        return path, str(sheet_name).strip() or None
    return ref, None


def clean_token(value: object) -> str:
    text = str(value or "").strip()
    return "" if text.lower() == "nan" else text


def normalize_label(value: object) -> str:
    return " ".join(str(value or "").strip().lower().split())


def build_code_match_method(levels_up: int) -> str:
    levels_up = max(int(levels_up), 0)
    if levels_up == 0:
        return "direct_code_match"
    return f"{PARENT_CODE_MATCH_PREFIX}{levels_up}_levels_up"


def normalize_match_method(value: object) -> str:
    text = normalize_label(value)
    if not text:
        return ""
    if text == "code_exact":
        return build_code_match_method(0)
    if text == "direct_code_match":
        return text
    old_parent_match = re.fullmatch(r"code_ancestor_(\d+)", text)
    if old_parent_match:
        return build_code_match_method(int(old_parent_match.group(1)))
    new_parent_match = re.fullmatch(rf"{PARENT_CODE_MATCH_PREFIX}(\d+)_levels_up", text)
    if new_parent_match:
        return build_code_match_method(int(new_parent_match.group(1)))
    remap = {
        "independent_exact": "independent_table_direct_match",
        "independent_exact_x": "independent_table_direct_match_x_category",
        "nonspecified_fallback": "nonspecified_category_fallback",
        "independent_reverse": "independent_table_reverse_lookup",
        "unmatched_esto_pair_nonzero": "unmapped_nonzero_esto_pair",
        "explicit_override": "manual_override",
        "skip_unallocated_or_x": "skipped_x_or_unallocated",
    }
    return remap.get(text, text)


def is_parent_code_match(value: object) -> bool:
    return normalize_match_method(value).startswith(PARENT_CODE_MATCH_PREFIX)


def is_reverse_independent_match(value: object) -> bool:
    return normalize_match_method(value) == "independent_table_reverse_lookup"


def split_sector_codes(raw_value: object) -> list[str]:
    text = str(raw_value or "").strip()
    if not text or text.lower() == "nan":
        return []
    parts = re.split(r"\s*(?:,|;|\||\band\b)\s*", text, flags=re.IGNORECASE)
    seen: set[str] = set()
    out: list[str] = []
    for part in parts:
        token = part.strip()
        if not token:
            continue
        key = token.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(token)
    return out


def load_sheet_map(path: Path = DEFAULT_SHEET_MAP) -> pd.DataFrame:
    df = read_config_table(path)
    df.columns = [c.strip().lower() for c in df.columns]
    if "projection_fuel_filter" not in df.columns:
        df["projection_fuel_filter"] = ""
    if "active" in df.columns:
        df = df[df["active"].astype(str).str.lower().isin({"true", "1", "yes"})]
    if "sheet_name" in df.columns:
        df["sheet_name"] = df["sheet_name"].astype(str).str.strip()
    if "sector_code_9th" in df.columns:
        df["sector_code_9th"] = df["sector_code_9th"].astype(str).str.strip()
    df["projection_fuel_filter"] = df["projection_fuel_filter"].fillna("").astype(str).str.strip()
    if "category_type" in df.columns:
        df["category_type"] = (
            df["category_type"]
            .astype(str)
            .str.strip()
            .str.lower()
            .replace({"": "fuel", "nan": "fuel"})
        )
    return df.reset_index(drop=True)


def load_canonical_pairs(path: ConfigTableRef = DEFAULT_NINTH_TO_ESTO, *, strict: bool = False) -> tuple[pd.DataFrame, pd.DataFrame]:
    path, sheet_name = split_config_table_ref(path)
    path = Path(path)
    if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        raw = read_config_table(path, sheet_name=sheet_name)
    else:
        raw = read_config_table(path)
    raw.columns = [str(c).strip().lower() for c in raw.columns]
    required = ["9th_sector", "9th_fuel", "esto_flow", "esto_product"]
    if any(c not in raw.columns for c in required) and path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        try:
            raw = read_config_table(path, sheet_name="leap_combined_mapping")
            raw.columns = [str(c).strip().lower() for c in raw.columns]
            raw = raw.rename(
                columns={
                    "ninth_sector": "9th_sector",
                    "ninth_fuel": "9th_fuel",
                }
            )
        except Exception:
            pass
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError(f"Canonical pairs file missing required columns: {missing}")

    optional = [c for c in ["sector_match_method", "fuel_match_method", "mapping_note"] if c in raw.columns]
    keep_cols = required + optional
    df = raw[keep_cols].copy()
    for col in keep_cols:
        df[col] = df[col].map(clean_token)
    for col in ["sector_match_method", "fuel_match_method"]:
        if col in df.columns:
            df[col] = df[col].map(normalize_match_method)
    df = df[(df["9th_sector"] != "") & (df["9th_fuel"] != "")]
    df = df[(df["esto_flow"] != "") & (df["esto_product"] != "")]
    if df.empty:
        empty_conflicts = pd.DataFrame(columns=["9th_sector", "9th_fuel", "issue", "details"])
        return df, empty_conflicts

    grouped = (
        df.groupby(["9th_sector", "9th_fuel"], dropna=False)[["esto_flow", "esto_product"]]
        .nunique(dropna=False)
        .reset_index()
    )
    bad = grouped[(grouped["esto_flow"] > 1) | (grouped["esto_product"] > 1)][["9th_sector", "9th_fuel"]]
    if bad.empty:
        conflicts = pd.DataFrame(columns=["9th_sector", "9th_fuel", "issue", "details"])
    else:
        rows: list[dict[str, str]] = []
        for sector, fuel in set(zip(bad["9th_sector"], bad["9th_fuel"])):
            vals = df[(df["9th_sector"] == sector) & (df["9th_fuel"] == fuel)][["esto_flow", "esto_product"]].drop_duplicates()
            details = "; ".join(f"{r.esto_flow} | {r.esto_product}" for r in vals.itertuples(index=False))
            rows.append(
                {
                    "9th_sector": sector,
                    "9th_fuel": fuel,
                    "issue": "duplicate_canonical_key_inconsistent_target",
                    "details": details,
                }
            )
        conflicts = pd.DataFrame(rows).sort_values(["9th_sector", "9th_fuel"]).reset_index(drop=True)
        if strict:
            raise ValueError(f"Conflicting canonical mappings found: {len(conflicts)} key(s)")

    clean = df.drop_duplicates(subset=keep_cols).copy()
    clean = clean.sort_values(["9th_sector", "9th_fuel", "esto_flow", "esto_product"] + optional).reset_index(drop=True)
    return clean, conflicts


def build_product_fuel_crosswalk(canonical_pairs: pd.DataFrame) -> pd.DataFrame:
    df = canonical_pairs.copy()
    for col in ["9th_fuel", "esto_product", "9th_sector", "esto_flow"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].map(clean_token)
    out = (
        df[df["9th_fuel"].ne("") & df["esto_product"].ne("")]
        .groupby(["esto_product", "9th_fuel"], as_index=False)
        .agg(
            sector_count=("9th_sector", "nunique"),
            flow_count=("esto_flow", "nunique"),
        )
        .sort_values(["esto_product", "9th_fuel"])
        .reset_index(drop=True)
    )
    return out


def build_flow_sector_crosswalk(canonical_pairs: pd.DataFrame) -> pd.DataFrame:
    df = canonical_pairs.copy()
    for col in ["9th_sector", "esto_flow", "9th_fuel", "esto_product"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].map(clean_token)
    out = (
        df[df["9th_sector"].ne("") & df["esto_flow"].ne("")]
        .groupby(["esto_flow", "9th_sector"], as_index=False)
        .agg(
            fuel_count=("9th_fuel", "nunique"),
            product_count=("esto_product", "nunique"),
        )
        .sort_values(["esto_flow", "9th_sector"])
        .reset_index(drop=True)
    )
    return out


def build_sector_to_esto_flow_lookup(codebook_path: Path = DEFAULT_CODEBOOK) -> dict[str, str]:
    df = _safe_read_codebook_sheet(codebook_path, "code_to_name")
    if df.empty:
        return {}
    lookup: dict[str, str] = {}
    valid_cols = {c.lower() for c in SECTOR_COLUMNS}
    for _, row in df.iterrows():
        ninth = clean_token(row.get("9th_label")).lower()
        ninth_col = clean_token(row.get("9th_column")).lower()
        esto = clean_token(row.get("esto_label"))
        if ninth and ninth_col in valid_cols and esto:
            lookup[ninth] = esto
    return lookup


def build_leap_label_crosswalk(
    codebook_path: Path = DEFAULT_CODEBOOK,
    override_path: Path | str | None = DEFAULT_BACKUP_LEAP_MAPPINGS,
) -> pd.DataFrame:
    leap_df = _safe_read_codebook_sheet(codebook_path, "ESTO_LEAP_names")
    if leap_df.empty:
        return pd.DataFrame(columns=["leap_label", "leap_label_norm", "esto_product", "mapping_source"])
    leap_df = leap_df[leap_df["category"].astype(str).str.strip().str.lower() == "products"].copy()
    rows: list[dict[str, str]] = []
    for _, row in leap_df.iterrows():
        leap_label = clean_token(row.get("leap_name"))
        esto_product = clean_token(row.get("original_label"))
        if leap_label and esto_product:
            rows.append(
                {
                    "leap_label": leap_label,
                    "leap_label_norm": normalize_label(leap_label),
                    "esto_product": esto_product,
                    "mapping_source": "codebook_fallback",
                }
            )

    code_df = _safe_read_codebook_sheet(codebook_path, "code_to_name")
    if code_df.empty:
        return pd.DataFrame(columns=["leap_label", "leap_label_norm", "esto_product", "mapping_source"])
    code_rows = code_df.copy()
    if "esto_column" in code_rows.columns:
        code_rows = code_rows[code_rows["esto_column"].astype(str).str.strip().str.lower() == "products"]
    code_rows = code_rows[["name", "esto_label"]].copy()
    code_rows["name"] = code_rows["name"].map(clean_token)
    code_rows["esto_label"] = code_rows["esto_label"].map(clean_token)
    for _, row in code_rows.iterrows():
        leap_label = row["name"]
        esto_product = row["esto_label"]
        if leap_label and esto_product:
            rows.append(
                {
                    "leap_label": leap_label,
                    "leap_label_norm": normalize_label(leap_label),
                    "esto_product": esto_product,
                    "mapping_source": "codebook_fallback",
                }
            )

    # Synthetic placeholders.
    for leap_label, esto_product in [("hydrogen", "16.09 Other sources"), ("efuel", "16.09 Other sources"), ("ammonia", "16.09 Other sources")]:
        rows.append(
            {
                "leap_label": leap_label,
                "leap_label_norm": normalize_label(leap_label),
                "esto_product": esto_product,
                "mapping_source": "codebook_fallback",
            }
        )

    out = pd.DataFrame(rows).drop_duplicates(subset=["leap_label_norm", "esto_product"], keep="first")

    if override_path and config_table_exists(override_path):
        p = Path(override_path)
        if p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
            ov = read_config_table(p)
        else:
            ov = read_config_table(p)
        ov.columns = [c.strip().lower() for c in ov.columns]
        for _, row in ov.iterrows():
            label = clean_token(row.get("leap_fuel_label"))
            if not label:
                continue
            out = out[out["leap_label_norm"] != normalize_label(label)]
            out = pd.concat(
                [
                    out,
                    pd.DataFrame(
                        [
                            {
                                "leap_label": label,
                                "leap_label_norm": normalize_label(label),
                                "esto_product": clean_token(row.get("esto_product_override")),
                                "mapping_source": "override",
                            }
                        ]
                    ),
                ],
                ignore_index=True,
            )

    out = out.sort_values(["leap_label_norm", "mapping_source"]).drop_duplicates(subset=["leap_label_norm"], keep="first")
    return out.reset_index(drop=True)


def build_leap_branch_crosswalk(
    sheet_map_path: Path = DEFAULT_SHEET_MAP,
    codebook_path: Path = DEFAULT_CODEBOOK,
) -> pd.DataFrame:
    sheet_map = load_sheet_map(sheet_map_path)
    sector_to_flow = build_sector_to_esto_flow_lookup(codebook_path)
    rows: list[dict[str, str]] = []
    for _, row in sheet_map.iterrows():
        sheet = clean_token(row.get("sheet_name"))
        sectors = split_sector_codes(row.get("sector_code_9th"))
        if not sectors:
            sectors = [clean_token(row.get("sector_code_9th"))]
        for sector in sectors:
            flow = clean_token(row.get("esto_flow_override"))
            flow_source = "sheet_override" if flow else ""
            if not flow:
                flow = clean_token(sector_to_flow.get(str(sector).lower(), ""))
                flow_source = "flow_sector_map" if flow else ""
            rows.append(
                {
                    "sheet_name": sheet,
                    "9th_sector": sector,
                    "esto_flow": flow,
                    "flow_source": flow_source,
                }
            )
    return pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)


def load_fuel_aliases(
    alias_path: Path | str | None = DEFAULT_BACKUP_LEAP_MAPPINGS,
    codebook_path: Path = DEFAULT_CODEBOOK,
) -> dict[str, dict[str, str]]:
    crosswalk = build_leap_label_crosswalk(codebook_path=codebook_path, override_path=alias_path)
    mapping: dict[str, dict[str, str]] = {}
    for _, row in crosswalk.iterrows():
        key = normalize_label(row.get("leap_label"))
        mapping[key] = {
            "ninth_fuel": "",
            "esto_product": clean_token(row.get("esto_product")),
            "esto_flow": "",
            "mapping_source": clean_token(row.get("mapping_source")) or "codebook_fallback",
            "flow_source": "",
            "fuel_source": "inferred",
        }
    return mapping


def map_fuel_label(
    fuel_label: str,
    fuel_mapping: dict[str, dict[str, str]],
    fallback_codebook: dict[str, str] | None = None,
) -> dict[str, str]:
    key = normalize_label(fuel_label)
    entry = dict(fuel_mapping.get(key, {}))
    ninth = clean_token(entry.get("ninth_fuel", ""))
    if not ninth and fallback_codebook:
        ninth = clean_token(fallback_codebook.get(key, ""))
    return {
        "ninth_fuel": ninth,
        "esto_product": clean_token(entry.get("esto_product", "")),
        "esto_flow": clean_token(entry.get("esto_flow", "")),
        "mapping_source": clean_token(entry.get("mapping_source", "")),
        "flow_source": clean_token(entry.get("flow_source", "")),
        "fuel_source": clean_token(entry.get("fuel_source", "")),
    }
