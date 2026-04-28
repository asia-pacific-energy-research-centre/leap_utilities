#%%
# Summary: Label and optionally export subtotal rows in the ESTO (Matt) dataset.
import os
import json
import hashlib
from pathlib import Path
import pandas as pd

from codebase.utilities.master_config import config_table_exists, read_config_table

from codebase.mappings.canonical_mapping import load_canonical_pairs
from codebase.utilities.leap_results_dashboard_utils import (
    apply_explicit_sector_reassignments,
    load_explicit_sector_fuel_mappings,
    load_explicit_sector_reassignments,
)
from codebase.utilities.leap_results_dashboard_v2.reference_loader import (
    append_synthetic_reference_rows,
    load_synthetic_reference_rows_config,
)
#%%

#%%
######### CONSTANTS (UNLIKELY TO CHANGE) #########
ENABLE_DEBUG_BREAKPOINTS = True
#%%

#%%
######### FUNCTIONS #########
def ensure_repo_root():
    """Move to repo root if running from the scrapbook folder."""
    try:
        if os.getcwd().endswith("scrapbook"):
            os.chdir("../../")
    except Exception as exc:
        print(f"Failed to set repo root: {exc}")
        try_debug_breakpoint()
        raise


def try_debug_breakpoint():
    """Trigger a debug breakpoint when enabled (safe to call anywhere)."""
    if not ENABLE_DEBUG_BREAKPOINTS:
        return
    try:
        breakpoint()
    except Exception as breakpoint_exc:
        print(f"Debug breakpoint failed: {breakpoint_exc}")


def apply_matt_subtotal_mapping(df, mapping_path):
    """Merge subtotal mapping onto the Matt/ESTO data to flag subtotal rows.

    Inputs:
        df: ESTO (Matt) dataframe with flows/products columns.
        mapping_path: Excel mapping file path.
    Outputs:
        Dataframe with boolean is_subtotal column.
    Side effects:
        Reads from disk.
    """
    try:
        if "is_subtotal" in df.columns:
            matt = df.copy()
            matt["is_subtotal"] = (
                matt["is_subtotal"]
                .fillna(False)
                .astype(str)
                .str.strip()
                .str.lower()
                .isin(["true", "1", "yes"])
            )
            return matt

        mapping = read_config_table(mapping_path, dtype=str)
        normalized_cols = {col: str(col).strip().lower() for col in mapping.columns}
        mapping = mapping.rename(columns=normalized_cols)

        # New mapping format: single sheet with flow/product/is_subtotal columns.
        if {"flow", "product", "is_subtotal"}.issubset(mapping.columns) or {
            "flows",
            "products",
            "is_subtotal",
        }.issubset(mapping.columns):
            if {"flows", "products"}.issubset(mapping.columns):
                mapping = mapping.rename(columns={"flows": "flow", "products": "product"})
            mapping = mapping[["flow", "product", "is_subtotal"]].copy()
            mapping["flow"] = mapping["flow"].astype(str).str.strip()
            mapping["product"] = mapping["product"].astype(str).str.strip()
            mapping["is_subtotal"] = (
                mapping["is_subtotal"]
                .astype(str)
                .str.strip()
                .str.lower()
                .isin(["true", "1", "yes"])
            )
            mapping = (
                mapping.groupby(["flow", "product"], dropna=False)["is_subtotal"]
                .any()
                .reset_index()
            )

            matt = df.copy()
            matt["flow"] = matt["flows"].astype(str).str.strip()
            matt["product"] = matt["products"].astype(str).str.strip()
            matt = matt.merge(mapping, on=["flow", "product"], how="left")
            matt["is_subtotal"] = matt["is_subtotal"].fillna(False)
            matt = matt.drop(columns=[col for col in ["flow", "product"] if col in matt.columns])
        else:
            # Legacy mapping format: separate FLOW/PRODUCT sheets and codebase/name columns.
            mapping["is_subtotal"] = (
                mapping["is_subtotal"]
                .astype(str)
                .str.strip()
                .str.lower()
                .isin(["true", "1", "yes"])
            )

            flows_mapping = mapping[mapping["type"] == "FLOWS"].copy()
            flows_mapping["flow_value"] = (
                flows_mapping["new"].astype(str).str.strip()
                + " "
                + flows_mapping["new_name"].astype(str).str.strip()
            )

            matt = df.copy()
            matt["flow_value"] = matt["flows"].astype(str).str.strip()
            matt = matt.merge(
                flows_mapping[["flow_value", "is_subtotal"]],
                on="flow_value",
                how="left",
            )
            matt = matt.rename(columns={"is_subtotal": "flow_is_subtotal"})

            products_mapping = mapping[mapping["type"] == "PRODUCTS"].copy()
            products_mapping["product_value"] = (
                products_mapping["new"].astype(str).str.strip()
                + " "
                + products_mapping["new_name"].astype(str).str.strip()
            )
            matt["product_value"] = matt["products"].astype(str).str.strip()
            matt = matt.merge(
                products_mapping[["product_value", "is_subtotal"]],
                on="product_value",
                how="left",
            )
            matt = matt.rename(columns={"is_subtotal": "product_is_subtotal"})

            matt["flow_is_subtotal"] = matt["flow_is_subtotal"].fillna(False)
            matt["product_is_subtotal"] = matt["product_is_subtotal"].fillna(False)
            matt["is_subtotal"] = matt["flow_is_subtotal"] | matt["product_is_subtotal"]
            drop_cols = [
                "flow_value",
                "flow_is_subtotal",
                "product_value",
                "product_is_subtotal",
            ]
            matt = matt.drop(columns=[col for col in drop_cols if col in matt.columns])
        year_cols = [col for col in matt.columns if str(col).isdigit()]
        leading_cols = [
            col for col in ["economy", "flows", "products"] if col in matt.columns
        ]
        other_cols = [
            col
            for col in matt.columns
            if col not in leading_cols and col != "is_subtotal" and col not in year_cols
        ]
        ordered_cols = leading_cols + ["is_subtotal"] + other_cols + year_cols
        matt = matt[ordered_cols]
        return matt
    except Exception as exc:
        print(f"Failed to apply subtotal mapping to Matt data: {exc}")
        try_debug_breakpoint()
        raise


def filter_matt_subtotals(df):
    """Drop subtotal rows from the Matt/ESTO dataset when flagged."""
    try:
        if "is_subtotal" not in df.columns:
            return df
        return df[df["is_subtotal"] == False].copy()
    except Exception as exc:
        print(f"Failed to filter Matt subtotals: {exc}")
        try_debug_breakpoint()
        raise


def save_subtotal_labeled_data(df, output_path, label):
    """Save a subtotal-labeled dataset for inspection."""
    try:
        if df is None or df.empty:
            print(f"No data to save for {label}; skipping {output_path}.")
            return
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        df.to_csv(output_path, index=False)
        print(f"Saved {label} with subtotal labels to {output_path}")
    except Exception as exc:
        print(f"Failed to save {label} to {output_path}: {exc}")
        try_debug_breakpoint()
        raise


def _file_signature(path):
    path = Path(path)
    if not path.exists():
        return {"path": str(path), "exists": False}
    stat = path.stat()
    return {
        "path": str(path.resolve()),
        "exists": True,
        "size": int(stat.st_size),
        "mtime_ns": int(stat.st_mtime_ns),
    }


def _build_reference_cache_key(payload):
    blob = json.dumps(payload, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(blob).hexdigest()[:16]


def load_augmented_reference_tables(
    *,
    esto_path,
    ninth_path,
    subtotal_mapping_path=None,
    explicit_reassignments_path=None,
    explicit_mappings_path=None,
    canonical_pairs_path=None,
    synthetic_rules_path="config/synthetic_reference_rows.csv",
    cache_dir="data/.cache/augmented_reference_tables",
    apply_esto_subtotal_map=False,
    filter_esto_subtotals_flag=False,
    filter_ninth_subtotals_flag=False,
):
    """Load ESTO + 9th tables with optional subtotal cleanup, explicit reassignments,
    and synthetic-row augmentation, then cache the augmented result on disk.

    The cache is invalidated automatically when any source/config file or option changes.
    """
    ensure_repo_root()
    esto_path = Path(esto_path)
    ninth_path = Path(ninth_path)
    subtotal_mapping_path = Path(subtotal_mapping_path) if subtotal_mapping_path else None
    explicit_reassignments_path = (
        Path(explicit_reassignments_path) if explicit_reassignments_path else None
    )
    explicit_mappings_path = Path(explicit_mappings_path) if explicit_mappings_path else None
    canonical_pairs_path = Path(canonical_pairs_path) if canonical_pairs_path else None
    synthetic_rules_path = Path(synthetic_rules_path) if synthetic_rules_path else None
    cache_dir = Path(cache_dir)
    cache_dir.mkdir(parents=True, exist_ok=True)

    cache_payload = {
        "esto": _file_signature(esto_path),
        "ninth": _file_signature(ninth_path),
        "subtotal_mapping": _file_signature(subtotal_mapping_path) if subtotal_mapping_path else {},
        "explicit_reassignments": _file_signature(explicit_reassignments_path) if explicit_reassignments_path else {},
        "explicit_mappings": _file_signature(explicit_mappings_path) if explicit_mappings_path else {},
        "canonical_pairs": _file_signature(canonical_pairs_path) if canonical_pairs_path else {},
        "synthetic_rules": _file_signature(synthetic_rules_path) if synthetic_rules_path else {},
        "options": {
            "apply_esto_subtotal_map": bool(apply_esto_subtotal_map),
            "filter_esto_subtotals_flag": bool(filter_esto_subtotals_flag),
            "filter_ninth_subtotals_flag": bool(filter_ninth_subtotals_flag),
        },
    }
    cache_key = _build_reference_cache_key(cache_payload)
    esto_cache = cache_dir / f"{cache_key}_esto.csv"
    ninth_cache = cache_dir / f"{cache_key}_ninth.csv"
    meta_cache = cache_dir / f"{cache_key}_meta.json"

    if esto_cache.exists() and ninth_cache.exists() and meta_cache.exists():
        return read_config_table(esto_cache), read_config_table(ninth_cache)

    esto_df = read_config_table(esto_path)
    ninth_df = read_config_table(ninth_path, low_memory=False)

    if apply_esto_subtotal_map and subtotal_mapping_path:
        esto_df = apply_matt_subtotal_mapping(esto_df, subtotal_mapping_path)
    if filter_esto_subtotals_flag:
        esto_df = filter_matt_subtotals(esto_df)
    if filter_ninth_subtotals_flag:
        for subtotal_col in ["subtotal_results", "subtotal_layout"]:
            if subtotal_col in ninth_df.columns:
                mask = (
                    ninth_df[subtotal_col]
                    .fillna(False)
                    .astype(str)
                    .str.strip()
                    .str.lower()
                    .isin({"true", "1", "yes"})
                )
                ninth_df = ninth_df.loc[~mask].copy()

    explicit_reassignments = (
        load_explicit_sector_reassignments(explicit_reassignments_path)
        if explicit_reassignments_path and config_table_exists(explicit_reassignments_path)
        else pd.DataFrame()
    )
    if not explicit_reassignments.empty:
        esto_df, ninth_df, _ = apply_explicit_sector_reassignments(
            esto_df,
            ninth_df,
            explicit_reassignments,
        )

    explicit_mappings = (
        load_explicit_sector_fuel_mappings(explicit_mappings_path)
        if explicit_mappings_path and explicit_mappings_path.exists()
        else pd.DataFrame()
    )
    canonical_pairs = (
        load_canonical_pairs(canonical_pairs_path, strict=False)[0]
        if canonical_pairs_path and config_table_exists(canonical_pairs_path)
        else pd.DataFrame()
    )
    synthetic_rules = (
        load_synthetic_reference_rows_config(synthetic_rules_path)
        if synthetic_rules_path and config_table_exists(synthetic_rules_path)
        else pd.DataFrame()
    )
    if not synthetic_rules.empty:
        esto_df, ninth_df, _ = append_synthetic_reference_rows(
            esto_df=esto_df,
            ninth_df=ninth_df,
            rules=synthetic_rules,
            explicit_mappings=explicit_mappings,
            canonical_pairs=canonical_pairs,
        )

    esto_df.to_csv(esto_cache, index=False)
    ninth_df.to_csv(ninth_cache, index=False)
    meta_cache.write_text(json.dumps(cache_payload, indent=2))
    return esto_df, ninth_df
#%%

#%%
######### CONSTANTS (LIKELY TO CHANGE) #########
RUN_LABEL_SUBTOTALS = False
ESTO_DATA_PATH = "data/00APEC_2024_low.csv"
SUBTOTAL_MAPPING_PATH = "config/ESTO_subtotal_mapping.xlsx"
SAVE_ESTO_SUBTOTAL_LABELED = False
ESTO_SUBTOTAL_LABELED_OUTPUT_PATH = "data/00APEC_2024_low_with_subtotals.csv"
#%%

#%%
######### RUN LABELING (TOGGLE) #########
if RUN_LABEL_SUBTOTALS:
    try:
        ensure_repo_root()
        raw_df = read_config_table(ESTO_DATA_PATH)
        labeled = apply_matt_subtotal_mapping(raw_df, SUBTOTAL_MAPPING_PATH)
        if SAVE_ESTO_SUBTOTAL_LABELED:
            save_subtotal_labeled_data(
                labeled,
                ESTO_SUBTOTAL_LABELED_OUTPUT_PATH,
                "ESTO (Matt) data",
            )
        cleaned = filter_matt_subtotals(labeled)
        print(
            "Subtotal labeling complete. "
            f"Rows: raw={raw_df.shape[0]}, labeled={labeled.shape[0]}, "
            f"cleaned={cleaned.shape[0]}"
        )
    except Exception as exc:
        print(f"Failed to run subtotal labeling: {exc}")
        try_debug_breakpoint()
#%%
