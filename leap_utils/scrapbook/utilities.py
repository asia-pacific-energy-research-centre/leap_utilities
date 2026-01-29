#%%
# Summary: Label and optionally export subtotal rows in the ESTO (Matt) dataset.
import os
import pandas as pd
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
        mapping = pd.read_excel(mapping_path, dtype=str)
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
            # Legacy mapping format: separate FLOW/PRODUCT sheets and code/name columns.
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
        raw_df = pd.read_csv(ESTO_DATA_PATH)
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
