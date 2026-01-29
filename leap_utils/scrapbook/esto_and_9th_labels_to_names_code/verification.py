#%%
# Summary: Apply manual name harmonization fixes to the code_to_name worksheet and write outputs back to Excel.
from __future__ import annotations

from pathlib import Path
from typing import Dict, Optional, Tuple

import pandas as pd
#%%

#%%
######### CONSTANTS (UNLIKELY TO CHANGE) #########
ENABLE_DEBUG_BREAKPOINTS = True
CODE_TO_NAME_REQUIRED_COLUMNS = {"code", "name", "source_sheet"}
LABEL_REQUIRED_COLUMNS = {"9th_label", "esto_label", "name"}
MAPPING_COLUMNS = ["9th_label", "9th_column", "esto_label", "esto_column", "name"]
NINTH_LABEL_COLUMNS = [
    "fuels",
    "subfuels",
    "sectors",
    "sub1sectors",
    "sub2sectors",
    "sub3sectors",
    "sub4sectors",
]
NINTH_SECTOR_COLUMNS = ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]
NINTH_FUEL_COLUMNS = ["fuels", "subfuels"]
ESTO_COLUMN_CANDIDATES = ["flows", "products"]
ESTO_LABEL_COLUMN = "full code"
#%%

#%%
######### FUNCTIONS #########
def try_debug_breakpoint() -> None:
    """Trigger a debug breakpoint when enabled (safe to call anywhere)."""
    if not ENABLE_DEBUG_BREAKPOINTS:
        return
    try:
        breakpoint()
    except Exception as breakpoint_exc:
        print(f"Debug breakpoint failed: {breakpoint_exc}")


def load_code_to_name_sheet(workbook_path: Path, sheet_name: str) -> pd.DataFrame:
    """Load the code_to_name sheet from the workbook.

    Inputs:
        workbook_path: Path to the Excel workbook.
        sheet_name: Sheet name to read.
    Outputs:
        Dataframe with the sheet contents.
    Side effects:
        Reads from disk.
    """
    try:
        return pd.read_excel(workbook_path, sheet_name=sheet_name, dtype=str).fillna("")
    except Exception as exc:
        print(f"Failed to read {sheet_name} from {workbook_path}: {exc}")
        try_debug_breakpoint()
        raise


def load_reference_sheet(workbook_path: Path, sheet_name: str) -> pd.DataFrame:
    """Load a reference sheet (9th or ESTO) from the workbook.

    Inputs:
        workbook_path: Path to the Excel workbook.
        sheet_name: Sheet name to read.
    Outputs:
        Dataframe with the sheet contents.
    Side effects:
        Reads from disk.
    """
    try:
        return pd.read_excel(workbook_path, sheet_name=sheet_name, dtype=str).fillna("")
    except Exception as exc:
        print(f"Failed to read {sheet_name} from {workbook_path}: {exc}")
        try_debug_breakpoint()
        raise


def add_label_columns_for_fixes(df: pd.DataFrame) -> pd.DataFrame:
    """Add 9th_label / esto_label columns from the code_to_name structure.

    Inputs:
        df: Dataframe from the code_to_name sheet.
    Outputs:
        Dataframe with 9th_label and esto_label columns added.
    Side effects:
        None.
    """
    try:
        if LABEL_REQUIRED_COLUMNS.issubset(df.columns):
            working = df.copy()
            working["9th_label"] = working["9th_label"].fillna("").astype(str).str.strip()
            working["esto_label"] = working["esto_label"].fillna("").astype(str).str.strip()
            return working

        missing = CODE_TO_NAME_REQUIRED_COLUMNS - set(df.columns)
        if missing:
            raise ValueError(
                "Missing required columns for code_to_name input. "
                f"Need either {sorted(LABEL_REQUIRED_COLUMNS)} or {sorted(CODE_TO_NAME_REQUIRED_COLUMNS)}. "
                f"Missing: {sorted(missing)}"
            )

        working = df.copy()
        working["source_sheet_clean"] = working["source_sheet"].astype(str).str.strip().str.lower()
        working["code_clean"] = working["code"].astype(str).str.strip()
        working["9th_label"] = ""
        working["esto_label"] = ""

        is_9th = working["source_sheet_clean"] == "9th"
        is_esto = working["source_sheet_clean"] == "esto"

        working.loc[is_9th, "9th_label"] = working.loc[is_9th, "code_clean"]
        working.loc[is_esto, "esto_label"] = working.loc[is_esto, "code_clean"]
        return working
    except Exception as exc:
        print(f"Failed to prepare labels for fixes: {exc}")
        try_debug_breakpoint()
        raise


def apply_name_fixes(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Apply manual name harmonisation fixes based on 9th/ESTO codes.

    Inputs:
        df: Dataframe with name, 9th_label, and esto_label columns.
    Outputs:
        Tuple of (updated dataframe, audit dataframe).
    Side effects:
        None.
    """
    df_out = df.copy()

    # Canonical name rules keyed by exact ESTO labels.
    esto_label_to_canonical: Dict[str, str] = {
        "16 Other sector": "Others",
        "13 Total final energy consumption": "Tide, wave, ocean",
        "15.01 Domestic air transport": "Fuelwood & woodwaste",
        "08.01 Recycled products": "Natural gas",
        "11 Statistical discrepancy": "Geothermal",
        "10 Losses & own use": "Hydro",
        "15.03 Rail": "Charcoal",
        "15.02 Road": "Bagasse",
        "16.02 Residential": "Industrial waste",
        "08.02 Interproduct transfers": "LNG",
    }

    # Canonical name rules keyed by exact 9th labels.
    ninth_label_to_canonical_exact: Dict[str, str] = {
        "09_13_06_others": "Others",
        "09_nuclear": "Nuclear",
        "19_total": "Total",
        "13_tide_wave_ocean": "Tide wave ocean",
        "08_01_natural_gas": "Natural gas",
        "11_geothermal": "Geothermal",
        "10_hydro": "Hydro",
        "15_03_charcoal": "Charcoal",
        "15_02_bagasse": "Bagasse",
        "16_02_industrial_waste": "Industrial waste",
        "08_02_lng": "LNG",
        "14_wind": "Wind",
    }

    # For 9th labels, match by numeric prefix at the start.
    ninth_prefix_to_canonical: Dict[str, str] = {}

    def canonical_for_row(row: pd.Series) -> Optional[str]:
        esto_label = (row.get("esto_label") or "").strip()
        ninth_label = (row.get("9th_label") or "").strip()

        if esto_label in esto_label_to_canonical:
            return esto_label_to_canonical[esto_label]

        if ninth_label in ninth_label_to_canonical_exact:
            return ninth_label_to_canonical_exact[ninth_label]

        for pref, canon in ninth_prefix_to_canonical.items():
            if ninth_label.startswith(pref):
                return canon

        return None

    old_name = df_out["name"].astype(str)
    canon_name = df_out.apply(canonical_for_row, axis=1).fillna("")

    # Update only where we have a canonical name and it differs.
    mask = (canon_name != "") & (old_name != canon_name)
    df_out.loc[mask, "name"] = canon_name[mask]

    # Audit table.
    audit_columns = [col for col in ["code", "source_sheet", "9th_label", "esto_label"] if col in df_out.columns]
    changes = df_out.loc[mask, audit_columns].copy()
    changes.insert(0, "old_name", old_name[mask].values)
    changes.insert(1, "new_name", canon_name[mask].values)

    return df_out, changes


def write_excel_sheet(workbook_path: Path, output_df: pd.DataFrame, sheet_name: str) -> None:
    """Write a dataframe to a sheet (replace if it exists).

    Inputs:
        workbook_path: Excel workbook path.
        output_df: Dataframe to write.
        sheet_name: Sheet name to create or replace.
    Outputs:
        None.
    Side effects:
        Writes to disk.
    """
    try:
        with pd.ExcelWriter(
            workbook_path, mode="a", if_sheet_exists="replace", engine="openpyxl"
        ) as writer:
            output_df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Wrote {output_df.shape[0]} rows to {workbook_path}:{sheet_name}")
    except Exception as exc:
        print(f"Failed to write {sheet_name} to {workbook_path}: {exc}")
        try_debug_breakpoint()
        raise


def report_missing_labels(df: pd.DataFrame, label_column: str) -> pd.DataFrame:
    """Report rows missing a specific label column.

    Inputs:
        df: Dataframe containing label columns.
        label_column: Column name to check for missing/blank values.
    Outputs:
        Dataframe of rows with missing labels.
    Side effects:
        None.
    """
    if label_column not in df.columns:
        print(f"Missing column for label check: {label_column}")
        return pd.DataFrame()
    missing_mask = df[label_column].fillna("").astype(str).str.strip() == ""
    available_columns = [col for col in MAPPING_COLUMNS if col in df.columns]
    return df.loc[missing_mask, available_columns].copy()


def report_duplicate_values(df: pd.DataFrame, column: str) -> pd.DataFrame:
    """Report duplicate non-blank values for a column.

    Inputs:
        df: Dataframe to check.
        column: Column name to test for duplicates.
    Outputs:
        Dataframe of duplicate values and counts.
    Side effects:
        None.
    """
    if column not in df.columns:
        print(f"Missing column for duplicate check: {column}")
        return pd.DataFrame()
    values = df[column].fillna("").astype(str).str.strip()
    counts = values[values != ""].value_counts()
    duplicates = counts[counts > 1].rename("count").reset_index().rename(columns={"index": column})
    return duplicates


def run_validation_checks(df: pd.DataFrame) -> None:
    """Run missing label and duplicate checks for the code_to_name dataset.

    Inputs:
        df: Dataframe containing the code_to_name data.
    Outputs:
        None.
    Side effects:
        Prints diagnostic information.
    """
    try:
        missing_9th = report_missing_labels(df, "9th_label")
        missing_esto = report_missing_labels(df, "esto_label")

        print(f"Missing 9th labels: {len(missing_9th)}")
        if len(missing_9th) and FAIL_ON_MISSING_LABELS:
            raise ValueError("Missing 9th labels detected.")

        print(f"Missing ESTO labels: {len(missing_esto)}")
        if len(missing_esto) and FAIL_ON_MISSING_LABELS:
            raise ValueError("Missing ESTO labels detected.")

        for column in MAPPING_COLUMNS:
            duplicates = report_duplicate_values(df, column)
            print(f"Duplicate values in {column}: {len(duplicates)}")
            if len(duplicates):
                if column in {"9th_column", "esto_label", "esto_column"}:
                    print(duplicates.to_string(index=False))
                else:
                    print(duplicates.head(10).to_string(index=False))
                if FAIL_ON_DUPLICATES:
                    raise ValueError(f"Duplicate values detected in {column}.")
    except Exception as exc:
        print(f"Validation checks failed: {exc}")
        try_debug_breakpoint()
        raise


def run_mapping_cardinality_checks(df: pd.DataFrame) -> pd.DataFrame:
    """Check one-to-one/one-to-many/many-to-one/many-to-many mappings.

    Inputs:
        df: Dataframe containing the code_to_name data.
    Outputs:
        Dataframe of many-to-many rows (empty if none).
    Side effects:
        Prints diagnostic counts.
    """
    try:
        if "9th_label" not in df.columns or "esto_label" not in df.columns:
            print("Missing 9th_label or esto_label for mapping cardinality checks.")
            return pd.DataFrame()

        working = df.copy()
        working["9th_label_clean"] = working["9th_label"].fillna("").astype(str).str.strip()
        working["esto_label_clean"] = working["esto_label"].fillna("").astype(str).str.strip()
        working = working[
            (working["9th_label_clean"] != "") & (working["esto_label_clean"] != "")
        ]
        working = working.drop_duplicates(
            subset=["9th_label_clean", "esto_label_clean"]
        )

        if working.empty:
            print("No rows with both 9th and ESTO labels; skipping mapping cardinality checks.")
            return pd.DataFrame()

        ninth_counts = working["9th_label_clean"].value_counts()
        esto_counts = working["esto_label_clean"].value_counts()
        working["9th_degree"] = working["9th_label_clean"].map(ninth_counts)
        working["esto_degree"] = working["esto_label_clean"].map(esto_counts)

        one_to_one = (working["9th_degree"] == 1) & (working["esto_degree"] == 1)
        one_to_many = (working["9th_degree"] > 1) & (working["esto_degree"] == 1)
        many_to_one = (working["9th_degree"] == 1) & (working["esto_degree"] > 1)
        many_to_many = (working["9th_degree"] > 1) & (working["esto_degree"] > 1)

        print(f"One-to-one mappings: {int(one_to_one.sum())}")
        print(f"One-to-many mappings: {int(one_to_many.sum())}")
        print(f"Many-to-one mappings: {int(many_to_one.sum())}")
        print(f"Many-to-many mappings: {int(many_to_many.sum())}")

        output_cols = [col for col in MAPPING_COLUMNS if col in working.columns]
        many_to_many_df = working.loc[many_to_many, output_cols].copy()
        if len(many_to_many_df):
            with pd.option_context(
                "display.max_rows",
                None,
                "display.max_columns",
                None,
                "display.max_colwidth",
                None,
            ):
                print(many_to_many_df.to_string(index=False))
            if FAIL_ON_MANY_TO_MANY:
                raise ValueError("Many-to-many mappings detected.")
        else:
            print("No many-to-many mappings detected.")
        return many_to_many_df
    except Exception as exc:
        print(f"Mapping cardinality checks failed: {exc}")
        try_debug_breakpoint()
        raise


def collect_9th_labels(ninth_df: pd.DataFrame) -> pd.Series:
    """Collect unique 9th labels from all hierarchy columns.

    Inputs:
        ninth_df: Dataframe for the 9th sheet.
    Outputs:
        Series of unique labels.
    Side effects:
        None.
    """
    columns = [col for col in NINTH_LABEL_COLUMNS if col in ninth_df.columns]
    if not columns:
        return pd.Series([], dtype=str)
    combined = pd.concat([ninth_df[col] for col in columns], ignore_index=True)
    combined = combined.fillna("").astype(str).str.strip()
    combined = combined[(combined != "") & (combined.str.lower() != "x")]
    return combined.drop_duplicates().reset_index(drop=True)


def collect_esto_labels(esto_df: pd.DataFrame) -> pd.Series:
    """Collect unique ESTO labels from full code or flows/products columns.

    Inputs:
        esto_df: Dataframe for the ESTO sheet.
    Outputs:
        Series of unique labels.
    Side effects:
        None.
    """
    if ESTO_LABEL_COLUMN in esto_df.columns:
        labels = esto_df[ESTO_LABEL_COLUMN].fillna("").astype(str).str.strip()
        labels = labels[labels != ""]
        return labels.drop_duplicates().reset_index(drop=True)

    available = [col for col in ESTO_COLUMN_CANDIDATES if col in esto_df.columns]
    if not available:
        return pd.Series([], dtype=str)

    combined = pd.concat([esto_df[col] for col in available], ignore_index=True)
    combined = combined.fillna("").astype(str).str.strip()
    combined = combined[combined != ""]
    return combined.drop_duplicates().reset_index(drop=True)


def build_label_column_mapping(
    reference_df: pd.DataFrame,
    candidate_columns: list[str],
    ignore_values: Optional[set[str]] = None,
) -> Dict[str, set[str]]:
    """Build a mapping of label -> set(columns) from the reference sheet.

    Inputs:
        reference_df: Reference dataframe to scan for labels.
        candidate_columns: Columns to scan for labels.
        ignore_values: Optional set of lowercase labels to ignore.
    Outputs:
        Dict of label -> set of columns where it appears.
    Side effects:
        None.
    """
    mapping: Dict[str, set[str]] = {}
    ignore_values = ignore_values or set()
    for column in candidate_columns:
        if column not in reference_df.columns:
            continue
        labels = reference_df[column].fillna("").astype(str).str.strip()
        for label in labels:
            if label == "":
                continue
            if label.lower() in ignore_values:
                continue
            mapping.setdefault(label, set()).add(column)
    return mapping


def resolve_label_column(
    label: str,
    label_mapping: Dict[str, set[str]],
    preferred_order: list[str],
) -> Tuple[str, list[str]]:
    """Resolve the preferred column for a label from a mapping.

    Inputs:
        label: Label to resolve.
        label_mapping: Mapping of label -> set(columns).
        preferred_order: Column order preference.
    Outputs:
        Tuple of (preferred_column, candidate_columns).
    Side effects:
        None.
    """
    candidates = sorted(label_mapping.get(label, set()))
    if not candidates:
        return "", []
    ordered = [col for col in preferred_order if col in candidates]
    preferred = ordered[0] if ordered else candidates[0]
    return preferred, candidates


def align_label_columns_with_reference(
    code_to_name_df: pd.DataFrame,
    ninth_df: pd.DataFrame,
    esto_df: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Align 9th/ESTO column fields with reference sheet locations.

    Inputs:
        code_to_name_df: Dataframe with 9th/ESTO labels and columns.
        ninth_df: Reference 9th sheet dataframe.
        esto_df: Reference ESTO sheet dataframe.
    Outputs:
        Tuple of (updated dataframe, audit dataframe).
    Side effects:
        None.
    """
    try:
        updated_df = code_to_name_df.copy()
        audit_rows = []

        ninth_mapping = build_label_column_mapping(
            ninth_df, NINTH_LABEL_COLUMNS, ignore_values={"x"}
        )
        esto_mapping = build_label_column_mapping(esto_df, ESTO_COLUMN_CANDIDATES)

        for idx, row in updated_df.iterrows():
            ninth_label = (row.get("9th_label") or "").strip()
            esto_label = (row.get("esto_label") or "").strip()

            if ninth_label:
                expected_column, candidates = resolve_label_column(
                    ninth_label, ninth_mapping, NINTH_LABEL_COLUMNS
                )
                current_column = (row.get("9th_column") or "").strip()
                if expected_column and current_column not in candidates:
                    updated_df.at[idx, "9th_column"] = expected_column
                    audit_rows.append(
                        {
                            "name": row.get("name", ""),
                            "label_type": "9th",
                            "label": ninth_label,
                            "old_column": current_column,
                            "new_column": expected_column,
                            "candidate_columns": ", ".join(candidates),
                        }
                    )

            if esto_label:
                expected_column, candidates = resolve_label_column(
                    esto_label, esto_mapping, ESTO_COLUMN_CANDIDATES
                )
                current_column = (row.get("esto_column") or "").strip()
                if expected_column and current_column not in candidates:
                    updated_df.at[idx, "esto_column"] = expected_column
                    audit_rows.append(
                        {
                            "name": row.get("name", ""),
                            "label_type": "ESTO",
                            "label": esto_label,
                            "old_column": current_column,
                            "new_column": expected_column,
                            "candidate_columns": ", ".join(candidates),
                        }
                    )

        audit_df = pd.DataFrame(audit_rows)
        return updated_df, audit_df
    except Exception as exc:
        print(f"Failed to align label columns: {exc}")
        try_debug_breakpoint()
        raise


def run_label_column_alignment_checks(
    workbook_path: Path, code_to_name_df: pd.DataFrame
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Run alignment checks against the 9th and ESTO reference sheets.

    Inputs:
        workbook_path: Path to the Excel workbook.
        code_to_name_df: Dataframe with code_to_name labels.
    Outputs:
        Tuple of (updated dataframe, audit dataframe).
    Side effects:
        Reads reference sheets and prints diagnostics.
    """
    try:
        ninth_df = load_reference_sheet(workbook_path, NINTH_SHEET_NAME)
        esto_df = load_reference_sheet(workbook_path, ESTO_SHEET_NAME)
        updated_df, audit_df = align_label_columns_with_reference(
            code_to_name_df, ninth_df, esto_df
        )

        print(f"Label column fixes applied: {len(audit_df)}")
        if len(audit_df):
            with pd.option_context(
                "display.max_rows",
                None,
                "display.max_columns",
                None,
                "display.max_colwidth",
                None,
            ):
                print(audit_df.to_string(index=False))
        else:
            print("No label column fixes needed.")
        return updated_df, audit_df
    except Exception as exc:
        print(f"Label column alignment checks failed: {exc}")
        try_debug_breakpoint()
        raise


def report_missing_reference_labels(
    code_to_name_df: pd.DataFrame,
    reference_labels: pd.Series,
    label_column: str,
    reference_name: str,
) -> pd.DataFrame:
    """Report labels missing from code_to_name compared to a reference set.

    Inputs:
        code_to_name_df: Dataframe with code_to_name labels.
        reference_labels: Series of labels from the reference sheet.
        label_column: Column in code_to_name to compare.
        reference_name: Label for reporting.
    Outputs:
        Dataframe of missing labels.
    Side effects:
        None.
    """
    if label_column not in code_to_name_df.columns:
        return pd.DataFrame()
    existing = code_to_name_df[label_column].fillna("").astype(str).str.strip()
    existing = existing[existing != ""]
    missing = sorted(set(reference_labels) - set(existing))
    return pd.DataFrame({label_column: missing, "reference": reference_name})


def run_reference_label_checks(workbook_path: Path, code_to_name_df: pd.DataFrame) -> None:
    """Check for reference labels missing from code_to_name.

    Inputs:
        workbook_path: Path to the Excel workbook.
        code_to_name_df: Dataframe with code_to_name labels.
    Outputs:
        None.
    Side effects:
        Reads reference sheets and prints diagnostics.
    """
    try:
        ninth_df = load_reference_sheet(workbook_path, NINTH_SHEET_NAME)
        esto_df = load_reference_sheet(workbook_path, ESTO_SHEET_NAME)

        ninth_labels = collect_9th_labels(ninth_df)
        esto_labels = collect_esto_labels(esto_df)

        ninth_in_mapping = (
            code_to_name_df.get("9th_label", pd.Series([], dtype=str))
            .fillna("")
            .astype(str)
            .str.strip()
        )
        ninth_in_mapping = ninth_in_mapping[ninth_in_mapping != ""].drop_duplicates()
        esto_in_mapping = (
            code_to_name_df.get("esto_label", pd.Series([], dtype=str))
            .fillna("")
            .astype(str)
            .str.strip()
        )
        esto_in_mapping = esto_in_mapping[esto_in_mapping != ""].drop_duplicates()

        print(f"Unique 9th labels (reference): {len(ninth_labels)}")
        print(f"Unique 9th labels (code_to_name): {len(ninth_in_mapping)}")
        print(f"Unique ESTO labels (reference): {len(esto_labels)}")
        print(f"Unique ESTO labels (code_to_name): {len(esto_in_mapping)}")

        missing_9th = report_missing_reference_labels(
            code_to_name_df, ninth_labels, "9th_label", "9th"
        )
        missing_esto = report_missing_reference_labels(
            code_to_name_df, esto_labels, "esto_label", "ESTO"
        )

        print(f"Missing 9th labels from code_to_name: {len(missing_9th)}")
        if len(missing_9th):
            print(missing_9th.to_string(index=False))

        print(f"Missing ESTO labels from code_to_name: {len(missing_esto)}")
        if len(missing_esto):
            print(missing_esto.to_string(index=False))

        if (len(missing_9th) or len(missing_esto)) and FAIL_ON_MISSING_REFERENCE_LABELS:
            raise ValueError("Missing reference labels detected.")
        return missing_9th, missing_esto
    except Exception as exc:
        print(f"Reference label checks failed: {exc}")
        try_debug_breakpoint()
        raise
#%%


def run_column_source_consistency_checks(code_to_name_df: pd.DataFrame) -> pd.DataFrame:
    """Check for 9th/ESTO column source mismatches in code_to_name.

    Inputs:
        code_to_name_df: Dataframe with 9th/ESTO column selections.
    Outputs:
        Dataframe of mismatched rows.
    Side effects:
        Prints diagnostics.
    """
    try:
        if "9th_column" not in code_to_name_df.columns or "esto_column" not in code_to_name_df.columns:
            print("Missing 9th_column or esto_column for source consistency checks.")
            return pd.DataFrame()

        working = code_to_name_df.copy()
        working["9th_column_clean"] = working["9th_column"].fillna("").astype(str).str.strip().str.lower()
        working["esto_column_clean"] = working["esto_column"].fillna("").astype(str).str.strip().str.lower()

        sector_set = {col.lower() for col in NINTH_SECTOR_COLUMNS}
        fuel_set = {col.lower() for col in NINTH_FUEL_COLUMNS}

        mismatches = []
        for _, row in working.iterrows():
            ninth_col = row["9th_column_clean"]
            esto_col = row["esto_column_clean"]
            if ninth_col == "" or esto_col == "":
                continue

            if esto_col == "flows" and ninth_col not in sector_set:
                mismatches.append(row)
                continue
            if esto_col == "products" and ninth_col not in fuel_set:
                mismatches.append(row)
                continue

        if not mismatches:
            print("No 9th/ESTO column source mismatches found.")
            return pd.DataFrame()

        mismatch_df = pd.DataFrame(mismatches)
        output_cols = [col for col in MAPPING_COLUMNS if col in mismatch_df.columns]
        mismatch_df = mismatch_df.loc[:, output_cols]

        print(f"9th/ESTO column source mismatches: {len(mismatch_df)}")
        with pd.option_context(
            "display.max_rows",
            None,
            "display.max_columns",
            None,
            "display.max_colwidth",
            None,
        ):
            print(mismatch_df.to_string(index=False))
        return mismatch_df
    except Exception as exc:
        print(f"Source consistency checks failed: {exc}")
        try_debug_breakpoint()
        raise


def run_esto_label_column_consistency_checks(
    workbook_path: Path, code_to_name_df: pd.DataFrame
) -> pd.DataFrame:
    """Check that ESTO labels match their source column (flows vs products).

    Inputs:
        workbook_path: Path to the Excel workbook.
        code_to_name_df: Dataframe with esto_label/esto_column.
    Outputs:
        Dataframe of mismatched rows.
    Side effects:
        Reads the ESTO sheet and prints diagnostics.
    """
    try:
        if "esto_label" not in code_to_name_df.columns or "esto_column" not in code_to_name_df.columns:
            print("Missing esto_label or esto_column for ESTO label consistency checks.")
            return pd.DataFrame()

        esto_df = load_reference_sheet(workbook_path, ESTO_SHEET_NAME)
        esto_mapping = build_label_column_mapping(esto_df, ESTO_COLUMN_CANDIDATES)

        mismatches = []
        for _, row in code_to_name_df.iterrows():
            esto_label = (row.get("esto_label") or "").strip()
            esto_column = (row.get("esto_column") or "").strip().lower()
            if esto_label == "" or esto_column == "":
                continue

            candidates = sorted(esto_mapping.get(esto_label, set()))
            if not candidates:
                continue
            if esto_column not in [col.lower() for col in candidates]:
                mismatches.append(
                    {
                        "name": row.get("name", ""),
                        "esto_label": esto_label,
                        "esto_column": row.get("esto_column", ""),
                        "expected_columns": ", ".join(candidates),
                    }
                )

        if not mismatches:
            print("No ESTO label/column mismatches found.")
            return pd.DataFrame()

        mismatch_df = pd.DataFrame(mismatches)
        print(f"ESTO label/column mismatches: {len(mismatch_df)}")
        with pd.option_context(
            "display.max_rows",
            None,
            "display.max_columns",
            None,
            "display.max_colwidth",
            None,
        ):
            print(mismatch_df.to_string(index=False))
        return mismatch_df
    except Exception as exc:
        print(f"ESTO label consistency checks failed: {exc}")
        try_debug_breakpoint()
        raise

#%%
######### CONSTANTS (LIKELY TO CHANGE) #########
WORKBOOK_PATH = Path(r"C:\Users\Work\github\leap_utilities\config\sector_fuel_codes_to_names.xlsx")
SOURCE_SHEET_NAME = "code_to_name"
NINTH_SHEET_NAME = "9th"
ESTO_SHEET_NAME = "ESTO"
OUTPUT_SHEET_NAME = "code_to_name_fixed"
AUDIT_SHEET_NAME = "code_to_name_name_audit"
RUN_NAME_FIXES = False
RUN_VALIDATION_CHECKS = True
RUN_REFERENCE_LABEL_CHECKS = True
RUN_MAPPING_CARDINALITY_CHECKS = True
RUN_LABEL_COLUMN_ALIGNMENT = True
RUN_COLUMN_SOURCE_CONSISTENCY_CHECKS = True
RUN_ESTO_LABEL_COLUMN_CONSISTENCY_CHECKS = True
COLUMN_ALIGNMENT_SHEET_NAME = "code_to_name_column_fixed"
COLUMN_ALIGNMENT_AUDIT_SHEET_NAME = "code_to_name_column_audit"
APPLY_LABEL_COLUMN_FIXES = False
APPLY_NAME_FIXES = False
FAIL_ON_MISSING_LABELS = False
FAIL_ON_DUPLICATES = False
FAIL_ON_MISSING_REFERENCE_LABELS = False
FAIL_ON_MANY_TO_MANY = False
#%%

#%%
######### RUN VALIDATION CHECKS (TOGGLE) #########
if RUN_VALIDATION_CHECKS:
    try:
        source_df = load_code_to_name_sheet(WORKBOOK_PATH, SOURCE_SHEET_NAME)
        prepared_df = add_label_columns_for_fixes(source_df)
        run_validation_checks(prepared_df)
    except Exception as exc:
        print(f"Failed to run validation checks: {exc}")
        try_debug_breakpoint()
#%%

#%%
######### RUN REFERENCE LABEL CHECKS (TOGGLE) #########
if RUN_REFERENCE_LABEL_CHECKS:
    try:
        source_df = load_code_to_name_sheet(WORKBOOK_PATH, SOURCE_SHEET_NAME)
        prepared_df = add_label_columns_for_fixes(source_df)

        missing_9th, missing_esto = run_reference_label_checks(WORKBOOK_PATH, prepared_df)
    except Exception as exc:
        print(f"Failed to run reference label checks: {exc}")
        try_debug_breakpoint()
#%%

#%%
######### RUN MAPPING CARDINALITY CHECKS (TOGGLE) #########
if RUN_MAPPING_CARDINALITY_CHECKS:
    try:
        source_df = load_code_to_name_sheet(WORKBOOK_PATH, SOURCE_SHEET_NAME)
        prepared_df = add_label_columns_for_fixes(source_df)
        run_mapping_cardinality_checks(prepared_df)
    except Exception as exc:
        print(f"Failed to run mapping cardinality checks: {exc}")
        try_debug_breakpoint()
#%%

#%%
######### RUN ESTO LABEL/COLUMN CONSISTENCY CHECKS (TOGGLE) #########
if RUN_ESTO_LABEL_COLUMN_CONSISTENCY_CHECKS:
    try:
        source_df = load_code_to_name_sheet(WORKBOOK_PATH, SOURCE_SHEET_NAME)
        prepared_df = add_label_columns_for_fixes(source_df)
        run_esto_label_column_consistency_checks(WORKBOOK_PATH, prepared_df)
    except Exception as exc:
        print(f"Failed to run ESTO label/column consistency checks: {exc}")
        try_debug_breakpoint()
#%%

#%%
######### RUN COLUMN SOURCE CONSISTENCY CHECKS (TOGGLE) #########
if RUN_COLUMN_SOURCE_CONSISTENCY_CHECKS:
    try:
        source_df = load_code_to_name_sheet(WORKBOOK_PATH, SOURCE_SHEET_NAME)
        prepared_df = add_label_columns_for_fixes(source_df)
        run_column_source_consistency_checks(prepared_df)
    except Exception as exc:
        print(f"Failed to run column source consistency checks: {exc}")
        try_debug_breakpoint()
#%%

#%%
######### RUN LABEL COLUMN ALIGNMENT (TOGGLE) #########
if RUN_LABEL_COLUMN_ALIGNMENT:
    try:
        source_df = load_code_to_name_sheet(WORKBOOK_PATH, SOURCE_SHEET_NAME)
        prepared_df = add_label_columns_for_fixes(source_df)
        aligned_df, audit_df = run_label_column_alignment_checks(WORKBOOK_PATH, prepared_df)

        if APPLY_LABEL_COLUMN_FIXES:
            try:
                write_excel_sheet(WORKBOOK_PATH, aligned_df, COLUMN_ALIGNMENT_SHEET_NAME)
                write_excel_sheet(WORKBOOK_PATH, audit_df, COLUMN_ALIGNMENT_AUDIT_SHEET_NAME)
            except Exception as write_exc:
                print(f"Failed to write label column alignment outputs: {write_exc}")
        else:
            print("Label column fixes computed but not written (APPLY_LABEL_COLUMN_FIXES=False).")
    except Exception as exc:
        print(f"Failed to run label column alignment: {exc}")
        try_debug_breakpoint()
#%%

#%%
######### RUN NAME FIXES (TOGGLE) #########
if RUN_NAME_FIXES:
    try:
        source_df = load_code_to_name_sheet(WORKBOOK_PATH, SOURCE_SHEET_NAME)
        prepared_df = add_label_columns_for_fixes(source_df)
        fixed_df, audit_df = apply_name_fixes(prepared_df)

        if APPLY_NAME_FIXES:
            write_excel_sheet(WORKBOOK_PATH, fixed_df, OUTPUT_SHEET_NAME)
            write_excel_sheet(WORKBOOK_PATH, audit_df, AUDIT_SHEET_NAME)
        else:
            print("Name fixes computed but not written (APPLY_NAME_FIXES=False).")
        if len(audit_df):
            with pd.option_context(
                "display.max_rows",
                None,
                "display.max_columns",
                None,
                "display.max_colwidth",
                None,
            ):
                print(audit_df.to_string(index=False))
        print(f"Rows changed: {len(audit_df)}")
    except Exception as exc:
        print(f"Failed to apply name fixes: {exc}")
        try_debug_breakpoint()
#%%
# a = pd.read_csv("../../data/00APEC_2024_low.csv")[['flows', 'products']].drop_duplicates().reset_index(drop=True)
# new = pd.DataFrame()
# #make legnth of new 100
# new = new.reindex(range(0,150))
# a = pd.read_csv("../../data/merged_file_energy_ALL_20250814.csv")
# for col in ['sectors','sub1sectors','sub2sectors','sub3sectors','sub4sectors','fuels','subfuels']:
#     b = a[[col]].drop_duplicates().copy()
#     b = b[b[col].notna() & (b[col] != "")]
#     new[col] = b[col].reset_index(drop=True)
# #drop nas in longest col
# new = new.dropna(subset=['sub2sectors'])
