#%%
from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Any

import pandas as pd

from codebase.utilities.master_config import read_config_table
from codebase.utilities.workflow_outputs import write_output_manifest


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


#%%
def _resolve(path: Path | str) -> Path:
    raw = str(path).replace("\\", "/")
    drive_match = re.match(r"^([a-zA-Z]):/(.*)$", raw)
    if drive_match:
        drive = drive_match.group(1).lower()
        rest = drive_match.group(2)
        return Path(f"/mnt/{drive}/{rest}")
    candidate = Path(raw)
    return candidate if candidate.is_absolute() else (REPO_ROOT / candidate)


def _coalesce_unique(series: pd.Series) -> str:
    cleaned = [str(v).strip() for v in series.fillna("").astype(str).tolist() if str(v).strip()]
    unique = sorted(set(cleaned))
    return unique[0] if unique else ""


def _build_one_side_mapping(
    *,
    frame: pd.DataFrame,
    left_col: str,
    right_col: str,
    value_col: str,
    left_label: str,
    right_label: str,
) -> pd.DataFrame:
    use = frame.copy()
    use[left_col] = use.get(left_col, "").fillna("").astype(str).str.strip()
    use[right_col] = use.get(right_col, "").fillna("").astype(str).str.strip()
    use[value_col] = pd.to_numeric(use.get(value_col, pd.NA), errors="coerce")
    use = use[use[left_col].ne("") & use[right_col].ne("")].copy()

    if use.empty:
        return pd.DataFrame(
            columns=[
                left_col,
                right_col,
                "row_count",
                "value_pj_abs_sum",
                "right_count_for_left",
                "left_count_for_right",
                "mapping_cardinality",
                f"{left_label}_confidence",
                "source",
            ]
        )

    pairs = (
        use[[left_col, right_col, value_col]]
        .groupby([left_col, right_col], as_index=False)
        .agg(
            row_count=(left_col, "size"),
            value_pj_abs_sum=(value_col, lambda s: s.abs().sum()),
        )
    )
    left_counts = pairs.groupby(left_col, as_index=False).size().rename(columns={"size": "right_count_for_left"})
    right_counts = pairs.groupby(right_col, as_index=False).size().rename(columns={"size": "left_count_for_right"})
    out = pairs.merge(left_counts, on=left_col, how="left").merge(right_counts, on=right_col, how="left")

    out["mapping_cardinality"] = "many_to_many"
    one_to_one = out["right_count_for_left"].eq(1) & out["left_count_for_right"].eq(1)
    one_to_many = out["right_count_for_left"].gt(1) & out["left_count_for_right"].eq(1)
    many_to_one = out["right_count_for_left"].eq(1) & out["left_count_for_right"].gt(1)
    out.loc[one_to_one, "mapping_cardinality"] = "one_to_one"
    out.loc[one_to_many, "mapping_cardinality"] = "one_to_many"
    out.loc[many_to_one, "mapping_cardinality"] = "many_to_one"

    out[f"{left_label}_confidence"] = "review"
    out.loc[out["mapping_cardinality"].eq("one_to_one"), f"{left_label}_confidence"] = "high"
    out.loc[out["mapping_cardinality"].isin(["many_to_one", "one_to_many"]), f"{left_label}_confidence"] = "medium"
    out["source"] = "existing_pair_mappings"

    return out.sort_values([left_col, "value_pj_abs_sum"], ascending=[True, False], kind="mergesort").reset_index(drop=True)


def _build_final_choice_sheet(
    *,
    mapping_df: pd.DataFrame,
    left_col: str,
    right_col: str,
    confidence_col: str,
) -> pd.DataFrame:
    if mapping_df.empty:
        return pd.DataFrame(
            columns=[
                left_col,
                right_col,
                "mapping_cardinality",
                "value_pj_abs_sum",
                confidence_col,
                "selection_method",
                "manual_override",
                "review_notes",
            ]
        )

    work = mapping_df.copy()
    priority = {"one_to_one": 0, "many_to_one": 1, "one_to_many": 2, "many_to_many": 3}
    work["_priority"] = work["mapping_cardinality"].map(priority).fillna(9).astype(int)
    work["_value_rank"] = pd.to_numeric(work["value_pj_abs_sum"], errors="coerce").fillna(0.0)
    work = work.sort_values([left_col, "_priority", "_value_rank"], ascending=[True, True, False], kind="mergesort")
    chosen = work.groupby(left_col, as_index=False).head(1).copy()
    chosen["selection_method"] = "auto_best_existing_mapping"
    chosen["manual_override"] = ""
    chosen["review_notes"] = ""
    chosen = chosen[
        [
            left_col,
            right_col,
            "mapping_cardinality",
            "value_pj_abs_sum",
            confidence_col,
            "selection_method",
            "manual_override",
            "review_notes",
        ]
    ].reset_index(drop=True)
    return chosen


def run_workflow(
    *,
    esto_mapping_workbook: Path | str = _resolve("config/leap_to_esto_balance_full_mapping_slim.xlsx"),
    ninth_mapping_workbook: Path | str = _resolve("config/leap_to_ninth_balance_full_mapping_slim.xlsx"),
    output_workbook: Path | str = _resolve("config/leap_individual_mappings_from_existing.xlsx"),
) -> dict[str, Any]:
    esto_path = _resolve(esto_mapping_workbook)
    ninth_path = _resolve(ninth_mapping_workbook)
    out_path = _resolve(output_workbook)

    esto_name = read_config_table(esto_path, sheet_name="leap_name_to_esto_pair").fillna("")
    ninth_name = read_config_table(ninth_path, sheet_name="leap_name_to_ninth_pair").fillna("")

    for df in [esto_name, ninth_name]:
        if "value_pj_abs_sum" not in df.columns:
            df["value_pj_abs_sum"] = pd.NA

    leap_sector_to_esto_flow = _build_one_side_mapping(
        frame=esto_name,
        left_col="leap_sector_name",
        right_col="esto_flow",
        value_col="value_pj_abs_sum",
        left_label="sector",
        right_label="flow",
    )
    leap_fuel_to_esto_product = _build_one_side_mapping(
        frame=esto_name,
        left_col="leap_fuel_name",
        right_col="esto_product",
        value_col="value_pj_abs_sum",
        left_label="fuel",
        right_label="product",
    )
    leap_sector_to_ninth_sector = _build_one_side_mapping(
        frame=ninth_name,
        left_col="leap_sector_name",
        right_col="ninth_sector",
        value_col="value_pj_abs_sum",
        left_label="sector",
        right_label="ninth_sector",
    )
    leap_fuel_to_ninth_fuel = _build_one_side_mapping(
        frame=ninth_name,
        left_col="leap_fuel_name",
        right_col="ninth_fuel",
        value_col="value_pj_abs_sum",
        left_label="fuel",
        right_label="ninth_fuel",
    )

    sector_flow_final = _build_final_choice_sheet(
        mapping_df=leap_sector_to_esto_flow,
        left_col="leap_sector_name",
        right_col="esto_flow",
        confidence_col="sector_confidence",
    )
    fuel_product_final = _build_final_choice_sheet(
        mapping_df=leap_fuel_to_esto_product,
        left_col="leap_fuel_name",
        right_col="esto_product",
        confidence_col="fuel_confidence",
    )
    sector_ninth_final = _build_final_choice_sheet(
        mapping_df=leap_sector_to_ninth_sector,
        left_col="leap_sector_name",
        right_col="ninth_sector",
        confidence_col="sector_confidence",
    )
    fuel_ninth_final = _build_final_choice_sheet(
        mapping_df=leap_fuel_to_ninth_fuel,
        left_col="leap_fuel_name",
        right_col="ninth_fuel",
        confidence_col="fuel_confidence",
    )

    summary = pd.DataFrame(
        [
            {
                "leap_sector_to_esto_flow_rows": len(leap_sector_to_esto_flow),
                "leap_fuel_to_esto_product_rows": len(leap_fuel_to_esto_product),
                "leap_sector_to_ninth_sector_rows": len(leap_sector_to_ninth_sector),
                "leap_fuel_to_ninth_fuel_rows": len(leap_fuel_to_ninth_fuel),
                "sector_to_esto_one_to_one": int((leap_sector_to_esto_flow["mapping_cardinality"] == "one_to_one").sum()),
                "fuel_to_esto_one_to_one": int((leap_fuel_to_esto_product["mapping_cardinality"] == "one_to_one").sum()),
                "sector_flow_final_rows": len(sector_flow_final),
                "fuel_product_final_rows": len(fuel_product_final),
                "sector_ninth_final_rows": len(sector_ninth_final),
                "fuel_ninth_final_rows": len(fuel_ninth_final),
            }
        ]
    )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        leap_sector_to_esto_flow.to_excel(writer, sheet_name="leap_sector_to_esto_flow", index=False)
        leap_fuel_to_esto_product.to_excel(writer, sheet_name="leap_fuel_to_esto_product", index=False)
        sector_flow_final.to_excel(writer, sheet_name="sector_flow_final", index=False)
        fuel_product_final.to_excel(writer, sheet_name="fuel_product_final", index=False)
        leap_sector_to_ninth_sector.to_excel(writer, sheet_name="leap_sector_to_ninth_sector", index=False)
        leap_fuel_to_ninth_fuel.to_excel(writer, sheet_name="leap_fuel_to_ninth_fuel", index=False)
        sector_ninth_final.to_excel(writer, sheet_name="sector_ninth_final", index=False)
        fuel_ninth_final.to_excel(writer, sheet_name="fuel_ninth_final", index=False)
    manifest_path = write_output_manifest(
        out_dir=out_path.parent,
        primary_outputs={"output_workbook": str(out_path)},
        supporting_outputs={},
        primary_output_descriptions={
            "output_workbook": "Workbook scaffold of individual LEAP-to-ESTO and LEAP-to-9th mappings.",
        },
        notes=[
            "This workflow produces a single primary workbook.",
        ],
    )

    return {
        "output_workbook": str(out_path),
        "summary": summary.iloc[0].to_dict(),
        "output_manifest_json": str(manifest_path),
    }


#%%
RUN_WORKFLOW = False
WORKFLOW_RESULT: dict[str, Any] | None = None
if RUN_WORKFLOW:
    WORKFLOW_RESULT = run_workflow()
    print("[OK] Individual mapping scaffold workflow complete.")
    print(WORKFLOW_RESULT)
