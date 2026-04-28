from __future__ import annotations

from pathlib import Path

import pandas as pd


def _load_csv(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    return pd.read_csv(path)


def compare_outputs(
    *,
    v1_output_dir: Path,
    v2_output_dir: Path,
    out_path: Path,
    value_tolerance: float = 1e-6,
) -> Path:
    v1_long = _load_csv(v1_output_dir / "comparison_long.csv")
    v2_long = _load_csv(v2_output_dir / "comparison_long.csv")

    rows: list[dict[str, object]] = []
    rows.append({"metric": "v1_rows", "value": int(len(v1_long))})
    rows.append({"metric": "v2_rows", "value": int(len(v2_long))})
    rows.append({"metric": "v1_columns", "value": "|".join(sorted(v1_long.columns.astype(str).tolist()))})
    rows.append({"metric": "v2_columns", "value": "|".join(sorted(v2_long.columns.astype(str).tolist()))})

    if not v1_long.empty and not v2_long.empty:
        key_cols = ["sheet", "fuel_label", "scenario", "source", "year"]
        for col in key_cols:
            if col in v1_long.columns:
                v1_long[col] = v1_long[col].astype(str)
            if col in v2_long.columns:
                v2_long[col] = v2_long[col].astype(str)

        if all(c in v1_long.columns for c in key_cols) and all(c in v2_long.columns for c in key_cols):
            merged = v1_long[key_cols + ["value"]].rename(columns={"value": "value_v1"}).merge(
                v2_long[key_cols + ["value"]].rename(columns={"value": "value_v2"}),
                on=key_cols,
                how="outer",
            )
            merged["value_v1"] = pd.to_numeric(merged["value_v1"], errors="coerce")
            merged["value_v2"] = pd.to_numeric(merged["value_v2"], errors="coerce")
            merged["abs_diff"] = (merged["value_v1"] - merged["value_v2"]).abs()
            diff_rows = merged[(merged["abs_diff"] > value_tolerance) | (merged["value_v1"].isna() ^ merged["value_v2"].isna())]
            rows.append({"metric": "value_diff_rows", "value": int(len(diff_rows))})
            if not diff_rows.empty:
                detail = out_path.with_name("shadow_compare_value_diffs.csv")
                diff_rows.to_csv(detail, index=False)
                rows.append({"metric": "value_diff_detail", "value": str(detail)})

    out_path.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows).to_csv(out_path, index=False)
    return out_path
