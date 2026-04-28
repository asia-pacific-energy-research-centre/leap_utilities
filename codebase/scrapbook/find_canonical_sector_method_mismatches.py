from __future__ import annotations

import os
import re
import sys
from pathlib import Path

import pandas as pd

from codebase.utilities.master_config import read_config_table


REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.mappings.canonical_mapping import build_code_match_method, normalize_match_method


DEFAULT_CANONICAL_PAIRS = Path("config/ninth_pairs_to_esto_pairs.xlsx")
DEFAULT_OUTPUT_DIR = Path("config/computer_generated_config/leap_mapping_views")


def ensure_repo_root() -> None:
    cwd = Path.cwd()
    if cwd != REPO_ROOT:
        os.chdir(REPO_ROOT)


def _resolve(path_like: Path | str) -> Path:
    text = str(path_like).replace("\\", "/")
    path = Path(text)
    return path if path.is_absolute() else REPO_ROOT / path


def _extract_ninth_code(label: object) -> str:
    match = re.match(r"^\s*(\d{2}(?:_\d{2})*)", str(label or ""))
    return match.group(1) if match else ""


def _extract_esto_code(label: object) -> str:
    match = re.match(r"^\s*(\d{2}(?:\.\d{2})*)", str(label or ""))
    return match.group(1).replace(".", "_") if match else ""


def _expected_code_method(ninth_sector: object, esto_flow: object) -> str:
    ninth_code = _extract_ninth_code(ninth_sector)
    flow_code = _extract_esto_code(esto_flow)
    if not ninth_code or not flow_code:
        return ""
    ninth_parts = ninth_code.split("_")
    flow_parts = flow_code.split("_")
    if len(flow_parts) > len(ninth_parts):
        return ""
    if ninth_parts[: len(flow_parts)] != flow_parts:
        return ""
    return build_code_match_method(len(ninth_parts) - len(flow_parts))


def build_report(
    canonical_pairs_path: Path | str = DEFAULT_CANONICAL_PAIRS,
    output_dir: Path | str = DEFAULT_OUTPUT_DIR,
) -> tuple[Path, Path]:
    ensure_repo_root()
    pairs_path = _resolve(canonical_pairs_path)
    out_dir = _resolve(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    pairs = read_config_table(pairs_path)
    pairs.columns = [str(col).strip().lower() for col in pairs.columns]
    for col in ["9th_sector", "9th_fuel", "esto_flow", "esto_product", "sector_match_method", "fuel_match_method", "mapping_note"]:
        if col not in pairs.columns:
            pairs[col] = ""
        pairs[col] = pairs[col].fillna("").astype(str).str.strip()
    for col in ["sector_match_method", "fuel_match_method"]:
        pairs[col] = pairs[col].map(normalize_match_method)

    pairs["expected_sector_match_method"] = pairs.apply(
        lambda row: _expected_code_method(row.get("9th_sector", ""), row.get("esto_flow", "")),
        axis=1,
    )

    exact_to_ancestor = pairs[
        pairs["sector_match_method"].eq("direct_code_match")
        & pairs["expected_sector_match_method"].str.startswith("parent_code_match_")
    ].copy()
    if not exact_to_ancestor.empty:
        exact_to_ancestor["issue_type"] = "direct_code_match_should_be_parent_code_match"

    broader_exact_like = pairs[
        pairs["sector_match_method"].isin(
            {
                "direct_code_match",
                "independent_table_direct_match",
                "independent_table_direct_match_x_category",
            }
        )
        & pairs["expected_sector_match_method"].str.startswith("parent_code_match_")
    ].copy()
    if not broader_exact_like.empty:
        broader_exact_like["issue_type"] = "exact_like_method_uses_parent_flow"

    exact_path = out_dir / "sector_match_method_direct_to_parent.csv"
    broader_path = out_dir / "sector_match_method_exact_like_parent_flow.csv"
    exact_to_ancestor.to_csv(exact_path, index=False)
    broader_exact_like.to_csv(broader_path, index=False)
    return exact_path, broader_path


if __name__ == "__main__":
    exact_path, broader_path = build_report()
    print(f"[INFO] Wrote strict direct-code mismatch report: {exact_path}")
    print(f"[INFO] Wrote broader exact-like parent-flow report: {broader_path}")
