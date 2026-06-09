#%%
from __future__ import annotations

import os
import sys
from pathlib import Path

import pandas as pd

from codebase.utilities.master_config import config_table_exists, read_config_table

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from codebase.mappings.canonical_mapping import (  # noqa: E402
    build_flow_sector_crosswalk,
    build_leap_branch_crosswalk,
    build_sector_to_esto_flow_lookup,
    build_product_fuel_crosswalk,
    normalize_label,
    split_sector_codes,
    load_canonical_pairs,
    load_sheet_map,
)


def ensure_repo_root() -> None:
    """Make relative file operations stable in notebooks/scripts."""
    cwd = Path.cwd()
    if cwd != REPO_ROOT:
        os.chdir(REPO_ROOT)


def _resolve_path(path_like: Path | str) -> Path:
    """
    Resolve a path against REPO_ROOT and normalize Windows-style separators.
    """
    text = str(path_like)
    # Allow notebook users to pass Windows-style relative paths like "config\\file.xlsx".
    text = text.replace("\\", "/")
    p = Path(text)
    return p if p.is_absolute() else (REPO_ROOT / p)


def _safe_write_csv(df: pd.DataFrame, path: Path) -> None:
    try:
        df.to_csv(path, index=False)
    except PermissionError:
        print(f"[WARN] Could not write locked file: {path}. Close it and rerun to refresh this artifact.")


def _read_override_table(path: Path) -> pd.DataFrame:
    cols = [
        "leap_fuel_label",
        "ninth_fuel_override",
        "esto_product_override",
        "esto_flow_override",
    ]
    if not config_table_exists(path):
        return pd.DataFrame(columns=cols + ["label_norm"])
    if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        df = read_config_table(path)
    else:
        df = read_config_table(path)
    df.columns = [str(c).strip().lower() for c in df.columns]
    for col in cols:
        if col not in df.columns:
            df[col] = ""
    out = df[cols].copy()
    for col in cols:
        out[col] = out[col].fillna("").astype(str).str.strip()
    out["label_norm"] = out["leap_fuel_label"].map(normalize_label)
    return out


def _build_leap_label_map(codebook_path: Path) -> pd.DataFrame:
    leap_df = read_config_table(codebook_path, sheet_name="ESTO_LEAP_names")
    leap_df = leap_df[leap_df["category"].astype(str).str.strip().str.lower() == "products"].copy()
    leap_df["leap_label"] = leap_df["leap_name"].fillna("").astype(str).str.strip()
    leap_df["leap_label_norm"] = leap_df["leap_label"].map(normalize_label)
    leap_df["esto_product"] = leap_df["original_label"].fillna("").astype(str).str.strip()
    leap_df["source"] = "esto_leap_names"

    code_df = read_config_table(codebook_path, sheet_name="code_to_name")
    code_df["name"] = code_df["name"].fillna("").astype(str).str.strip()
    code_df["name_norm"] = code_df["name"].map(normalize_label)
    code_df["esto_label"] = code_df["esto_label"].fillna("").astype(str).str.strip()
    code_df = code_df[(code_df["name_norm"] != "") & (code_df["esto_label"] != "")]

    base = leap_df[["leap_label", "leap_label_norm", "esto_product", "source"]].copy()
    mapped = set(base["leap_label_norm"].tolist())
    extra = code_df[~code_df["name_norm"].isin(mapped)].copy()
    if not extra.empty:
        extra = extra.rename(
            columns={"name": "leap_label", "name_norm": "leap_label_norm", "esto_label": "esto_product"}
        )
        extra["source"] = "code_to_name_name"
        base = pd.concat([base, extra[["leap_label", "leap_label_norm", "esto_product", "source"]]], ignore_index=True)

    base = base[(base["leap_label_norm"] != "") & (base["esto_product"] != "")]
    base = base.drop_duplicates(subset=["leap_label_norm", "esto_product", "source"])
    return base.sort_values(["leap_label_norm", "source"]).reset_index(drop=True)


def build_views(
    canonical_pairs_path: Path,
    codebook_path: Path,
    sheet_map_path: Path,
    output_dir: Path,
    *, 
    override_path: Path,
    leap_long_path: Path | None,
    projection_table_path: Path | None,
) -> int:
    output_dir.mkdir(parents=True, exist_ok=True)

    canonical_pairs, canonical_conflicts = load_canonical_pairs(canonical_pairs_path, strict=False)
    _safe_write_csv(canonical_pairs, output_dir / "canonical_pairs_clean.csv")
    _safe_write_csv(build_product_fuel_crosswalk(canonical_pairs), output_dir / "canonical_product_fuel_crosswalk.csv")
    _safe_write_csv(build_flow_sector_crosswalk(canonical_pairs), output_dir / "canonical_flow_sector_crosswalk.csv")
    _safe_write_csv(
        build_leap_branch_crosswalk(sheet_map_path=sheet_map_path, codebook_path=codebook_path),
        output_dir / "leap_branch_sector_flow_crosswalk.csv",
    )

    label_map = _build_leap_label_map(codebook_path)
    _safe_write_csv(label_map, output_dir / "leap_label_to_esto_product.csv")

    sheet_map = load_sheet_map(sheet_map_path)
    coverage_rows: list[dict[str, object]] = []
    for _, row in sheet_map.iterrows():
        sheet = str(row.get("sheet_name") or "").strip()
        sectors = split_sector_codes(row.get("sector_code_9th"))
        if not sectors:
            sectors = [str(row.get("sector_code_9th") or "").strip()]
        for sector in sectors:
            matches = canonical_pairs[canonical_pairs["9th_sector"].astype(str).str.lower() == sector.lower()]
            coverage_rows.append(
                {
                    "sheet_name": sheet,
                    "sector_code_9th": sector,
                    "canonical_pairs": int(len(matches)),
                    "distinct_ninth_fuels": int(matches["9th_fuel"].nunique()),
                    "distinct_esto_products": int(matches["esto_product"].nunique()),
                    "covered": bool(not matches.empty),
                }
            )
    coverage_df = pd.DataFrame(coverage_rows)
    _safe_write_csv(coverage_df, output_dir / "sheet_sector_coverage_report.csv")

    overrides = _read_override_table(override_path)
    override_by_label = {
        r.label_norm: {
            "ninth_fuel_override": str(r.ninth_fuel_override or "").strip(),
            "esto_product_override": str(r.esto_product_override or "").strip(),
            "esto_flow_override": str(r.esto_flow_override or "").strip(),
        }
        for r in overrides.itertuples(index=False)
        if str(r.label_norm or "").strip()
    }

    label_to_product = (
        label_map.sort_values("source").drop_duplicates(subset=["leap_label_norm"], keep="first").set_index("leap_label_norm")["esto_product"].to_dict()
    )
    sector_flow_mapping = build_sector_to_esto_flow_lookup(codebook_path)
    product_to_fuels = (
        canonical_pairs.groupby(canonical_pairs["esto_product"].astype(str).str.strip().str.lower())["9th_fuel"]
        .apply(lambda s: sorted(set(v for v in s.astype(str).str.strip() if v)))
        .to_dict()
    )

    sector_projection_fuels: dict[str, set[str]] = {}
    if projection_table_path and projection_table_path.exists():
        proj = read_config_table(projection_table_path)
        years = [str(y) for y in range(2023, 2071) if str(y) in proj.columns]
        if years:
            for y in years:
                proj[y] = pd.to_numeric(proj[y], errors="coerce")
            proj["projection_nonzero"] = proj[years].abs().sum(axis=1) > 0
            proj = proj[proj["projection_nonzero"]]
        sector_cols = [c for c in ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"] if c in proj.columns]
        fuel_cols = [c for c in ["fuels", "subfuels"] if c in proj.columns]
        for _, row in proj.iterrows():
            sectors = [str(row.get(c) or "").strip().lower() for c in sector_cols]
            sectors = [s for s in sectors if s and s != "x"]
            fuels = [str(row.get(c) or "").strip() for c in fuel_cols]
            fuels = [f for f in fuels if f and f.lower() != "x"]
            if not sectors or not fuels:
                continue
            for s in sectors:
                sector_projection_fuels.setdefault(s, set()).update(fuels)

    def _resolve_sector_flow(sectors: list[str]) -> str:
        for sector_code in sectors:
            key = str(sector_code or "").strip().lower()
            if not key:
                continue
            if key in sector_flow_mapping:
                return str(sector_flow_mapping[key]).strip()
            prefix = key + "_"
            candidates = [k for k in sector_flow_mapping if str(k).strip().lower().startswith(prefix)]
            if candidates:
                best = min(candidates, key=len)
                return str(sector_flow_mapping.get(best, "")).strip()
        return ""

    conflict_rows: list[dict[str, object]] = []
    if not canonical_conflicts.empty:
        for _, r in canonical_conflicts.iterrows():
            conflict_rows.append(
                {
                    "issue": str(r.get("issue", "")),
                    "hard_conflict": True,
                    "sheet": "",
                    "fuel_label": "",
                    "9th_sector": str(r.get("9th_sector", "")),
                    "9th_fuel": str(r.get("9th_fuel", "")),
                    "details": str(r.get("details", "")),
                    "suggested_override": "Fix canonical key in ninth_pairs_to_esto_pairs.xlsx",
                }
            )

    unresolved_rows: list[dict[str, object]] = []
    synthetic_rows: list[dict[str, object]] = []
    leap_rows = pd.DataFrame(columns=["sheet_name", "fuel_label"])
    if leap_long_path and leap_long_path.exists():
        leap_long = read_config_table(leap_long_path)
        if {"sheet_name", "fuel_label"}.issubset(leap_long.columns):
            leap_rows = leap_long[["sheet_name", "fuel_label"]].drop_duplicates().copy()

    if not leap_rows.empty:
        canonical_pairs = canonical_pairs.copy()
        canonical_pairs["sector_norm"] = canonical_pairs["9th_sector"].astype(str).str.strip().str.lower()
        canonical_pairs["fuel_norm"] = canonical_pairs["9th_fuel"].astype(str).str.strip().str.lower()
        canonical_pairs["esto_product_norm"] = canonical_pairs["esto_product"].astype(str).str.strip().str.lower()
        global_unique_fuel_by_product: dict[str, str] = {}
        global_unique_flow_by_product: dict[str, str] = {}
        for product, g in canonical_pairs.groupby("esto_product_norm", dropna=False):
            fuels = sorted(g["9th_fuel"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())
            flows = sorted(g["esto_flow"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())
            if len(fuels) == 1:
                global_unique_fuel_by_product[str(product)] = fuels[0]
            if len(flows) == 1:
                global_unique_flow_by_product[str(product)] = flows[0]

        sheet_to_sectors = {
            str(r.sheet_name): (split_sector_codes(r.sector_code_9th) or [str(r.sector_code_9th or "").strip()])
            for r in sheet_map.itertuples(index=False)
        }

        for r in leap_rows.itertuples(index=False):
            sheet = str(r.sheet_name)
            fuel_label = str(r.fuel_label)
            key = normalize_label(fuel_label)
            sectors = sheet_to_sectors.get(sheet, [])
            sector_keys = {s.strip().lower() for s in sectors if str(s).strip()}
            ov = override_by_label.get(key, {})
            ninth_override = str(ov.get("ninth_fuel_override") or "").strip()
            product_override = str(ov.get("esto_product_override") or "").strip()
            esto_product = product_override or label_to_product.get(key, "")
            candidates = canonical_pairs[
                canonical_pairs["sector_norm"].isin(sector_keys)
                & (canonical_pairs["esto_product_norm"] == str(esto_product).strip().lower())
            ]
            if candidates.empty:
                child_matches = canonical_pairs[
                    canonical_pairs["sector_norm"].map(
                        lambda s: any(str(s).startswith(sk + "_") for sk in sector_keys)
                    )
                    & (canonical_pairs["esto_product_norm"] == str(esto_product).strip().lower())
                ]
                candidates = child_matches
            candidate_fuels = sorted(candidates["9th_fuel"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())

            issue = ""
            hard = False
            if ninth_override:
                continue
            if not esto_product:
                issue = "missing_esto_product_for_label"
            elif len(candidate_fuels) == 0:
                if not _resolve_sector_flow(sectors):
                    issue = "no_canonical_match_for_sector_product"
            elif len(candidate_fuels) > 1:
                issue = "many_to_many_ambiguity_without_override"
                hard = True

            if issue:
                product_key = str(esto_product).strip().lower()
                suggested_fuel = candidate_fuels[0] if len(candidate_fuels) == 1 else global_unique_fuel_by_product.get(product_key, "")
                suggested_flow = ""
                if suggested_fuel:
                    sf = candidates[candidates["9th_fuel"] == suggested_fuel]["esto_flow"]
                    suggested_flow = str(sf.iloc[0]) if not sf.empty else ""
                if not suggested_flow:
                    suggested_flow = global_unique_flow_by_product.get(product_key, "")

                issue_detail = ""
                recommended_action = ""
                if issue == "missing_esto_product_for_label":
                    issue_detail = (
                        "LEAP label is not harmonized to any ESTO product in codebook mappings, "
                        "so canonical sector+product lookup cannot start."
                    )
                    recommended_action = (
                        "Add leap_fuel_label + esto_product_override (and optionally ninth_fuel_override) "
                        "in backup_leap_mappings.xlsx."
                    )
                elif issue == "no_canonical_match_for_sector_product":
                    issue_detail = (
                        "Label resolves to an ESTO product, but canonical pairs contain no row for this "
                        "sheet sector + ESTO product combination."
                    )
                    recommended_action = (
                        "Either add canonical pair(s) in ninth_pairs_to_esto_pairs.xlsx, or add explicit override."
                    )
                elif issue == "many_to_many_ambiguity_without_override":
                    issue_detail = (
                        "Multiple canonical 9th fuels match this sector + ESTO product; deterministic choice "
                        "requires explicit rule/override."
                    )
                    recommended_action = "Add ninth_fuel_override in backup_leap_mappings.xlsx or disambiguation rule."

                # Suggest projection-only synthetic canonical pair when sector+product has a unique
                # projection fuel candidate (e.g., zero in ESTO base year but present in projections).
                synthetic_sector = ""
                synthetic_fuel = ""
                projection_support = False
                if issue == "no_canonical_match_for_sector_product" and product_key:
                    global_fuels = set(product_to_fuels.get(product_key, []))
                    sector_fuels = set()
                    for s in sectors:
                        sector_fuels |= sector_projection_fuels.get(str(s).strip().lower(), set())
                    overlap = sorted(global_fuels.intersection(sector_fuels))
                    if len(overlap) == 1:
                        synthetic_sector = sectors[0] if sectors else ""
                        synthetic_fuel = overlap[0]
                        projection_support = True
                        if not suggested_fuel:
                            suggested_fuel = synthetic_fuel
                        if not suggested_flow:
                            suggested_flow = _resolve_sector_flow(sectors)
                        synthetic_rows.append(
                            {
                                "9th_sector": synthetic_sector,
                                "9th_fuel": synthetic_fuel,
                                "esto_flow": suggested_flow,
                                "esto_product": esto_product,
                                "mapping_note": "projection_only_synthetic_pair (ESTO base-year absent/zero)",
                                "source_sheet": sheet,
                                "source_fuel_label": fuel_label,
                            }
                        )
                        issue_detail = (
                            issue_detail
                            + " Projection data indicates a unique matching 9th fuel; synthetic pair suggested."
                        ).strip()
                        recommended_action = (
                            "Add suggested synthetic canonical row (mark as projection_only_synthetic_pair) "
                            "or add explicit override."
                        )

                unresolved_rows.append(
                    {
                        "sheet": sheet,
                        "fuel_label": fuel_label,
                        "issue": issue,
                        "issue_detail": issue_detail,
                        "recommended_action": recommended_action,
                        "esto_product_candidate": esto_product,
                        "candidate_ninth_fuels": " | ".join(candidate_fuels),
                        "projection_fuel_candidates": " | ".join(sorted(set().union(*[sector_projection_fuels.get(str(s).strip().lower(), set()) for s in sectors])) if sectors else []),
                        "synthetic_pair_suggested": bool(projection_support),
                        "synthetic_9th_sector": synthetic_sector,
                        "synthetic_9th_fuel": synthetic_fuel,
                        "suggest_leap_fuel_label": fuel_label,
                        "suggest_ninth_fuel_override": suggested_fuel,
                        "suggest_esto_product_override": esto_product,
                        "suggest_esto_flow_override": suggested_flow,
                    }
                )
                if hard:
                    conflict_rows.append(
                        {
                            "issue": issue,
                            "hard_conflict": True,
                            "sheet": sheet,
                            "fuel_label": fuel_label,
                            "9th_sector": " | ".join(sectors),
                            "9th_fuel": " | ".join(candidate_fuels),
                            "details": f"ESTO product '{esto_product}' maps to multiple 9th fuels for mapped sector(s).",
                            "suggested_override": "Add ninth_fuel_override in backup_leap_mappings.xlsx",
                        }
                    )

    unresolved_df = pd.DataFrame(unresolved_rows)
    _safe_write_csv(unresolved_df, output_dir / "unresolved_leap_labels.csv")
    synthetic_df = pd.DataFrame(synthetic_rows).drop_duplicates()
    if synthetic_df.empty:
        synthetic_df = pd.DataFrame(
            columns=[
                "9th_sector",
                "9th_fuel",
                "esto_flow",
                "esto_product",
                "mapping_note",
                "source_sheet",
                "source_fuel_label",
            ]
        )
    _safe_write_csv(synthetic_df, output_dir / "synthetic_projection_only_pairs.csv")

    conflicts_df = pd.DataFrame(conflict_rows)
    if conflicts_df.empty:
        conflicts_df = pd.DataFrame(
            columns=[
                "issue",
                "hard_conflict",
                "sheet",
                "fuel_label",
                "9th_sector",
                "9th_fuel",
                "details",
                "suggested_override",
            ]
        )
    _safe_write_csv(conflicts_df, output_dir / "mapping_conflicts.csv")

    hard_count = int(conflicts_df["hard_conflict"].fillna(False).astype(bool).sum())
    print(f"[INFO] canonical_pairs_clean.csv rows: {len(canonical_pairs)}")
    print(
        "[INFO] crosswalks: "
        "canonical_product_fuel_crosswalk.csv, "
        "canonical_flow_sector_crosswalk.csv, "
        "leap_branch_sector_flow_crosswalk.csv"
    )
    print(f"[INFO] leap_label_to_esto_product.csv rows: {len(label_map)}")
    print(f"[INFO] sheet_sector_coverage_report.csv rows: {len(coverage_df)}")
    print(f"[INFO] unresolved_leap_labels.csv rows: {len(unresolved_df)}")
    print(f"[INFO] synthetic_projection_only_pairs.csv rows: {len(synthetic_df)}")
    print(f"[INFO] mapping_conflicts.csv rows: {len(conflicts_df)} (hard={hard_count})")
    return 1 if hard_count > 0 else 0


def run_mapping_views_workflow(
    *,
    canonical_pairs_path: Path | str = Path("config/ninth_pairs_to_esto_pairs.xlsx"),
    codebook_path: Path | str = Path("config/sector_fuel_codes_to_names.xlsx"),
    sheet_map_path: Path | str = Path("config/leap_results_sheet_map.csv"),
    override_path: Path | str = Path("config/backup_leap_mappings.xlsx"),
    leap_long_path: Path | str | None = None,
    projection_table_path: Path | str = Path("data/merged_file_energy_ALL_20251106.csv"),
    output_dir: Path | str = Path("config/computer_generated_config/leap_mapping_views/USA"),
    fail_on_hard_conflicts: bool = False,
) -> dict[str, object]:
    """
    Notebook-friendly entry point.
    Returns a summary dict and does not call SystemExit.
    """
    ensure_repo_root()
    canonical_pairs_path = _resolve_path(canonical_pairs_path)
    codebook_path = _resolve_path(codebook_path)
    sheet_map_path = _resolve_path(sheet_map_path)
    override_path = _resolve_path(override_path)
    leap_long_path = _resolve_path(leap_long_path) if leap_long_path else None
    projection_table_path = _resolve_path(projection_table_path)
    output_dir = _resolve_path(output_dir)

    effective_leap_long = leap_long_path if leap_long_path and leap_long_path.exists() else None
    exit_code = build_views(
        canonical_pairs_path=canonical_pairs_path,
        codebook_path=codebook_path,
        sheet_map_path=sheet_map_path,
        output_dir=output_dir,
        override_path=override_path,
        leap_long_path=effective_leap_long,
        projection_table_path=projection_table_path if projection_table_path.exists() else None,
    )
    summary = {
        "output_dir": str(output_dir),
        "hard_conflicts_found": bool(exit_code != 0),
        "exit_code": int(exit_code),
        "canonical_pairs_clean_csv": str(output_dir / "canonical_pairs_clean.csv"),
        "canonical_product_fuel_crosswalk_csv": str(output_dir / "canonical_product_fuel_crosswalk.csv"),
        "canonical_flow_sector_crosswalk_csv": str(output_dir / "canonical_flow_sector_crosswalk.csv"),
        "leap_branch_sector_flow_crosswalk_csv": str(output_dir / "leap_branch_sector_flow_crosswalk.csv"),
        "leap_label_to_esto_product_csv": str(output_dir / "leap_label_to_esto_product.csv"),
        "sheet_sector_coverage_report_csv": str(output_dir / "sheet_sector_coverage_report.csv"),
        "unresolved_leap_labels_csv": str(output_dir / "unresolved_leap_labels.csv"),
        "synthetic_projection_only_pairs_csv": str(output_dir / "synthetic_projection_only_pairs.csv"),
        "mapping_conflicts_csv": str(output_dir / "mapping_conflicts.csv"),
    }
    if fail_on_hard_conflicts and exit_code != 0:
        raise RuntimeError(
            "Hard conflicts found in canonical mapping views. "
            f"See: {output_dir / 'mapping_conflicts.csv'}"
        )
    return summary

#%%
# Notebook-only module: call `run_mapping_views_workflow(...)` directly.
if __name__ == "__main__":  # pragma: no cover
    summary = run_mapping_views_workflow(fail_on_hard_conflicts=False)
    print("[INFO] Mapping views workflow finished.")
    for key, value in summary.items():
        print(f"- {key}: {value}")
#%%
