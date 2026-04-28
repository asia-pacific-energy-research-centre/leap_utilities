from __future__ import annotations

from pathlib import Path

import pandas as pd

from codebase.utilities.master_config import config_table_exists, read_config_table

from codebase.utilities.leap_results_dashboard_utils import apply_explicit_sector_reassignments


SYNTHETIC_REFERENCE_COLUMNS = [
    "rule_name",
    "rule_mode",
    "source_esto_flow",
    "source_esto_product",
    "source_ninth_sectors",
    "source_ninth_sub1sectors",
    "source_ninth_sub2sectors",
    "source_ninth_sub3sectors",
    "source_ninth_sub4sectors",
    "source_ninth_fuels",
    "source_ninth_subfuels",
    "create_esto",
    "create_ninth",
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
NINTH_SECTOR_COLUMNS = ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]
NINTH_FUEL_COLUMNS = ["fuels", "subfuels"]


def _clean_token(value: object) -> str:
    return str(value).strip() if value is not None and not pd.isna(value) else ""


def _pick_first_non_empty(*values: object) -> str:
    for value in values:
        token = _clean_token(value)
        if token:
            return token
    return ""

def _find_ninth_template(
    *,
    ninth_df: pd.DataFrame,
    sector_code: str,
    fuel_code: str,
) -> pd.Series | None:
    if ninth_df.empty:
        return None
    working = ninth_df.copy()
    sector_code = _clean_token(sector_code)
    fuel_code = _clean_token(fuel_code)
    if sector_code:
        sector_mask = pd.Series(False, index=working.index)
        for col in NINTH_SECTOR_COLUMNS:
            if col in working.columns:
                sector_mask |= working[col].fillna("").astype(str).str.strip().eq(sector_code)
        working = working.loc[sector_mask].copy()
    if working.empty:
        return None
    if fuel_code:
        fuel_mask = pd.Series(False, index=working.index)
        for col in NINTH_FUEL_COLUMNS:
            if col in working.columns:
                fuel_mask |= working[col].fillna("").astype(str).str.strip().eq(fuel_code)
        exact = working.loc[fuel_mask].copy()
        if not exact.empty:
            working = exact
    if working.empty:
        return None
    return working.reset_index(drop=True).iloc[0]


def _resolve_rule_targets(
    *,
    rule: pd.Series,
    ninth_df: pd.DataFrame,
    explicit_mappings: pd.DataFrame,
) -> dict[str, object]:
    resolved = {col: rule.get(col, "") for col in SYNTHETIC_REFERENCE_COLUMNS}
    if _clean_token(rule.get("rule_mode")).lower() != "derived":
        return resolved

    template_sector_code = _pick_first_non_empty(
        resolved.get("target_sub4sectors"),
        resolved.get("target_sub3sectors"),
        resolved.get("target_sub2sectors"),
        resolved.get("target_sub1sectors"),
        resolved.get("target_sectors"),
        resolved.get("source_ninth_sub4sectors"),
        resolved.get("source_ninth_sub3sectors"),
        resolved.get("source_ninth_sub2sectors"),
        resolved.get("source_ninth_sub1sectors"),
        resolved.get("source_ninth_sectors"),
    )
    template_fuel_code = _pick_first_non_empty(
        resolved.get("target_subfuels"),
        resolved.get("target_fuels"),
        resolved.get("source_ninth_subfuels"),
        resolved.get("source_ninth_fuels"),
    )
    template_row = _find_ninth_template(
        ninth_df=ninth_df,
        sector_code=template_sector_code,
        fuel_code=template_fuel_code,
    )
    if template_row is not None:
        for col in NINTH_SECTOR_COLUMNS + NINTH_FUEL_COLUMNS:
            target_col = f"target_{col}"
            if not _clean_token(resolved.get(target_col)):
                resolved[target_col] = _clean_token(template_row.get(col))

    if not _clean_token(resolved.get("target_esto_flow")):
        resolved["target_esto_flow"] = _pick_first_non_empty(resolved.get("source_esto_flow"))
    if not _clean_token(resolved.get("target_esto_product")):
        resolved["target_esto_product"] = _pick_first_non_empty(resolved.get("source_esto_product"))
    return resolved


def _expand_resolved_rule_targets(
    *,
    resolved_rule: dict[str, object],
    canonical_pairs: pd.DataFrame | None = None,
) -> list[dict[str, object]]:
    if _clean_token(resolved_rule.get("rule_mode")).lower() != "derived":
        return [resolved_rule]

    canonical_pairs = canonical_pairs if canonical_pairs is not None else pd.DataFrame()
    if canonical_pairs.empty:
        return [resolved_rule]

    pair_df = canonical_pairs.copy()
    for col in ["9th_sector", "9th_fuel", "esto_flow", "esto_product"]:
        if col not in pair_df.columns:
            pair_df[col] = ""
        pair_df[col] = pair_df[col].map(_clean_token)

    source_ninth_sector = _pick_first_non_empty(
        resolved_rule.get("source_ninth_sub4sectors"),
        resolved_rule.get("source_ninth_sub3sectors"),
        resolved_rule.get("source_ninth_sub2sectors"),
        resolved_rule.get("source_ninth_sub1sectors"),
        resolved_rule.get("source_ninth_sectors"),
    )
    source_ninth_fuel = _pick_first_non_empty(
        resolved_rule.get("source_ninth_subfuels"),
        resolved_rule.get("source_ninth_fuels"),
    )
    source_esto_flow = _clean_token(resolved_rule.get("source_esto_flow"))
    source_esto_product = _clean_token(resolved_rule.get("source_esto_product"))

    expanded_rules: list[dict[str, object]] = []

    if bool(resolved_rule.get("create_esto")) and not _clean_token(resolved_rule.get("target_esto_flow")):
        flow_candidates = pair_df.copy()
        if source_ninth_fuel:
            flow_candidates = flow_candidates[flow_candidates["9th_fuel"].eq(source_ninth_fuel)]
        if source_ninth_sector:
            sector_exact = flow_candidates[flow_candidates["9th_sector"].eq(source_ninth_sector)]
            if not sector_exact.empty:
                flow_candidates = sector_exact
        if source_esto_flow:
            flow_exact = flow_candidates[flow_candidates["esto_flow"].eq(source_esto_flow)]
            if not flow_exact.empty:
                flow_candidates = flow_exact
        flow_values = [value for value in flow_candidates["esto_flow"].drop_duplicates().tolist() if value]
        if flow_values:
            for flow_value in flow_values:
                variant = dict(resolved_rule)
                variant["target_esto_flow"] = flow_value
                expanded_rules.append(variant)

    if bool(resolved_rule.get("create_ninth")) and not _pick_first_non_empty(
        resolved_rule.get("target_sub4sectors"),
        resolved_rule.get("target_sub3sectors"),
        resolved_rule.get("target_sub2sectors"),
        resolved_rule.get("target_sub1sectors"),
        resolved_rule.get("target_sectors"),
    ):
        sector_candidates = pair_df.copy()
        target_or_source_product = _pick_first_non_empty(
            resolved_rule.get("target_esto_product"),
            source_esto_product,
        )
        if target_or_source_product:
            sector_candidates = sector_candidates[sector_candidates["esto_product"].eq(target_or_source_product)]
        if source_esto_flow:
            flow_exact = sector_candidates[sector_candidates["esto_flow"].eq(source_esto_flow)]
            if not flow_exact.empty:
                sector_candidates = flow_exact
        sector_pairs = (
            sector_candidates[["9th_sector", "9th_fuel"]]
            .drop_duplicates()
            .to_dict("records")
        )
        if sector_pairs:
            for pair in sector_pairs:
                variant = dict(resolved_rule)
                variant["target_sub2sectors"] = _pick_first_non_empty(
                    variant.get("target_sub2sectors"),
                    pair.get("9th_sector"),
                )
                variant["target_subfuels"] = _pick_first_non_empty(
                    variant.get("target_subfuels"),
                    pair.get("9th_fuel"),
                )
                expanded_rules.append(variant)

    if expanded_rules:
        seen: set[tuple[object, ...]] = set()
        deduped: list[dict[str, object]] = []
        key_cols = [
            "target_esto_flow",
            "target_esto_product",
            "target_sectors",
            "target_sub1sectors",
            "target_sub2sectors",
            "target_sub3sectors",
            "target_sub4sectors",
            "target_fuels",
            "target_subfuels",
        ]
        for variant in expanded_rules:
            key = tuple(_clean_token(variant.get(col)) for col in key_cols)
            if key in seen:
                continue
            seen.add(key)
            deduped.append(variant)
        return deduped

    return [resolved_rule]


def load_synthetic_reference_rows_config(path: Path) -> pd.DataFrame:
    if not config_table_exists(path):
        return pd.DataFrame(columns=SYNTHETIC_REFERENCE_COLUMNS)
    df = read_config_table(path)
    df.columns = [str(c).strip().lower() for c in df.columns]
    if "rule_mode" not in df.columns:
        df["rule_mode"] = "manual"
    if "create_esto" not in df.columns and "create_base" in df.columns:
        df["create_esto"] = df["create_base"]
    if "active" in df.columns:
        df = df[df["active"].astype(str).str.strip().str.lower().isin({"true", "1", "yes"})]
    for col in SYNTHETIC_REFERENCE_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    for col in SYNTHETIC_REFERENCE_COLUMNS:
        if col in {"create_esto", "create_ninth"}:
            df[col] = df[col].fillna("").astype(str).str.strip().str.lower().isin({"true", "1", "yes"})
        else:
            df[col] = df[col].fillna("").astype(str).str.strip()
    df["rule_mode"] = df["rule_mode"].replace("", "manual")
    return df[SYNTHETIC_REFERENCE_COLUMNS].reset_index(drop=True)


def append_synthetic_reference_rows(
    *,
    esto_df: pd.DataFrame,
    ninth_df: pd.DataFrame,
    rules: pd.DataFrame,
    explicit_mappings: pd.DataFrame | None = None,
    canonical_pairs: pd.DataFrame | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    status_cols = [
        "rule_name",
        "source_esto_matches",
        "esto_rows_added",
        "ninth_rows_added",
        "target_esto_flow",
        "target_esto_product",
        "target_ninth_sub2sectors",
        "notes",
    ]
    if rules.empty:
        return esto_df.copy(), ninth_df.copy(), pd.DataFrame(columns=status_cols)

    esto_out = esto_df.copy()
    ninth_out = ninth_df.copy()

    esto_year_cols = [c for c in esto_out.columns if str(c).strip().isdigit() and len(str(c).strip()) == 4]
    ninth_year_cols = [c for c in ninth_out.columns if str(c).strip().isdigit() and len(str(c).strip()) == 4]

    ninth_key_cols = [
        col
        for col in [
            "economy",
            "scenarios",
            "sectors",
            "sub1sectors",
            "sub2sectors",
            "sub3sectors",
            "sub4sectors",
            "fuels",
            "subfuels",
        ]
        if col in ninth_out.columns
    ]
    esto_key_cols = [col for col in ["economy", "flows", "products"] if col in esto_out.columns]

    scenario_templates = (
        ninth_out[
            [col for col in ninth_out.columns if col in {"economy", "scenarios", "subtotal_layout", "subtotal_results"}]
        ]
        .drop_duplicates()
        .reset_index(drop=True)
    )
    if scenario_templates.empty:
        scenario_templates = pd.DataFrame([{}])
    esto_templates = (
        esto_out[[col for col in esto_out.columns if col in {"economy"}]].drop_duplicates().reset_index(drop=True)
    )
    if esto_templates.empty:
        esto_templates = pd.DataFrame([{}])

    status_rows: list[dict[str, object]] = []
    new_ninth_rows: list[dict[str, object]] = []
    new_esto_rows: list[dict[str, object]] = []
    explicit_mappings = explicit_mappings if explicit_mappings is not None else pd.DataFrame()
    for _, rule in rules.iterrows():
        resolved_rule = _resolve_rule_targets(
            rule=rule,
            ninth_df=ninth_out,
            explicit_mappings=explicit_mappings,
        )
        expanded_rules = _expand_resolved_rule_targets(
            resolved_rule=resolved_rule,
            canonical_pairs=canonical_pairs,
        )
        source_flow = _clean_token(resolved_rule.get("source_esto_flow"))
        source_product = _clean_token(resolved_rule.get("source_esto_product"))
        source_esto_matches = esto_df.copy()
        if source_flow and "flows" in source_esto_matches.columns:
            source_esto_matches = source_esto_matches[
                source_esto_matches["flows"].fillna("").astype(str).str.strip().eq(source_flow)
            ]
        if source_product and "products" in source_esto_matches.columns:
            source_esto_matches = source_esto_matches[
                source_esto_matches["products"].fillna("").astype(str).str.strip().eq(source_product)
            ]
        source_esto_match_count = int(len(source_esto_matches))

        total_esto_rows_added = 0
        total_ninth_rows_added = 0
        status_targets: list[tuple[str, str]] = []
        for expanded_rule in expanded_rules:
            rule_name = _clean_token(expanded_rule.get("rule_name")) or "unnamed_rule"
            notes = _clean_token(expanded_rule.get("notes"))
            create_esto = bool(expanded_rule.get("create_esto"))
            create_ninth = bool(expanded_rule.get("create_ninth"))
            local_esto_rows_added = 0
            if create_esto and not esto_out.empty:
                for template in esto_templates.to_dict("records"):
                    candidate = {col: pd.NA for col in esto_out.columns}
                    for col, value in template.items():
                        candidate[col] = value
                    if "flows" in candidate:
                        candidate["flows"] = _clean_token(expanded_rule.get("target_esto_flow")) or source_flow
                    if "products" in candidate:
                        candidate["products"] = _clean_token(expanded_rule.get("target_esto_product")) or source_product or "x"
                    for year_col in esto_year_cols:
                        candidate[year_col] = 0
                    if esto_key_cols:
                        duplicate_mask = pd.Series(True, index=esto_out.index)
                        for key in esto_key_cols:
                            duplicate_mask &= esto_out[key].fillna("").astype(str).str.strip().eq(
                                str(candidate.get(key) or "").strip()
                            )
                        if bool(duplicate_mask.any()):
                            continue
                        duplicate_new = any(
                            all(str(existing.get(key) or "").strip() == str(candidate.get(key) or "").strip() for key in esto_key_cols)
                            for existing in new_esto_rows
                        )
                        if duplicate_new:
                            continue
                    candidate["_synthetic_esto_row"] = True
                    candidate["_synthetic_rule_name"] = rule_name
                    new_esto_rows.append(candidate)
                    local_esto_rows_added += 1

            local_ninth_rows_added = 0
            if create_ninth and not ninth_out.empty:
                for template in scenario_templates.to_dict("records"):
                    candidate = {col: pd.NA for col in ninth_out.columns}
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
                            candidate[col] = _clean_token(expanded_rule.get(rule_key)) or "x"
                    if "subtotal_layout" in candidate:
                        candidate["subtotal_layout"] = False
                    if "subtotal_results" in candidate:
                        candidate["subtotal_results"] = False
                    for year_col in ninth_year_cols:
                        candidate[year_col] = 0

                    if ninth_key_cols:
                        duplicate_mask = pd.Series(True, index=ninth_out.index)
                        for key in ninth_key_cols:
                            duplicate_mask &= ninth_out[key].fillna("").astype(str).str.strip().eq(str(candidate.get(key) or "").strip())
                        if bool(duplicate_mask.any()):
                            continue
                        duplicate_new = any(
                            all(str(existing.get(key) or "").strip() == str(candidate.get(key) or "").strip() for key in ninth_key_cols)
                            for existing in new_ninth_rows
                        )
                        if duplicate_new:
                            continue
                    candidate["_synthetic_ninth_row"] = True
                    candidate["_synthetic_rule_name"] = rule_name
                    new_ninth_rows.append(candidate)
                    local_ninth_rows_added += 1
            total_esto_rows_added += local_esto_rows_added
            total_ninth_rows_added += local_ninth_rows_added
            status_targets.append(
                (
                    _clean_token(expanded_rule.get("target_esto_flow")) or source_flow,
                    _clean_token(expanded_rule.get("target_sub2sectors")),
                )
            )

        status_rows.append(
            {
                "rule_name": _clean_token(resolved_rule.get("rule_name")) or "unnamed_rule",
                "source_esto_matches": source_esto_match_count,
                "esto_rows_added": total_esto_rows_added,
                "ninth_rows_added": total_ninth_rows_added,
                "target_esto_flow": "; ".join(sorted({flow for flow, _ in status_targets if flow})),
                "target_esto_product": _clean_token(resolved_rule.get("target_esto_product")) or source_product,
                "target_ninth_sub2sectors": "; ".join(sorted({sector for _, sector in status_targets if sector})),
                "notes": _clean_token(resolved_rule.get("notes")),
            }
        )

    if new_esto_rows:
        esto_out = pd.concat([esto_out, pd.DataFrame(new_esto_rows)], ignore_index=True, sort=False)
    if new_ninth_rows:
        ninth_out = pd.concat([ninth_out, pd.DataFrame(new_ninth_rows)], ignore_index=True, sort=False)
    return esto_out, ninth_out, pd.DataFrame(status_rows, columns=status_cols)


def drop_all_zero_year_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    year_cols = [c for c in df.columns if str(c).strip().isdigit() and len(str(c).strip()) == 4]
    if not year_cols:
        return df
    values = df[year_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    keep_mask = values.ne(0).any(axis=1)
    return df.loc[keep_mask].copy()


def load_reference_tables(
    *,
    esto_table_path: Path,
    projection_table_path: Path,
    explicit_reassignments: pd.DataFrame,
    explicit_mappings: pd.DataFrame | None = None,
    canonical_pairs: pd.DataFrame | None = None,
    synthetic_reference_rows_path: Path | None = None,
    drop_all_zero_base_rows: bool,
    drop_all_zero_projection_rows: bool,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    esto_df = read_config_table(esto_table_path)
    ninth_df = read_config_table(projection_table_path)
    esto_df, ninth_df, status = apply_explicit_sector_reassignments(esto_df, ninth_df, explicit_reassignments)
    synthetic_rules = (
        load_synthetic_reference_rows_config(synthetic_reference_rows_path)
        if synthetic_reference_rows_path
        else pd.DataFrame()
    )
    esto_df, ninth_df, synthetic_status = append_synthetic_reference_rows(
        esto_df=esto_df,
        ninth_df=ninth_df,
        rules=synthetic_rules,
        explicit_mappings=explicit_mappings,
        canonical_pairs=canonical_pairs,
    )
    # Preserve all-zero ESTO rows. Structural zeros at exact sheet level are
    # meaningful for mapping resolution; dropping them forces false parent-flow
    # fallbacks later in comparison build.
    if drop_all_zero_projection_rows:
        ninth_df = drop_all_zero_year_rows(ninth_df)
    return esto_df, ninth_df, status, synthetic_status
