from pathlib import Path

import pandas as pd

from codebase.utilities.energy_balance_template_extractor import TemplateBalanceExtractor


def _make_extractor(
    *,
    explicit_pair_mappings_only: bool,
    allow_descendant_mapping_expansion: bool = True,
    present_child: bool = False,
) -> TemplateBalanceExtractor:
    extractor = TemplateBalanceExtractor(
        template_sheet="EBal|2060",
        mapping_pairs_path=Path("config/leap_mappings.xlsx"),
        codebook_path=Path("config/esto_9th_leap_codebook.xlsx"),
        explicit_pair_mappings_only=explicit_pair_mappings_only,
        allow_descendant_mapping_expansion=allow_descendant_mapping_expansion,
    )

    # Keep this test focused on explicit full-path mappings and descendant reuse.
    extractor._flow_name_to_codes = {}
    extractor._fuel_name_to_codes = {}
    extractor._flow_name_to_esto = {}
    extractor._fuel_name_to_esto = {}
    extractor._canonical_pair_to_esto = {}
    extractor._balance_name_pair_to_esto = {}
    extractor._balance_code_pair_to_esto = {}
    extractor._balance_name_pair_to_ninth = {}
    extractor._flow_code_cache = {}
    extractor._fuel_code_cache = {}
    extractor._flow_esto_cache = {}
    extractor._fuel_esto_cache = {}
    extractor._balance_full_path_pairs_to_remove = set()
    extractor._balance_full_path_pairs_with_removed_rows = set()

    fuel_key = extractor._canonicalize_label("Coal")
    parent_key = extractor._canonicalize_path_key("Transformation")
    child_key = extractor._canonicalize_path_key("Transformation/Electricity plants")

    extractor._balance_full_path_pair_to_esto = {
        (child_key, fuel_key): [
            {
                "esto_flow": "09.01.01 Electricity plants",
                "esto_product": "01 Coal",
                "candidate_leap_sector_name_full_path": "Transformation/Electricity plants",
                "candidate_leap_fuel_name": "Coal",
                "candidate_rule": "test_child_mapping",
                "pair_mapping_cardinality": "one_to_one",
            }
        ]
    }
    extractor._balance_full_path_pair_to_ninth = {
        (child_key, fuel_key): [
            {
                "ninth_sector": "09_01_01_electricity_plants",
                "ninth_fuel": "01_coal",
                "candidate_leap_sector_name_full_path": "Transformation/Electricity plants",
                "candidate_leap_fuel_name": "Coal",
                "candidate_rule": "test_child_mapping",
                "pair_mapping_cardinality": "one_to_one",
            }
        ]
    }

    present_keys = {(parent_key, fuel_key)}
    if present_child:
        present_keys.add((child_key, fuel_key))
    extractor._balance_present_source_keys_by_sheet = {"Balance": present_keys}
    return extractor


def _parent_row() -> pd.Series:
    return pd.Series(
        {
            "source_sheet": "Balance",
            "leap_sector_name": "Transformation",
            "leap_sector_name_full_path": "Transformation",
            "leap_sector_name_original": "Transformation",
            "leap_fuel_name": "Coal",
        }
    )


def test_non_explicit_mode_uses_absent_child_descendant_mapping_by_default() -> None:
    extractor = _make_extractor(explicit_pair_mappings_only=False, present_child=False)

    records = extractor._map_row_records(_parent_row())

    assert len(records) == 1
    assert records[0]["mapping_status"] == "mapped"
    assert records[0]["mapping_method"] == "module_full_path_pair"
    assert records[0]["match_resolution"] == "module_only"
    assert records[0]["esto_flow"] == "09.01.01 Electricity plants"


def test_descendant_mapping_expansion_switch_disables_absent_child_mapping() -> None:
    extractor = _make_extractor(
        explicit_pair_mappings_only=False,
        allow_descendant_mapping_expansion=False,
        present_child=False,
    )

    records = extractor._map_row_records(_parent_row())

    assert len(records) == 1
    assert records[0]["mapping_status"] == "unmapped"
    assert records[0]["mapping_method"] == ""
    assert records[0]["match_resolution"] == "detailed"
    assert records[0]["esto_flow"] == ""


def test_explicit_pair_mode_blocks_absent_child_descendant_mapping() -> None:
    extractor = _make_extractor(explicit_pair_mappings_only=True, present_child=False)

    records = extractor._map_row_records(_parent_row())

    assert len(records) == 1
    assert records[0]["mapping_status"] == "unmapped"
    assert records[0]["mapping_method"] == ""
    assert records[0]["match_resolution"] == "detailed"
    assert records[0]["esto_flow"] == ""


def test_parent_row_does_not_reuse_descendant_mapping_when_child_source_is_present() -> None:
    extractor = _make_extractor(explicit_pair_mappings_only=False, present_child=True)

    records = extractor._map_row_records(_parent_row())

    assert len(records) == 1
    assert records[0]["mapping_status"] == "unmapped"
    assert records[0]["remove_row"] is True
    assert records[0]["esto_flow"] == ""
