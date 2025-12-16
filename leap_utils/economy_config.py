"""
Configuration for transport LEAP runs, keyed by economy code.

Use `get_transport_run_config("<economy>")` to fetch all file paths and
defaults for that economy in one place.
"""

scenario_dict = {
    "Current Accounts": {
        "scenario_name": "Current Accounts",
        "scenario_code": "CA",
        "scenario_id": 1,
        },
    "Target": {
        "scenario_name": "Target",
        "scenario_code": "TGT",
        "scenario_id": 3,
        },
    "Reference": {
        "scenario_name": "Reference",
        "scenario_code": "REF",
        "scenario_id": 4,
        }
}
region_id_name_dict = {
    "12_NZ": {
        "region_id": 2,
        "region_name": "New Zealand",
        "region_code": "12_NZ",
    },
    "20_USA": {
        "region_id": 1,
        "region_name": "United States of America",
        "region_code": "20_USA",
    },
}