from codebase.leap_series_analysis_workflow import (  # noqa: F401
    BASE_YEAR,
    BRANCH_SECTOR_MAPPING_CSV,
    CODE_TO_NAME_PATH,
    CODE_TO_NAME_SHEET,
    ECONOMY,
    ESTO_DATA_PATH,
    FUEL_ALIASES_CSV,
    LEAP_RESULTS_FILE,
    NINTH_DATA_PATH,
    NINTH_SCENARIO,
    NINTH_TO_ESTO_MAPPING_PATH,
    OUTPUT_DIR,
    PROJECTION_END_YEAR,
    PROJECTION_START_YEAR,
    REGION,
    SCENARIO,
    SHARE_YEAR_OFFSET,
    SUBTOTAL_MAPPING_PATH,
    build_config,
    run_with_config,
)


if __name__ == "__main__":
    run_with_config()
