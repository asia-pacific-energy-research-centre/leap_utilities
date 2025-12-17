#%%
import pandas as pd
from pathlib import Path

from .economy_config import region_id_name_dict, scenario_dict

# def get_leap_metadata(measure):
#     """Fetch LEAP_units, LEAP_Scale, LEAP_Per from LEAP_MEASURE_CONFIG if available."""
#     for shortname, config_group in LEAP_MEASURE_CONFIG.items():
#         if measure == shortname or measure in config_group:
#             entry = config_group if isinstance(config_group, dict) else config_group[measure]
#             units = entry.get("LEAP_units", entry.get("unit", ""))
#             scale = entry.get("LEAP_Scale", "")
#             per = entry.get("LEAP_Per", "")
#             return units, scale, per
#     return "", "", ""


def create_import_instructions_sheet(writer):
    """Create instructions sheet inside the Excel file."""
    instructions = pd.DataFrame({
        "Step": range(1, 6),
        "Action": [
            "Open LEAP → Settings → Import Data...",
            "Select this Excel file and choose the 'Data' sheet.",
            "Map Branch Path → Branch, Variable → Variable, Years → Years.",
            "Select import options (Overwrite, Add, etc.).",
            "Run import and review LEAP’s message window."
        ]
    })
    instructions.to_excel(writer, sheet_name="Instructions", index=False)


def finalise_export_df(log_df, scenario, region, base_year, final_year
):
    """
    Create a LEAP-compatible Excel import file using LEAP_MEASURE_CONFIG metadata.
    Matches official LEAP Excel import format.
    """
    
    print(f"\n=== Creating LEAP Import File (structured) ===")

    if log_df is None or log_df.empty:
        print("[ERROR] No data available for export.")
        return None

    # --- Filter years ---
    log_df = log_df[(log_df["Date"] >= base_year) & (log_df["Date"] <= final_year)]
    
    # --- Pivot to wide format ---
    #just so we dont get an empty pivot, if any cols are fully None or na, fille them with str version of na then repalce once pivoted
    for col in ['Units', 'Scale', 'Per...']:
        if log_df[col].isna().all():
            log_df[col] = 'N/A'
        elif log_df[col].isnull().all():
            log_df[col] = 'null'
        elif (log_df[col] == '').all():
            log_df[col] = 'empty'
        elif (log_df[col] == None).all():
            log_df[col] = 'None'
       
    pivot_df = (
        log_df.pivot(
            index=["Branch_Path",'Scenario', "Measure", "Units", "Scale", "Per..."],
            columns="Date",
            values="Value"
        )
        .reset_index()
    )
    
    #now replace back the na values
    for col in ['Units', 'Scale', 'Per...']:
        pivot_df[col] = pivot_df[col].replace({'N/A': pd.NA, 'null': pd.NA, 'empty': '', 'None': None})
        #and do it to log df too just in case
        log_df[col] = log_df[col].replace({'N/A': pd.NA, 'null': pd.NA, 'empty': '', 'None': None})

    # --- Identify and sort year columns ---
    year_cols = sorted([int(c) for c in pivot_df.columns if isinstance(c, (int, float))])

    # --- Fill metadata columns ---
    pivot_df["Branch Path"] = pivot_df["Branch_Path"]
    pivot_df["Variable"] = pivot_df["Measure"]
    pivot_df["Region"] = region
    # pivot_df["Method"] = ''

    # # Fetch LEAP metadata from measure config
    # breakpoint()#check if htis is working. seems to me skipping steps
    # meta = pivot_df["Variable"].apply(lambda m: pd.Series(get_leap_metadata(m)))
    # meta.columns = ["Units", "Scale", "Per..."]
    # pivot_df = pd.concat([pivot_df, meta], axis=1)

    # --- Add Level 1–N columns ---
    max_levels = pivot_df["Branch_Path"].apply(lambda x: len(str(x).split("\\"))).max()
    for i in range(1, max_levels + 1):
        pivot_df[f"Level {i}"] = pivot_df["Branch_Path"].apply(
            lambda x: str(x).split("\\")[i - 1] if len(str(x).split("\\")) >= i else ""
        )

    # --- Sort variables within each branch in LEAP-like order ---
    var_order = [
        "Total Activity",
        "Activity Level",
        "Final Energy Intensity",
        "Total Final Energy Consumption",
        "Stock",
        "Sales Share",
        "Efficiency",
        "Turnover Rate",
        "Occupancy or Load",
    ]
    pivot_df["Variable_sort_order"] = pivot_df["Variable"].apply(
        lambda v: var_order.index(v) if v in var_order else len(var_order)
    )
    pivot_df = pivot_df.sort_values(by=["Branch_Path", "Variable_sort_order"]).drop(columns="Variable_sort_order")

    # --- Column order ---
    base_cols = ["Branch Path", "Variable", "Scenario", "Region", "Scale", "Units", "Per..."]#, "Method"
    level_cols = [f"Level {i}" for i in range(1, max_levels + 1)]
    export_df = pivot_df[base_cols + year_cols + level_cols].copy()

    # --- Add trailing placeholder column for #N/A ---
    # export_df.loc[:, "#N/A"] = pd.NA
    
    return export_df


def save_export_files(leap_export_df, export_df_for_viewing, leap_export_filename, base_year, final_year, model_name):
    """Save the export DataFrame and log DataFrame to an Excel file."""
    # --- Write to Excel ---
    out_path = Path(leap_export_filename)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    
    # Create two header rows for LEAP format
    # Row 0: Area/Version info
    # Row 1: Empty
    # Row 2: Column headers (will be set by pandas)
    
    leap_export_df2 = leap_export_df.copy()
    export_df_for_viewing2 = export_df_for_viewing.copy()
    
    # Create header rows with all columns from the dataframe
    # Row 0: Area and Version info
    header_data_0 = {col: '' for col in leap_export_df2.columns}
    header_data_0['Branch Path'] = 'Area:'
    header_data_0['Variable'] = model_name
    header_data_0['Scenario'] = 'Ver:'
    header_data_0['Region'] = '2'
    header_row_0 = pd.DataFrame([header_data_0])
    
    # Row 1: Empty row
    nas = pd.DataFrame([{col: pd.NA for col in leap_export_df2.columns}])
    
    header_row_2 = pd.DataFrame([leap_export_df2.columns], columns=leap_export_df2.columns)
    
    # Concatenate header rows with data
    leap_export_df2 = pd.concat([header_row_0, nas, header_row_2, leap_export_df2], ignore_index=True)
    # Same for viewing sheet
    header_data_0_view = {col: '' for col in export_df_for_viewing2.columns}
    header_data_0_view['Branch Path'] = 'Area:'
    header_data_0_view['Variable'] = model_name
    header_data_0_view['Scenario'] = 'Ver:'
    header_data_0_view['Region'] = '2'
    header_row_0_view = pd.DataFrame([header_data_0_view])
    
    nas = pd.DataFrame([{col: pd.NA for col in export_df_for_viewing2.columns}])
    
    # Row 2: Column names row (for compatibility with some LEAP import formats)
    header_row_2_view = pd.DataFrame([export_df_for_viewing2.columns], columns=export_df_for_viewing2.columns)
    
    export_df_for_viewing2 = pd.concat([header_row_0_view, nas, header_row_2_view, export_df_for_viewing2], ignore_index=True)
    
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        export_df_for_viewing2.to_excel(writer, sheet_name="FOR_VIEWING", index=False, header=False, startrow=0)
        leap_export_df2.to_excel(writer, sheet_name="LEAP", index=False, header=False, startrow=0)
    print(f"✅ Created file for importing into leap, and viewing at {leap_export_filename}, with {len(export_df_for_viewing)} entries.")
    print(f" - Years covered: {base_year}–{final_year}")
    print(f" - Variables: {leap_export_df['Variable'].nunique()}")
    print(f" - Branches: {export_df_for_viewing['Branch Path'].nunique()}")
    print("=" * 60)


def check_scenario_and_region_ids(import_df, scenario, region):
    #check the region id and name dict for this region and change the values in import df to match those. If they are not availble then raise an error since this will cause issues later that will be harder to identify than here. it can be imported from transport_economy_config
    # if there are multiple region in the import df then keep only the ones matching the region
    dict_regions = [region['region_name'] for region in region_id_name_dict.values()]
    if region not in dict_regions:#what is region
        breakpoint()
        raise ValueError(f"[ERROR] The region {region} specified for structure checking is not found in the region_id_name_dict: {dict_regions}. Make sure to load the correct region data for structure checking.")
    import_df = import_df[import_df['Region'] == region]
    #CONCERN what if the regions in the dict are incorrect for a new user. I guess it will be noticed later on when the region ids dont match up...
    import_regions = import_df['Region'].unique()
    if len(import_regions) != 1:
        breakpoint()
        raise ValueError(f"[ERROR] Less than one region found in import_df during structure checks: {import_regions}.")
    if region not in import_regions:
        #we will change the region names in the import df to match those in the dict:
        import_df['Region'] = region
        region_id = [region_id_name_dict[key]['region_id'] for key in region_id_name_dict if region_id_name_dict[key]['region_name'] == region]
        if len(region_id) != 1:
            raise ValueError(f"[ERROR] Multiple region ids found for region {region} in region_id_name_dict.")
        import_df['RegionID'] = region_id[0]
        # raise ValueError(f"[ERROR] The region {region} specified for structure checking is not found in the import dataframe regions: {import_regions}. Make sure to load the correct region data for structure checking.")
        
    #do the same for scenario ids. if there are multiple scenarios in the import df then keep only the ones matching the target scenario and current accounts
    dict_scenarios = [scenario_dict[key]['scenario_name'] for key in scenario_dict]
    if scenario not in dict_scenarios:#what is scenario
        breakpoint()
        raise ValueError(f"[ERROR] The scenario {scenario} specified for structure checking is not found in the scenario_dict: {dict_scenarios}. Make sure to load the correct scenario data for structure checking.")
    #drop scnearios that arent the scneario or the current accounts scenario
    import_df = import_df[(import_df['Scenario'] == scenario) | (import_df['Scenario'] == "Current Accounts")]
    
    import_scenarios = import_df['Scenario'].unique()
    #drop current accounts from this list since we will handle it separately
    import_scenarios = [s for s in import_scenarios if s != "Current Accounts"]
    if len(import_scenarios) != 1:
        breakpoint()
        raise ValueError(f"[ERROR] More or less than one scenario found in import_df during structure checks: {import_scenarios}. There should be a Current Accounts scenario and the projected scenario.")
    if scenario not in import_scenarios:
        #we will change the scenario names in the import df to match those in the dict where Current Accounts is not the scenario:
        import_df.loc[import_df['Scenario'] != "Current Accounts", 'Scenario'] = scenario
        scenario_id = [scenario_dict[key]['scenario_id'] for key in scenario_dict if scenario_dict[key]['scenario_name'] == scenario]
        if len(scenario_id) != 1:
            raise ValueError(f"[ERROR] Multiple scenario ids found for scenario {scenario} in scenario_dict.")
        import_df.loc[import_df['Scenario'] != "Current Accounts", 'ScenarioID'] = scenario_id[0]
        # raise ValueError(f"[ERROR] The scenario {scenario} specified for structure checking is not found in the import dataframe scenarios: {import_scenarios}. Make sure to load the correct scenario data for structure checking.")
    
    return import_df

def join_and_check_import_structure_matches_export_structure(import_filename, export_df, export_df_for_viewing, scenario, region, STRICT_CHECKS=True, current_accounts_label="Current Accounts"):
    new_current_account_df = pd.DataFrame()
    current_accounts_only_rows = pd.DataFrame()
    import_df = pd.read_excel(import_filename, sheet_name='Export', header=2)    
    non_current_scenarios = export_df.loc[export_df["Scenario"] != current_accounts_label, "Scenario"].unique()
    if len(non_current_scenarios) != 1:
        breakpoint()
        raise ValueError("[ERROR] More or less than one non-Current Accounts scenario found in export_df during structure checks.")
    # breakpoint()
    import_df = check_scenario_and_region_ids(import_df, scenario, region)
    #########################
    #FIRST HANDLE CURRENT ACCOUNTS SPECIAL CASE
    #########################
    #this next block is to handle current accounts scenario checking. an example of rows that are in current accoutns but not in other scenarios are things like 'Stock shares' for the the branches in between the top level and the lowest level. these are needed for current accounts to define stocks but arent needed in other scenarios since they are calculated. 
    
    #also extract a current accounts set of data to make sure it matches the set we have. 
    imported_current_account_df = import_df[(import_df['Scenario'] == current_accounts_label) & (import_df['Region'] == region)]
    current_accounts_df = export_df[(export_df['Scenario'] == current_accounts_label) & (export_df['Region'] == region)]
    
    #check where there are rows in current accounts df that are not in the import df for the target scenario
    missing_in_import_df = imported_current_account_df.merge(
        current_accounts_df,
        how='outer',
        on=['Branch Path', 'Variable', 'Region'],
        indicator=True,
        suffixes=('', '_import')
    )
    
    if not missing_in_import_df[missing_in_import_df['_merge'] != 'both'].empty:
        #skip those with unneeded variables:
        unneeded_vars = [
            'First Sales Year',
            'Fraction of Scrapped Replaced',
            'Max Scrappage Fraction',
            'Scrappage',
            'Fuel Economy Correction Factor',
            'Mileage Correction Factor'
        ]#these default to reasonable numbers which are by default correct so we dont need to have it in current accounts output
        missing_in_import_df = missing_in_import_df[~missing_in_import_df['Variable'].isin(unneeded_vars)]
        #if there are some right only rows then that is kind of unexpected..
        if len(missing_in_import_df[missing_in_import_df['_merge'] == 'right_only']) > 0:
            if STRICT_CHECKS:
                breakpoint()#if this is occuring then it means we have rows in current accounts df that arent in the import df for the target scenario and need to invesitgate why
                raise ValueError(f"[ERROR] Some rows need to be removed from Current Accounts scenario that do not exist in the import dataframe for that scenario: {missing_in_import_df[missing_in_import_df['_merge'] == 'right_only']}")
            else:
                print("[WARN] Some rows are missing in Current Accounts scenario that exist in the import dataframe for the target scenario:")
        #search for the left only rows in the export df to see if they exist there. if they do, we will extract them now and add them to the new_current_account_df
        current_accounts_only_rows = missing_in_import_df[missing_in_import_df['_merge'] == 'left_only'][['Branch Path', 'Variable', 'Region', "BranchID", "VariableID", "ScenarioID", "RegionID"]]
        if len(current_accounts_only_rows) > 0:
            breakpoint()
            if STRICT_CHECKS:
                breakpoint()#if this is occuring then it means we have rows in the export df df that arent in the new_current_account_df for the target scenario and need to invesitgate why
                raise ValueError(f"[ERROR] Some rows are missing in the import dataframe for Current Accounts scenario that exist in the export dataframe for that scenario: {current_accounts_only_rows}")
            else:
                print("[WARN] Some rows are missing in the import dataframe for Current Accounts scenario that exist in the export dataframe for that scenario:")
        # rows_to_add = export_df.merge(
        #     current_accounts_only_rows,
        #     how='right',
        #     on=['Branch Path', 'Variable', 'Region'],
        #     indicator=True
        # )
        # rows_to_add['Scenario'] = current_accounts_label
        # #drop them from export_df sicne we only want them for the current accounts scenario
        # export_df = export_df[~export_df.set_index(['Branch Path', 'Variable', 'Region']).index.isin(current_accounts_only_rows.set_index(['Branch Path', 'Variable', 'Region']).index)]
        # #if there are any rows in current_accounts_only_rows that are not in the export df then raise an error
        # if len(rows_to_add[rows_to_add['_merge'] == 'right_only']) > 0:
        #     breakpoint()
        #     # raise ValueError("[ERROR] Some rows are missing in the export dataframe that exist in Current Accounts scenario:")
        #     print("[WARN] Some rows are missing in the export dataframe that exist in Current Accounts scenario:")
        # #now we can add the rows to the new_current_account_df which we will add to later.
        # new_current_account_df = pd.concat([new_current_account_df, rows_to_add[rows_to_add['_merge'] == 'both'].drop(columns=['_merge'])], ignore_index=True)
    
    #########################
    #NOW CHECK THE STRUCTURE OF THE IMPORT AND EXPORT DFs for main SCENARIO
    #########################
    
    #now we can move on to checking the structure of the import and export dfs for the main scneario and current accoutns scneairo:
    new_df = pd.DataFrame()
    for scenario_ in [current_accounts_label, scenario]:
        import_df_scenario = import_df[(import_df['Scenario'] == scenario_) & (import_df['Region'] == region)] 
        export_df_scenario = export_df[(export_df['Scenario'] == scenario_) & (export_df['Region'] == region)] 
        for col in export_df_scenario.columns:
            if col not in import_df_scenario.columns:
                if col not in []:#levels are not necessary 
                    breakpoint()
                    print(f"[WARN] Column {col} is missing in import dataframe")
                    # raise ValueError(f"Column {col} is missing in import dataframe")
        for col in import_df_scenario.columns:
            if col not in export_df_scenario.columns:
                if col not in ["BranchID", "VariableID", "ScenarioID", "RegionID"]:
                    if 'Unnamed' in col:
                        continue
                    if 'Level' in col:#levels are not necessary 
                        continue
                    breakpoint()
                    print(f"[WARN] Column {col} is missing in export dataframe")
                    # raise ValueError(f"Column {col} is missing in export dataframe")
                    
        print("Import and export dataframes have matching structure.")
        
        #now join them together for comparison
        comparison_df = import_df_scenario.merge(export_df_scenario, how='outer', on=['Branch Path', 'Variable', 'Scenario', 'Region'], suffixes=('_import', '_export'), indicator=True)
        #where valyes are not the same in the cols: Scale	Units	Per... then print them out
        different_cols = comparison_df.copy()
        # Fill NAs with 'NA' for comparison
        for col in ['Scale', 'Units', 'Per...']:
            different_cols[f'{col}_import'] = different_cols[f'{col}_import'].fillna('NA')
            different_cols[f'{col}_export'] = different_cols[f'{col}_export'].fillna('NA')
        #filter to only those rows where the cols are different
        different_cols = different_cols[
            ((different_cols['Scale_import'] != different_cols['Scale_export']) |
            (different_cols['Units_import'] != different_cols['Units_export']) |
                    (different_cols['Per..._import'] != different_cols['Per..._export'])) & (different_cols['_merge'] == 'both')
                ]
        
        unneeded_vars = [
        "Fraction of Scrapped Replaced",
        "Max Scrappage Fraction",
        "Scrappage",
        "Fuel Economy Correction Factor",
        "Mileage Correction Factor",
        "First Sales Year"
        ]
        
        #drop those with unneeded variables: (these arent needed becase they default to reasonable numbers which are by default correct so we dont need to have it in output)
        different_cols = different_cols[~different_cols['Variable'].isin(unneeded_vars)]
        if not different_cols.empty:
            print("[WARN] Differences found between import and export dataframes in Scale, Units, or Per... columns:")
            print(different_cols)
            if STRICT_CHECKS:
                breakpoint()
                raise ValueError("Differences found between import and export dataframes in Scale, Units, or Per... columns:")
            
        #also check for where merge is not 'both'. first drop unneeded vars
        comparison_df = comparison_df[~comparison_df['Variable'].isin(unneeded_vars)]
        #if there are rows where the variable is 'Stock' and the scenario is not current accountns and the _merge is left_only, then justdrop them since these are stock rows that only exist in current accounts but leap likes to have them in the import file even if they are not in the export file for that scenario... i dunno its weird
        comparison_df = comparison_df[~((comparison_df['Variable'] == 'Stock') & (comparison_df['Scenario'] != current_accounts_label) & (comparison_df['_merge'] == 'left_only'))]
        
        if not comparison_df[comparison_df['_merge'] != 'both'].empty:
            # breakpoint()
            print("[WARN] Some rows are missing in either import or export dataframes:")
            # raise ValueError("Some rows are missing in either import or export dataframes:")
            # print(comparison_df[comparison_df['_merge'] != 'both'])
            right_dfs = comparison_df[comparison_df['_merge'] == 'right_only']
            left_dfs = comparison_df[comparison_df['_merge'] == 'left_only']
            if not right_dfs.empty:
                breakpoint()
                print("Rows missing in import dataframe. This is usually where we have created values in this system for branches that dont exist in leap or are spelt differently:")
                print(right_dfs)
            if not left_dfs.empty:
                breakpoint()
                print("Rows missing in export dataframe. This is usually where we have created values for branches in leap that dont exist in this system or are spelt differently:")
                print(left_dfs)
            comparison_df = comparison_df[comparison_df['_merge'] == 'both']   
            
        #now drop all the extra cols except the new ones (BranchID	VariableID	ScenarioID)
        #first rename all cols to remove _export suffixes since we want to keep those
        comparison_df = comparison_df.rename(columns=lambda x: x.replace('_export', ''))
        #then drop the _import cols
        comparison_df = comparison_df.drop(columns=[col for col in comparison_df.columns if col.endswith('_import') or col == '_merge'])
        #then reorder cols to be in the follwoing roder:
        # BranchID	VariableID	ScenarioID	RegionID	Branch Path	Variable	Scenario	Region	Scale	Units	Per...	Expression		Level 1	Level 2	Level 3	Level 4	Level 5	Level 6	Level 7	Level 8...									
        base_cols = ["BranchID", "VariableID", "ScenarioID", "RegionID", "Branch Path", "Variable", "Scenario", "Region", "Scale", "Units", "Per...", "Expression"]
        level_cols = [f"Level {i}" for i in range(1, 15) if f"Level {i}" in comparison_df.columns] + [f"Level {i}..." for i in range(1, 15) if f"Level {i}..." in comparison_df.columns]
        other_cols = [col for col in comparison_df.columns if col not in base_cols + level_cols]
        if len(other_cols) > 0:
            print("In addition to the expect cols in the export df, we have these other cols:", other_cols)
            if STRICT_CHECKS:
                breakpoint()
                raise ValueError("Unexpected extra columns found in comparison dataframe.")
        comparison_df = comparison_df[base_cols + level_cols + other_cols]
        
        new_df = pd.concat([new_df, comparison_df], ignore_index=True)
    ##################################
    
    ##################################
    # #now we want to create a copy of that new_df which contains all of the current accounts rows as well. To make it simple we will just copy the rows from this comparison df and rename the scenario to current accounts, then add on the extra current accounts rows we found earlier:
    # if imported_current_account_df.empty:
    #     breakpoint()
    #     raise ValueError("No current accounts rows were found to add to the comparison dataframe. This wasnt expected")
    #     # print("[WARN] No current accounts rows were found to add to the comparison dataframe.")
    #     current_account_comparison_df = comparison_df.copy()
    #     current_account_comparison_df['Scenario'] = current_accounts_label
    # else:
    #     current_account_comparison_df = comparison_df.copy()
    #     current_account_comparison_df['Scenario'] = current_accounts_label
    #     #drop the BranchID	VariableID	ScenarioID	RegionID rows since they will be different for current accounts
    #     current_account_comparison_df = current_account_comparison_df.drop(columns=['BranchID', 'VariableID', 'ScenarioID', 'RegionID'])
    #     #join to get the right IDs from the current_account_df
    #     current_account_comparison_df = current_account_comparison_df.merge(
    #         imported_current_account_df[['Branch Path', 'Variable', 'Scenario', 'Region', 'RegionID', 'BranchID', 'VariableID', 'ScenarioID']],
    #         how='left',
    #         on=['Branch Path', 'Variable', 'Scenario', 'Region']
    #     )
                
    #     #find any cols missing in imported_current_account_df that are in current_account_comparison_df
    #     for col in current_account_comparison_df.columns:
    #         if col not in imported_current_account_df.columns:
    #             #double check its not a levelcol
    #             if 'Level' in col:
    #                 continue
    #             print(f"[WARN] Column {col} is missing in imported_current_account_df")
    #             if STRICT_CHECKS:
    #                 breakpoint()
    #                 raise ValueError(f"Column {col} is missing in current_account_df")
    #             # current_account_df[col] = pd.NA
                
    #     #concat the set of extra current accounts rows that are not in the export df which we found earlier
    #     new_current_account_df = pd.concat([new_current_account_df, current_account_comparison_df], ignore_index=True)
        
    # ################################
    # #deduplicate current accounts rows, preferring later entries (typically those sourced from the import file)
    # if not new_current_account_df.empty:
    #     dedup_keys = [col for col in ["Branch Path", "Variable", "Scenario", "Region"] if col in new_current_account_df.columns]
    #     new_current_account_df = new_current_account_df.drop_duplicates(subset=dedup_keys, keep="last")
    
    
    ################################
    # #finally we will concat the new_current_account_df to the comparison_df so we fianlly have all scenarios together
    # final_export_df = pd.concat([comparison_df, new_current_account_df], ignore_index=True)
        
    ################################
    
    ################################
    #lastly we are going to join on the for_viewing sheet to get the first three cols 
    
    #frist, do the meege on the main export df to get the first three cols.
    export_df_for_viewing = export_df_for_viewing.merge(new_df[['Branch Path', 'Variable', 'Scenario', 'Region', 'RegionID', 'BranchID', 'VariableID', 'ScenarioID']],
        how='outer',
        on=['Branch Path', 'Variable', 'Scenario', 'Region'],
        indicator=True
    )
    #if there are any left only rows then raise an error
    if len(export_df_for_viewing[export_df_for_viewing['_merge'] == 'left_only']) > 0:
        #drop any values in current_accounts_only_rows from the left only rows since we expect those to be missing
        #drop them from export_df sicne we only want them for the current accounts scenario
        export_df_for_viewing = export_df_for_viewing[~export_df_for_viewing.set_index(['Branch Path', 'Variable', 'Region']).index.isin(current_accounts_only_rows.set_index(['Branch Path', 'Variable', 'Region']).index)]
        if len(export_df_for_viewing[export_df_for_viewing['_merge'] == 'left_only']) > 0:
            if STRICT_CHECKS:
                breakpoint()
                raise ValueError(f"Some rows in export_df_for_viewing are missing in export dataframe for {scenario} scenario:")
            print(f"[WARN] Some rows in export_df_for_viewing are missing in export dataframe for {scenario} scenario:")
            # raise ValueError("[ERROR] Some rows in export_df_for_viewing are missing in comparison dataframe:")
            print(export_df_for_viewing[export_df_for_viewing['_merge'] == 'left_only'])
    #likewise for right only rows
    if len(export_df_for_viewing[export_df_for_viewing['_merge'] == 'right_only']) > 0:
        if STRICT_CHECKS:
            breakpoint()
            raise ValueError(f"Some rows in export_df are missing in export_df_for_viewing for {scenario} scenario:")
        print(f"[WARN] Some rows in export_df are missing in export_df_for_viewing for {scenario} scenario:")
        # raise ValueError("[ERROR] Some rows in export_df_for_viewing are missing in comparison dataframe:")
        print(export_df_for_viewing[export_df_for_viewing['_merge'] == 'right_only'])
    export_df_for_viewing = export_df_for_viewing.drop(columns=['_merge'])    
    
    # # Second create a version of the new rows in new_current_account_df with the expression col expanded into its respective years. The expression col has pattern: WORD(year1, value1, year2, value2, ...)
    # new_current_account_df_for_viewing = new_current_account_df.copy()
    # def expand_expression_col(row):
    #     expr = row['Expression']
    #     result = row.copy()
    #     if pd.isna(expr):
    #         return result
    #     parts = expr.split('(')
    #     if len(parts) < 2:
    #         return result
    #     args = parts[1].rstrip(')').split(',')
    #     year_value_pairs = list(zip(args[::2], args[1::2]))
    #     for year, value in year_value_pairs:
    #         year = year.strip()
    #         value = value.strip()
    #         try:
    #             year_str = str(year)
    #             value_float = float(value)
    #             result[year_str] = value_float
                
    #         except ValueError:
    #             breakpoint()
    #             continue
    #     #set method in row using the word
    #     result['Method'] = parts[0].strip()
    #     return result
    # new_current_account_df_for_viewing = new_current_account_df_for_viewing.apply(expand_expression_col, axis=1)
    # #drop the expression col and
    # #change the year cols names to all be ints
    # year_cols = [col for col in new_current_account_df_for_viewing.columns if str(col).isdigit() and len(str(col)) == 4]
    # for col in year_cols:
    #     new_current_account_df_for_viewing = new_current_account_df_for_viewing.rename(columns={col: int(col)})
        
    # new_current_account_df_for_viewing = new_current_account_df_for_viewing.drop(columns=['Expression'])
    # export_df_for_viewing=pd.concat([new_current_account_df_for_viewing, export_df_for_viewing], ignore_index=True)
    # dedup_keys_view = [col for col in ["Branch Path", "Variable", "Scenario", "Region"] if col in export_df_for_viewing.columns]
    # export_df_for_viewing = export_df_for_viewing.drop_duplicates(subset=dedup_keys_view, keep="last")
    ################################
    #make sure that RegionID	BranchID	VariableID	ScenarioID are at the front and also are ints
    for col in ['RegionID', 'BranchID', 'VariableID', 'ScenarioID']:
        new_df[col] = new_df[col].astype('Int64')
        export_df_for_viewing[col] = export_df_for_viewing[col].astype('Int64')
        
        #if any are na then raise a warning
        if new_df[col].isna().any():
            if STRICT_CHECKS:
                breakpoint()
                raise ValueError(f"Some rows in final_export_df have NA values in column {col}")
            breakpoint()
            print(f"[WARN] Some rows in final_export_df have NA values in column {col}")
        if export_df_for_viewing[col].isna().any():
            if STRICT_CHECKS:
                breakpoint()
                raise ValueError(f"Some rows in export_df_for_viewing have NA values in column {col}")
            breakpoint()
            print(f"[WARN] Some rows in export_df_for_viewing have NA values in column {col}")
        
    export_df_for_viewing = export_df_for_viewing[['RegionID', 'BranchID', 'VariableID', 'ScenarioID'] + [col for col in export_df_for_viewing.columns if col not in ['RegionID', 'BranchID', 'VariableID', 'ScenarioID']]]
    
    return new_df, export_df_for_viewing
         
         
def separate_current_accounts_from_scenario(export_df, base_year,scenario, current_accounts_label="Current Accounts"):
    """
    Clone the generated export data to populate a Current Accounts scenario with a few adjustments.
    """
    #set scneario
    export_df["Scenario"] = scenario
    ca_export_df = export_df.copy()
    ca_export_df["Scenario"] = current_accounts_label
    #drop all years except the base year in year col
    ca_export_df = ca_export_df[ca_export_df.Date == base_year]
    
    #we need to make it so some variables are only in current accounts and not in the main scenario. these are typically things like stock shares for the branches in between the top level and the lowest level. these are needed for current accounts to define stocks but arent needed in other scenarios since they are calculated.
    #list of variables to only keep in current accounts
    vars_to_only_keep_in_ca = [
        'Stock Share',
        'Stock'
    ]
    
    export_df = export_df[~export_df["Measure"].isin(vars_to_only_keep_in_ca)]
    
    combined_export_df = pd.concat([export_df, ca_export_df], ignore_index=True)

    return combined_export_df

# def create_key_assumptions_branches():
#     """Create key assumptions branches inside LEAP import file."""
#     print(f"\n=== Creating Key Assumptions Branches ===")
#     # Create a DataFrame for key assumptions branches
#     key_assumptions_data = {
#         "Branch Path": [
#             "Key Assumptions",
            
             
def copy_energy_spreadsheet_into_leap_import_file(
    leap_export_filename='../results/leap_balances_export_file.xlsx',
    energy_spreadsheet_filename='../data/merged_file_energy_ALL_20250814.csv',
    ECONOMY='20_USA',
    BASE_YEAR=2022,
    SUBTOTAL_COLUMN='subtotal_results',
    SCENARIO="reference",
    ROOT=r"Key Assumptions\Energy Balances",
    REGION="Region 1",
    DROP_ZERO_BRANCHES=True,
    sheet_name="Energy_Balances",
    variable_col_value='Key Assumption',
    units="PJ",
    filters_dict=None,
):
    """
    
    NOTE THAT THIS FUCTION HAS BEEN BUILT TO WORK INDEPENDTLY OF THE TRANSPORT BASED SYSTEM.
    
    Create a LEAP import-style sheet from an energy balance spreadsheet.

    Branch paths are constructed as:
        ROOT\\sector\\sub1\\sub2\\sub3\\sub4\\fuel\\subfuel
    The resulting dataframe includes Level 1-8 columns derived from that path.
    Branches with no energy across all years are dropped when DROP_ZERO_BRANCHES=True.

    Returns the constructed dataframe and, if leap_export_filename is provided,
    writes/replaces the sheet named ``sheet_name`` in that workbook.
    
    filters_dict can be set to choose only specific sectors or fuels to include. if they are None then all sectors and fuels are included.
    """
    source_path = energy_spreadsheet_filename
    if '.csv' in source_path:
        energy_df = pd.read_csv(source_path)
    else:
        energy_df = pd.read_excel(source_path)

    # Filter down to the requested economy/scenario and exclude subtotal rows
    filtered = energy_df[
        (energy_df["economy"] == ECONOMY)
        & (energy_df["scenarios"] == SCENARIO.lower())
        & (energy_df[SUBTOTAL_COLUMN] == False)
    ].copy()
    if filters_dict is not None:
        for column in filters_dict:
            allowed_values = filters_dict[column]
            filtered = filtered[filtered[column].isin(allowed_values)]
    hierarchy_cols = ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors", "fuels", "subfuels"]

    # Remove numeric prefixes (and leading whitespace) from hierarchy parts. they can be like 15_01_01_passenger, 15_transport_sector or so on. 
    for col in hierarchy_cols:
        filtered[col] = (
            filtered[col]
            .astype("string")
            .str.strip()
            .str.replace(r'^(?:\d+[_\-\s]*)+', '', regex=True)
        )
    # breakpoint()#check that all hierarchy cols done right
    
    def _clean_part(val):
        if pd.isna(val):
            return None
        val_str = str(val).strip()
        if val_str.lower() in {"", "x", "nan"}:
            return None
        return val_str

    def _build_branch_path(row):
        parts = [ROOT]
        for col in hierarchy_cols:
            part = _clean_part(row[col])
            if part:
                parts.append(part)
        return "\\".join(parts)

    filtered["Branch Path"] = filtered.apply(_build_branch_path, axis=1)

    # Identify numeric or str year columns
    year_cols = sorted([c for c in filtered.columns if isinstance(c, (int, float, str))])
    
    # Only keep the base year
    year_cols = [c for c in year_cols if str(c) == str(BASE_YEAR)]
    
    if not year_cols:
        breakpoint()
        print(f"[WARN] Base year {BASE_YEAR} not found in data columns.")
        return pd.DataFrame()

    if DROP_ZERO_BRANCHES:
        energy_totals = filtered[year_cols].fillna(0).sum(axis=1)
        filtered = filtered[energy_totals != 0]

    if filtered.empty:
        breakpoint()
        print("[WARN] No energy rows remain after filtering; nothing to copy into LEAP import file.")
        return pd.DataFrame()
    
    export_df = filtered[["Branch Path"] + year_cols].copy()
    export_df.insert(1, "Variable", variable_col_value)
    export_df.insert(2, "Scenario", SCENARIO)
    export_df.insert(3, "Region", REGION)
    export_df.insert(4, "Scale", pd.NA)
    export_df.insert(5, "Units", units)
    export_df.insert(6, "Per...", pd.NA)

    # Add Level 1-8 based on the branch path structure
    max_levels = min(8, export_df["Branch Path"].str.split("\\").str.len().max())
    for i in range(1, max_levels + 1):
        export_df[f"Level {i}"] = export_df["Branch Path"].apply(
            lambda x: x.split("\\")[i - 1] if len(x.split("\\")) >= i else ""
        )

    base_cols = ["Branch Path", "Variable", "Scenario", "Region", "Scale", "Units", "Per..."]
    level_cols = [f"Level {i}" for i in range(1, max_levels + 1)]
    export_df = export_df[base_cols + year_cols + level_cols]
    
    if leap_export_filename:
        with pd.ExcelWriter(leap_export_filename, engine="openpyxl", mode="w") as writer:
            export_df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"[INFO] Energy balances written to {leap_export_filename} (sheet '{sheet_name}').")

    # return export_df


#%%
#%%

# filtered = energy_df[
#     (energy_df["economy"] == ECONOMY)
#     & (energy_df["scenarios"] == SCENARIO.lower())
#     & (energy_df[SUBTOTAL_COLUMN] == False)
# ].copy()
