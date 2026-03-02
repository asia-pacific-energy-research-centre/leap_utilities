
###############################
import pandas as pd
#%%

#load data\temp\new_labels.csv and map the labels to their 9th column so they have cols 'value'	'col'. save that as unmapped.xlsx. then put them through the code below

new_labels = pd.read_csv("../../data/temp/new_labels.csv")["value"].tolist()
ninth_data = pd.read_csv("../../data/merged_file_energy_ALL_20250814_pre_trump.csv")[["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors", "fuels", "subfuels"]].drop_duplicates()  # Use merged_file_energy_ALL_20251106.csv and merged_file_energy_00_APEC_20251106 for exact 9th edition projection matching.

unmapped_ninth = []

for label in new_labels:
    if label == "" or pd.isna(label):
        continue
    #strip wjhitespace
    label = label.strip()
    match = ninth_data[
        (ninth_data["sectors"] == label) |
        (ninth_data["sub1sectors"] == label) |
        (ninth_data["sub2sectors"] == label) |
        (ninth_data["sub3sectors"] == label) |
        (ninth_data["sub4sectors"] == label) |
        (ninth_data["fuels"] == label) |
        (ninth_data["subfuels"] == label)
    ]
    if not match.empty:
        for col in ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors", "fuels", "subfuels"]:
            if match.iloc[0][col] == label:
                unmapped_ninth.append((label, col))
                break
    else:
        unmapped_ninth.append((label, "unknown_column"))

#put in df with cols value and col
unmapped_ninth_data = pd.DataFrame(unmapped_ninth, columns=["value", "col"])


#%%
# #load data/temp/unmapped.xlsx and map the labels to the 9th sector and fuel cols :

# # unmapped_ninth_data = pd.read_excel("../data/temp/unmapped.xlsx")[["value","col"]] 
# ninth_data = pd.read_csv("../data/merged_file_energy_ALL_20250814_pre_trump.csv")[["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors", "fuels", "subfuels"]].drop_duplicates()  # Use merged_file_energy_ALL_20251106.csv and merged_file_energy_00_APEC_20251106 for exact 9th edition projection matching.

#pivot unmapped_ninth_data so col makes the cols and value is the value in those cols. the cols will end up having many nas. it should have the same amount of rows as unmapped_ninth_data
unmapped_pivoted = unmapped_ninth_data.assign(row_id=range(len(unmapped_ninth_data))).pivot_table(index="row_id", columns="col", values="value", aggfunc='first').reset_index(drop=True)

# Create key and key_col columns from the original unmapped_ninth_data
unmapped_pivoted['key'] = unmapped_ninth_data['value'].values
unmapped_pivoted['key_col'] = unmapped_ninth_data['col'].values

existing_cols = unmapped_pivoted.columns.tolist()
fuels = unmapped_pivoted[[col for col in ["fuels", "subfuels", "key", "key_col"] if col in existing_cols]].drop_duplicates()
for col in ["fuels", "subfuels"]:
    if col not in fuels.columns:
        fuels[col] = pd.NA  # Add missing fuel columns initialized with NA
        
existing_cols = unmapped_pivoted.columns.tolist()
sectors = unmapped_pivoted[[col for col in ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors", "key", "key_col"] if col in existing_cols]].drop_duplicates()
for col in ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]:
    if col not in sectors.columns:
        sectors[col] = pd.NA  # Add missing sector columns initialized with NA

#filter out where there are only nas in the rows (excluding key and key_col)
fuel_data_cols = [col for col in fuels.columns if col not in ["key", "key_col"]]
fuels = fuels.dropna(subset=fuel_data_cols, how="all")

sector_data_cols = [col for col in sectors.columns if col not in ["key", "key_col"]]
sectors = sectors.dropna(subset=sector_data_cols, how="all")


# Fill missing hierarchy columns for fuels
for col in ["fuels", "subfuels"]:
    if col in fuels.columns:
        # Create a mapping from ninth_data for this column
        mapping = ninth_data[ninth_data[col].notna()][["fuels", "subfuels"]].drop_duplicates()
        
        # Merge to fill in the missing columns
        fuels = fuels.merge(
            mapping,
            on=col,
            how="left",
            suffixes=("", "_filled")
        )
        
        # Fill NaN values in other columns with the merged values
        for fill_col in ["fuels", "subfuels"]:
            if f"{fill_col}_filled" in fuels.columns:
                if fill_col in fuels.columns:
                    fuels[fill_col] = fuels[fill_col].fillna(fuels[f"{fill_col}_filled"])
                else:
                    fuels[fill_col] = fuels[f"{fill_col}_filled"]
                fuels = fuels.drop(columns=[f"{fill_col}_filled"])

# Fill missing hierarchy columns for sectors
for col in ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]:
    if col in sectors.columns:
        
        # Create a mapping from ninth_data for this column
        mapping = ninth_data[ninth_data[col].notna()][["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]].drop_duplicates()
        sectors_ = sectors[sectors[col].notna()]
        if sectors_.empty:
            continue
        # Merge to fill in the missing columns
        sectors = sectors.merge(
            mapping,
            on=col,
            how="left",
            suffixes=("", "_filled")
        )
        
        # Fill NaN values in other columns with the merged values
        for fill_col in ["sectors", "sub1sectors", "sub2sectors", "sub3sectors", "sub4sectors"]:
            if f"{fill_col}_filled" in sectors.columns:
                if fill_col in sectors.columns:
                    sectors[fill_col] = sectors[fill_col].fillna(sectors[f"{fill_col}_filled"])
                else:
                    sectors[fill_col] = sectors[f"{fill_col}_filled"]
                sectors = sectors.drop(columns=[f"{fill_col}_filled"])
    #drop duplicates again
sectors = sectors.drop_duplicates(['key'])
fuels = fuels.drop_duplicates(['key'])
#%%
#make the key cols show at the start and then save as csvs
sector_cols = [col for col in sectors.columns if col not in ['key', 'key_col']]
sectors = sectors[['key', 'key_col'] + sector_cols]
fuel_cols = [col for col in fuels.columns if col not in ['key', 'key_col']]
fuels = fuels[['key', 'key_col'] + fuel_cols]
sectors.to_csv("../../data/temp/unmapped_sectors.csv", index=False)
fuels.to_csv("../../data/temp/unmapped_fuels.csv", index=False)
#finally creat dicts
