#%%
# leap_utils/scrapbook/build_ninth_sector_fuel_pairs.py
import pandas as pd
from pathlib import Path

DATA_PATH = Path("../../data/merged_file_energy_ALL_20250814.csv")
OUTPUT_PATH = Path("../../outputs/ninth_sector_fuel_pairs.csv")

SECTOR_COLS = ["sub4sectors", "sub3sectors", "sub2sectors", "sub1sectors", "sectors"]
FUEL_COLS = ["subfuels", "fuels"]

def most_specific(row, cols):
    for col in cols:
        val = row.get(col, "")
        if pd.notna(val) and str(val).strip() and str(val).strip() != "x":
            return str(val).strip()
    return ""

def main():
    df = pd.read_csv(DATA_PATH, low_memory=False, dtype=str).fillna("")

    # Build most-detailed sector/fuel labels
    df["sector_pair"] = df.apply(lambda r: most_specific(r, SECTOR_COLS), axis=1)
    df["fuel_pair"] = df.apply(lambda r: most_specific(r, FUEL_COLS), axis=1)

    # Keep unique mappings only
    mapping = (
        df[SECTOR_COLS + FUEL_COLS + ["sector_pair", "fuel_pair"]]
        .drop_duplicates()
        .sort_values(["sector_pair", "fuel_pair"])
    )

    mapping.to_csv(OUTPUT_PATH, index=False)
    print(f"Wrote {len(mapping)} unique mappings to {OUTPUT_PATH}")
#%%
if __name__ == "__main__":
    main()
#%%
