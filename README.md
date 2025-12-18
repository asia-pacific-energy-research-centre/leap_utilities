# leap-utilities

Shared LEAP helpers (COM connection, branch utilities, Excel import/export, energy-use reconciliation) extracted from the transport toolkit. Suitable for reuse across sectors. Talk to finn if you want help with the use of these.

## Setup

### 1) Conda env install
```bash
cd leap_utilities
conda env create -f environment.yml
conda activate leap_utilities
pip install -e .
```

## Using in another repo
```python
from leap_utils.leap_core import connect_to_leap, build_expr
from leap_utils.leap_excel_io import finalise_export_df
from leap_utils.energy_use_reconciliation import build_branch_rules_from_mapping
```
These utilities were designed first for transport applications, so some functions accept transport-specific mappings but they are not required (e.g., vehicle types, modes). 

## Modules
- `leap_core`: COM helpers, expression building, branch creation/fill utilities (transport mappings optional/injectable).
- `leap_excel_io`: helpers to build LEAP import Excel files and merge/view sheets.
- `energy_use_reconciliation`: ESTO/LEAP reconciliation helpers (transport checks optional).

## Notes
- Requires Windows/pywin32 for COM access.
- Keep transport-specific mappings in your transport repo and inject them as needed; this package stays generic.
- If struggling talk to finn, he understands that it might be tricky! He shared this to try and encourage sharing of methods within the team.
- If you don't want to install, add the repo root to `PYTHONPATH`/`sys.path` before importing `leap_utils`, but `pip install -e .` is recommended.


# Industry example:
`leap_utils/industry_mapping_example.py` shows the minimal pattern for moving data between LEAP industry models using an Excel export/import mapping (the same format you get from LEAP’s `Analysis > Export to Excel Template`). You can generate that file in LEAP or build it yourself in the same shape (Branch Path, Variable, Scenario, Region, Scale, Units, Per..., years…).

### How to use the example:
- Open `leap_utils/industry_mapping_example.py` and point `leap_export_filename` to your mapping file (export from source model, or a custom file structured like a LEAP import/export sheet).
- Set `SCENARIO` and `REGION` to the target values in the destination LEAP area; adjust `sheet_name` if your Excel sheet differs from `"Export"`.
- If you need to create the branch structure in the destination model, set `CREATE_BRANCHES_FROM_EXPORT_FILE = True` (uses `create_branches_from_export_file`).
- To write the data into existing branches, keep `FILL_BRANCHES_FROM_EXPORT_FILE = True` (uses `fill_branches_from_export_file`) and optional `SET_UNITS=True` to carry over units from the sheet. > note the issue with setting scale values from the sheet that requires a manual fix within LEAP (see code comments).
- Run the script after making sure your Python environment is ready (e.g. pywin32 is installed) and LEAP is installed and open in the right area, region and scenario, with the right Fuels set. The helper will connect via `connect_to_leap()`, then create/fill branches based on your file.

### Notes/ideas:
- The same pattern works for other sectors—swap in a different export file or build one programmatically (see usage in the APERC `leap_transport` and `power_fish` repos).
- For percentage/share variables you may need to confirm the Scale in the LEAP GUI after import (e.g., set unit to “share” so LEAP assigns the correct scale).

Image below shows the end result of running the example script to copy data from the LEAP industry model (i.e. USA industry area) to the LEAP transport model (i.e. USA transport area), creating branches as needed and filling in data from the export file. It also shows how the scale and units are set correctly for the variables imported - after a manual fix for the scale issue mentioned above.

![image showing usa transport model with industry model in leap](docs/images/usa-transport-industry.png)

# Balance tables example:
This was a quick project to generate balance tables from the 9th edition energy dataset. See `leap_utils/balance_table_example.py` for an example of how to use the `copy_energy_spreadsheet_into_leap_import_file` module to build balance tables within LEAP for checking against the ESTO data while modelling. The script connects to LEAP, extracts energy use data, and generates branches and data within the assumptions folder for this.

![balance table example](docs/images/balance-table-example.png)

# Common issues:
- Units need to be manually set within the LEAP GUI to ensure correct scale value if it is not already. This is because it seems that when we use the create_branches_from_export_file() funciton to create branches, they seem to default to some unknown value that seems to be making LEAP project incorrect values. See Industry example comments for more details.