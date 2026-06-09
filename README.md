# leap-utilities

Shared LEAP helpers (branch utilities, Excel import/export, energy-use reconciliation and LEAP API interactions). It was originally intended to be flexible and reusable but perhaps is just a bit of a grab-bag of utilities that have been built up over time for interacting with LEAP in various ways.

A large portion of the codebase was developed for setting up LEAP areas for the APERC 10th edition with the intended modelling structure of that project in mind. A lot of code is built to extract and process ESTO data to put it into the strucutre used by the LEAP areas. This uses a lot of mappings, which have been built for reuse in other similar applications. 

Also there is a large workflow called results_supply_link_workflow which combines the demand, transformation and supply workflows to create a full end-to-end workflow for initialising the data for a new LEAP area and iterating on its results to get to a good state.

## Setup

### Prerequisites

- **Windows** — LEAP uses a COM/win32 interface that only works on Windows. Note that as of 9June2026 we consider the LEAP API too buggy to use, so we have switched to using Excel import/export as the main method of moving data in and out of LEAP, and use manual processes for other API tasks like branch creation.
- **LEAP** — must be installed and licensed on the machine. The workflows read and write LEAP data through its COM API and through exported Excel workbooks.
- **Python 3.11** — via conda (recommended) or a standalone install.
- **Both repos cloned as siblings** — `leap_utilities` and `leap_dashboard` must sit in the same parent folder if you want to use leap_dashboard to create dashboards (e.g. `github/leap_utilities` and `github/leap_dashboard`). The dashboard workflow finds `leap_utilities` by looking one level up.

### 1) Install dependencies

#### Option A — conda (recommended)

```bash
cd leap_utilities
conda env create -f environment.yml
conda activate leap_utilities
```

#### Option B — pip only

```bash
cd leap_utilities
pip install -e .
# environment.yml lists the same deps: pandas openpyxl matplotlib pywin32
```

Do the same in `leap_dashboard` if you are running the dashboard workflow:

```bash
cd leap_dashboard
pip install -e .
# deps: pandas openpyxl plotly jinja2
```

### 2) Data files required

The main workflows expect the following files to already be present (they are not generated — they come from external datasets or LEAP exports):

| File | Where used |
| ---- | ---------- |
| `data/00APEC_2025_low_with_subtotals.csv` | ESTO historical reference |
| `data/merged_file_energy_ALL_20251106.csv` | 9th Outlook projection data |
| `data/full model export.xlsx` | Branch path reference template |
| `data/leap balances exports/<economy>/` | LEAP balance exports (manual export from LEAP) |

See `data/README.md` for full descriptions and `data/leap balances exports/README.md` for the expected filename format of LEAP exports.

### 3) Running `results_supply_link_workflow.py`

This is the main workflow for syncing a new LEAP area. No environment variables need to be set — the repo root is resolved automatically from the file location.

Open the script, check the `ACTIVE_PRESET` at the bottom (either `_PRESET_BASELINE_SEED` or `_PRESET_RESULTS_UPDATE`), set the economy and scenario in `codebase/configuration/workflow_config.py`, then run the script. See the *Syncing a new area* section of the researcher guide for the full iterative process.

### Troubleshooting

- **`ImportError: No module named 'win32com'`** — pywin32 is not installed or the wrong Python environment is active.
- **`FileNotFoundError` on a data file** — check `data/README.md` for which files need to be present before running.
- **LEAP COM errors on startup** — LEAP must be open and the correct area/scenario must be active before running any workflow that uses the COM API.

## Modules

- `leap_core`: COM helpers, expression building, branch creation/fill utilities (transport mappings optional/injectable).
- `leap_excel_io`: helpers to build LEAP import Excel files and merge/view sheets.
- `leap_exports`: packaged helpers for export filename formatting, workbook creation, and workbook discovery/validation.
- `leap_api`: packaged helpers for LEAP API availability checks and workbook import operations.
- `energy_use_reconciliation`: ESTO/LEAP reconciliation helpers (transport checks optional).
- `power_workflow`: standalone power import workflow for `data/power export.xlsx` with scenario alignment, ESTO fuel validation, skip reporting, and hardcoded override hooks. > not completed yet but a work in progress.

## Notes

- Requires Windows/pywin32 for COM access.
- If struggling talk to finn, he understands that it might be tricky!
- If you don't want to install, add the repo root to `PYTHONPATH`/`sys.path` before importing `code`, but `pip install -e .` is recommended.

# Industry example:

Note that this pat of the code is redundant now that Industry model has been fully migrated to the new LEAP structure, but it is left here as an example of how to use the Excel import/export pattern for moving data between LEAP models. The `codebase/industry_mapping_workflow.py` shows the minimal pattern for moving data between LEAP industry models using an Excel export/import mapping (the same format you get from LEAP’s `Analysis > Export to Excel Template`). You can generate that file in LEAP or build it yourself in the same shape (Branch Path, Variable, Scenario, Region, Scale, Units, Per..., years…). Could be applied to other sectors too, but industry was the original use case for this pattern so it’s the example here.

### How to use the example:

- Open `codebase/industry_mapping_workflow.py` and point `leap_export_filename` to your mapping file (export from source model, or a custom file structured like a LEAP import/export sheet).
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

This was a quick project to generate balance tables from the 9th edition energy dataset. See `codebase/balance_table_example.py` for an example of how to use the `copy_energy_spreadsheet_into_leap_import_file` module to build balance tables within LEAP for checking against the ESTO data while modelling. The script connects to LEAP, extracts energy use data, and generates branches and data within the assumptions folder for this.

![balance table example](docs/images/balance-table-example.png)

# Common issues:

- Units need to be manually set within the LEAP GUI to ensure correct scale value if it is not already. This is because it seems that when we use the create_branches_from_export_file() funciton to create branches, they seem to default to some unknown value that seems to be making LEAP project incorrect values. See Industry example comments for more details.
