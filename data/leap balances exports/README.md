# LEAP Balance Exports

This folder stores pre-scraped LEAP energy balance workbooks used by the balance
table and results-supply workflows. Keep one economy folder per LEAP economy
code, for example:

```text
data/leap balances exports/20_USA/
```

## Quick Extraction Summary

For now, LEAP results are extracted manually from the LEAP Energy Balance table
view rather than through the LEAP API. Export the Energy Balance results to
Excel, with separate balance sheets for the relevant scenario/year combinations,
then save the exported workbooks in the economy folder here.

The Python workflows read those exported workbooks directly. The extractor uses
the sheet layout as follows:

- cell `A1` identifies the LEAP area
- cell `A2` identifies scenario, year, and units
- row `3` contains fuel columns
- column `A`, from row `4` down, contains LEAP balance rows/sectors
- each numeric cell becomes one long-format record

The intermediate table keeps source workbook, source sheet, area, scenario,
year, LEAP sector, LEAP fuel, units, and value. Values are converted to
petajoules, then mapped to ESTO flow/product pairs using `config/leap_mappings.xlsx`.
The mapped long table is used by `codebase/leap_balance_to_esto_long_workflow.py`,
the dashboard workflows, and `codebase/results_supply_link_workflow.py`.

## Expected Filename Format

Active workbooks should use this filename pattern:

```text
full model output all years <date_id> <scenario_code>.xlsx
```

Examples:

```text
full model output all years 04092026 REF.xlsx
full model output all years 04092026 TGT.xlsx
```

Scenario codes are resolved as:

- `REF` or `Reference` -> `REF`
- `TGT` or `Target` -> `TGT`

## Date IDs

The resolver accepts compact date IDs for latest-file selection:

- `492026` means April 9, 2026.
- `04092026` also means April 9, 2026.
- `4212026` means April 21, 2026.
- `04212026` also means April 21, 2026.

When no date ID is pinned in workflow code, the resolver parses the date IDs
and selects the latest workbook for the requested economy and scenario.

When a workflow pins a date ID explicitly, the filename token must match
exactly. For example, `date_id="4292026"` matches `4292026`, not `04292026`.
Use the exact token in the filename for reproducible reruns.

## Archive Folder

The resolver only scans workbook files directly inside the economy folder, such
as `20_USA/`. Files under `archive/` are ignored. Move older or superseded
workbooks into `archive/` when they should not be selected by default.

## Code Path

Workbook selection is handled by:

```text
codebase/utilities/leap_balance_export_resolver.py
```

Workflow constants normally look like this:

```python
BALANCE_EXPORT_ECONOMY = "20_USA"
REF_BALANCE_EXPORT_DATE_ID = None
TGT_BALANCE_EXPORT_DATE_ID = None
```

Set `REF_BALANCE_EXPORT_DATE_ID` or `TGT_BALANCE_EXPORT_DATE_ID` only when you
need to pin a specific workbook date.
