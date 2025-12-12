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
