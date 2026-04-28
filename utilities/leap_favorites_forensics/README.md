# LEAP Favorites Forensics Helpers (Read-Only)

## Python

- Inventory:
  - `python utilities/leap_favorites_forensics/leap_area_forensics.py inventory --area "C:\\LEAP_Areas\\MyArea" --outdir "outputs\\leap_forensics\\myarea"`
- Hash files:
  - `python utilities/leap_favorites_forensics/leap_area_forensics.py hash --area "C:\\LEAP_Areas\\MyArea" --outdir "outputs\\leap_forensics\\myarea"`
- Search text + binary strings:
  - `python utilities/leap_favorites_forensics/leap_area_forensics.py search --area "C:\\LEAP_Areas\\MyArea" --outdir "outputs\\leap_forensics\\myarea"`
- SQLite schema probe:
  - `python utilities/leap_favorites_forensics/leap_area_forensics.py sqlite-probe --area "C:\\LEAP_Areas\\MyArea" --outdir "outputs\\leap_forensics\\myarea"`
- Diff two areas:
  - `python utilities/leap_favorites_forensics/leap_area_forensics.py diff --area-a "C:\\LEAP_Areas\\AreaA" --area-b "C:\\LEAP_Areas\\AreaB" --outdir "outputs\\leap_forensics\\diff_A_B"`

## PowerShell

- Inventory:
  - `pwsh utilities/leap_favorites_forensics/leap_area_forensics.ps1 -Mode inventory -Area "C:\\LEAP_Areas\\MyArea" -OutDir "outputs\\leap_forensics\\myarea"`
- Hash:
  - `pwsh utilities/leap_favorites_forensics/leap_area_forensics.ps1 -Mode hash -Area "C:\\LEAP_Areas\\MyArea" -OutDir "outputs\\leap_forensics\\myarea"`
- Search:
  - `pwsh utilities/leap_favorites_forensics/leap_area_forensics.ps1 -Mode search -Area "C:\\LEAP_Areas\\MyArea" -OutDir "outputs\\leap_forensics\\myarea"`
- SQLite probe:
  - `pwsh utilities/leap_favorites_forensics/leap_area_forensics.ps1 -Mode sqlite-probe -Area "C:\\LEAP_Areas\\MyArea" -OutDir "outputs\\leap_forensics\\myarea"`
- Diff:
  - `pwsh utilities/leap_favorites_forensics/leap_area_forensics.ps1 -Mode diff -AreaA "C:\\LEAP_Areas\\AreaA" -AreaB "C:\\LEAP_Areas\\AreaB" -OutDir "outputs\\leap_forensics\\diff_A_B"`

## Favorites PoC Extractor

- `python utilities/leap_favorites_forensics/extract_report_favorites.py --area "C:\\LEAP_Areas\\MyArea" --out "outputs\\leap_forensics\\myarea\\favorites_from_reportini.json"`

This extractor is read-only and only parses `ReportINI.txt`.
