#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


def norm(p: str) -> Path:
    return Path(p.replace("\\\\", "/")).expanduser().resolve()


def parse_report_ini(path: Path) -> dict:
    lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()
    sections: dict[str, dict[str, str]] = {}
    current: str | None = None
    for line in lines:
        m = re.match(r"^\[(.*)\]$", line.strip())
        if m:
            current = m.group(1)
            sections[current] = {}
            continue
        if current and "=" in line:
            k, v = line.split("=", 1)
            sections[current][k.strip()] = v.strip()

    favorites = []
    for name, data in sections.items():
        if name.startswith("_"):
            continue
        favorites.append(
            {
                "favorite_key": name,
                "chart_type": data.get("ChartType"),
                "result_variable_id": data.get("ResultVariableID"),
                "x_axis_id": data.get("XAxisID"),
                "legend_id": data.get("LegendID"),
                "parentbranchid": data.get("parentbranchid"),
                "unitid": data.get("unitid"),
                "scenario_subset": data.get("ScenarioSubset"),
                "region_subset": data.get("RegionSubset"),
                "region_id": data.get("RegionID"),
                "fuel_id": data.get("FuelID"),
                "input_fuel_type_id": data.get("InputFuelTypeID"),
                "fav_notes": data.get("FavNotes"),
                "raw": data,
            }
        )

    return {
        "source_report": str(path),
        "section_count": len(sections),
        "favorite_section_count": len(favorites),
        "favorites": favorites,
    }


def main() -> None:
    ap = argparse.ArgumentParser(description="Read-only extractor for favorite-like records in ReportINI.txt")
    ap.add_argument("--area", required=True, help="LEAP area folder")
    ap.add_argument("--out", required=True, help="Output JSON path")
    args = ap.parse_args()

    area = norm(args.area)
    report = area / "ReportINI.txt"
    if not report.exists():
        raise FileNotFoundError(f"Missing ReportINI.txt in {area}")

    data = parse_report_ini(report)
    out = norm(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"Wrote {out}")
    print(f"Favorite-like sections: {data['favorite_section_count']}")


if __name__ == "__main__":
    main()
