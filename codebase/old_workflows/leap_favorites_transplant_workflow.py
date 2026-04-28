#%%
from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path
import re
from typing import Iterable

REPO_ROOT = Path(__file__).resolve().parents[1]
if REPO_ROOT.exists() and str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))


# Notebook-focused configuration block.
# Edit these values in a notebook cell (or directly in this file), then call:
#   run_notebook_workflow()
RUN_MODE = "transplant"  # transplant | rollback
SOURCE_AREA = r"C:\LEAP_Areas\clean slate - python version (Recovered 04-07-26) - Copy"
DESTINATION_AREA = r"C:\LEAP_Areas\clean slate daniel elec fixing - Copy"
TARGET_AREA = r""
BACKUP_DIR = r""
IN_PLACE = False
CREATE_WORKING_COPY = True
WORKING_COPY_SUFFIX = " - CodexWorkingCopy-"
OVERWRITE_EXISTING_WORKING_COPY = False
DRY_RUN = True


def _resolve(path_like: str | Path) -> Path:
    text = str(path_like).replace("\\", "/").strip()
    m = re.match(r"^([A-Za-z]):/(.*)$", text)
    if m:
        drive = m.group(1).lower()
        rest = m.group(2)
        text = f"/mnt/{drive}/{rest}"
    p = Path(text)
    if not p.is_absolute():
        p = (REPO_ROOT / p).resolve()
    return p.resolve()


def _now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d-%H%M%S")


def _parse_report_sections(path: Path) -> tuple[list[str], list[tuple[str, list[str]]]]:
    lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()
    preamble: list[str] = []
    sections: list[tuple[str, list[str]]] = []
    current_name: str | None = None
    buf: list[str] = []

    for line in lines:
        m = re.match(r"^\[(.*)\]$", line.strip())
        if m:
            if current_name is None and not sections:
                preamble = buf
            if current_name is not None:
                sections.append((current_name, buf))
            current_name = m.group(1)
            buf = [line]
        else:
            buf.append(line)

    if current_name is not None:
        sections.append((current_name, buf))

    return preamble, sections


def _write_report_sections(path: Path, preamble: list[str], sections: list[tuple[str, list[str]]]) -> None:
    out: list[str] = []
    out.extend(preamble)
    if out and out[-1] != "":
        out.append("")

    for idx, (_, block) in enumerate(sections):
        out.extend(block)
        if idx != len(sections) - 1 and (not out or out[-1] != ""):
            out.append("")

    path.write_text("\r\n".join(out) + "\r\n", encoding="utf-8", errors="ignore")


def _count_nondefault_sections(sections: Iterable[tuple[str, list[str]]]) -> int:
    return sum(1 for name, _ in sections if not name.startswith("_"))


def _copytree_overwrite(src: Path, dst: Path) -> None:
    if dst.exists():
        shutil.rmtree(dst)
    shutil.copytree(src, dst)


def _transplant(
    source_area: Path,
    destination_area: Path,
    *,
    create_working_copy: bool,
    in_place: bool,
    working_copy_suffix: str,
    overwrite_existing_working_copy: bool,
    dry_run: bool,
) -> dict:
    if not source_area.exists() or not source_area.is_dir():
        raise FileNotFoundError(f"Source area not found: {source_area}")
    if not destination_area.exists() or not destination_area.is_dir():
        raise FileNotFoundError(f"Destination area not found: {destination_area}")

    timestamp = _now_stamp()
    if in_place:
        target_area = destination_area
    else:
        target_area = destination_area.parent / f"{destination_area.name}{working_copy_suffix}{timestamp}"

    source_report = source_area / "ReportINI.txt"
    if not source_report.exists():
        raise FileNotFoundError(f"Missing source ReportINI.txt: {source_report}")

    destination_report = target_area / "ReportINI.txt"

    if create_working_copy and not in_place:
        if target_area.exists():
            if overwrite_existing_working_copy:
                if not dry_run:
                    shutil.rmtree(target_area)
            else:
                raise FileExistsError(
                    f"Working copy already exists: {target_area}. "
                    "Use --overwrite-existing-working-copy to replace it."
                )
        if not dry_run:
            shutil.copytree(destination_area, target_area)

    if not in_place and not target_area.exists() and not dry_run:
        raise RuntimeError(f"Working copy was not created: {target_area}")

    if in_place:
        destination_report = destination_area / "ReportINI.txt"

    if not dry_run and not destination_report.exists():
        raise FileNotFoundError(f"Missing destination ReportINI.txt: {destination_report}")

    source_pre, source_sections = _parse_report_sections(source_report)
    source_map = {name: block for name, block in source_sections}
    source_nondefault_names = [name for name, _ in source_sections if not name.startswith("_")]

    destination_pre: list[str] = []
    destination_sections: list[tuple[str, list[str]]] = []
    if not dry_run:
        destination_pre, destination_sections = _parse_report_sections(destination_report)

    destination_default_kept = [
        (name, block)
        for name, block in destination_sections
        if name.startswith("_") and name.lower() != "_overviews"
    ]

    merged_sections: list[tuple[str, list[str]]] = []
    merged_sections.extend(destination_default_kept)
    if "_Overviews" in source_map:
        merged_sections.append(("_Overviews", source_map["_Overviews"]))
    for name in source_nondefault_names:
        merged_sections.append((name, source_map[name]))

    backup_dir: Path | None = None
    report_backup_path: Path | None = None
    previews_backup_path: Path | None = None

    source_previews_dir = source_area / "FavoritePreviews"
    target_previews_dir = target_area / "FavoritePreviews"

    source_preview_files = sorted([p for p in source_previews_dir.glob("*") if p.is_file()]) if source_previews_dir.exists() else []

    if not dry_run:
        backup_dir = target_area / "_FavoritesTransplantBackup" / timestamp
        backup_dir.mkdir(parents=True, exist_ok=True)

        report_backup_path = backup_dir / "ReportINI.before.txt"
        shutil.copy2(destination_report, report_backup_path)

        if target_previews_dir.exists():
            previews_backup_path = backup_dir / "FavoritePreviews.before"
            shutil.copytree(target_previews_dir, previews_backup_path)

        _write_report_sections(destination_report, destination_pre, merged_sections)

        if source_previews_dir.exists():
            target_previews_dir.mkdir(parents=True, exist_ok=True)
            for src_file in source_preview_files:
                shutil.copy2(src_file, target_previews_dir / src_file.name)

    # post metrics
    post_sections: list[tuple[str, list[str]]] = merged_sections
    destination_nondefault_after = _count_nondefault_sections(post_sections)
    has_overviews_after = any(name == "_Overviews" for name, _ in post_sections)
    destination_preview_count_after = len(list(target_previews_dir.glob("*"))) if (target_previews_dir.exists() and not dry_run) else len(source_preview_files)

    result = {
        "timestamp": timestamp,
        "dry_run": dry_run,
        "source_area": str(source_area),
        "destination_area_input": str(destination_area),
        "target_area": str(target_area),
        "in_place": in_place,
        "create_working_copy": create_working_copy,
        "source_nondefault_sections": len(source_nondefault_names),
        "destination_nondefault_sections_after": destination_nondefault_after,
        "destination_has_overviews_after": has_overviews_after,
        "source_preview_file_count": len(source_preview_files),
        "destination_preview_file_count_after": destination_preview_count_after,
        "report_backup_path": str(report_backup_path) if report_backup_path else None,
        "previews_backup_path": str(previews_backup_path) if previews_backup_path else None,
        "backup_dir": str(backup_dir) if backup_dir else None,
    }

    if not dry_run and backup_dir is not None:
        json_path = backup_dir / "transplant_summary.json"
        md_path = backup_dir / "transplant_steps.md"
        json_path.write_text(json.dumps(result, indent=2), encoding="utf-8")
        md_path.write_text(
            "\n".join(
                [
                    "# Favorites Transplant Steps",
                    "",
                    f"- timestamp: {timestamp}",
                    f"- source_area: {source_area}",
                    f"- destination_input: {destination_area}",
                    f"- target_area: {target_area}",
                    f"- in_place: {in_place}",
                    f"- source_nondefault_sections: {len(source_nondefault_names)}",
                    f"- destination_nondefault_after: {destination_nondefault_after}",
                    f"- destination_preview_files_after: {destination_preview_count_after}",
                    "",
                    "## Actions",
                    "- Backed up destination ReportINI.txt.",
                    "- Rebuilt destination ReportINI favorites blocks from source.",
                    "- Copied FavoritePreviews files from source to target.",
                    "",
                    "## Rollback",
                    f"- Restore {report_backup_path} to target ReportINI.txt.",
                    f"- If previews backup exists, restore {previews_backup_path} to target FavoritePreviews.",
                ]
            ),
            encoding="utf-8",
        )

    return result


def _rollback(target_area: Path, backup_dir: Path, dry_run: bool) -> dict:
    if not target_area.exists():
        raise FileNotFoundError(f"Target area not found: {target_area}")
    if not backup_dir.exists():
        raise FileNotFoundError(f"Backup directory not found: {backup_dir}")

    backup_report = backup_dir / "ReportINI.before.txt"
    if not backup_report.exists():
        raise FileNotFoundError(f"Missing backup report: {backup_report}")

    target_report = target_area / "ReportINI.txt"
    target_previews = target_area / "FavoritePreviews"
    backup_previews = backup_dir / "FavoritePreviews.before"

    if not dry_run:
        shutil.copy2(backup_report, target_report)
        if backup_previews.exists():
            _copytree_overwrite(backup_previews, target_previews)

    return {
        "dry_run": dry_run,
        "target_area": str(target_area),
        "backup_dir": str(backup_dir),
        "restored_report": str(target_report),
        "restored_previews": str(target_previews) if backup_previews.exists() else None,
    }


def run_notebook_workflow(
    *,
    run_mode: str | None = None,
    source_area: str | Path | None = None,
    destination_area: str | Path | None = None,
    target_area: str | Path | None = None,
    backup_dir: str | Path | None = None,
    in_place: bool | None = None,
    create_working_copy: bool | None = None,
    working_copy_suffix: str | None = None,
    overwrite_existing_working_copy: bool | None = None,
    dry_run: bool | None = None,
) -> dict:
    """
    Notebook-first entrypoint.

    Typical notebook usage:
    1) Edit config constants near top of this file, or pass explicit kwargs.
    2) Call run_notebook_workflow().
    """
    mode = run_mode or RUN_MODE
    mode = str(mode).strip().lower()

    if mode == "transplant":
        if source_area is None:
            source_area = SOURCE_AREA
        if destination_area is None:
            destination_area = DESTINATION_AREA
        if in_place is None:
            in_place = IN_PLACE
        if create_working_copy is None:
            create_working_copy = CREATE_WORKING_COPY
        if working_copy_suffix is None:
            working_copy_suffix = WORKING_COPY_SUFFIX
        if overwrite_existing_working_copy is None:
            overwrite_existing_working_copy = OVERWRITE_EXISTING_WORKING_COPY
        if dry_run is None:
            dry_run = DRY_RUN

        result = _transplant(
            source_area=_resolve(source_area),
            destination_area=_resolve(destination_area),
            create_working_copy=bool(create_working_copy),
            in_place=bool(in_place),
            working_copy_suffix=str(working_copy_suffix),
            overwrite_existing_working_copy=bool(overwrite_existing_working_copy),
            dry_run=bool(dry_run),
        )
    elif mode == "rollback":
        if target_area is None:
            target_area = TARGET_AREA
        if backup_dir is None:
            backup_dir = BACKUP_DIR
        if dry_run is None:
            dry_run = DRY_RUN
        if not str(target_area).strip():
            raise ValueError("rollback mode requires target_area (or TARGET_AREA config).")
        if not str(backup_dir).strip():
            raise ValueError("rollback mode requires backup_dir (or BACKUP_DIR config).")

        result = _rollback(
            target_area=_resolve(target_area),
            backup_dir=_resolve(backup_dir),
            dry_run=bool(dry_run),
        )
    else:
        raise ValueError(f"Unsupported run_mode: {mode}")

    print(json.dumps(result, indent=2))
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description="Safe workflow for LEAP favorites transplant.")
    sub = parser.add_subparsers(dest="mode", required=True)

    p_tx = sub.add_parser("transplant", help="Transplant favorites from source area into destination or working copy.")
    p_tx.add_argument("--source-area", required=True)
    p_tx.add_argument("--destination-area", required=True)
    p_tx.add_argument("--in-place", action="store_true", help="Write directly to destination area (unsafe; use only after testing).")
    p_tx.add_argument("--no-working-copy", action="store_true", help="Do not create working copy.")
    p_tx.add_argument(
        "--working-copy-suffix",
        default=" - CodexWorkingCopy-",
        help="Suffix prefix used when creating the working copy folder name.",
    )
    p_tx.add_argument("--overwrite-existing-working-copy", action="store_true")
    p_tx.add_argument("--dry-run", action="store_true")

    p_rb = sub.add_parser("rollback", help="Restore ReportINI and previews from a backup folder.")
    p_rb.add_argument("--target-area", required=True)
    p_rb.add_argument("--backup-dir", required=True)
    p_rb.add_argument("--dry-run", action="store_true")

    args = parser.parse_args()

    if args.mode == "transplant":
        result = _transplant(
            source_area=_resolve(args.source_area),
            destination_area=_resolve(args.destination_area),
            create_working_copy=not args.no_working_copy,
            in_place=args.in_place,
            working_copy_suffix=args.working_copy_suffix,
            overwrite_existing_working_copy=args.overwrite_existing_working_copy,
            dry_run=args.dry_run,
        )
    else:
        result = _rollback(
            target_area=_resolve(args.target_area),
            backup_dir=_resolve(args.backup_dir),
            dry_run=args.dry_run,
        )

    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
