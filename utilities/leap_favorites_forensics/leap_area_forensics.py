#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import sqlite3
import subprocess
from pathlib import Path
from typing import Iterable

TEXT_EXTS = {".txt", ".ini", ".cfg", ".json", ".xml", ".md", ".log", ".csv"}
DB_EXTS = {".db", ".sqlite", ".sqlite3", ".mdb", ".accdb", ".nx1"}
KEYWORDS = [
    "favorite",
    "favourites",
    "favorites",
    "fave",
    "chart",
    "table",
    "results",
    "foldername",
    "favename",
]


def norm(p: str) -> Path:
    return Path(p.replace("\\\\", "/")).expanduser().resolve()


def iter_files(root: Path) -> Iterable[Path]:
    for p in root.rglob("*"):
        if p.is_file():
            yield p


def sha256(path: Path, chunk: int = 1024 * 1024) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def inventory(area: Path, outdir: Path) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    inv_path = outdir / "inventory.tsv"
    ext_path = outdir / "extension_summary.tsv"
    candidates_path = outdir / "candidate_files.tsv"
    recent_path = outdir / "recent_files.tsv"
    small_path = outdir / "small_files_lt20k.tsv"

    rows = []
    ext_counts: dict[str, int] = {}
    for p in iter_files(area):
        st = p.stat()
        ext = p.suffix.lower() if p.suffix else "[noext]"
        ext_counts[ext] = ext_counts.get(ext, 0) + 1
        rows.append(
            {
                "type": "f",
                "size": st.st_size,
                "mtime_epoch": st.st_mtime,
                "mtime_iso": st.st_mtime_ns,
                "path": str(p),
                "ext": ext,
            }
        )
    for p in area.rglob("*"):
        if p.is_dir():
            st = p.stat()
            rows.append(
                {
                    "type": "d",
                    "size": 0,
                    "mtime_epoch": st.st_mtime,
                    "mtime_iso": st.st_mtime_ns,
                    "path": str(p),
                    "ext": "",
                }
            )

    with inv_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["type", "size", "mtime_epoch", "path"])
        for r in sorted(rows, key=lambda x: x["path"]):
            w.writerow([r["type"], r["size"], f"{r['mtime_epoch']:.6f}", r["path"]])

    with ext_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["extension", "count"])
        for ext, count in sorted(ext_counts.items(), key=lambda kv: (-kv[1], kv[0])):
            w.writerow([ext, count])

    candidate_exts = {
        ".db",
        ".sqlite",
        ".sqlite3",
        ".mdb",
        ".accdb",
        ".xml",
        ".json",
        ".ini",
        ".cfg",
        ".txt",
        ".dat",
        ".bin",
        ".zip",
        ".7z",
        ".nx1",
    }
    with candidates_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["size", "mtime_epoch", "ext", "path"])
        for p in iter_files(area):
            ext = p.suffix.lower()
            if ext in candidate_exts:
                st = p.stat()
                w.writerow([st.st_size, f"{st.st_mtime:.6f}", ext or "[noext]", str(p)])

    files = [p for p in iter_files(area)]
    with recent_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["mtime_epoch", "size", "path"])
        for p in sorted(files, key=lambda q: q.stat().st_mtime, reverse=True):
            st = p.stat()
            w.writerow([f"{st.st_mtime:.6f}", st.st_size, str(p)])

    with small_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["size", "mtime_epoch", "path"])
        for p in sorted(files, key=lambda q: q.stat().st_size):
            st = p.stat()
            if st.st_size < 20 * 1024:
                w.writerow([st.st_size, f"{st.st_mtime:.6f}", str(p)])


def hash_files(area: Path, outdir: Path) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    out = outdir / "file_hashes_sha256.tsv"
    with out.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["sha256", "size", "path"])
        for p in sorted(iter_files(area)):
            st = p.stat()
            w.writerow([sha256(p), st.st_size, str(p)])


def looks_text(path: Path) -> bool:
    if path.suffix.lower() in TEXT_EXTS:
        return True
    try:
        b = path.read_bytes()[:4096]
    except Exception:
        return False
    if not b:
        return True
    if b"\x00" in b:
        return False
    bad = sum(1 for x in b if x < 9 or (13 < x < 32))
    return bad / len(b) < 0.05


def search(area: Path, outdir: Path, terms: list[str]) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    text_out = outdir / "text_matches.tsv"
    bin_out = outdir / "binary_strings_matches.tsv"
    pattern = re.compile("|".join(re.escape(t) for t in terms), re.IGNORECASE)

    with text_out.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["path", "line_no", "line"])
        for p in sorted(iter_files(area)):
            if not looks_text(p):
                continue
            try:
                text = p.read_text(encoding="utf-8", errors="ignore")
            except Exception:
                continue
            for i, line in enumerate(text.splitlines(), 1):
                if pattern.search(line):
                    w.writerow([str(p), i, line.strip()])

    with bin_out.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["path", "string_line_no", "string"])
        for p in sorted(iter_files(area)):
            if looks_text(p):
                continue
            try:
                cp = subprocess.run(
                    ["strings", "-a", "-n", "4", str(p)],
                    check=False,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="ignore",
                )
            except FileNotFoundError:
                break
            if cp.returncode not in (0, 1):
                continue
            for i, line in enumerate(cp.stdout.splitlines(), 1):
                if pattern.search(line):
                    w.writerow([str(p), i, line.strip()])


def sqlite_probe(area: Path, outdir: Path, terms: list[str]) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    out = outdir / "sqlite_probe.json"
    data: dict[str, object] = {"databases": []}
    pattern = re.compile("|".join(re.escape(t) for t in terms), re.IGNORECASE)
    for p in sorted(iter_files(area)):
        low = p.suffix.lower()
        if low not in {".db", ".sqlite", ".sqlite3"}:
            continue
        db_info: dict[str, object] = {"path": str(p), "tables": []}
        try:
            con = sqlite3.connect(f"file:{p}?mode=ro", uri=True)
            cur = con.cursor()
            tables = [r[0] for r in cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")]
            for t in tables:
                cols = [r[1] for r in cur.execute(f"PRAGMA table_info('{t}')")]
                hit = bool(pattern.search(t) or any(pattern.search(c) for c in cols))
                sample = []
                if hit:
                    for row in cur.execute(f"SELECT * FROM '{t}' LIMIT 5"):
                        sample.append([str(x)[:500] for x in row])
                db_info["tables"].append({"name": t, "columns": cols, "keyword_hit": hit, "sample_rows": sample})
            con.close()
        except Exception as e:
            db_info["error"] = str(e)
        data["databases"].append(db_info)
    out.write_text(json.dumps(data, indent=2), encoding="utf-8")


def diff_areas(area_a: Path, area_b: Path, outdir: Path) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    out = outdir / "area_diff.tsv"

    def fmap(root: Path) -> dict[str, Path]:
        return {p.relative_to(root).as_posix(): p for p in iter_files(root)}

    ma, mb = fmap(area_a), fmap(area_b)
    keys = sorted(set(ma) | set(mb))
    with out.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["status", "relpath", "size_a", "size_b", "sha256_a", "sha256_b"])
        for rel in keys:
            pa, pb = ma.get(rel), mb.get(rel)
            if pa is None:
                w.writerow(["only_in_b", rel, "", pb.stat().st_size, "", sha256(pb)])
                continue
            if pb is None:
                w.writerow(["only_in_a", rel, pa.stat().st_size, "", sha256(pa), ""])
                continue
            sa, sb = pa.stat().st_size, pb.stat().st_size
            if sa != sb:
                w.writerow(["size_diff", rel, sa, sb, "", ""])
            else:
                ha, hb = sha256(pa), sha256(pb)
                if ha != hb:
                    w.writerow(["hash_diff", rel, sa, sb, ha, hb])


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Read-only LEAP area forensics helper")
    sub = p.add_subparsers(dest="cmd", required=True)

    p_inv = sub.add_parser("inventory")
    p_inv.add_argument("--area", required=True)
    p_inv.add_argument("--outdir", required=True)

    p_hash = sub.add_parser("hash")
    p_hash.add_argument("--area", required=True)
    p_hash.add_argument("--outdir", required=True)

    p_search = sub.add_parser("search")
    p_search.add_argument("--area", required=True)
    p_search.add_argument("--outdir", required=True)
    p_search.add_argument("--terms", nargs="*", default=KEYWORDS)

    p_sql = sub.add_parser("sqlite-probe")
    p_sql.add_argument("--area", required=True)
    p_sql.add_argument("--outdir", required=True)
    p_sql.add_argument("--terms", nargs="*", default=KEYWORDS)

    p_diff = sub.add_parser("diff")
    p_diff.add_argument("--area-a", required=True)
    p_diff.add_argument("--area-b", required=True)
    p_diff.add_argument("--outdir", required=True)

    return p.parse_args()


def main() -> None:
    args = parse_args()
    if args.cmd == "inventory":
        inventory(norm(args.area), norm(args.outdir))
    elif args.cmd == "hash":
        hash_files(norm(args.area), norm(args.outdir))
    elif args.cmd == "search":
        search(norm(args.area), norm(args.outdir), args.terms)
    elif args.cmd == "sqlite-probe":
        sqlite_probe(norm(args.area), norm(args.outdir), args.terms)
    elif args.cmd == "diff":
        diff_areas(norm(args.area_a), norm(args.area_b), norm(args.outdir))
    else:
        raise ValueError(args.cmd)


if __name__ == "__main__":
    main()
