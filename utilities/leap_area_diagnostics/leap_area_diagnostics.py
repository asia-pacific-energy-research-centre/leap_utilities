#!/usr/bin/env python3
"""Read-only diagnostics for comparing two LEAP area folders."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path
from typing import Callable


def normalize_path(raw: str) -> Path:
    return Path(raw.replace("\\\\", "/")).resolve()


def iter_files(root: Path):
    for dp, _, fns in os.walk(root):
        dpp = Path(dp)
        for fn in fns:
            p = dpp / fn
            try:
                st = p.stat()
            except OSError:
                continue
            rel = str(p.relative_to(root)).replace("\\", "/")
            yield p, rel, st.st_size


def inventory(root: Path, top_n: int = 50) -> dict:
    ext_count = Counter()
    ext_size = Counter()
    dir_size = defaultdict(int)
    files = []
    total_size = 0
    total_count = 0

    for _, rel, size in iter_files(root):
        total_size += size
        total_count += 1
        ext = Path(rel).suffix.lower() or "[no_ext]"
        ext_count[ext] += 1
        ext_size[ext] += size
        parent = Path(rel).parent
        if not parent.parts:
            dir_size["."] += size
        else:
            cur = []
            for part in parent.parts:
                cur.append(part)
                dir_size["/".join(cur)] += size
            dir_size["."] += size
        files.append({"rel_path": rel, "size": size, "ext": ext})

    largest = sorted(files, key=lambda x: x["size"], reverse=True)[:top_n]
    return {
        "root": str(root),
        "total_size": total_size,
        "file_count": total_count,
        "ext_count": dict(ext_count),
        "ext_size": dict(ext_size),
        "dir_size": dict(dir_size),
        "largest": largest,
        "files": files,
    }


def diff_trees(a: dict, b: dict, out_csv: Path | None = None) -> dict:
    fa = {r["rel_path"]: r for r in a["files"]}
    fb = {r["rel_path"]: r for r in b["files"]}
    only_a = sorted(set(fa) - set(fb))
    only_b = sorted(set(fb) - set(fa))

    deltas = []
    for rel in sorted(set(fa) | set(fb)):
        sa = fa.get(rel, {}).get("size", 0)
        sb = fb.get(rel, {}).get("size", 0)
        if sa != sb:
            deltas.append({"rel_path": rel, "size_a": sa, "size_b": sb, "delta_b_minus_a": sb - sa})
    deltas.sort(key=lambda r: abs(r["delta_b_minus_a"]), reverse=True)

    if out_csv:
        with out_csv.open("w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=["rel_path", "size_a", "size_b", "delta_b_minus_a"])
            w.writeheader()
            w.writerows(deltas)

    return {
        "only_in_a": only_a,
        "only_in_b": only_b,
        "size_deltas": deltas,
    }


def inspect_db_schema(root: Path, include_ext: set[str]) -> dict:
    rows = []
    for p, rel, size in iter_files(root):
        ext = p.suffix.lower()
        if ext not in include_ext:
            continue
        rec = {"path": rel, "size": size, "ext": ext, "sqlite": False}
        try:
            con = sqlite3.connect(f"file:{p}?mode=ro", uri=True)
            cur = con.cursor()
            tables = [r[0] for r in cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY 1")]
            rec["sqlite"] = True
            rec["tables"] = tables
            rec["row_counts"] = {}
            for t in tables:
                try:
                    rec["row_counts"][t] = cur.execute(f'SELECT COUNT(*) FROM "{t}"').fetchone()[0]
                except sqlite3.Error:
                    pass
            con.close()
        except Exception as e:  # noqa: BLE001
            rec["error"] = str(e)
        rows.append(rec)
    return {"root": str(root), "db_candidates": rows}


def search_keywords(root: Path, regex: re.Pattern, include_ext: set[str]) -> dict:
    hits = {}
    for p, rel, _ in iter_files(root):
        if p.suffix.lower() not in include_ext:
            continue
        try:
            text = p.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        m = regex.findall(text)
        if m:
            c = Counter(x.lower() for x in m)
            hits[rel] = dict(c)
    return {"root": str(root), "hits": hits}


def hash_files(root: Path, algo: str = "sha256") -> dict:
    out = []
    hfn: Callable[[], "hashlib._Hash"] = getattr(hashlib, algo)
    for p, rel, _ in iter_files(root):
        h = hfn()
        try:
            with p.open("rb") as f:
                for chunk in iter(lambda: f.read(1024 * 1024), b""):
                    h.update(chunk)
        except OSError:
            continue
        out.append({"rel_path": rel, "hash": h.hexdigest()})
    return {"root": str(root), "algo": algo, "hashes": out}


def main() -> None:
    parser = argparse.ArgumentParser(description="Read-only LEAP area diagnostics")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_inv = sub.add_parser("inventory")
    p_inv.add_argument("--area", action="append", required=True)
    p_inv.add_argument("--top", type=int, default=50)

    p_diff = sub.add_parser("diff")
    p_diff.add_argument("--area-a", required=True)
    p_diff.add_argument("--area-b", required=True)
    p_diff.add_argument("--csv", default="")

    p_db = sub.add_parser("schema")
    p_db.add_argument("--area", action="append", required=True)
    p_db.add_argument("--ext", default=".db,.sqlite,.sqlite3,.mdb,.accdb,.sdf,.nx1,.bin")

    p_kw = sub.add_parser("keywords")
    p_kw.add_argument("--area", action="append", required=True)
    p_kw.add_argument("--pattern", required=True)
    p_kw.add_argument("--ext", default=".txt,.ini,.cfg,.conf,.xml,.json,.vbs,.js,.py,.csv")

    p_hash = sub.add_parser("hash")
    p_hash.add_argument("--area", action="append", required=True)
    p_hash.add_argument("--algo", default="sha256")

    args = parser.parse_args()

    if args.cmd == "inventory":
        print(json.dumps([inventory(normalize_path(a), top_n=args.top) for a in args.area], indent=2))
    elif args.cmd == "diff":
        inv_a = inventory(normalize_path(args.area_a), top_n=0)
        inv_b = inventory(normalize_path(args.area_b), top_n=0)
        out_csv = Path(args.csv) if args.csv else None
        print(json.dumps(diff_trees(inv_a, inv_b, out_csv=out_csv), indent=2))
    elif args.cmd == "schema":
        include_ext = {e.strip().lower() for e in args.ext.split(",") if e.strip()}
        print(json.dumps([inspect_db_schema(normalize_path(a), include_ext) for a in args.area], indent=2))
    elif args.cmd == "keywords":
        include_ext = {e.strip().lower() for e in args.ext.split(",") if e.strip()}
        regex = re.compile(args.pattern, re.IGNORECASE)
        print(json.dumps([search_keywords(normalize_path(a), regex, include_ext) for a in args.area], indent=2))
    elif args.cmd == "hash":
        print(json.dumps([hash_files(normalize_path(a), algo=args.algo) for a in args.area], indent=2))


if __name__ == "__main__":
    main()
