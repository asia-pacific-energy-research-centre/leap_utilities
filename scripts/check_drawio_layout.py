#!/usr/bin/env python3
"""Check draw.io layout for overflows and unintended overlaps."""
from __future__ import annotations

import argparse
import re
from pathlib import Path

ATTR_RE = re.compile(r'(\w+)="([^"]*)"')


def load_cells(path: Path):
    text = path.read_text(encoding='utf-8')
    cells = {}
    geom = {}
    parents = {}

    current_id = None
    for line in text.splitlines():
        if '<mxCell' in line:
            attrs = dict(ATTR_RE.findall(line))
            cid = attrs.get('id')
            if cid:
                cells[cid] = attrs
                parents[cid] = attrs.get('parent')
                if attrs.get('vertex') == '1':
                    current_id = cid
                else:
                    current_id = None
        if current_id and '<mxGeometry' in line:
            attrs = dict(ATTR_RE.findall(line))
            if 'width' in attrs and 'height' in attrs:
                geom[current_id] = {
                    'x': float(attrs.get('x', '0') or 0),
                    'y': float(attrs.get('y', '0') or 0),
                    'w': float(attrs.get('width', '0') or 0),
                    'h': float(attrs.get('height', '0') or 0),
                }
            current_id = None
    return cells, geom, parents


def build_ancestors(parents: dict[str, str | None]):
    ancestors = {}
    for cid in parents:
        chain = []
        cur = parents.get(cid)
        while cur:
            chain.append(cur)
            cur = parents.get(cur)
        ancestors[cid] = chain
    return ancestors


def absolute_positions(geom: dict[str, dict], parents: dict[str, str | None]):
    abs_geom = {}
    for cid, g in geom.items():
        x = g['x']
        y = g['y']
        cur = parents.get(cid)
        while cur:
            if cur in geom:
                x += geom[cur]['x']
                y += geom[cur]['y']
            cur = parents.get(cur)
        abs_geom[cid] = {'x': x, 'y': y, 'w': g['w'], 'h': g['h']}
    return abs_geom


def intersects(a, b):
    return (
        a['x'] < b['x'] + b['w']
        and a['x'] + a['w'] > b['x']
        and a['y'] < b['y'] + b['h']
        and a['y'] + a['h'] > b['y']
    )


def check(path: Path):
    cells, geom, parents = load_cells(path)
    ancestors = build_ancestors(parents)
    abs_geom = absolute_positions(geom, parents)

    overflows = []
    for cid, g in geom.items():
        parent = parents.get(cid)
        if not parent or parent not in geom:
            continue
        pg = geom[parent]
        if (
            g['x'] < 0
            or g['y'] < 0
            or g['x'] + g['w'] > pg['w']
            or g['y'] + g['h'] > pg['h']
        ):
            overflows.append((cid, parent))

    vertex_ids = [cid for cid in geom if cells.get(cid, {}).get('vertex') == '1']
    collisions = []
    for i, a in enumerate(vertex_ids):
        for b in vertex_ids[i + 1 :]:
            if b in ancestors.get(a, []) or a in ancestors.get(b, []):
                continue
            if intersects(abs_geom[a], abs_geom[b]):
                collisions.append((a, b))

    return overflows, collisions


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('path', type=Path)
    args = parser.parse_args()

    overflows, collisions = check(args.path)
    print(f'Overflows: {len(overflows)}')
    for cid, parent in overflows:
        print(f'- {cid} overflows {parent}')

    print(f'Collisions: {len(collisions)}')
    for a, b in collisions:
        print(f'- {a} intersects {b}')

    if overflows or collisions:
        raise SystemExit(1)


if __name__ == '__main__':
    main()
