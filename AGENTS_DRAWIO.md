# AGENTS_DRAWIO.md
#
# Draw.io specific instructions for this repo.

## When editing draw.io diagrams

- If you modify any `docs/leap-system*.drawio` file, create a png version of it, and check how it looks.
- In addition run `python3 scripts/check_drawio_layout.py docs/leap-system.drawio` (and any other touched `.drawio` files) and ensure it passes.
- The layout check must detect any child box overflowing its parent container.
- The layout check must also flag parent containers that are significantly larger than their contents; shrink or reflow so containers fit their children with modest padding.
- The layout check must flag vertical overlaps between sibling boxes in the same column (e.g., stacked tooling cards). Resolve by adjusting y-positions/heights to add clear spacing.
- The layout check must detect any overlaps between boxes that are not in an ancestor/descendant relationship, using absolute (parent-accumulated) positions so negative x/y or wide boxes don't hide overlaps.
- Also check for overlaps between sibling groups in the same layer (e.g., `G_*` demand boxes) and for vertical collisions where a group's bottom intersects the next row/section below (e.g., Transport demand touching Transformation). Resolve unintended overlaps by reflowing/reparenting; if an overlap is intentional (e.g., container frames), call it out explicitly.
- If the overflow check script is missing, implement a lightweight check in the change (or describe the overflow risk explicitly if you cannot run checks).
- double check for duplicates of arrows and boxes that may have been accidentally created during editing.

### Layout and spacing

- Avoid overlaps between boxes unless intentional container frames; reflow to remove unintended collisions.
- Use consistent spacing within rows/columns; align edges and keep gutters uniform.
- Keep a minimum margin from the canvas edges so elements do not feel cramped.
- Snap to a grid and keep box edges aligned across columns.
- Keep box sizes consistent for similar elements (e.g., step boxes, sector cards).
- Avoid extreme resizing; adjust layout before increasing box size.
- Keep section headers the same width as their section container where possible.

### Typography and content

- Keep text padding consistent inside boxes; avoid text touching borders.
- Keep labels centered/aligned consistently for a given row or section.
- Use consistent font sizes per hierarchy (headers, body, notes); avoid mixing too many sizes in one section.
- Keep connector label font size consistent; avoid multi-line labels unless necessary.
- If icons are used, keep style/size consistent and align them with text baselines.

### Colors

- Use a consistent palette per section (same fill/stroke/text colors for all boxes in that section).
- Avoid low-contrast text; ensure readable contrast against fills.
- Reserve accent colors for key headers or important flows only.
- Keep neutral colors for background containers; avoid overly saturated fills.
- Use dashed borders with lighter fills for group/region containers.
- Do not introduce new colors unless they signal a new semantic group.

### Strokes

- Keep stroke widths consistent within a section; avoid mixing thick and thin outlines.

### Arrows

- Shortest path first: Keep connectors as short as possible; prefer straight vertical/horizontal lines over multi-turn routes.
- Avoid crossings: Do not let lines pass through or over boxes/text; if a crossing is unavoidable, reflow the boxes instead of adding more bends.
- Route outside groups: When a line must cross a region, route it along the outside edge of group containers (left/right/top/bottom gutters).
- One gutter per side: Use consistent lanes (e.g., a single left margin lane) rather than multiple parallel long lines that drift.
- Parallel spacing: If two lines must run in parallel, keep consistent spacing and avoid lines running too close together.
- Downward flow: For top-to-bottom relationships, route vertically with minimal horizontal offsets; arrowheads should land on the top edge of the target box.
- Entry/exit from clean sides: Prefer left/right entry for lateral links, top/bottom for vertical links; avoid corners and offset endpoints slightly from corners.
- Corner exception: If a lateral line has to move vertically near a box connection, connect to the top of the box instead; if a vertical line has to move horizontally near a box connection, connect to the left/right side instead.
- Avoid last-moment jogs: Do not add unnecessary bends immediately before connecting to a box; straighten the final segment when possible.
- Keep labels close: Label text should sit near the destination box or mid-segment, not far away on long routes.
- No line underlays: Avoid routing behind text-heavy boxes where lines are hard to see.
- Visually verify arrowheads sit on box edges (not floating inside).

### Double-check

- since there are many rules here, after making changes, always run through the checklist again to ensure compliance.
- allow a longer timeout time for draw.io changes since they are complex (up to 10minutes is ok).

## Formatting and hygiene

- Prefer minimal diffs in `.drawio` files.
- Keep text ASCII unless the file already uses non-ASCII characters.
