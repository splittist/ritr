# Milestone 3: readable tables

Baseline: `3ff8901` (numbering and indentation).
Status: implementation in progress.

## Goal and scope

Display tables as recognizable rows and columns while preserving the existing
text-span editing contract. Stories gain an explicit block hierarchy: paragraphs
and tables, with cells containing blocks. The flat token projection remains
available for search, source lookup and CLI diagnostics.

- Approximate grid/preferred widths, cell margins and simple borders.
- Horizontal spans and vertical merges, with diagnostics for invalid geometry.
- Numbering, indentation, codes and inspector editing inside cells.
- Selectable table/row/cell properties with source bindings.
- Nested tables where their geometry can be represented safely.
- Preserve source XML and package parts during inspection and cell text edits.

Table creation, row/column insertion, merge editing, exact Word autofit/layout,
floating placement, full conditional table styles and justification are deferred.
Inline typing remains the next editing slice after table display.

## Acceptance

Open fixtures containing merged cells and numbered cell paragraphs. Recognize
their layout in both views; inspect cell properties; edit cell text through the
existing preview/commit flow; undo/redo and save/reopen. Verify table definitions
and unrelated package parts are unchanged. Add engine and desktop tests, inspect
screenshots, and compare representative table structure with Word where possible.
