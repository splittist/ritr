# Milestone 3: readable tables

Baseline: `3ff8901` (numbering and indentation).
Status: implemented and verified.

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

## Delivered

Stories now expose a block hierarchy referencing the original token stream.
`blocks.ts` owns hierarchy, `table.ts` owns read-only table interpretation,
`StoryView.tsx` owns composition, and `table-layout.ts` owns CSS approximation.
The mutation API and preservation layer are unchanged.

The desktop shows grid spans, vertical merges, nested tables, preferred/grid
widths, margins, simple borders, shading and vertical alignment. Table, row and
cell buttons open source XML and resolved properties. Numbering and indentation
work inside cells; text edits use the existing inspector preview/commit flow.

Malformed merges, content-bearing merge continuations, legacy `hMerge`, unfamiliar
row wrappers, right-to-left geometry, and display-limit violations use a diagnosed
linear view. This keeps content visible instead of inventing a plausible grid.
Limits are six nested display levels, 500 rows per table and 128 columns.

## Validation

- `npm run check`: 52 tests pass, formatting and TypeScript checks pass, production
  build succeeds. Tests cover hierarchy/ranges, grid holes, merges, safe fallback,
  tracked-cell protection, nesting and exact XML preservation during cell edits.
- `npm run test:desktop`: existing editing/list tests plus merged/nested table
  rendering, cell editing, undo/redo, property inspection and verified Save As.
  Clean and Reveal Codes screenshots inspected.
- Original and desktop-edited fixtures opened read-only in an isolated Word
  instance; both contained 28 paragraphs. `scripts/verify-tables.ps1` passed for
  both files, checking representative table cell coordinates and nested content.

## Deliberate approximations

Widths are display hints, not Word's autofit algorithm. Borders use a simplified
CSS model; missing borders get dotted editor guides, explicit nil borders remain
hidden. Conditional table styles, row-height rules, pagination, repeated headers,
floating placement and rotated text are not reproduced. Source properties remain
available for inspection and unchanged by text edits. Large tables also need a
future virtualization pass: each visible cell text region currently owns an editor.

Next: bounded inline editing within an existing text span, retaining previewable
commands and protected boundaries. Paragraph split/join needs separate identity
and structural-editing work.
