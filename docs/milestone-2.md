# Milestone 2: list labels and paragraph shape

Baseline: `07621f8` (the verified preservation-first workbench).
Status: implementation in progress.

## Goal

Display generated numbering, bullets and approximate paragraph indentation in
both Reveal Codes and the clean review view. Keep labels separate from document
text and make their definitions and inherited properties inspectable.

## Scope and approach

- Add explicit paragraph metadata while keeping existing source-bound text commands.
- Separate paragraph property inheritance, numbering definitions/counters, and
  UI layout into small modules.
- Resolve direct and style-based list membership, multilevel labels, common
  numeric formats, bullets, continuation, start overrides and restart rules.
- Display left/right, first-line and hanging indentation on a consistent scale.
- Preserve every source byte on inspection; text edits still patch only `w:t`.
- Diagnose unsupported or ambiguous numbering rather than inventing exact labels.

## Acceptance checks

1. Synthetic fixtures exercise legal outlines, heading numbering, bullets,
   interruptions, independent lists, restarts, overrides and wrapped text.
2. Expected labels and resolved indentation are covered by engine tests.
3. Desktop tests cover labels with codes on/off, property inspection and wrapping.
4. Compare representative list labels with Microsoft Word where available.
5. Edit text, undo/redo, save/reopen, and assert numbering/style parts are unchanged.

Pagination, exact Word line breaking, list editing, full table/conditional-style
layout, picture bullets and all international numbering systems are outside this
slice. The current review view includes revisions; ambiguous revision/opaque
numbering must be identified rather than represented as Word's final view.

Verification results and any refinements to this scope will be recorded here.
