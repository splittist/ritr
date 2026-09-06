# Milestone 1: preservation-first editing workbench

Implemented on 2026-09-06. The original architecture remains the long-term
direction; this milestone is a coherent, intentionally bounded vertical slice.

## Architecture review and chosen scope

The source package must remain authoritative for preservation. Building a
general rich-text editor first would leave that contract untested. This
implementation therefore starts with lexical source bindings, targeted text
patches, package comparisons, and transactional snapshots, then exposes them
through both a CLI and desktop UI.

The planned TypeScript/Electron/React/CodeMirror direction is followed. The
initial monorepo is simplified to one package with module boundaries. The
semantic model is currently stories and source-bound tokens, not the full
planned block/run graph. The command interface is explicit methods and typed
change sets, without a registry until there are more command families.

## Delivered versus the original phases

| Phase | Result |
| --- | --- |
| 0 | TypeScript setup, tests, synthetic fixture provenance, per-part byte reports |
| 1 | ZIP integrity/resource checks, XML source retention, relationships/content types, exact no-edit saves |
| 2 | Story and token projection, basic inherited/direct properties, table/hyperlink/review boundaries, opaque content |
| 3 | Desktop workspace, story navigation, atomic Reveal Codes widgets, source inspector, clean review view |
| 4 | Span-local insert/delete/replace, previews, undo/redo, targeted XML patches and verified Save As |
| 6, partial | Literal workspace search/replacement, protected matches, atomic in-memory transaction and new-directory batch export |

These are **partial slices of phases 2–4 and 6**, not claims that every original
phase exit condition is complete. Paragraph splitting/joining, direct typing,
the command palette, full style semantics, and existing-file transactions remain.

## Verification

- All 27 unit and integration tests pass. They cover file/part identity, exact lexical changes,
  XML/ZIP failures, Unicode, source identities, fields, revisions, comments,
  tables, hyperlinks, notes, headers, protected documents, stale previews,
  multi-document undo, failed staging, and exclusive Save As.
- All three checked-in `docxfix` documents survive no-edit saves with exact
  archive identity. Editing near annotations changes only the intended part.
- The production TypeScript, Electron and UI builds pass.
- The hidden Electron smoke test exercises actual controls for three open
  documents, span editing, transaction preview, multi-story replacement,
  undo/redo, Save As, source inspection, and code visibility. Its screenshot
  was visually inspected.
- The dependency audit reports zero known vulnerabilities at implementation time.
- Microsoft Word successfully opened the edited plain, review, and section
  fixtures read-only and found the expected replacement text in all three.

`scripts/verify-word.ps1` provides an additional optional Windows/Word check.
It opens supplied fixtures read-only in a hidden Word instance and can assert
expected replacement text. It never saves them. This checks basic opening and
text interoperability; it does not certify visual equivalence, every OOXML
feature, or modern comment semantics.

## Recommended next milestone

Implement direct typing and paragraph split/join with explicit formatting
inheritance and persistent source identities across structural edits. Add
boundary fixtures before widening mutation permissions. After that, add direct
formatting commands and cross-run search. Existing-file workspace replacement
should wait for backup and recovery-journal design, as the original architecture
requires.
