# ritr

`ritr` is a local-first, text-oriented word processor for editing Microsoft Word
documents while exposing their underlying formatting and structure.

Its primary editing experience is inspired by WordPerfect's Reveal Codes:
formatting boundaries and document structures are visible, inspectable editor
objects alongside fixed-width text. The application is also designed around
multi-document workspaces and safe batch operations.

This is a personal tool and experimental codebase, not a commercial product.

## Project principles

- Preserve DOCX content that the application does not understand.
- Make formatting and structure inspectable instead of implicit.
- Route every mutation through a reusable command interface.
- Treat multi-document changes as previewable, atomic transactions.
- Prefer safe refusal and useful diagnostics over silent document damage.
- Keep the editing engine independent from the desktop UI and preview renderer.

## Initial documents

- [Architecture and design decisions](docs/architecture-and-design.md)
- [Implementation plan](docs/implementation-plan.md)

## Try it

Requires Node.js 22.12+ (tested on Node 24) and npm. No server or account is needed.

```powershell
npm install
npx install-electron
npm start
```

`install-electron` downloads the desktop runtime; run it once after installation.
The explicit step also works with npm versions that block dependency install
scripts. To open the bundled sample documents immediately:

```powershell
npm start -- fixtures/generated/plain.docx fixtures/generated/review.docx fixtures/generated/sections.docx
```

Select text in the document, edit it in the inspector, and choose **Preview text
edit**, then **Apply transaction**. Clicking a code shows its source XML.
For inline editing, type at a text selection or double-click a span, then choose
**Preview inline edit** and **Apply transaction**. Escape or **Cancel inline edit**
discards the draft. Inline edits stay within one span, including table-cell text.
Workspace replacement uses the same preview/commit flow. Undo reverses the entire
transaction. Save As creates a new verified DOCX and refuses existing paths.

## Implemented

- Preservation-oriented ZIP/OOXML package layer and per-part byte comparison.
- Snapshot-owned source identities, retained through targeted XML patches and undo/redo.
- Body, header/footer, note and comment stories, source-bound text, formatting
  inspection, tables/hyperlinks, revision boundaries, and opaque content codes.
- Generated multilevel numbers and bullets, style-linked headings, restart and
  continuation handling, plus approximate paragraph and hanging indentation.
- Table grids, horizontal/vertical merges, nested tables, approximate widths and
  borders, with source-bound table/cell inspection and the same safe text editing.
- Insert/delete/replace within text spans, previewable workspace transactions,
  cross-run literal search/replacement, protected-match reporting, and session undo/redo.
- Electron/React desktop with a CodeMirror Reveal Codes projection and inspector.
- Diagnostic CLI, safe Save As, and staged export to a new batch directory.
- Synthetic `docxfix` fixtures, preservation/command/save tests, and desktop
  end-to-end verification.

This is a working text-span editor and document workbench. Paragraph split/join,
formatting changes, comment/revision editing,
paginated preview, and in-place saves are still future work.
See the precise [editing contract](docs/editing-contract.md).

## CLI examples

```powershell
npm run cli -- inspect fixtures/generated/review.docx
npm run cli -- inspect fixtures/generated/review.docx --json
npm run cli -- roundtrip fixtures/generated/plain.docx copy.docx
npm run cli -- compare fixtures/generated/plain.docx copy.docx
npm run cli -- search "Replace this phrase" fixtures/generated/plain.docx
npm run cli -- replace "Replace this phrase" "Updated wording" fixtures/generated/plain.docx
npm run cli -- replace "Replace this phrase" "Updated wording" fixtures/generated/plain.docx fixtures/generated/sections.docx --out-dir edited-copies
```

Replacement previews by default; `--out-dir` applies it and exports all supplied
documents into a new directory. Search crosses formatting runs but stops at
structure and review boundaries. Replacements inherit the first matched run's
formatting; the preview shows every affected span.
`compare` exits with code 2 when parts differ; invalid operations exit with 1.

## Understand and develop it

Start with the [codebase guide](docs/codebase-guide.md), then the first test in
`tests/engine.test.ts`. One TypeScript package keeps the initial codebase small;
the package adapter, engine, filesystem operations, desktop shell and UI have
separate folders under `src/`.

```powershell
npm run check          # Formatting, tests, type checking and production build
npm run test:desktop   # Hidden Electron smoke test; run after building
npm run format        # Consistent readable source formatting
npm run fixtures      # Regenerate fixtures using ../docxfix/.venv
```

`npm run dev` serves only the UI for styling; local file operations require the
Electron bridge, so use `npm start` for the working application. Desktop test
screenshots and output documents go to ignored `test-results/` directories.

The [first milestone report](docs/milestone-1.md) records the editing foundation;
[Milestone 2](docs/milestone-2.md) documents numbering, indentation and Word
comparisons. Try the new display features with:

```powershell
npm start -- fixtures/generated/lists.docx fixtures/generated/legal.docx fixtures/generated/headings.docx
```

[Milestone 3](docs/milestone-3.md) adds readable tables. Try it with
`npm start -- fixtures/generated/tables.docx`.

The [fixture guide](fixtures/README.md) records provenance and generation.
