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

Type directly in the document. **Enter** splits a paragraph; **Backspace** at its
start or **Delete** at its end joins adjacent paragraphs. Text editing crosses
formatting runs, and paste may introduce paragraph breaks. **Undo** reverses each
input operation. Structural editing stays within ordinary paragraphs in one
container, including a table cell; protected content and section boundaries stop it.

Use the formatting toolbar for bold, italic, underline, text color, and OOXML
highlighting. With a selection it formats that text; at a caret it controls subsequent
typing. Ctrl+B, Ctrl+I, and Ctrl+U are remappable editor commands. Continuous typing
or deletion is grouped for undo, with a new step after a pause or caret movement.
Enter at the end of a non-list paragraph uses its declared next style when available.
In an empty list item, Enter moves up one level, or exits a top-level list.
Alt+Shift+Right and Alt+Shift+Left increase or decrease list level, including across
selected list paragraphs. Both commands are remappable and available in the palette.
Tab retains its existing focus/navigation behaviour.

The inspector and workspace replacement still use **Preview text edit** and
**Apply transaction**. Clicking a code shows its source XML. Save As creates a new
verified DOCX and refuses existing paths.

Click **Commands** or press **Ctrl+Shift+P** to search the available actions.
Use arrow keys and Enter to choose an action, or Escape to return to your edit.
Unavailable commands remain visible with an explanation.

| Shortcut | Action |
| --- | --- |
| Ctrl+O | Open documents |
| Ctrl+Shift+S | Save As |
| Ctrl+F | Focus workspace search |
| Ctrl+Enter | Preview the inspector text edit |
| Ctrl+Shift+Enter | Apply the displayed transaction |
| Ctrl+Z | Undo a workspace transaction |
| Ctrl+Y / Ctrl+Shift+Z | Redo a workspace transaction |
| Ctrl+Shift+E | Toggle Reveal Codes |
| Ctrl+B / Ctrl+I / Ctrl+U | Toggle bold / italic / underline |
| Enter | Split paragraph; use next style at paragraph end |
| Backspace / Delete | Delete text or join at a paragraph edge |

Inside text fields, undo/redo shortcuts retain native text-editing behavior.
The document editor uses workspace undo/redo, preserving formatting and source
identities. Use the toolbar or palette for workspace undo while an inspector field
has focus.

Expand **Keybindings** in the sidebar to remap commands with JSON, for example
`{"commands.open": ["Ctrl+k"]}`. Use `[]` to unbind a command or `{}` to restore
defaults. Changes apply immediately and persist locally.

## Implemented

- Preservation-oriented ZIP/OOXML package layer and per-part byte comparison.
- Snapshot-owned source identities, retained through targeted XML patches and undo/redo.
- Body, header/footer, note and comment stories, source-bound text, formatting
  inspection, tables/hyperlinks, revision boundaries, and opaque content codes.
- Fixed-width text with editable bold, italic, underline, RGB color, and OOXML highlighting.
- Generated multilevel numbers and bullets, style-linked headings, restart and
  continuation handling, plus approximate paragraph and hanging indentation.
- Table grids, horizontal/vertical merges, nested tables, approximate widths and
  borders, with source-bound table/cell inspection and the same safe text editing.
- Native inline editing, bounded paragraph split/join and multiline paste, with preserved formatting,
  previewable workspace transactions,
  cross-run literal search/replacement, protected-match reporting, and session undo/redo.
- Electron/React desktop with a CodeMirror Reveal Codes projection and inspector.
- Shared GUI command registry, searchable palette, contextual availability, and remappable shortcuts.
- Diagnostic CLI, safe Save As, and staged export to a new batch directory.
- Synthetic `docxfix` fixtures, preservation/command/save tests, and desktop
  end-to-end verification.

This is a working document editor and workbench. Comment/revision editing,
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
