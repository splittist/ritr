# Codebase guide

The implementation is one TypeScript package with explicit source folders. A
monorepo would currently add more configuration than useful isolation. The
engine can still be extracted later: it has no Electron, React, filesystem,
or browser imports.

## Start with one edit

Read the first test in `tests/engine.test.ts`. It opens a document, creates a
preview, commits it, checks exactly which XML bytes changed, then undoes and
redoes the transaction. That is the application in miniature.

Then follow this path:

```text
UI inspector / CLI
       |
       v
Workspace.preview()          validate intent and prepare every document
       |
       v
textPatch() + patchXml()      replace only selected w:t elements
       |
       v
DocxPackage.withXmlPatches()  construct a snapshot with retained source identities
       |
       v
Workspace.commit(id)         publish all snapshots together
       |
       v
readDocument()               rebuild semantic stories and display tokens
```

A preview contains before/after text for review, but committing uses the staged
transaction held privately by the workspace. Modifying the returned preview
cannot change what gets committed. Any workspace mutation expires pending
previews. Undo/redo moves whole transactions between two stacks.

## Where things live

| File | Responsibility |
| --- | --- |
| `src/package/zip.ts` | ZIP bounds, entry checks, CRC verification, inflation and compression |
| `src/package/xml.ts` | Namespace-aware XML parsing, source ranges, lexical text patches |
| `src/package/xml-identity.ts` | Patch-based node identity retention, fresh IDs, and identity comparison |
| `src/package/docx.ts` | Immutable package snapshots, relationships, content types, diagnostics, byte comparison |
| `src/engine/document.ts` | Read-only semantic interpretation: stories, source-bound spans, code tokens, editing policy |
| `src/engine/styles.ts` | Shared style index, inheritance chains, document defaults |
| `src/engine/numbering.ts` | List definitions, style links, label formats and story counter streams |
| `src/engine/paragraph.ts` | Paragraph list membership and indentation with property provenance |
| `src/ui/paragraph-layout.ts` | Source indentation to bounded, approximate display geometry |
| `src/engine/workspace.ts` | Exclusive mutation API: search, preview, commit, history, events, saved state |
| `src/engine/text-range.ts` | Shared text boundaries, source positions, caret affinity, and formatted piece edits |
| `src/ui/keymap.ts` | Default chords, scoped matching, validated overrides, and shortcut presentation |
| `src/engine/search.ts` | Literal cross-run matching with structural boundaries and source ranges |
| `src/io/save.ts` | Filesystem boundary: stage, validate, sync, publish a new file or directory |
| `src/cli.ts` | Thin command-line client of the same engine |
| `src/desktop/protocol.ts` | Typed contract between desktop processes |
| `src/desktop/main.ts` | Owns workspace and dialogs; validates IPC callers and arguments |
| `src/desktop/preload.ts` | Exposes named operations; no generic IPC or filesystem access |
| `src/ui/App.tsx` | Workspace navigation, inspector, search, previews and reports |
| `src/ui/commands.ts` | Typed command IDs, GUI labels, availability, and guarded dispatch |
| `src/ui/CommandPalette.tsx` | Modal command search, keyboard selection, and unavailable-action explanations |
| `src/ui/RevealEditor.tsx` | CodeMirror projection with atomic code widgets and source selection |

The useful core nouns are `DocxPackage`, `Story`, `TextSpan`, `TextEdit`,
`ChangePreview`, and `Workspace`. There is no service container, plugin system,
generic command bus, or custom reactive framework to learn. `Paragraph` and
`ListLabel` add resolved display metadata without widening the mutation API.

For search, start with `tests/search.test.ts`. `searchStory()` joins text across
formatting codes and maps each match back to source span ranges. Other code
categories terminate a search segment. `previewReplace()` expands editable
matches into ordinary `TextEdit` commands, putting replacement text in the first
matched run and removing matched slices from later runs. The existing transaction
and preservation checks apply to the complete change.

For numbering, start with `tests/numbering.test.ts`. `Numbering` reads definitions
once; each story creates `Paragraphs` with a fresh `NumberingSequence`. Property
resolution happens before label generation, and indentation merges document
defaults, list properties, paragraph styles, then direct paragraph properties.
The renderer receives generated labels and measurements; it does not run counters.
The Word comparison findings and exact supported scope are recorded in
[Milestone 2](milestone-2.md).

For tables, start with `tests/table.test.ts`. `blocks.ts` builds a hierarchy over
half-open ranges in the existing story token stream; it does not duplicate or
replace the text model. `table.ts` resolves grid coordinates, merged cells and
direct display properties. Cells contain blocks, so nested tables use the same
representation. Unsupported geometry retains a diagnosed linear projection.

`StoryView.tsx` renders HTML tables and groups adjacent text blocks into embedded
`RevealEditor` instances. `table-layout.ts` converts source measurements into
approximate CSS. Selecting cell text still submits the existing `TextEdit`;
neither HTML nor CodeMirror becomes the document's source of truth.

## Preservation and identity

The package owns its original byte buffers. A no-edit save returns a copy of
the original archive, so it preserves both file and part identity. A changed
package shares untouched part buffers internally; public byte access returns
copies. This explicitly handles Node's `Buffer.slice()` aliasing behavior.

The XML parser records namespace URI, local name, attributes, and lexical
start/end offsets. Offsets count UTF-16 code units, matching JavaScript strings
and CodeMirror. Text is decoded semantically, but the entire original XML string
is retained for patching. Saxes validates XML; it is not used as a serializer.

Each opened element receives a part-scoped identity such as
`word/document.xml#12`. Package snapshots retain a private identity table for each
changed XML part and reapply it on every parse. Source offsets are still recomputed;
never use saved offsets as long-lived identities.

`withXmlPatches()` maps unchanged opening tags through the patch offsets. A node
keeps its ID only if its opening tag survives and its expanded name still matches.
Rewritten or moved nodes can be retained explicitly with a patch's `retain` entries:
each names an old node ID and the new opening-tag offset inside `replacement`.
The targets must be valid elements of the same expanded name, and retention is
one-to-one. `textPatch()` declares retention for the text element it rewrites.
Other new nodes receive fresh session IDs that are never reused across branches.

Start with `tests/identity.test.ts` for insertion, deletion, wrapper split/join,
explicit subtree moves, semantic bindings, and invalid claims. The standalone
`compareXmlIdentities()` report lists retained, created, and removed IDs. Raw
`withXml()` replacements have no continuity proof, so every node in a changed
part receives a fresh identity; use targeted patches for engine commands.

Identity tables live in immutable package snapshots and therefore travel with
workspace undo/redo. They never appear in saved OOXML. Reopening a file starts
a new identity session. These IDs are scoped to the opened document and part;
they are not persistent cross-file identifiers. Node continuity also does not
map a caret from a removed span into another span: future split/join commands
must define that position mapping and their own structural editing policy.

Replacing text rewrites just its `w:t` element, escaping text and setting
`xml:space="preserve"`. Other elements, attributes, namespace declarations,
comments, whitespace, and package parts remain byte-identical. XML embedded
inside a text element causes editing refusal.

## What the UI owns

React owns navigation, drafts, and the currently displayed preview. The
workspace in Electron's main process owns committed document state. The
CodeMirror buffer is read-only; it contains display placeholders for codes,
which are never written into the DOCX. Selecting a span opens the text inspector.
The inspector submits a `TextEdit`, exactly like any other engine client.
Inline editing uses a native textarea widget for a contiguous editable text segment.
`App.tsx` owns the source revision and draft history; `inline-edit.ts` maintains one
piece per source span and reconciles individual browser input events. The desktop
stages those pieces with `previewPieces`, preserving all untouched run properties.
`text-range.ts` supplies shared segmentation and source-position mapping.

`commands.ts` is a small GUI action registry. Buttons, palette entries, and global
shortcuts share command metadata and availability checks. `App.tsx` supplies one
typed handler per command ID; these handlers still call the named desktop API and
the existing engine transactions. No generic IPC execution endpoint is exposed.
Navigation and direct text input remain ordinary selection/draft interactions.

Dispatch rechecks the current context and serializes operations. A disabled
action returns its reason without invoking its handler. The modal palette keeps
unavailable actions visible and restores focus on close. Native text fields keep
their own undo/redo shortcuts, while the document projection uses workspace
history. Preview and apply have distinct shortcuts. The current source selection,
including keyboard selection and protected boundaries, drives edit availability.
Adding an action means defining its ID and metadata, implementing its typed
handler, and exposing any desired button; the palette picks it up automatically.

The current UI deliberately does not translate arbitrary CodeMirror changes
into commands. That requires paragraph and formatting-boundary semantics first.

## How to add behavior

1. Write a small fixture covering the construct and its neighboring content.
2. Add read-only interpretation and a useful code token in `document.ts`.
3. Specify the allowed editing boundary in `editing-contract.md`.
4. Add a command that stages source patches without changing current state.
5. Assert no-edit preservation, the exact edited region, protected boundaries,
   multi-document failure behavior, undo/redo, and save/reopen.
6. Expose the operation through the IPC contract and UI.

The next substantial slice should add conservative paragraph split/join commands
using the identity-aware patch path. The read-only block hierarchy is not yet a
structural editing graph. Commands must still specify allowed paragraph properties,
review and opaque-content boundaries, and how selections move across a split/join.

## Tradeoffs to revisit

Parsing is synchronous and repeated on queries. That keeps the prototype easy
to audit, but large documents need cached parsed trees and a worker boundary.
History retains snapshots for the session and has no memory budget or journal.
Style inspection reports a basic cascade, not Word's fully resolved appearance.
Package validation checks integrity and consistency, not the entire OOXML schema.

Library references used for these boundaries:
[fflate](https://github.com/101arrowz/fflate),
[saxes](https://github.com/lddubeau/saxes), and
[Electron security guidance](https://www.electronjs.org/docs/latest/tutorial/security).

Key chords and focus scopes are defined separately in `keymap.ts`. The sidebar's
JSON overrides are validated, persisted in local storage, and applied to both
matching and shortcut labels. Editor deletion and draft history also use this map.
