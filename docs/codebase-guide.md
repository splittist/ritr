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
DocxPackage.withXml()         construct a new immutable package snapshot
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
| `src/package/docx.ts` | Immutable package snapshots, relationships, content types, diagnostics, byte comparison |
| `src/engine/document.ts` | Read-only semantic interpretation: stories, source-bound spans, code tokens, editing policy |
| `src/engine/styles.ts` | Shared style index, inheritance chains, document defaults |
| `src/engine/numbering.ts` | List definitions, style links, label formats and story counter streams |
| `src/engine/paragraph.ts` | Paragraph list membership and indentation with property provenance |
| `src/ui/paragraph-layout.ts` | Source indentation to bounded, approximate display geometry |
| `src/engine/workspace.ts` | Exclusive mutation API: search, preview, commit, history, events, saved state |
| `src/io/save.ts` | Filesystem boundary: stage, validate, sync, publish a new file or directory |
| `src/cli.ts` | Thin command-line client of the same engine |
| `src/desktop/protocol.ts` | Typed contract between desktop processes |
| `src/desktop/main.ts` | Owns workspace and dialogs; validates IPC callers and arguments |
| `src/desktop/preload.ts` | Exposes named operations; no generic IPC or filesystem access |
| `src/ui/App.tsx` | Workspace navigation, inspector, search, previews and reports |
| `src/ui/RevealEditor.tsx` | CodeMirror projection with atomic code widgets and source selection |

The useful core nouns are `DocxPackage`, `Story`, `TextSpan`, `TextEdit`,
`ChangePreview`, and `Workspace`. There is no service container, plugin system,
generic command bus, or custom reactive framework to learn. `Paragraph` and
`ListLabel` add resolved display metadata without widening the mutation API.

For numbering, start with `tests/numbering.test.ts`. `Numbering` reads definitions
once; each story creates `Paragraphs` with a fresh `NumberingSequence`. Property
resolution happens before label generation, and indentation merges document
defaults, list properties, paragraph styles, then direct paragraph properties.
The renderer receives generated labels and measurements; it does not run counters.
The Word comparison findings and exact supported scope are recorded in
[Milestone 2](milestone-2.md).

## Preservation and identity

The package owns its original byte buffers. A no-edit save returns a copy of
the original archive, so it preserves both file and part identity. A changed
package shares untouched part buffers internally; public byte access returns
copies. This explicitly handles Node's `Buffer.slice()` aliasing behavior.

The XML parser records namespace URI, local name, attributes, and lexical
start/end offsets. Offsets count UTF-16 code units, matching JavaScript strings
and CodeMirror. Text is decoded semantically, but the entire original XML string
is retained for patching. Saxes validates XML; it is not used as a serializer.

Each parsed element receives a part-scoped identity such as
`word/document.xml#12`. These identities remain stable over the current set of
text-only commands because element topology never changes. Offsets are recomputed
after edits. **Do not extend this assumption to insertion/removal of elements.**
Before paragraph split/join or formatting commands, add an explicit identity
remapping layer or retain mutable bound XML nodes. Do not use saved offsets as
long-lived identities.

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

The next substantial slice should be direct typing within a span followed by
paragraph split/join. Structural identity management comes before widening the
editing surface. A full block/run graph will become useful then; the current
semantic projection is deliberately smaller than the long-term model.

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
