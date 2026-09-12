# Initial Implementation Plan

Implementation update: the first working vertical slice is recorded in
[Milestone 1](milestone-1.md); numbering and indentation display are recorded in
[Milestone 2](milestone-2.md). The phases below remain the broader roadmap;
several have partial implementations rather than completed exit conditions.

This plan favors vertical, testable slices. Each phase should leave behind a
usable diagnostic capability rather than only internal infrastructure.

## Phase 0: Repository and test foundation

- Establish the TypeScript monorepo or workspace structure.
- Add unit, integration, and fixture-test runners.
- Define fixture licensing and provenance metadata.
- Create structural ZIP/XML comparison utilities.
- Record test results per package part.

Exit condition: a test can open a DOCX, write a copy without changes, compare
the two packages, and explain every difference.

## Phase 1: Preservation-oriented package layer

- Read ZIP entries without discarding unknown parts.
- Parse content types and relationships.
- Parse selected WordprocessingML parts while retaining source XML.
- Assign stable internal identities and source bindings.
- Write unchanged parts without semantic regeneration.
- Validate relationship and content-type consistency.

Exit condition: representative documents survive a no-edit round trip, with
unchanged parts preserved and all differences reported.

## Phase 2: Read-only semantic engine

- Model documents, stories, blocks, inline content, and structural positions.
- Resolve styles and direct versus inherited formatting.
- Model tables, lists, hyperlinks, comments, notes, and revisions.
- Represent unsupported structures as opaque content.
- Produce diagnostics instead of silently discarding data.

Exit condition: command-line inspection can print the semantic structure and
source binding of every visible document element in the initial corpus.

## Phase 3: Read-only Reveal Codes application

- Create the Electron and React shell.
- Add workspace and document navigation.
- Implement a CodeMirror projection with text and atomic code decorations.
- Map UI selections to structural engine positions.
- Add code-category filtering and property inspection.
- Add a searchable command palette backed by a command registry.

Exit condition: several documents can be opened simultaneously and navigated
in a responsive, fixed-width Reveal Codes view.

## Phase 4: Safe text editing

- Define commands, change sets, transactions, and events.
- Implement text insertion, deletion, paragraph split, and paragraph join.
- Specify editing behavior at every formatting boundary.
- Add undo and redo.
- Translate supported changes into targeted OOXML mutations.
- Save only to a new file and validate it before reporting success.

Exit condition: edited documents reopen correctly in Microsoft Word or
LibreOffice, and unrelated package content survives.

## Phase 5: Formatting and structure editing

- Apply and remove direct run properties.
- Set paragraph properties and named styles.
- Add and remove hyperlinks.
- Edit lists, tables, page breaks, and section breaks conservatively.
- Display effective, inherited, and direct formatting distinctly.

Exit condition: formatting operations are fully command-driven, undoable, and
covered by boundary-focused round-trip fixtures.

## Phase 6: Workspace search and replacement

- Index all supported stories across open documents.
- Support literal, regular-expression, and formatting-aware queries.
- Generate a reviewable multi-document change set.
- Write, validate, and atomically commit all affected documents.
- Implement transaction-wide undo and failure recovery.

Exit condition: a replacement across several documents can be previewed,
applied safely, and undone as one operation.

## Phase 7: Comments and revisions

- Display comment anchors and comment threads.
- Add, edit, resolve, and remove comments where representable.
- Display tracked insertions, deletions, and property changes.
- Accept and reject revisions at selection, story, document, and workspace
  scopes.
- Preserve revision metadata and unsupported revision forms.

Exit condition: common reviewed documents can be inspected and processed
without losing unrelated comments or revisions.

## Phase 8: Preview and in-place save

- Integrate SuperDoc behind a preview adapter.
- Materialize preview documents from current engine state.
- Add atomic in-place replacement, backup policy, and recovery journals.
- Compare preview output with Word or LibreOffice rendering fixtures where
  useful.

Exit condition: users can confidently edit, preview, and replace original files
with a documented recovery path.

## First milestone acceptance scenario

The first combined milestone spans Phases 0 through 4:

1. Open several ordinary DOCX files.
2. Display body text, paragraph boundaries, and run formatting codes.
3. Insert and delete text at ordinary and formatted locations.
4. Save each result as a new DOCX.
5. Reopen the results in Word or LibreOffice.
6. Produce a report showing that unrelated package content survived.

## Immediate next steps

The foundation, numbering, and table display slices are implemented. Literal
workspace search and replacement now cross formatting runs, using source-bound
span edits and stopping at structural and review boundaries.

1. Add conservative paragraph split/join commands with preservation fixtures and
   explicit caret mapping, using the implemented identity-aware package patches.
2. Expand formatting inspection and add new actions to the shared command registry.

### Implemented slice: bounded inline typing

Users can edit text where they read it through a native inline draft field.
The Reveal Codes buffer remains read-only; applying a draft uses the engine's
existing source-bound text commands.

- Map selections to offsets within one editable text span, including empty spans.
- Keep inline changes as a draft with explicit apply/cancel, using the existing
  preview/commit transaction path and workspace undo/redo.
- Preserve caret and selection through projection refreshes. Handle composition,
  paste, and grapheme-aware deletion without exposing code placeholders as text.
- Refuse edits across spans, protected content, or structural boundaries; explain
  the boundary in the UI. Paragraph split/join is a separate structural command slice.

Acceptance: type, paste, delete, cancel, apply, undo and redo in ordinary text
and table cells, with codes shown and hidden. Save and reopen the result and
verify that only the intended text elements changed. Cover empty spans, Unicode
composition, and stale drafts after another workspace transaction.

Validation: 60 unit/integration tests and the hidden Electron smoke test pass.
The desktop test exercises typing, paste, deletion, draft cancel/apply, code
visibility changes, empty spans, table cells, protected/cross-span refusal,
undo/redo, Save As part preservation, and stale revisions. Composition coverage
uses synthetic browser events; operating-system IME behavior still needs manual
verification with the user's input methods.

### Implemented slice: source identity remapping

Package snapshots now carry private XML identity tables. Targeted patches retain
untouched nodes, support explicit retention of rewritten or moved nodes, and give
new nodes fresh IDs across staging branches. Text transactions use this path;
undo/redo restores identity tables together with package content. Raw whole-part
replacement retires identities in the changed part.

The package tests exercise insertion/deletion, identical neighbors, wrapper
split/join, subtree movement, follow-up text edits, table/header bindings,
namespace changes, multipart failure, invalid retention, and save/reopen without
adding identity metadata. The structural cases test the package foundation;
paragraph split/join is not yet exposed as an engine or desktop command.

Validation: all 70 unit/integration tests, formatting, type checking, production
build, and the hidden Electron smoke test pass.

### Implemented slice: GUI commands and palette

Existing GUI actions now have typed IDs, shared labels, availability reasons,
and selected keyboard shortcuts. Buttons and the searchable Commands palette
use the same guarded dispatcher and existing named desktop operations. The
palette supports keyboard navigation, retains unavailable actions with reasons,
and restores focus to a draft on close. Workspace undo shortcuts defer to native
text undo inside fields. Selection tracking now handles protected span endpoints
and keyboard selections so command availability follows the actual source range.

The next structural commands should enter through this registry and retain the
engine's preview/commit and source-preservation contracts.

Validation: all 74 unit/integration tests, formatting, type checking, production
build, and desktop smoke checks pass. The desktop coverage includes palette
filtering, unavailable/protected actions, keyboard navigation, focus and field
selection restoration, native text undo, workspace history shortcuts, and the
separate preview/apply shortcuts.

### Implemented slice: remappable keybindings and cross-run inline editing

Commands and key chords are separate. A validated, persisted override map drives
shortcut matching and displayed hints; editor deletion and draft history are named
actions with focus-specific bindings.

Inline drafts now cover contiguous editable segments, retain source-run pieces
through every input, and stage one transaction. Replacement uses the first selected
character's format; deletion keeps that typing format. Caret ownership survives
apply, and explicit empty-run editing retains the chosen properties. Structural
and protected boundaries remain enforced. Paragraph split/join remains next.
