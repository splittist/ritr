# Initial Implementation Plan

Implementation update: the first working vertical slice is recorded in
[Milestone 1](milestone-1.md). The phases below remain the broader roadmap;
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

1. Choose the package-manager and repository layout.
2. Create a minimal set of synthetic DOCX fixtures.
3. Define the package comparison report format.
4. Implement the no-edit round-trip test before building UI components.
