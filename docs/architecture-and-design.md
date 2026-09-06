# Architecture and Design Decisions

Status: Accepted initial direction
Date: 2026-09-06

## 1. Purpose

`ritr` is a Windows-first, local-only desktop word processor focused on:

- fixed-width text editing;
- an interactive Reveal Codes representation of formatting and structure;
- reading, editing, writing, and safely round-tripping DOCX files;
- working with several documents as a single workspace;
- command-driven operation, including cross-document search and replacement.

The project is a personal application. It is not intended to become a hosted
service, collaboration platform, or polished product for third-party users.

## 2. Product boundaries

The application has no server component, user accounts, cloud storage,
telemetry, marketplace, or real-time collaboration.

The initial target is Windows. Cross-platform portability is welcome where it
does not compromise the design, but it is not a milestone requirement.

Macros in DOCM packages may eventually be preserved as opaque package content.
The application will not execute VBA, interpret macro behavior, or provide a
macro editor.

Pixel-identical Microsoft Word layout is explicitly out of scope. Paginated
rendering is a secondary preview and validation aid, not the primary editor.

## 3. Technology direction

The initial implementation will use:

- TypeScript for the engine and application;
- Electron for the desktop shell;
- React for application UI;
- CodeMirror 6 for the fixed-width Reveal Codes editor;
- SuperDoc, initially, as an optional document preview adapter.

The engine must remain independent of Electron, React, CodeMirror, browser DOM
selection, filesystem dialogs, and any particular preview implementation.
CodeMirror displays and edits a projection of engine state; its text buffer is
not the canonical document model.

## 4. Architectural boundaries

The system is divided into four principal layers:

```text
DOCX package adapter
        <->
Document engine
        <->
Commands and transactions
        <->
UI projections and preview adapters
```

This is not intended as textbook MVC. The critical rule is:

> The UI may query the engine, but every mutation passes through the command
> layer.

### 4.1 DOCX package adapter

The package adapter owns OOXML-specific concerns:

- ZIP package parsing and creation;
- content types, relationships, and XML parts;
- source bindings between engine nodes and OOXML nodes;
- preservation of unknown parts, elements, attributes, and extensions;
- translation of engine changes into targeted XML mutations;
- package validation and save diagnostics.

The adapter is a first-class subsystem rather than a lossy import/export
converter.

### 4.2 Document engine

The engine owns word-processing semantics:

- documents and workspaces;
- stories and structural positions;
- paragraphs, runs, styles, lists, tables, and hyperlinks;
- headers, footers, footnotes, and endnotes;
- comments and anchored ranges;
- tracked insertions, deletions, and property changes;
- selections, formatting inheritance, and validation;
- opaque placeholders for preserved unsupported content.

The main body, headers, footers, footnotes, and endnotes are modeled as variants
of a common `Story` or content-stream abstraction. Locations include document,
story, and structural position so commands and searches can work across all
supported stories.

Comments are external records with anchored document ranges and their own
content. Tracked changes are structural entities, not merely display flags.

### 4.3 Commands and transactions

Commands are the exclusive mutation API. They are usable by the editor,
command palette, keyboard shortcuts, tests, batch operations, and possible
future scripting.

The command system distinguishes:

- **Query:** reads state and cannot mutate it.
- **Command:** expresses user intent.
- **Change set:** describes calculated concrete changes, diagnostics, and a
  preview.
- **Transaction:** applies a change set atomically.
- **Event:** reports committed changes to projections and other consumers.

Representative commands include inserting and deleting text, applying
formatting, changing paragraph properties, adding comments, accepting or
rejecting revisions, and replacing matches across a workspace.

Undo and redo operate on committed transactions. A workspace-wide operation is
one transaction even when it affects several files.

### 4.4 UI projections

The UI consumes purpose-specific projections rather than directly exposing the
OOXML tree. Expected projections include:

- Reveal Codes;
- clean text;
- outline;
- comments and revisions;
- workspace search results;
- preview document.

The Reveal Codes projection is a sequence of text and atomic display tokens.
Each display token maps back to an engine location or property. Formatting
codes are not literal markup stored in the document.

## 5. Source of truth and preservation model

An opened DOCX package remains the source for preserving package content. The
engine maintains a semantic graph, bindings back to source XML, diagnostics,
and committed changes.

The normal pipeline is:

```text
Original DOCX package
        -> parsed XML plus semantic index
        -> commands and targeted changes
        -> patch affected XML
        -> validate and write a new package
```

The application will not normally regenerate an entire DOCX from a simplified
model. Unsupported content is represented by opaque engine nodes or annotations
whose edit policy is preservation.

Source bindings must use stable identities assigned during parsing rather than
depending only on numeric XML indexes or fragile XPath expressions.

## 6. Round-trip contract

The following behavior is required:

1. Opening and saving without semantic edits preserves every package part
   byte-for-byte where practical.
2. Editing supported content may rewrite the smallest safely addressable XML
   region.
3. Unaffected and unsupported package content is preserved.
4. When a requested edit cannot safely preserve surrounding content, the
   application warns or refuses the operation.
5. The application never silently promises exact visual equivalence with Word.
6. Saving produces diagnostics describing material normalization, preservation
   limitations, or unsupported affected content.

ZIP container metadata and ordering may prevent the complete output file from
being byte-identical even when every package part is unchanged. Tests therefore
distinguish whole-file identity, package-part identity, semantic identity, and
visual interoperability.

## 7. Initial content support

The first editable feature set covers:

- body text and paragraphs;
- text runs and direct character formatting;
- paragraph properties and named styles;
- lists and numbering;
- hyperlinks;
- tables;
- page and section breaks;
- headers and footers;
- footnotes and endnotes;
- comments;
- existing tracked changes, including accept and reject operations.

Creation of complex tracked property changes may follow later.

Initially preserved but not necessarily editable:

- images, drawings, and floating objects;
- fields;
- content controls;
- equations;
- embedded objects;
- unfamiliar extension markup;
- VBA projects in macro-enabled packages.

These are displayed as inspectable opaque objects rather than silently omitted.

## 8. Reveal Codes interaction contract

Formatting and structure codes are atomic editor objects. Initially:

- cursor keys move across codes predictably;
- clicking selects a code;
- Enter opens its properties;
- deletion invokes an appropriate semantic command;
- destructive structural deletion may require confirmation;
- code categories can be shown, hidden, expanded, or collapsed;
- clean-text and Reveal Codes views share one engine selection;
- inherited formatting is visually distinct from direct formatting;
- comment anchors and revisions are visible structural boundaries.

Typing, pasting, deleting, splitting paragraphs, and joining paragraphs at
formatting boundaries require explicit behavioral specifications and tests.

## 9. Workspaces and multi-document operations

A workspace is a set of open or indexed documents. Search results identify a
document, story, and structural location.

Cross-document mutations always create a previewable change set. Applying the
change writes all outputs to temporary locations, validates them, and only then
commits replacements. Failure before commit leaves originals untouched.

Workspace-wide undo reverses the complete transaction rather than individual
files independently.

## 10. Save safety

Early versions default to Save As. Before in-place saving is enabled, the system
must support:

- atomic replacement;
- configurable backups;
- package validation before replacement;
- recovery after application interruption;
- clear reporting of documents that changed;
- transaction journals sufficient for workspace undo.

## 11. Testing strategy

The project begins with a representative, legally redistributable test corpus.
Tests operate at several levels:

- package-part preservation;
- XML structural comparison with defined normalization rules;
- semantic engine comparison;
- command and transaction behavior;
- save/reopen interoperability;
- rendering or visual comparison for selected fixtures;
- corruption and failure recovery.

Every newly supported OOXML construct requires fixtures for unedited
preservation, local editing near the construct, and intentional editing of the
construct where supported.

## 12. Deferred questions

These decisions are intentionally deferred until prototypes provide evidence:

- the final on-disk workspace format;
- whether to persist a command journal between sessions;
- the exact syntax and visual language for codes;
- how much OOXML normalization is acceptable after an affected region changes;
- whether a public scripting interface is useful;
- support for non-Windows platforms.
