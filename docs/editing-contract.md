# Current editing and preservation contract

This describes implemented behavior, not the entire architecture roadmap.

## Text commands

- A location is a document ID, source-bound text span ID, and UTF-16 offset.
- Insert, delete, and replace operate inside one `w:t`. The desktop inspector
  replaces the selected span's full text; the engine also accepts subranges.
- Text inherits the existing run's properties because the run is not replaced.
- Empty results keep the existing text element. Existing empty text elements
  are editable; an empty paragraph without a text element cannot yet receive text.
- Surrogate pairs cannot be split. Grapheme clusters are not treated as atomic
  by the engine; use full-grapheme ranges when implementing direct typing.
- Tabs, line breaks, invalid XML characters, overlapping ranges, and stale
  previews are refused. Adjacent replacement ranges are allowed.
- Search is literal and crosses adjacent text spans and formatting codes.
  Every structure, review, or opaque code is a hard boundary, including paragraph
  and cell boundaries, tabs/breaks, hyperlinks, bookmark/comment anchors, fields,
  and revision boundaries. Search works within a hyperlink or revision boundary.
  Optional case-insensitive matching uses JavaScript Unicode regex folding
  without changing source offsets. Matches are non-overlapping within each segment.
- Each search match exposes ordered source `ranges`; its top-level span and
  offsets identify the first range for navigation. Cross-run replacement writes
  the replacement into that first range and deletes the matched text in later
  ranges. Unmatched prefixes/suffixes, run properties, and empty elements remain.
  Replacement text inherits the first matched run's formatting. An identical
  replacement leaves the original text distribution and formatting untouched.
- Replacement previews include a protected-match count. Protected matches are
  skipped in full if any matched slice is protected; direct attempts to edit
  protected spans throw an error.
- Workspace commit and undo/redo affect every staged document together.

## Inline drafts

Typing or pasting at a selection within one editable span, double-clicking a span,
or pressing Enter opens a native text field in that span's position. The field
contains source text only, including an empty string for an empty text element.
Empty spans have a dotted underline so they can be selected with codes hidden.

The draft is local to the UI until **Preview inline edit** and **Apply transaction**.
Cancel or Escape discards it. Editing a draft dismisses its previous preview.
Only one draft is active at a time; navigation, other editing commands, and Save As
wait for apply or cancel. Workspace undo/redo clears a draft with a status message.
The desktop checks the draft's source revision before staging it; stale drafts
are refused. Closing the window with a draft offers to keep editing.

The field supports native text composition and plain-text paste. Backspace/Delete
expand deletion to whole graphemes, including combining sequences and ZWJ emoji.
Tabs, line breaks, object placeholders, and invalid XML text are refused. Deletion
at the field edge stops there. Cross-span selections and protected text cannot
start an edit. Caret/selection offsets survive code visibility changes and are
restored into the projection after apply/cancel (clamped to the resulting span).
Typing inside a draft does not add workspace undo entries; applying it adds one.

## Source identity

Node IDs are scoped to an opened document and part. Targeted package patches
retain IDs for untouched opening tags and for rewritten/moved nodes explicitly
identified by the command. New nodes get fresh IDs; removed IDs cannot silently
become the identity of a later element. Retention requires a unique destination
element with the same namespace URI and local name. Invalid claims are refused
before returning a new package snapshot.

Identity tables are restored with package snapshots during undo/redo. No identity
metadata is written into DOCX files; reopening starts a new identity session.
Raw whole-part XML replacement retires that part's IDs instead of guessing node
continuity. These are package-layer capabilities, not permission to edit arbitrary
document structure. Paragraph split/join and caret mapping across those operations
remain unimplemented in the engine and desktop.

## GUI commands

The Commands button and Ctrl+Shift+P open a searchable modal palette. Existing
open/save, workspace history, inline/inspector preview, apply/dismiss, search,
replacement, code visibility, and package comparison actions share typed command
IDs and availability checks with their buttons and shortcuts. Unavailable actions
explain the missing document, editable selection, draft, query, or preview.
Checks run again at invocation; busy operations cannot be invoked twice.

Ctrl+Enter previews the current text draft. Ctrl+Shift+Enter applies the displayed
preview. Changing an inline or inspector draft dismisses its old preview. The
palette does not bypass preview/commit, protected-content policy, or stale-revision
checks. Ctrl+Z and redo shortcuts operate on workspace history in the projection;
inside text fields they remain native text undo/redo. Escape closes the palette
and returns focus without cancelling an underlying inline draft.

## Protected content

Fields (including results across paragraph boundaries), revisions, tracked run
properties, comment records, content controls, drawings, equations, and unknown
XML wrappers are preserved. Known revision text is visible but read-only.
Unknown wrappers appear as opaque codes, without flattening their descendants.
Digital signatures, enforced document protection, enabled Track Changes, and
package errors disable text editing. The tool does not create revisions.

Text adjacent to comment/bookmark anchors, inside ordinary hyperlinks and table
cells, or in ordinary header/footer/note stories can be edited without moving
those boundaries. Comment metadata and modern comment extension parts remain
untouched. This is not modern-comment editing support.

The plain-text view is a **review projection**, not Word's final/accepted text:
it includes visible revision text and placeholders for opaque content.

## Paragraph display

Generated numbers and bullets appear in both views. They are selectable display
objects, excluded from editable spans and search. The inspector exposes their
definition and property sources. Ordinary text commands do not renumber or
rewrite list definitions. Explicit paragraph metadata includes resolved list
membership and left/right/first-line indentation in source twips.

Common list formats, style-linked numbering, start/level overrides and Word's
restart/continuation behavior are supported. Shared abstract definitions share
counter streams; independent abstract definitions have separate streams.
Counters are scoped to individual stories. Unsupported/ambiguous numbering is
shown as `?` with diagnostics, including following paragraph-bearing opaque or
structurally revised content in a story.

Indentation is approximate in the UI. Display values are bounded, long labels
may widen hanging space, and inline Reveal Codes affect wrapping. Exact source
values remain visible in the inspector. This is not pagination or full layout;
custom tabs, character-unit indentation and right-to-left layout are not fully
resolved. See [Milestone 2](milestone-2.md) for formats and verification.

## Save and validation

No-edit saves preserve the original archive bytes. Edited saves recompress the
archive; ZIP timestamps, order, compression, comments, and extra fields are not
promised to survive. Untouched uncompressed part bytes do survive.

Each single-file Save As:

1. Validates the package and its relationship/content-type diagnostics.
2. Writes a uniquely named sibling temporary file, flushes it, and rereads it.
3. Reopens and compares the package parts.
4. Publishes it using an exclusive hard link and removes the temporary name.

An existing destination is refused, even after a native dialog's overwrite
confirmation. Filesystems without hard-link support cause a clear save failure;
there is no unsafe overwrite fallback. A crash can leave a `.ritr-*.tmp` file,
but cannot overwrite an original through this path.

CLI batch export stages every document in a sibling directory and publishes
that directory with one rename. The destination must be new. Failed staging
cleans up its own temporary directory. This is **Save As for a batch**, not
atomic replacement of a set of existing files. A crash can leave a complete or
partial `.ritr-export-*` staging directory. Power-loss durability of directory
metadata is not guaranteed. There is no recovery journal or disk-level undo.

ZIP input limits: 64 MiB archive, 128 MiB expanded total, 32 MiB per part,
10,000 entries. XML limits: depth 256, 250,000 elements per part. Only UTF-8 XML,
ordinary single-disk non-ZIP64 archives, and stored/deflated entries are supported.
DTD, encrypted ZIP, duplicate names, path traversal, malformed XML, and CRC
failures are refused. Relationship targets are checked; external targets are
preserved but never fetched. This is not full OPC/OOXML schema validation.

## Not implemented yet

Paragraph split/join; arbitrary cross-run selection edits; direct formatting changes; list,
table, hyperlink, and section restructuring; full style/layout resolution and
exotic numbering formats;
comment or revision mutation; granular code-category filters;
paginated preview; in-place saves; persisted workspaces or undo journals;
installer packaging. The desktop is an inspector/editor prototype, not yet a
complete replacement for Word.
