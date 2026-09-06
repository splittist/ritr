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
- Search is literal and limited to individual spans. It does not find phrases
  crossing run boundaries. Optional case-insensitive matching uses JavaScript
  Unicode regex folding without changing source offsets.
- Replacement previews include a protected-match count. Protected matches are
  skipped explicitly; direct attempts to edit them throw an error.
- Workspace commit and undo/redo affect every staged document together.

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

Paragraph split/join; cross-run text edits; direct formatting changes; list,
table, hyperlink, and section restructuring; full style/numbering resolution;
comment or revision mutation; command palette; granular code-category filters;
paginated preview; in-place saves; persisted workspaces or undo journals;
installer packaging. The desktop is an inspector/editor prototype, not yet a
complete replacement for Word.
