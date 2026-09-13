# Current editing and preservation contract

This describes implemented behavior, not the entire architecture roadmap.

## Text commands

- A location is a document ID, source-bound text span ID, and UTF-16 offset.
- Low-level text edits operate inside one `w:t`. The inspector replaces a selected
  span; range commands and inline drafts compose edits across adjacent formatting runs.
- Surviving text retains its source run properties. Insertions inherit the caret's
  source run; at an ambiguous boundary the preceding nonempty run wins, falling
  back to the following run at segment start. An explicit source position (including
  a selected empty run) wins over that default.
- Cross-run replacement inherits the first selected character's run, independently
  of selection direction. Deletion retains that run as the subsequent typing format.
  Empty source elements remain; arbitrary empty runs do not override caret ownership.
- `previewRange` accepts source endpoints and an expected revision, validates the
  complete range, and returns a source-bound caret. `previewPieces` validates a
  draft's ordered source spans and revision before staging its changed pieces.
- Empty results keep the existing text element. Existing empty text elements
  are editable; ordinary empty paragraphs can receive text through the direct editor.
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

## Direct inline editing and paragraph commands

The CodeMirror document is editable in place. Text input no longer opens a separate
textarea and does not require Preview/Apply. Each native input operation is staged,
validated, and committed as one workspace transaction. Inspector and workspace
replacement commands retain their explicit preview/commit workflow.

The projection contains source text and separators between paragraphs, with no
phantom line after the final paragraph. Formatting, list labels, bookmark/comment
range and proofing markers, and the final section-properties code are widgets
without text offsets; other document objects occupy protected
slots. Showing/hiding codes never changes editable offsets. A serialized input queue
preserves fast typing while engine snapshots catch up. Failed input restores the
last confirmed snapshot and discards queued dependent input. Navigation and Save As
wait for pending input. Protected text and object edits are also refused locally.

Enter (`paragraph.split`) splits at the caret; Backspace at paragraph start or
Delete at paragraph end joins neighboring paragraphs. Selection replacement and
plain-text paste may span ordinary paragraphs or introduce multiple paragraphs.
One such replacement/paste is one atomic undo entry. Unicode deletion consumes whole
graphemes across run boundaries. Tabs, embedded objects, and invalid XML text remain
unsupported input. Composition is committed when composition ends; automated coverage
uses synthetic composition events and native text input, not an operating-system IME.

Structural edits work on paragraphs containing direct runs, text, and supported
markers and run objects within one body, header/footer, note, or table cell.
Bookmarks, comment range and proofing markers, tabs, breaks, drawings, and note
references retain their XML, identities, and order on their source side of a split.
Hidden range markers add no phantom caret positions. A boundary split uses the
selected text span to determine its side of a marker or object. Section properties,
revisions, fields, unsupported wrappers, and protected text still prevent restructuring;
refusals name the unsupported element. Text replacements across anchors remain refused.
Joins require adjacent sibling paragraphs and never cross cells. Paragraph-local
namespace declarations on the removed paragraph are conservatively refused.

Splitting retains the original paragraph identity on the left. At run boundaries,
whole runs move to the appropriate paragraph with their XML and identities intact.
Only a split inside a run creates a second run; its surviving original text keeps
its identity. Empty half-runs/text elements are not generated. Unchanged following
runs keep their identities. The new paragraph copies the existing paragraph settings,
including style and numbering. A dedicated Enter at a non-list paragraph's end
applies its declared `next` style if it names an existing paragraph style. Middle
splits and pasted breaks retain the current style. List items continue their existing
numbering; automatic empty-item exit is not yet implemented.
The split run's formatting is copied to its new half. New paragraph wrappers do not
duplicate source paragraph IDs or unrelated attributes.

Joining retains the first paragraph's settings and removes the second paragraph's
settings. Surviving run XML and identities remain intact. Consequently, inherited
formatting from a different second paragraph style follows the first paragraph after
a join; direct run formatting remains unchanged. Unknown package parts are untouched.

New empty split paragraphs carry typing formatting in paragraph-mark properties,
without a placeholder run. Existing source empty runs remain preserved.
Ordinary empty paragraphs are editable. Existing empty run formatting is reused;
otherwise paragraph-mark run properties supply the initial typing format. Deletion
retains the first selected run's typing format for subsequent input at the same caret.
Undo/redo restores package snapshots, source identities, and the logical selection
for direct edits. Adjacent typing or same-direction deletion shares an undo entry
until a pause longer than one second, selection/caret movement, focus change, paste,
composition, formatting, or paragraph command. Save, undo, redo, and intervening
workspace transactions seal the previous group. Each input still commits and advances
the revision immediately; merging history retains the first before-snapshot and last
after-snapshot. Escape does not discard already committed input.

Reveal Codes orders tags at the same text offset in source order, so adjacent runs
appear as siblings rather than as nested opening and closing tags.

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
document structure. The bounded paragraph commands above use this identity-aware
path and keep their selection mapping with each undo transaction.

## GUI commands

The Commands button and Ctrl+Shift+P open a searchable modal palette. Existing
open/save, workspace history, inspector preview, apply/dismiss, search,
replacement, code visibility, and package comparison actions share typed command
IDs and availability checks with their buttons and shortcuts. Unavailable actions
explain the missing document, editable selection, query, or preview.
Checks run again at invocation; busy operations cannot be invoked twice.

Ctrl+Enter previews the inspector text edit. Ctrl+Shift+Enter applies the displayed
preview. Changing an inspector draft or typing directly dismisses its old preview. The
palette does not bypass preview/commit, protected-content policy, or stale-revision
checks. Ctrl+Z and redo shortcuts operate on workspace history in the projection;
inside ordinary text fields they remain native text undo/redo. Escape closes the
palette and returns focus to the document.

Key chords live in `src/ui/keymap.ts`, independently of command definitions and
actions. The active map drives matching, palette shortcut labels, tooltips, and
accessibility hints. **Keybindings** in the sidebar accepts a JSON object mapping
command IDs to chord arrays; `[]` unbinds a command and `{}` restores defaults.
Overrides persist locally and apply immediately. Custom bindings replace conflicting
defaults; ambiguous custom bindings and malformed configuration are refused.
Workspace history bindings defer to ordinary text fields. Editor actions have
their own focus scope. IME composition bypasses shortcut matching.

Example:

```json
{
  "commands.open": ["Ctrl+k"],
  "view.codes": [],
  "text.deleteBackward": ["Alt+h"]
}
```

Editor command IDs include `paragraph.split`, `text.deleteBackward`,
`text.deleteForward`, `format.bold`, `format.italic`, and `format.underline`. The earlier `text.edit` and draft-history IDs are retired.
Chords support Ctrl/Mod (either Control or Meta), Meta, Alt, and Shift. These are
simultaneous chords; sequential bindings such as Ctrl+K then Ctrl+C are not implemented.

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

## Text display

The document projection remains fixed-width and fixed-size, while showing bold,
italic, underline, text color, and the fixed OOXML `w:highlight` palette. Direct
properties override inherited properties; bold and italic toggle through style
chains. Explicit off/none values clear inherited formatting.

Six-digit RGB colors are displayed; `auto` uses the UI text color. Theme colors
use their stored RGB fallback when present; theme tint/shade resolution is not
implemented. Highlighting supports all 16 named colors and `none`, independently
of run shading (`w:shd`). Underline variants use the nearest CSS line style;
word-only, heavy, and compound variants are approximate. Font families and sizes
from the document are intentionally ignored. Native inline editing keeps the formatted projection in place.

The palette follows [OOXML HighlightColorValues](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.highlightcolorvalues?view=openxml-3.0.1).
The formatting toolbar applies bold, italic, underline, RGB/automatic color, and
OOXML highlighting/none to a text selection. Ctrl+B/I/U invoke the same toggles and
are remappable independently of commands. A mixed boolean selection toggles on;
a uniformly enabled selection toggles off. Formatting is an atomic workspace edit,
with undo/redo restoring selection and source snapshots. Only selection boundaries
split runs; full selected runs and unaffected child identities remain intact.
Already-effective formatting is a no-op, avoiding a new run per typed character.
Protected text, selected objects, invalid colors, and surrogate-boundary selections
are refused before publication. Original fonts, sizes, and unrelated run properties
remain preserved.

At a collapsed caret, formatting choices are local typing overrides and do not dirty
the document until text is entered. The inserted text and its format commit together.
Caret movement and history restoration clear overrides; a formatting change starts a
new undo group. Formatting across paragraph breaks affects selected text, not the
paragraph mark or list definitions.

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

Unrestricted paragraph restructuring; automatic empty-list exit; list,
table, hyperlink, and section restructuring; full style/layout resolution and
exotic numbering formats;
comment or revision mutation; granular code-category filters;
paginated preview; in-place saves; persisted workspaces or undo journals;
installer packaging. The desktop is an inspector/editor prototype, not yet a
complete replacement for Word.
