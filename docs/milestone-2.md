# Milestone 2: list labels and paragraph shape

Baseline: `07621f8` (the verified preservation-first workbench).
Status: implemented and verified.

## Goal

Display generated numbering, bullets and approximate paragraph indentation in
both Reveal Codes and the clean review view. Keep labels separate from document
text and make their definitions and inherited properties inspectable.

## Scope and approach

- Add explicit paragraph metadata while keeping existing source-bound text commands.
- Separate paragraph property inheritance, numbering definitions/counters, and
  UI layout into small modules.
- Resolve direct and style-based list membership, multilevel labels, common
  numeric formats, bullets, continuation, start overrides and restart rules.
- Display left/right, first-line and hanging indentation on a consistent scale.
- Preserve every source byte on inspection; text edits still patch only `w:t`.
- Diagnose unsupported or ambiguous numbering rather than inventing exact labels.

## Acceptance checks

1. Synthetic fixtures exercise legal outlines, heading numbering, bullets,
   interruptions, independent lists, restarts, overrides and wrapped text.
2. Expected labels and resolved indentation are covered by engine tests.
3. Desktop tests cover labels with codes on/off, property inspection and wrapping.
4. Compare representative list labels with Microsoft Word where available.
5. Edit text, undo/redo, save/reopen, and assert numbering/style parts are unchanged.

Pagination, exact Word line breaking, list editing, full table/conditional-style
layout, picture bullets and all international numbering systems are outside this
slice. The current review view includes revisions; ambiguous revision/opaque
numbering must be identified rather than represented as Word's final view.

## Delivered

Paragraphs now have explicit metadata alongside the existing source-bound token
stream. `styles.ts` indexes inheritance, `numbering.ts` resolves definitions and
maintains counters, `paragraph.ts` resolves membership/indentation with source
attribution, and `ui/paragraph-layout.ts` maps source measurements to display.

The workbench and CLI show generated labels for decimal, padded decimal,
alphabetic, Roman, and bullet formats. Multilevel templates, style-linked levels,
numbering-style references, start overrides, and level overrides are supported.
Click a label to inspect its list instance, level, resolved template, definition
XML, indentation, and property sources. Labels stay visible with codes hidden
and never become editable document text or search matches.

Indentation covers left/right and first-line/hanging values, including inherited
properties. The UI uses 15 twips per CSS pixel with navigation-friendly limits.
Long labels can widen hanging space to keep labels separate and wrapped lines
aligned in the fixed-width font. This does not change source measurements or
promise Word-equivalent line breaks. Inline codes themselves occupy space in
Reveal Codes; the clean view is the better layout comparison.

## Interoperability findings

Direct Word comparisons informed two regression fixes:

- List instances sharing an abstract definition share a counter stream in Word.
  A start override resets that stream once when that instance/level is first
  encountered. Distinct abstract definitions provide independent streams.
- A direct first-line indent replaces an inherited hanging indent. If both are
  explicitly present in the same indentation element, hanging takes precedence.

Restart handling follows Word's behavior: a specified restart level also resets
at earlier levels. Word ignores `lvlRestart` inside a level override.
See [Microsoft's implementation notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/140a2cd9-2f26-456d-9760-ae6ecef4e8b5),
[style-based numbering](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.numberingproperties?view=openxml-2.13.1),
and [start overrides](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.startoverridenumberingvalue?view=openxml-3.0.1).

## Verification

- New fixture documents: `legal.docx`, `headings.docx`, and `lists.docx`.
- Word read-only comparison: all 34 fixture paragraphs match generated labels
  and left/right/first-line indentation. Another 13 probe paragraphs cover
  shared counters, switching instances, overrides, and indentation precedence.
- Engine tests cover the above plus unsupported formats, legacy bullet glyphs,
  style-link cycles, missing levels, custom restarts, tables and preservation.
- Inspection preserves whole-file bytes; text edits preserve numbering/styles
  parts; labels remain stable after undo/redo and save/reopen.
- Desktop smoke tests exercise labels with codes on/off, scrolling through
  virtualized paragraphs, source inspection and measured wrapping alignment.
  Screenshots of both views were visually inspected.

To repeat the optional Word checks on Windows:

```powershell
node --import tsx scripts/numbering-report.ts
node --import tsx scripts/numbering-probe.ts
./scripts/verify-numbering.ps1
./scripts/verify-numbering.ps1 -Expected test-results/numbering-probes.json
```

## Known boundaries

Unsupported formats, unknown private-font glyphs, picture bullets, section-break
restart extensions and unresolved definitions produce `?` labels and diagnostics.
Paragraph-bearing opaque markup or structural paragraph revisions make following
numbering uncertain for that story. This intentionally avoids claiming final-view
labels for content the review projection cannot interpret completely.

Character-unit indentation and right-to-left layout are diagnosed as approximate;
full table/conditional-style formatting, custom tab-stop layout, and numbering
changes inside tracked revisions remain outside this slice. Word verification
covers the listed fixtures, not every possible numbering construct.
