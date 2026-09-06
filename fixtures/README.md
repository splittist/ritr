# Fixture provenance

`specs/*.json` are original synthetic ritr test data, dedicated to the public
domain under CC0-1.0. They contain no third-party or personal document content.
`generated/*.docx` are produced from these specifications with the local
Apache-2.0-licensed [docxfix](../../docxfix) utility (version 0.1.0).
The generated files are checked in so normal tests do not require Python.

Run `npm run fixtures` with a sibling `../docxfix` checkout and its installed
`.venv`. Override the location with `DOCXFIX_ROOT`. ZIP timestamps and generator
metadata may vary; reproducibility here means document features, not ZIP identity.

- `plain`: paragraphs, Unicode, ordinary replacement.
- `review`: headings, lists, deletion/insertion, comment anchors and comment parts.
- `sections`: section properties, landscape page, headers and footers.
- `legal`: interrupted multilevel legal list with lower-level restarts.
- `headings`: numbering inherited from Heading1–Heading4 styles.
- `lists`: bullets, multilevel numbering, start overrides, independent lists,
  alphabetic/Roman formats, and inherited/direct indentation with wrapped text.
- `tables`: merged header and vertical cells, numbered cell paragraphs, nested
  table, direct borders/margins/shading, and a borderless signature table.

`tables` uses a docxfix seed plus independent synthetic XML augmentation in
`scripts/table-fixture.ts`. Regenerate it with `npm run fixtures -- tables`.

`lists` starts as a docxfix document, then `scripts/list-fixture.ts` applies
explicit synthetic XML variants for features outside its input spec. That script
does not use the numbering resolver it tests. Run `npm run fixtures -- lists` to
regenerate just that fixture, or `npm run fixtures -- legal headings lists` for
the complete numbering corpus. Ordinary tests need only the checked-in files.

The tests also construct small OOXML packages in `tests/helpers.ts` to isolate
formatting, relationships, tables, notes, opaque extensions, corruption, and
editing boundaries. These synthetic packages have the same CC0 dedication.

Modern comment extension data is tested for byte preservation only. No claim is
made about its completeness or Microsoft Word interoperability.
