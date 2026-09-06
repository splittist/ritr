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

The tests also construct small OOXML packages in `tests/helpers.ts` to isolate
formatting, relationships, tables, notes, opaque extensions, corruption, and
editing boundaries. These synthetic packages have the same CC0 dedication.

Modern comment extension data is tested for byte preservation only. No claim is
made about its completeness or Microsoft Word interoperability.
