# Public DOCX fixtures

These fixtures are intentionally tiny generated Open XML packages committed for public regression testing.

- `simple.docx` contains one main document part with the text `Hello DOCX SAX`.
- `unknown-inline.docx` contains a custom `x:unknown` element in namespace `urn:docx-sax:test` to verify low-level XML preservation.
- `corrupt.docx` is not a ZIP/OpenXML package and is used to verify clean CLI failure behavior.

No private documents or private validation corpora are included.
