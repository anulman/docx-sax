# CLI JSONL adapter

The CLI is intentionally a thin adapter over the typed `DocxSaxReader` event stream:

```bash
docx-sax parse input.docx --jsonl
```

It writes deterministic JSON Lines to stdout and reserves stderr for human-readable diagnostics/errors. A corrupt or invalid DOCX exits nonzero without mixing diagnostics into stdout.

JSONL is not the core API. The library continues to expose typed events (`PackageEvent`, `PartEvent`, `RelationshipEvent`, `ElementStartEvent`, `TextEvent`, `ElementEndEvent`, and `DiagnosticEvent`). The CLI maps those records to stable lowercase event `type` values: `package`, `part`, `relationship`, `element`, `text`, `end`, and `diagnostic`.

Public regression fixtures live under `test/DocxSax.Tests/Fixtures`, with golden snapshots in `test/DocxSax.Tests/Golden`.
