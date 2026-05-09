# Core typed event API

The core reader emits typed .NET records rather than JSON. Adapter layers should translate these records into their own wire formats without changing the core parser contract.

Initial event families:

- `PackageEvent` for package start/end.
- `PartEvent` for OpenXML part boundaries.
- `RelationshipEvent` for package/part relationships.
- `ElementStartEvent` and `ElementEndEvent` for raw XML element boundaries.
- `TextEvent` for raw XML text-like nodes.
- `DiagnosticEvent` for non-fatal reader observations.

The core intentionally does not infer paragraphs, headings, comments, layout, or other app-level document semantics.
