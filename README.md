# docx-sax

`docx-sax` is a standalone developer tool for parsing `.docx`/OpenXML packages as a typed, low-level event stream.

The core is a .NET 8 library built on the Microsoft Open XML SDK. It exposes OpenXML-oriented events such as package, part, relationship, element, text, and diagnostic records. It deliberately preserves low-level details: part URIs, relationship types, XML names/namespaces, attributes, depth, paths, and event ordinals.

## Current status

This repository currently contains the core .NET library scaffold, a minimal forward-only reader for simple Word documents, and a CLI JSONL adapter over the typed event stream. Later stacked work can build Node N-API, WASM, and demo surfaces on top of the typed core API.

## Non-goals

- No JSON-first API in the core library.
- No application-level semantic events such as paragraphs, headings, comments, or product-specific concepts.
- No authoring workflow or document transformation API.
- No product-layer concepts, this is a walker-style SAX parser

## Repository layout

```text
DocxSax.sln
src/DocxSax/          # .NET 8 typed event reader library
src/DocxSax.Tool/     # .NET 8 CLI/global-tool JSONL adapter
test/DocxSax.Tests/   # generated DOCX fixtures, golden JSONL, and tests
docs/                 # design notes as the project grows
```

## Stack plan

1. Core .NET typed event API.
2. CLI JSONL adapter over the typed API (this layer).
3. Node N-API bindings.
4. WASM package.
5. Next.js demo using the public adapters.

## Build and test

```bash
dotnet restore
dotnet build --configuration Release
dotnet test --configuration Release
```

## CLI usage

```bash
dotnet run --project src/DocxSax.Tool -- parse document.docx --jsonl
```

The `docx-sax` global tool command shape is:

```bash
docx-sax parse input.docx --jsonl
```

JSONL is an adapter format, not the core primitive. Each line is one typed event serialized deterministically to stdout. Human diagnostics and parse failures are written to stderr; invalid or corrupt input exits nonzero.

Event `type` values:

- `package` with `phase: "start" | "end"`
- `part` with `phase: "start" | "end"`, part URI, content type, and relationship type
- `relationship` with source URI, relationship id/type, target URI, and external flag
- `element` for XML element starts, including names, namespace, depth, path, empty-element flag, and attributes
- `text` for XML text-like nodes, including text, depth, path, and whitespace flag
- `end` for XML element ends
- `diagnostic` for non-fatal reader diagnostics

## Minimal library usage

```csharp
await using var stream = File.OpenRead("document.docx");
var reader = new DocxSaxReader();

foreach (DocxEvent docxEvent in reader.Read(stream))
{
    // Pattern-match PackageEvent, PartEvent, RelationshipEvent,
    // ElementStartEvent, ElementEndEvent, TextEvent, DiagnosticEvent.
}
```

## License

MIT
