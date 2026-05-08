# docx-sax

`docx-sax` is a standalone developer tool for parsing `.docx`/OpenXML packages as a typed, low-level event stream.

The core is a .NET 8 library built on the Microsoft Open XML SDK. It exposes OpenXML-oriented events such as package, part, relationship, element, text, and diagnostic records. It deliberately preserves low-level details: part URIs, relationship types, XML names/namespaces, attributes, depth, paths, and event ordinals.

## Current status

This repository currently contains the core .NET library scaffold, a minimal forward-only reader for simple Word documents, a CLI JSONL adapter over the typed event stream, and an initial Linux x64 Node N-API wrapper backed by a .NET Native AOT shared library.

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
src/DocxSax.Native/   # .NET 8 Native AOT C ABI bridge for Node/native hosts
packages/docx-sax/   # Node N-API wrapper package
test/DocxSax.Tests/   # generated DOCX fixtures, golden JSONL, and tests
docs/                 # design notes as the project grows
```

## Stack plan

1. Core .NET typed event API.
2. CLI JSONL adapter over the typed API (this layer).
3. Node N-API bindings (initial Linux x64 v0 is present).
4. WASM package.
5. Next.js demo using the public adapters.

## Build, test, and validation

```bash
dotnet restore
dotnet build --configuration Release
dotnet test --configuration Release
dotnet format --verify-no-changes --verbosity minimal
dotnet pack --configuration Release --output artifacts/packages

cd packages/docx-sax
npm install
npm run build
npm test
```

CI runs these checks on Ubuntu, Windows, and macOS with .NET 8. It also uploads Cobertura coverage artifacts, validates that CLI JSONL output parses line-by-line as JSON, installs the packed `DocxSax.Tool` from the local package output, and performs a Native AOT publish check for the CLI on each runner RID. The Linux leg additionally builds and tests the Node N-API wrapper.

NuGet package IDs:

- `DocxSax` for the core library.
- `DocxSax.Tool` for the `docx-sax` .NET global/local tool command.

## Version strategy

`DocxSax` and `DocxSax.Tool` use one repo-wide SemVer, centralized in `Directory.Build.props`.

Early public packages start at `0.1.0-alpha.1`. The project stays in `0.x` while the typed event contract and JSONL schema are evolving. Use prerelease versions for all early packages, and publish only the top merged artifact from a stacked PR series rather than every intermediate branch.

Version bump guidance:

- Patch/prerelease increments (`0.1.0-alpha.N`) for validation, fixtures, docs, CI, and compatible implementation changes.
- Minor prerelease increments (`0.2.0-alpha.N`, etc.) for meaningful event-model or schema changes.
- `1.0.0` only after typed event names/properties are stable, JSONL compatibility is documented, and Node/WASM adapters have proven the core API shape.

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

## Node wrapper usage

```js
import { parseFile, parseFileBatches } from 'docx-sax/node';

for await (const event of parseFile('document.docx')) {
    console.log(event.type, event.ordinal);
}

for await (const batch of parseFileBatches('document.docx', { batchSize: 256 })) {
    // batch is an array of low-level events from the native bridge
}
```

See [`docs/node-native.md`](docs/node-native.md) for local validation and current transport caveats. The v0 package validates on Linux x64 first.

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
