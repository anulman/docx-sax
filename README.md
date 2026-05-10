# docx-sax

`docx-sax` is a standalone developer tool for parsing `.docx`/OpenXML packages as a typed, low-level event stream.

The core is a .NET 8 library built on the Microsoft Open XML SDK. It exposes OpenXML-oriented events such as package, part, relationship, element, text, and diagnostic records. It deliberately preserves low-level details: part URIs, relationship types, XML names/namespaces, attributes, depth, paths, and event ordinals.

For adapter-selection guidance written for humans and LLM/agent consumers, see [LLMS.md](LLMS.md). The Next.js demo also serves the same file at `/LLMS.md` when deployed.

## Current status

This repository currently contains the core .NET library scaffold, a minimal forward-only reader for simple Word documents, a CLI JSONL adapter over the typed event stream, an initial Linux x64 Node N-API wrapper backed by a .NET Native AOT shared library, and a browser/WASM bridge spike, and namespaced npm runtime plugin packages.

## JavaScript API model

The public JavaScript wrappers share one event model and differ only by input transport:

- `@docx-sax/node` accepts file paths with `parseFile(path, options)` and `parseFileBatches(path, options)` and loads the `@docx-sax/native-linux-x64` payload on Linux x64.
- `@docx-sax/browser` accepts bytes/blob inputs with `parseBytes(input, options)` and `parseBytesBatches(input, options)` and owns the browser/WASM adapter.

All four functions return async iterables over the same `DocxSaxEvent` object union. The non-batched helpers yield one event at a time; the batched helpers yield `DocxSaxEvent[]` chunks and accept the shared `{ batchSize }` option. Transport-specific options stay transport-specific (`nativeLibraryPath` for Node, `dotnetModuleUrl` for browser/WASM).

The package ships TypeScript declarations for the common event union: package/part lifecycle events, relationships, XML element starts/ends, text events, diagnostics, attributes, ordinals, paths, namespaces, and part URIs. The wrappers intentionally do not expose a whole-document array API; consumers should iterate the stream or batches.

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
src/DocxSax.Browser/  # .NET 8 browser-wasm JSExport bridge spike
packages/node/                  # @docx-sax/node user-facing Node adapter
packages/native-linux-x64/              # @docx-sax/native-linux-x64 Linux x64 native runtime payload
packages/native-darwin-arm64/ # @docx-sax/native-darwin-arm64 placeholder payload
packages/native-darwin-x64/   # @docx-sax/native-darwin-x64 placeholder payload
packages/native-win32-x64/    # @docx-sax/native-win32-x64 placeholder payload
packages/browser-wasm/          # @docx-sax/browser browser adapter + WASM assets
demos/nextjs-wasm/    # Next.js browser/WASM demo
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

npm install
npm run build --workspace docx-sax
npm run test --workspace docx-sax
npm run build --workspace @docx-sax/native-linux-x64
npm run test --workspace @docx-sax/native-linux-x64
npm run build --workspace @docx-sax/browser
npx playwright install chromium
npm run test --workspace @docx-sax/browser
```

CI runs these checks on Ubuntu, Windows, and macOS with .NET 8. It also uploads Cobertura coverage artifacts, validates that CLI JSONL output parses line-by-line as JSON, installs the packed `DocxSax.Tool` from the local package output, and performs a Native AOT publish check for the CLI on each runner RID. The Linux leg additionally builds/tests the Node N-API wrapper and runs the browser/WASM Vite + Playwright smoke.

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
import { parseFile, parseFileBatches } from '@docx-sax/node';

for await (const event of parseFile('document.docx')) {
    console.log(event.type, event.ordinal);
}

for await (const batch of parseFileBatches('document.docx', { batchSize: 256 })) {
    // batch is DocxSaxEvent[] from the same model used by @docx-sax/browser
}
```

See [`docs/native-node.md`](docs/native-node.md) for local validation and current transport caveats. The v0 native plugin validates on Linux x64 first and is published as `@docx-sax/native-linux-x64`; `@docx-sax/node` is the ergonomic Node import path.

## Browser/WASM usage

```js
import { parseBytes, parseBytesBatches } from '@docx-sax/browser';

const response = await fetch('/document.docx');
const bytes = new Uint8Array(await response.arrayBuffer());

for await (const batch of parseBytesBatches(bytes, { batchSize: 256 })) {
    // batch is DocxSaxEvent[] from the same model used by @docx-sax/node
}

for await (const event of parseBytes(bytes)) {
    console.log(event.type, event.ordinal);
}
```

See [`docs/browser-wasm.md`](docs/browser-wasm.md) for local validation, artifact-size measurements, and current caveats. The spike validates that OpenXML SDK can parse a small generated DOCX fixture in browser WASM, but the published runtime is still large and remains an alpha plugin package (`@docx-sax/browser`).

## Next.js WASM demo

`demos/nextjs-wasm` is an isolated Next.js app that loads `@docx-sax/browser` in a client component, accepts a `.docx` upload, parses it through the browser WASM bridge, and renders a simple text preview from `text` events grouped by paragraph ends. It also exposes the safe public smoke fixture at `/fixtures/simple.docx`.

```bash
cd demos/nextjs-wasm
npm install
npm run build:wasm   # publishes packages/browser-wasm WASM assets, then copies _framework/ into public/docx-sax/
npm run build        # regular Next/Vercel build; intentionally does not run dotnet publish
npm run test:e2e     # uploads public/fixtures/simple.docx and verifies WASM parse + preview render
```

For a single local validation command, run `npm test`. Vercel should use the plain `npm run build` path after WASM assets have been prepared by CI or a package/artifact step; generic Vercel builds should not republish the heavy .NET browser workload.

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
