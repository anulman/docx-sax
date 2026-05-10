# docx-sax for agents and integrators

This file is the short integration guide for humans, LLMs, and coding agents that need to choose the right `docx-sax` runtime adapter without reading the whole repository.

## What docx-sax is

`docx-sax` parses `.docx`/OpenXML packages as a low-level SAX-style event stream. The canonical parser is the .NET `DocxSax` library, built on the Microsoft Open XML SDK. It emits typed events for package boundaries, parts, relationships, XML element starts/ends, text nodes, and diagnostics.

The key design point is runtime neutrality: .NET, CLI JSONL, Node, and browser/WASM adapters expose the same event model instead of inventing separate document-shaped APIs. `docx-sax` does **not** infer paragraphs, headings, comments, layout, or product-level semantics.

## Namespaced plugin/adapter model

JavaScript v0 uses scoped packages under the `@docx-sax` org:

- `@docx-sax/node` for Node.js file-path parsing through `@docx-sax/native-linux-x64` (Linux x64 alpha).
- `@docx-sax/browser` for browser/Vercel/Next.js parsing through browser WASM; `@docx-sax/browser` owns the WASM `_framework` assets.

There is intentionally no generic root import in the current package. Pick the plugin/adapter that matches the runtime so bundlers, agents, and applications do not accidentally pull in the wrong transport.

The .NET packages are the parser foundation:

- `DocxSax` is the core typed event reader library.
- `DocxSax.Tool` provides the `docx-sax` CLI JSONL adapter.

Current version in this repository: `0.1.0-alpha.1`. Treat the event schema, JS adapter packaging, and runtime support as preview/pre-1.0.

## Install current pieces

Packages may not be published from this branch yet. If a package is not available from NuGet/npm, build from source using the commands below.

### .NET library

When published:

```bash
dotnet add package DocxSax --prerelease
```

From this repository:

```bash
dotnet restore
dotnet build --configuration Release
```

### .NET CLI JSONL tool

When published:

```bash
dotnet tool install --global DocxSax.Tool --prerelease
docx-sax parse document.docx --jsonl
```

From this repository:

```bash
dotnet run --project src/DocxSax.Tool -- parse document.docx --jsonl
```

### JavaScript package

When published:

```bash
# Node.js adapter:
npm install @docx-sax/node
# Browser/WASM adapter:
npm install @docx-sax/browser
```

From this repository:

```bash
npm install
npm run build --workspace @docx-sax/native-linux-x64
npm run test --workspace @docx-sax/native-linux-x64
npm run build --workspace @docx-sax/browser
npm run test --workspace @docx-sax/browser
```

The current Native adapter validates on Linux x64 first. The browser/WASM adapter is functional but still a large preview artifact; see `docs/browser-wasm.md` before treating it as production-ready.

## Minimal usage

### .NET core library

```csharp
using DocxSax;

await using var stream = File.OpenRead("document.docx");
var reader = new DocxSaxReader();

foreach (DocxEvent docxEvent in reader.Read(stream))
{
    switch (docxEvent)
    {
        case TextEvent text when !text.IsWhitespace:
            Console.WriteLine(text.Text);
            break;
        case DiagnosticEvent diagnostic:
            Console.Error.WriteLine(diagnostic.Message);
            break;
    }
}
```

### CLI JSONL process boundary

```bash
docx-sax parse document.docx --jsonl
```

Each stdout line is one JSON event. Diagnostics and parse errors go to stderr; invalid/corrupt input exits nonzero.

Example consumer:

```js
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';

const child = spawn('docx-sax', ['parse', 'document.docx', '--jsonl'], {
  stdio: ['ignore', 'pipe', 'inherit'],
});

for await (const line of createInterface({ input: child.stdout })) {
  const event = JSON.parse(line);
  if (event.type === 'text' && !event.isWhitespace) {
    console.log(event.text);
  }
}
```

### Node.js native adapter

```js
import { parseFile, parseFileBatches } from '@docx-sax/node';

for await (const event of parseFile('document.docx')) {
  console.log(event.type, event.ordinal);
}

for await (const batch of parseFileBatches('document.docx', { batchSize: 256 })) {
  // batch is DocxSaxEvent[] from the shared event model.
}
```

### Browser/WASM adapter

```js
import {
  parseBytes,
  parseBytesBatches,
  preloadRuntime,
  warmupRuntime,
} from '@docx-sax/browser';

await preloadRuntime({ dotnetModuleUrl: '/docx-sax/_framework/dotnet.js' });
await warmupRuntime({ dotnetModuleUrl: '/docx-sax/_framework/dotnet.js' });

const response = await fetch('/document.docx');
const bytes = new Uint8Array(await response.arrayBuffer());

for await (const event of parseBytes(bytes, {
  dotnetModuleUrl: '/docx-sax/_framework/dotnet.js',
})) {
  console.log(event.type, event.ordinal);
}

for await (const batch of parseBytesBatches(bytes, {
  batchSize: 256,
  dotnetModuleUrl: '/docx-sax/_framework/dotnet.js',
})) {
  // batch is DocxSaxEvent[], matching @docx-sax/node batches.
}
```

The Next.js demo in `demos/nextjs-wasm` hosts the WASM runtime under `/docx-sax/_framework/` and should expose this file at `/LLMS.md` when deployed.

## Choosing an adapter

- Use **`DocxSax` (.NET library)** for native/server workloads where you want typed records, direct stream handling, and no process boundary.
- Use **`DocxSax.Tool` / `docx-sax parse --jsonl`** for shell pipelines, debugging, cross-language process boundaries, and agents that can safely consume line-delimited JSON.
- Use **`@docx-sax/node`** for Node applications that can use the native bridge and currently target the validated Linux x64 path.
- Use **`@docx-sax/browser`** for browser, Next.js, and Vercel demos where users upload private DOCX files and parsing should happen client-side through WASM.

## Event model checklist

Expect events with these lowercase `type` values across adapters:

- `package` with `phase: "start" | "end"`
- `part` with `phase: "start" | "end"`, part URI, content type, and relationship type
- `relationship` with source URI, relationship id/type, target URI, and external flag
- `element` for XML element starts, names, namespace fields, depth, path, empty-element flag, and attributes
- `text` for XML text-like nodes, text, depth, path, and whitespace flag
- `end` for XML element ends
- `diagnostic` for non-fatal parser observations

For full details, read `README.md`, `docs/core-api.md`, `docs/cli-jsonl.md`, `docs/native-node.md`, and `docs/browser-wasm.md`.
