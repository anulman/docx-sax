# Browser/WASM bridge spike

The browser bridge is a v0 spike around the same typed low-level event model as the .NET core, CLI JSONL adapter, and Node native bridge.

The public API is transport-neutral above the byte/blob input boundary: `parseBytes()` yields the same `DocxSaxEvent` objects that `docx-sax/node` yields from `parseFile()`, and `parseBytesBatches()` yields `DocxSaxEvent[]` batches. The shared option is `{ batchSize }`; browser/WASM additionally supports `{ dotnetModuleUrl }` when assets are hosted somewhere other than the package default. `preloadRuntime()` starts the runtime early, and `warmupRuntime()` can then exercise OpenXML/XML/event-mapping cold paths with an in-memory DOCX before a user selects a private document.

## Shape

- `src/DocxSax.Browser` is a `Microsoft.NET.Sdk.WebAssembly` project targeting `net8.0`/`browser-wasm`.
- `BrowserBridge.BeginParseBytesBatches(byte[] bytes, int batchSize)`, `ReadNextBatch(sessionId)`, and `EndParseBytesBatches(sessionId)` are exported to JavaScript with `[JSExport]` as a pull-based batch stream. The bridge returns positional primitive/object-array rows; `packages/docx-sax/browser.js` maps those rows to public `DocxSaxEvent` objects.
- `packages/docx-sax/browser.js` hides the .NET runtime details and exposes async generators:
  - `preloadRuntime({ dotnetModuleUrl })` loads and initializes the singleton .NET browser runtime.
  - `warmupRuntime({ dotnetModuleUrl })` parses a tiny generated DOCX in memory and initializes OpenXML package/XML plus DocxSax event mapping paths. The warmup promise is cached per runtime URL.
  - `parseBytesBatches(input, { batchSize })` yields `DocxSaxEvent[]` batches.
  - `parseBytes(input, { batchSize })` flattens those batches to individual `DocxSaxEvent` objects.

- The browser runtime boundary does not use JSON strings. JSON remains reserved for the CLI JSONL adapter; browser and Node both map typed/native fields to the shared JS event object model.

## Usage

```js
import { parseBytes, parseBytesBatches, preloadRuntime, warmupRuntime } from 'docx-sax/browser';

requestIdleCallback(async () => {
  await preloadRuntime();
  await warmupRuntime();
});

const response = await fetch('/document.docx');
const bytes = new Uint8Array(await response.arrayBuffer());

for await (const batch of parseBytesBatches(bytes, { batchSize: 256 })) {
  // batch is DocxSaxEvent[], matching docx-sax/node parseFileBatches()
}

for await (const event of parseBytes(bytes)) {
  console.log(event.type, event.ordinal);
}
```

TypeScript declarations are included for the shared `DocxSaxEvent` union: `package`, `part`, `relationship`, `element`, `text`, `end`, and `diagnostic` events plus shared XML attribute shapes.

The default loader expects the published .NET assets at `./dist/wasm/wwwroot/_framework/dotnet.js` relative to `packages/docx-sax/browser.js`. Pass `dotnetModuleUrl` if a bundler relocates the assets:

```js
for await (const batch of parseBytesBatches(bytes, {
  dotnetModuleUrl: '/assets/docx-sax/_framework/dotnet.js',
})) {
  // ...
}
```

## Local validation

```bash
cd packages/docx-sax
npm ci
npx playwright install chromium
npm run test:browser
```

`npm run test:browser` publishes the .NET browser-wasm project, starts a Vite server, loads the bridge in Chromium, fetches the public/generated `simple.docx` fixture, and verifies package/part/element/text/end events are observed.

CI runs the same browser smoke on the Linux leg. It is intentionally Linux-only for now because Playwright browser installation is the heaviest new dependency and the Node native wrapper is also Linux-first.

## Feasibility notes

Observed on this spike:

- OpenXML SDK and `System.IO.Packaging` run in .NET 8 browser WASM for the generated fixtures covered by browser smoke.
- The browser project now publishes with trimming enabled. The current trimmed `_framework` directory is roughly **13.9 MiB** on disk including generated gzip/brotli sidecars, or **8.8 MiB** for the raw files that npm actually includes after excluding `*.pdb*`, `*.br`, and `*.gz` sidecars.
- The largest raw browser payloads are currently `DocumentFormat.OpenXml.wasm` (~3.5 MiB), `dotnet.native.wasm` (~1.4 MiB), `System.Private.CoreLib.wasm` (~1.4 MiB), `System.Private.Xml.wasm` (~0.5 MiB), and `System.Linq.Expressions.wasm` (~0.3 MiB).
- Trimming still emits a framework linker warning from `System.Linq.Expressions` (`IL2104`), so the browser project keeps `ILLinkTreatWarningsAsErrors=false`. Browser smoke/e2e coverage is the safety net until that upstream/library warning can be removed or rooted more precisely.
- `npm pack --dry-run` for this combined Node+browser package is currently about **13.0 MiB packed** / **34.5 MiB unpacked** / **41 files**. Most unpacked size is the Linux native library plus the trimmed browser runtime.
- The browser bridge uses a pull-based parse session so JavaScript can request one event-object batch at a time and yield during large documents. In a headless Chromium profile of the 77,370-event exchanged Big Computer DOCX, pull-based parsing showed first preview at ~476ms median before warmup; adding idle `warmupRuntime()` reduced the measured first preview to ~139ms median, largest parse long task from ~428ms to ~97ms, and total parse/render from ~3.89s to ~3.39s.
- Cold start includes loading the .NET browser runtime and OpenXML assemblies. Consumers should call `preloadRuntime()` and then `warmupRuntime()` during idle time; `warmupRuntime()` avoids private fixtures by generating a minimal DOCX inside the browser bridge.

## Follow-ups

- Provide bundler examples for Vite/Next asset copying once packaging strategy is settled.
- Document a recommended Web Worker wrapper for UI apps that need stronger main-thread isolation than pull-based batches alone provide.
- Add explicit cancellation/cleanup semantics to the browser parse session API if product usage shows abandoned parses are common.
- Keep reducing browser payload size if the trimmed runtime remains too heavy for the eventual release target.
