# Browser/WASM bridge spike

The browser bridge is a v0 spike around the same typed low-level event model as the .NET core, CLI JSONL adapter, and Node native bridge.

The public API is transport-neutral above the byte/blob input boundary: `parseBytes()` yields the same `DocxSaxEvent` objects that `docx-sax/node` yields from `parseFile()`, and `parseBytesBatches()` yields `DocxSaxEvent[]` batches. The shared option is `{ batchSize }`; browser/WASM additionally supports `{ dotnetModuleUrl }` when assets are hosted somewhere other than the package default. `preloadRuntime()` starts the runtime early, and `warmupRuntime()` can then exercise OpenXML/XML/JSON cold paths with an in-memory DOCX before a user selects a private document.

## Shape

- `src/DocxSax.Browser` is a `Microsoft.NET.Sdk.WebAssembly` project targeting `net8.0`/`browser-wasm`.
- `BrowserBridge.BeginParseBytesJsonBatches(byte[] bytes, int batchSize)`, `ReadNextJsonBatch(sessionId)`, and `EndParseBytesJsonBatches(sessionId)` are exported to JavaScript with `[JSExport]` as a pull-based batch stream. The previous one-shot `ParseBytesJsonBatchFrames(...)` remains available as a compatibility fallback.
- `packages/docx-sax/browser.js` hides the .NET runtime details and exposes async generators:
  - `preloadRuntime({ dotnetModuleUrl })` loads and initializes the singleton .NET browser runtime.
  - `warmupRuntime({ dotnetModuleUrl })` parses a tiny generated DOCX in memory and initializes OpenXML package/XML plus DocxSax JSON serialization paths. The warmup promise is cached per runtime URL.
  - `parseBytesBatches(input, { batchSize })` yields `DocxSaxEvent[]` batches.
  - `parseBytes(input, { batchSize })` flattens those batches to individual `DocxSaxEvent` objects.

- The transport is newline-framed JSON batches. Each frame is a JSON array containing the stable event payloads used by the CLI and Node native bridge. This keeps the public JS API batch-oriented instead of requiring a whole-document event array or one JS callback per XML node.

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

The default loader expects the published .NET assets at `./dist/wasm/wwwroot/_framework/dotnet.js` relative to `packages/docx-sax/browser.js`. Pass `dotnetModuleUrl` if a bundler relocates the assets:
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

- OpenXML SDK and `System.IO.Packaging` do run in .NET 8 browser WASM for the small generated fixture.
- A Release publish with trimming disabled produced roughly **46 MiB** in `packages/docx-sax/dist/wasm` with about **17 MiB** of gzip/brotli sidecars. The largest uncompressed payloads are `DocumentFormat.OpenXml.wasm` (~6.1 MiB), `System.Private.CoreLib.wasm` (~4.0 MiB), `System.Private.Xml.wasm` (~3.0 MiB), and `dotnet.native.wasm` (~2.7 MiB).
- Publish trimming currently fails because framework/OpenXML dependencies produce trim warnings, so the project sets `<PublishTrimmed>false</PublishTrimmed>`. That is acceptable for the spike but too large for a polished browser package.
- The browser bridge now uses a pull-based parse session so JavaScript can request one JSON batch at a time and yield to the main thread during large documents. In a headless Chromium profile of the 77,370-event exchanged Big Computer DOCX, pull-based parsing showed first preview at ~476ms median before warmup; adding idle `warmupRuntime()` reduced the measured first preview to ~139ms median, largest parse long task from ~428ms to ~97ms, and total parse/render from ~3.89s to ~3.39s.
- Cold start includes loading the .NET browser runtime and OpenXML assemblies. Consumers should call `preloadRuntime()` and then `warmupRuntime()` during idle time; `warmupRuntime()` avoids private fixtures by generating a minimal DOCX inside the browser bridge. Worker isolation may still be useful if per-batch synchronous parsing remains too visible in UI apps.

## Follow-ups

- Move parsing into a WebWorker wrapper if the remaining per-batch main-thread work is still too visible for production UI.
- Consider a callback/JS-import frame sink if pull-based batches prove insufficient for backpressure or cancellation semantics.
- Investigate trim warnings/root descriptors or a lower-level ZIP/XML path if browser artifact size is a release blocker.
- Add bundler examples for Vite/Next asset copying once packaging strategy is settled.
