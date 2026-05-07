# Browser/WASM bridge spike

The browser bridge is a v0 spike around the same typed low-level event model as the .NET core, CLI JSONL adapter, and Node native bridge.

## Shape

- `src/DocxSax.Browser` is a `Microsoft.NET.Sdk.WebAssembly` project targeting `net8.0`/`browser-wasm`.
- `BrowserBridge.ParseBytesJsonBatchFrames(byte[] bytes, int batchSize)` is exported to JavaScript with `[JSExport]`.
- `packages/docx-sax/index.js` hides the .NET runtime details and exposes async generators:
  - `parseBytesBatches(input, { batchSize })` yields arrays of event objects.
  - `parseBytes(input, { batchSize })` flattens those batches to individual events.
- The transport is newline-framed JSON batches. Each frame is a JSON array containing the stable event payloads used by the CLI and Node native bridge. This keeps the public JS API batch-oriented instead of requiring a whole-document event array or one JS callback per XML node.

## Usage

```js
import { parseBytes, parseBytesBatches } from 'docx-sax/browser';

const response = await fetch('/document.docx');
const bytes = new Uint8Array(await response.arrayBuffer());

for await (const batch of parseBytesBatches(bytes, { batchSize: 256 })) {
  // batch is an array of package/part/relationship/element/text/end events
}

for await (const event of parseBytes(bytes)) {
  console.log(event.type, event.ordinal);
}
```

The default loader expects the published .NET assets at `./dist/wasm/wwwroot/_framework/dotnet.js` relative to `packages/docx-sax/index.js`. Pass `dotnetModuleUrl` if a bundler relocates the assets:

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
npm test
```

`npm test` publishes the .NET browser-wasm project, starts a Vite server, loads the bridge in Chromium, fetches the public/generated `simple.docx` fixture, and verifies package/part/element/text/end events are observed.

CI runs the same browser smoke on the Linux leg. It is intentionally Linux-only for now because Playwright browser installation is the heaviest new dependency and the Node native wrapper is also Linux-first.

## Feasibility notes

Observed on this spike:

- OpenXML SDK and `System.IO.Packaging` do run in .NET 8 browser WASM for the small generated fixture.
- A Release publish with trimming disabled produced roughly **46 MiB** in `packages/docx-sax/dist/wasm` with about **17 MiB** of gzip/brotli sidecars. The largest uncompressed payloads are `DocumentFormat.OpenXml.wasm` (~6.1 MiB), `System.Private.CoreLib.wasm` (~4.0 MiB), `System.Private.Xml.wasm` (~3.0 MiB), and `dotnet.native.wasm` (~2.7 MiB).
- Publish trimming currently fails because framework/OpenXML dependencies produce trim warnings, so the project sets `<PublishTrimmed>false</PublishTrimmed>`. That is acceptable for the spike but too large for a polished browser package.
- The current C# export returns all batch frames as one string after parsing. The public JavaScript API is batch-oriented, but true streaming/backpressure would require a follow-up callback, JS import, or WebWorker message channel design.
- Cold start includes loading the .NET browser runtime and OpenXML assemblies. Consumers should keep the runtime singleton warm and prefer worker isolation for UI apps.

## Follow-ups

- Move parsing into a WebWorker wrapper so DOCX parsing and .NET startup do not block the UI thread.
- Replace the one-shot string return with a real streaming frame sink once the JS interop shape is proven.
- Investigate trim warnings/root descriptors or a lower-level ZIP/XML path if browser artifact size is a release blocker.
- Add bundler examples for Vite/Next asset copying once packaging strategy is settled.
