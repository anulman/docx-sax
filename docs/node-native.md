# Node native wrapper

`packages/docx-sax` contains a v0 Node wrapper for the typed `DocxSaxReader` event stream. It is intentionally not a CLI/stdout adapter: Node calls a N-API addon, the addon loads the .NET Native AOT shared library, and the native library calls back with JSON batches.

The public API is transport-neutral above the file input boundary: `parseFile()` yields the same `DocxSaxEvent` objects that `@docxdocx-sax/browser` yields from `parseBytes()`, and `parseFileBatches()` yields `DocxSaxEvent[]` batches. The shared option is `{ batchSize }`; Node additionally supports `{ nativeLibraryPath }` for local/native bridge validation.

## Usage

```js
import { parseFile, parseFileBatches } from 'docx-sax/node';

for await (const event of parseFile('document.docx')) {
  console.log(event.type, event.ordinal);
}

for await (const batch of parseFileBatches('document.docx', { batchSize: 256 })) {
  // batch is DocxSaxEvent[], matching @docxdocx-sax/browser parseBytesBatches()
}
```

The default export also exposes `parseFile` and `parseFileBatches`. TypeScript declarations are included for the shared `DocxSaxEvent` union: `package`, `part`, `relationship`, `element`, `text`, `end`, and `diagnostic` events plus shared XML attribute shapes.

## Local validation

From the repo root:

```bash
cd packages/docx-sax
npm install
npm run build
npm test
```

`npm run build:node` publishes `src/DocxSax.Native` as a Native AOT shared library into `packages/docx-sax/native/linux-x64/`, then builds the N-API addon with `node-gyp`.

## Transport caveats

- Current validation is Linux x64 only. The C# export is portable in shape, but package scripts and CI only build/test the `linux-x64` Native AOT library in this PR.
- The bridge transports batches as JSON strings across the C ABI. That keeps the ABI small and version-tolerant while the event schema is still pre-1.0.
- `parseFileBatches()` exposes a batched async iterable backed by a live native parse on a libuv worker. The addon queues only a small number of JSON batches at a time and blocks the native callback until JavaScript consumes more, so the Node bridge no longer buffers a whole-document event list.
- Native parse failures reject the async iterator promise rather than writing diagnostics to stdout/stderr.
