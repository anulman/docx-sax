# Node native wrapper

`packages/node` contains a v0 Node wrapper for the typed `DocxSaxReader` event stream. It is intentionally not a CLI/stdout adapter: Node calls a N-API addon, the addon loads the .NET Native AOT shared library, and the native library calls back with JSON batches.

## Usage

```js
import { parseFile, parseFileBatches } from '@docx-sax/node';

for await (const event of parseFile('document.docx')) {
  console.log(event.type, event.ordinal);
}

for await (const batch of parseFileBatches('document.docx', { batchSize: 256 })) {
  // batch is an array of low-level package/part/relationship/element/text/end events
}
```

The default export also exposes `parseFile` and `parseFileBatches`.

## Local validation

From the repo root:

```bash
cd packages/node
npm install
npm run build
npm test
```

`npm run build` publishes `src/DocxSax.Native` as a Native AOT shared library into `packages/node/native/linux-x64/`, then builds the N-API addon with `node-gyp`.

## Transport caveats

- Current validation is Linux x64 only. The C# export is portable in shape, but package scripts and CI only build/test the `linux-x64` Native AOT library in this PR.
- The bridge transports batches as JSON strings across the C ABI. That keeps the ABI small and version-tolerant while the event schema is still pre-1.0.
- `parseFileBatches()` exposes a batched async iterable. Internally the v0 N-API worker completes a native parse on a libuv worker and then resolves collected JSON batches; it avoids a whole-document event-array API, but it is not yet true backpressured native streaming.
- Native parse failures reject the async iterator promise rather than writing diagnostics to stdout/stderr.
