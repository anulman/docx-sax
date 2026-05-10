const runtimePromises = new Map();
const warmupPromises = new Map();

async function toUint8Array(input) {
  if (input instanceof Uint8Array) {
    return input;
  }

  if (input instanceof ArrayBuffer) {
    return new Uint8Array(input);
  }

  if (ArrayBuffer.isView(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }

  if (typeof Blob !== 'undefined' && input instanceof Blob) {
    return new Uint8Array(await input.arrayBuffer());
  }

  throw new TypeError('parseBytesBatches(input) expects a Uint8Array, ArrayBuffer, typed-array view, or Blob');
}

async function loadRuntime(options = {}) {
  const dotnetModuleUrl = options.dotnetModuleUrl ?? './dist/wasm/wwwroot/_framework/dotnet.js';
  let runtimePromise = runtimePromises.get(dotnetModuleUrl);
  if (!runtimePromise) {
    runtimePromise = import(/* @vite-ignore */ /* webpackIgnore: true */ dotnetModuleUrl).then(async ({ dotnet }) => {
      const runtime = await dotnet.withDiagnosticTracing(false).create();
      const config = runtime.getConfig();
      const exports = await runtime.getAssemblyExports(config.mainAssemblyName);
      return { runtime, exports };
    });
    runtimePromises.set(dotnetModuleUrl, runtimePromise);
  }

  return runtimePromise;
}

function browserBridge(exports) {
  return exports.DocxSax.Browser.BrowserBridge;
}

function supportsPullBatches(bridge) {
  return typeof bridge.BeginParseBytesBatches === 'function'
    && typeof bridge.ReadNextBatch === 'function'
    && typeof bridge.EndParseBytesBatches === 'function';
}

function mapAttributeRow(row) {
  return {
    name: row[0],
    localName: row[1],
    prefix: row[2],
    namespaceUri: row[3],
    value: row[4],
  };
}

function mapEventRow(row) {
  switch (row[0]) {
    case 0:
    case 1:
      return { type: 'package', phase: row[2], ordinal: row[1] };
    case 2:
    case 3:
      return {
        type: 'part',
        phase: row[2],
        ordinal: row[1],
        uri: row[3],
        contentType: row[4],
        relationshipType: row[5],
      };
    case 4:
      return {
        type: 'relationship',
        ordinal: row[1],
        sourceUri: row[6],
        id: row[7],
        relationshipType: row[5],
        targetUri: row[8],
        isExternal: row[9],
      };
    case 5:
      return {
        type: 'element',
        ordinal: row[1],
        partUri: row[10],
        name: row[11],
        localName: row[12],
        prefix: row[13],
        namespaceUri: row[14],
        depth: row[15],
        path: row[16],
        isEmptyElement: row[17],
        attributes: row[18].map(mapAttributeRow),
      };
    case 6:
      return {
        type: 'end',
        ordinal: row[1],
        partUri: row[10],
        name: row[11],
        localName: row[12],
        prefix: row[13],
        namespaceUri: row[14],
        depth: row[15],
        path: row[16],
      };
    case 7:
      return {
        type: 'text',
        ordinal: row[1],
        partUri: row[10],
        text: row[19],
        depth: row[15],
        path: row[16],
        isWhitespace: row[20],
      };
    case 8: {
      const diagnostic = { type: 'diagnostic', ordinal: row[1], message: row[21] };
      if (row[10] != null) {
        diagnostic.partUri = row[10];
      }
      return diagnostic;
    }
    default:
      throw new Error(`Unknown DocxSax browser event kind: ${row[0]}`);
  }
}

function mapEventBatch(batch) {
  return batch.map(mapEventRow);
}

function taskYield() {
  if (typeof scheduler !== 'undefined' && typeof scheduler.yield === 'function') {
    return scheduler.yield();
  }

  return new Promise((resolve) => setTimeout(resolve, 0));
}

/**
 * Start loading and initializing the browser WASM runtime before the first parse.
 *
 * The initialized runtime is cached by dotnetModuleUrl and reused by parseBytes/parseBytesBatches.
 * Applications should schedule this during idle time when possible so startup work does not compete
 * with first paint or other critical interactions.
 *
 * @param {{ dotnetModuleUrl?: string }} [options]
 * @returns {Promise<void>}
 */
export async function preloadRuntime(options = {}) {
  await loadRuntime(options);
}

/**
 * Warm OpenXML package, XML reader, and DocxSax event mapping paths after the WASM runtime loads.
 *
 * This parses a tiny in-memory DOCX generated inside the bridge, so applications can call it during
 * idle time without exposing a user document. The warmup is cached by dotnetModuleUrl and runs once
 * per loaded runtime.
 *
 * @param {{ dotnetModuleUrl?: string }} [options]
 * @returns {Promise<void>}
 */
export async function warmupRuntime(options = {}) {
  const dotnetModuleUrl = options.dotnetModuleUrl ?? './dist/wasm/wwwroot/_framework/dotnet.js';
  let warmupPromise = warmupPromises.get(dotnetModuleUrl);
  if (!warmupPromise) {
    warmupPromise = loadRuntime(options).then(({ exports }) => {
      const bridge = browserBridge(exports);
      if (typeof bridge.Warmup !== 'function') {
        return;
      }

      bridge.Warmup();
    });
    warmupPromises.set(dotnetModuleUrl, warmupPromise);
  }

  return warmupPromise;
}

/**
 * Parse DOCX bytes/blob through the browser WASM bridge and yield arrays of transport-neutral DocxSax events.
 *
 * @param {Uint8Array | ArrayBuffer | ArrayBufferView | Blob} input
 * @param {{ batchSize?: number, dotnetModuleUrl?: string, mainThreadYieldIntervalMs?: number }} [options]
 * @returns {AsyncGenerator<import('./index.d.ts').DocxSaxEvent[], void, void>}
 */
export async function* parseBytesBatches(input, options = {}) {
  const bytes = await toUint8Array(input);
  const batchSize = Number.isInteger(options.batchSize) && options.batchSize > 0 ? options.batchSize : 128;
  const mainThreadYieldIntervalMs = Number.isFinite(options.mainThreadYieldIntervalMs)
    ? Math.max(16, options.mainThreadYieldIntervalMs)
    : 64;
  const { exports } = await loadRuntime(options);
  const bridge = browserBridge(exports);

  if (!supportsPullBatches(bridge)) {
    throw new Error('docx-sax browser WASM runtime does not expose pull batch parsing');
  }

  const parseSessionId = bridge.BeginParseBytesBatches(bytes, batchSize);
  let lastMainThreadYield = typeof performance === 'undefined' ? 0 : performance.now();
  try {
    while (true) {
      const batch = bridge.ReadNextBatch(parseSessionId);
      if (batch == null || batch.length === 0) {
        break;
      }

      yield mapEventBatch(batch);

      if (typeof performance !== 'undefined' && performance.now() - lastMainThreadYield >= mainThreadYieldIntervalMs) {
        await taskYield();
        lastMainThreadYield = performance.now();
      }
    }
  } finally {
    bridge.EndParseBytesBatches(parseSessionId);
  }
}

/**
 * Parse DOCX bytes/blob through the browser WASM bridge and yield transport-neutral DocxSax events.
 *
 * @param {Uint8Array | ArrayBuffer | ArrayBufferView | Blob} input
 * @param {{ batchSize?: number, dotnetModuleUrl?: string }} [options]
 * @returns {AsyncGenerator<import('./index.d.ts').DocxSaxEvent, void, void>}
 */
export async function* parseBytes(input, options = {}) {
  for await (const batch of parseBytesBatches(input, options)) {
    for (const event of batch) {
      yield event;
    }
  }
}

export default {
  parseBytes,
  parseBytesBatches,
  preloadRuntime,
  warmupRuntime,
};
