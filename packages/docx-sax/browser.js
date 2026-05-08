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
  return typeof bridge.BeginParseBytesJsonBatches === 'function'
    && typeof bridge.ReadNextJsonBatch === 'function'
    && typeof bridge.EndParseBytesJsonBatches === 'function';
}

function animationFrame() {
  if (typeof requestAnimationFrame === 'function') {
    return new Promise((resolve) => requestAnimationFrame(() => resolve()));
  }

  return Promise.resolve();
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
 * Warm OpenXML package, XML reader, and DocxSax JSON serialization paths after the WASM runtime loads.
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
 * @param {{ batchSize?: number, dotnetModuleUrl?: string }} [options]
 * @returns {AsyncGenerator<import('./index.d.ts').DocxSaxEvent[], void, void>}
 */
export async function* parseBytesBatches(input, options = {}) {
  const bytes = await toUint8Array(input);
  const batchSize = Number.isInteger(options.batchSize) && options.batchSize > 0 ? options.batchSize : 128;
  const { exports } = await loadRuntime(options);
  const bridge = browserBridge(exports);

  if (!supportsPullBatches(bridge)) {
    const frames = bridge.ParseBytesJsonBatchFrames(bytes, batchSize);
    for (const frame of frames.split('\n')) {
      if (frame.length > 0) {
        yield JSON.parse(frame);
      }
    }
    return;
  }

  const parseSessionId = bridge.BeginParseBytesJsonBatches(bytes, batchSize);
  let lastMainThreadYield = typeof performance === 'undefined' ? 0 : performance.now();
  try {
    while (true) {
      const frame = bridge.ReadNextJsonBatch(parseSessionId);
      if (frame == null || frame.length === 0) {
        break;
      }

      yield JSON.parse(frame);

      if (typeof performance !== 'undefined' && performance.now() - lastMainThreadYield >= 16) {
        await animationFrame();
        lastMainThreadYield = performance.now();
      }
    }
  } finally {
    bridge.EndParseBytesJsonBatches(parseSessionId);
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
