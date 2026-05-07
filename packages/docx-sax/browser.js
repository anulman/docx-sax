const runtimePromises = new Map();

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
    runtimePromise = import(dotnetModuleUrl).then(async ({ dotnet }) => {
      const runtime = await dotnet.withDiagnosticTracing(false).create();
      const config = runtime.getConfig();
      const exports = await runtime.getAssemblyExports(config.mainAssemblyName);
      return { runtime, exports };
    });
    runtimePromises.set(dotnetModuleUrl, runtimePromise);
  }

  return runtimePromise;
}

export async function* parseBytesBatches(input, options = {}) {
  const bytes = await toUint8Array(input);
  const batchSize = Number.isInteger(options.batchSize) && options.batchSize > 0 ? options.batchSize : 128;
  const { exports } = await loadRuntime(options);
  const frames = exports.DocxSax.Browser.BrowserBridge.ParseBytesJsonBatchFrames(bytes, batchSize);

  for (const frame of frames.split('\n')) {
    if (frame.length > 0) {
      yield JSON.parse(frame);
    }
  }
}

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
};
