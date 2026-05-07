import { createRequire } from 'node:module';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const nativeAddon = require('./build/Release/docx_sax_node.node');

const packageDir = dirname(fileURLToPath(import.meta.url));

function defaultNativeLibraryPath() {
  if (process.platform !== 'linux' || process.arch !== 'x64') {
    throw new Error(`@docx-sax/node v0 currently ships a linux-x64 native bridge only (got ${process.platform}-${process.arch})`);
  }

  return join(packageDir, 'native', 'linux-x64', 'DocxSax.Native.so');
}

function normalizeOptions(options = {}) {
  const batchSize = Number.isInteger(options.batchSize) && options.batchSize > 0 ? options.batchSize : 128;
  const nativeLibraryPath = options.nativeLibraryPath === undefined
    ? defaultNativeLibraryPath()
    : resolve(String(options.nativeLibraryPath));

  return { batchSize, nativeLibraryPath };
}

export async function* parseFileBatches(path, options = {}) {
  if (typeof path !== 'string' || path.length === 0) {
    throw new TypeError('parseFileBatches(path) requires a non-empty string path');
  }

  const { batchSize, nativeLibraryPath } = normalizeOptions(options);
  const jsonBatches = await nativeAddon.parseFileBatchesJson(resolve(path), batchSize, nativeLibraryPath);

  for (const jsonBatch of jsonBatches) {
    yield JSON.parse(jsonBatch);
  }
}

export async function* parseFile(path, options = {}) {
  for await (const batch of parseFileBatches(path, options)) {
    for (const event of batch) {
      yield event;
    }
  }
}

export default {
  parseFile,
  parseFileBatches,
};
