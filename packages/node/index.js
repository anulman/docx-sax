let nativePluginPromise;

async function loadNativePlugin() {
  nativePluginPromise ??= import('@docx-sax/native-linux-x64').catch((error) => {
    if (error?.code === 'ERR_MODULE_NOT_FOUND' && String(error.message).includes('@docx-sax/native-linux-x64')) {
      throw new Error('@docx-sax/node requires the @docx-sax/native-linux-x64 payload. Install it on linux-x64 or install @docx-sax/browser for the browser/WASM runtime.', { cause: error });
    }
    throw error;
  });

  return nativePluginPromise;
}

/**
 * Parse a DOCX file path through the Native bridge and yield arrays of transport-neutral DocxSax events.
 *
 * @param {string} path
 * @param {{ batchSize?: number, nativeLibraryPath?: string }} [options]
 * @returns {AsyncGenerator<import('./node.d.ts').DocxSaxEvent[], void, void>}
 */
export async function* parseFileBatches(path, options = {}) {
  const plugin = await loadNativePlugin();
  yield* plugin.parseFileBatches(path, options);
}

/**
 * Parse a DOCX file path through the Native bridge and yield transport-neutral DocxSax events.
 *
 * @param {string} path
 * @param {{ batchSize?: number, nativeLibraryPath?: string }} [options]
 * @returns {AsyncGenerator<import('./node.d.ts').DocxSaxEvent, void, void>}
 */
export async function* parseFile(path, options = {}) {
  const plugin = await loadNativePlugin();
  yield* plugin.parseFile(path, options);
}

export default {
  parseFile,
  parseFileBatches,
};
