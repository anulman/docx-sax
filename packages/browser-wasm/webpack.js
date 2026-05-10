import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { listRuntimeFiles, runtimeFrameworkDir } from './lib/assets.js';

function normalizeMount(mount) {
  return String(mount ?? 'docx-sax').replace(/^\/+|\/+$/g, '');
}

export class DocxSaxWasmWebpackPlugin {
  constructor(options = {}) {
    this.mount = normalizeMount(options.mount);
  }

  apply(compiler) {
    const pluginName = 'DocxSaxWasmWebpackPlugin';
    compiler.hooks.thisCompilation.tap(pluginName, (compilation) => {
      const webpack = compiler.webpack;
      const stage = webpack.Compilation.PROCESS_ASSETS_STAGE_ADDITIONS;
      compilation.hooks.processAssets.tapPromise({ name: pluginName, stage }, async () => {
        const sourceRoot = runtimeFrameworkDir();
        const files = await listRuntimeFiles(sourceRoot);
        for (const rel of files) {
          const bytes = await readFile(path.join(sourceRoot, rel));
          const assetName = path.posix.join(this.mount, '_framework', rel.split(path.sep).join('/'));
          compilation.emitAsset(assetName, new webpack.sources.RawSource(bytes));
        }
      });
    });
  }
}

export function docxSaxWasm(options = {}) {
  return new DocxSaxWasmWebpackPlugin(options);
}

export default DocxSaxWasmWebpackPlugin;
