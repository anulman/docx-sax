import { createReadStream } from 'node:fs';
import path from 'node:path';
import { copyRuntimeAssets, runtimeFrameworkDir } from './lib/assets.js';

function normalizeMount(mount) {
  const value = mount ?? '/docx-sax';
  return `/${String(value).replace(/^\/+|\/+$/g, '')}`;
}

function publicPathPrefix(base, mount) {
  const basePart = String(base ?? '/').replace(/^\/+|\/+$/g, '');
  const mountPart = String(mount).replace(/^\/+|\/+$/g, '');
  return `/${[basePart, mountPart, '_framework'].filter(Boolean).join('/')}/`;
}

function contentType(file) {
  if (file.endsWith('.js')) return 'text/javascript; charset=utf-8';
  if (file.endsWith('.json')) return 'application/json; charset=utf-8';
  if (file.endsWith('.wasm')) return 'application/wasm';
  return 'application/octet-stream';
}

export function docxSaxWasm(options = {}) {
  const mount = normalizeMount(options.mount);
  let resolvedConfig;
  return {
    name: '@docx-sax/browser-wasm',
    configResolved(config) {
      resolvedConfig = config;
    },
    configureServer(server) {
      const prefix = publicPathPrefix(resolvedConfig?.base, mount);
      server.middlewares.use((req, res, next) => {
        const url = new URL(req.url ?? '/', 'http://localhost');
        if (!url.pathname.startsWith(prefix)) return next();
        const rel = decodeURIComponent(url.pathname.slice(prefix.length));
        const sourceRoot = runtimeFrameworkDir();
        const file = path.resolve(sourceRoot, rel);
        if (file !== sourceRoot && !file.startsWith(`${sourceRoot}${path.sep}`)) return next();
        res.setHeader('Content-Type', contentType(file));
        createReadStream(file).on('error', next).pipe(res);
      });
    },
    async closeBundle() {
      if (!resolvedConfig || resolvedConfig.command !== 'build') return;
      const outDir = path.resolve(resolvedConfig.root, resolvedConfig.build.outDir);
      const destination = path.join(outDir, mount);
      await copyRuntimeAssets(destination, { clean: options.clean !== false });
    },
  };
}

export default docxSaxWasm;
