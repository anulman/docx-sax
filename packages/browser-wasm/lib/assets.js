import { cp, mkdir, readdir, rm, stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const packageRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

export function runtimeFrameworkDir() {
  return path.join(packageRoot, 'dist/wasm/wwwroot/_framework');
}

export function frameworkTargetDir(destination) {
  return path.basename(destination) === '_framework' ? destination : path.join(destination, '_framework');
}

export async function listRuntimeFiles(source = runtimeFrameworkDir(), prefix = '') {
  const entries = await readdir(source, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const rel = path.join(prefix, entry.name);
    const full = path.join(source, entry.name);
    if (entry.isDirectory()) {
      files.push(...await listRuntimeFiles(full, rel));
    } else if (!/\.pdb(\.|$)/i.test(entry.name) && !/\.(br|gz)$/i.test(entry.name)) {
      files.push(rel);
    }
  }
  return files;
}

export async function copyRuntimeAssets(destination, options = {}) {
  const source = options.source ?? runtimeFrameworkDir();
  await stat(path.join(source, 'dotnet.js'));
  const target = frameworkTargetDir(path.resolve(destination));
  if (options.clean !== false) {
    await rm(target, { recursive: true, force: true });
  }
  await mkdir(path.dirname(target), { recursive: true });
  await cp(source, target, {
    recursive: true,
    force: true,
    filter: (entry) => {
      const base = path.basename(entry);
      return !/\.pdb(\.|$)/i.test(base) && !/\.(br|gz)$/i.test(base);
    },
  });
  return target;
}
