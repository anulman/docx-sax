import { cp, mkdir, rm, stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const demoRoot = path.resolve(here, '..');
const repoRoot = path.resolve(demoRoot, '../..');
const source = path.join(repoRoot, 'packages/browser-wasm/dist/wasm/wwwroot/_framework');
const target = path.join(demoRoot, 'public/docx-sax/_framework');

try {
  await stat(path.join(source, 'dotnet.js'));
} catch {
  throw new Error(`Missing browser WASM assets at ${source}. Run \`npm run build:wasm\` from demos/nextjs-wasm or \`npm --prefix ../../packages/browser-wasm run build:dotnet\` first.`);
}

await rm(target, { recursive: true, force: true });
await mkdir(path.dirname(target), { recursive: true });
await cp(source, target, {
  recursive: true,
  force: true,
  filter: (entry) => !/\.pdb(\.|$)/i.test(path.basename(entry)),
});
console.log(`Copied docx-sax WASM assets to ${path.relative(demoRoot, target)}`);
