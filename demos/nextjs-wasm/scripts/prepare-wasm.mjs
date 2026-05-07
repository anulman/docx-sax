import { cp, mkdir, stat } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const demoRoot = path.resolve(here, '..');
const repoRoot = path.resolve(demoRoot, '../..');
const source = path.join(repoRoot, 'packages/docx-sax/dist/wasm/wwwroot/_framework');
const target = path.join(demoRoot, 'public/docx-sax/_framework');

try {
  await stat(path.join(source, 'dotnet.js'));
} catch {
  throw new Error(`Missing browser WASM assets at ${source}. Run \`npm run build:wasm\` from demos/nextjs-wasm or \`npm --prefix ../../packages/docx-sax run build:dotnet\` first.`);
}

await mkdir(path.dirname(target), { recursive: true });
await cp(source, target, { recursive: true, force: true });
console.log(`Copied docx-sax WASM assets to ${path.relative(demoRoot, target)}`);
