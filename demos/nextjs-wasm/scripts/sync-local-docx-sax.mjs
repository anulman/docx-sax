import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const demoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const repoRoot = path.resolve(demoRoot, '..', '..');
const sourceRoot = path.join(repoRoot, 'packages', 'docx-sax');
const targetRoot = path.join(demoRoot, 'node_modules', 'docx-sax');

const files = [
  'package.json',
  'browser.js',
  'browser.d.ts',
  'node.js',
  'node.d.ts',
];

if (!fs.existsSync(sourceRoot)) {
  throw new Error(`Missing local docx-sax package at ${sourceRoot}`);
}

fs.mkdirSync(targetRoot, { recursive: true });
for (const file of files) {
  const source = path.join(sourceRoot, file);
  const target = path.join(targetRoot, file);
  if (!fs.existsSync(source)) {
    throw new Error(`Missing local docx-sax package file ${source}`);
  }
  fs.copyFileSync(source, target);
}

console.log(`Synced local docx-sax package entrypoints to ${path.relative(demoRoot, targetRoot)}`);
