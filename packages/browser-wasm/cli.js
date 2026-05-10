#!/usr/bin/env node
import { copyRuntimeAssets } from './lib/assets.js';

function usage() {
  console.error('Usage: npx @docx-sax/browser copy-wasm <dest>');
}

const [, , command, dest] = process.argv;
if (command !== 'copy-wasm' || !dest) {
  usage();
  process.exit(1);
}

try {
  const target = await copyRuntimeAssets(dest);
  console.log(`Copied docx-sax WASM runtime to ${target}`);
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
}
