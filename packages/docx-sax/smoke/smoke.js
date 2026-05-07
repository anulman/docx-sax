import { parseBytesBatches } from '../browser.js';

const response = await fetch('/test/DocxSax.Tests/Fixtures/simple.docx');
if (!response.ok) {
  throw new Error(`fixture fetch failed: ${response.status}`);
}

const bytes = new Uint8Array(await response.arrayBuffer());
const seen = new Set();
let batchCount = 0;
let eventCount = 0;

for await (const batch of parseBytesBatches(bytes, { batchSize: 16 })) {
  batchCount += 1;
  eventCount += batch.length;
  for (const event of batch) {
    seen.add(event.type === 'package' ? `${event.type}:${event.phase}` : event.type);
  }
}

const required = ['package:start', 'part', 'element', 'text', 'end', 'package:end'];
const missing = required.filter((type) => !seen.has(type));
if (missing.length > 0) {
  throw new Error(`missing expected events: ${missing.join(', ')}`);
}

window.__docxSaxSmokeResult = { batchCount, eventCount, seen: [...seen].sort() };
console.log('docx-sax browser smoke ok', window.__docxSaxSmokeResult);
