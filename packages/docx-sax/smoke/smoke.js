import { parseBytes, parseBytesBatches, preloadRuntime, warmupRuntime } from '../browser.js';

const response = await fetch('/test/DocxSax.Tests/Fixtures/simple.docx');
if (!response.ok) {
  throw new Error(`fixture fetch failed: ${response.status}`);
}

const bytes = new Uint8Array(await response.arrayBuffer());
await preloadRuntime();
await warmupRuntime();
// A second call should reuse the same warmup promise and remain safe before parsing real bytes.
await warmupRuntime();

const batchEvents = [];
let batchCount = 0;

for await (const batch of parseBytesBatches(bytes, { batchSize: 16 })) {
  batchCount += 1;
  batchEvents.push(...batch);
}

const streamEvents = [];
for await (const event of parseBytes(bytes, { batchSize: 11 })) {
  streamEvents.push(event);
}

const seen = new Set(batchEvents.map((event) => event.type));
for (const type of ['package', 'part', 'element', 'text', 'end']) {
  if (!seen.has(type)) {
    throw new Error(`missing ${type} event`);
  }
}

if (batchCount < 2) {
  throw new Error(`expected multiple batches, got ${batchCount}`);
}

if (streamEvents.length !== batchEvents.length) {
  throw new Error(`parseBytes yielded ${streamEvents.length} events; parseBytesBatches yielded ${batchEvents.length}`);
}

window.__docxSaxSmokeResult = { batchCount, eventCount: batchEvents.length, seen: [...seen].sort() };
console.log('docx-sax browser smoke ok', window.__docxSaxSmokeResult);
