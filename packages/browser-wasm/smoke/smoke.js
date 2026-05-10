import { parseBytes, parseBytesBatches, preloadRuntime, setRuntimeBaseUrl, warmupRuntime } from '../index.js';
import { assertCommonDocxSaxEventModel, collectEventShapeSummary } from '../../../test/js/common-event-shape.mjs';
import { createGeneratedDocxBytes } from '../../../test/js/generated-docx.mjs';

setRuntimeBaseUrl('/packages/browser-wasm/dist/wasm/wwwroot/_framework');

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

const batchSummary = assertCommonDocxSaxEventModel(batchEvents);

const streamEvents = [];
for await (const event of parseBytes(bytes, { batchSize: 11 })) {
  streamEvents.push(event);
}

const streamSummary = assertCommonDocxSaxEventModel(streamEvents);

if (JSON.stringify(batchSummary.keysByKind) !== JSON.stringify(streamSummary.keysByKind)) {
  throw new Error('parseBytes and parseBytesBatches yielded different event object shapes');
}

const largeBytes = createGeneratedDocxBytes({ paragraphs: 6000 });
const largeBatchSize = 32;
const largeGenerator = parseBytesBatches(largeBytes, { batchSize: largeBatchSize });
const firstBatchStartedAt = performance.now();
const firstBatchResult = await largeGenerator.next();
const firstBatchMs = performance.now() - firstBatchStartedAt;

if (firstBatchResult.done) {
  throw new Error('large generated DOCX unexpectedly completed before yielding its first browser batch');
}

if (firstBatchResult.value.length === 0 || firstBatchResult.value.length > largeBatchSize) {
  throw new Error(`first browser batch size ${firstBatchResult.value.length} exceeded requested batch size ${largeBatchSize}`);
}

const largeEvents = [...firstBatchResult.value];
let largeBatchCount = 1;
const remainingStartedAt = performance.now();
for await (const batch of largeGenerator) {
  largeBatchCount += 1;
  largeEvents.push(...batch);
}
const remainingMs = performance.now() - remainingStartedAt;
const totalLargeMs = firstBatchMs + remainingMs;

if (largeBatchCount < 100) {
  throw new Error(`expected large generated DOCX to require many browser batches, got ${largeBatchCount}`);
}

if (largeEvents.length < 20_000) {
  throw new Error(`expected large generated DOCX to produce many events, got ${largeEvents.length}`);
}

if (firstBatchMs >= totalLargeMs * 0.75) {
  throw new Error(`first browser batch took too much of total parse time: first=${firstBatchMs.toFixed(1)}ms total=${totalLargeMs.toFixed(1)}ms`);
}

const largeSummary = collectEventShapeSummary(largeEvents);

window.__docxSaxSmokeResult = {
  batchCount,
  eventCount: batchEvents.length,
  seen: batchSummary.seen,
  keysByKind: batchSummary.keysByKind,
  streaming: {
    batchSize: largeBatchSize,
    firstBatchEvents: firstBatchResult.value.length,
    largeBatchCount,
    largeEventCount: largeEvents.length,
    firstBatchMs,
    totalLargeMs,
    seen: largeSummary.seen,
  },
};
console.log('docx-sax browser smoke ok', window.__docxSaxSmokeResult);
