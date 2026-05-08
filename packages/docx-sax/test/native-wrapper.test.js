import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';

import { parseFile, parseFileBatches } from '../node.js';
import { assertCommonDocxSaxEventModel, collectEventShapeSummary } from '../../../test/js/common-event-shape.mjs';
import { createGeneratedDocxBytes } from '../../../test/js/generated-docx.mjs';

const repoRoot = fileURLToPath(new URL('../../..', import.meta.url));
const fixture = join(repoRoot, 'test', 'DocxSax.Tests', 'Fixtures', 'simple.docx');

test('parseFileBatches yields native JSON event batches', async () => {
  const batches = [];
  for await (const batch of parseFileBatches(fixture, { batchSize: 5 })) {
    batches.push(batch);
  }

  assert.ok(batches.length > 1, 'expected multiple native batches with batchSize=5');
  assert.equal(batches[0][0].type, 'package');
  assert.equal(batches[0][0].phase, 'start');
  assert.ok(batches.some((batch) => batch.some((event) => event.type === 'part' && event.phase === 'start')));
});

test('parseFile streams package, part, element, text, and end events from native path', async () => {
  const events = [];
  for await (const event of parseFile(fixture, { batchSize: 4 })) {
    events.push(event);
  }

  assert.ok(events.length > 0);
  assert.deepEqual(events.map((event) => event.ordinal), events.map((_, index) => index));
  assert.equal(events[0].type, 'package');
  assert.equal(events[0].phase, 'start');
  assert.equal(events.at(-1).type, 'package');
  assert.equal(events.at(-1).phase, 'end');

  assert.ok(events.some((event) => event.type === 'part' && event.uri === '/word/document.xml'));
  assert.ok(events.some((event) => event.type === 'element' && event.name === 'w:document'));
  assert.ok(events.some((event) => event.type === 'text' && event.text === 'Hello DOCX SAX'));
  assert.ok(events.some((event) => event.type === 'end' && event.name === 'w:document'));
});

test('parseFile can be stopped early without waiting for the whole native parse', async () => {
  await Promise.race([
    (async () => {
      for await (const event of parseFile(fixture, { batchSize: 1 })) {
        assert.equal(event.type, 'package');
        assert.equal(event.phase, 'start');
        break;
      }
    })(),
    new Promise((_, reject) => {
      const timeout = setTimeout(() => reject(new Error('timed out waiting for early stream disposal')), 1000);
      timeout.unref();
    }),
  ]);
});

test('parseFile yields the shared DocxSaxEvent object model', async () => {
  const events = [];
  for await (const event of parseFile(fixture, { batchSize: 7 })) {
    events.push(event);
  }

  const summary = assertCommonDocxSaxEventModel(events);
  assert.deepEqual(summary, collectEventShapeSummary(events));
  assert.deepEqual(summary.keysByKind['package:start'], ['ordinal', 'phase', 'type']);
  assert.deepEqual(summary.keysByKind.text, ['depth', 'isWhitespace', 'ordinal', 'partUri', 'path', 'text', 'type']);
});

test('parseFileBatches yields bounded batches incrementally for a large generated DOCX after native warmup', async () => {
  // Warm the addon/native runtime first so this measures the parse stream shape, not one-time load cost.
  // Drain the warmup parse completely: cancelling a Native AOT parse mid-callback can exercise
  // shutdown/disposal paths rather than the steady-state streaming behavior this test owns.
  for await (const _batch of parseFileBatches(fixture, { batchSize: 32 })) {
    // drain
  }

  const tempDir = await mkdtemp(join(tmpdir(), 'docx-sax-node-stream-'));
  const generatedFixture = join(tempDir, 'large-generated.docx');

  try {
    await writeFile(generatedFixture, createGeneratedDocxBytes({ paragraphs: 6000 }));

    const batchSize = 32;
    const iterator = parseFileBatches(generatedFixture, { batchSize });
    const firstBatchStartedAt = performance.now();
    const first = await iterator.next();
    const firstBatchMs = performance.now() - firstBatchStartedAt;

    assert.equal(first.done, false, 'expected a first native batch before parse completion');
    assert.ok(first.value.length > 0, 'first native batch should contain events');
    assert.ok(first.value.length <= batchSize, `first native batch should be bounded by requested batch size ${batchSize}`);
    assert.equal(first.value[0].ordinal, 0);

    const events = [...first.value];
    let batchCount = 1;
    const remainingStartedAt = performance.now();
    for await (const batch of iterator) {
      batchCount += 1;
      assert.ok(batch.length <= batchSize, `native batch ${batchCount} should be bounded by requested batch size ${batchSize}`);
      events.push(...batch);
    }
    const remainingMs = performance.now() - remainingStartedAt;
    const totalMs = firstBatchMs + remainingMs;

    assert.ok(batchCount >= 100, `expected large generated DOCX to require many native batches, got ${batchCount}`);
    assert.ok(events.length >= 20_000, `expected large generated DOCX to produce many events, got ${events.length}`);
    assert.ok(
      firstBatchMs < totalMs * 0.75,
      `first native batch should arrive well before complete parse: first=${firstBatchMs.toFixed(1)}ms total=${totalMs.toFixed(1)}ms`,
    );
    assert.deepEqual(events.map((event) => event.ordinal), events.map((_, index) => index));
    assert.equal(events.at(-1).type, 'package');
    assert.equal(events.at(-1).phase, 'end');
  } finally {
    await rm(tempDir, { recursive: true, force: true });
  }
});
