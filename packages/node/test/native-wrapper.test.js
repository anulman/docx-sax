import assert from 'node:assert/strict';
import { join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';

import { parseFile, parseFileBatches } from '../index.js';

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
