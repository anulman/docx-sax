'use client';

import { useEffect, useMemo, useState } from 'react';
import { parseBytesBatches, preloadRuntime, warmupRuntime } from '@docx-sax/browser';

const DOTNET_MODULE_URL = '/docx-sax/_framework/dotnet.js';
const INITIAL_STATUS = 'Choose a .docx file to parse it in your browser.';
const PREVIEW_UPDATE_INTERVAL_MS = 100;
const MAX_DIAGNOSTICS = 20;
const TRACKED_CHANGE_ELEMENTS = new Set(['ins', 'del', 'moveFrom', 'moveTo']);
const STYLE_REFERENCE_ELEMENTS = new Set(['pStyle', 'rStyle', 'tblStyle', 'numStyleLink', 'styleLink']);
const REFERENCE_ELEMENTS = new Set(['commentRangeStart', 'commentRangeEnd', 'commentReference', 'footnoteReference', 'endnoteReference']);

function emptySummary() {
  return { events: 0, textEvents: 0, parts: new Set(), diagnostics: [] };
}

function cloneSummary(summary) {
  return {
    events: summary.events,
    textEvents: summary.textEvents,
    parts: new Set(summary.parts),
    diagnostics: [...summary.diagnostics],
  };
}

function emptyPreviewDocument() {
  return {
    partSections: [],
    comments: [],
    notes: [],
    trackedChanges: [],
    styleObjects: [],
    references: [],
  };
}

function createPreviewState() {
  return {
    parts: new Map(),
    comments: new Map(),
    notes: new Map(),
    trackedChanges: [],
    styleObjects: [],
    references: [],
    trackedStack: [],
  };
}

function attrValue(event, localName) {
  return event.attributes?.find((attr) => attr.localName === localName)?.value;
}

function attrsObject(event) {
  return Object.fromEntries((event.attributes ?? []).map((attr) => [attr.name, attr.value]));
}

function attrsLabel(attrs) {
  const entries = Object.entries(attrs ?? {});
  if (entries.length === 0) return '';
  return entries.map(([key, value]) => `${key}=${JSON.stringify(value)}`).join(' ');
}

function sectionForPart(preview, partUri) {
  const key = partUri || '(package)';
  let section = preview.parts.get(key);
  if (!section) {
    section = { partUri: key, paragraphs: [], fallbackText: [], currentParagraph: '' };
    preview.parts.set(key, section);
  }
  return section;
}

function appendParagraph(section) {
  const text = section.currentParagraph.trim();
  if (text.length > 0) section.paragraphs.push(text);
  section.currentParagraph = '';
}

function appendText(preview, event) {
  if (event.isWhitespace || !event.text) return;
  const section = sectionForPart(preview, event.partUri);
  section.currentParagraph += event.text;

  const trimmed = event.text.trim();
  if (trimmed) section.fallbackText.push(trimmed);
  for (const change of preview.trackedStack) change.text += event.text;
}

function appendPreviewBatch(preview, batch) {
  for (const event of batch) {
    if (event.type === 'element') {
      if (TRACKED_CHANGE_ELEMENTS.has(event.localName)) {
        const change = {
          type: event.localName,
          partUri: event.partUri,
          attrs: attrsObject(event),
          text: '',
        };
        preview.trackedChanges.push(change);
        preview.trackedStack.push(change);
      }

      if (REFERENCE_ELEMENTS.has(event.localName)) {
        preview.references.push({
          type: event.localName,
          partUri: event.partUri,
          id: attrValue(event, 'id'),
          attrs: attrsObject(event),
        });
      }

      if (event.partUri === '/word/styles.xml' || STYLE_REFERENCE_ELEMENTS.has(event.localName)) {
        preview.styleObjects.push({
          type: event.localName,
          partUri: event.partUri,
          path: event.path,
          attrs: attrsObject(event),
        });
      }
    }

    if (event.type === 'text') {
      appendText(preview, event);
    }

    if (event.type === 'end') {
      if (event.localName === 'p') {
        appendParagraph(sectionForPart(preview, event.partUri));
      }

      if (TRACKED_CHANGE_ELEMENTS.has(event.localName)) {
        const index = preview.trackedStack.findLastIndex((change) => change.type === event.localName);
        if (index !== -1) preview.trackedStack.splice(index, 1);
      }
    }
  }
}

function clonePreviewDocument(preview, includeTrailing = false) {
  const partSections = [...preview.parts.values()].map((section) => {
    const paragraphs = [...section.paragraphs];
    const trailing = section.currentParagraph.trim();
    if (includeTrailing && trailing) paragraphs.push(trailing);

    return {
      partUri: section.partUri,
      paragraphs: paragraphs.length > 0 ? paragraphs : [...section.fallbackText],
    };
  }).filter((section) => section.paragraphs.length > 0);

  return {
    partSections,
    comments: partSections.filter((section) => /comments/i.test(section.partUri)),
    notes: partSections.filter((section) => /(footnotes|endnotes)/i.test(section.partUri)),
    trackedChanges: preview.trackedChanges
      .filter((change) => change.text.trim())
      .map((change) => ({ ...change, text: change.text.trim() })),
    styleObjects: preview.styleObjects.map((style) => ({ ...style, attrs: { ...style.attrs } })),
    references: preview.references.map((ref) => ({ ...ref, attrs: { ...ref.attrs } })),
  };
}

function PartText({ section }) {
  return (
    <section className="doc-section">
      <h3>{section.partUri}</h3>
      {section.paragraphs.map((paragraph, index) => <p key={`${section.partUri}-${index}`}>{paragraph}</p>)}
    </section>
  );
}

function MetadataList({ title, items, renderItem }) {
  if (items.length === 0) return null;
  return (
    <section className="doc-section metadata-section">
      <h3>{title}</h3>
      <ul>
        {items.map((item, index) => <li key={index}>{renderItem(item)}</li>)}
      </ul>
    </section>
  );
}

export default function Home() {
  const [fileName, setFileName] = useState('');
  const [status, setStatus] = useState(INITIAL_STATUS);
  const [busy, setBusy] = useState(false);
  const [previewDocument, setPreviewDocument] = useState(emptyPreviewDocument());
  const [summary, setSummary] = useState(emptySummary());
  const [error, setError] = useState('');

  const partList = useMemo(() => [...summary.parts].sort(), [summary]);

  useEffect(() => {
    let cancelled = false;

    const preload = async () => {
      try {
        await preloadRuntime({ dotnetModuleUrl: DOTNET_MODULE_URL });
        if (!cancelled) {
          setStatus((currentStatus) => currentStatus === INITIAL_STATUS
            ? 'WASM runtime ready. Choose a .docx file to parse it in your browser.'
            : currentStatus);
        }

        await warmupRuntime({ dotnetModuleUrl: DOTNET_MODULE_URL });
        if (!cancelled) {
          setStatus((currentStatus) => currentStatus === 'WASM runtime ready. Choose a .docx file to parse it in your browser.'
            ? 'WASM runtime ready and warmed. Choose a .docx file to parse it in your browser.'
            : currentStatus);
        }
      } catch (err) {
        // Keep the upload path resilient: parseBytesBatches will retry/report the runtime error on demand.
        console.warn('docx-sax WASM preload/warmup failed', err);
      }
    };

    if ('requestIdleCallback' in window) {
      const idleId = window.requestIdleCallback(preload, { timeout: 2_000 });
      return () => {
        cancelled = true;
        window.cancelIdleCallback(idleId);
      };
    }

    const timeoutId = window.setTimeout(preload, 250);
    return () => {
      cancelled = true;
      window.clearTimeout(timeoutId);
    };
  }, []);

  async function parseFile(file) {
    setBusy(true);
    setFileName(file.name);
    setError('');
    setPreviewDocument(emptyPreviewDocument());
    setSummary(emptySummary());
    setStatus('Loading WASM runtime and parsing DOCX…');

    try {
      const preview = createPreviewState();
      const nextSummary = emptySummary();
      let lastPreviewUpdate = 0;

      // parseBytesBatches(file) yields the same DocxSaxEvent object batches as @docx-sax/node
      // parseFileBatches(path); the browser wrapper only differs in accepting bytes/blob input and
      // a dotnetModuleUrl. The WASM bridge is pull-based, so each loop gets a real parsed batch.
      for await (const batch of parseBytesBatches(file, { batchSize: 128, dotnetModuleUrl: DOTNET_MODULE_URL })) {
        appendPreviewBatch(preview, batch);
        for (const event of batch) {
          nextSummary.events += 1;
          if (event.partUri) nextSummary.parts.add(event.partUri);
          if (event.type === 'text' && !event.isWhitespace) nextSummary.textEvents += 1;
          if (event.type === 'diagnostic' && nextSummary.diagnostics.length < MAX_DIAGNOSTICS) {
            nextSummary.diagnostics.push(event.message);
          }
        }

        const now = performance.now();
        if (now - lastPreviewUpdate >= PREVIEW_UPDATE_INTERVAL_MS) {
          lastPreviewUpdate = now;
          setPreviewDocument(clonePreviewDocument(preview));
          setSummary(cloneSummary(nextSummary));
          setStatus(`Parsed ${nextSummary.events} events from ${file.name}…`);
        }
      }

      setPreviewDocument(clonePreviewDocument(preview, true));
      setSummary(cloneSummary(nextSummary));
      setStatus(`Parsed ${nextSummary.events} events from ${file.name}.`);
    } catch (err) {
      console.error(err);
      setError(err instanceof Error ? err.message : String(err));
      setStatus('Parse failed.');
    } finally {
      setBusy(false);
    }
  }

  async function handleChange(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    await parseFile(file);
  }

  const hasPreview = previewDocument.partSections.length > 0
    || previewDocument.trackedChanges.length > 0
    || previewDocument.styleObjects.length > 0
    || previewDocument.references.length > 0;

  return (
    <main className="shell">
      <section className="hero">
        <p className="eyebrow">docx-sax browser/WASM</p>
        <h1>Next.js DOCX event preview</h1>
        <p>
          Upload a Word document and this demo loads <code>@docx-sax/browser</code> client-side,
          parses the DOCX through the WASM bridge, then renders text collected from the shared
          DocxSaxEvent stream used by both browser and Node wrappers.
        </p>
      </section>

      <section className="grid">
        <div className="card">
          <h2>Upload</h2>
          <label className="dropzone">
            <input aria-label="DOCX file" type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" onChange={handleChange} disabled={busy} />
            <span>{busy ? 'Parsing…' : 'Select a .docx file'}</span>
          </label>
          <p className="hint">
            Smoke tests use the public fixture at <a href="/fixtures/simple.docx">/fixtures/simple.docx</a>.
          </p>
          <p role="status" className="status">{status}</p>
          {fileName && <p className="muted">Current file: {fileName}</p>}
          {error && <pre role="alert" className="error">{error}</pre>}
        </div>

        <div className="card preview">
          <h2>Preview</h2>
          {!hasPreview ? (
            <p className="muted">Parsed document text will appear here.</p>
          ) : (
            <article aria-label="Rendered DOCX preview">
              {previewDocument.partSections.map((section) => <PartText key={section.partUri} section={section} />)}
              <MetadataList title="Comments" items={previewDocument.comments} renderItem={(section) => <span>{section.partUri}: {section.paragraphs.join(' ')}</span>} />
              <MetadataList title="Footnotes / endnotes" items={previewDocument.notes} renderItem={(section) => <span>{section.partUri}: {section.paragraphs.join(' ')}</span>} />
              <MetadataList title="Tracked changes" items={previewDocument.trackedChanges} renderItem={(change) => <span><strong>{change.type}</strong> {attrsLabel(change.attrs)} — {change.text}</span>} />
              <MetadataList title="Comment / note references" items={previewDocument.references} renderItem={(ref) => <span><strong>{ref.type}</strong> {ref.partUri} id={ref.id ?? '(none)'}</span>} />
              <MetadataList title="Style objects" items={previewDocument.styleObjects} renderItem={(style) => <span><strong>{style.type}</strong> {style.partUri} {style.path} {attrsLabel(style.attrs)}</span>} />
            </article>
          )}
        </div>
      </section>

      <section className="card details">
        <h2>Event summary</h2>
        <dl>
          <div><dt>Events</dt><dd>{summary.events}</dd></div>
          <div><dt>Text events</dt><dd>{summary.textEvents}</dd></div>
          <div><dt>Parts</dt><dd>{partList.length}</dd></div>
          <div><dt>Tracked changes</dt><dd>{previewDocument.trackedChanges.length}</dd></div>
          <div><dt>Style objects</dt><dd>{previewDocument.styleObjects.length}</dd></div>
        </dl>
        {partList.length > 0 && <ul>{partList.map((part) => <li key={part}>{part}</li>)}</ul>}
        {summary.diagnostics.length > 0 && <pre className="error">{summary.diagnostics.join('\n')}</pre>}
      </section>
    </main>
  );
}
