'use client';

import { useEffect, useMemo, useState } from 'react';
import { parseBytesBatches, preloadRuntime, warmupRuntime } from 'docx-sax/browser';

const DOTNET_MODULE_URL = '/docx-sax/_framework/dotnet.js';
const INITIAL_STATUS = 'Choose a .docx file to parse it in your browser.';
const PREVIEW_UPDATE_INTERVAL_MS = 100;
const MAX_PREVIEW_PARAGRAPHS = 80;
const MAX_FALLBACK_SNIPPETS = 40;
const MAX_DIAGNOSTICS = 20;

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

function createPreviewState() {
  return {
    paragraphs: [],
    fallbackSnippets: [],
    fallbackSeen: new Set(),
    currentParagraph: '',
  };
}

function appendPreviewBatch(preview, batch) {
  for (const event of batch) {
    if (event.type === 'text' && !event.isWhitespace && event.text) {
      preview.currentParagraph += event.text;
      const snippet = event.text.trim();
      if (snippet && preview.fallbackSnippets.length < MAX_FALLBACK_SNIPPETS && !preview.fallbackSeen.has(snippet)) {
        preview.fallbackSeen.add(snippet);
        preview.fallbackSnippets.push(snippet);
      }
    }

    if (event.type === 'end' && event.localName === 'p') {
      const text = preview.currentParagraph.trim();
      if (text.length > 0 && preview.paragraphs.length < MAX_PREVIEW_PARAGRAPHS) {
        preview.paragraphs.push(text);
      }
      preview.currentParagraph = '';
    }
  }
}

function previewParagraphs(preview, includeTrailing = false) {
  const paragraphs = [...preview.paragraphs];
  const trailing = preview.currentParagraph.trim();
  if (includeTrailing && trailing.length > 0 && paragraphs.length > 0 && paragraphs.length < MAX_PREVIEW_PARAGRAPHS) {
    paragraphs.push(trailing);
  }

  if (paragraphs.length > 0) {
    return paragraphs;
  }

  // Some valid DOCX packages carry useful text outside Word paragraphs, for example chart
  // caches under c:v/c:f. Surface those text events instead of leaving Preview blank.
  return preview.fallbackSnippets;
}

export default function Home() {
  const [fileName, setFileName] = useState('');
  const [status, setStatus] = useState(INITIAL_STATUS);
  const [busy, setBusy] = useState(false);
  const [paragraphs, setParagraphs] = useState([]);
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
    setParagraphs([]);
    setSummary(emptySummary());
    setStatus('Loading WASM runtime and parsing DOCX…');

    try {
      const preview = createPreviewState();
      const nextSummary = emptySummary();
      let lastPreviewUpdate = 0;

      // parseBytesBatches(file) yields the same DocxSaxEvent object batches as docx-sax/node
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
          setParagraphs(previewParagraphs(preview));
          setSummary(cloneSummary(nextSummary));
          setStatus(`Parsed ${nextSummary.events} events from ${file.name}…`);
        }
      }

      setParagraphs(previewParagraphs(preview, true));
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

  return (
    <main className="shell">
      <section className="hero">
        <p className="eyebrow">docx-sax browser/WASM</p>
        <h1>Next.js DOCX event preview</h1>
        <p>
          Upload a Word document and this demo loads <code>docx-sax/browser</code> client-side,
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
          {paragraphs.length === 0 ? (
            <p className="muted">Parsed document text will appear here.</p>
          ) : (
            <article aria-label="Rendered DOCX preview">
              {paragraphs.map((paragraph, index) => <p key={index}>{paragraph}</p>)}
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
        </dl>
        {partList.length > 0 && <ul>{partList.map((part) => <li key={part}>{part}</li>)}</ul>}
        {summary.diagnostics.length > 0 && <pre className="error">{summary.diagnostics.join('\n')}</pre>}
      </section>
    </main>
  );
}
