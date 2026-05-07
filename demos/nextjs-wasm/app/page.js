'use client';

import { useMemo, useState } from 'react';
import { parseBytes } from 'docx-sax/browser';

const DOTNET_MODULE_URL = '/docx-sax/_framework/dotnet.js';

function emptySummary() {
  return { events: 0, textEvents: 0, parts: new Set(), diagnostics: [] };
}

function renderPreview(events) {
  const paragraphs = [];
  let current = '';

  for (const event of events) {
    if (event.type === 'text' && !event.isWhitespace && event.text) {
      current += event.text;
    }

    if (event.type === 'end' && event.localName === 'p') {
      const text = current.trim();
      if (text.length > 0) {
        paragraphs.push(text);
      }
      current = '';
    }
  }

  const trailing = current.trim();
  if (trailing.length > 0) {
    paragraphs.push(trailing);
  }

  return paragraphs;
}

export default function Home() {
  const [fileName, setFileName] = useState('');
  const [status, setStatus] = useState('Choose a .docx file to parse it in your browser.');
  const [busy, setBusy] = useState(false);
  const [paragraphs, setParagraphs] = useState([]);
  const [summary, setSummary] = useState(emptySummary());
  const [error, setError] = useState('');

  const partList = useMemo(() => [...summary.parts].sort(), [summary]);

  async function parseFile(file) {
    setBusy(true);
    setFileName(file.name);
    setError('');
    setParagraphs([]);
    setSummary(emptySummary());
    setStatus('Loading WASM runtime and parsing DOCX…');

    try {
      const events = [];
      const nextSummary = emptySummary();

      for await (const event of parseBytes(file, { batchSize: 128, dotnetModuleUrl: DOTNET_MODULE_URL })) {
        events.push(event);
        nextSummary.events += 1;
        if (event.partUri) nextSummary.parts.add(event.partUri);
        if (event.type === 'text' && !event.isWhitespace) nextSummary.textEvents += 1;
        if (event.type === 'diagnostic') nextSummary.diagnostics.push(event.message);
      }

      const nextParagraphs = renderPreview(events);
      setParagraphs(nextParagraphs);
      setSummary(nextSummary);
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
          parses the DOCX through the WASM bridge, then renders text collected from the low-level
          event stream.
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
