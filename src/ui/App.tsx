import { useEffect, useState } from 'react';
import type { DesktopApi, Snapshot } from '../desktop/protocol';
import type { ChangePreview, SearchMatch } from '../engine/workspace';
import type { Token } from '../engine/document';
import type { PartDifference } from '../package/docx';
import { RevealEditor } from './RevealEditor';

declare global {
  interface Window {
    ritr?: DesktopApi;
  }
}
const empty: Snapshot = { documents: [], canUndo: false, canRedo: false };

export function App() {
  const [state, setState] = useState(empty);
  const [documentId, setDocumentId] = useState('');
  const [storyId, setStoryId] = useState('');
  const [selected, setSelected] = useState<Token>();
  const [draft, setDraft] = useState('');
  const [query, setQuery] = useState('');
  const [replacement, setReplacement] = useState('');
  const [caseSensitive, setCaseSensitive] = useState(true);
  const [matches, setMatches] = useState<SearchMatch[]>();
  const [preview, setPreview] = useState<ChangePreview>();
  const [report, setReport] = useState<PartDifference[]>();
  const [codes, setCodes] = useState(true);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('Open a DOCX to explore its text and structure.');
  const [error, setError] = useState('');
  const document = state.documents.find((d) => d.id === documentId) ?? state.documents[0];
  const story = document?.model.stories.find((s) => s.id === storyId) ?? document?.model.stories[0];
  const api = window.ritr;
  async function run(action: () => Promise<void>) {
    setBusy(true);
    setError('');
    try {
      await action();
    } catch (e) {
      setError(String(e).replace(/^Error: /, ''));
    } finally {
      setBusy(false);
    }
  }
  function refresh(snapshot: Snapshot) {
    setState(snapshot);
    setPreview(undefined);
    setMatches(undefined);
    setSelected(undefined);
    setReport(undefined);
  }
  function select(token: Token) {
    setSelected(token);
    setDraft(token.kind === 'text' ? token.span.text : '');
  }
  useEffect(() => {
    if (api)
      void run(async () => {
        setState(await api.snapshot());
      });
  }, []);
  const open = () =>
    api &&
    run(async () => {
      refresh(await api.open());
      setMessage('Documents opened. Select text or a code to inspect it.');
    });
  const save = () =>
    api &&
    document &&
    run(async () => {
      const result = await api.save(document.id);
      refresh(result.snapshot);
      setReport(result.report);
      if (result.path) setMessage(`Saved and verified ${result.path}`);
    });
  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          ritr<span>DOCUMENT WORKBENCH</span>
        </div>
        <div className="toolbar">
          <button className="primary" disabled={!api || busy} onClick={open}>
            Open documents
          </button>
          <button disabled={!document || busy} onClick={save}>
            Save As…
          </button>
          <span className="divider" />
          <button
            disabled={!state.canUndo || busy}
            onClick={() => void run(async () => refresh(await api!.undo()))}
          >
            Undo
          </button>
          <button
            disabled={!state.canRedo || busy}
            onClick={() => void run(async () => refresh(await api!.redo()))}
          >
            Redo
          </button>
        </div>
        <span className="local-badge">● Local workspace</span>
      </header>
      {!api && (
        <div className="notice">
          Desktop bridge unavailable. Run <code>npm start</code> to open local documents.
        </div>
      )}
      <div className="workspace">
        <aside className="sidebar">
          <div className="section-heading">
            WORKSPACE <span>{state.documents.length}</span>
          </div>
          {!state.documents.length && (
            <p className="muted">Your open documents will appear here.</p>
          )}
          <nav aria-label="Documents">
            {state.documents.map((d) => (
              <button
                key={d.id}
                className={`document-link ${document?.id === d.id ? 'active' : ''}`}
                onClick={() => {
                  setDocumentId(d.id);
                  setStoryId('');
                  setSelected(undefined);
                  setReport(undefined);
                }}
              >
                <span>▤</span>
                <span>
                  {d.name}
                  {d.dirty ? ' •' : ''}
                  <small>{d.model.stories.length} stories</small>
                </span>
              </button>
            ))}
          </nav>
          {document && (
            <>
              <div className="section-heading stories-heading">STORIES</div>
              <nav aria-label="Stories">
                {document.model.stories.map((s) => (
                  <button
                    key={s.id}
                    className={`story-link ${story?.id === s.id ? 'active' : ''}`}
                    onClick={() => {
                      setStoryId(s.id);
                      setSelected(undefined);
                    }}
                  >
                    {s.label}
                    <small>{s.tokens.filter((t) => t.kind === 'text').length} text spans</small>
                  </button>
                ))}
              </nav>
            </>
          )}
          <div className="sidebar-foot">
            Original files stay untouched.
            <br />
            Changes live here until Save As.
          </div>
        </aside>
        <main>
          <section className="document-heading">
            <div>
              <div className="eyebrow">{story?.label ?? 'GET STARTED'}</div>
              <h1>{document?.name ?? 'See what your document is made of.'}</h1>
            </div>
            {document && (
              <label className="toggle">
                <input
                  type="checkbox"
                  checked={codes}
                  onChange={(e) => setCodes(e.target.checked)}
                />{' '}
                Reveal codes
              </label>
            )}
          </section>
          {story ? (
            <>
              <div className="editor-caption">
                <span>Select text to edit. Click a code to inspect its source.</span>
                <span>{document?.dirty ? 'Unsaved changes' : 'Saved snapshot'}</span>
              </div>
              <RevealEditor story={story} codes={codes} onSelect={select} />
            </>
          ) : (
            <section className="welcome">
              <div className="sample-code">Paragraph</div>
              <p>Text, with its structure in view.</p>
              <div className="sample-code">Run · bold</div>
              <p className="welcome-description">
                Open one or several Word documents. Inspect formatting, explore stories, and preview
                changes across your workspace.
              </p>
              <button className="primary" disabled={!api || busy} onClick={open}>
                Open your first document
              </button>
              <div className="welcome-notes">
                01 &nbsp; Explore the codes
                <br />
                02 &nbsp; Edit a text span
                <br />
                03 &nbsp; Save a verified copy
              </div>
            </section>
          )}
          <section className="search-panel">
            <div className="section-heading">WORKSPACE SEARCH & REPLACE</div>
            <div className="search-fields">
              <input
                aria-label="Find text"
                placeholder="Find literal text…"
                value={query}
                onChange={(e) => {
                  setQuery(e.target.value);
                  setMatches(undefined);
                  setPreview(undefined);
                }}
              />
              <input
                aria-label="Replace with"
                placeholder="Replace with…"
                value={replacement}
                onChange={(e) => {
                  setReplacement(e.target.value);
                  setPreview(undefined);
                }}
              />
              <button
                disabled={!query || !document || busy}
                onClick={() =>
                  void run(async () => {
                    setMatches(await api!.search(query, caseSensitive));
                    setPreview(undefined);
                  })
                }
              >
                Find
              </button>
              <button
                disabled={!query || !document || busy}
                onClick={() =>
                  void run(async () =>
                    setPreview(await api!.previewReplace(query, replacement, caseSensitive)),
                  )
                }
              >
                Preview replacement
              </button>
            </div>
            <div className="search-options">
              <label>
                <input
                  type="checkbox"
                  checked={caseSensitive}
                  onChange={(e) => {
                    setCaseSensitive(e.target.checked);
                    setPreview(undefined);
                    setMatches(undefined);
                  }}
                />{' '}
                Match case
              </label>
              <span>Within text spans · all open documents and stories</span>
            </div>
            {matches && (
              <div className="matches">
                <p>
                  {matches.length} matches · {matches.filter((m) => !m.editable).length} protected
                </p>
                {matches.map((m, i) => (
                  <button
                    key={i}
                    onClick={() => {
                      setDocumentId(m.documentId);
                      setStoryId(m.storyId);
                      const token = state.documents
                        .find((d) => d.id === m.documentId)
                        ?.model.stories.find((s) => s.id === m.storyId)
                        ?.tokens.find((t) => t.kind === 'text' && t.span.id === m.spanId);
                      if (token) select(token);
                    }}
                  >
                    {state.documents.find((d) => d.id === m.documentId)?.name} · {m.before}{' '}
                    {m.editable ? '' : '🔒'}
                  </button>
                ))}
              </div>
            )}
          </section>
          {preview && (
            <section className="preview-panel" aria-label="Change preview">
              <div className="section-heading">
                CHANGE PREVIEW{' '}
                <span>
                  {preview.edits.length} spans · {preview.skipped} protected matches skipped
                </span>
              </div>
              {preview.edits.map((e, i) => (
                <div className="diff" key={i}>
                  <small>{e.name}</small>
                  <del>{e.before || '(empty)'}</del>
                  <ins>{e.after || '(empty)'}</ins>
                </div>
              ))}
              {!preview.edits.length && <p>No editable changes.</p>}
              <button
                className="primary"
                disabled={!preview.edits.length || busy}
                onClick={() =>
                  void run(async () => {
                    refresh(await api!.commit(preview.id));
                    setMessage('Transaction applied. Undo reverses the complete change.');
                  })
                }
              >
                Apply transaction
              </button>
              <button onClick={() => setPreview(undefined)}>Dismiss</button>
            </section>
          )}
        </main>
        <aside className="inspector">
          <div className="section-heading">INSPECTOR</div>
          {!selected && (
            <p className="muted">
              Select a text span or a formatting code to see its properties and source binding.
            </p>
          )}
          {selected?.kind === 'text' && (
            <section>
              <h2>Text span</h2>
              <span className={`status ${selected.span.editable ? '' : 'protected'}`}>
                {selected.span.editable ? 'Editable' : 'Preserved · read-only'}
              </span>
              {selected.span.reason && <p>{selected.span.reason}</p>}
              <label className="field-label" htmlFor="span-text">
                Text content
              </label>
              <textarea
                id="span-text"
                value={draft}
                disabled={!selected.span.editable || busy}
                onChange={(e) => setDraft(e.target.value)}
              />
              <button
                className="primary full"
                disabled={!selected.span.editable || busy || draft === selected.span.text}
                onClick={() =>
                  void run(async () =>
                    setPreview(
                      await api!.previewEdit({
                        documentId: document!.id,
                        spanId: selected.span.id,
                        from: 0,
                        to: selected.span.text.length,
                        text: draft,
                      }),
                    ),
                  )
                }
              >
                Preview text edit
              </button>
              <h3>Direct formatting</h3>
              <pre>{JSON.stringify(selected.span.direct, null, 2)}</pre>
              <h3>Inherited formatting</h3>
              <pre>{JSON.stringify(selected.span.inherited, null, 2)}</pre>
              <h3>Source binding</h3>
              <pre>{JSON.stringify(selected.span.source, null, 2)}</pre>
            </section>
          )}
          {selected?.kind === 'code' && (
            <section>
              <h2>{selected.label}</h2>
              <span className="status">{selected.category}</span>
              <h3>Source binding</h3>
              <pre>{JSON.stringify(selected.source, null, 2)}</pre>
              <h3>Source XML</h3>
              <pre>{String(selected.details.xml)}</pre>
              <p className="muted">
                Structure and formatting codes are inspectable in this version. Their properties are
                preserved during text edits.
              </p>
            </section>
          )}
          {document && (
            <section className="diagnostics">
              <div className="section-heading">PRESERVATION</div>
              <button
                className="full"
                disabled={busy}
                onClick={() => void run(async () => setReport(await api!.report(document.id)))}
              >
                Compare package parts
              </button>
              {report && (
                <div className="part-report">
                  <p>
                    {report.filter((p) => p.status === 'unchanged').length} unchanged ·{' '}
                    {report.filter((p) => p.status !== 'unchanged').length} changed
                  </p>
                  {report.map((p) => (
                    <div key={p.part} className={p.status !== 'unchanged' ? 'changed' : ''}>
                      <span>{p.part}</span>
                      <small>{p.status}</small>
                    </div>
                  ))}
                </div>
              )}
              <h3>Diagnostics · {document.model.diagnostics.length}</h3>
              {document.model.diagnostics.length ? (
                document.model.diagnostics.map((d, i) => (
                  <p className={`diagnostic ${d.severity}`} key={i}>
                    {d.message}
                  </p>
                ))
              ) : (
                <p className="muted">No package diagnostics.</p>
              )}
            </section>
          )}
        </aside>
      </div>
      <footer role="status" className={error ? 'error' : ''}>
        {busy ? 'Working…' : error || message}
        <span>ritr / 0.1 · text editing prototype</span>
      </footer>
    </div>
  );
}
