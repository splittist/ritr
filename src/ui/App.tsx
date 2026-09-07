import { useEffect, useState } from 'react';
import type { DesktopApi, Snapshot } from '../desktop/protocol';
import type { ChangePreview, SearchMatch } from '../engine/workspace';
import type { Token } from '../engine/document';
import type { PartDifference } from '../package/docx';
import { StoryView } from './StoryView';
import { validInlineText, type InlineDraft, type InlineEditing } from './inline-edit';

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
  const [inlineDraft, setInlineDraft] = useState<InlineDraft & { revision: number }>();
  const [restore, setRestore] = useState<InlineEditing['restore']>();
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
    if (inlineDraft) setMessage('Inline draft cleared because the workspace changed.');
    setInlineDraft(undefined);
    setRestore(undefined);
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
  const inline: InlineEditing = {
    busy,
    draft: inlineDraft,
    restore,
    message: setError,
    start: (spanId, anchor, head, insert) => {
      if (busy) return;
      if (inlineDraft) {
        setError('Apply or cancel the current inline draft before editing another span.');
        return;
      }
      const token = story?.tokens.find((t) => t.kind === 'text' && t.span.id === spanId);
      if (!document || token?.kind !== 'text' || !token.span.editable) return;
      const text =
        insert === undefined
          ? token.span.text
          : token.span.text.slice(0, anchor) + insert + token.span.text.slice(head);
      if (!validInlineText(text)) {
        setError('This text cannot be edited inline.');
        return;
      }
      setPreview(undefined);
      setRestore(undefined);
      setError('');
      setInlineDraft({
        spanId,
        text,
        revision: document.revision,
        anchor: insert === undefined ? anchor : anchor + insert.length,
        head: insert === undefined ? head : anchor + insert.length,
      });
    },
    change: (text, anchor, head) => {
      if (inlineDraft && !busy) {
        if (text !== inlineDraft.text) setPreview(undefined);
        setInlineDraft({ ...inlineDraft, text, anchor, head });
      }
    },
    cancel: () => {
      if (busy) return;
      if (inlineDraft)
        setRestore({
          spanId: inlineDraft.spanId,
          anchor: inlineDraft.anchor,
          head: inlineDraft.head,
        });
      setInlineDraft(undefined);
      setPreview(undefined);
      setError('');
    },
  };
  useEffect(() => {
    if (!inlineDraft) return;
    const preventClose = (event: BeforeUnloadEvent) => {
      event.preventDefault();
      event.returnValue = '';
    };
    window.addEventListener('beforeunload', preventClose);
    return () => window.removeEventListener('beforeunload', preventClose);
  }, [!!inlineDraft]);
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
          <button className="primary" disabled={!api || busy || !!inlineDraft} onClick={open}>
            Open documents
          </button>
          <button disabled={!document || busy || !!inlineDraft} onClick={save}>
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
                disabled={!!inlineDraft}
                className={`document-link ${document?.id === d.id ? 'active' : ''}`}
                onClick={() => {
                  setDocumentId(d.id);
                  setRestore(undefined);
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
                    disabled={!!inlineDraft}
                    className={`story-link ${story?.id === s.id ? 'active' : ''}`}
                    onClick={() => {
                      setStoryId(s.id);
                      setRestore(undefined);
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
                <span>
                  Type in a span, or double-click to edit inline. Click a code to inspect it.
                </span>
                <span>{document?.dirty ? 'Unsaved changes' : 'Saved snapshot'}</span>
              </div>
              {inlineDraft && (
                <div className="inline-toolbar" aria-label="Inline draft">
                  <span>Inline draft · one text span · not yet applied</span>
                  <button
                    disabled={busy}
                    onClick={() =>
                      void run(async () => {
                        const token = story.tokens.find(
                          (t) => t.kind === 'text' && t.span.id === inlineDraft.spanId,
                        );
                        if (token?.kind !== 'text')
                          throw new Error('The draft span is no longer available.');
                        setPreview(
                          await api!.previewEdit({
                            documentId: document!.id,
                            spanId: inlineDraft.spanId,
                            from: 0,
                            to: token.span.text.length,
                            text: inlineDraft.text,
                            expectedRevision: inlineDraft.revision,
                          }),
                        );
                      })
                    }
                  >
                    Preview inline edit
                  </button>
                  <button disabled={busy} onClick={inline.cancel}>
                    Cancel inline edit
                  </button>
                </div>
              )}
              <StoryView story={story} codes={codes} onSelect={select} inline={inline} />
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
                disabled={!query || !document || busy || !!inlineDraft}
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
                disabled={!query || !document || busy || !!inlineDraft}
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
              <span>Across formatting runs · stops at structure and review codes</span>
            </div>
            {matches && (
              <div className="matches">
                <p>
                  {matches.length} matches · {matches.filter((m) => !m.editable).length} protected
                </p>
                {matches.map((m, i) => (
                  <button
                    key={i}
                    disabled={!!inlineDraft}
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
              <p>Cross-run replacements use the first matched run’s formatting.</p>
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
                    if (inlineDraft)
                      setRestore({
                        spanId: inlineDraft.spanId,
                        anchor: inlineDraft.anchor,
                        head: inlineDraft.head,
                      });
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
                disabled={!selected.span.editable || busy || !!inlineDraft}
                onChange={(e) => setDraft(e.target.value)}
              />
              <button
                className="primary full"
                disabled={
                  !selected.span.editable || busy || !!inlineDraft || draft === selected.span.text
                }
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
              {(['table', 'row', 'cell'] as const).map((kind) =>
                selected.details[kind] ? (
                  <div key={kind}>
                    <h3>
                      {kind === 'table'
                        ? 'Table properties'
                        : kind === 'row'
                          ? 'Row properties'
                          : 'Cell properties'}
                    </h3>
                    <pre>{JSON.stringify(selected.details[kind], null, 2)}</pre>
                  </div>
                ) : null,
              )}
              {selected.paragraph && (
                <>
                  <h3>Paragraph properties</h3>
                  <p className="muted">
                    {selected.paragraph.styleId ? `Style: ${selected.paragraph.styleId}. ` : ''}
                    Indentation shown approximately; source measurements below are in twips (1/20
                    point).
                  </p>
                  <pre>
                    {JSON.stringify(
                      {
                        layout: selected.paragraph.layout,
                        propertySources: selected.paragraph.propertySources,
                      },
                      null,
                      2,
                    )}
                  </pre>
                  {selected.paragraph.numbering && (
                    <>
                      <h3>Generated list label</h3>
                      <pre>
                        {JSON.stringify(
                          { ...selected.paragraph.numbering, definitionXml: undefined },
                          null,
                          2,
                        )}
                      </pre>
                      <h3>Numbering definition</h3>
                      <pre>{selected.paragraph.numbering.definitionXml ?? 'Unresolved'}</pre>
                    </>
                  )}
                  {selected.paragraph.warnings.map((warning, i) => (
                    <p key={i} className="diagnostic warning">
                      {warning}
                    </p>
                  ))}
                </>
              )}
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
