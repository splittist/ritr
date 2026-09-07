import { useEffect, useRef, useState } from 'react';
import type { DesktopApi, Snapshot } from '../desktop/protocol';
import type { ChangePreview, SearchMatch } from '../engine/workspace';
import type { Token } from '../engine/document';
import type { PartDifference } from '../package/docx';
import { StoryView } from './StoryView';
import { validInlineText, type InlineDraft, type InlineEditing } from './inline-edit';
import { CommandPalette, captureCommandFocus } from './CommandPalette';
import {
  commands,
  commandReason,
  invokeCommand,
  shortcutCommand,
  type CommandActions,
  type CommandContext,
  type CommandId,
} from './commands';

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
  const [paletteOpen, setPaletteOpen] = useState(false);
  const paletteFocus = useRef<() => void>(() => {});
  const running = useRef(false);
  const searchInput = useRef<HTMLInputElement>(null);
  const document = state.documents.find((d) => d.id === documentId) ?? state.documents[0];
  const story = document?.model.stories.find((s) => s.id === storyId) ?? document?.model.stories[0];
  const api = window.ritr;
  async function run(action: () => Promise<void>) {
    if (running.current) return;
    running.current = true;
    setBusy(true);
    setError('');
    try {
      await action();
    } catch (e) {
      setError(String(e).replace(/^Error: /, ''));
    } finally {
      running.current = false;
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
  function select(token?: Token) {
    setSelected(token);
    setDraft(token?.kind === 'text' ? token.span.text : '');
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
    cancel: () => execute('edit.cancelInline'),
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
  const selectionReason =
    selected?.kind !== 'text'
      ? 'Select an editable text span.'
      : !selected.span.editable
        ? (selected.span.reason ?? 'This text is protected.')
        : undefined;
  const inlineToken = story?.tokens.find(
    (t) => t.kind === 'text' && t.span.id === inlineDraft?.spanId,
  );
  const context: CommandContext = {
    connected: !!api,
    busy,
    hasDocument: !!document,
    canUndo: state.canUndo,
    canRedo: state.canRedo,
    hasInlineDraft: !!inlineDraft,
    selectionReason,
    editReason: inlineDraft
      ? inlineToken?.kind !== 'text'
        ? 'The draft span is no longer available.'
        : inlineDraft.text === inlineToken.span.text
          ? 'The draft has no changes.'
          : undefined
      : (selectionReason ??
        (selected?.kind === 'text' && draft === selected.span.text
          ? 'Edit the selected span first.'
          : undefined)),
    hasQuery: !!query,
    hasPreview: !!preview,
    hasChanges: !!preview?.edits.length,
  };
  const actions: CommandActions = {
    'commands.open': () => {
      if (!paletteOpen) paletteFocus.current = captureCommandFocus();
      setPaletteOpen((open) => !open);
    },
    'file.open': async () => {
      refresh(await api!.open());
      setMessage('Documents opened. Select text or a code to inspect it.');
    },
    'file.saveAs': async () => {
      const result = await api!.save(document!.id);
      refresh(result.snapshot);
      setReport(result.report);
      if (result.path) setMessage(`Saved and verified ${result.path}`);
    },
    'history.undo': async () => refresh(await api!.undo()),
    'history.redo': async () => refresh(await api!.redo()),
    'edit.inline': () => {
      if (selected?.kind === 'text') inline.start(selected.span.id, 0, selected.span.text.length);
    },
    'edit.preview': async () => {
      if (inlineDraft && inlineToken?.kind === 'text') {
        setPreview(
          await api!.previewEdit({
            documentId: document!.id,
            spanId: inlineDraft.spanId,
            from: 0,
            to: inlineToken.span.text.length,
            text: inlineDraft.text,
            expectedRevision: inlineDraft.revision,
          }),
        );
      } else if (selected?.kind === 'text') {
        setPreview(
          await api!.previewEdit({
            documentId: document!.id,
            spanId: selected.span.id,
            from: 0,
            to: selected.span.text.length,
            text: draft,
            expectedRevision: document!.revision,
          }),
        );
      }
    },
    'edit.cancelInline': () => {
      if (inlineDraft)
        setRestore({
          spanId: inlineDraft.spanId,
          anchor: inlineDraft.anchor,
          head: inlineDraft.head,
        });
      setInlineDraft(undefined);
      setPreview(undefined);
    },
    'search.focus': () => {
      searchInput.current?.focus();
      searchInput.current?.select();
    },
    'search.find': async () => {
      setMatches(await api!.search(query, caseSensitive));
      setPreview(undefined);
    },
    'search.replace': async () =>
      setPreview(await api!.previewReplace(query, replacement, caseSensitive)),
    'transaction.apply': async () => {
      refresh(await api!.commit(preview!.id));
      if (inlineDraft)
        setRestore({
          spanId: inlineDraft.spanId,
          anchor: inlineDraft.anchor,
          head: inlineDraft.head,
        });
      setMessage('Transaction applied. Undo reverses the complete change.');
    },
    'transaction.dismiss': () => setPreview(undefined),
    'view.codes': () => setCodes((shown) => !shown),
    'document.compare': async () => setReport(await api!.report(document!.id)),
  };
  const currentCommands = useRef({ context, actions, paletteOpen });
  currentCommands.current = { context, actions, paletteOpen };
  function execute(id: CommandId) {
    void invokeCommand(
      id,
      () => ({
        ...currentCommands.current.context,
        busy: currentCommands.current.context.busy || running.current,
      }),
      () =>
        id === 'commands.open'
          ? currentCommands.current.actions[id]()
          : run(async () => {
              await currentCommands.current.actions[id]();
            }),
    ).then((reason) => {
      if (reason) setError(reason);
    });
  }
  function commandProps(id: CommandId) {
    const command = commands.find((c) => c.id === id)!;
    const reason = commandReason(id, context);
    return {
      children: command.label,
      disabled: !!reason,
      title:
        reason ??
        `${command.description}${'shortcuts' in command ? ` (${command.shortcuts.join(' / ')})` : ''}`,
      'aria-keyshortcuts':
        'shortcuts' in command
          ? command.shortcuts.map((key) => key.replace('Ctrl', 'Control')).join(' ')
          : undefined,
      onClick: () => execute(id),
    };
  }
  useEffect(() => {
    const keydown = (event: KeyboardEvent) => {
      const target = event.target instanceof HTMLElement ? event.target : undefined;
      const nativeText =
        !!target?.closest('input, textarea') ||
        (!!target?.isContentEditable && !target.closest('.cm-content'));
      const id = shortcutCommand(event, nativeText);
      if (!id) return;
      event.preventDefault();
      event.stopPropagation();
      if (!currentCommands.current.paletteOpen || id === 'commands.open') execute(id);
    };
    window.addEventListener('keydown', keydown, true);
    return () => window.removeEventListener('keydown', keydown, true);
  }, []);
  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          ritr<span>DOCUMENT WORKBENCH</span>
        </div>
        <div className="toolbar">
          <button className="primary" {...commandProps('file.open')} />
          <button {...commandProps('file.saveAs')} />
          <span className="divider" />
          <button {...commandProps('history.undo')} />
          <button {...commandProps('history.redo')} />
          <button {...commandProps('commands.open')}>Commands</button>
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
                  disabled={!!commandReason('view.codes', context)}
                  title={commandReason('view.codes', context)}
                  onChange={() => execute('view.codes')}
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
                  <button {...commandProps('edit.preview')}>Preview inline edit</button>
                  <button {...commandProps('edit.cancelInline')} />
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
              <button className="primary" {...commandProps('file.open')}>
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
                ref={searchInput}
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
              <button {...commandProps('search.find')} />
              <button {...commandProps('search.replace')} />
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
              <button className="primary" {...commandProps('transaction.apply')} />
              <button {...commandProps('transaction.dismiss')}>Dismiss</button>
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
              <button className="full" {...commandProps('edit.inline')}>
                Edit inline
              </button>
              <label className="field-label" htmlFor="span-text">
                Text content
              </label>
              <textarea
                id="span-text"
                value={draft}
                disabled={!selected.span.editable || busy || !!inlineDraft}
                onChange={(e) => {
                  setDraft(e.target.value);
                  setPreview(undefined);
                }}
              />
              <button className="primary full" {...commandProps('edit.preview')} />
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
              <button className="full" {...commandProps('document.compare')} />
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
      {paletteOpen && (
        <CommandPalette
          context={context}
          onClose={() => setPaletteOpen(false)}
          onExecute={execute}
          restoreFocus={paletteFocus.current}
        />
      )}
    </div>
  );
}
