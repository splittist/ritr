import { FormattingToolbar } from './FormattingToolbar';
import type { ProjectionEdit } from '../engine/projection';
import type { DirectEditing, FormattingTarget } from './direct-edit';
import { commandShortcuts, shortcutCommand, resolveKeymap } from './keymap';
import { useEffect, useRef, useState } from 'react';
import type { DesktopApi, Snapshot } from '../desktop/protocol';
import type { ChangePreview, SearchMatch } from '../engine/workspace';
import type { Token } from '../engine/document';
import type { PartDifference } from '../package/docx';
import { StoryView } from './StoryView';
import { CommandPalette, captureCommandFocus } from './CommandPalette';
import {
  commands,
  commandReason,
  invokeCommand,
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
  const [keymap, setKeymap] = useState(() => {
    try {
      return resolveKeymap(JSON.parse(localStorage.getItem('ritr.keymap') ?? '{}'));
    } catch {
      return resolveKeymap();
    }
  });
  const [keymapText, setKeymapText] = useState(() => localStorage.getItem('ritr.keymap') ?? '{}');
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
  const [pending, setPending] = useState(false);
  const [composing, setComposing] = useState(false);
  const [focus, setFocus] = useState<DirectEditing['focus']>();
  const [paletteOpen, setPaletteOpen] = useState(false);
  const [formatting, setFormatting] = useState<{
    documentId: string;
    storyId: string;
    target?: FormattingTarget;
  }>();
  const latest = useRef(state);
  const queue = useRef<ProjectionEdit[]>([]);
  const draining = useRef(false);
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
    latest.current = snapshot;
    setFocus(undefined);
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
  const inline: DirectEditing = {
    activate: (target) => {
      if (document && story) setFormatting({ documentId: document.id, storyId: story.id, target });
    },
    format: (edit) => {
      if (!document || !story || busy || pending || running.current) return;
      void run(async () => {
        refresh(
          await api!.formatProjection({
            ...edit,
            documentId: document.id,
            storyId: story.id,
            expectedRevision: document.revision,
          }),
        );
        setMessage('Text formatting applied. Undo restores the previous formatting.');
      });
    },
    get busy() {
      return busy || running.current;
    },
    composition: setComposing,
    pending,
    restore:
      state.selection?.documentId === document?.id && state.selection?.storyId === story?.id
        ? state.selection
        : undefined,
    keymap,
    focus,
    message: setError,
    change: (edit) => {
      if (!document || !story || busy || running.current) return;
      queue.current.push({
        ...edit,
        documentId: document.id,
        storyId: story.id,
        expectedRevision: 0,
      });
      setPreview(undefined);
      setMatches(undefined);
      setReport(undefined);
      setError('');
      setPending(true);
      if (draining.current) return;
      draining.current = true;
      void (async () => {
        try {
          while (queue.current.length) {
            const next = queue.current.shift()!;
            const revision = latest.current.documents.find(
              (d) => d.id === next.documentId,
            )?.revision;
            if (revision === undefined) throw new Error('The editing document is no longer open');
            latest.current = await api!.editProjection({ ...next, expectedRevision: revision });
          }
          setMessage('Text updated. Undo reverses the last edit.');
        } catch (e) {
          queue.current.length = 0;
          setError(String(e).replace(/^Error: /, ''));
          try {
            latest.current = await api!.snapshot();
          } catch {
            /* retain the last confirmed snapshot */
          }
        } finally {
          setState(latest.current);
          setSelected(undefined);
          draining.current = false;
          setPending(false);
        }
      })();
    },
  };
  useEffect(() => {
    if (!pending && !composing) return;
    const preventClose = (event: BeforeUnloadEvent) => {
      event.preventDefault();
      event.returnValue = '';
    };
    window.addEventListener('beforeunload', preventClose);
    return () => window.removeEventListener('beforeunload', preventClose);
  }, [pending, composing]);
  useEffect(() => {
    if (api)
      void run(async () => {
        refresh(await api.snapshot());
      });
  }, []);
  const selectionReason =
    selected?.kind !== 'text'
      ? 'Select editable document text.'
      : !selected.span.editable
        ? (selected.span.reason ?? 'This text is protected.')
        : undefined;
  const context: CommandContext = {
    connected: !!api,
    busy: busy || pending || composing,
    hasDocument: !!document,
    canUndo: state.canUndo,
    canRedo: state.canRedo,
    selectionReason,
    listReason:
      formatting?.target &&
      formatting.documentId === document?.id &&
      formatting.storyId === story?.id
        ? formatting.target.listReason
        : 'Place the caret in a list item.',
    editReason:
      selectionReason ??
      (selected?.kind === 'text' && draft === selected.span.text
        ? 'Edit the selected span first.'
        : undefined),
    hasQuery: !!query,
    hasPreview: !!preview,
    hasChanges: !!preview?.edits.length,
  };
  const actions: CommandActions = {
    'list.indent': () => formatting?.target?.list('indent'),
    'list.outdent': () => formatting?.target?.list('outdent'),
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
      if (selected?.kind === 'text') setFocus({ spanId: selected.span.id });
    },
    'edit.preview': async () => {
      if (selected?.kind === 'text')
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
        id === 'commands.open' || id === 'list.indent' || id === 'list.outdent'
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
        `${command.description}${commandShortcuts(id, keymap).length ? ` (${commandShortcuts(id, keymap).join(' / ')})` : ''}`,
      'aria-keyshortcuts':
        commandShortcuts(id, keymap)
          .map((key) => key.replace('Ctrl', 'Control'))
          .join(' ') || undefined,
      onClick: () => execute(id),
    };
  }
  useEffect(() => {
    const keydown = (event: KeyboardEvent) => {
      const target = event.target instanceof HTMLElement ? event.target : undefined;
      const nativeText =
        !!target?.closest('input, textarea') ||
        (!!target?.isContentEditable && !target.closest('.cm-content'));
      const id = shortcutCommand(event, nativeText, keymap);
      if (!id) return;
      event.preventDefault();
      event.stopPropagation();
      if (!currentCommands.current.paletteOpen || id === 'commands.open') execute(id);
    };
    window.addEventListener('keydown', keydown, true);
    return () => window.removeEventListener('keydown', keydown, true);
  }, [keymap]);
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
                disabled={pending || composing}
                className={`document-link ${document?.id === d.id ? 'active' : ''}`}
                onClick={() => {
                  setDocumentId(d.id);
                  setFocus(undefined);
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
                    disabled={pending || composing}
                    className={`story-link ${story?.id === s.id ? 'active' : ''}`}
                    onClick={() => {
                      setStoryId(s.id);
                      setFocus(undefined);
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
          <details className="keymap-settings">
            <summary>Keybindings</summary>
            <p>Map command IDs to key chords. An empty array removes a binding.</p>
            <textarea
              aria-label="Keybindings JSON"
              value={keymapText}
              onChange={(e) => setKeymapText(e.target.value)}
            />
            <button
              onClick={() => {
                try {
                  const overrides = JSON.parse(keymapText);
                  const next = resolveKeymap(overrides);
                  localStorage.setItem('ritr.keymap', JSON.stringify(overrides));
                  setKeymap(next);
                  setError('');
                  setMessage('Keybindings updated.');
                } catch (e) {
                  setError(String(e).replace(/^Error: /, ''));
                }
              }}
            >
              Apply keybindings
            </button>
          </details>
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
                  Type directly in the document. Enter splits a paragraph; Backspace/Delete joins at
                  an edge.
                </span>
                <span>
                  {pending
                    ? 'Applying input…'
                    : document?.dirty
                      ? 'Unsaved changes'
                      : 'Saved snapshot'}
                </span>
              </div>
              <FormattingToolbar
                busy={busy || pending || composing}
                target={
                  formatting?.documentId === document?.id && formatting?.storyId === story?.id
                    ? formatting.target
                    : undefined
                }
              />
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
                    disabled={pending || composing}
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
                disabled={!selected.span.editable || busy || pending}
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
          keymap={keymap}
          context={context}
          onClose={() => setPaletteOpen(false)}
          onExecute={execute}
          restoreFocus={paletteFocus.current}
        />
      )}
    </div>
  );
}
