import { effectiveFormat, type TextFormat } from '../engine/format';
import { useEffect, useRef } from 'react';
import { Annotation, EditorState, StateEffect, StateField } from '@codemirror/state';
import { Decoration, EditorView, WidgetType, type DecorationSet } from '@codemirror/view';
import type { CodeToken, Story, Token } from '../engine/document';
import { projectTokens, tokenId, type InputHistory } from '../engine/projection';
import { textStyle } from './text-format';
import { editorCommand } from './keymap';
import { deletionRange } from './inline-edit';
import { paragraphGeometry, paragraphLineStyle } from './paragraph-layout';
import type { DirectEditing } from './direct-edit';

const external = Annotation.define<boolean>();
const replaceDecorations = StateEffect.define<DecorationSet>();
const decorations = StateField.define<DecorationSet>({
  create: () => Decoration.none,
  update: (value, transaction) => {
    const replacement = transaction.effects.find((e) => e.is(replaceDecorations));
    return replacement ? replacement.value : value.map(transaction.changes);
  },
  provide: (field) => EditorView.decorations.from(field),
});

class CodeWidget extends WidgetType {
  constructor(
    readonly token: CodeToken,
    readonly select: (token: Token) => void,
    readonly markerWidth = 0,
  ) {
    super();
  }
  toDOM(): HTMLElement {
    const button = document.createElement('button');
    const marker = this.token.role === 'list-label';
    button.className = marker ? 'list-marker' : `code-token ${this.token.category}`;
    button.textContent = this.token.label;
    button.title = marker
      ? `Inspect list ${this.token.paragraph?.numbering?.numId}, level ${(this.token.paragraph?.numbering?.level ?? 0) + 1}`
      : `Inspect ${this.token.label}`;
    if (marker && this.token.paragraph) {
      const list = this.token.paragraph.numbering!;
      if (list.suffix === 'tab') {
        button.style.width = `${this.markerWidth}px`;
        button.style.paddingRight = list.text ? '8px' : '0';
      }
      if (list.suffix === 'space') button.style.paddingRight = '0.6em';
      button.style.textAlign = ['left', 'right', 'center'].includes(list.alignment)
        ? list.alignment
        : 'left';
      button.setAttribute('aria-label', `List label ${this.token.label || '(none)'}`);
      if (list.warnings.length) button.classList.add('unresolved');
    }
    button.onclick = () => this.select(this.token);
    return button;
  }
  ignoreEvent(): boolean {
    return true;
  }
}

class EmptyWidget extends WidgetType {
  constructor(
    readonly spanId: string,
    readonly activate: () => void,
  ) {
    super();
  }
  toDOM() {
    const span = document.createElement('span');
    span.className = 'editable-span empty-span';
    span.dataset.spanId = this.spanId;
    span.title = 'Empty text span';
    span.textContent = '\u200b';
    span.onmousedown = (event) => {
      event.preventDefault();
      this.activate();
    };
    return span;
  }
}
class HiddenWidget extends WidgetType {
  toDOM() {
    const span = document.createElement('span');
    span.style.display = 'none';
    return span;
  }
}

/** Native CodeMirror text input, with source mutations serialized through the engine. */
export function RevealEditor({
  story,
  codes,
  onSelect,
  inline,
  embedded = false,
}: {
  story: Story;
  codes: boolean;
  onSelect: (token?: Token) => void;
  inline: DirectEditing;
  embedded?: boolean;
}) {
  const host = useRef<HTMLDivElement>(null);
  const viewRef = useRef<EditorView>(undefined);
  const current = useRef({ story, codes, onSelect, inline });
  current.current = { story, codes, onSelect, inline };
  const projection = useRef(projectTokens(story.tokens));
  const composing = useRef<string | undefined>(undefined);
  const preferred = useRef<string | undefined>(undefined);
  const lastRestore = useRef<DirectEditing['restore']>(undefined);
  const lastFocus = useRef<DirectEditing['focus']>(undefined);
  const sync = useRef(() => {});
  useEffect(() => {
    let alive = true;
    let typingFormat: TextFormat | undefined;
    const applyShortcut: { current?: (request: TextFormat | 'b' | 'i' | 'u') => void } = {};
    let history: (InputHistory & { time: number }) | undefined;
    const origin = () => tokenId(current.current.story.tokens[0]!);
    const selected = () => {
      const view = viewRef.current!;
      const { from, to } = view.state.selection.main;
      return (
        projection.current.locations.find(
          (l) =>
            l.token.kind === 'text' &&
            l.from <= from &&
            to <= l.to &&
            (from < l.to || l.from === l.to),
        ) ??
        projection.current.locations.find(
          (l) => l.token.kind === 'text' && l.from <= from && to <= l.to,
        )
      );
    };
    const activate = () => {
      const view = viewRef.current!;
      const selection = view.state.selection.main;
      const locations = selection.empty
        ? [selected()].filter(Boolean)
        : projection.current.locations.filter(
            (l) => l.token.kind === 'text' && l.from < selection.to && l.to > selection.from,
          );
      const textSpans = locations.flatMap((l) => (l?.token.kind === 'text' ? [l.token.span] : []));
      const values: TextFormat = {
        ...effectiveFormat(textSpans[0]),
        ...(selection.empty ? typingFormat : {}),
      };
      if (!selection.empty)
        for (const key of ['b', 'i', 'u'] as const)
          values[key] = textSpans.length > 0 && textSpans.every((s) => effectiveFormat(s)[key]);
      const editable =
        textSpans.every((s) => s.editable) &&
        !view.state.doc.sliceString(selection.from, selection.to).includes('\ufffc') &&
        (selection.empty
          ? projection.current.paragraphs.some(
              (p) => p.from <= selection.from && selection.from <= p.to,
            )
          : textSpans.length > 0);
      const applyFormat = (request: TextFormat | 'b' | 'i' | 'u') => {
        if (
          !alive ||
          current.current.inline.busy ||
          current.current.inline.pending ||
          composing.current !== undefined
        )
          return;
        if (!editable) {
          current.current.inline.message(
            'Select editable text or place the caret in an editable paragraph.',
          );
          return;
        }
        const format = typeof request === 'string' ? { [request]: !values[request] } : request;
        history = undefined;
        const range = view.state.selection.main;
        if (range.empty) {
          typingFormat = { ...typingFormat, ...format };
          activate();
        } else {
          typingFormat = undefined;
          current.current.inline.format({
            origin: origin(),
            from: range.from,
            to: range.to,
            format,
          });
        }
        view.focus();
      };
      applyShortcut.current = applyFormat;
      current.current.inline.activate({
        values,
        editable,
        apply: applyFormat,
      });
    };
    const send = (from: number, to: number, text: string, group?: InputHistory, enter?: boolean) =>
      current.current.inline.change({
        origin: origin(),
        from,
        to,
        text,
        typingSpan: preferred.current,
        history: group,
        typingFormat,
        enter,
      });
    const flushComposition = () => {
      if (!alive || composing.current === undefined) return;
      const before = composing.current,
        after = viewRef.current!.state.doc.toString();
      composing.current = undefined;
      let from = 0,
        to = before.length,
        end = after.length;
      while (from < to && from < end && before[from] === after[from]) from++;
      while (to > from && end > from && before[to - 1] === after[end - 1]) {
        to--;
        end--;
      }
      if (from > 0 && /[\ud800-\udbff]/.test(before[from - 1]!)) from--;
      if (to < before.length && /[\udc00-\udfff]/.test(before[to]!)) {
        to++;
        end++;
      }
      if (before !== after) send(from, to, after.slice(from, end));
      else sync.current();
      current.current.inline.composition(false);
    };
    const view = new EditorView({
      parent: host.current!,
      state: EditorState.create({
        doc: projection.current.text,
        extensions: [
          decorations,
          EditorView.lineWrapping,
          EditorView.contentAttributes.of({
            'aria-label': 'Reveal Codes document',
            spellcheck: 'false',
          }),
          EditorState.transactionFilter.of((transaction) => {
            if (!transaction.docChanged || transaction.annotation(external)) return transaction;
            if (current.current.inline.busy || transaction.startState.selection.ranges.length !== 1)
              return [];
            let valid = true;
            transaction.changes.iterChanges((from, to, _a, _b, inserted) => {
              if (
                transaction.startState.doc.sliceString(from, to).includes('\ufffc') ||
                /[\u0000-\u0009\u000b-\u001f\ufffc\ufffe\uffff]/u.test(inserted.toString())
              )
                valid = false;
            });
            transaction.changes.iterChanges((from, to) => {
              const protectedRange = projection.current.locations.find(
                (l) =>
                  l.token.kind === 'text' &&
                  !l.token.span.editable &&
                  ((from < l.to && to > l.from) ||
                    (from === to &&
                      from >= l.from &&
                      from <= l.to &&
                      (preferred.current === l.token.span.id || (from > l.from && from < l.to)))),
              );
              if (protectedRange) valid = false;
            });
            if (!valid) {
              current.current.inline.message(
                'This edit would change a protected object, tab, or unsupported character.',
              );
              return [];
            }
            return transaction;
          }),
          EditorView.updateListener.of((update) => {
            const nativeChange =
              update.docChanged && !update.transactions.some((t) => t.annotation(external));
            if (nativeChange) {
              if (composing.current === undefined) {
                const changes: { from: number; to: number; text: string }[] = [];
                update.changes.iterChanges((from, to, _a, _b, inserted) =>
                  changes.push({ from, to, text: inserted.toString() }),
                );
                const tx = update.transactions[0];
                const kind = tx?.isUserEvent('input.type')
                  ? 'typing'
                  : tx?.isUserEvent('delete.backward')
                    ? 'backspace'
                    : tx?.isUserEvent('delete.forward')
                      ? 'delete'
                      : undefined;
                const change = changes[0];
                const groupable =
                  kind &&
                  changes.length === 1 &&
                  update.startState.selection.main.empty &&
                  change &&
                  !change.text.includes('\n') &&
                  !update.startState.doc.sliceString(change.from, change.to).includes('\n');
                const now = Date.now();
                if (!groupable) history = undefined;
                else if (!history || history.kind !== kind || now - history.time > 1000)
                  history = { id: crypto.randomUUID(), kind, time: now };
                else history.time = now;
                // Descending offsets compose independently within one native transaction.
                for (const change of changes.reverse())
                  send(
                    change.from,
                    change.to,
                    change.text,
                    history,
                    tx?.isUserEvent('input.paragraph'),
                  );
                preferred.current = undefined;
              }
              projection.current = {
                ...projection.current,
                text: update.state.doc.toString(),
                locations: projection.current.locations.map((l) => ({
                  ...l,
                  from: update.changes.mapPos(l.from, -1),
                  to: update.changes.mapPos(l.to, 1),
                })),
              };
            }
            if (
              update.selectionSet &&
              !nativeChange &&
              !update.transactions.some((t) => t.annotation(external))
            ) {
              history = undefined;
              typingFormat = undefined;
              const location = selected();
              preferred.current =
                location?.token.kind === 'text' ? location.token.span.id : undefined;
              current.current.onSelect(location?.token);
            }
            if (viewRef.current?.hasFocus && (nativeChange || update.selectionSet)) activate();
          }),
          EditorView.domEventHandlers({
            focus: () => {
              activate();
              return false;
            },
            blur: () => {
              history = undefined;
              return false;
            },
            mousedown: () => {
              history = undefined;
              return false;
            },
            compositionstart: () => {
              history = undefined;
              composing.current = view.state.doc.toString();
              current.current.inline.composition(true);
              return false;
            },
            compositionend: () => {
              setTimeout(flushComposition, 0);
              return false;
            },
            keydown: (event, editor) => {
              if (event.isComposing || composing.current !== undefined) return false;
              const command = editorCommand(event, current.current.inline.keymap);
              if (
                command === 'format.bold' ||
                command === 'format.italic' ||
                command === 'format.underline'
              ) {
                event.preventDefault();
                const key =
                  command === 'format.bold' ? 'b' : command === 'format.italic' ? 'i' : 'u';
                // Route keyboard and toolbar actions through the same active editor target.
                activate();
                applyShortcut.current?.(key);
                return true;
              }
              const deleting =
                command === 'text.deleteBackward' || command === 'text.deleteForward';
              if (command === 'paragraph.split' || deleting) {
                event.preventDefault();
                const selection = editor.state.selection.main;
                const range = deleting
                  ? deletionRange(
                      editor.state.doc.toString(),
                      selection.from,
                      selection.to,
                      command === 'text.deleteBackward',
                    )
                  : { from: selection.from, to: selection.to };
                const text = deleting ? '' : '\n';
                if (range.from !== range.to || text)
                  editor.dispatch({
                    changes: { ...range, insert: text },
                    selection: { anchor: range.from + text.length },
                    userEvent: deleting
                      ? command === 'text.deleteBackward'
                        ? 'delete.backward'
                        : 'delete.forward'
                      : 'input.paragraph',
                  });
                return true;
              }
              if (
                (event.ctrlKey || event.metaKey) &&
                !event.altKey &&
                ['b', 'i', 'u'].includes(event.key.toLowerCase())
              ) {
                event.preventDefault();
                return true;
              }
              // An unbound editing key must not fall through to browser behavior.
              if (
                ['Enter', 'Backspace', 'Delete'].includes(event.key) &&
                !event.ctrlKey &&
                !event.metaKey &&
                !event.altKey
              ) {
                event.preventDefault();
                return true;
              }
              return false;
            },
          }),
        ],
      }),
    });
    viewRef.current = view;
    sync.current = () => {
      if (!alive || current.current.inline.pending || composing.current !== undefined) return;
      const { story, codes, inline } = current.current;
      const next = projectTokens(story.tokens);
      const ranges: ReturnType<Decoration['range']>[] = [];
      const context = document.createElement('canvas').getContext('2d')!;
      context.font = '14px Consolas, "Cascadia Code", monospace';
      const widths = new Map(
        story.paragraphs.map((p) => [
          p.id,
          Math.ceil(context.measureText(p.numbering?.text ?? '').width) + 8,
        ]),
      );
      for (const location of next.locations) {
        const { token, from, to } = location;
        if (token.kind === 'text') {
          if (from !== to)
            ranges.push(
              Decoration.mark({
                class: token.span.editable ? 'editable-span' : 'protected-span',
                inclusive: false,
                attributes: { 'data-span-id': token.span.id, style: textStyle(token.span) },
              }).range(from, to),
            );
          else
            ranges.push(
              Decoration.widget({
                widget: new EmptyWidget(token.span.id, () => {
                  preferred.current = token.span.id;
                  view.dispatch({ selection: { anchor: from }, annotations: external.of(true) });
                  view.focus();
                  current.current.onSelect(token);
                }),
                side: -1,
              }).range(from),
            );
        } else {
          if (token.role === 'paragraph-start' && token.paragraph)
            ranges.push(
              Decoration.line({
                attributes: {
                  class: 'paragraph-line',
                  'data-paragraph': token.paragraph.id,
                  style: paragraphLineStyle(
                    token.paragraph.layout,
                    false,
                    widths.get(token.paragraph.id),
                  ),
                },
              }).range(from),
            );
          const visible = codes || token.category === 'opaque' || token.role === 'list-label';
          const widget = visible
            ? new CodeWidget(
                token,
                (t) => current.current.onSelect(t),
                token.paragraph
                  ? paragraphGeometry(token.paragraph.layout, widths.get(token.paragraph.id))
                      .markerWidth
                  : 0,
              )
            : new HiddenWidget();
          if (to > from && token.role !== 'paragraph-end')
            ranges.push(Decoration.replace({ widget }).range(from, to));
          else if (visible)
            ranges.push(
              Decoration.widget({
                widget,
                // Equal-position widgets follow source order. Giving closing tags a
                // different side puts the next Run before the preceding /Run.
                side: -1,
              }).range(from),
            );
        }
      }
      const old = view.state.doc.toString();
      let { anchor, head } = view.state.selection.main;
      const restoreSelection =
        inline.restore &&
        inline.restore !== lastRestore.current &&
        inline.restore.origin === origin();
      if (restoreSelection) {
        history = undefined;
        typingFormat = undefined;
        anchor = inline.restore!.anchor;
        head = inline.restore!.head;
        lastRestore.current = inline.restore;
        preferred.current = undefined;
      }
      const requestFocus = inline.focus && inline.focus !== lastFocus.current;
      if (requestFocus) {
        const target = next.locations.find(
          (l) => l.token.kind === 'text' && l.token.span.id === inline.focus?.spanId,
        );
        if (target) {
          anchor = target.from;
          head = target.to;
        }
        lastFocus.current = inline.focus;
      }
      projection.current = next;
      view.dispatch({
        changes: old === next.text ? undefined : { from: 0, to: old.length, insert: next.text },
        selection: {
          anchor: Math.min(anchor, next.text.length),
          head: Math.min(head, next.text.length),
        },
        effects: replaceDecorations.of(Decoration.set(ranges, true)),
        annotations: external.of(true),
      });
      if (requestFocus || restoreSelection) view.focus();
      if (view.hasFocus) activate();
    };
    sync.current();
    return () => {
      alive = false;
      view.destroy();
      viewRef.current = undefined;
    };
  }, []);
  useEffect(() => {
    sync.current();
  }, [story, codes, inline.pending, inline.focus, inline.restore]);
  return <div className={`reveal-editor${embedded ? ' embedded' : ''}`} ref={host} />;
}
