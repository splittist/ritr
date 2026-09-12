import { textStyle } from './text-format';
import { editorCommand } from './keymap';
import { type TextPosition } from '../engine/text-range';
import { useEffect, useRef } from 'react';
import { EditorState, RangeSetBuilder } from '@codemirror/state';
import { Decoration, EditorView, WidgetType, type DecorationSet } from '@codemirror/view';
import type { CodeToken, Story, Token } from '../engine/document';
import { paragraphGeometry, paragraphLineStyle } from './paragraph-layout';
import {
  deletionRange,
  editDraft,
  inputDraft,
  selectDraft,
  validInlineText,
  type InlineEditing,
} from './inline-edit';

class InlineWidget extends WidgetType {
  constructor(readonly editing: () => InlineEditing) {
    super();
  }
  toDOM(): HTMLElement {
    const input = document.createElement('textarea');
    let draft = this.editing().draft!;
    const { past, future } = this.editing().history;
    input.className = 'inline-text';
    input.setAttribute('aria-label', 'Inline text');
    input.value = draft.text;
    input.rows = 1;
    input.spellcheck = false;
    let composing = false;
    let selection: { from: number; to: number } | undefined;
    const resize = () => {
      input.style.width = `${Math.max(12, Math.min(70, input.value.length + 2))}ch`;
      input.style.height = 'auto';
      input.style.height = `${Math.max(28, input.scrollHeight)}px`;
    };
    const render = () => {
      input.value = draft.text;
      input.setSelectionRange(draft.anchor, draft.head);
      resize();
    };
    const accept = (next: typeof draft, record = true) => {
      if (record && next.pieces.some((p, i) => p.text !== draft.pieces[i]?.text)) {
        past.push(draft);
        future.length = 0;
      }
      draft = next;
      this.editing().change(draft);
    };
    const history = (redo: boolean) => {
      const source = redo ? future : past,
        target = redo ? past : future;
      const next = source.pop();
      if (next) {
        target.push(draft);
        accept(next, false);
        render();
      }
    };
    const update = () => {
      if (this.editing().busy) {
        render();
        return;
      }
      try {
        accept(inputDraft(draft, input.value, input.selectionStart, input.selectionEnd, selection));
        selection = undefined;
      } catch (e) {
        render();
        this.editing().message(String(e).replace(/^Error: /, ''));
      }
      resize();
    };
    input.oninput = () => {
      if (!composing) update();
    };
    input.onbeforeinput = (event) => {
      if (this.editing().busy) {
        event.preventDefault();
        return;
      }
      if (event.inputType === 'historyUndo' || event.inputType === 'historyRedo') {
        event.preventDefault();
        history(event.inputType === 'historyRedo');
        return;
      }
      if (!composing) selection = { from: input.selectionStart, to: input.selectionEnd };
    };
    input.addEventListener('compositionstart', () => {
      selection = { from: input.selectionStart, to: input.selectionEnd };
      composing = true;
    });
    input.addEventListener('compositionend', () => {
      composing = false;
      update();
    });
    input.onselect = () => {
      if (!composing && input.value === draft.text)
        accept(selectDraft(draft, input.selectionStart, input.selectionEnd), false);
    };
    input.onkeydown = (event) => {
      if (this.editing().busy) {
        event.preventDefault();
        return;
      }
      if (event.isComposing || composing) return;
      const command = editorCommand(event, this.editing().keymap, true);
      if (command === 'text.undo' || command === 'text.redo') {
        event.preventDefault();
        history(command === 'text.redo');
        return;
      }
      if (
        !command &&
        ['Backspace', 'Delete'].includes(event.key) &&
        !event.ctrlKey &&
        !event.metaKey &&
        !event.altKey
      ) {
        event.preventDefault();
        return;
      }
      if (command === 'text.cancel') {
        event.preventDefault();
        this.editing().cancel();
      }
      if (command === 'text.edit') {
        event.preventDefault();
        this.editing().message(
          'Paragraph split is not supported yet. Preview and apply this text edit.',
        );
      }
      if (command === 'text.deleteBackward' || command === 'text.deleteForward') {
        event.preventDefault();
        const range = deletionRange(
          input.value,
          input.selectionStart,
          input.selectionEnd,
          command === 'text.deleteBackward',
        );
        if (range.from === range.to) {
          this.editing().message('This edit stops at a structural boundary.');
          return;
        }
        try {
          accept(editDraft(draft, range.from, range.to, ''));
          render();
        } catch (e) {
          this.editing().message(String(e));
        }
      }
    };
    input.onpaste = (event) => {
      const text = event.clipboardData?.getData('text/plain');
      if (text !== undefined && !validInlineText(text)) {
        event.preventDefault();
        this.editing().message('Paste plain text without tabs or line breaks.');
      }
    };
    requestAnimationFrame(() => {
      if (!input.isConnected) return;
      resize();
      input.focus();
      input.setSelectionRange(draft.anchor, draft.head);
    });
    return input;
  }
  ignoreEvent(): boolean {
    return true;
  }
}

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

/** The projection stays read-only; a draft replaces one contiguous editable segment visually. */
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
  inline: InlineEditing;
  embedded?: boolean;
}) {
  const host = useRef<HTMLDivElement>(null);
  const select = useRef(onSelect);
  select.current = onSelect;
  const editing = useRef(inline);
  editing.current = inline;
  const caret = useRef<InlineEditing['restore']>(undefined);
  const lastRestore = useRef<InlineEditing['restore']>(undefined);
  useEffect(() => {
    let text = '';
    const decorations = new RangeSetBuilder<Decoration>();
    const atomic = new RangeSetBuilder<Decoration>();
    const lines: ReturnType<Decoration['range']>[] = [];
    const locations: { from: number; to: number; token: Token }[] = [];
    const markers = new Map(
      story.tokens.flatMap((t) =>
        t.kind === 'code' && t.role === 'list-label' && t.paragraph
          ? [[t.paragraph.id, t] as const]
          : [],
      ),
    );
    const widths = new Map<string, number>();
    const context = document.createElement('canvas').getContext('2d')!;
    context.font = '14px Consolas, "Cascadia Code", monospace';
    for (const p of story.paragraphs) {
      const label = p.numbering;
      if (label?.suffix === 'tab' && label.text)
        widths.set(p.id, Math.ceil(context.measureText(label.text).width) + 8);
    }
    const addWidget = (token: CodeToken) => {
      const from = text.length;
      text += '\ufffc';
      const decoration = Decoration.replace({
        widget: new CodeWidget(
          token,
          (t) => select.current(t),
          token.paragraph
            ? paragraphGeometry(token.paragraph.layout, widths.get(token.paragraph.id)).markerWidth
            : 0,
        ),
      });
      decorations.add(from, text.length, decoration);
      atomic.add(from, text.length, decoration);
    };
    const draftIds = new Set(editing.current.draft?.pieces.map((p) => p.spanId));
    let draftStarted = false;
    let paragraph: CodeToken['paragraph'];
    for (const token of story.tokens) {
      if (token.kind === 'code' && token.role === 'list-label') continue;
      if (token.kind === 'code' && token.role === 'paragraph-start' && token.paragraph) {
        if (text && !text.endsWith('\n')) text += '\n';
        paragraph = token.paragraph;
        lines.push(
          Decoration.line({
            attributes: {
              class: 'paragraph-line',
              style: paragraphLineStyle(paragraph.layout, false, widths.get(paragraph.id)),
              'data-paragraph': paragraph.id,
            },
          }).range(text.length),
        );
        const marker = markers.get(paragraph.id);
        if (marker?.kind === 'code') addWidget(marker);
      }
      const from = text.length;
      if (draftStarted && token.kind === 'code' && token.category === 'format') continue;
      if (token.kind === 'code' && token.category !== 'format') draftStarted = false;
      if (
        token.kind === 'text' &&
        draftIds.has(token.span.id) &&
        token.span.id !== editing.current.draft?.spanId
      )
        continue;
      if (token.kind === 'text') {
        if (token.span.id === editing.current.draft?.spanId) draftStarted = true;
        // Empty source spans remain selectable without inserting literal document text.
        text += token.span.text || '\u200b';
        decorations.add(
          from,
          text.length,
          token.span.id === editing.current.draft?.spanId
            ? Decoration.replace({ widget: new InlineWidget(() => editing.current) })
            : Decoration.mark({
                class: `${token.span.editable ? 'editable-span' : 'protected-span'}${token.span.text ? '' : ' empty-span'}`,
                attributes: {
                  'data-span-id': token.span.id,
                  style: textStyle(token.span),
                  ...(token.span.text ? {} : { title: 'Empty text span' }),
                },
              }),
        );
        locations.push({ from, to: text.length, token });
      } else if (codes || token.category === 'opaque') {
        addWidget(token);
      }
      if (
        token.kind === 'code' &&
        (token.label === '¶' || token.label === 'br' || token.label === 'cr')
      ) {
        text += '\n';
        if (token.role === 'paragraph-end') paragraph = undefined;
        else if (paragraph)
          lines.push(
            Decoration.line({
              attributes: {
                class: 'paragraph-line',
                style: paragraphLineStyle(paragraph.layout, true, widths.get(paragraph.id)),
              },
            }).range(text.length),
          );
      }
      if (token.kind === 'code' && token.label === 'tab') text += '\t';
    }
    const ranges = decorations.finish();
    const atomicRanges = atomic.finish();
    const position = (
      offset: number,
      side: 'left' | 'right',
      preferred?: TextPosition,
    ): TextPosition | undefined => {
      const candidates = locations.filter(
        (l) => offset >= l.from && offset <= l.to && l.token.kind === 'text',
      );
      const location =
        candidates.find(
          (l) =>
            l.token.kind === 'text' &&
            l.token.span.id === preferred?.spanId &&
            Math.min(l.token.span.text.length, offset - l.from) === preferred.offset,
        ) ??
        (side === 'left'
          ? candidates.find((l) => offset > l.from)
          : candidates.find((l) => offset < l.to)) ??
        candidates[0];
      if (location?.token.kind !== 'text') return;
      return {
        spanId: location.token.span.id,
        offset: Math.min(location.token.span.text.length, offset - location.from),
        affinity: side,
      };
    };
    const selectedLocation = (editor: EditorView) => {
      const { from, to } = editor.state.selection.main;
      return (
        locations.find((l) => from >= l.from && from < l.to && to <= l.to) ??
        locations.find((l) => from >= l.from && to <= l.to)
      );
    };
    const begin = (editor: EditorView, insert?: string, backward?: boolean) => {
      const { from, to } = editor.state.selection.main;
      const anchor = position(
        from,
        from === to ? 'left' : 'right',
        from === to ? activeCaret : undefined,
      );
      const head = from === to ? anchor : position(to, 'left');
      if (!anchor || !head) {
        editing.current.message(
          'Select text within one text segment; structural boundaries are preserved.',
        );
        return;
      }
      const forward = editor.state.selection.main.anchor <= editor.state.selection.main.head;
      editing.current.start(forward ? anchor : head, forward ? head : anchor, insert, backward);
    };
    const restore =
      editing.current.restore !== lastRestore.current ? editing.current.restore : caret.current;
    lastRestore.current = editing.current.restore;
    const projectionOffset = (p?: TextPosition) => {
      const location = locations.find(
        (l) => l.token.kind === 'text' && l.token.span.id === p?.spanId,
      );
      return location?.token.kind === 'text' && p
        ? location.from + Math.min(p.offset, location.token.span.text.length)
        : undefined;
    };
    const restoredAnchor = projectionOffset(restore?.anchor),
      restoredHead = projectionOffset(restore?.head);
    const restored = restoredAnchor !== undefined && restoredHead !== undefined;
    let activeCaret = restored && restoredAnchor === restoredHead ? restore?.anchor : undefined;
    const view = new EditorView({
      parent: host.current!,
      state: EditorState.create({
        doc: text,
        selection: restored ? { anchor: restoredAnchor!, head: restoredHead! } : undefined,
        extensions: [
          EditorState.readOnly.of(true),
          EditorView.lineWrapping,
          EditorView.updateListener.of((update) => {
            if (update.selectionSet) {
              select.current(selectedLocation(update.view)?.token);
              const { from, to } = update.state.selection.main;
              const a = position(
                from,
                from === to ? 'left' : 'right',
                from === to ? activeCaret : undefined,
              );
              activeCaret = from === to ? a : undefined;
              editing.current.selectRange(a, from === to ? a : position(to, 'left'));
            }
          }),
          EditorView.decorations.of(ranges),
          EditorView.decorations.of(Decoration.set(lines, true)),
          EditorView.atomicRanges.of(() => atomicRanges as DecorationSet),
          EditorView.contentAttributes.of({
            'aria-label': 'Reveal Codes document',
            spellcheck: 'false',
          }),
          EditorView.domEventHandlers({
            dblclick: (event, editor) => {
              const target = (event.target as HTMLElement).closest('.empty-span');
              const empty = locations.find(
                (l) =>
                  l.token.kind === 'text' &&
                  l.token.span.id === target?.getAttribute('data-span-id'),
              );
              if (empty?.token.kind === 'text') {
                const point: TextPosition = {
                  spanId: empty.token.span.id,
                  offset: 0,
                  affinity: 'right',
                };
                editing.current.start(point, point);
              } else begin(editor);
              return true;
            },
            mouseup: (_event, editor) => {
              select.current(selectedLocation(editor)?.token);
              const { from, to } = editor.state.selection.main;
              const a = position(
                from,
                from === to ? 'left' : 'right',
                from === to ? activeCaret : undefined,
              );
              activeCaret = from === to ? a : undefined;
              editing.current.selectRange(a, from === to ? a : position(to, 'left'));
              return false;
            },
            keydown: (event, editor) => {
              if (event.isComposing) return false;
              const command = editorCommand(event, editing.current.keymap);
              const deleting =
                command === 'text.deleteBackward' || command === 'text.deleteForward';
              if (
                !deleting &&
                command !== 'text.edit' &&
                (event.ctrlKey || event.metaKey || event.altKey || event.key.length !== 1)
              )
                return false;
              begin(
                editor,
                deleting ? '' : command === 'text.edit' ? undefined : event.key,
                deleting ? command === 'text.deleteBackward' : undefined,
              );
              event.preventDefault();
              return true;
            },
            paste: (event, editor) => {
              const value = event.clipboardData?.getData('text/plain');
              if (value !== undefined) {
                event.preventDefault();
                if (validInlineText(value)) begin(editor, value);
                else editing.current.message('Paste plain text without tabs or line breaks.');
              }
              return true;
            },
          }),
        ],
      }),
    });
    if (restored && !editing.current.draft) view.focus();
    return () => {
      const { anchor, head } = view.state.selection.main;
      const a = position(anchor, 'left', anchor === head ? activeCaret : undefined),
        h = anchor === head ? a : position(head, 'left');
      caret.current = a && h ? { anchor: a, head: h } : undefined;
      view.destroy();
    };
  }, [story, codes, inline.draft?.spanId, inline.restore]);
  return <div className={`reveal-editor${embedded ? ' embedded' : ''}`} ref={host} />;
}
