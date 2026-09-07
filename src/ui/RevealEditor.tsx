import { useEffect, useRef } from 'react';
import { EditorState, RangeSetBuilder } from '@codemirror/state';
import { Decoration, EditorView, WidgetType, type DecorationSet } from '@codemirror/view';
import type { CodeToken, Story, Token } from '../engine/document';
import { paragraphGeometry, paragraphLineStyle } from './paragraph-layout';
import { deletionRange, validInlineText, type InlineEditing } from './inline-edit';

class InlineWidget extends WidgetType {
  constructor(readonly editing: () => InlineEditing) {
    super();
  }
  toDOM(): HTMLElement {
    const input = document.createElement('textarea');
    const draft = this.editing().draft!;
    input.className = 'inline-text';
    input.setAttribute('aria-label', 'Inline text');
    input.value = draft.text;
    input.rows = 1;
    input.spellcheck = false;
    let accepted = draft.text;
    let composing = false;
    const resize = () => {
      input.style.width = `${Math.max(12, Math.min(70, input.value.length + 2))}ch`;
      input.style.height = 'auto';
      input.style.height = `${Math.max(28, input.scrollHeight)}px`;
    };
    const update = () => {
      if (this.editing().busy) {
        input.value = accepted;
        return;
      }
      if (!validInlineText(input.value)) {
        input.value = accepted;
        this.editing().message('Inline edits cannot contain tabs, line breaks, or invalid text.');
      } else accepted = input.value;
      this.editing().change(input.value, input.selectionStart, input.selectionEnd);
      resize();
    };
    input.oninput = () => {
      if (!composing) update();
    };
    input.onbeforeinput = (event) => {
      if (this.editing().busy) event.preventDefault();
    };
    input.addEventListener('compositionstart', () => {
      composing = true;
    });
    input.addEventListener('compositionend', () => {
      composing = false;
      update();
    });
    input.onselect = () => {
      if (!composing) this.editing().change(input.value, input.selectionStart, input.selectionEnd);
    };
    input.onkeydown = (event) => {
      if (this.editing().busy) {
        event.preventDefault();
        return;
      }
      if (event.isComposing || composing) return;
      if (event.key === 'Escape') {
        event.preventDefault();
        this.editing().cancel();
      }
      if (event.key === 'Enter') {
        event.preventDefault();
        this.editing().message(
          'Paragraph split is not supported yet. Preview and apply this span edit.',
        );
      }
      if (event.key === 'Backspace' || event.key === 'Delete') {
        event.preventDefault();
        const range = deletionRange(
          input.value,
          input.selectionStart,
          input.selectionEnd,
          event.key === 'Backspace',
        );
        if (range.from === range.to) {
          this.editing().message('This edit stops at the text span boundary.');
          return;
        }
        input.setRangeText('', range.from, range.to, 'start');
        update();
      }
    };
    input.onpaste = (event) => {
      const text = event.clipboardData?.getData('text/plain');
      if (text !== undefined && !validInlineText(text)) {
        event.preventDefault();
        this.editing().message('Paste plain text without tabs or line breaks into this span.');
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

/** The projection stays read-only; an inline draft replaces one source span visually. */
export function RevealEditor({
  story,
  codes,
  onSelect,
  inline,
  embedded = false,
}: {
  story: Story;
  codes: boolean;
  onSelect: (token: Token) => void;
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
      if (token.kind === 'text') {
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
    const begin = (editor: EditorView, insert?: string, backward?: boolean) => {
      const { from, to } = editor.state.selection.main;
      const location =
        locations.find((l) => from >= l.from && from < l.to && to <= l.to) ??
        locations.find((l) => from >= l.from && to <= l.to);
      if (!location || location.token.kind !== 'text') {
        editing.current.message(
          'Select text within one span to edit; codes and boundaries are preserved.',
        );
        return;
      }
      const span = location.token.span;
      if (!span.editable) {
        editing.current.message(span.reason ?? 'This text is protected.');
        return;
      }
      let anchor = Math.min(span.text.length, from - location.from);
      let head = Math.min(span.text.length, to - location.from);
      if (backward !== undefined) {
        const range = deletionRange(span.text, anchor, head, backward);
        if (range.from === range.to) {
          editing.current.message('This edit stops at the text span boundary.');
          return;
        }
        anchor = range.from;
        head = range.to;
      }
      editing.current.start(span.id, anchor, head, insert);
    };
    const restore =
      editing.current.restore !== lastRestore.current ? editing.current.restore : caret.current;
    lastRestore.current = editing.current.restore;
    const restored = locations.find(
      (l) => l.token.kind === 'text' && l.token.span.id === restore?.spanId,
    );
    const restoredLength = restored?.token.kind === 'text' ? restored.token.span.text.length : 0;
    const view = new EditorView({
      parent: host.current!,
      state: EditorState.create({
        doc: text,
        selection:
          restored && restore
            ? {
                anchor: restored.from + Math.min(restore.anchor, restoredLength),
                head: restored.from + Math.min(restore.head, restoredLength),
              }
            : undefined,
        extensions: [
          EditorState.readOnly.of(true),
          EditorView.lineWrapping,
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
              if (empty) editor.dispatch({ selection: { anchor: empty.from, head: empty.from } });
              begin(editor);
              return true;
            },
            mouseup: (_event, editor) => {
              const position = editor.state.selection.main.head;
              const location =
                locations.find((l) => position >= l.from && position < l.to) ??
                locations.find((l) => position === l.to);
              if (location) select.current(location.token);
              return false;
            },
            keydown: (event, editor) => {
              if (event.ctrlKey || event.metaKey || event.altKey || event.isComposing) return false;
              const deleting = event.key === 'Backspace' || event.key === 'Delete';
              if (event.key !== 'Enter' && event.key.length !== 1 && !deleting) return false;
              begin(
                editor,
                deleting ? '' : event.key === 'Enter' ? undefined : event.key,
                deleting ? event.key === 'Backspace' : undefined,
              );
              event.preventDefault();
              return true;
            },
            paste: (event, editor) => {
              const value = event.clipboardData?.getData('text/plain');
              if (value !== undefined) {
                event.preventDefault();
                if (validInlineText(value)) begin(editor, value);
                else
                  editing.current.message(
                    'Paste plain text without tabs or line breaks into one span.',
                  );
              }
              return true;
            },
          }),
        ],
      }),
    });
    if (restored && !editing.current.draft) view.focus();
    return () => {
      const { anchor, head, from, to } = view.state.selection.main;
      const location = locations.find((l) => from >= l.from && to <= l.to);
      caret.current =
        location?.token.kind === 'text'
          ? {
              spanId: location.token.span.id,
              anchor: anchor - location.from,
              head: head - location.from,
            }
          : undefined;
      view.destroy();
    };
  }, [story, codes, inline.draft?.spanId, inline.restore]);
  return <div className={`reveal-editor${embedded ? ' embedded' : ''}`} ref={host} />;
}
