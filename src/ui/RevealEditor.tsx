import { useEffect, useRef } from 'react';
import { EditorState, RangeSetBuilder } from '@codemirror/state';
import { Decoration, EditorView, WidgetType, type DecorationSet } from '@codemirror/view';
import type { CodeToken, Story, Token } from '../engine/document';
import { paragraphGeometry, paragraphLineStyle } from './paragraph-layout';

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

/** CodeMirror is a read-only projection. Selecting text opens a command-based editor. */
export function RevealEditor({
  story,
  codes,
  onSelect,
}: {
  story: Story;
  codes: boolean;
  onSelect: (token: Token) => void;
}) {
  const host = useRef<HTMLDivElement>(null);
  const select = useRef(onSelect);
  select.current = onSelect;
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
          Decoration.mark({ class: token.span.editable ? 'editable-span' : 'protected-span' }),
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
    const view = new EditorView({
      parent: host.current!,
      state: EditorState.create({
        doc: text,
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
            mouseup: (_event, editor) => {
              const position = editor.state.selection.main.head;
              const location =
                locations.find((l) => position >= l.from && position < l.to) ??
                locations.find((l) => position === l.to);
              if (location) select.current(location.token);
              return false;
            },
            keydown: (event, editor) => {
              if (event.key !== 'Enter') return false;
              const position = editor.state.selection.main.head;
              const location = locations.find((l) => position >= l.from && position <= l.to);
              if (location) select.current(location.token);
              event.preventDefault();
              return true;
            },
          }),
        ],
      }),
    });
    return () => view.destroy();
  }, [story, codes]);
  return <div className="reveal-editor" ref={host} />;
}
