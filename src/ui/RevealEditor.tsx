import { useEffect, useRef } from 'react';
import { EditorState, RangeSetBuilder } from '@codemirror/state';
import { Decoration, EditorView, WidgetType, type DecorationSet } from '@codemirror/view';
import type { CodeToken, Story, Token } from '../engine/document';

class CodeWidget extends WidgetType {
  constructor(
    readonly token: CodeToken,
    readonly select: (token: Token) => void,
  ) {
    super();
  }
  toDOM(): HTMLElement {
    const button = document.createElement('button');
    button.className = `code-token ${this.token.category}`;
    button.textContent = this.token.label;
    button.title = `Inspect ${this.token.label}`;
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
    const locations: { from: number; to: number; token: Token }[] = [];
    for (const token of story.tokens) {
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
        text += '\ufffc';
        const decoration = Decoration.replace({
          widget: new CodeWidget(token, (t) => select.current(t)),
        });
        decorations.add(from, text.length, decoration);
        atomic.add(from, text.length, decoration);
      }
      if (
        token.kind === 'code' &&
        (token.label === '¶' || token.label === 'br' || token.label === 'cr')
      )
        text += '\n';
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
