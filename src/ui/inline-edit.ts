import {
  positionAt,
  positionOffset,
  replacePieces,
  type TextPiece,
  type TextPosition,
} from '../engine/text-range';
import type { KeyBinding } from './keymap';

/** A draft keeps one piece per source span, including empty spans. */
export interface InlineDraft {
  spanId: string;
  pieces: TextPiece[];
  original: TextPiece[];
  text: string;
  anchor: number;
  head: number;
  typing: TextPosition;
  initialAnchor: TextPosition;
  initialHead: TextPosition;
}
export interface InlineEditing {
  history: { past: InlineDraft[]; future: InlineDraft[] };
  busy: boolean;
  keymap: readonly KeyBinding[];
  draft?: InlineDraft;
  restore?: { anchor: TextPosition; head: TextPosition };
  selectRange: (anchor?: TextPosition, head?: TextPosition) => void;
  start: (anchor: TextPosition, head: TextPosition, insert?: string, backward?: boolean) => void;
  change: (draft: InlineDraft) => void;
  cancel: () => void;
  message: (message: string) => void;
}

export function editDraft(draft: InlineDraft, from: number, to: number, text: string): InlineDraft {
  if (!validInlineText(text))
    throw new Error('Inline edits cannot contain tabs, line breaks, or invalid text.');
  const result = replacePieces(draft.pieces, from, to, text, draft.typing);
  const offset = positionOffset(result.pieces, result.caret);
  return {
    ...draft,
    pieces: result.pieces,
    text: result.pieces.map((p) => p.text).join(''),
    anchor: offset,
    head: offset,
    typing: result.caret,
  };
}
export function selectDraft(draft: InlineDraft, anchor: number, head: number): InlineDraft {
  return {
    ...draft,
    anchor,
    head,
    typing:
      anchor === draft.anchor && head === draft.head
        ? draft.typing
        : positionAt(draft.pieces, Math.min(anchor, head)),
  };
}
/** Reconcile one browser input, preferring its actual selection over an ambiguous string diff. */
export function inputDraft(
  draft: InlineDraft,
  text: string,
  anchor: number,
  head: number,
  selection?: { from: number; to: number },
): InlineDraft {
  if (text === draft.text) return selectDraft(draft, anchor, head);
  let from = selection?.from ?? Math.min(draft.anchor, draft.head);
  let to = selection?.to ?? Math.max(draft.anchor, draft.head);
  const insertedLength = text.length - (draft.text.length - (to - from));
  if (
    insertedLength < 0 ||
    text.slice(0, from) !== draft.text.slice(0, from) ||
    text.slice(from + insertedLength) !== draft.text.slice(to)
  ) {
    from = 0;
    while (from < text.length && from < draft.text.length && text[from] === draft.text[from])
      from++;
    to = draft.text.length;
    let end = text.length;
    while (to > from && end > from && draft.text[to - 1] === text[end - 1]) {
      to--;
      end--;
    }
    // Shared UTF-16 prefixes can end inside an emoji.
    if (from > 0 && /[\ud800-\udbff]/.test(draft.text[from - 1]!)) from--;
    if (to < draft.text.length && /[\udc00-\udfff]/.test(draft.text[to]!)) to++;
  }
  const length = text.length - (draft.text.length - (to - from));
  return { ...editDraft(draft, from, to, text.slice(from, from + length)), anchor, head };
}

export function validInlineText(text: string): boolean {
  // Match the engine's text-only contract, and exclude the projection object marker.
  return !/[\u0000-\u001f\ufffc\ufffe\uffff]/u.test(text) && !/[\ud800-\udfff]/u.test(text);
}

/** Expand deletion to complete graphemes, including combining marks and ZWJ emoji. */
export function deletionRange(text: string, from: number, to: number, backward: boolean) {
  const boundaries = [
    ...Array.from(
      new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(text),
      (segment) => segment.index,
    ),
    text.length,
  ];
  const reversed = [...boundaries].reverse();
  if (from === to) {
    if (backward) from = reversed.find((offset) => offset < from) ?? 0;
    else to = boundaries.find((offset) => offset > to) ?? text.length;
  }
  return {
    from: reversed.find((offset) => offset <= from) ?? 0,
    to: boundaries.find((offset) => offset >= to) ?? text.length,
  };
}
