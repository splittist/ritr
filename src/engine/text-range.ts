import type { Story, TextSpan } from './document';
import { assertText } from '../package/xml';

export interface TextPosition {
  spanId: string;
  offset: number;
  affinity?: 'left' | 'right';
}
export interface TextPiece {
  spanId: string;
  text: string;
}
export interface TextSegment {
  spans: TextSpan[];
  text: string;
}

/** Formatting is transparent; structure, review, and opaque codes are boundaries. */
export function textSegments(story: Story): TextSegment[] {
  const segments: TextSegment[] = [];
  let spans: TextSpan[] = [];
  const flush = () => {
    if (spans.length) segments.push({ spans, text: spans.map((s) => s.text).join('') });
    spans = [];
  };
  for (const token of story.tokens) {
    if (token.kind === 'text') spans.push(token.span);
    else if (token.category !== 'format') flush();
  }
  flush();
  return segments;
}

/** Read-only text also bounds an inline draft, so editable neighbors remain usable. */
export function editingSegments(story: Story): TextSegment[] {
  return textSegments(story).flatMap((segment) => {
    const groups: TextSegment[] = [];
    for (const span of segment.spans) {
      const last = groups.at(-1);
      if (last && last.spans[0]!.editable === span.editable) {
        last.spans.push(span);
        last.text += span.text;
      } else groups.push({ spans: [span], text: span.text });
    }
    return groups;
  });
}

export function positionOffset(pieces: readonly TextPiece[], position: TextPosition): number {
  let offset = 0;
  for (const piece of pieces) {
    if (piece.spanId === position.spanId) {
      if (
        !Number.isInteger(position.offset) ||
        position.offset < 0 ||
        position.offset > piece.text.length
      )
        throw new Error('Invalid text position');
      return offset + position.offset;
    }
    offset += piece.text.length;
  }
  throw new Error('Selection crosses a structural boundary');
}

export function positionAt(
  pieces: readonly TextPiece[],
  offset: number,
  affinity: 'left' | 'right' = 'left',
  preferred?: string,
): TextPosition {
  let start = 0;
  const candidates: TextPosition[] = [];
  for (const piece of pieces) {
    if (offset >= start && offset <= start + piece.text.length)
      candidates.push({ spanId: piece.spanId, offset: offset - start, affinity });
    start += piece.text.length;
  }
  const chosen =
    candidates.find((p) => p.spanId === preferred) ??
    (affinity === 'left'
      ? candidates.find((p) => p.offset > 0)
      : [...candidates]
          .reverse()
          .find((p) => p.offset < pieces.find((s) => s.spanId === p.spanId)!.text.length)) ??
    candidates.find((p) => pieces.find((s) => s.spanId === p.spanId)!.text.length > 0) ??
    candidates[0];
  if (!chosen) throw new Error('Invalid text offset');
  return chosen;
}

/** Pieces remain in source order; empty source elements are retained. */
export function replacePieces(
  pieces: readonly TextPiece[],
  from: number,
  to: number,
  text: string,
  typing?: TextPosition,
) {
  assertText(text);
  const before = pieces.map((p) => p.text).join('');
  if (
    !Number.isInteger(from) ||
    !Number.isInteger(to) ||
    from < 0 ||
    to < from ||
    to > before.length
  )
    throw new Error('Invalid text range');
  for (const offset of [from, to])
    if (
      offset > 0 &&
      offset < before.length &&
      /[\ud800-\udbff]/.test(before[offset - 1]!) &&
      /[\udc00-\udfff]/.test(before[offset]!)
    )
      throw new Error('Split-surrogate text range');
  const insertion =
    from === to && typing && positionOffset(pieces, typing) === from
      ? typing
      : positionAt(pieces, from, from === to ? 'left' : 'right');
  // Identical replacements preserve the original distribution of formatting.
  if (before.slice(from, to) === text)
    return {
      pieces: pieces.map((p) => ({ ...p })),
      caret: from === to && !text ? insertion : positionAt(pieces, from + text.length),
    };
  let start = 0;
  let caret = insertion;
  const result = pieces.map((piece) => {
    const end = start + piece.text.length;
    const a = Math.max(0, Math.min(piece.text.length, from - start));
    const b = Math.max(a, Math.min(piece.text.length, to - start));
    const inserted = piece.spanId === insertion.spanId ? text : '';
    const value = piece.text.slice(0, a) + inserted + piece.text.slice(b);
    if (piece.spanId === insertion.spanId) caret = { ...insertion, offset: a + text.length };
    start = end;
    return { ...piece, text: value };
  });
  return { pieces: result, caret };
}

export function resolveTextRange(story: Story, anchor: TextPosition, head: TextPosition) {
  const segment = editingSegments(story).find((s) => s.spans.some((p) => p.id === anchor.spanId));
  if (!segment) throw new Error('Unknown text span');
  const pieces = segment.spans.map((s) => ({ spanId: s.id, text: s.text }));
  const a = positionOffset(pieces, anchor),
    h = positionOffset(pieces, head);
  const from = Math.min(a, h),
    to = Math.max(a, h);
  let start = 0;
  for (const span of segment.spans) {
    const end = start + span.text.length;
    if (
      !span.editable &&
      ((from < end && to > start) || (from === to && span.id === anchor.spanId))
    )
      throw new Error(span.reason ?? 'Protected text');
    start = end;
  }
  return { segment, pieces, from, to };
}
