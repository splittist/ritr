/** UI drafts never contain projection placeholders or cross source-span boundaries. */
export interface InlineDraft {
  spanId: string;
  text: string;
  anchor: number;
  head: number;
}

export interface InlineEditing {
  busy: boolean;
  draft?: InlineDraft;
  restore?: { spanId: string; anchor: number; head: number };
  start: (spanId: string, anchor: number, head: number, insert?: string) => void;
  change: (text: string, anchor: number, head: number) => void;
  cancel: () => void;
  message: (message: string) => void;
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
