import type { ParagraphLayout } from '../engine/paragraph';

/** Fixed display scale, bounded to keep unusually large/negative indents navigable. */
export function paragraphGeometry(layout: ParagraphLayout, minimumMarkerWidth = 0) {
  const bounded = (twips: number, min: number, max: number) =>
    Math.max(min, Math.min(max, twips / 15));
  let left = bounded(layout.left, -12, 240);
  const right = bounded(layout.right, -12, 160);
  let firstLine = bounded(layout.firstLine, -Math.max(left, 0) - 12, 120);
  // Long labels may outgrow the document's hanging space in our fixed-width
  // font. Widen that space so the label cannot touch text and wrapping aligns.
  if (firstLine < 0 && minimumMarkerWidth > -firstLine) {
    left += minimumMarkerWidth + firstLine;
    firstLine = -minimumMarkerWidth;
  }
  return { left, right, firstLine, markerWidth: Math.max(minimumMarkerWidth, -firstLine, 0) };
}

export function paragraphLineStyle(
  layout: ParagraphLayout,
  continuation = false,
  minimumMarkerWidth = 0,
): string {
  const g = paragraphGeometry(layout, minimumMarkerWidth);
  return `margin-left:${g.left}px;margin-right:${g.right}px;text-indent:${continuation ? 0 : g.firstLine}px;`;
}
