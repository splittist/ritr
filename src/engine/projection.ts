import type { TextFormat } from './format';
import type { Story, Token } from './document';

export type ListAction = 'indent' | 'outdent' | 'enter';
export interface ProjectedParagraph {
  id: string;
  from: number;
  to: number;
  numbering?: { numId: string; level: number };
}
export const tokenId = (token: Token) => (token.kind === 'text' ? token.span.id : token.id);
/** Codes never add editable characters. Non-format objects occupy a protected slot. */
export function projectTokens(tokens: readonly Token[]) {
  let text = '';
  const locations: { from: number; to: number; token: Token }[] = [];
  const paragraphs: ProjectedParagraph[] = [];
  let lastParagraphEnd = -1;
  tokens.forEach((token, index) => {
    if (token.kind === 'code' && token.role === 'paragraph-end') lastParagraphEnd = index;
  });
  for (const [index, token] of tokens.entries()) {
    const from = text.length;
    if (token.kind === 'text') text += token.span.text;
    else if (token.role === 'paragraph-start')
      paragraphs.push({
        id: token.source.nodeId,
        from,
        to: from,
        numbering: token.paragraph?.numbering
          ? { numId: token.paragraph.numbering.numId, level: token.paragraph.numbering.level }
          : undefined,
      });
    else if (token.role === 'paragraph-end') {
      const paragraph = paragraphs.at(-1);
      if (paragraph) paragraph.to = from;
      // Separators belong between paragraphs. A terminal newline creates a phantom
      // editable line with no source paragraph to receive typing or a join.
      if (index !== lastParagraphEnd) text += '\n';
    } else if (
      token.category !== 'format' &&
      token.role !== 'list-label' &&
      token.role !== 'anchor' &&
      token.label !== 'sectPr'
    )
      text += '\ufffc';
    locations.push({ from, to: text.length, token });
  }
  return { text, locations, paragraphs };
}
export function projectStory(story: Story, origin: string) {
  const index = story.tokens.findIndex((t) => tokenId(t) === origin);
  if (index < 0) throw new Error('The editing region is stale');
  return projectTokens(story.tokens.slice(index));
}
export interface InputHistory {
  id: string;
  kind: 'typing' | 'backspace' | 'delete';
}
export interface ProjectionEdit {
  documentId: string;
  storyId: string;
  origin: string;
  from: number;
  to: number;
  text: string;
  expectedRevision: number;
  /** Explicit ownership for an empty run or a boundary caret. */
  typingSpan?: string;
  history?: InputHistory;
  typingFormat?: TextFormat;
  /** Dedicated Enter command, distinguished from pasted paragraph breaks. */
  enter?: boolean;
  listAction?: ListAction;
}

export interface ProjectionSelection {
  documentId: string;
  storyId: string;
  origin: string;
  anchor: number;
  head: number;
}

/** A selection ending at the next paragraph start does not include that paragraph. */
export function selectedParagraphs(
  paragraphs: readonly ProjectedParagraph[],
  from: number,
  to: number,
) {
  return paragraphs.filter((p) =>
    from === to ? p.from <= from && from <= p.to : p.from < to && p.to >= from,
  );
}
/** Keep semantic keyboard decisions current while native text waits for engine confirmation. */
export function mapParagraphs(
  paragraphs: readonly ProjectedParagraph[],
  from: number,
  to: number,
  text: string,
): ProjectedParagraph[] {
  const first = paragraphs.findIndex((p) => p.from <= from && from <= p.to);
  const last = paragraphs.findIndex((p) => p.from <= to && to <= p.to);
  if (first < 0 || last < first) return [...paragraphs];
  const delta = text.length - (to - from),
    chunks = text.split('\n');
  let start = paragraphs[first]!.from,
    end = from;
  const replacements = chunks.map((chunk, index) => {
    end += chunk.length;
    const p: ProjectedParagraph = {
      ...paragraphs[first]!,
      id: index ? `${paragraphs[first]!.id}:pending:${start}` : paragraphs[first]!.id,
      from: start,
      to: index === chunks.length - 1 ? paragraphs[last]!.to + delta : end,
    };
    start = end + 1;
    end++;
    return p;
  });
  return [
    ...paragraphs.slice(0, first),
    ...replacements,
    ...paragraphs.slice(last + 1).map((p) => ({ ...p, from: p.from + delta, to: p.to + delta })),
  ];
}
export function mapListAction(
  paragraphs: readonly ProjectedParagraph[],
  from: number,
  to: number,
  action: ListAction,
): ProjectedParagraph[] {
  const selected = new Set(selectedParagraphs(paragraphs, from, to));
  return paragraphs.map((p) => {
    if (!selected.has(p) || !p.numbering) return p;
    const level = p.numbering.level + (action === 'indent' ? 1 : -1);
    return { ...p, numbering: level < 0 ? undefined : { ...p.numbering, level } };
  });
}
