import type { Token } from './document';
import { descendants, type XmlNode } from '../package/xml';
import { readTable, type Table } from './table';

/** Half-open ranges refer to Story.tokens, so text is indexed once, not copied. */
export interface TokenRange {
  from: number;
  to: number;
}
export type Block =
  | { kind: 'paragraph'; id: string; range: TokenRange }
  | { kind: 'table'; id: string; range: TokenRange; table: Table }
  | { kind: 'codes'; id: string; range: TokenRange };

/** Add hierarchy to the source-bound stream without changing its order or identities. */
export function buildBlocks(
  root: XmlNode,
  tokens: Token[],
  warn: (message: string) => void,
): Block[] {
  const nodes = new Map(descendants(root).map((n) => [n.id, n]));
  const ranges = new Map<string, TokenRange>();
  tokens.forEach((token, index) => {
    const id = token.kind === 'text' ? token.span.source.nodeId : token.source.nodeId;
    const range = ranges.get(id);
    if (range) range.to = index + 1;
    else ranges.set(id, { from: index, to: index + 1 });
  });
  function read(from: number, to: number, depth = 0): Block[] {
    const blocks: Block[] = [];
    let index = from;
    while (index < to) {
      const token = tokens[index]!;
      if (token.kind === 'code' && (token.role === 'paragraph-start' || token.label === 'tbl')) {
        const range = ranges.get(token.source.nodeId)!;
        if (range.from === index && range.to <= to && range.to > index + 1) {
          if (token.role === 'paragraph-start')
            blocks.push({ kind: 'paragraph', id: token.source.nodeId, range });
          else {
            const node = nodes.get(token.source.nodeId)!;
            const table = readTable(
              node,
              tokens,
              ranges,
              (cellRange) => read(cellRange.from + 1, cellRange.to - 1, depth + 1),
              depth,
            );
            for (const message of table.warnings) warn(`${table.id}: ${message}`);
            blocks.push({ kind: 'table', id: table.id, range, table });
          }
          index = range.to;
          continue;
        }
      }
      const previous = blocks.at(-1);
      if (previous?.kind === 'codes' && previous.range.to === index) previous.range.to++;
      else
        blocks.push({ kind: 'codes', id: `codes-${index}`, range: { from: index, to: index + 1 } });
      index++;
    }
    return blocks;
  }
  return read(0, tokens.length);
}
