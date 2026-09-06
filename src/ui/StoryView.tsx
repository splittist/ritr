import { useMemo } from 'react';
import type { Block, TokenRange } from '../engine/blocks';
import type { CodeToken, Story, Token } from '../engine/document';
import type { Table } from '../engine/table';
import { RevealEditor } from './RevealEditor';
import { cellStyle, tableStyle } from './table-layout';

interface Props {
  story: Story;
  codes: boolean;
  onSelect: (token: Token) => void;
}

/** HTML owns table geometry; each text region reuses the same CodeMirror projection. */
export function StoryView(props: Props) {
  return (
    <div className="story-view">
      <BlockView {...props} blocks={props.story.blocks} />
    </div>
  );
}
function TextRegion({ range, ...props }: Props & { range: TokenRange }) {
  const region = useMemo(() => {
    const tokens = props.story.tokens.slice(range.from, range.to);
    const paragraphIds = new Set(
      tokens.flatMap((t) => (t.kind === 'code' && t.paragraph ? [t.paragraph.id] : [])),
    );
    return {
      ...props.story,
      tokens,
      paragraphs: props.story.paragraphs.filter((p) => paragraphIds.has(p.id)),
      blocks: [],
    };
  }, [props.story, range.from, range.to]);
  return <RevealEditor story={region} codes={props.codes} onSelect={props.onSelect} embedded />;
}
function BlockView({ blocks, ...props }: Props & { blocks: Block[] }) {
  // Keep adjacent paragraph/code blocks in one editor, not one editor per paragraph.
  const regions: (TokenRange | (Block & { kind: 'table' }))[] = [];
  for (const block of blocks) {
    if (block.kind === 'table') regions.push(block);
    else {
      const previous = regions.at(-1);
      if (previous && !('kind' in previous) && previous.to === block.range.from)
        previous.to = block.range.to;
      else regions.push({ ...block.range });
    }
  }
  return (
    <>
      {regions.map((region, index) =>
        'kind' in region ? (
          <TableView key={region.id} {...props} table={region.table} range={region.range} />
        ) : (
          <TextRegion key={`text-${index}`} {...props} range={region} />
        ),
      )}
    </>
  );
}
function Inspect({
  token,
  onSelect,
  label,
}: {
  token: CodeToken;
  onSelect: Props['onSelect'];
  label?: string;
}) {
  return (
    <button className="table-code" onClick={() => onSelect(token)}>
      {label ?? token.label}
    </button>
  );
}
function TableView({ table, range, ...props }: Props & { table: Table; range: TokenRange }) {
  const totalWidth = table.columns.reduce((a, b) => a + b, 0);
  return (
    <section className="table-region" data-table-id={table.id} aria-label="Document table">
      <div className="table-caption">
        <Inspect
          token={table.inspector}
          onSelect={props.onSelect}
          label={`Table · ${table.rows.length} rows × ${table.columns.length} columns`}
        />
        <span>{table.linear ? 'Linear fallback' : 'Approximate layout'}</span>
      </div>
      {table.warnings.length > 0 && (
        <details className="table-warnings">
          <summary>
            {table.warnings.length} display {table.warnings.length === 1 ? 'note' : 'notes'}
          </summary>
          {table.warnings.map((w, i) => (
            <p key={i}>{w}</p>
          ))}
        </details>
      )}
      {table.linear ? (
        <TextRegion {...props} range={range} />
      ) : (
        <div className="table-scroll">
          <table
            className="document-table"
            style={tableStyle(table)}
            aria-label="Document table grid"
          >
            <colgroup>
              {table.columns.map((width, i) => (
                <col key={i} style={{ width: `${(width / totalWidth) * 100}%` }} />
              ))}
            </colgroup>
            <tbody>
              {table.rows.map((row, rowIndex) => {
                const used = row.before + row.cells.reduce((n, cell) => n + cell.colSpan, 0);
                const after = Math.max(0, table.columns.length - used);
                const firstVisible = row.cells.find((c) => !c.mergedInto)?.id;
                return (
                  <tr
                    key={row.id}
                    data-row-id={row.id}
                    className={row.header ? 'table-header-row' : ''}
                  >
                    {row.before > 0 && (
                      <td
                        className="grid-gap"
                        colSpan={row.before}
                        aria-label="Omitted grid columns"
                      />
                    )}
                    {row.cells
                      .filter((cell) => !cell.mergedInto)
                      .map((cell) => (
                        <td
                          key={cell.id}
                          data-cell-id={cell.id}
                          colSpan={cell.colSpan}
                          rowSpan={cell.rowSpan}
                          style={cellStyle(table, cell, rowIndex)}
                        >
                          <div className={`cell-controls ${props.codes ? '' : 'quiet'}`}>
                            {cell.id === firstVisible && (
                              <Inspect token={row.inspector} onSelect={props.onSelect} />
                            )}
                            <Inspect token={cell.inspector} onSelect={props.onSelect} />
                          </div>
                          <BlockView {...props} blocks={cell.blocks} />
                        </td>
                      ))}
                    {after > 0 && (
                      <td className="grid-gap" colSpan={after} aria-label="Omitted grid columns" />
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}
