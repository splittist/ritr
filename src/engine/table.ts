import type { Block, TokenRange } from './blocks';
import type { CodeToken, SourceBinding, Token } from './document';
import { child, isWord, wordAttr, type XmlNode } from '../package/xml';
import { value } from './styles';

export interface PreferredWidth {
  type: 'auto' | 'dxa' | 'pct';
  value: number;
}
export interface Border {
  style: 'none' | 'solid' | 'dotted' | 'dashed' | 'double';
  size: number;
  color: string;
}
export type Borders = Partial<
  Record<'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV', Border>
>;
export interface CellMargins {
  top: number;
  bottom: number;
  left: number;
  right: number;
}
export interface TableCell {
  id: string;
  source: SourceBinding;
  range: TokenRange;
  blocks: Block[];
  column: number;
  colSpan: number;
  rowSpan: number;
  merge?: 'restart' | 'continue';
  mergedInto?: string;
  width: PreferredWidth;
  margins: CellMargins;
  borders: Borders;
  verticalAlign: 'top' | 'center' | 'bottom';
  shading?: string;
  inspector: CodeToken;
}
export interface TableRow {
  id: string;
  cells: TableCell[];
  before: number;
  after: number;
  header: boolean;
  inspector: CodeToken;
}
export interface Table {
  id: string;
  source: SourceBinding;
  columns: number[];
  rows: TableRow[];
  width: PreferredWidth;
  alignment: 'left' | 'center' | 'right';
  indent: number;
  margins: CellMargins;
  borders: Borders;
  styleId?: string;
  /** Unsafe/unrepresentable geometry uses the existing linear projection. */
  linear: boolean;
  warnings: string[];
  inspector: CodeToken;
}

const on = (node: XmlNode | undefined) =>
  !!node && !['0', 'false', 'off'].includes(wordAttr(node, 'val') ?? '');
function amount(
  raw: string | undefined,
  fallback: number,
  warnings: string[],
  label: string,
  max = 31680,
): number {
  if (raw === undefined) return fallback;
  if (!/^\d+$/.test(raw) || Number(raw) > max) {
    warnings.push(`Unsupported ${label}: ${raw}`);
    return fallback;
  }
  return Number(raw);
}
function width(node: XmlNode | undefined, warnings: string[]): PreferredWidth {
  if (!node) return { type: 'auto', value: 0 };
  const type = wordAttr(node, 'type') ?? 'dxa';
  if (type === 'auto' || type === 'nil') return { type: 'auto', value: 0 };
  if (type !== 'dxa' && type !== 'pct') {
    warnings.push(`Unsupported width unit ${type}`);
    return { type: 'auto', value: 0 };
  }
  const raw = wordAttr(node, 'w');
  // Strict OOXML can encode percentages explicitly; transitional uses fiftieths.
  if (type === 'pct' && raw && /^\d+(\.\d+)?%$/.test(raw))
    return { type, value: Math.min(5000, Math.round(parseFloat(raw) * 50)) };
  return {
    type,
    value: amount(raw, 0, warnings, 'preferred width', type === 'pct' ? 5000 : 31680),
  };
}
function margins(node: XmlNode | undefined, base: CellMargins, warnings: string[]): CellMargins {
  const result = { ...base };
  if (!node) return result;
  for (const edge of ['top', 'bottom', 'left', 'right'] as const) {
    const side =
      child(node, edge === 'left' ? 'start' : edge === 'right' ? 'end' : edge) ?? child(node, edge);
    if (!side) continue;
    const unit = wordAttr(side, 'type') ?? 'dxa';
    if (unit === 'nil') result[edge] = 0;
    else if (unit === 'dxa')
      result[edge] = amount(wordAttr(side, 'w'), result[edge], warnings, 'cell margin');
    else warnings.push(`Unsupported cell margin unit ${unit}`);
  }
  return result;
}
function borders(node: XmlNode | undefined, warnings: string[]): Borders {
  const result: Borders = {};
  if (!node) return result;
  for (const edge of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'] as const) {
    const side =
      child(node, edge === 'left' ? 'start' : edge === 'right' ? 'end' : edge) ?? child(node, edge);
    if (!side) continue;
    const rawStyle = wordAttr(side, 'val') ?? 'single';
    const styles: Record<string, Border['style']> = {
      nil: 'none',
      none: 'none',
      single: 'solid',
      dotted: 'dotted',
      dashed: 'dashed',
      double: 'double',
    };
    if (!styles[rawStyle]) warnings.push(`Border ${rawStyle} displayed as a single line`);
    const rawColor = wordAttr(side, 'color');
    result[edge] = {
      style: styles[rawStyle] ?? 'solid',
      size: amount(wordAttr(side, 'sz'), 4, warnings, 'border size', 96),
      color: rawColor && /^[0-9a-f]{6}$/i.test(rawColor) ? `#${rawColor}` : '#73877c',
    };
    if (wordAttr(side, 'themeColor')) warnings.push('Theme border colors are approximated');
  }
  return result;
}

export function readTable(
  node: XmlNode,
  tokens: Token[],
  ranges: Map<string, TokenRange>,
  blocks: (range: TokenRange) => Block[],
  depth: number,
): Table {
  const warnings: string[] = [];
  const tokenFor = (n: XmlNode, label: string): CodeToken => {
    const original = tokens[ranges.get(n.id)!.from] as CodeToken;
    return { ...original, label, details: { ...original.details } };
  };
  const props = child(node, 'tblPr');
  const table: Table = {
    id: node.id,
    source: tokenFor(node, 'Table').source,
    columns: [],
    rows: [],
    width: width(props && child(props, 'tblW'), warnings),
    alignment: 'left',
    indent: 0,
    margins: margins(
      props && child(props, 'tblCellMar'),
      { top: 0, bottom: 0, left: 108, right: 108 },
      warnings,
    ),
    borders: borders(props && child(props, 'tblBorders'), warnings),
    styleId: value(props, 'tblStyle'),
    linear: false,
    warnings,
    inspector: tokenFor(node, 'Table'),
  };
  const fallback = (message: string) => {
    warnings.push(message);
    table.linear = true;
  };
  const geometry = (raw: string | undefined, defaultValue: number, label: string) => {
    if (
      raw !== undefined &&
      (!/^\d+$/.test(raw) || Number(raw) > 128 || (label === 'gridSpan' && Number(raw) === 0))
    )
      fallback(`Invalid ${label}; shown linearly`);
    return amount(raw, defaultValue, warnings, label, 128);
  };
  const unsupported = (parent: XmlNode, allowed: string[]) =>
    parent.children.some((n) => !isWord(n) || !allowed.includes(n.local));
  if (depth >= 6 || node.children.filter((n) => isWord(n, 'tr')).length > 500) {
    fallback('Table exceeds display depth/row limits; shown as linear codes');
    return table;
  }
  if (unsupported(node, ['tblPr', 'tblGrid', 'tr']))
    fallback('Wrapped or unfamiliar table rows are shown linearly');
  if (table.styleId)
    warnings.push(
      `Table style ${table.styleId} is preserved; conditional style formatting is not resolved`,
    );
  if (props && child(props, 'tblpPr'))
    warnings.push('Floating table placement is approximated inline');
  if (on(props && child(props, 'bidiVisual')))
    fallback('Right-to-left table geometry is shown linearly');
  const align = value(props, 'jc');
  if (align && ['left', 'center', 'right'].includes(align))
    table.alignment = align as Table['alignment'];
  const ind = props && child(props, 'tblInd');
  if (ind) table.indent = width(ind, warnings).value;
  const grid = child(node, 'tblGrid');
  table.columns = (grid?.children.filter((n) => isWord(n, 'gridCol')) ?? []).map((n) =>
    amount(wordAttr(n, 'w'), 0, warnings, 'grid width'),
  );
  if (table.columns.length > 128) {
    fallback('Table exceeds 128 display columns; shown linearly');
    table.columns = [];
    return table;
  }
  let active = new Map<number, TableCell>();
  for (const rowNode of node.children.filter((n) => isWord(n, 'tr'))) {
    const rowProps = child(rowNode, 'trPr');
    const row: TableRow = {
      id: rowNode.id,
      cells: [],
      before: geometry(value(rowProps, 'gridBefore'), 0, 'gridBefore'),
      after: geometry(value(rowProps, 'gridAfter'), 0, 'gridAfter'),
      header: on(rowProps && child(rowProps, 'tblHeader')),
      inspector: tokenFor(rowNode, `Row ${table.rows.length + 1}`),
    };
    if (unsupported(rowNode, ['trPr', 'tc', 'tblPrEx']))
      fallback('Wrapped or unfamiliar row cells are shown linearly');
    if (child(rowNode, 'tblPrEx'))
      warnings.push('Row-specific table property exceptions are not resolved');
    let column = row.before;
    const nextActive = new Map<number, TableCell>();
    for (const cellNode of rowNode.children.filter((n) => isWord(n, 'tc'))) {
      const cellProps = child(cellNode, 'tcPr');
      const range = ranges.get(cellNode.id)!;
      const span = geometry(value(cellProps, 'gridSpan'), 1, 'gridSpan');
      if (span === 0) fallback('Zero-width cell span is shown linearly');
      const mergeNode = cellProps && child(cellProps, 'vMerge');
      const merge = mergeNode ? (wordAttr(mergeNode, 'val') ?? 'continue') : undefined;
      if (merge && !['restart', 'continue'].includes(merge))
        fallback('Invalid vertical merge is shown linearly');
      if (cellProps && child(cellProps, 'hMerge'))
        fallback('Legacy hMerge geometry is shown linearly; gridSpan is supported');
      const alignment = value(cellProps, 'vAlign') ?? 'top';
      const shade = cellProps && child(cellProps, 'shd');
      const fill = shade && wordAttr(shade, 'fill');
      const cell: TableCell = {
        id: cellNode.id,
        source: tokenFor(cellNode, 'Cell').source,
        range,
        blocks: blocks(range),
        column,
        colSpan: Math.max(1, span),
        rowSpan: 1,
        merge: merge as TableCell['merge'],
        width: width(cellProps && child(cellProps, 'tcW'), warnings),
        margins: margins(cellProps && child(cellProps, 'tcMar'), table.margins, warnings),
        borders: borders(cellProps && child(cellProps, 'tcBorders'), warnings),
        verticalAlign: ['top', 'center', 'bottom'].includes(alignment)
          ? (alignment as TableCell['verticalAlign'])
          : 'top',
        shading: fill && /^[0-9a-f]{6}$/i.test(fill) ? `#${fill}` : undefined,
        inspector: tokenFor(cellNode, `Cell ${table.rows.length + 1}, ${column + 1}`),
      };
      if (value(cellProps, 'textDirection') && value(cellProps, 'textDirection') !== 'lrTb')
        warnings.push('Rotated cell text is displayed horizontally');
      if (merge === 'continue') {
        const origin = active.get(column);
        if (!origin || origin.colSpan !== cell.colSpan)
          fallback('Vertical merge has no matching cell above; shown linearly');
        else {
          const hiddenContent = tokens
            .slice(range.from + 1, range.to - 1)
            .some((t) =>
              t.kind === 'text'
                ? t.span.text.length > 0
                : t.category === 'opaque' ||
                  t.category === 'review' ||
                  t.role === 'list-label' ||
                  ['tab', 'br', 'cr', 'tbl'].includes(t.label),
            );
          if (hiddenContent)
            fallback(
              'A vertical-merge continuation contains content; shown linearly to keep it visible',
            );
          origin.rowSpan++;
          cell.mergedInto = origin.id;
          nextActive.set(column, origin);
        }
      } else if (merge === 'restart') nextActive.set(column, cell);
      row.cells.push(cell);
      column += cell.colSpan;
    }
    const count = column + row.after;
    if (count > 128) fallback('Row exceeds 128 display columns; shown linearly');
    else while (table.columns.length < count) table.columns.push(0);
    table.rows.push(row);
    active = nextActive;
  }
  if (!table.rows.length || !table.columns.length)
    fallback('Empty or unresolved table grid is shown linearly');
  // Infer only missing grid widths; explicit grid columns remain authoritative.
  for (const row of table.rows)
    for (const cell of row.cells) {
      if (cell.width.type === 'dxa')
        for (
          let i = cell.column;
          i < Math.min(cell.column + cell.colSpan, table.columns.length);
          i++
        )
          if (!table.columns[i]) table.columns[i] = cell.width.value / cell.colSpan;
    }
  table.columns = table.columns.map((w) => w || 1440);
  table.warnings = [...new Set(warnings)];
  table.inspector.details.table = {
    columns: table.columns,
    width: table.width,
    alignment: table.alignment,
    indent: table.indent,
    styleId: table.styleId,
    margins: table.margins,
    borders: table.borders,
    linear: table.linear,
    warnings: table.warnings,
  };
  for (const row of table.rows) {
    row.inspector.details.row = {
      before: row.before,
      after: row.after,
      header: row.header,
      cells: row.cells.map((c) => c.id),
    };
    for (const cell of row.cells)
      cell.inspector.details.cell = {
        column: cell.column,
        colSpan: cell.colSpan,
        rowSpan: cell.rowSpan,
        mergedInto: cell.mergedInto,
        width: cell.width,
        margins: cell.margins,
        borders: cell.borders,
        verticalAlign: cell.verticalAlign,
        shading: cell.shading,
      };
  }
  return table;
}
