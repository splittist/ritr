import type { CSSProperties } from 'react';
import type { Border, Table, TableCell } from '../engine/table';

function borderStyle(border: Border | undefined): string {
  if (!border) return '1px dotted #c6cec2'; // editor guide, not a claimed document border
  if (border.style === 'none') return 'none';
  return `${Math.max(0.5, border.size / 6)}px ${border.style} ${border.color}`;
}
export function tableStyle(table: Table): CSSProperties {
  const width =
    table.width.type === 'pct' && table.width.value > 0
      ? `${table.width.value / 50}%`
      : table.width.type === 'dxa' && table.width.value > 0
        ? `${table.width.value / 15}px`
        : `${table.columns.reduce((a, b) => a + b, 0) / 15}px`;
  return {
    width,
    maxWidth: '100%',
    tableLayout: 'fixed',
    borderCollapse: 'collapse',
    marginLeft:
      table.alignment === 'center' || table.alignment === 'right'
        ? 'auto'
        : `${Math.min(table.indent / 15, 120)}px`,
    marginRight: table.alignment === 'center' ? 'auto' : undefined,
  };
}
export function cellStyle(table: Table, cell: TableCell, rowIndex: number): CSSProperties {
  const edge = (name: 'top' | 'bottom' | 'left' | 'right') => {
    const outside =
      name === 'top'
        ? rowIndex === 0
        : name === 'bottom'
          ? rowIndex + cell.rowSpan === table.rows.length
          : name === 'left'
            ? cell.column === 0
            : cell.column + cell.colSpan === table.columns.length;
    return borderStyle(
      cell.borders[name] ??
        table.borders[outside ? name : name === 'top' || name === 'bottom' ? 'insideH' : 'insideV'],
    );
  };
  return {
    padding: `${Math.min(cell.margins.top / 15, 40)}px ${Math.min(cell.margins.right / 15, 60)}px ${Math.min(cell.margins.bottom / 15, 40)}px ${Math.min(cell.margins.left / 15, 60)}px`,
    verticalAlign: cell.verticalAlign === 'center' ? 'middle' : cell.verticalAlign,
    backgroundColor: cell.shading,
    borderTop: edge('top'),
    borderBottom: edge('bottom'),
    borderLeft: edge('left'),
    borderRight: edge('right'),
  };
}
