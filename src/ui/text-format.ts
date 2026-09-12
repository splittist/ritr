import type { TextSpan } from '../engine/document';

/** ST_Highlight is a fixed palette, distinct from arbitrary run shading (w:shd). */
export const highlightColors: Readonly<Record<string, string>> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00ffff',
  green: '#00ff00',
  magenta: '#ff00ff',
  red: '#ff0000',
  yellow: '#ffff00',
  white: '#ffffff',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#c0c0c0',
};
const on = (value?: string) => value === 'on' || value === 'true' || value === '1';
const underlines: Readonly<Record<string, string>> = {
  on: 'solid',
  single: 'solid',
  words: 'solid',
  thick: 'solid',
  double: 'double',
  dotted: 'dotted',
  dottedHeavy: 'dotted',
  dash: 'dashed',
  dashedHeavy: 'dashed',
  dashLong: 'dashed',
  dashLongHeavy: 'dashed',
  dotDash: 'dashed',
  dashDotHeavy: 'dashed',
  dotDotDash: 'dashed',
  dashDotDotHeavy: 'dashed',
  wave: 'wavy',
  wavyHeavy: 'wavy',
  wavyDouble: 'wavy',
};

/** Only allowlisted CSS values enter the projection; source fonts and sizes are ignored. */
export function textStyle(span: Pick<TextSpan, 'direct' | 'inherited'>): string {
  const format = { ...span.inherited, ...span.direct };
  const underline = Object.hasOwn(underlines, format.u ?? '') ? underlines[format.u!] : undefined;
  const color = /^[0-9a-f]{6}$/i.test(format.color ?? '') ? `#${format.color}` : 'inherit';
  const highlight = Object.hasOwn(highlightColors, format.highlight ?? '')
    ? highlightColors[format.highlight!]
    : 'transparent';
  return [
    `font-weight:${on(format.b) ? '700' : '400'}`,
    `font-style:${on(format.i) ? 'italic' : 'normal'}`,
    `text-decoration-line:${underline ? 'underline' : 'none'}`,
    `text-decoration-style:${underline ?? 'solid'}`,
    `color:${color}`,
    `background-color:${highlight}`,
  ].join(';');
}
