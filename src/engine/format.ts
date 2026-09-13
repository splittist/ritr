import type { TextSpan } from './document';
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

export interface TextFormat {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  color?: string;
  highlight?: string;
}
export function validateFormat(format: TextFormat) {
  if (!format || typeof format !== 'object' || Array.isArray(format) || !Object.keys(format).length)
    throw new Error('Choose a text format');
  for (const [key, value] of Object.entries(format)) {
    if (['b', 'i', 'u'].includes(key) && typeof value === 'boolean') continue;
    if (key === 'color' && typeof value === 'string' && /^(?:[0-9a-f]{6}|auto)$/i.test(value))
      continue;
    if (
      key === 'highlight' &&
      typeof value === 'string' &&
      (value === 'none' || Object.hasOwn(highlightColors, value))
    )
      continue;
    throw new Error(`Invalid text format: ${key}`);
  }
}
export function effectiveFormat(span?: Pick<TextSpan, 'direct' | 'inherited'>): TextFormat {
  const f = { ...span?.inherited, ...span?.direct };
  const on = (v?: string) => ['on', 'true', '1'].includes(v ?? '');
  return {
    b: on(f.b),
    i: on(f.i),
    u: !!f.u && !['none', 'off', '0', 'false'].includes(f.u),
    color: f.color ?? 'auto',
    highlight: f.highlight ?? 'none',
  };
}
