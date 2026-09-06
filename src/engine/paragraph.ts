import { child, wordAttr, type XmlNode } from '../package/xml';
import type { SourceBinding } from './document';
import { Numbering, NumberingSequence, type ListLabel } from './numbering';
import { Styles, value, type PropertyLayer } from './styles';

export interface ParagraphLayout {
  /** Measurements are source twips (1/20 point); firstLine is relative to left. */
  left: number;
  right: number;
  firstLine: number;
  sources: Record<string, string>;
}
export interface Paragraph {
  id: string;
  source: SourceBinding;
  styleId?: string;
  layout: ParagraphLayout;
  numbering?: ListLabel;
  propertySources: Record<string, string>;
  warnings: string[];
}

/** Resolve individual indentation attributes, retaining the winning source. */
export function resolveIndentation(layers: PropertyLayer[], warnings: string[]): ParagraphLayout {
  const attrs: Record<string, string> = {},
    sources: Record<string, string> = {};
  for (const { name, node } of layers) {
    const indent = child(node, 'ind');
    if (!indent) continue;
    // First-line and hanging are two encodings of one signed offset. A later
    // layer replaces that offset; within one element, hanging takes precedence.
    if (
      ['firstLine', 'hanging', 'firstLineChars', 'hangingChars'].some(
        (key) => wordAttr(indent, key) !== undefined,
      )
    ) {
      for (const key of ['firstLine', 'hanging', 'firstLineChars', 'hangingChars']) {
        delete attrs[key];
        delete sources[key];
      }
    }
    for (const key of [
      'left',
      'right',
      'start',
      'end',
      'firstLine',
      'hanging',
      'leftChars',
      'rightChars',
      'startChars',
      'endChars',
      'firstLineChars',
      'hangingChars',
    ]) {
      const v = wordAttr(indent, key);
      if (v !== undefined) {
        attrs[key] = v;
        sources[key] = name;
      }
    }
  }
  const measure = (key: string) => {
    if (attrs[key] === undefined) return 0;
    const n = Number(attrs[key]);
    if (!/^-?\d+$/.test(attrs[key]!) || !Number.isSafeInteger(n) || Math.abs(n) > 31680) {
      warnings.push(`Unsupported indentation ${key}=${attrs[key]}`);
      return 0;
    }
    return n;
  };
  if (Object.keys(attrs).some((k) => k.endsWith('Chars')))
    warnings.push('Character-unit indentation is approximated using available twip values');
  // Logical start/end supersede physical left/right in left-to-right paragraphs.
  const left = measure(attrs.start !== undefined ? 'start' : 'left');
  const right = measure(attrs.end !== undefined ? 'end' : 'right');
  const firstLine = (attrs.hanging !== undefined ? -measure('hanging') : measure('firstLine')) || 0;
  return { left, right, firstLine, sources };
}

/** Stateless property resolution plus one counter sequence per story. */
export class Paragraphs {
  private readonly sequence: NumberingSequence;
  constructor(
    private readonly styles: Styles,
    private readonly numbering: Numbering,
  ) {
    this.sequence = new NumberingSequence(numbering);
  }
  read(node: XmlNode, part: string, uncertain: string[] = []): Paragraph {
    const warnings: string[] = [];
    const direct = child(node, 'pPr');
    const styleId = value(direct, 'pStyle') ?? this.styles.defaultParagraph;
    const inherited = this.styles.paragraphLayers(styleId, warnings);
    const propertySources: Record<string, string> = {};
    let numId: string | undefined, index: number | undefined;
    const membership = [
      ...inherited,
      ...(direct ? [{ name: 'Direct paragraph', node: direct }] : []),
    ];
    for (const layer of membership) {
      const numPr = child(layer.node, 'numPr');
      const id = value(numPr, 'numId');
      if (id !== undefined) {
        numId = id;
        propertySources.numId = layer.name;
      }
      // ilvl in a paragraph style is ignored: the numbering level's pStyle binds it.
      const level = layer.node === direct ? value(numPr, 'ilvl') : undefined;
      if (level !== undefined) {
        index = Number(level);
        propertySources.level = layer.name;
      }
    }
    let list: ListLabel | undefined;
    let levelLayer: PropertyLayer[] = [];
    if (numId !== undefined && numId !== '0') {
      const definition = this.numbering.definition(numId);
      if (index === undefined && propertySources.numId !== 'Direct paragraph') {
        const chain = this.styles.chain(styleId, []).reverse();
        for (const style of chain) {
          const match = [...definition.levels.values()].find(
            (l) => l.paragraphStyle === wordAttr(style, 'styleId'),
          );
          if (match) {
            index = match.index;
            propertySources.level = `Numbering level linked to ${wordAttr(style, 'styleId')}`;
            break;
          }
        }
      }
      index ??= 0;
      const issues = [...uncertain, ...warnings];
      if (!Number.isInteger(index) || index < 0 || index > 8)
        issues.push('Invalid numbering level');
      list = this.sequence.next(numId, index, issues);
      const level = definition.levels.get(index);
      const pPr = level && child(level.node, 'pPr');
      if (pPr) levelLayer = [{ name: `List ${numId}, level ${index + 1}`, node: pPr }];
      warnings.push(...list.warnings);
    }
    const defaults = this.styles.defaults('pPr');
    const layers = [
      ...(defaults ? [{ name: 'Document defaults', node: defaults }] : []),
      ...levelLayer,
      ...inherited,
      ...(direct ? [{ name: 'Direct paragraph', node: direct }] : []),
    ];
    for (const layer of layers)
      if (
        child(layer.node, 'bidi') &&
        !['0', 'false', 'off'].includes(value(layer.node, 'bidi') ?? '')
      )
        warnings.push('Right-to-left paragraph layout is not yet resolved');
    return {
      id: node.id,
      source: { part, nodeId: node.id, start: node.start, end: node.end },
      styleId,
      numbering: list,
      layout: resolveIndentation(layers, warnings),
      propertySources,
      warnings: [...new Set(warnings)],
    };
  }
}
