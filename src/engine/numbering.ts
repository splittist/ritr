import type { DocxPackage } from '../package/docx';
import { child, isWord, wordAttr, type XmlNode } from '../package/xml';
import { Styles, value } from './styles';

export interface NumberingLevel {
  index: number;
  start: number;
  instanceStart?: boolean;
  format: string;
  template: string;
  /** One-based restart threshold; 0 means never. Word also resets at higher levels. */
  restart: number;
  paragraphStyle?: string;
  suffix: 'tab' | 'space' | 'nothing';
  alignment: string;
  legal: boolean;
  font?: string;
  node: XmlNode;
  warnings: string[];
}
export interface ListDefinition {
  numId: string;
  abstractId: string;
  counterId: string;
  levels: Map<number, NumberingLevel>;
  warnings: string[];
}
export interface ListLabel {
  numId: string;
  abstractId?: string;
  level: number;
  text: string;
  format?: string;
  template?: string;
  suffix: 'tab' | 'space' | 'nothing';
  alignment: string;
  font?: string;
  definitionId?: string;
  definitionXml?: string;
  warnings: string[];
}

function integer(
  raw: string | undefined,
  fallback: number,
  warnings: string[],
  name: string,
): number {
  if (raw === undefined) return fallback;
  if (!/^\d+$/.test(raw) || !Number.isSafeInteger(Number(raw)) || Number(raw) > 1_000_000) {
    warnings.push(`Unsupported ${name}: ${raw}`);
    return fallback;
  }
  return Number(raw);
}

/** Package numbering definitions are immutable; counters live in a separate story session. */
export class Numbering {
  private readonly abstract = new Map<string, XmlNode>();
  private readonly instances = new Map<string, XmlNode>();
  private readonly cache = new Map<string, ListDefinition>();
  readonly source: string;
  constructor(
    pkg: DocxPackage,
    private readonly styles: Styles,
  ) {
    const rel = pkg.relationships.find(
      (r) => r.source === pkg.mainPart && r.type.endsWith('/numbering') && !r.external,
    );
    this.source = rel && pkg.names().includes(rel.target) ? pkg.text(rel.target) : '';
    if (!this.source || !rel) return;
    for (const n of pkg.xml(rel.target).children) {
      if (isWord(n, 'abstractNum')) this.abstract.set(wordAttr(n, 'abstractNumId') ?? '', n);
      if (isWord(n, 'num')) this.instances.set(wordAttr(n, 'numId') ?? '', n);
    }
  }
  definition(numId: string, seen = new Set<string>()): ListDefinition {
    const cached = this.cache.get(numId);
    if (cached) return cached;
    const result: ListDefinition = {
      numId,
      abstractId: '',
      counterId: '',
      levels: new Map(),
      warnings: [],
    };
    if (seen.has(numId)) {
      result.warnings.push('Numbering style link cycle');
      return result;
    }
    seen.add(numId);
    const instance = this.instances.get(numId);
    const abstractId = value(instance, 'abstractNumId');
    const abstract = abstractId !== undefined ? this.abstract.get(abstractId) : undefined;
    if (!instance || !abstract) {
      result.warnings.push(`Missing numbering definition for list ${numId}`);
      return result;
    }
    result.abstractId = abstractId!;
    result.counterId = abstractId!;
    const styleLink = value(abstract, 'numStyleLink');
    if (styleLink) {
      let linkedId: string | undefined;
      for (const style of this.styles.chain(styleLink, result.warnings)) {
        const pPr = child(style, 'pPr');
        linkedId = value(pPr && child(pPr, 'numPr'), 'numId') ?? linkedId;
      }
      if (linkedId) {
        const linked = this.definition(linkedId, seen);
        result.levels = new Map(linked.levels);
        result.counterId = linked.counterId;
        result.warnings.push(...linked.warnings);
      } else result.warnings.push(`Missing numbering style target ${styleLink}`);
    }
    for (const node of abstract.children.filter((n) => isWord(n, 'lvl'))) {
      const level = this.readLevel(node);
      if (level.index > 8) {
        result.warnings.push('Numbering levels beyond 9 are unsupported');
        continue;
      }
      result.levels.set(level.index, level);
    }
    for (const override of instance.children.filter((n) => isWord(n, 'lvlOverride'))) {
      const index = integer(wordAttr(override, 'ilvl'), 0, result.warnings, 'override level');
      const replacement = child(override, 'lvl');
      const base = result.levels.get(index);
      const level = replacement
        ? this.readLevel(replacement)
        : base && { ...base, warnings: [...base.warnings] };
      if (!level) {
        result.warnings.push(`Missing overridden level ${index}`);
        continue;
      }
      // Word ignores lvlRestart inside an override (MS-OI29500 §2.1.282).
      if (replacement) level.restart = base?.restart ?? index;
      level.start = integer(
        value(override, 'startOverride'),
        level.start,
        level.warnings,
        'start override',
      );
      level.instanceStart =
        value(override, 'startOverride') !== undefined ||
        (replacement !== undefined && value(replacement, 'start') !== undefined);
      result.levels.set(index, level);
    }
    if (
      Object.entries(abstract.attrs).some(
        ([key, v]) =>
          key.endsWith('}restartNumberingAfterBreak') && ['1', 'true', 'on'].includes(v),
      )
    )
      result.warnings.push('Section-break numbering restarts are not yet resolved');
    this.cache.set(numId, result);
    return result;
  }
  private readLevel(node: XmlNode): NumberingLevel {
    const warnings: string[] = [];
    const index = integer(wordAttr(node, 'ilvl'), 0, warnings, 'level');
    let restart = integer(value(node, 'lvlRestart'), index, warnings, 'restart');
    if (restart > index) restart = index;
    const format = value(node, 'numFmt') ?? 'decimal';
    const template = value(node, 'lvlText') ?? '';
    if (
      ![
        'decimal',
        'decimalZero',
        'upperLetter',
        'lowerLetter',
        'upperRoman',
        'lowerRoman',
        'bullet',
        'none',
      ].includes(format)
    )
      warnings.push(`Unsupported numbering format ${format}`);
    if (child(node, 'lvlPicBulletId')) warnings.push('Picture bullets are not rendered');
    if (
      /%[1-9]/.test(template) &&
      [...template.matchAll(/%([1-9])/g)].some((m) => Number(m[1]) > index + 1)
    )
      warnings.push('Label references a deeper numbering level');
    const rPr = child(node, 'rPr'),
      fonts = rPr && child(rPr, 'rFonts');
    const font = fonts && (wordAttr(fonts, 'ascii') ?? wordAttr(fonts, 'hAnsi'));
    const legalNode = child(node, 'isLgl');
    const suffix = value(node, 'suff') ?? 'tab';
    if (!['tab', 'space', 'nothing'].includes(suffix))
      warnings.push(`Unsupported numbering suffix ${suffix}`);
    return {
      index,
      start: integer(value(node, 'start'), 0, warnings, 'start'),
      format,
      template,
      restart,
      paragraphStyle: value(node, 'pStyle'),
      suffix: ['space', 'nothing'].includes(suffix) ? (suffix as 'space' | 'nothing') : 'tab',
      alignment: value(node, 'lvlJc') ?? 'left',
      legal: !!legalNode && !['0', 'false', 'off'].includes(wordAttr(legalNode, 'val') ?? ''),
      font,
      node,
      warnings,
    };
  }
}

export function formatCounter(number: number, format: string): string | undefined {
  if (!Number.isSafeInteger(number) || number < 0) return undefined;
  if (format === 'decimal') return String(number);
  if (format === 'decimalZero') return String(number).padStart(2, '0');
  if (format === 'none') return '';
  if (format === 'upperLetter' || format === 'lowerLetter') {
    if (number < 1 || number > 780) return undefined;
    // Word's alphabetic numbering repeats letters after Z: AA, BB, ... ZZ.
    const letter = String.fromCharCode(65 + ((number - 1) % 26)).repeat(
      Math.floor((number - 1) / 26) + 1,
    );
    return format === 'lowerLetter' ? letter.toLowerCase() : letter;
  }
  if (format === 'upperRoman' || format === 'lowerRoman') {
    if (number < 1 || number > 3999) return undefined;
    let result = '';
    for (const [amount, symbol] of [
      [1000, 'M'],
      [900, 'CM'],
      [500, 'D'],
      [400, 'CD'],
      [100, 'C'],
      [90, 'XC'],
      [50, 'L'],
      [40, 'XL'],
      [10, 'X'],
      [9, 'IX'],
      [5, 'V'],
      [4, 'IV'],
      [1, 'I'],
    ] as const) {
      while (number >= amount) {
        result += symbol;
        number -= amount;
      }
    }
    return format === 'lowerRoman' ? result.toLowerCase() : result;
  }
  return undefined;
}

/** A fresh counter set for each story; unnumbered paragraphs do not reset a list. */
export class NumberingSequence {
  private counters = new Map<string, Map<number, number>>();
  private visited = new Set<string>();
  constructor(private readonly numbering: Numbering) {}
  next(numId: string, index: number, uncertainty: string[] = []): ListLabel {
    const definition = this.numbering.definition(numId);
    const level = definition.levels.get(index);
    const result: ListLabel = {
      numId,
      abstractId: definition.abstractId,
      level: index,
      text: '?',
      suffix: level?.suffix ?? 'tab',
      alignment: level?.alignment ?? 'left',
      warnings: [...definition.warnings, ...(level?.warnings ?? []), ...uncertainty],
    };
    if (!level) {
      result.warnings.push(`Missing level ${index} in list ${numId}`);
      return result;
    }
    Object.assign(result, {
      format: level.format,
      template: level.template,
      font: level.font,
      definitionId: level.node.id,
      definitionXml: this.numbering.source.slice(level.node.start, level.node.end),
    });
    // Word shares counters across instances of the same abstract definition.
    // A start override resets that stream once when this instance is first used.
    const counters = this.counters.get(definition.counterId) ?? new Map<number, number>();
    const visit = `${numId}:${index}`;
    const startsInstance = level.instanceStart && !this.visited.has(visit);
    counters.set(
      index,
      !startsInstance && counters.has(index) ? counters.get(index)! + 1 : level.start,
    );
    this.visited.add(visit);
    for (const [other, spec] of definition.levels)
      if (other > index && spec.restart > 0 && index < spec.restart) counters.delete(other);
    this.counters.set(definition.counterId, counters);
    if (level.format === 'bullet') {
      result.text = level.template;
      // Common legacy bullet glyphs use font-private codepoints. Normalize known
      // ones for portable display while retaining font/template in the inspector.
      if (level.font?.toLowerCase() === 'symbol' && level.template === '\uf0b7') result.text = '•';
      else if (level.font?.toLowerCase() === 'courier new' && level.template === 'o')
        result.text = '◦';
      else if (level.font?.toLowerCase() === 'wingdings' && level.template === '\uf0a7')
        result.text = '▪';
      else if (/[\ue000-\uf8ff]/u.test(level.template))
        result.warnings.push('Unsupported font-specific bullet glyph');
    } else if (level.format === 'none') result.text = '';
    else
      result.text = level.template.replace(/%([1-9])/g, (_match, digit: string) => {
        const referenced = Number(digit) - 1,
          spec = definition.levels.get(referenced);
        const formatted =
          spec &&
          formatCounter(
            counters.get(referenced) ?? spec.start,
            level.legal ? 'decimal' : spec.format,
          );
        if (formatted === undefined)
          result.warnings.push(`Cannot format level ${referenced} counter`);
        return formatted ?? '?';
      });
    if (result.warnings.length) result.text = '?';
    return result;
  }
}
