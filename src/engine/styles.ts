import type { DocxPackage } from '../package/docx';
import { child, isWord, wordAttr, type XmlNode } from '../package/xml';

export interface PropertyLayer {
  name: string;
  node: XmlNode;
}
export const value = (node: XmlNode | undefined, name: string): string | undefined => {
  const element = node && child(node, name);
  return element && wordAttr(element, 'val');
};

/** A shared style index. Chains are returned base-first, with cycles diagnosed. */
export class Styles {
  readonly root?: XmlNode;
  readonly entries = new Map<string, XmlNode>();
  readonly defaultParagraph?: string;
  constructor(pkg: DocxPackage) {
    const rel = pkg.relationships.find(
      (r) => r.source === pkg.mainPart && r.type.endsWith('/styles') && !r.external,
    );
    if (rel && pkg.names().includes(rel.target)) this.root = pkg.xml(rel.target);
    for (const node of this.root?.children ?? []) {
      const id = wordAttr(node, 'styleId');
      if (isWord(node, 'style') && id) this.entries.set(id, node);
    }
    this.defaultParagraph = [...this.entries].find(
      ([, n]) =>
        wordAttr(n, 'type') === 'paragraph' &&
        ['1', 'true', 'on'].includes(wordAttr(n, 'default') ?? ''),
    )?.[0];
  }
  chain(id: string | undefined, warnings: string[]): XmlNode[] {
    const result: XmlNode[] = [],
      seen = new Set<string>();
    while (id) {
      if (seen.has(id)) {
        warnings.push(`Style inheritance cycle at ${id}`);
        break;
      }
      seen.add(id);
      const style = this.entries.get(id);
      if (!style) {
        warnings.push(`Missing style ${id}`);
        break;
      }
      result.unshift(style);
      id = value(style, 'basedOn');
    }
    return result;
  }
  defaults(kind: 'pPr' | 'rPr'): XmlNode | undefined {
    const defaults = this.root && child(this.root, 'docDefaults');
    const wrapper = defaults && child(defaults, `${kind}Default`);
    return wrapper && child(wrapper, kind);
  }
  paragraphLayers(styleId: string | undefined, warnings: string[]): PropertyLayer[] {
    return this.chain(styleId ?? this.defaultParagraph, warnings).flatMap((style) => {
      const node = child(style, 'pPr');
      return node ? [{ name: `Style ${wordAttr(style, 'styleId')}`, node }] : [];
    });
  }
}
