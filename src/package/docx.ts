import { attr, child, isWord, parseXml, type XmlNode } from './xml';
import { equalBytes, readZip, writeZip } from './zip';

const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const decoder = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true });
const encoder = new TextEncoder();
export interface Diagnostic {
  severity: 'info' | 'warning' | 'error';
  code: string;
  part?: string;
  message: string;
}
export interface Relationship {
  source: string;
  id: string;
  type: string;
  target: string;
  external: boolean;
}
export interface PartDifference {
  part: string;
  status: 'unchanged' | 'changed' | 'added' | 'removed';
  before: number;
  after: number;
}

function resolveTarget(source: string, target: string): string {
  const decoded = decodeURIComponent(target.split('#')[0]!);
  if (decoded.includes('\\') || decoded.includes(':'))
    throw new Error(`Invalid relationship target: ${target}`);
  const pieces = decoded.startsWith('/') ? [] : source.split('/').slice(0, -1);
  for (const segment of decoded.split('/')) {
    if (segment === '..') {
      if (!pieces.length) throw new Error(`Relationship escapes package: ${target}`);
      pieces.pop();
    } else if (segment && segment !== '.') pieces.push(segment);
  }
  return pieces.join('/');
}

/** Immutable package snapshot. Byte buffers are never exposed without a copy. */
export class DocxPackage {
  readonly diagnostics: readonly Diagnostic[];
  readonly relationships: readonly Relationship[];
  readonly mainPart: string;
  private constructor(
    private readonly parts: Map<string, Uint8Array>,
    private readonly original?: Uint8Array,
  ) {
    const diagnostics: Diagnostic[] = [];
    const types = this.xml('[Content_Types].xml');
    if (types.local !== 'Types' || types.uri !== CT) throw new Error('Invalid content types root');
    const defaults = new Map<string, string>();
    const overrides = new Map<string, string>();
    for (const node of types.children) {
      if (node.uri !== CT) continue;
      const key = attr(node, node.local === 'Default' ? 'Extension' : 'PartName');
      const value = attr(node, 'ContentType');
      if (!key || !value) throw new Error('Invalid content type declaration');
      const map = node.local === 'Default' ? defaults : overrides;
      if (map.has(key)) throw new Error(`Duplicate content type: ${key}`);
      map.set(key, value);
    }
    const relationships: Relationship[] = [];
    for (const name of this.names()) {
      if (name.endsWith('/') || name === '[Content_Types].xml') continue;
      if (!overrides.has(`/${name}`) && !defaults.has(name.split('.').at(-1)!))
        diagnostics.push({
          severity: 'error',
          code: 'missing-content-type',
          part: name,
          message: 'No content type declaration',
        });
      if (name.endsWith('.xml') || name.endsWith('.rels')) this.xml(name);
      if (!name.endsWith('.rels')) continue;
      const source =
        name === '_rels/.rels' ? '' : name.replace(/(^|\/)\_rels\//, '$1').replace(/\.rels$/, '');
      if (source && !parts.has(source))
        diagnostics.push({
          severity: 'error',
          code: 'missing-relationship-source',
          part: name,
          message: `Missing source ${source}`,
        });
      const root = this.xml(name);
      if (root.uri !== REL || root.local !== 'Relationships')
        throw new Error(`Invalid relationships root: ${name}`);
      const ids = new Set<string>();
      for (const node of root.children) {
        if (node.uri !== REL || node.local !== 'Relationship')
          throw new Error(`Invalid relationship element: ${name}`);
        const id = attr(node, 'Id'),
          type = attr(node, 'Type'),
          target = attr(node, 'Target');
        if (!id || !type || !target || ids.has(id))
          throw new Error(`Invalid or duplicate relationship in ${name}`);
        ids.add(id);
        const external = attr(node, 'TargetMode') === 'External';
        const resolved = external ? target : resolveTarget(source, target);
        if (!external && !parts.has(resolved))
          diagnostics.push({
            severity: 'error',
            code: 'missing-target',
            part: name,
            message: `${id} targets missing ${resolved}`,
          });
        relationships.push({ source, id, type, target: resolved, external });
      }
    }
    const main = relationships.filter(
      (r) => !r.source && r.type.endsWith('/officeDocument') && !r.external,
    );
    if (main.length !== 1 || !parts.has(main[0]!.target))
      throw new Error('Expected one internal officeDocument relationship');
    this.mainPart = main[0]!.target;
    const document = this.xml(this.mainPart);
    if (!isWord(document, 'document') || !child(document, 'body'))
      throw new Error('Main part must contain a WordprocessingML document and body');
    if (this.names().some((n) => n.startsWith('_xmlsignatures/')))
      diagnostics.push({
        severity: 'warning',
        code: 'signed-package',
        message:
          'Digital signatures are preserved; editing is disabled because it would invalidate them.',
      });
    this.relationships = Object.freeze(relationships.map((r) => Object.freeze(r)));
    this.diagnostics = Object.freeze(diagnostics.map((d) => Object.freeze(d)));
  }
  static open(bytes: Uint8Array): DocxPackage {
    // Buffer.slice() aliases memory; normalize to a real owned Uint8Array first.
    const owned = new Uint8Array(bytes);
    return new DocxPackage(readZip(owned), owned);
  }
  names(): string[] {
    return [...this.parts.keys()];
  }
  bytes(name: string): Uint8Array {
    const value = this.parts.get(name);
    if (!value) throw new Error(`Missing package part: ${name}`);
    return value.slice();
  }
  text(name: string): string {
    const text = decoder.decode(this.bytes(name));
    const encoding = text.match(/<\?xml[^?]*encoding\s*=\s*["']([^"']+)/i)?.[1];
    if (encoding && !/^utf-?8$/i.test(encoding))
      throw new Error(`${name}: only UTF-8 XML is supported`);
    return text;
  }
  xml(name: string): XmlNode {
    return parseXml(this.text(name), name);
  }
  withXml(changes: ReadonlyMap<string, string>): DocxPackage {
    const parts = new Map(this.parts);
    for (const [name, value] of changes) {
      if (!parts.has(name)) throw new Error(`Cannot patch nonexistent part: ${name}`);
      parts.set(name, encoder.encode(value));
    }
    return new DocxPackage(parts);
  }
  save(): Uint8Array {
    const errors = this.diagnostics.filter((d) => d.severity === 'error');
    if (errors.length)
      throw new Error(`Package validation failed: ${errors.map((d) => d.message).join('; ')}`);
    return this.original?.slice() ?? writeZip(this.parts);
  }
}

export function comparePackages(before: DocxPackage, after: DocxPackage): PartDifference[] {
  const left = new Set(before.names()),
    right = new Set(after.names());
  return [...new Set([...left, ...right])].sort().map((part) => {
    const a = left.has(part) ? before.bytes(part) : undefined;
    const b = right.has(part) ? after.bytes(part) : undefined;
    return {
      part,
      status: !a ? 'added' : !b ? 'removed' : equalBytes(a, b) ? 'unchanged' : 'changed',
      before: a?.length ?? 0,
      after: b?.length ?? 0,
    };
  });
}
