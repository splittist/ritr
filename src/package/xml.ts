import { SaxesParser } from 'saxes';

/** Source offsets are UTF-16 string offsets, never byte offsets or XPath indexes. */
export interface XmlNode {
  id: string;
  name: string;
  local: string;
  uri: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  text: string;
  start: number;
  openEnd: number;
  closeStart: number;
  end: number;
  selfClosing: boolean;
}

export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const STRICT_W = 'http://purl.oclc.org/ooxml/wordprocessingml/main';
export const isWord = (node: XmlNode, local?: string) =>
  (node.uri === W || node.uri === STRICT_W) && (!local || node.local === local);

export function parseXml(source: string, part: string): XmlNode {
  const parser = new SaxesParser({ xmlns: true });
  const stack: XmlNode[] = [];
  let root: XmlNode | undefined;
  let nextId = 0;
  let start = 0;
  parser.on('doctype', () => {
    throw new Error(`${part}: DTDs are not supported`);
  });
  parser.on('opentagstart', () => {
    start = source.lastIndexOf('<', parser.position - 1);
  });
  parser.on('opentag', (tag) => {
    if (stack.length >= 256 || nextId >= 250_000) throw new Error('XML structural limit exceeded');
    const node: XmlNode = {
      id: `${part}#${nextId++}`,
      name: tag.name,
      local: tag.local,
      uri: tag.uri,
      attrs: {},
      children: [],
      text: '',
      start,
      openEnd: parser.position,
      closeStart: parser.position,
      end: parser.position,
      selfClosing: tag.isSelfClosing,
    };
    for (const attr of Object.values(tag.attributes)) {
      node.attrs[`{${attr.uri}}${attr.local}`] = attr.value;
    }
    if (stack.length) stack.at(-1)!.children.push(node);
    else root = node;
    stack.push(node);
  });
  parser.on('text', (text) => {
    if (stack.length) stack.at(-1)!.text += text;
  });
  parser.on('cdata', (text) => {
    if (stack.length) stack.at(-1)!.text += text;
  });
  parser.on('closetag', () => {
    const node = stack.pop()!;
    node.end = parser.position;
    node.closeStart = node.selfClosing
      ? node.openEnd
      : source.lastIndexOf('<', parser.position - 1);
  });
  try {
    parser.write(source).close();
  } catch (error) {
    throw new Error(`${part}: ${String(error)}`);
  }
  if (!root) throw new Error(`${part}: empty XML`);
  return root;
}

export const attr = (node: XmlNode, name: string, uri = '') => node.attrs[`{${uri}}${name}`];
export const wordAttr = (node: XmlNode, name: string) => attr(node, name, node.uri);
export const child = (node: XmlNode, name: string) => node.children.find((n) => isWord(n, name));
export function descendants(node: XmlNode): XmlNode[] {
  return [node, ...node.children.flatMap(descendants)];
}
export const escapeXml = (value: string) =>
  value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export function assertText(value: string): void {
  // Newlines/tabs have structural OOXML representations; commands cannot invent them here.
  if (/[\u0000-\u001f\ufffe\uffff]/u.test(value) || /[\ud800-\udfff]/u.test(value)) {
    throw new Error('Text must contain valid XML characters and no tabs or line breaks.');
  }
}

export interface XmlPatch {
  start: number;
  end: number;
  replacement: string;
  /** Existing nodes retained at element-start offsets relative to replacement. */
  retain?: readonly { nodeId: string; offset: number }[];
}
export function orderedPatches(source: string, patches: readonly XmlPatch[]): XmlPatch[] {
  const ordered = [...patches].sort((a, b) => a.start - b.start);
  let previous: XmlPatch | undefined;
  for (const patch of ordered) {
    if (
      !Number.isInteger(patch.start) ||
      !Number.isInteger(patch.end) ||
      patch.start < 0 ||
      patch.end > source.length ||
      patch.end < patch.start ||
      typeof patch.replacement !== 'string' ||
      (previous && (patch.start < previous.end || patch.start === previous.start))
    )
      throw new Error('Overlapping or invalid XML patches');
    previous = patch;
  }
  return ordered;
}
export function patchXml(source: string, patches: readonly XmlPatch[]): string {
  const chunks: string[] = [];
  let offset = 0;
  for (const patch of orderedPatches(source, patches)) {
    chunks.push(source.slice(offset, patch.start), patch.replacement);
    offset = patch.end;
  }
  chunks.push(source.slice(offset));
  return chunks.join('');
}

/** Only the selected w:t is rewritten. All surrounding lexical XML survives. */
export function textPatch(source: string, node: XmlNode, value: string): XmlPatch {
  assertText(value);
  let opening = source.slice(node.start, node.openEnd).replace(/\s*\/>$/, '>');
  let hasSpace = false;
  // Consume complete quoted attributes so text resembling xml:space inside an
  // unfamiliar attribute value is never accidentally rewritten.
  opening = opening.replace(/([^\s=<>/]+)\s*=\s*("[^"]*"|'[^']*')/g, (raw, name: string) => {
    if (name !== 'xml:space') return raw;
    hasSpace = true;
    return 'xml:space="preserve"';
  });
  if (!hasSpace) opening = opening.slice(0, -1) + ' xml:space="preserve">';
  return {
    start: node.start,
    end: node.end,
    replacement: `${opening}${escapeXml(value)}</${node.name}>`,
    retain: [{ nodeId: node.id, offset: 0 }],
  };
}
