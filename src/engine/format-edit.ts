import { DocxPackage } from '../package/docx';
import { child, descendants, isWord, textPatch, type XmlNode, type XmlPatch } from '../package/xml';
import { readDocument } from './document';
import { projectStory } from './projection';
import { effectiveFormat, validateFormat, type TextFormat } from './format';
export interface FormatEdit {
  documentId: string;
  storyId: string;
  origin: string;
  from: number;
  to: number;
  expectedRevision: number;
  format: TextFormat;
}
const opening = (source: string, n: XmlNode) =>
  source.slice(n.start, n.openEnd).replace(/\s*\/>$/, '>');
const newOpening = (source: string, n: XmlNode) =>
  `<${n.name}${(opening(source, n).match(/\s+xmlns(?::[^\s=]+)?\s*=\s*(?:"[^"]*"|'[^']*')/g) ?? []).join('')}>`;
const order = [
  'rStyle',
  'rFonts',
  'b',
  'bCs',
  'i',
  'iCs',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'noProof',
  'snapToGrid',
  'vanish',
  'webHidden',
  'color',
  'spacing',
  'w',
  'kern',
  'position',
  'sz',
  'szCs',
  'highlight',
  'u',
  'effect',
  'bdr',
  'shd',
  'fitText',
  'vertAlign',
  'rtl',
  'cs',
  'em',
  'lang',
  'eastAsianLayout',
  'specVanish',
  'oMath',
  'rPrChange',
];
function setProperties(pkg: DocxPackage, part: string, runId: string, format: TextFormat) {
  for (const [key, value] of Object.entries(format)) {
    const run = descendants(pkg.xml(part)).find((n) => n.id === runId)!;
    const props = child(run, 'rPr');
    if (props && descendants(props).some((n) => isWord(n, 'rPrChange')))
      throw new Error('Tracked formatting is protected');
    const source = pkg.text(part);
    const prefix = run.name.includes(':') ? run.name.split(':')[0] + ':' : '';
    const val =
      typeof value === 'boolean'
        ? key === 'u'
          ? value
            ? 'single'
            : 'none'
          : value
            ? '1'
            : '0'
        : value;
    const replacement = `<${prefix}${key} ${prefix}val="${val}"/>`;
    const existing = props?.children.filter((n) => isWord(n, key)) ?? [];
    if (existing.length > 1) throw new Error(`Duplicate run property: ${key}`);
    let patch: XmlPatch;
    if (existing[0])
      patch = {
        start: existing[0].start,
        end: existing[0].end,
        replacement,
        retain: [{ nodeId: existing[0].id, offset: 0 }],
      };
    else if (!props)
      patch = {
        start: run.openEnd,
        end: run.openEnd,
        replacement: `<${prefix}rPr>${replacement}</${prefix}rPr>`,
      };
    else if (props.selfClosing)
      patch = {
        start: props.start,
        end: props.end,
        replacement: opening(source, props) + replacement + `</${props.name}>`,
        retain: [{ nodeId: props.id, offset: 0 }],
      };
    else {
      const next = props.children.find(
        (n) => isWord(n) && order.indexOf(n.local) > order.indexOf(key),
      );
      const start = next?.start ?? props.closeStart;
      patch = { start, end: start, replacement };
    }
    pkg = pkg.withXmlPatches(new Map([[part, [patch]]]));
  }
  return pkg;
}
/** Isolate just the selected portion of one source run, retaining every untouched child. */
function isolate(
  pkg: DocxPackage,
  part: string,
  runId: string,
  firstId: string,
  from: number,
  lastId: string,
  to: number,
) {
  const nodes = descendants(pkg.xml(part));
  const run = nodes.find((n) => n.id === runId)!,
    first = nodes.find((n) => n.id === firstId)!,
    last = nodes.find((n) => n.id === lastId)!;
  const props = child(run, 'rPr'),
    source = pkg.text(part);
  const content = run.children.filter((n) => !isWord(n, 'rPr'));
  if (
    content[0]?.id === first.id &&
    from === 0 &&
    content.at(-1)?.id === last.id &&
    to === last.text.length
  )
    return { pkg, runId };
  const hasLeft = from > 0 || content[0]?.id !== first.id;
  const hasRight = to < last.text.length || content.at(-1)?.id !== last.id;
  let replacement = '';
  const retain: { nodeId: string; offset: number }[] = [];
  const copy = (start: number, end: number, identities = true) => {
    const offset = replacement.length;
    replacement += source.slice(start, end);
    if (identities)
      for (const n of nodes)
        if (n.start >= start && n.start < end && n.end <= end)
          retain.push({ nodeId: n.id, offset: offset + n.start - start });
  };
  const openRun = (original: boolean) => {
    if (original) retain.push({ nodeId: run.id, offset: replacement.length });
    replacement += original ? opening(source, run) : newOpening(source, run);
    if (props) {
      copy(run.openEnd, props.start, false);
      copy(props.start, props.end, original);
    }
  };
  const rewriteText = (n: XmlNode, text: string, identity: boolean) => {
    if (identity) retain.push({ nodeId: n.id, offset: replacement.length });
    replacement += textPatch(source, n, text).replacement;
  };
  if (hasLeft) {
    openRun(true);
    copy(props?.end ?? run.openEnd, first.start);
    if (from) rewriteText(first, first.text.slice(0, from), false);
    replacement += `</${run.name}>`;
  }
  const selectedOffset = replacement.length;
  openRun(!hasLeft);
  if (!hasLeft) copy(props?.end ?? run.openEnd, first.start);
  if (first.id === last.id) rewriteText(first, first.text.slice(from, to), true);
  else {
    rewriteText(first, first.text.slice(from), true);
    copy(first.end, last.start);
    rewriteText(last, last.text.slice(0, to), true);
  }
  if (!hasRight) copy(last.end, run.closeStart);
  replacement += `</${run.name}>`;
  if (hasRight) {
    openRun(false);
    if (to < last.text.length) rewriteText(last, last.text.slice(to), false);
    copy(last.end, run.closeStart);
    replacement += `</${run.name}>`;
  }
  pkg = pkg.withXmlPatches(
    new Map([[part, [{ start: run.start, end: run.end, replacement, retain }]]]),
  );
  const selected = descendants(pkg.xml(part)).find((n) => n.start === run.start + selectedOffset)!;
  return { pkg, runId: selected.id };
}
export function formatProjection(pkg: DocxPackage, edit: FormatEdit): DocxPackage {
  validateFormat(edit.format);
  const story = readDocument(pkg).stories.find((s) => s.id === edit.storyId);
  if (!story) throw new Error('Unknown story');
  const projection = projectStory(story, edit.origin);
  if (
    !Number.isInteger(edit.from) ||
    !Number.isInteger(edit.to) ||
    edit.from < 0 ||
    edit.to <= edit.from ||
    edit.to > projection.text.length
  )
    throw new Error('Select text to format');
  if (projection.text.slice(edit.from, edit.to).includes('\ufffc'))
    throw new Error('Format text without selecting document objects');
  const nodes = descendants(pkg.xml(story.part));
  const groups = new Map<string, { first: string; last: string; from: number; to: number }>();
  for (const l of projection.locations) {
    if (l.token.kind !== 'text' || l.from >= edit.to || l.to <= edit.from || l.from === l.to)
      continue;
    const span = l.token.span;
    if (!span.editable) throw new Error(span.reason ?? 'Protected text');
    const t = nodes.find((n) => n.id === span.id)!;
    const run = nodes.find((n) => n.children.some((c) => c.id === t.id));
    if (
      !run ||
      !isWord(run, 'r') ||
      run.children.some(
        (n) => isWord(n, 'rPr') && descendants(n).some((c) => isWord(c, 'rPrChange')),
      )
    )
      throw new Error('Formatting requires an ordinary editable run');
    const from = Math.max(0, edit.from - l.from),
      to = Math.min(span.text.length, edit.to - l.from);
    for (const offset of [from, to])
      if (
        offset > 0 &&
        offset < span.text.length &&
        /[\ud800-\udbff]/.test(span.text[offset - 1]!) &&
        /[\udc00-\udfff]/.test(span.text[offset]!)
      )
        throw new Error('Split-surrogate format range');
    const current = effectiveFormat(span);
    if (
      Object.entries(edit.format).every(([key, value]) => {
        const existing = current[key as keyof TextFormat];
        return typeof existing === 'string' && typeof value === 'string'
          ? existing.toLowerCase() === value.toLowerCase()
          : existing === value;
      })
    )
      continue;
    const group = groups.get(run.id);
    if (group) {
      group.last = t.id;
      group.to = to;
    } else groups.set(run.id, { first: t.id, last: t.id, from, to });
  }
  if (!groups.size) return pkg;
  for (const [runId, g] of [...groups].reverse()) {
    const isolated = isolate(pkg, story.part, runId, g.first, g.from, g.last, g.to);
    pkg = setProperties(isolated.pkg, story.part, isolated.runId, edit.format);
  }
  return pkg;
}
