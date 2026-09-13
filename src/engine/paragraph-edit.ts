import { Styles, value } from './styles';
import { DocxPackage } from '../package/docx';
import {
  assertText,
  escapeXml,
  wordAttr,
  child,
  descendants,
  isWord,
  textPatch,
  type XmlNode,
  type XmlPatch,
} from '../package/xml';
import { readDocument, spans, type Story } from './document';
import { positionAt, replacePieces, resolveTextRange } from './text-range';
import { projectStory, type ProjectionEdit } from './projection';

const nodeIn = (pkg: DocxPackage, part: string, id: string) => {
  const node = descendants(pkg.xml(part)).find((n) => n.id === id);
  if (!node) throw new Error('Source paragraph is stale');
  return node;
};
const apply = (pkg: DocxPackage, part: string, patches: readonly XmlPatch[]) =>
  pkg.withXmlPatches(new Map([[part, patches]]));
const opening = (source: string, node: XmlNode) =>
  source.slice(node.start, node.openEnd).replace(/\s*\/>$/, '>');
// New wrappers inherit namespace declarations, never duplicate paragraph IDs or unknown attributes.
const newOpening = (source: string, node: XmlNode) =>
  `<${node.name}${(opening(source, node).match(/\s+xmlns(?::[^\s=]+)?\s*=\s*(?:"[^"]*"|'[^']*')/g) ?? []).join('')}>`;

function editableParagraph(pkg: DocxPackage, story: Story, id: string) {
  const node = nodeIn(pkg, story.part, id);
  const parent = descendants(pkg.xml(story.part)).find((n) =>
    n.children.some((c) => c.id === node.id),
  );
  if (
    !isWord(node, 'p') ||
    !parent ||
    !['body', 'tc', 'hdr', 'ftr', 'footnote', 'endnote'].some((n) => isWord(parent, n))
  )
    throw new Error('Paragraph restructuring requires an ordinary paragraph in one container');
  if (descendants(node).some((n) => isWord(n, 'rPrChange')))
    throw new Error('Tracked run formatting cannot be split or joined');
  const props = child(node, 'pPr');
  if (
    node.text.trim() ||
    (props && node.children[0]?.id !== props.id) ||
    node.children.filter((n) => isWord(n, 'pPr')).length > 1
  )
    throw new Error('Unsupported paragraph content order');
  if (
    props &&
    descendants(props).some((n) =>
      ['sectPr', 'pPrChange', 'rPrChange', 'numberingChange', 'ins', 'del'].some((name) =>
        isWord(n, name),
      ),
    )
  )
    throw new Error('Section and review boundaries cannot be split or joined');
  // These nodes move with their source side of the split; they are never rewritten.
  const markers = [
    'bookmarkStart',
    'bookmarkEnd',
    'commentRangeStart',
    'commentRangeEnd',
    'proofErr',
  ];
  const runContent = [
    'rPr',
    't',
    'tab',
    'br',
    'cr',
    'lastRenderedPageBreak',
    'drawing',
    'pict',
    'object',
    'sym',
    'footnoteReference',
    'endnoteReference',
    'commentReference',
  ];
  const unsupported = node.children.flatMap((n) =>
    isWord(n, 'r')
      ? n.children.filter((c) => !runContent.some((name) => isWord(c, name)))
      : isWord(n, 'pPr') || markers.some((name) => isWord(n, name))
        ? []
        : [n],
  )[0];
  if (unsupported)
    throw new Error(
      `Cannot split or join this paragraph: ${unsupported.name} is not yet supported`,
    );
  const sourceSpans = spans(readDocument(pkg)).filter(
    (s) => s.source.part === story.part && s.source.start >= node.start && s.source.end <= node.end,
  );
  if (sourceSpans.some((s) => !s.editable))
    throw new Error('Protected paragraphs cannot be split or joined');
  // Also protect paragraphs with no text nodes, including empty field results.
  if (!sourceSpans.length) {
    const model = readDocument(pkg);
    if (
      model.diagnostics.some(
        (d) => d.severity === 'error' || ['signed-package', 'protected-editing'].includes(d.code),
      ) ||
      story.kind === 'comment'
    )
      throw new Error('Protected paragraphs cannot be edited');
    let depth = 0;
    for (const n of descendants(pkg.xml(story.part))) {
      if (n.start >= node.start) break;
      if (isWord(n, 'fldChar')) {
        const type = n.attrs[`{${n.uri}}fldCharType`];
        if (type === 'begin') depth++;
        else if (type === 'end') depth--;
      }
    }
    if (depth) throw new Error('An empty field result is protected');
  }
  return { node, parent, props, sourceSpans };
}

function ensureText(pkg: DocxPackage, story: Story, id: string): DocxPackage {
  const { node, props, sourceSpans } = editableParagraph(pkg, story, id);
  if (sourceSpans.length) return pkg;
  const prefix = node.name.includes(':') ? node.name.split(':')[0] + ':' : '';
  const source = pkg.text(story.part);
  const text = `<${prefix}t xml:space="preserve"></${prefix}t>`;
  const existingRun = node.children.filter((n) => isWord(n, 'r')).at(-1);
  if (existingRun)
    return apply(
      pkg,
      story.part,
      existingRun.selfClosing
        ? [
            {
              start: existingRun.start,
              end: existingRun.end,
              replacement: opening(source, existingRun) + text + `</${existingRun.name}>`,
              retain: [{ nodeId: existingRun.id, offset: 0 }],
            },
          ]
        : [{ start: existingRun.closeStart, end: existingRun.closeStart, replacement: text }],
    );
  const mark = props && child(props, 'rPr');
  const run = `<${prefix}r>${mark ? source.slice(mark.start, mark.end) : ''}${text}</${prefix}r>`;
  return apply(
    pkg,
    story.part,
    node.selfClosing
      ? [
          {
            start: node.start,
            end: node.end,
            replacement: opening(source, node) + run + `</${node.name}>`,
            retain: [{ nodeId: node.id, offset: 0 }],
          },
        ]
      : [{ start: node.closeStart, end: node.closeStart, replacement: run }],
  );
}

/** Empty split paragraphs carry typing properties on the paragraph mark, not a dummy run. */
function markPatches(source: string, node: XmlNode, runProps?: XmlNode): XmlPatch[] {
  const props = child(node, 'pPr');
  const mark = props && child(props, 'rPr');
  const replacement = runProps ? source.slice(runProps.start, runProps.end) : '';
  if (mark) return [{ start: mark.start, end: mark.end, replacement }];
  if (!replacement) return [];
  if (props)
    return props.selfClosing
      ? [
          {
            start: props.start,
            end: props.end,
            replacement: opening(source, props) + replacement + `</${props.name}>`,
            retain: [{ nodeId: props.id, offset: 0 }],
          },
        ]
      : [{ start: props.closeStart, end: props.closeStart, replacement }];
  const prefix = node.name.includes(':') ? node.name.split(':')[0] + ':' : '';
  return [
    {
      start: node.openEnd,
      end: node.openEnd,
      replacement: `<${prefix}pPr>${replacement}</${prefix}pPr>`,
    },
  ];
}
function copiedProperties(source: string, node: XmlNode, runProps?: XmlNode, empty = false) {
  const props = child(node, 'pPr');
  let value = props ? source.slice(props.start, props.end) : '';
  if (!empty) return value;
  for (const patch of markPatches(source, node, runProps)) {
    const start = props ? patch.start - props.start : 0;
    const end = props ? patch.end - props.start : 0;
    value = value.slice(0, start) + patch.replacement + value.slice(end);
  }
  return value;
}

export function splitParagraph(
  pkg: DocxPackage,
  story: Story,
  id: string,
  offset: number,
  typingSpan?: string,
) {
  const { node, props, sourceSpans } = editableParagraph(pkg, story, id);
  if (!sourceSpans.length) {
    if (offset !== 0) throw new Error('Invalid text offset');
    if (node.children.some((n) => isWord(n, 'r') && n.children.some((c) => !isWord(c, 'rPr'))))
      throw new Error('Place the split caret in text beside the document object');
    const source = pkg.text(story.part);
    const lastRun = node.children.filter((n) => isWord(n, 'r')).at(-1);
    const runProps = lastRun ? child(lastRun, 'rPr') : props && child(props, 'rPr');
    const left = node.selfClosing ? opening(source, node) : '';
    const prefix = left + `</${node.name}>`;
    const start = node.selfClosing ? node.start : node.closeStart;
    const updated = apply(pkg, story.part, [
      {
        start,
        end: node.selfClosing ? node.end : start,
        replacement:
          prefix +
          newOpening(source, node) +
          copiedProperties(source, node, runProps, true) +
          (node.selfClosing ? `</${node.name}>` : ''),
        retain: node.selfClosing ? [{ nodeId: node.id, offset: 0 }] : undefined,
      },
    ]);
    const right = descendants(updated.xml(story.part)).find(
      (n) => n.start === start + prefix.length,
    )!;
    return { pkg: updated, rightId: right.id, rightSpan: undefined };
  }
  const pieces = sourceSpans.map((s) => ({ spanId: s.id, text: s.text }));
  const point = positionAt(pieces, offset, 'left', typingSpan);
  // Validate offset and Unicode boundaries without changing text.
  replacePieces(pieces, offset, offset, '');
  const t = nodeIn(pkg, story.part, point.spanId);
  const run = node.children.find((n) => n.children.some((c) => c.id === t.id))!;
  const source = pkg.text(story.part),
    rPr = child(run, 'rPr');
  const content = run.children.filter((n) => !isWord(n, 'rPr'));
  const beforeRun = point.offset === 0 && content[0]?.id === t.id;
  const afterRun = point.offset === t.text.length && content.at(-1)?.id === t.id;
  if (beforeRun || afterRun) {
    const cut = beforeRun ? run.start : run.end;
    const leftEmpty = !node.children.some((n) => isWord(n, 'r') && n.end <= cut);
    const rightEmpty = !node.children.some((n) => isWord(n, 'r') && n.start >= cut);
    const prefix =
      `</${node.name}>` +
      newOpening(source, node) +
      copiedProperties(source, node, rPr, rightEmpty);
    const patches = leftEmpty ? markPatches(source, node, rPr) : [];
    // The first run can start exactly at the paragraph opening tag's end.
    const samePoint = patches.find((p) => p.start === cut && p.end === cut);
    const insertion = (samePoint?.replacement ?? '') + prefix;
    const updated = apply(pkg, story.part, [
      ...patches.filter((p) => p !== samePoint),
      { start: cut, end: cut, replacement: insertion },
    ]);
    const adjustedCut =
      cut +
      patches
        .filter((p) => p !== samePoint)
        .reduce((n, p) => n + p.replacement.length - (p.end - p.start), 0);
    const created = descendants(updated.xml(story.part)).find(
      (n) =>
        n.start === adjustedCut + (samePoint?.replacement.length ?? 0) + `</${node.name}>`.length,
    )!;
    const rightText = descendants(created).find((n) => isWord(n, 't'));
    return { pkg: updated, rightId: created.id, rightSpan: rightText?.id };
  }
  const left = textPatch(source, t, t.text.slice(0, point.offset));
  const right = textPatch(source, t, t.text.slice(point.offset));
  const leftXml = point.offset === 0 && t.text.length ? '' : left.replacement;
  const rightXml = point.offset === t.text.length && t.text.length ? '' : right.replacement;
  const paragraphStart = leftXml.length + `</${run.name}></${node.name}>`.length;
  const prefix =
    leftXml +
    `</${run.name}></${node.name}>` +
    newOpening(source, node) +
    (props ? source.slice(props.start, props.end) : '') +
    newOpening(source, run) +
    (rPr ? source.slice(rPr.start, rPr.end) : '');
  const updated = apply(pkg, story.part, [
    {
      ...left,
      replacement: prefix + rightXml,
      retain: [{ nodeId: t.id, offset: leftXml ? 0 : prefix.length }],
    },
  ]);
  const created = descendants(updated.xml(story.part)).find(
    (n) => n.start === t.start + paragraphStart,
  )!;
  const rightText = descendants(created).find((n) => isWord(n, 't'));
  return { pkg: updated, rightId: created.id, rightSpan: rightText?.id };
}

export function joinParagraphs(pkg: DocxPackage, story: Story, leftId: string, rightId: string) {
  const left = editableParagraph(pkg, story, leftId),
    right = editableParagraph(pkg, story, rightId);
  if (
    left.parent.id !== right.parent.id ||
    left.parent.children.findIndex((n) => n.id === right.node.id) !==
      left.parent.children.findIndex((n) => n.id === left.node.id) + 1
  )
    throw new Error('Only adjacent paragraphs in the same container can be joined');
  if (
    Object.keys(right.node.attrs).some((key) => key.startsWith('{http://www.w3.org/2000/xmlns/}'))
  )
    throw new Error('Joining a paragraph with local namespace declarations is not supported');
  // The first paragraph owns paragraph formatting; run XML and identities survive.
  if (left.node.selfClosing)
    pkg = apply(pkg, story.part, [
      {
        start: left.node.start,
        end: left.node.end,
        replacement: opening(pkg.text(story.part), left.node) + `</${left.node.name}>`,
        retain: [{ nodeId: left.node.id, offset: 0 }],
      },
    ]);
  const a = nodeIn(pkg, story.part, leftId),
    b = nodeIn(pkg, story.part, rightId);
  const props = child(b, 'pPr'),
    source = pkg.text(story.part);
  if (b.selfClosing)
    return apply(pkg, story.part, [{ start: b.start, end: b.end, replacement: '' }]);
  return apply(pkg, story.part, [
    {
      start: a.closeStart,
      end: props?.end ?? b.openEnd,
      replacement: source.slice(a.end, b.start),
    },
  ]);
}

function replaceInParagraph(
  pkg: DocxPackage,
  story: Story,
  id: string,
  from: number,
  to: number,
  text: string,
  typingSpan?: string,
) {
  if (from === to && !text) return pkg;
  const paragraph = nodeIn(pkg, story.part, id);
  let sourceSpans = spans(readDocument(pkg)).filter(
    (s) =>
      s.source.part === story.part &&
      s.source.start >= paragraph.start &&
      s.source.end <= paragraph.end,
  );
  if (!sourceSpans.length) {
    pkg = ensureText(pkg, story, id);
    const current = nodeIn(pkg, story.part, id);
    sourceSpans = spans(readDocument(pkg)).filter(
      (s) =>
        s.source.part === story.part &&
        s.source.start >= current.start &&
        s.source.end <= current.end,
    );
  }
  const pieces = sourceSpans.map((s) => ({ spanId: s.id, text: s.text }));
  const anchor = positionAt(pieces, from, from === to ? 'left' : 'right', typingSpan);
  const head = from === to ? anchor : positionAt(pieces, to);
  const currentStory = readDocument(pkg).stories.find((s) => s.id === story.id)!;
  resolveTextRange(currentStory, anchor, head);
  const result = replacePieces(pieces, from, to, text, anchor);
  const source = pkg.text(story.part);
  const patches = result.pieces.flatMap((piece, i) =>
    piece.text === pieces[i]!.text
      ? []
      : [textPatch(source, nodeIn(pkg, story.part, piece.spanId), piece.text)],
  );
  return patches.length ? apply(pkg, story.part, patches) : pkg;
}

function applyNextStyle(pkg: DocxPackage, story: Story, originalId: string, rightId: string) {
  const paragraph = story.paragraphs.find((p) => p.id === originalId);
  if (!paragraph || paragraph.numbering) return pkg;
  const styles = new Styles(pkg);
  const current = styles.entries.get(paragraph.styleId ?? styles.defaultParagraph ?? '');
  const next = current && value(current, 'next');
  const target = next && styles.entries.get(next);
  if (!next || !target || wordAttr(target, 'type') !== 'paragraph' || next === paragraph.styleId)
    return pkg;
  const node = nodeIn(pkg, story.part, rightId),
    props = child(node, 'pPr');
  const existing = props && child(props, 'pStyle'),
    source = pkg.text(story.part);
  const prefix = node.name.includes(':') ? node.name.split(':')[0] + ':' : '';
  const replacement = `<${prefix}pStyle ${prefix}val="${escapeXml(next).replace(/"/g, '&quot;')}"/>`;
  if (existing)
    return apply(pkg, story.part, [
      {
        start: existing.start,
        end: existing.end,
        replacement,
        retain: [{ nodeId: existing.id, offset: 0 }],
      },
    ]);
  if (!props)
    return apply(pkg, story.part, [
      {
        start: node.openEnd,
        end: node.openEnd,
        replacement: `<${prefix}pPr>${replacement}</${prefix}pPr>`,
      },
    ]);
  if (props.selfClosing)
    return apply(pkg, story.part, [
      {
        start: props.start,
        end: props.end,
        replacement: opening(source, props) + replacement + `</${props.name}>`,
        retain: [{ nodeId: props.id, offset: 0 }],
      },
    ]);
  return apply(pkg, story.part, [{ start: props.openEnd, end: props.openEnd, replacement }]);
}

/** Stage the complete input (including multi-paragraph paste) before publication. */
export function editProjection(pkg: DocxPackage, edit: ProjectionEdit) {
  const story = readDocument(pkg).stories.find((s) => s.id === edit.storyId);
  if (!story) throw new Error('Unknown story');
  const projection = projectStory(story, edit.origin);
  if (
    !Number.isInteger(edit.from) ||
    !Number.isInteger(edit.to) ||
    edit.from < 0 ||
    edit.to < edit.from ||
    edit.to > projection.text.length
  )
    throw new Error('Invalid projection range');
  const selected = projection.text.slice(edit.from, edit.to);
  if (selected.includes('\ufffc') || edit.text.includes('\ufffc'))
    throw new Error('Document objects and anchors cannot be edited as text');
  assertText(edit.text.replace(/\n/g, ''));
  if (edit.enter !== undefined && typeof edit.enter !== 'boolean')
    throw new Error('Invalid Enter command');
  const first = projection.paragraphs.find((p) => p.from <= edit.from && edit.from <= p.to);
  const last = projection.paragraphs.find((p) => p.from <= edit.to && edit.to <= p.to);
  if (!first || !last) throw new Error('Select text inside a paragraph');
  const group = projection.paragraphs.slice(
    projection.paragraphs.indexOf(first),
    projection.paragraphs.indexOf(last) + 1,
  );
  if (group.length > 1 || edit.text.includes('\n')) {
    for (const p of group) editableParagraph(pkg, story, p.id);
    for (const p of group.slice(1)) pkg = joinParagraphs(pkg, story, first.id, p.id);
  }
  // Object slots are not text offsets. For text-only edits count source characters only.
  const from = projection.text.slice(first.from, edit.from).replace(/\ufffc/g, '').length;
  const to = from + selected.replace(/\n/g, '').length;
  const chunks = edit.text.split('\n');
  const candidates = projection.locations.filter(
    (l) => l.token.kind === 'text' && l.from <= edit.from && edit.from <= l.to,
  );
  const owner =
    candidates.find((l) => l.token.kind === 'text' && l.token.span.id === edit.typingSpan) ??
    candidates.find((l) => edit.from < l.to && (edit.from !== edit.to || edit.from > l.from)) ??
    candidates.find((l) => edit.from > l.from) ??
    candidates[0];
  pkg = replaceInParagraph(
    pkg,
    story,
    first.id,
    from,
    to,
    chunks.join(''),
    owner?.token.kind === 'text' ? owner.token.span.id : undefined,
  );
  let splitOwner = owner?.token.kind === 'text' ? owner.token.span.id : undefined;
  let paragraphId = first.id,
    splitOffset = from + chunks[0]!.length;
  for (const chunk of chunks.slice(1)) {
    const split = splitParagraph(pkg, story, paragraphId, splitOffset, splitOwner);
    pkg = split.pkg;
    paragraphId = split.rightId;
    splitOwner = split.rightSpan;
    splitOffset = chunk.length;
  }
  if (edit.enter && edit.text === '\n' && edit.from === edit.to && edit.from === first.to)
    pkg = applyNextStyle(pkg, story, first.id, paragraphId);
  const updatedStory = readDocument(pkg).stories.find((s) => s.id === story.id)!;
  const updated = projectStory(updatedStory, edit.origin);
  const offset = edit.from + edit.text.length;
  const atCaret = updated.locations.filter(
    (l) => l.token.kind === 'text' && l.from <= offset && offset <= l.to,
  );
  const retainedOwner =
    !edit.text.includes('\n') && owner?.token.kind === 'text' ? owner.token.span.id : undefined;
  const target =
    atCaret.find((l) => l.token.kind === 'text' && l.token.span.id === retainedOwner) ??
    atCaret.find((l) => offset > l.from) ??
    atCaret[0];
  return { pkg, typingSpan: target?.token.kind === 'text' ? target.token.span.id : undefined };
}
