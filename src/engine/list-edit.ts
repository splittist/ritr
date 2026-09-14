import { DocxPackage } from '../package/docx';
import { child, descendants, isWord, type XmlNode, type XmlPatch } from '../package/xml';
import { readDocument } from './document';
import { editableParagraph } from './paragraph-edit';
import { Numbering } from './numbering';
import { Styles } from './styles';
import { projectStory, selectedParagraphs, type ProjectionEdit } from './projection';
const pOrder = [
  'pStyle',
  'keepNext',
  'keepLines',
  'pageBreakBefore',
  'framePr',
  'widowControl',
  'numPr',
  'suppressLineNumbers',
  'pBdr',
  'shd',
  'tabs',
  'suppressAutoHyphens',
  'kinsoku',
  'wordWrap',
  'overflowPunct',
  'topLinePunct',
  'autoSpaceDE',
  'autoSpaceDN',
  'bidi',
  'adjustRightInd',
  'snapToGrid',
  'spacing',
  'ind',
  'contextualSpacing',
  'mirrorIndents',
  'suppressOverlap',
  'jc',
  'textDirection',
  'textAlignment',
  'textboxTightWrap',
  'outlineLvl',
  'divId',
  'cnfStyle',
  'rPr',
  'sectPr',
  'pPrChange',
];
const opening = (source: string, node: XmlNode) =>
  source.slice(node.start, node.openEnd).replace(/\s*\/>$/, '>');
function setChild(
  pkg: DocxPackage,
  part: string,
  parentId: string,
  name: string,
  content: (prefix: string) => string,
  order: string[],
) {
  const parent = descendants(pkg.xml(part)).find((n) => n.id === parentId)!;
  const existing = parent.children.filter((n) => isWord(n, name));
  if (existing.length > 1) throw new Error(`Ambiguous paragraph property: ${name}`);
  const prefix = parent.name.includes(':') ? parent.name.split(':')[0] + ':' : '';
  const replacement = content(prefix),
    source = pkg.text(part);
  let patch: XmlPatch;
  if (existing[0])
    patch = {
      start: existing[0].start,
      end: existing[0].end,
      replacement,
      retain: [{ nodeId: existing[0].id, offset: 0 }],
    };
  else if (parent.selfClosing)
    patch = {
      start: parent.start,
      end: parent.end,
      replacement: opening(source, parent) + replacement + `</${parent.name}>`,
      retain: [{ nodeId: parent.id, offset: 0 }],
    };
  else {
    const next = parent.children.find(
      (n) => isWord(n) && order.indexOf(n.local) > order.indexOf(name),
    );
    const start =
      isWord(parent, 'p') && name === 'pPr' ? parent.openEnd : (next?.start ?? parent.closeStart);
    patch = { start, end: start, replacement };
  }
  return pkg.withXmlPatches(new Map([[part, [patch]]]));
}
function membership(pkg: DocxPackage, part: string, id: string, numId: string, level: number) {
  let node = descendants(pkg.xml(part)).find((n) => n.id === id)!;
  if (!child(node, 'pPr'))
    pkg = setChild(pkg, part, id, 'pPr', (prefix) => `<${prefix}pPr/>`, ['pPr', 'r']);
  node = descendants(pkg.xml(part)).find((n) => n.id === id)!;
  const props = child(node, 'pPr')!;
  if (!child(props, 'numPr'))
    pkg = setChild(pkg, part, props.id, 'numPr', (prefix) => `<${prefix}numPr/>`, pOrder);
  const num = child(
    descendants(pkg.xml(part)).find((n) => n.id === props.id)!,
    'numPr',
  )!;
  pkg = setChild(
    pkg,
    part,
    num.id,
    'ilvl',
    (prefix) => `<${prefix}ilvl ${prefix}val="${level}"/>`,
    ['ilvl', 'numId', 'numberingChange', 'ins'],
  );
  return setChild(
    pkg,
    part,
    num.id,
    'numId',
    (prefix) => `<${prefix}numId ${prefix}val="${numId}"/>`,
    ['ilvl', 'numId', 'numberingChange', 'ins'],
  );
}
/** Change only direct list membership; definitions, runs, bookmarks and direct indentation survive. */
export function editListProjection(pkg: DocxPackage, edit: ProjectionEdit) {
  if (
    !edit.listAction ||
    !['indent', 'outdent', 'enter'].includes(edit.listAction) ||
    edit.text !== ''
  )
    throw new Error('Invalid list command');
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
    throw new Error('Invalid list selection');
  const selected = selectedParagraphs(projection.paragraphs, edit.from, edit.to);
  if (!selected.length || selected.some((p) => !p.numbering))
    throw new Error('Place the caret in a list item or select only list paragraphs');
  if (
    edit.listAction === 'enter' &&
    (edit.from !== edit.to || selected.length !== 1 || selected[0]!.from !== selected[0]!.to)
  )
    throw new Error('Empty-item Enter requires an empty list paragraph');
  const numbering = new Numbering(pkg, new Styles(pkg));
  let parent: string | undefined;
  const plans = selected.map((p) => {
    const guarded = editableParagraph(pkg, story, p.id);
    if (parent && parent !== guarded.parent.id)
      throw new Error('List commands cannot cross table cells or containers');
    parent = guarded.parent.id;
    if ((guarded.props?.children.filter((n) => isWord(n, 'numPr')).length ?? 0) > 1)
      throw new Error('Ambiguous paragraph property: numPr');
    const list = p.numbering!;
    if (
      !/^\d+$/.test(list.numId) ||
      !Number.isInteger(list.level) ||
      list.level < 0 ||
      list.level > 8
    )
      throw new Error('Unsupported list membership');
    const level = list.level + (edit.listAction === 'indent' ? 1 : -1);
    if (level > 8) throw new Error('This item is already at the deepest list level');
    if (level >= 0 && !numbering.definition(list.numId).levels.has(level))
      throw new Error(`This list has no definition for level ${level + 1}`);
    return { id: p.id, numId: level < 0 ? '0' : list.numId, level: Math.max(0, level) };
  });
  for (const plan of plans) pkg = membership(pkg, story.part, plan.id, plan.numId, plan.level);
  return {
    pkg,
    typingSpan: undefined,
    selection: { anchor: edit.from, head: edit.to },
    label:
      edit.listAction === 'indent'
        ? 'Increase list level'
        : plans.every((p) => p.numId === '0')
          ? 'Exit list'
          : 'Decrease list level',
  };
}
