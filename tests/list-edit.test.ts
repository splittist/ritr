import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workspace } from '../src/engine/workspace';
import {
  projectStory,
  tokenId,
  mapParagraphs,
  mapListAction,
  type ListAction,
} from '../src/engine/projection';
import { descendants, W } from '../src/package/xml';
import { fixture, relation, relationships } from './helpers';
const item = (text: string, level = 0, extra = '') =>
  `<w:p w:rsidR="12345678"><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="1"/></w:numPr>${extra}</w:pPr><w:bookmarkStart w:id="3" w:name="marker"/><w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p>`;
function setup(body = item('first') + item('second'), levels = [0, 1, 2], styles = '') {
  const w = new Workspace(),
    id = w.open(
      'lists.docx',
      fixture(body, {
        'word/_rels/document.xml.rels': relationships(
          relation('n', 'numbering', 'numbering.xml') + relation('s', 'styles', 'styles.xml'),
        ),
        'word/numbering.xml': `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0">${levels.map((level) => `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%${level + 1}."/><w:pPr><w:ind w:left="${(level + 1) * 720}" w:hanging="360"/></w:pPr></w:lvl>`).join('')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
        'word/styles.xml': `<w:styles xmlns:w="${W}">${styles}</w:styles>`,
        'custom.bin': new Uint8Array([9, 8, 7]),
      }),
    );
  const request = () => {
    const doc = w.document(id),
      story = doc.model.stories[0]!;
    return {
      documentId: id,
      storyId: story.id,
      origin: tokenId(story.tokens.find((t) => t.kind === 'code' && t.role === 'paragraph-start')!),
      expectedRevision: doc.revision,
    };
  };
  const change = (action: ListAction, from = 0, to = from) =>
    w.applyProjection({ ...request(), from, to, text: '', listAction: action });
  const enter = (from = 0) =>
    w.applyProjection({ ...request(), from, to: from, text: '\n', enter: true });
  const paragraphs = () => w.document(id).model.stories[0]!.paragraphs;
  return { w, id, request, change, enter, paragraphs };
}
test('list level changes preserve runs, markers, definitions and package parts, with atomic undo and caret restoration', () => {
  const { w, id, change, paragraphs } = setup();
  const before = w.package(id),
    nodes = descendants(before.xml(before.mainPart));
  const preserved = nodes
    .filter((n) => ['r', 't', 'bookmarkStart', 'bookmarkEnd'].includes(n.local))
    .map((n) => ({ id: n.id, xml: before.text(before.mainPart).slice(n.start, n.end) }));
  change('indent', 2);
  assert.equal(paragraphs()[0]!.numbering!.level, 1);
  assert.equal(paragraphs()[0]!.layout.left, 1440);
  assert.equal(paragraphs()[1]!.numbering!.level, 0);
  for (const n of preserved) {
    const after = descendants(w.package(id).xml(before.mainPart)).find((x) => x.id === n.id)!;
    assert.ok(after);
    assert.equal(w.package(id).text(before.mainPart).slice(after.start, after.end), n.xml);
  }
  assert.deepEqual(
    w
      .report(id)
      .filter((p) => p.status !== 'unchanged')
      .map((p) => p.part),
    ['word/document.xml'],
  );
  w.undo();
  assert.equal(w.package(id), before);
  assert.equal(w.selection?.head, 2);
  w.redo();
  assert.equal(paragraphs()[0]!.layout.left, 1440);
  assert.equal(w.selection?.head, 2);
  const reopened = new Workspace(),
    copy = reopened.open('copy', w.package(id).save());
  assert.equal(reopened.document(copy).model.stories[0]!.paragraphs[0]!.numbering!.level, 1);
});
test('empty-item Enter moves up a level, exits the top level, then creates a normal paragraph', () => {
  const { w, id, enter, paragraphs, request } = setup(item('', 1));
  enter();
  assert.equal(paragraphs().length, 1);
  assert.equal(paragraphs()[0]!.numbering!.level, 0);
  enter();
  assert.equal(paragraphs().length, 1);
  assert.equal(paragraphs()[0]!.numbering, undefined);
  assert.equal(paragraphs()[0]!.layout.left, 0);
  enter();
  assert.equal(paragraphs().length, 2);
  assert.ok(paragraphs().every((p) => !p.numbering));
  w.undo();
  assert.equal(paragraphs().length, 1);
  assert.equal(w.selection?.head, 0);
  w.undo();
  assert.equal(paragraphs()[0]!.numbering!.level, 0);
  w.undo();
  assert.equal(paragraphs()[0]!.numbering!.level, 1);
  w.applyProjection({ ...request(), from: 0, to: 0, text: '\n' });
  assert.equal(paragraphs().length, 2, 'Pasted breaks do not exit a list');
  assert.ok(paragraphs().every((p) => p.numbering?.level === 1));
});
test('nonempty Enter continues numbering and multi-item commands are one undo entry', () => {
  const { w, change, enter, paragraphs } = setup(item('a') + item('b'));
  change('indent', 0, 2);
  assert.deepEqual(
    paragraphs().map((p) => p.numbering?.level),
    [1, 0],
    'A selection ending at the next paragraph start excludes it',
  );
  w.undo();
  change('indent', 0, 3);
  assert.deepEqual(
    paragraphs().map((p) => p.numbering?.level),
    [1, 1],
  );
  w.undo();
  assert.deepEqual(
    paragraphs().map((p) => p.numbering?.level),
    [0, 0],
  );
  enter(1);
  assert.deepEqual(
    paragraphs().map((p) => p.numbering?.text),
    ['1.', '2.', '3.'],
  );
});
test('inherited list membership is overridden locally while explicit indentation and styles survive', () => {
  const styles =
    '<w:style w:type="paragraph" w:default="1" w:styleId="List"><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>';
  const { w, id, change, paragraphs } = setup(
    '<w:p><w:bookmarkStart w:id="1" w:name="m"/><w:r><w:t>text</w:t></w:r></w:p>',
    [0, 1],
    styles,
  );
  change('outdent');
  assert.equal(paragraphs()[0]!.numbering, undefined);
  assert.equal(paragraphs()[0]!.styleId, 'List');
  assert.equal(
    descendants(w.package(id).xml(w.package(id).mainPart)).find((n) => n.local === 'p')!
      .children[0]!.local,
    'pPr',
  );
  const explicit = setup(item('text', 0, '<w:ind w:left="1800"/>'));
  explicit.change('indent');
  assert.equal(explicit.paragraphs()[0]!.layout.left, 1800);
});
test('unsupported levels, protection, cell crossings and stale requests fail without partial mutations', () => {
  for (const body of [
    item('x', 1),
    item('x', 8),
    item('x', 0, '<w:sectPr/>'),
    item('x', 0, '<w:numPr><w:numId w:val="1"/></w:numPr>'),
    '<w:p><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:ins w:id="1"><w:r><w:t>x</w:t></w:r></w:ins></w:p>',
  ]) {
    const { w, id, change } = setup(body, [0, 1, 8]);
    const before = w.package(id);
    assert.throws(() => change('indent'));
    assert.equal(w.package(id), before);
  }
  const multi = setup(item('a', 0) + item('b', 1), [0, 1]);
  const before = multi.w.package(multi.id);
  assert.throws(() => multi.change('indent', 0, 3));
  assert.equal(multi.w.package(multi.id), before);
  const cells = setup(
    '<w:tbl><w:tr><w:tc>' + item('a') + '</w:tc><w:tc>' + item('b') + '</w:tc></w:tr></w:tbl>',
  );
  const original = cells.w.package(cells.id),
    r = cells.request(),
    story = cells.w.document(cells.id).model.stories[0]!,
    projected = projectStory(story, r.origin);
  assert.throws(() => cells.change('indent', 0, projected.paragraphs[1]!.to));
  assert.equal(cells.w.package(cells.id), original);
  cells.change('indent');
  assert.throws(
    () => cells.w.applyProjection({ ...r, from: 0, to: 0, text: '', listAction: 'outdent' }),
    /stale/,
  );
});
test('optimistic paragraph state supports rapid Enter through nested-list exit and queued typing', () => {
  let ps = [{ id: 'p', from: 0, to: 1, numbering: { numId: '1', level: 1 } }];
  ps = mapParagraphs(ps, 1, 1, '\n') as typeof ps;
  assert.deepEqual(
    ps.map((p) => [p.from, p.to, p.numbering?.level]),
    [
      [0, 1, 1],
      [2, 2, 1],
    ],
  );
  ps = mapListAction(ps, 2, 2, 'enter') as typeof ps;
  assert.equal(ps[1]!.numbering?.level, 0);
  ps = mapListAction(ps, 2, 2, 'enter') as typeof ps;
  assert.equal(ps[1]!.numbering, undefined);
  ps = mapParagraphs(ps, 2, 2, '\n') as typeof ps;
  assert.equal(ps.length, 3);
  assert.equal(ps[2]!.numbering, undefined);
  ps = mapParagraphs(ps, 3, 3, 'typed') as typeof ps;
  assert.equal(ps[2]!.to, 8);
});
