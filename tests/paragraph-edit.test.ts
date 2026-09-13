import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workspace } from '../src/engine/workspace';
import { spans } from '../src/engine/document';
import { projectTokens, tokenId } from '../src/engine/projection';
import { descendants, child } from '../src/package/xml';
import { fixture, paragraph, relation, relationships } from './helpers';

function setup(body = paragraph('hello') + paragraph('world')) {
  const w = new Workspace(),
    id = w.open('paragraphs.docx', fixture(body, { 'custom.bin': new Uint8Array([1, 2, 3]) }));
  const edit = (from: number, to: number, text: string, typingSpan?: string) => {
    const document = w.document(id),
      story = document.model.stories[0]!;
    w.applyProjection({
      documentId: id,
      storyId: story.id,
      origin: tokenId(story.tokens[0]!),
      expectedRevision: document.revision,
      from,
      to,
      text,
      typingSpan,
    });
  };
  const text = () => projectTokens(w.document(id).model.stories[0]!.tokens).text;
  return { w, id, edit, text };
}
test('split and join preserve runs, paragraph identity, other parts, and whole-operation undo/redo', () => {
  const { w, id, edit, text } = setup(
    '<w:p w:rsidR="12345678"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>hello</w:t></w:r><w:r><w:t> tail</w:t></w:r></w:p>' +
      paragraph('world'),
  );
  const original = w.package(id),
    initial = w.document(id).model.stories[0]!,
    ids = spans(w.document(id).model).map((s) => s.id);
  edit(2, 2, '\n');
  assert.equal(text(), 'he\nllo tail\nworld');
  assert.equal(w.document(id).model.stories[0]!.paragraphs[0]!.id, initial.paragraphs[0]!.id);
  assert.equal(spans(w.document(id).model)[0]!.id, ids[0]);
  assert.equal(spans(w.document(id).model)[2]!.id, ids[1]);
  assert.equal(spans(w.document(id).model)[1]!.direct.b, 'on');
  const ps = child(w.package(id).xml(w.package(id).mainPart), 'body')!.children.filter(
    (n) => n.local === 'p',
  );
  assert.deepEqual(
    ps.slice(0, 2).map((p) => child(p, 'pPr')?.children[0]?.local),
    ['jc', 'jc'],
  );
  assert.equal(
    Object.keys(ps[1]!.attrs).length,
    0,
    'New paragraph does not duplicate source IDs/attributes',
  );
  assert.deepEqual(
    w
      .report(id)
      .filter((p) => p.status !== 'unchanged')
      .map((p) => p.part),
    ['word/document.xml'],
  );
  w.undo();
  assert.equal(w.package(id), original);
  w.redo();
  edit(2, 3, '');
  assert.equal(text(), 'hello tail\nworld');
  assert.ok(spans(w.document(id).model).some((s) => s.id === ids[1]));
  const reopened = new Workspace(),
    copy = reopened.open('copy', w.package(id).save());
  assert.equal(projectTokens(reopened.document(copy).model.stories[0]!.tokens).text, text());
});
test('Enter at both edges and multi-paragraph paste work, including empty paragraphs', () => {
  const { w, id, edit, text } = setup('<w:p/>' + paragraph('tail'));
  edit(0, 0, 'a\nb\nc');
  assert.ok(text().startsWith('a\nb\nc\ntail'));
  w.undo();
  assert.ok(text().startsWith('\ntail'));
  edit(0, 0, '\n');
  assert.ok(text().startsWith('\n\ntail'));
  edit(0, 1, '');
  assert.ok(text().startsWith('\ntail'));
  edit(0, 0, 'hello');
  edit(5, 5, '\n');
  assert.ok(text().startsWith('hello\n\ntail'));
  assert.equal(
    new Set(descendants(w.package(id).xml(w.package(id).mainPart)).map((n) => n.id)).size,
    descendants(w.package(id).xml(w.package(id).mainPart)).length,
  );
});
test('selection replacement across adjacent paragraphs is atomic and retains surviving formatting', () => {
  const { w, id, edit, text } = setup(
    paragraph('abc') +
      '<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>DEF</w:t></w:r></w:p>' +
      paragraph('ghi'),
  );
  edit(1, 9, 'X\nY');
  assert.ok(text().startsWith('aX\nYhi'));
  w.undo();
  assert.ok(text().startsWith('abc\nDEF\nghi'));
  edit(2, 5, '');
  assert.ok(text().startsWith('abEF\nghi'));
  assert.equal(spans(w.document(id).model).find((s) => s.text === 'EF')!.direct.i, 'on');
});
test('splits and joins refuse section boundaries, fields, revisions and cell crossings without changing snapshots', () => {
  for (const body of [
    '<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>hello</w:t></w:r></w:p>',
    '<w:p><w:ins w:id="1"><w:r><w:t>hello</w:t></w:r></w:ins></w:p>',
    '<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:t>hello</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>',
  ]) {
    const { w, id, edit } = setup(body);
    const before = w.package(id);
    assert.throws(() => edit(2, 2, '\n'));
    assert.equal(w.package(id), before);
  }
  const { w, id } = setup(
    '<w:tbl><w:tr><w:tc>' +
      paragraph('a') +
      paragraph('b') +
      '</w:tc><w:tc>' +
      paragraph('c') +
      '</w:tc></w:tr></w:tbl>',
  );
  const story = w.document(id).model.stories[0]!,
    first = story.tokens.find((t) => t.kind === 'code' && t.role === 'paragraph-start')!;
  const origin = tokenId(first),
    before = w.package(id);
  w.applyProjection({
    documentId: id,
    storyId: story.id,
    origin,
    expectedRevision: 0,
    from: 1,
    to: 2,
    text: '',
  });
  assert.notEqual(w.package(id), before, 'Join within a cell works');
  assert.throws(() =>
    w.applyProjection({
      documentId: id,
      storyId: story.id,
      origin,
      expectedRevision: 1,
      from: 2,
      to: 6,
      text: '',
    }),
  );
});

test('typing after deletion retains the deleted run format and undo/redo restores logical selections', () => {
  const { w, id, edit } = setup(
    '<w:p><w:r><w:t>ab</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>CD</w:t></w:r></w:p>',
  );
  edit(2, 4, '');
  edit(2, 2, 'X');
  assert.deepEqual(
    spans(w.document(id).model).map((s) => s.text),
    ['ab', 'X'],
  );
  assert.equal(spans(w.document(id).model)[1]!.direct.b, 'on');
  w.undo();
  assert.equal(w.selection?.anchor, 2);
  assert.equal(w.selection?.head, 2);
  w.undo();
  assert.equal(w.selection?.anchor, 2);
  assert.equal(w.selection?.head, 4);
  w.redo();
  assert.equal(w.selection?.anchor, 2);
  assert.equal(w.selection?.head, 2);
});

test('text adjacent to anchors stays on the selected side and stale edits cannot mutate the document', () => {
  const { w, id, edit, text } = setup(
    '<w:p><w:r><w:t>ab</w:t></w:r><w:bookmarkStart w:id="1" w:name="x"/><w:r><w:t>CD</w:t></w:r></w:p>',
  );
  edit(2, 2, 'X', spans(w.document(id).model)[1]!.id);
  assert.equal(text(), 'abXCD');
  assert.deepEqual(
    spans(w.document(id).model).map((s) => s.text),
    ['ab', 'XCD'],
  );
  const story = w.document(id).model.stories[0]!,
    before = w.package(id);
  assert.throws(
    () =>
      w.applyProjection({
        documentId: id,
        storyId: story.id,
        origin: tokenId(story.tokens[0]!),
        from: 0,
        to: 0,
        text: 'stale',
        expectedRevision: 0,
      }),
    /stale/,
  );
  assert.equal(w.package(id), before);
});

test('empty paragraphs inherit an existing empty run or paragraph-mark typing properties', () => {
  for (const body of [
    '<w:p><w:r><w:rPr><w:b/></w:rPr></w:r></w:p>',
    '<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr></w:p>',
  ]) {
    const { w, id, edit, text } = setup(body);
    edit(0, 0, 'bold');
    assert.ok(text().startsWith('bold'));
    assert.equal(spans(w.document(id).model)[0]!.direct.b, 'on');
    edit(2, 2, '\n');
    assert.deepEqual(
      spans(w.document(id).model).map((s) => s.direct.b),
      ['on', 'on'],
    );
  }
});

test('the final paragraph has no phantom line and supports split/join at the document end', () => {
  const { w, id, edit, text } = setup(paragraph('hello'));
  assert.equal(text(), 'hello');
  edit(text().length, text().length, '\n');
  assert.equal(text(), 'hello\n');
  assert.equal(w.document(id).model.stories[0]!.paragraphs.length, 2);
  edit(text().length - 1, text().length, '');
  assert.equal(text(), 'hello');
  assert.equal(w.document(id).model.stories[0]!.paragraphs.length, 1);
  edit(text().length, text().length, ' world');
  assert.equal(text(), 'hello world');
});

test('split and join preserve bookmarks, proofing markers, comments and run objects away from the caret', () => {
  const markers =
    '<w:bookmarkStart w:id="1" w:name="mark"/><w:proofErr w:type="spellStart"/><w:commentRangeStart w:id="2"/>';
  const endings =
    '<w:commentRangeEnd w:id="2"/><w:proofErr w:type="spellEnd"/><w:bookmarkEnd w:id="1"/>';
  for (const object of [
    '<w:tab/>',
    '<w:br/>',
    '<w:lastRenderedPageBreak/>',
    '<w:drawing><x:picture/></w:drawing>',
    '<w:footnoteReference w:id="3"/>',
  ]) {
    const { w, id, edit, text } = setup(
      '<w:p>' +
        markers +
        '<w:r><w:rPr><w:b/></w:rPr><w:t>ab</w:t>' +
        object +
        '<w:t>CD</w:t></w:r>' +
        endings +
        '</w:p>',
    );
    const before = w.package(id);
    const preserved = descendants(before.xml(before.mainPart))
      .filter((n) =>
        [
          'bookmarkStart',
          'bookmarkEnd',
          'proofErr',
          'commentRangeStart',
          'commentRangeEnd',
          'tab',
          'br',
          'lastRenderedPageBreak',
          'drawing',
          'footnoteReference',
        ].includes(n.local),
      )
      .map((n) => ({ id: n.id, xml: before.text(before.mainPart).slice(n.start, n.end) }));
    assert.equal(text(), 'ab\ufffcCD');
    // The right-hand text at an object boundary owns the split, not the preceding run text.
    edit(3, 3, '\n');
    assert.equal(text(), 'ab\ufffc\nCD');
    assert.equal(spans(w.document(id).model).at(-1)!.direct.b, 'on');
    edit(3, 4, '');
    assert.equal(text(), 'ab\ufffcCD');
    for (const item of preserved) {
      const node = descendants(w.package(id).xml(before.mainPart)).find((n) => n.id === item.id)!;
      assert.ok(node);
      assert.equal(w.package(id).text(before.mainPart).slice(node.start, node.end), item.xml);
    }
    assert.deepEqual(
      w
        .report(id)
        .filter((p) => p.status !== 'unchanged')
        .map((p) => p.part),
      ['word/document.xml'],
    );
    w.undo();
    assert.equal(text(), 'ab\ufffc\nCD');
    w.undo();
    assert.equal(w.package(id), before);
  }
});

test('hidden boundary markers survive joins and do not contribute caret offsets', () => {
  const { w, id, edit, text } = setup(
    '<w:p><w:r><w:t>left</w:t></w:r><w:bookmarkStart w:id="1" w:name="m"/></w:p><w:p><w:bookmarkEnd w:id="1"/><w:r><w:t>right</w:t></w:r></w:p>',
  );
  assert.equal(text(), 'left\nright');
  edit(4, 5, '');
  assert.equal(text(), 'leftright');
  const right = spans(w.document(id).model).find((s) => s.text === 'right')!;
  edit(4, 4, '\n', right.id);
  assert.equal(text(), 'left\nright');
  const ps = child(w.package(id).xml(w.package(id).mainPart), 'body')!.children.filter(
    (n) => n.local === 'p',
  );
  assert.ok(
    ps[0]!.children.some((n) => n.local === 'bookmarkEnd'),
    'A split after the marker keeps it on the left',
  );
});

test('repeated split/join at run edges preserves whole runs without empty or nested runs', () => {
  for (const offset of [0, 1, 2, 3, 4]) {
    const { w, id, edit, text } = setup(
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>ab</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>cd</w:t></w:r></w:p>',
    );
    const initial = spans(w.document(id).model).map((s) => s.id);
    for (let i = 0; i < 5; i++) {
      edit(offset, offset, '\n');
      edit(offset, offset + 1, '');
      assert.equal(text(), 'abcd');
      const root = w.package(id).xml(w.package(id).mainPart);
      const runs = descendants(root).filter((n) => n.local === 'r');
      assert.equal(runs.length, [1, 3].includes(offset) ? 3 : 2);
      for (const run of runs) {
        assert.ok(run.children.some((n) => n.local === 't' && n.text.length));
        assert.ok(
          !descendants(run)
            .slice(1)
            .some((n) => n.local === 'r'),
          'Runs must not nest',
        );
      }
      for (const spanId of initial)
        assert.ok(spans(w.document(id).model).some((s) => s.id === spanId));
    }
  }
});

test('empty split paragraphs carry typing formatting without introducing empty runs', () => {
  for (const offset of [0, 4]) {
    const { w, id, edit, text } = setup(
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:p>',
    );
    edit(offset, offset, '\n');
    assert.equal(
      descendants(w.package(id).xml(w.package(id).mainPart)).filter((n) => n.local === 'r').length,
      1,
    );
    const caret = offset === 0 ? 0 : 5;
    edit(caret, caret, 'X');
    assert.equal(text(), offset === 0 ? 'X\nbold' : 'bold\nX');
    assert.equal(spans(w.document(id).model).find((s) => s.text === 'X')!.direct.b, 'on');
  }
  const { w, id, edit, text } = setup('<w:p/>');
  for (let i = 0; i < 4; i++) edit(i, i, '\n');
  for (let i = 0; i < 4; i++) edit(0, 1, '');
  assert.equal(text(), '');
  assert.equal(
    descendants(w.package(id).xml(w.package(id).mainPart)).filter((n) => n.local === 'r').length,
    0,
  );
});

test('splitting between text nodes within a run adds no empty text elements', () => {
  const { w, id, edit, text } = setup(
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>ab</w:t><w:t>cd</w:t></w:r></w:p>',
  );
  edit(2, 2, '\n');
  assert.equal(text(), 'ab\ncd');
  assert.deepEqual(
    spans(w.document(id).model).map((s) => s.text),
    ['ab', 'cd'],
  );
});

test('Enter at paragraph end uses a valid next style, while middle splits, paste and lists retain their style', () => {
  for (const scenario of ['end', 'middle', 'paste', 'list', 'missing'] as const) {
    const w = new Workspace();
    const body =
      '<w:p><w:pPr><w:pStyle w:val="Heading"/>' +
      (scenario === 'list' ? '<w:numPr><w:numId w:val="1"/></w:numPr>' : '') +
      '</w:pPr><w:r><w:t>title</w:t></w:r></w:p>';
    const styles =
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading"><w:next w:val="' +
      (scenario === 'missing' ? 'Missing' : 'Body') +
      '"/><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Body"/></w:styles>';
    const id = w.open(
      'next-style',
      fixture(body, {
        'word/styles.xml': styles,
        'word/_rels/document.xml.rels': relationships(relation('styles', 'styles', 'styles.xml')),
      }),
    );
    const before = w.package(id),
      story = w.document(id).model.stories[0]!;
    const from = scenario === 'middle' ? 2 : 5;
    w.applyProjection({
      documentId: id,
      storyId: story.id,
      origin: tokenId(story.tokens[0]!),
      expectedRevision: 0,
      from,
      to: from,
      text: '\n',
      enter: scenario !== 'paste',
    });
    const paragraphs = w.document(id).model.stories[0]!.paragraphs;
    assert.equal(paragraphs[0]!.styleId, 'Heading');
    assert.equal(paragraphs[1]!.styleId, scenario === 'end' ? 'Body' : 'Heading', scenario);
    assert.equal(w.package(id).text('word/styles.xml'), styles);
    w.undo();
    assert.equal(w.package(id), before);
    w.redo();
    assert.equal(
      w.document(id).model.stories[0]!.paragraphs[1]!.styleId,
      scenario === 'end' ? 'Body' : 'Heading',
    );
  }
});
