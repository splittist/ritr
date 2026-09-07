import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workspace } from '../src/engine/workspace';
import { spans, readDocument } from '../src/engine/document';
import { comparePackages, DocxPackage } from '../src/package/docx';
import { fixture, paragraph, relation, relationships } from './helpers';

const run = (text: string, props = '') => `<w:r>${props}<w:t>${text}</w:t></w:r>`;
const texts = (ws: Workspace, id: string) => spans(ws.document(id).model).map((s) => s.text);

test('cross-run replacement preserves source markup and other parts, reopens, and undoes atomically', () => {
  const ws = new Workspace();
  const body = `<w:p>${run('Before hel', '<w:rPr><w:b/></w:rPr>')}${run('lo')}${run(' world after', '<w:rPr><w:i/></w:rPr>')}</w:p>`;
  const a = ws.open('a.docx', fixture(body, { 'opaque.bin': new Uint8Array([7, 8, 9]) }));
  const b = ws.open('b.docx', fixture(body));
  const originals = [ws.package(a), ws.package(b)];
  const matches = ws.search('hello world');
  assert.equal(matches.length, 2);
  assert.deepEqual(
    matches[0]!.ranges.map(({ from, to }) => [from, to]),
    [
      [7, 10],
      [0, 2],
      [0, 6],
    ],
  );
  const preview = ws.previewReplace('hello world', 'new & 😀');
  assert.equal(preview.edits.length, 6);
  assert.equal(ws.package(a), originals[0]);
  ws.commit(preview.id);
  for (const [i, id] of [a, b].entries()) {
    assert.deepEqual(texts(ws, id), ['Before new & 😀', '', ' after']);
    assert.equal(
      ws.package(id).text('word/document.xml'),
      originals[i]!.text('word/document.xml')
        .replace('<w:t>Before hel</w:t>', '<w:t xml:space="preserve">Before new &amp; 😀</w:t>')
        .replace('<w:t>lo</w:t>', '<w:t xml:space="preserve"></w:t>')
        .replace('<w:t> world after</w:t>', '<w:t xml:space="preserve"> after</w:t>'),
    );
    assert.deepEqual(
      comparePackages(originals[i]!, ws.package(id))
        .filter((p) => p.status !== 'unchanged')
        .map((p) => p.part),
      ['word/document.xml'],
    );
    assert.deepEqual(
      spans(readDocument(DocxPackage.open(ws.package(id).save()))).map((s) => s.text),
      texts(ws, id),
    );
  }
  ws.undo();
  assert.equal(ws.package(a), originals[0]);
  assert.equal(ws.package(b), originals[1]);
  ws.redo();
  assert.deepEqual(texts(ws, b), ['Before new & 😀', '', ' after']);
});

test('adjacent matches can share spans; empty spans and literal Unicode case folding retain offsets', () => {
  const ws = new Workspace();
  const id = ws.open(
    'repeated.docx',
    fixture(`<w:p>${run('a')}${run('')}${run('aaa')}${run('aa')}</w:p>`),
  );
  assert.equal(ws.search('aa').length, 3);
  ws.commit(ws.previewReplace('aa', 'X').id);
  assert.deepEqual(texts(ws, id), ['X', '', 'X', 'X']);
  ws.undo();
  ws.commit(ws.previewReplace('aa', '').id);
  assert.deepEqual(texts(ws, id), ['', '', '', '']);
  const unicode = ws.open('unicode.docx', fixture(`<w:p>${run('İ😀A')}${run('.B😀')}</w:p>`));
  const [match] = ws.search('a.b😀', false);
  assert.equal(match!.documentId, unicode);
  assert.deepEqual(
    match!.ranges.map(({ from, to }) => [from, to]),
    [
      [3, 4],
      [0, 4],
    ],
  );
  assert.equal(ws.search('a.b😀').length, 0);
  assert.equal(ws.search('[').length, 0);
  ws.commit(ws.previewReplace('a.b😀', 'ok', false).id);
  assert.deepEqual(texts(ws, unicode), ['İ😀ok', '']);
});

test('an identical replacement keeps cross-run formatting and does not create undo history', () => {
  const ws = new Workspace();
  const id = ws.open(
    'same.docx',
    fixture(`<w:p>${run('hel')}${run('lo', '<w:rPr><w:b/></w:rPr>')}</w:p>`),
  );
  const original = ws.package(id);
  const preview = ws.previewReplace('hello', 'hello');
  assert.equal(preview.edits.length, 0);
  ws.commit(preview.id);
  assert.equal(ws.package(id), original);
  assert.equal(ws.canUndo, false);
});

test('search stops at structure, anchors, revisions, fields and opaque content', () => {
  const boundaries = [
    '</w:p><w:p>',
    '<w:r><w:tab/></w:r>',
    '<w:r><w:br/></w:r>',
    '<w:r><w:cr/></w:r>',
    '<w:bookmarkStart w:id="0" w:name="b"/>',
    '<w:bookmarkEnd w:id="0"/>',
    '<w:commentRangeStart w:id="0"/>',
    '<w:commentRangeEnd w:id="0"/>',
    '<w:r><w:commentReference w:id="0"/></w:r>',
    '<w:r><w:footnoteReference w:id="1"/></w:r>',
    '<x:future/>',
    '<w:r><w:drawing/></w:r>',
    '<w:sdt/>',
    '<w:ins w:id="1"/>',
    '<w:r><w:fldChar w:fldCharType="begin"/><w:fldChar w:fldCharType="end"/></w:r>',
  ];
  for (const boundary of boundaries) {
    const ws = new Workspace();
    ws.open('boundary.docx', fixture(`<w:p>${run('hel')}${boundary}${run('lo')}</w:p>`));
    assert.equal(ws.search('hello').length, 0, boundary);
  }
  const ws = new Workspace();
  ws.open(
    'links.docx',
    fixture(
      `<w:p>${run('out')}<w:hyperlink w:anchor="b">${run('hel')}${run('lo')}</w:hyperlink>${run('side')}</w:p>`,
    ),
  );
  assert.equal(ws.search('outhello').length, 0);
  assert.equal(ws.search('helloside').length, 0);
  assert.equal(ws.search('hello').length, 1);
  ws.open(
    'cells.docx',
    fixture(
      `<w:tbl><w:tr><w:tc>${paragraph('hel')}</w:tc><w:tc>${paragraph('lo')}</w:tc></w:tr></w:tbl>`,
    ),
  );
  assert.equal(ws.search('hello').length, 1);
});

test('a match touching protected formatting is skipped in full; revisions remain searchable inside their boundary', () => {
  const ws = new Workspace();
  const id = ws.open(
    'protected.docx',
    fixture(
      `<w:p>${run('hel')}${run('lo', '<w:rPr><w:rPrChange w:id="1"/></w:rPr>')}</w:p><w:p><w:ins w:id="2">${run('hel')}${run('lo')}</w:ins></w:p>`,
    ),
  );
  const matches = ws.search('hello');
  assert.equal(matches.length, 2);
  assert.ok(matches.every((m) => !m.editable && m.reason));
  const preview = ws.previewReplace('hello', 'changed');
  assert.equal(preview.skipped, 2);
  assert.deepEqual(preview.edits, []);
  ws.commit(preview.id);
  assert.deepEqual(texts(ws, id), ['hel', 'lo', 'hel', 'lo']);
});

test('cross-run matches in headers share the body transaction and stale previews are refused', () => {
  const ws = new Workspace();
  const body = `<w:p>${run('hel')}${run('lo')}</w:p>`;
  const id = ws.open(
    'stories.docx',
    fixture(body, {
      'word/_rels/document.xml.rels': relationships(relation('header', 'header', 'header.xml')),
      'word/header.xml': `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${body}</w:hdr>`,
    }),
  );
  assert.equal(new Set(ws.search('hello').map((m) => m.storyId)).size, 2);
  ws.commit(ws.previewReplace('hello', 'updated').id);
  assert.deepEqual(texts(ws, id), ['updated', '', 'updated', '']);
  const stale = ws.previewReplace('updated', 'stale');
  ws.undo();
  assert.throws(() => ws.commit(stale.id), /stale/);
  assert.deepEqual(texts(ws, id), ['hel', 'lo', 'hel', 'lo']);
});
