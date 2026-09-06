import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { Workspace } from '../src/engine/workspace';
import { readDocument, spans, plainText } from '../src/engine/document';
import { DocxPackage, comparePackages } from '../src/package/docx';
import { W } from '../src/package/xml';
import { fixture, paragraph, relation, relationships } from './helpers';

const texts = (ws: Workspace, id: string) => spans(ws.document(id).model).map((s) => s.text);
test('preview is pure, targeted edit preserves every surrounding byte, undo/redo restores snapshots', () => {
  const ws = new Workspace();
  const id = ws.open(
    'test.docx',
    fixture(
      `<w:p x:future="yes"><w:r><w:rPr><w:b/></w:rPr><w:t>hello</w:t></w:r><x:future a="1"/><w:r><w:t>world</w:t></w:r></w:p>`,
      { 'opaque.bin': new Uint8Array([1, 2, 3]) },
    ),
  );
  const original = ws.package(id);
  const span = spans(ws.document(id).model)[0]!;
  const proposal = ws.preview('insert', [
    { documentId: id, spanId: span.id, from: 5, to: 5, text: ' & 😀 ' },
  ]);
  assert.equal(ws.package(id), original);
  assert.equal(ws.document(id).dirty, false);
  ws.commit(proposal.id);
  assert.deepEqual(texts(ws, id), ['hello & 😀 ', 'world']);
  assert.equal(ws.document(id).dirty, true);
  assert.equal(
    ws.package(id).text('word/document.xml'),
    original
      .text('word/document.xml')
      .replace('<w:t>hello</w:t>', '<w:t xml:space="preserve">hello &amp; 😀 </w:t>'),
  );
  assert.deepEqual(
    comparePackages(original, ws.package(id))
      .filter((p) => p.status !== 'unchanged')
      .map((p) => p.part),
    ['word/document.xml'],
  );
  assert.equal(spans(ws.document(id).model)[0]!.id, span.id);
  ws.undo();
  assert.equal(ws.package(id), original);
  assert.equal(ws.document(id).dirty, false);
  ws.redo();
  assert.deepEqual(texts(ws, id), ['hello & 😀 ', 'world']);
  assert.deepEqual(textsFromSaved(ws, id), ['hello & 😀 ', 'world']);
});
function textsFromSaved(ws: Workspace, id: string) {
  return spans(readDocument(DocxPackage.open(ws.package(id).save()))).map((s) => s.text);
}

test('multi-document replace is one transaction; previews expire on state changes', () => {
  const ws = new Workspace();
  const a = ws.open('a.docx', fixture(paragraph('aaaa')));
  const b = ws.open('b.docx', fixture(paragraph('aa')));
  const preview = ws.previewReplace('aa', 'b');
  assert.equal(preview.edits.length, 2);
  ws.commit(preview.id);
  assert.deepEqual(texts(ws, a), ['bb']);
  assert.deepEqual(texts(ws, b), ['b']);
  ws.undo();
  assert.deepEqual(texts(ws, a), ['aaaa']);
  assert.deepEqual(texts(ws, b), ['aa']);
  ws.redo();
  const stale = ws.previewReplace('b', 'c');
  ws.undo();
  assert.throws(() => ws.commit(stale.id), /stale/);
});

test('failed staging never changes any document', () => {
  const ws = new Workspace();
  const a = ws.open('a.docx', fixture()),
    b = ws.open('b.docx', fixture());
  const pkg = ws.package(a);
  assert.throws(() =>
    ws.preview('invalid', [
      { documentId: a, spanId: spans(ws.document(a).model)[0]!.id, from: 0, to: 1, text: 'good' },
      { documentId: b, spanId: 'missing', from: 0, to: 1, text: 'bad' },
    ]),
  );
  assert.equal(ws.package(a), pkg);
  assert.equal(ws.canUndo, false);
});

test('commands reject overlapping, structural, and split-surrogate edits', () => {
  const ws = new Workspace();
  const id = ws.open('unicode.docx', fixture(paragraph('a😀b')));
  const spanId = spans(ws.document(id).model)[0]!.id;
  const edit = { documentId: id, spanId, from: 0, to: 1, text: '' };
  for (const text of ['\n', '\t', '\0', '\ud800'])
    assert.throws(() => ws.preview('invalid', [{ ...edit, text }]));
  assert.throws(() => ws.preview('invalid', [{ ...edit, from: 2, to: 2 }]));
  assert.throws(() => ws.preview('invalid', [edit, edit]));
  ws.commit(ws.preview('delete', [{ ...edit, from: 1, to: 3 }]).id);
  assert.deepEqual(texts(ws, id), ['ab']);
});

test('fields spanning paragraphs, revisions, content controls and extension markup stay protected', () => {
  const body =
    `<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:instrText>DATE</w:instrText><w:t>field one</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t>field two</w:t><w:fldChar w:fldCharType="end"/><w:t>editable</w:t></w:r><w:ins w:author="A"><w:r><w:t>inserted</w:t></w:r></w:ins><w:del><w:r><w:delText>deleted</w:delText></w:r></w:del><w:sdt><w:sdtContent>${paragraph('hidden')}</w:sdtContent></w:sdt><x:future>${paragraph('opaque')}</x:future></w:p>`;
  const ws = new Workspace();
  const id = ws.open('protected.docx', fixture(body));
  const items = spans(ws.document(id).model);
  assert.deepEqual(
    items.map((t) => [t.text, t.editable]),
    [
      ['field one', false],
      ['field two', false],
      ['editable', true],
      ['inserted', false],
      ['deleted', false],
    ],
  );
  const proposal = ws.previewReplace('field', 'changed');
  assert.equal(proposal.skipped, 2);
  assert.equal(proposal.edits.length, 0);
  assert.throws(
    () =>
      ws.preview('invalid', [{ documentId: id, spanId: items[0]!.id, from: 0, to: 0, text: 'x' }]),
    /Field/,
  );
  assert.ok(
    ws
      .document(id)
      .model.stories[0]!.tokens.some((t) => t.kind === 'code' && t.category === 'opaque'),
  );
});

test('formatting inheritance, tables, hyperlinks and notes have inspectable source bindings', () => {
  const body = `<w:tbl><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val="Heading"/></w:pPr><w:hyperlink r:id="link"><w:r><w:rPr><w:i/></w:rPr><w:t>inside table</w:t></w:r></w:hyperlink></w:p></w:tc></w:tr></w:tbl>`;
  const pkg = DocxPackage.open(
    fixture(body, {
      'word/_rels/document.xml.rels': relationships(
        relation('styles', 'styles', 'styles.xml') +
          relation('foot', 'footnotes', 'footnotes.xml') +
          relation('end', 'endnotes', 'endnotes.xml') +
          relation('link', 'hyperlink', 'https://example.test', true),
      ),
      'word/styles.xml': `<w:styles xmlns:w="${W}"><w:style w:styleId="Base"><w:rPr><w:b/></w:rPr></w:style><w:style w:styleId="Heading"><w:basedOn w:val="Base"/><w:rPr><w:color w:val="123456"/></w:rPr></w:style></w:styles>`,
      'word/footnotes.xml': `<w:footnotes xmlns:w="${W}"><w:footnote w:id="1">${paragraph('Footnote')}</w:footnote></w:footnotes>`,
      'word/endnotes.xml': `<w:endnotes xmlns:w="${W}"><w:endnote w:id="1">${paragraph('Endnote')}</w:endnote></w:endnotes>`,
    }),
  );
  const model = readDocument(pkg);
  assert.deepEqual(
    model.stories.map((s) => s.kind),
    ['body', 'footnote', 'endnote'],
  );
  const first = spans(model)[0]!;
  assert.deepEqual(first.direct, { i: 'on' });
  assert.deepEqual(first.inherited, { b: 'on', color: '123456' });
  assert.ok(model.stories[0]!.tokens.some((t) => t.kind === 'code' && t.label === 'tbl'));
});

test('docxfix review can be edited beside annotations without changing any review parts', async () => {
  const ws = new Workspace();
  const id = ws.open('review.docx', await readFile('fixtures/generated/review.docx'));
  const before = ws.package(id);
  ws.commit(ws.previewReplace('Replace this phrase', 'Edited phrase').id);
  const changed = comparePackages(before, ws.package(id)).filter((p) => p.status !== 'unchanged');
  assert.deepEqual(
    changed.map((p) => p.part),
    ['word/document.xml'],
  );
  assert.ok(ws.document(id).model.stories.some((s) => s.kind === 'comment'));
  assert.ok(textsFromSaved(ws, id).some((t) => t.includes('Edited phrase')));
});

test('docxfix headers and body are searched and replaced as one transaction', async () => {
  const ws = new Workspace();
  const id = ws.open('sections.docx', await readFile('fixtures/generated/sections.docx'));
  assert.equal(ws.search('Replace this phrase').length, 2);
  ws.commit(ws.previewReplace('Replace this phrase', 'Updated').id);
  assert.equal(ws.search('Updated').length, 2);
  assert.equal(ws.report(id).filter((p) => p.status === 'changed').length, 2);
  ws.undo();
  assert.equal(ws.search('Replace this phrase').length, 2);
});

test('search is literal and case folding preserves source offsets', () => {
  const ws = new Workspace();
  ws.open('a.docx', fixture(paragraph('İ a.b A.B')));
  assert.deepEqual(
    ws.search('a.b', false).map((m) => m.from),
    [2, 6],
  );
  assert.equal(ws.search('.').length, 2);
  assert.equal(ws.search('[').length, 0);
});

test('saved snapshot and undo dirty state, no-op transactions, observer isolation', () => {
  const ws = new Workspace();
  const id = ws.open('a.docx', fixture());
  ws.subscribe(() => {
    throw new Error('observer');
  });
  ws.commit(ws.previewReplace('Hello', 'Hi').id);
  const saved = ws.package(id);
  ws.markSaved(id, saved);
  assert.equal(ws.document(id).dirty, false);
  ws.undo();
  assert.equal(ws.document(id).dirty, true);
  ws.redo();
  assert.equal(ws.document(id).dirty, false);
  ws.commit(ws.previewReplace('Hi', 'Hi').id);
  ws.undo();
  assert.deepEqual(texts(ws, id), ['Hello world']);
});

test('signed packages, enforced protection and tracking disable editing', () => {
  const cases: Record<string, string>[] = [
    { '_xmlsignatures/sig.xml': '<signature/>' },
    {
      'word/_rels/document.xml.rels': relationships(relation('s', 'settings', 'settings.xml')),
      'word/settings.xml': `<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="1"/></w:settings>`,
    },
    {
      'word/_rels/document.xml.rels': relationships(relation('s', 'settings', 'settings.xml')),
      'word/settings.xml': `<w:settings xmlns:w="${W}"><w:trackRevisions/></w:settings>`,
    },
  ];
  for (const extras of cases) {
    const model = readDocument(DocxPackage.open(fixture(paragraph('protected'), extras)));
    assert.equal(spans(model)[0]!.editable, false);
  }
});

test('opaque content is represented in the plain-text review view', () => {
  const model = readDocument(DocxPackage.open(fixture(paragraph('a') + '<w:drawing/>')));
  assert.match(plainText(model.stories[0]!), /Opaque: w:drawing/);
});

test('fields starting inside opaque XML still protect their visible results', () => {
  const model = readDocument(
    DocxPackage.open(
      fixture(
        `<x:future><w:fldChar w:fldCharType="begin"/></x:future><w:p><w:r><w:t>result</w:t><w:fldChar w:fldCharType="end"/><w:t>after</w:t></w:r></w:p>`,
      ),
    ),
  );
  assert.deepEqual(
    spans(model).map((s) => [s.text, s.editable]),
    [
      ['result', false],
      ['after', true],
    ],
  );
});
