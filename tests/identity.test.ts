import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DocxPackage, comparePackages } from '../src/package/docx';
import {
  child,
  descendants,
  textPatch,
  patchXml,
  type XmlPatch,
  type XmlNode,
} from '../src/package/xml';
import { compareXmlIdentities } from '../src/package/xml-identity';
import { readDocument, spans } from '../src/engine/document';
import { Workspace } from '../src/engine/workspace';
import { fixture, paragraph, relation, relationships } from './helpers';

const nodes = (pkg: DocxPackage) => descendants(pkg.xml(pkg.mainPart));
const body = (pkg: DocxPackage) => child(pkg.xml(pkg.mainPart), 'body')!;
const apply = (pkg: DocxPackage, patches: readonly XmlPatch[]) =>
  pkg.withXmlPatches(new Map([[pkg.mainPart, patches]]));
const ids = (root: XmlNode) => descendants(root).map((node) => node.id);

test('insertion and deletion preserve identities of identical neighboring content and opaque nodes', () => {
  const before = DocxPackage.open(
    fixture(paragraph('same') + '<x:future><x:child/></x:future>' + paragraph('same')),
  );
  const initial = nodes(before);
  const first = body(before).children[0]!;
  const inserted = apply(before, [
    { start: first.start, end: first.start, replacement: paragraph('same') },
  ]);
  const report = compareXmlIdentities(before.xml(before.mainPart), inserted.xml(inserted.mainPart));
  assert.deepEqual(
    report.retained,
    initial.map((node) => node.id),
  );
  assert.equal(report.created.length, 3);
  assert.deepEqual(report.removed, []);
  for (const node of initial) {
    const current = nodes(inserted).find((n) => n.id === node.id)!;
    assert.equal(current.local, node.local);
    if (node.start >= first.start)
      assert.equal(current.start, node.start + paragraph('same').length);
  }
  const originalFirst = nodes(inserted).find((node) => node.id === first.id)!;
  const removed = apply(inserted, [
    { start: originalFirst.start, end: originalFirst.end, replacement: '' },
  ]);
  assert.deepEqual(
    compareXmlIdentities(inserted.xml(inserted.mainPart), removed.xml(removed.mainPart)).removed,
    ids(originalFirst),
  );
  assert.equal(spans(readDocument(removed)).length, 2);
  assert.equal(new Set(nodes(removed).map((node) => node.id)).size, nodes(removed).length);
  assert.equal(before.text(before.mainPart), DocxPackage.open(before.save()).text(before.mainPart));
});

test('splitting and joining wrappers retain reparented runs without ordinal or text matching', () => {
  const before = DocxPackage.open(
    fixture('<w:p><w:r><w:t>left</w:t></w:r><w:r><w:t>right</w:t></w:r></w:p>' + paragraph('tail')),
  );
  const p = body(before).children[0]!;
  const right = p.children[1]!;
  const split = apply(before, [
    { start: right.start, end: right.start, replacement: '</w:p><w:p>' },
  ]);
  const splitParagraphs = body(split).children.filter((node) => node.local === 'p');
  assert.equal(splitParagraphs[0]!.id, p.id);
  assert.equal(splitParagraphs[1]!.children[0]!.id, right.id);
  assert.deepEqual(
    compareXmlIdentities(before.xml(before.mainPart), split.xml(split.mainPart)).removed,
    [],
  );
  assert.equal(
    compareXmlIdentities(before.xml(before.mainPart), split.xml(split.mainPart)).created.length,
    1,
  );
  const joined = apply(split, [
    { start: splitParagraphs[0]!.closeStart, end: splitParagraphs[1]!.openEnd, replacement: '' },
  ]);
  assert.equal(joined.text(joined.mainPart), before.text(before.mainPart));
  assert.deepEqual(ids(joined.xml(joined.mainPart)), ids(before.xml(before.mainPart)));
  assert.deepEqual(
    compareXmlIdentities(split.xml(split.mainPart), joined.xml(joined.mainPart)).removed,
    [splitParagraphs[1]!.id],
  );
  assert.deepEqual(
    spans(readDocument(joined)).map((span) => span.text),
    ['left', 'right', 'tail'],
  );
});

test('explicit retention carries rewritten split nodes; subsequent text edits keep those bindings', () => {
  const before = DocxPackage.open(fixture(paragraph('hello') + paragraph('tail')));
  const p = body(before).children[0]!;
  const replacement = paragraph('hel') + paragraph('lo');
  const split = apply(before, [
    {
      start: p.start,
      end: p.end,
      replacement,
      retain: [
        { nodeId: p.id, offset: 0 },
        { nodeId: p.children[0]!.id, offset: 5 },
        { nodeId: p.children[0]!.children[0]!.id, offset: 10 },
      ],
    },
  ]);
  const newSpans = spans(readDocument(split));
  assert.equal(newSpans[0]!.id, p.children[0]!.children[0]!.id);
  assert.notEqual(newSpans[1]!.id, newSpans[0]!.id);
  assert.equal(newSpans[2]!.id, spans(readDocument(before))[1]!.id);
  const selected = nodes(split).find((node) => node.id === newSpans[1]!.id)!;
  const edited = apply(split, [textPatch(split.text(split.mainPart), selected, '😀 & later')]);
  assert.deepEqual(ids(edited.xml(edited.mainPart)), ids(split.xml(split.mainPart)));
  assert.deepEqual(
    spans(readDocument(edited)).map((span) => span.text),
    ['hel', '😀 & later', 'tail'],
  );
  assert.ok(
    comparePackages(split, edited)
      .filter((part) => part.status !== 'unchanged')
      .every((part) => part.part === split.mainPart),
  );
  assert.deepEqual(
    spans(readDocument(DocxPackage.open(edited.save()))).map((span) => span.text),
    ['hel', '😀 & later', 'tail'],
  );
});

test('explicit moves retain a removed subtree at its declared destination', () => {
  const before = DocxPackage.open(fixture(paragraph('move') + paragraph('stay')));
  const first = body(before).children[0]!;
  const last = body(before).children[1]!;
  const source = before.text(before.mainPart);
  const moved = apply(before, [
    {
      start: last.end,
      end: last.end,
      replacement: source.slice(first.start, first.end),
      retain: descendants(first).map((node) => ({
        nodeId: node.id,
        offset: node.start - first.start,
      })),
    },
    { start: first.start, end: first.end, replacement: '' },
  ]);
  assert.deepEqual(
    spans(readDocument(moved)).map((span) => span.text),
    ['stay', 'move'],
  );
  assert.deepEqual(
    new Set(ids(moved.xml(moved.mainPart))),
    new Set(ids(before.xml(before.mainPart))),
  );
  const mapped = body(moved).children[1]!;
  assert.equal(mapped.id, first.id);
  assert.equal(
    moved.text(moved.mainPart).slice(mapped.start, mapped.end),
    source.slice(first.start, first.end),
  );
});

test('invalid or ambiguous identity claims are refused before a snapshot can be published', () => {
  const before = DocxPackage.open(fixture(paragraph('a') + paragraph('b')));
  const first = body(before).children[0]!;
  const second = body(before).children[1]!;
  const replacement = paragraph('x');
  const base = { start: first.start, end: first.end, replacement };
  for (const retain of [
    [{ nodeId: 'missing', offset: 0 }],
    [{ nodeId: second.id, offset: 0 }],
    [{ nodeId: first.id, offset: 1 }],
    [{ nodeId: first.id, offset: 5 }],
    [{ nodeId: first.id, offset: -1 }],
    [{ nodeId: first.id, offset: 0.5 }],
    [{ nodeId: first.id, offset: replacement.length }],
    [
      { nodeId: first.id, offset: 0 },
      { nodeId: first.id, offset: 0 },
    ],
  ])
    assert.throws(() => apply(before, [{ ...base, retain }]), /retained|Retained|identity/);
  assert.throws(
    () =>
      apply(before, [
        { ...base, replacement: '<x:p/>', retain: [{ nodeId: first.id, offset: 0 }] },
      ]),
    /target/,
  );
  assert.throws(
    () =>
      apply(before, [
        {
          ...base,
          replacement: paragraph('x') + paragraph('y'),
          retain: [
            { nodeId: first.id, offset: 0 },
            { nodeId: first.id, offset: paragraph('x').length },
          ],
        },
      ]),
    /Duplicate/,
  );
  const both = { start: first.start, end: second.end, replacement };
  assert.throws(
    () =>
      apply(before, [
        {
          ...both,
          retain: [
            { nodeId: first.id, offset: 0 },
            { nodeId: second.id, offset: 0 },
          ],
        },
      ]),
    /Duplicate/,
  );
  assert.deepEqual(
    spans(readDocument(before)).map((span) => span.text),
    ['a', 'b'],
  );
});

test('raw replacement retires changed-part bindings; tables and callers cannot mutate snapshot identities', () => {
  const before = DocxPackage.open(
    fixture(paragraph('a'), { 'custom.xml': '<root><item/></root>' }),
  );
  assert.equal(before.withXml(new Map([[before.mainPart, before.text(before.mainPart)]])), before);
  const raw = before.withXml(
    new Map([[before.mainPart, before.text(before.mainPart).replace('>a<', '>b<')]]),
  );
  const report = compareXmlIdentities(before.xml(before.mainPart), raw.xml(raw.mainPart));
  assert.equal(report.retained.length, 0);
  assert.deepEqual(report.removed, ids(before.xml(before.mainPart)));
  assert.deepEqual(ids(raw.xml('custom.xml')), ids(before.xml('custom.xml')));
  const span = nodes(raw).find((node) => node.local === 't')!;
  const patch = textPatch(raw.text(raw.mainPart), span, 'changed');
  const edited = apply(raw, [patch]);
  (patch.retain as { nodeId: string; offset: number }[])[0]!.nodeId = 'corrupt';
  nodes(edited)[0]!.id = 'corrupt';
  assert.deepEqual(ids(edited.xml(edited.mainPart)), ids(raw.xml(raw.mainPart)));
  assert.deepEqual(edited.bytes('custom.xml'), before.bytes('custom.xml'));
});

test('new identities are unique across branches, while workspace text undo/redo restores identity tables', () => {
  const bytes = fixture(paragraph('before'));
  const base = DocxPackage.open(bytes);
  const at = body(base).openEnd;
  const patches = [{ start: at, end: at, replacement: paragraph('new') }];
  const a = apply(base, patches),
    b = apply(base, patches);
  const createdA = compareXmlIdentities(base.xml(base.mainPart), a.xml(a.mainPart)).created;
  const createdB = compareXmlIdentities(base.xml(base.mainPart), b.xml(b.mainPart)).created;
  assert.ok(createdA.every((id) => !createdB.includes(id)));
  const ws = new Workspace();
  const id = ws.open('history.docx', bytes);
  const original = ws.package(id);
  const initialIds = ids(original.xml(original.mainPart));
  ws.commit(ws.previewReplace('before', 'a much longer text 😀').id);
  const applied = ws.package(id);
  assert.deepEqual(ids(applied.xml(applied.mainPart)), initialIds);
  ws.undo();
  assert.equal(ws.package(id), original);
  ws.redo();
  assert.equal(ws.package(id), applied);
  ws.commit(ws.previewReplace('longer', 'short').id);
  assert.deepEqual(ids(ws.package(id).xml(original.mainPart)), initialIds);
});

test('patch bounds reject nonintegral and ambiguous edits; insertion at a tag end keeps its identity', () => {
  for (const patch of [
    { start: NaN, end: 0, replacement: '' },
    { start: 0, end: Infinity, replacement: '' },
    { start: 0.5, end: 1, replacement: '' },
  ])
    assert.throws(() => patchXml('<a/>', [patch]), /invalid/);
  assert.throws(
    () =>
      patchXml('<a/>', [
        { start: 0, end: 0, replacement: ' ' },
        { start: 0, end: 0, replacement: ' ' },
      ]),
    /invalid/,
  );
  const before = DocxPackage.open(fixture(paragraph('😀')));
  const p = body(before).children[0]!;
  const after = apply(before, [
    { start: p.openEnd, end: p.openEnd, replacement: '<w:r><w:t/></w:r>' },
  ]);
  assert.equal(body(after).children[0]!.id, p.id);
  assert.deepEqual(
    compareXmlIdentities(before.xml(before.mainPart), after.xml(after.mainPart)).removed,
    [],
  );
});

test('table and header source bindings survive a body insertion; multipart failure leaves all snapshots intact', () => {
  const before = DocxPackage.open(
    fixture(paragraph('lead') + `<w:tbl><w:tr><w:tc>${paragraph('cell')}</w:tc></w:tr></w:tbl>`, {
      'word/_rels/document.xml.rels': relationships(relation('header', 'header', 'header.xml')),
      'word/header.xml':
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        paragraph('header') +
        '</w:hdr>',
      'opaque.bin': new Uint8Array([7, 8, 9]),
    }),
  );
  const initial = readDocument(before);
  const table = initial.stories[0]!.blocks.find((block) => block.kind === 'table')!;
  const at = body(before).openEnd;
  const after = apply(before, [{ start: at, end: at, replacement: paragraph('inserted') }]);
  const current = readDocument(after);
  const updatedTable = current.stories[0]!.blocks.find((block) => block.kind === 'table')!;
  assert.equal(updatedTable.id, table.id);
  for (const span of spans(initial)) {
    const updated = spans(current).find((s) => s.id === span.id)!;
    assert.equal(updated.text, span.text);
    assert.equal(updated.source.nodeId, span.source.nodeId);
    const xmlNode = descendants(after.xml(updated.source.part)).find(
      (node) => node.id === updated.id,
    )!;
    assert.equal(updated.source.start, xmlNode.start);
    assert.equal(updated.source.end, xmlNode.end);
  }
  assert.deepEqual(after.bytes('word/header.xml'), before.bytes('word/header.xml'));
  assert.deepEqual(after.bytes('opaque.bin'), before.bytes('opaque.bin'));
  const bytes = after.save();
  const initialIds = ids(after.xml(after.mainPart));
  assert.throws(() =>
    after.withXmlPatches(
      new Map([
        [after.mainPart, [{ start: at, end: at, replacement: paragraph('never published') }]],
        ['word/header.xml', [{ start: 0, end: 1, replacement: '' }]],
      ]),
    ),
  );
  assert.deepEqual(after.save(), bytes);
  assert.deepEqual(ids(after.xml(after.mainPart)), initialIds);
  assert.deepEqual(
    spans(readDocument(after)).map((span) => span.text),
    ['inserted', 'lead', 'cell', 'header'],
  );
});

test('namespace changes retire affected identities even when descendant tag bytes are unchanged', () => {
  const before = DocxPackage.open(
    fixture(paragraph('text'), {
      'custom.xml': '<root xmlns:x="urn:old"><x:item/><keep/></root>',
    }),
  );
  const xml = before.text('custom.xml');
  const at = xml.indexOf('urn:old');
  const after = before.withXmlPatches(
    new Map([
      [
        'custom.xml',
        [
          {
            start: at,
            end: at + 'urn:old'.length,
            replacement: 'urn:new',
          },
        ],
      ],
    ]),
  );
  const old = descendants(before.xml('custom.xml'));
  const next = descendants(after.xml('custom.xml'));
  assert.notEqual(next[0]!.id, old[0]!.id);
  assert.notEqual(next[1]!.id, old[1]!.id);
  assert.equal(next[2]!.id, old[2]!.id);
  assert.deepEqual(
    compareXmlIdentities(before.xml('custom.xml'), after.xml('custom.xml')).removed,
    [old[0]!.id, old[1]!.id],
  );
});
