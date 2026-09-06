import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { DocxPackage, comparePackages } from '../src/package/docx';
import { descendants, parseXml, textPatch, patchXml, W } from '../src/package/xml';
import { readZip } from '../src/package/zip';
import { fixture, paragraph, relation, relationships } from './helpers';

for (const name of ['plain', 'review', 'sections'])
  test(`docxfix ${name}: whole-file and all-part no-edit identity`, async () => {
    const original = await readFile(`fixtures/generated/${name}.docx`);
    const pkg = DocxPackage.open(original);
    assert.deepEqual(pkg.save(), new Uint8Array(original));
    assert.ok(
      comparePackages(pkg, DocxPackage.open(pkg.save())).every((p) => p.status === 'unchanged'),
    );
    assert.equal(pkg.diagnostics.filter((d) => d.severity === 'error').length, 0);
  });

test('XML bindings capture exact lexical regions, including self-closing and Unicode', () => {
  const source = `<?xml version="1.0"?><a xmlns:w="${W}">😀<!-- hi --><w:t a='>'>a&amp;é</w:t><w:t /></a>`;
  const root = parseXml(source, 'part');
  const nodes = descendants(root).filter((n) => n.local === 't');
  assert.equal(source.slice(nodes[0]!.start, nodes[0]!.end), "<w:t a='>'>a&amp;é</w:t>");
  assert.equal(source.slice(nodes[1]!.start, nodes[1]!.end), '<w:t />');
  const edited = patchXml(source, [
    textPatch(source, nodes[0]!, ' <&😀 '),
    textPatch(source, nodes[1]!, 'new'),
  ]);
  assert.match(edited, /a='>' xml:space="preserve"> &lt;&amp;😀 <\/w:t>/);
  assert.deepEqual(
    descendants(parseXml(edited, 'part'))
      .filter((n) => n.local === 't')
      .map((n) => n.text),
    [' <&😀 ', 'new'],
  );
});

test('malformed XML and DTD are refused', () => {
  assert.throws(() => parseXml('<a><b></a>', 'bad'));
  assert.throws(() => parseXml('<!DOCTYPE a [<!ENTITY x "hello">]><a>&x;</a>', 'bad'), /DTD/);
  assert.throws(() => parseXml('<a>'.repeat(300) + '</a>'.repeat(300), 'deep'), /structural limit/);
});

test('relationship targets resolve relative to source, external links are never fetched', () => {
  const pkg = DocxPackage.open(
    fixture(paragraph('a'), {
      'word/_rels/document.xml.rels': relationships(
        relation('style', 'styles', '../custom/styles.xml') +
          relation('web', 'hyperlink', 'https://example.test/', true),
      ),
      'custom/styles.xml': `<w:styles xmlns:w="${W}"/>`,
    }),
  );
  assert.equal(pkg.relationships.find((r) => r.id === 'style')!.target, 'custom/styles.xml');
  assert.equal(pkg.relationships.find((r) => r.id === 'web')!.external, true);
  assert.equal(pkg.diagnostics.length, 0);
});

test('missing relationships and content types prevent saving', () => {
  const pkg = DocxPackage.open(
    fixture(paragraph('a'), {
      'word/_rels/document.xml.rels': relationships(
        relation('missing', 'image', 'media/missing.bin'),
      ),
      'untyped.unknown': new Uint8Array([1]),
    }),
  );
  assert.ok(pkg.diagnostics.some((d) => d.code === 'missing-target'));
  assert.ok(pkg.diagnostics.some((d) => d.code === 'missing-content-type'));
  assert.throws(() => pkg.save(), /validation failed/);
});

test('unknown binary parts survive and package byte access cannot mutate the snapshot', () => {
  const pkg = DocxPackage.open(
    fixture(paragraph('a'), { 'custom/data.bin': new Uint8Array([0, 255, 17]) }),
  );
  pkg.bytes('custom/data.bin').fill(0);
  assert.deepEqual(pkg.bytes('custom/data.bin'), new Uint8Array([0, 255, 17]));
  const copy = DocxPackage.open(pkg.save());
  assert.deepEqual(copy.bytes('custom/data.bin'), new Uint8Array([0, 255, 17]));
});

test('caller-owned Node Buffers and save results cannot mutate a package', () => {
  const input = Buffer.from(fixture());
  const pkg = DocxPackage.open(input);
  const expected = pkg.save();
  input.fill(0);
  pkg.save().fill(0);
  assert.deepEqual(pkg.save(), expected);
});

test('editing xml:space never rewrites lookalike text inside extension attributes', () => {
  const source = `<w:t xmlns:w="${W}" xmlns:x="urn:future" x:note='keep xml:space="default" exactly' xml:space='default'>old</w:t>`;
  const node = parseXml(source, 'text');
  const result = patchXml(source, [textPatch(source, node, ' new ')]);
  assert.match(result, /x:note='keep xml:space="default" exactly'/);
  assert.match(result, /xml:space="preserve"> new <\/w:t>/);
  parseXml(result, 'text');
});

test('ZIP resource limits, unsafe paths, CRC damage, and truncation are refused', () => {
  assert.throws(() => readZip(fixture('', { '../escape.xml': '<a/>' })), /Unsafe/);
  const bytes = fixture();
  assert.throws(() => readZip(bytes.subarray(0, bytes.length - 10)));
  const crcDamaged = bytes.slice();
  const view = new DataView(crcDamaged.buffer);
  const directory = view.getUint32(crcDamaged.length - 6, true);
  view.setUint32(directory + 16, 0, true);
  assert.throws(() => readZip(crcDamaged), /integrity/);
  const oversized = bytes.slice();
  new DataView(oversized.buffer).setUint32(directory + 24, 200_000_000, true);
  assert.throws(() => readZip(oversized), /oversized/);
});

test('comparison reports changed, added, and removed parts', () => {
  const a = DocxPackage.open(fixture(paragraph('a'), { 'a.bin': new Uint8Array([1]) }));
  const b = DocxPackage.open(fixture(paragraph('b'), { 'b.bin': new Uint8Array([2]) }));
  const report = comparePackages(a, b);
  assert.equal(report.find((p) => p.part === 'a.bin')!.status, 'removed');
  assert.equal(report.find((p) => p.part === 'b.bin')!.status, 'added');
  assert.equal(report.find((p) => p.part === 'word/document.xml')!.status, 'changed');
});
