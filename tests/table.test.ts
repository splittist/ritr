import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { DocxPackage, comparePackages } from '../src/package/docx';
import { readDocument } from '../src/engine/document';
import { Workspace } from '../src/engine/workspace';
import { fixture } from './helpers';

const p = (text = '') => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text = '', props = '') => `<w:tc><w:tcPr>${props}</w:tcPr>${p(text)}</w:tc>`;
const row = (cells: string, props = '') => `<w:tr><w:trPr>${props}</w:trPr>${cells}</w:tr>`;
const table = (rows: string) => `<w:tbl>${rows}</w:tbl>`;
function model(body: string) {
  return readDocument(DocxPackage.open(fixture(body))).stories[0]!;
}
function first(body: string) {
  const block = model(body).blocks.find((b) => b.kind === 'table');
  assert.ok(block?.kind === 'table');
  return block.table;
}

test('block ranges partition the token stream and nest table cell paragraphs', () => {
  const story = model(p('Before') + table(row(cell('Inside'))) + p('After'));
  let end = 0;
  for (const block of story.blocks) {
    assert.equal(block.range.from, end);
    end = block.range.to;
  }
  assert.equal(end, story.tokens.length);
  const block = story.blocks.find((b) => b.kind === 'table')!;
  assert.ok(block.kind === 'table');
  assert.ok(block.table.rows[0]!.cells[0]!.blocks.some((b) => b.kind === 'paragraph'));
  assert.equal(block.table.linear, false);
});

test('grid holes, horizontal spans and vertical continuations retain their coordinates', () => {
  const props = '<w:gridSpan w:val="2"/>';
  const t = first(
    table(
      row(cell('A', props + '<w:vMerge w:val="restart"/>'), '<w:gridBefore w:val="1"/>') +
        row(cell('', props + '<w:vMerge/>'), '<w:gridBefore w:val="1"/>'),
    ),
  );
  assert.equal(t.linear, false);
  assert.equal(t.columns.length, 3);
  const a = t.rows[0]!.cells[0]!,
    b = t.rows[1]!.cells[0]!;
  assert.deepEqual([a.column, a.colSpan, a.rowSpan], [1, 2, 2]);
  assert.equal(b.mergedInto, a.id);
  assert.equal(a.inspector.source.nodeId, a.source.nodeId);
});

test('unsafe geometry falls back without concealing continuation content', () => {
  for (const rows of [
    row(cell('Orphan', '<w:vMerge/>')),
    row(cell('A', '<w:vMerge w:val="restart"/>')) + row(cell('Keep me', '<w:vMerge/>')),
    row(cell('A', '<w:gridSpan w:val="129"/>')),
    row(cell('A', '<w:hMerge/>')),
    row(cell('A'), '<w:gridBefore w:val="-1"/>'),
  ]) {
    const t = first(table(rows));
    assert.equal(t.linear, true);
    assert.ok(t.warnings.length);
  }
});

test('tracked cell property changes protect text', () => {
  const story = model(
    table(row(cell('Protected', '<w:tcPrChange w:id="1"><w:tcPr/></w:tcPrChange>'))),
  );
  const token = story.tokens.find((t) => t.kind === 'text');
  assert.ok(token?.kind === 'text');
  assert.equal(token.span.editable, false);
});

test('table fixture: nested blocks, direct formatting, numbering and lossless cell editing', async () => {
  const bytes = new Uint8Array(await readFile('fixtures/generated/tables.docx'));
  const ws = new Workspace();
  const id = ws.open('tables.docx', bytes);
  const original = ws.package(id);
  const story = ws.document(id).model.stories[0]!;
  const tables = story.blocks.filter((b) => b.kind === 'table');
  assert.equal(tables.length, 2);
  const t = tables[0]!.table;
  assert.deepEqual(t.columns, [1800, 3600, 2400]);
  assert.equal(t.rows[0]!.cells[0]!.colSpan, 3);
  assert.equal(t.rows[1]!.cells[0]!.rowSpan, 2);
  assert.equal(t.rows[0]!.cells[0]!.shading, '#E3ECDF');
  assert.equal(t.rows[3]!.cells[2]!.blocks.filter((b) => b.kind === 'table').length, 1);
  assert.ok(story.paragraphs.filter((p) => p.numbering).length >= 3);
  assert.deepEqual(original.save(), bytes);
  const match = ws.search('Written confirmation')[0]!;
  ws.commit(ws.preview('Edit cell', [{ ...match, text: 'Signed confirmation' }]).id);
  assert.equal(
    ws.package(id).text('word/document.xml'),
    original
      .text('word/document.xml')
      .replace(
        '<w:t>Written confirmation</w:t>',
        '<w:t xml:space="preserve">Signed confirmation</w:t>',
      ),
  );
  assert.deepEqual(
    comparePackages(original, ws.package(id))
      .filter((p) => p.status !== 'unchanged')
      .map((p) => p.part),
    ['word/document.xml'],
  );
  assert.equal(
    readDocument(DocxPackage.open(ws.package(id).save())).stories[0]!.blocks.filter(
      (b) => b.kind === 'table',
    ).length,
    2,
  );
  ws.undo();
  assert.deepEqual(ws.package(id).save(), bytes);
  ws.redo();
  assert.equal(ws.search('Signed confirmation').length, 1);
});
