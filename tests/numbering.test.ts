import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { DocxPackage, comparePackages } from '../src/package/docx';
import { W } from '../src/package/xml';
import { readDocument, plainText } from '../src/engine/document';
import { Workspace } from '../src/engine/workspace';
import { formatCounter } from '../src/engine/numbering';
import { paragraphGeometry, paragraphLineStyle } from '../src/ui/paragraph-layout';
import { fixture, relation, relationships } from './helpers';

const p = (num = '1', level = 0, extra = '') =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${num}"/></w:numPr>${extra}</w:pPr><w:r><w:t>Text</w:t></w:r></w:p>`;
const lvl = (index: number, extra = '', format = 'decimal', template = `%${index + 1}.`) =>
  `<w:lvl w:ilvl="${index}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${template}"/>${extra}</w:lvl>`;
const num = (id = '1', overrides = '', abstract = '0') =>
  `<w:num w:numId="${id}"><w:abstractNumId w:val="${abstract}"/>${overrides}</w:num>`;
function packageFor(
  body: string,
  levels = lvl(0),
  instances = num(),
  styles = '',
  extraAbstract = '',
) {
  return DocxPackage.open(
    fixture(body, {
      'word/_rels/document.xml.rels': relationships(
        relation('n', 'numbering', 'numbering.xml') + relation('s', 'styles', 'styles.xml'),
      ),
      'word/numbering.xml': `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0">${levels}</w:abstractNum>${extraAbstract}${instances}</w:numbering>`,
      'word/styles.xml': `<w:styles xmlns:w="${W}">${styles}</w:styles>`,
    }),
  );
}
const labels = (pkg: DocxPackage) =>
  readDocument(pkg).stories[0]!.paragraphs.map((p) => p.numbering?.text ?? '');

test('lists continue through interruptions and numIds sharing an abstract definition', () => {
  const pkg = packageFor(
    p() + '<w:p><w:r><w:t>Interruption</w:t></w:r></w:p>' + p('2') + p() + p('2'),
    lvl(0),
    num() + num('2'),
  );
  assert.deepEqual(labels(pkg), ['1.', '', '2.', '3.', '4.']);
  assert.match(plainText(readDocument(pkg).stories[0]!), /1\.\tText/);
});
test('multilevel placeholders and restart after higher levels match Word rules', () => {
  const pkg = packageFor(
    p('1', 0) + p('1', 1) + p('1', 2) + p('1', 2) + p('1', 0) + p('1', 2),
    lvl(0) + lvl(1, '', 'lowerLetter', '%1(%2)') + lvl(2, '', 'lowerRoman', '%1.%2.%3'),
  );
  assert.deepEqual(labels(pkg), ['1.', '1(a)', '1.a.i', '1.a.ii', '2.', '2.a.i']);
});
test('never-restart and custom restart thresholds', () => {
  const body =
    p('1', 0) + p('1', 1) + p('1', 2) + p('1', 1) + p('1', 2) + p('1', 0) + p('1', 1) + p('1', 2);
  const levels = lvl(0) + lvl(1, '<w:lvlRestart w:val="0"/>') + lvl(2, '<w:lvlRestart w:val="1"/>');
  assert.deepEqual(labels(packageFor(body, levels)), [
    '1.',
    '1.',
    '1.',
    '2.',
    '2.',
    '2.',
    '3.',
    '1.',
  ]);
});
test('startOverride is used initially and after restart; full level overrides change labels', () => {
  const overrides =
    '<w:lvlOverride w:ilvl="1"><w:startOverride w:val="5"/>' +
    lvl(1, '<w:lvlRestart w:val="0"/>', 'upperRoman', '(%2)') +
    '</w:lvlOverride>';
  assert.deepEqual(
    labels(
      packageFor(
        p('1', 0) + p('1', 1) + p('1', 1) + p('1', 0) + p('1', 1),
        lvl(0) + lvl(1),
        num('1', overrides),
      ),
    ),
    ['1.', '(V)', '(VI)', '2.', '(V)'],
  );
});
test('legal numbering forces decimal ancestor values', () => {
  assert.deepEqual(
    labels(
      packageFor(
        p('1', 0) + p('1', 1),
        lvl(0, '', 'upperRoman') + lvl(1, '<w:isLgl/>', 'lowerLetter', '%1.%2'),
      ),
    ),
    ['I.', '1.1'],
  );
});
test('formats handle Word alphabetic sequences, Roman bounds and padded decimals', () => {
  assert.equal(formatCounter(28, 'upperLetter'), 'BB');
  assert.equal(formatCounter(53, 'lowerLetter'), 'aaa');
  assert.equal(formatCounter(1999, 'upperRoman'), 'MCMXCIX');
  assert.equal(formatCounter(4, 'lowerRoman'), 'iv');
  assert.equal(formatCounter(4, 'decimalZero'), '04');
  assert.equal(formatCounter(0, 'upperRoman'), undefined);
  assert.equal(formatCounter(4000, 'upperRoman'), undefined);
});
test('Unicode bullets and common legacy font glyphs have portable display labels', () => {
  assert.deepEqual(labels(packageFor(p() + p(), lvl(0, '', 'bullet', '•'))), ['•', '•']);
  assert.deepEqual(
    labels(
      packageFor(p(), lvl(0, '<w:rPr><w:rFonts w:ascii="Symbol"/></w:rPr>', 'bullet', '\uf0b7')),
    ),
    ['•'],
  );
  assert.deepEqual(
    labels(
      packageFor(p(), lvl(0, '<w:rPr><w:rFonts w:ascii="Wingdings"/></w:rPr>', 'bullet', '\uf0a7')),
    ),
    ['▪'],
  );
});
test('unsupported formats, glyphs, missing levels and bad references produce diagnostics', () => {
  for (const pkg of [
    packageFor(p(), lvl(0, '', 'chicago')),
    packageFor(p(), lvl(0, '', 'bullet', '\uf001')),
    packageFor(p('9')),
    packageFor(p('1', 8)),
    packageFor(p(), lvl(0, '', 'decimal', '%2')),
  ]) {
    assert.deepEqual(labels(pkg), ['?']);
    assert.ok(readDocument(pkg).diagnostics.some((d) => d.code === 'paragraph-display'));
  }
});
test('style membership ignores style ilvl and follows the level pStyle binding', () => {
  const styles =
    '<w:style w:type="paragraph" w:styleId="Heading"><w:pPr><w:numPr><w:ilvl w:val="8"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Derived"><w:basedOn w:val="Heading"/></w:style>';
  const body = '<w:p><w:pPr><w:pStyle w:val="Derived"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>';
  const model = readDocument(
    packageFor(
      body,
      lvl(0) + lvl(1, '<w:pStyle w:val="Heading"/>', 'decimal', '%1.%2'),
      num(),
      styles,
    ),
  );
  assert.equal(model.stories[0]!.paragraphs[0]!.numbering!.text, '1.1');
  assert.match(model.stories[0]!.paragraphs[0]!.propertySources.level!, /Heading/);
});
test('numId=0 suppresses inherited numbering', () => {
  const styles =
    '<w:style w:type="paragraph" w:styleId="List"><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>';
  assert.deepEqual(
    labels(packageFor(p('0', 0, '<w:pStyle w:val="List"/>'), lvl(0), num(), styles)),
    [''],
  );
});
test('numStyleLink resolves through numbering styles and cycles are diagnosed', () => {
  const styles =
    '<w:style w:type="numbering" w:styleId="Linked"><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>';
  const linked =
    '<w:abstractNum w:abstractNumId="2"><w:numStyleLink w:val="Linked"/></w:abstractNum>';
  assert.deepEqual(labels(packageFor(p('2'), lvl(0), num() + num('2', '', '2'), styles, linked)), [
    '1.',
  ]);
  const cycleStyles = styles.replace('w:val="1"', 'w:val="2"');
  assert.deepEqual(
    labels(packageFor(p('2'), lvl(0), num() + num('2', '', '2'), cycleStyles, linked)),
    ['?'],
  );
});
test('indentation merges defaults, list, style chain and direct attributes with provenance', () => {
  const styles =
    '<w:docDefaults><w:pPrDefault><w:pPr><w:ind w:right="144"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:ind w:left="1080"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Derived"><w:basedOn w:val="Base"/><w:pPr><w:ind w:right="360"/></w:pPr></w:style>';
  const para = readDocument(
    packageFor(
      p('1', 0, '<w:pStyle w:val="Derived"/><w:ind w:right="720"/>'),
      lvl(0, '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'),
      num(),
      styles,
    ),
  ).stories[0]!.paragraphs[0]!;
  assert.deepEqual([para.layout.left, para.layout.right, para.layout.firstLine], [1080, 720, -360]);
  assert.equal(para.layout.sources.left, 'Style Base');
  assert.equal(para.layout.sources.right, 'Direct paragraph');
  assert.equal(para.layout.sources.hanging, 'List 1, level 1');
});
test('source indentation maps to bounded display geometry without changing source measurements', () => {
  const layout = { left: 720, right: 360, firstLine: -360, sources: {} };
  assert.deepEqual(paragraphGeometry(layout), {
    left: 48,
    right: 24,
    firstLine: -24,
    markerWidth: 24,
  });
  assert.match(paragraphLineStyle(layout, true), /text-indent:0px/);
  assert.equal(paragraphGeometry({ ...layout, left: 20000 }).left, 240);
  assert.equal(layout.left, 720);
});

test('direct first-line indentation replaces inherited hanging, but same-element hanging wins', () => {
  const levels = lvl(0, '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>');
  const body =
    p('1', 0, '<w:ind w:firstLine="360"/>') + p('1', 0, '<w:ind w:firstLine="360" w:hanging="0"/>');
  assert.deepEqual(
    readDocument(packageFor(body, levels)).stories[0]!.paragraphs.map((p) => p.layout.firstLine),
    [360, 0],
  );
});

test('distinct abstract definitions are independent; start overrides reset a shared stream once', () => {
  const independent = packageFor(
    p() + p('2') + p(),
    lvl(0),
    num() + num('2', '', '2'),
    '',
    `<w:abstractNum w:abstractNumId="2">${lvl(0)}</w:abstractNum>`,
  );
  assert.deepEqual(labels(independent), ['1.', '1.', '2.']);
  const override = '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride>';
  assert.deepEqual(
    labels(
      packageFor(p() + p() + p('2') + p('2') + p() + p('2'), lvl(0), num() + num('2', override)),
    ),
    ['1.', '2.', '5.', '6.', '7.', '8.'],
  );
});
test('opaque numbered paragraphs make subsequent list labels explicitly uncertain', () => {
  const body = p() + `<w:sdt><w:sdtContent>${p()}</w:sdtContent></w:sdt>` + p();
  assert.deepEqual(labels(packageFor(body)), ['1.', '?']);
});
test('numbering in table cells continues with ordinary body paragraphs', () => {
  assert.deepEqual(
    labels(packageFor(p() + `<w:tbl><w:tr><w:tc>${p()}</w:tc></w:tr></w:tbl>` + p())),
    ['1.', '2.', '3.'],
  );
});

const expected: Record<string, string[]> = {
  legal: ['1.', '1.1.', '1.2.', '', '2.', '2.1.', '2.1.1.'],
  headings: ['1', '', '2', '2.1', '(a)', '(b)', '(i)', '2.2', '(a)'],
  lists: [
    '',
    '1.',
    '1.1.',
    '',
    '1.2.',
    '2.',
    '2.1.',
    '•',
    '◦',
    '•',
    '5.',
    '6.',
    '1.',
    '',
    '',
    'AA)',
    'BB)',
    'IV.',
  ],
};
for (const name of ['legal', 'headings', 'lists'])
  test(`docxfix ${name}: labels, preservation and undo/save/reopen`, async () => {
    const bytes = new Uint8Array(await readFile(`fixtures/generated/${name}.docx`));
    const pkg = DocxPackage.open(bytes);
    assert.deepEqual(labels(pkg), expected[name]);
    assert.deepEqual(pkg.save(), bytes);
    const ws = new Workspace();
    const id = ws.open(`${name}.docx`, bytes);
    const span = ws.document(id).model.stories[0]!.tokens.find((t) => t.kind === 'text')!;
    if (span.kind !== 'text') throw new Error('Missing text');
    ws.commit(
      ws.preview('Edit list text', [
        { documentId: id, spanId: span.span.id, from: 0, to: 0, text: 'Updated ' },
      ]).id,
    );
    assert.deepEqual(labels(ws.package(id)), expected[name]);
    assert.deepEqual(
      comparePackages(pkg, ws.package(id))
        .filter((p) => p.status !== 'unchanged')
        .map((p) => p.part),
      ['word/document.xml'],
    );
    assert.deepEqual(labels(DocxPackage.open(ws.package(id).save())), expected[name]);
    ws.undo();
    assert.deepEqual(ws.package(id).save(), bytes);
    ws.redo();
    assert.deepEqual(labels(ws.package(id)), expected[name]);
  });
