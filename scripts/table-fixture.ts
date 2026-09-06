import { readFile, writeFile } from 'node:fs/promises';
import { DocxPackage } from '../src/package/docx';
import { child, isWord, patchXml } from '../src/package/xml';

// Start from docxfix's known-good package. Insert explicit synthetic geometry;
// the fixture never uses the table reader or its interpretation of merges.
const path = 'fixtures/generated/tables.docx';
const pkg = DocxPackage.open(await readFile(path));
const xml = pkg.text(pkg.mainPart),
  body = child(pkg.xml(pkg.mainPart), 'body')!;
const paragraphs = body.children
  .filter((n) => isWord(n, 'p'))
  .map((p) => xml.slice(p.start, p.end));
const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (content: string, props = '') => `<w:tc><w:tcPr>${props}</w:tcPr>${content}</w:tc>`;
const borders = `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((edge) => `<w:${edge} w:val="single" w:sz="6" w:color="839787"/>`).join('')}</w:tblBorders>`;
const inner = `<w:tbl><w:tblPr><w:tblW w:type="pct" w:w="5000"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="1800"/></w:tblGrid><w:tr>${cell(p('Currency'))}${cell(p('CHF'))}</w:tr><w:tr>${cell(p('Due'))}${cell(p('Net 30'))}</w:tr></w:tbl>`;
const table = `<w:tbl><w:tblPr><w:tblW w:type="pct" w:w="5000"/><w:tblLayout w:type="fixed"/>${borders}<w:tblCellMar><w:top w:w="90" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="3600"/><w:gridCol w:w="2400"/></w:tblGrid>
<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell(p('RESPONSIBILITIES AND TERMS'), '<w:gridSpan w:val="3"/><w:shd w:fill="E3ECDF"/>')}</w:tr>
<w:tr>${cell(p('Delivery'), '<w:vMerge w:val="restart"/><w:vAlign w:val="center"/>')}${cell(paragraphs[2]! + paragraphs[3]!, '<w:gridSpan w:val="2"/>')}</w:tr>
<w:tr>${cell('<w:p/>', '<w:vMerge/>')}${cell(p('Acceptance criteria'))}${cell(p('Written confirmation'))}</w:tr>
<w:tr>${cell(p('Payment'))}${cell(paragraphs[4]!)}${cell(p('Commercial details') + inner + '<w:p/>')}</w:tr>
</w:tbl>`;
const borderless = `<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="6000"/><w:jc w:val="center"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((edge) => `<w:${edge} w:val="nil"/>`).join('')}</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell(p('Signed for the supplier'))}${cell(p('Signed for the customer'))}</w:tr></w:tbl>`;
const section = body.children.find((n) => isWord(n, 'sectPr'))!;
const replacement =
  paragraphs[0]! +
  paragraphs[1]! +
  table +
  paragraphs[5]! +
  borderless +
  '<w:p/>' +
  xml.slice(section.start, section.end);
const output = pkg.withXml(
  new Map([
    [pkg.mainPart, patchXml(xml, [{ start: body.openEnd, end: body.closeStart, replacement }])],
  ]),
);
await writeFile(path, output.save());
console.log('Added merged cells, nested table and borderless table to tables.docx');
