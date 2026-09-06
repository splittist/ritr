import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { DocxPackage } from '../src/package/docx';
import { child, patchXml } from '../src/package/xml';
import { readDocument } from '../src/engine/document';

// Focused interoperability probes for shared abstract definitions and overrides.
// These are generated under ignored test-results, never over user documents.
await mkdir('test-results', { recursive: true });
const base = DocxPackage.open(await readFile('fixtures/generated/lists.docx'));
const body = child(base.xml(base.mainPart), 'body')!;
const cases: Record<string, number[]> = {
  shared: [1, 7, 1, 7],
  overrides: [1, 1, 3, 3, 1, 3, 7],
  indent: [1, 1],
};
const report = [];
for (const [name, ids] of Object.entries(cases)) {
  const replacement = ids
    .map(
      (id, i) =>
        `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${id}"/></w:numPr>${name === 'indent' ? `<w:ind w:firstLine="360"${i === 1 ? ' w:hanging="0"' : ''}/>` : ''}</w:pPr><w:r><w:t>List ${id}</w:t></w:r></w:p>`,
    )
    .join('');
  const pkg = base.withXml(
    new Map([
      [
        base.mainPart,
        patchXml(base.text(base.mainPart), [
          { start: body.openEnd, end: body.closeStart, replacement },
        ]),
      ],
    ]),
  );
  const file = resolve(`test-results/numbering-probe-${name}.docx`);
  await writeFile(file, pkg.save());
  report.push({
    file,
    paragraphs: readDocument(pkg).stories[0]!.paragraphs.map((p) => ({
      label: p.numbering?.text ?? '',
      left: p.layout.left / 20,
      right: p.layout.right / 20,
      firstLine: p.layout.firstLine / 20,
    })),
  });
}
await writeFile('test-results/numbering-probes.json', JSON.stringify(report, null, 2));
