import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { DocxPackage } from '../src/package/docx';
import { readDocument } from '../src/engine/document';

await mkdir('test-results', { recursive: true });
const report = [];
for (const name of ['legal', 'headings', 'lists']) {
  const file = resolve(`fixtures/generated/${name}.docx`);
  const model = readDocument(DocxPackage.open(await readFile(file)));
  report.push({
    file,
    paragraphs: model.stories[0]!.paragraphs.map((p) => ({
      label: p.numbering?.text ?? '',
      left: p.layout.left / 20,
      right: p.layout.right / 20,
      firstLine: p.layout.firstLine / 20,
    })),
  });
}
await writeFile('test-results/numbering-expected.json', JSON.stringify(report, null, 2));
console.log('Expected Word labels and indentation written to test-results/numbering-expected.json');
