import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
await mkdir('test-results', { recursive: true });
const output = await mkdtemp(resolve('test-results/desktop-'));
const env = { ...process.env, RITR_SMOKE: '1' };
delete env.ELECTRON_RUN_AS_NODE;
const app = await electron.launch({
  executablePath: require('electron'),
  args: [
    '.',
    'fixtures/generated/plain.docx',
    'fixtures/generated/sections.docx',
    'fixtures/generated/review.docx',
    'fixtures/generated/lists.docx',
    'fixtures/generated/tables.docx',
  ],
  env,
});
try {
  console.log('Electron connected');
  const page = await app.firstWindow();
  page.setDefaultTimeout(15000);
  console.log('Window loaded:', await page.title());
  const errors = [];
  page.on('pageerror', (error) => errors.push(String(error)));
  await page.getByRole('heading', { name: 'plain.docx', exact: true }).waitFor();
  assert.equal(await page.locator('.document-link').count(), 5);
  assert.ok((await page.locator('.code-token').count()) > 0);
  await page.locator('.editable-span').first().click();
  await page.getByLabel('Text content').fill('A revised opening with café and 😀.');
  await page.getByRole('button', { name: 'Preview text edit' }).click();
  await page.getByLabel('Change preview').waitFor();
  await page.getByRole('button', { name: 'Apply transaction' }).click();
  await page.waitForFunction(() =>
    document.querySelector('.cm-content')?.textContent.includes('A revised opening'),
  );
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.waitForFunction(() =>
    document.querySelector('.cm-content')?.textContent.includes('A small document'),
  );
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await page.waitForFunction(() =>
    document.querySelector('.cm-content')?.textContent.includes('A revised opening'),
  );
  await page.getByLabel('Find text', { exact: true }).fill('Replace this phrase');
  await page.getByLabel('Replace with', { exact: true }).fill('Updated wording');
  await page.getByRole('button', { name: 'Preview replacement' }).click();
  await page.getByLabel('Change preview').waitFor();
  assert.equal(await page.locator('.diff').count(), 4);
  await page.getByRole('button', { name: 'Apply transaction' }).click();
  await page.waitForFunction(() =>
    document.querySelector('.cm-content')?.textContent.includes('Updated wording'),
  );
  await page.getByRole('button', { name: 'Compare package parts' }).click();
  await page.locator('.part-report').waitFor();
  assert.equal(await page.locator('.part-report .changed').count(), 1);
  const savedPath = join(output, 'edited.docx');
  await app.evaluate(({ dialog }, path) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: path });
  }, savedPath);
  await page.getByRole('button', { name: 'Save As…' }).click();
  await page.waitForFunction(() =>
    document.querySelector('footer')?.textContent.includes('Saved and verified'),
  );
  assert.ok((await readFile(savedPath)).length > 0);
  const snapshot = await page.evaluate(() => window.ritr.snapshot());
  assert.equal(snapshot.documents[0].dirty, false);
  assert.equal(snapshot.documents[1].dirty, true);
  await page.locator('.code-token').first().click();
  await page.getByRole('heading', { name: 'Source XML', exact: true }).waitFor();
  await page.screenshot({ path: join(output, 'workbench.png'), fullPage: true });
  await page.getByLabel('Reveal codes', { exact: true }).uncheck();
  assert.equal(await page.locator('.code-token:not(.opaque)').count(), 0);
  await page.getByRole('button', { name: /review.docx/ }).click();
  assert.ok((await page.locator('.protected-span').count()) > 0);
  await page.getByRole('button', { name: /lists.docx/ }).click();
  await page.getByRole('heading', { name: 'lists.docx', exact: true }).waitFor();
  const expectedLabels = [
    '1.',
    '1.1.',
    '1.2.',
    '2.',
    '2.1.',
    '•',
    '◦',
    '•',
    '5.',
    '6.',
    '1.',
    'AA)',
    'BB)',
    'IV.',
  ];
  // CodeMirror virtualizes offscreen lines. Check the initial visible prefix,
  // then scroll to verify the final labels rather than counting detached DOM.
  const visibleLabels = await page.locator('.list-marker').allTextContents();
  assert.deepEqual(visibleLabels, expectedLabels.slice(0, visibleLabels.length));
  assert.ok(visibleLabels.length > 5);
  await page.locator('main').evaluate((el) => {
    el.scrollTop = el.scrollHeight;
  });
  await page.getByRole('button', { name: 'List label IV.', exact: true }).waitFor();
  await page.getByRole('button', { name: 'List label BB)', exact: true }).waitFor();
  await page.locator('main').evaluate((el) => {
    el.scrollTop = 0;
  });
  await page.getByRole('button', { name: 'List label 1.1.', exact: true }).click();
  await page.getByRole('heading', { name: 'Generated list label', exact: true }).waitFor();
  await page.screenshot({ path: join(output, 'numbering-clean.png'), fullPage: true });
  const wrapping = await page
    .locator('.editable-span')
    .nth(2)
    .evaluate((element) => {
      const range = document.createRange();
      range.selectNodeContents(element);
      return [...range.getClientRects()].map((r) => ({ left: r.left, top: r.top, width: r.width }));
    });
  console.log('Wrapped list text geometry:', JSON.stringify(wrapping));
  assert.ok(wrapping.length > 1, 'Long numbered paragraph wraps');
  assert.ok(
    Math.abs(wrapping[0].left - wrapping[1].left) < 1,
    'Wrapped text aligns under the first line text',
  );
  await page.getByLabel('Reveal codes', { exact: true }).check();
  assert.ok((await page.locator('.list-marker').count()) > 0);
  assert.ok((await page.locator('.code-token').count()) > 0);
  await page.screenshot({ path: join(output, 'numbering-codes.png'), fullPage: true });
  const codeLabels = await page.locator('.list-marker').allTextContents();
  assert.deepEqual(codeLabels, expectedLabels.slice(0, codeLabels.length));
  await page.getByRole('button', { name: /tables.docx/ }).click();
  await page.getByRole('heading', { name: 'tables.docx', exact: true }).waitFor();
  await page.getByLabel('Reveal codes', { exact: true }).uncheck();
  assert.equal(await page.locator('.document-table').count(), 3);
  assert.equal(await page.locator('td[colspan="3"]').count(), 1);
  assert.equal(await page.locator('td[rowspan="2"]').count(), 1);
  const cellWrapping = await page
    .locator('.editable-span')
    .filter({ hasText: 'Deliver the materials' })
    .evaluate((element) => {
      const range = document.createRange();
      range.selectNodeContents(element);
      return [...range.getClientRects()].map((r) => ({ left: r.left, top: r.top }));
    });
  assert.ok(cellWrapping.length > 1, 'Numbered text wraps within the merged cell');
  assert.ok(
    Math.abs(cellWrapping[0].left - cellWrapping[1].left) < 1,
    'Cell continuation retains hanging indentation',
  );
  await page.locator('.editable-span').filter({ hasText: 'Written confirmation' }).click();
  await page.getByLabel('Text content').fill('Signed confirmation');
  await page.getByRole('button', { name: 'Preview text edit' }).click();
  await page.getByRole('button', { name: 'Apply transaction' }).click();
  await page.locator('.editable-span').filter({ hasText: 'Signed confirmation' }).waitFor();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.locator('.editable-span').filter({ hasText: 'Written confirmation' }).waitFor();
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await page.locator('.editable-span').filter({ hasText: 'Signed confirmation' }).waitFor();
  await page.screenshot({ path: join(output, 'tables-clean.png'), fullPage: true });
  await page.getByLabel('Reveal codes', { exact: true }).check();
  await page.getByRole('button', { name: 'Cell 2, 1', exact: true }).first().click();
  await page.getByRole('heading', { name: 'Cell properties', exact: true }).waitFor();
  await page.screenshot({ path: join(output, 'tables-codes.png'), fullPage: true });
  const tablePath = join(output, 'tables-edited.docx');
  await app.evaluate(({ dialog }, path) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: path });
  }, tablePath);
  await page.getByRole('button', { name: 'Save As…' }).click();
  await page.waitForFunction(() =>
    document.querySelector('footer')?.textContent.includes('Saved and verified'),
  );
  assert.ok((await readFile(tablePath)).length > 0);
  assert.deepEqual(errors, []);
  console.log(
    `Desktop smoke passed: editing, preview, multi-story replacement, undo/redo, Save As, code inspection, lists and merged/nested tables. Artifacts: ${output}`,
  );
} finally {
  await app.close();
}
