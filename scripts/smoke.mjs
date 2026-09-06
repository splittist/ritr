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
  assert.equal(await page.locator('.document-link').count(), 3);
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
  assert.deepEqual(errors, []);
  console.log(
    `Desktop smoke passed: editing, preview, multi-story replacement, undo/redo, Save As, code inspection, filtering. Artifacts: ${output}`,
  );
} finally {
  await app.close();
}
