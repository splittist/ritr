import { _electron as electron } from 'playwright';
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { resolve, join } from 'node:path';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
await mkdir('test-results', { recursive: true });
const output = await mkdtemp(resolve('test-results/desktop-'));
const crossRunPath = join(output, 'cross-run.docx');
const crossRunParts = unzipSync(await readFile('fixtures/generated/plain.docx'));
const crossRunXml = strFromU8(crossRunParts['word/document.xml']);
assert.ok(crossRunXml.includes('A small document with a clear beginning.'));
crossRunParts['word/document.xml'] = strToU8(
  crossRunXml
    .replace(
      'A small document with a clear beginning.',
      'Cross-run </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>needle',
    )
    .replace('Replace this phrase', 'Unrelated text'),
);
await writeFile(crossRunPath, zipSync(crossRunParts));
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
    crossRunPath,
  ],
  env,
});
try {
  console.log('Electron connected');
  const page = await app.firstWindow();
  page.on('dialog', (dialog) => void dialog.accept().catch(() => {}));
  page.setDefaultTimeout(15000);
  console.log('Window loaded:', await page.title());
  const errors = [];
  page.on('pageerror', (error) => errors.push(String(error)));
  await page.getByRole('heading', { name: 'plain.docx', exact: true }).waitFor();
  assert.equal(await page.locator('.document-link').count(), 6);
  assert.ok((await page.locator('.code-token').count()) > 0);
  // Registry availability, modal focus, and keyboard navigation use the same actions as buttons.
  const searchCommands = page.getByRole('combobox', { name: 'Search commands' });
  const palette = page.getByRole('dialog', { name: 'Commands', exact: true });
  await page.getByLabel('Find text', { exact: true }).focus();
  await page.keyboard.press('Control+Shift+P');
  await palette.waitFor();
  await searchCommands.fill('Undo');
  const undoOption = palette.locator('[data-command="history.undo"]');
  assert.equal(await undoOption.getAttribute('aria-disabled'), 'true');
  assert.ok((await undoOption.textContent()).includes('No workspace transaction to undo.'));
  await searchCommands.press('Enter');
  assert.equal(
    await palette.count(),
    1,
    'Unavailable commands do not execute or close the palette',
  );
  await searchCommands.fill('missing command xyz');
  await palette.getByText(/No commands match/).waitFor();
  await searchCommands.press('Escape');
  assert.equal(
    await page
      .getByLabel('Find text', { exact: true })
      .evaluate((input) => input === document.activeElement),
    true,
  );
  await page.getByRole('button', { name: 'Commands', exact: true }).click();
  await searchCommands.fill('codes');
  await searchCommands.press('Enter');
  assert.equal(await page.getByLabel('Reveal codes', { exact: true }).isChecked(), false);
  await page.keyboard.press('Control+Shift+E');
  assert.equal(await page.getByLabel('Reveal codes', { exact: true }).isChecked(), true);
  await page.locator('.editable-span').first().click();
  await page.getByLabel('Text content').fill('A revised opening with café and 😀.');
  await page.getByRole('button', { name: 'Preview text edit' }).click();
  await page.getByLabel('Change preview').waitFor();
  await page.getByLabel('Text content').fill('Another inspector draft');
  assert.equal(await page.getByLabel('Change preview').count(), 0);
  await page.getByLabel('Text content').fill('A revised opening with café and 😀.');
  await page.getByLabel('Text content').press('Control+Enter');
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
  await page.getByRole('button', { name: /cross-run.docx/ }).click();
  await page.getByLabel('Find text', { exact: true }).fill('Cross-run needle');
  await page.getByRole('button', { name: 'Find', exact: true }).click();
  await page.locator('.matches').getByText('1 matches · 0 protected').waitFor();
  await page.locator('.matches button').click();
  assert.equal(await page.getByLabel('Text content').inputValue(), 'Cross-run ');
  await page.getByLabel('Replace with', { exact: true }).fill('Joined text');
  await page.getByRole('button', { name: 'Preview replacement' }).click();
  await page.getByLabel('Change preview').waitFor();
  assert.equal(await page.locator('.diff').count(), 2);
  await page.getByRole('button', { name: 'Apply transaction' }).click();
  await page.locator('.editable-span').filter({ hasText: 'Joined text' }).waitFor();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.locator('.editable-span').filter({ hasText: 'needle' }).waitFor();
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await page.locator('.editable-span').filter({ hasText: 'Joined text' }).waitFor();
  // Inline drafts stay outside the package until preview/commit, including empty runs.
  await page.locator('.editable-span').filter({ hasText: 'Joined text' }).dblclick();
  const inlineInput = page.getByRole('textbox', { name: 'Inline text', exact: true });
  await inlineInput.fill('Draft café 👩‍💻');
  await page.screenshot({ path: join(output, 'inline-draft.png'), fullPage: true });
  await inlineInput.press('End');
  await inlineInput.press('Backspace');
  assert.equal(await inlineInput.inputValue(), 'Draft café ');
  await inlineInput.press('Backspace');
  await inlineInput.pressSequentially('!');
  await page.getByRole('button', { name: 'Preview inline edit' }).click();
  await page.getByLabel('Change preview').waitFor();
  assert.equal(await page.locator('.diff ins').textContent(), 'Draft café!');
  await inlineInput.fill('Draft revised');
  assert.equal(
    await page.getByLabel('Change preview').count(),
    0,
    'Typing expires the old preview',
  );
  await inlineInput.evaluate((input) => {
    input.setSelectionRange(2, 5);
    input.dispatchEvent(new Event('select', { bubbles: true }));
  });
  await page.getByLabel('Reveal codes', { exact: true }).uncheck();
  assert.equal(await inlineInput.inputValue(), 'Draft revised');
  await page.waitForFunction(() => {
    const input = document.querySelector('.inline-text');
    return input?.selectionStart === 2 && input?.selectionEnd === 5;
  });
  assert.deepEqual(
    await inlineInput.evaluate((input) => [input.selectionStart, input.selectionEnd]),
    [2, 5],
  );
  await page.getByRole('button', { name: 'Cancel inline edit' }).click();
  await page.locator('.editable-span').filter({ hasText: 'Joined text' }).waitFor();
  // A saved cross-run replacement left the second source span empty.
  const emptySpan = page
    .locator('.editable-span')
    .filter({ hasText: /^\u200b$/ })
    .first();
  await emptySpan.dblclick();
  assert.equal(await inlineInput.inputValue(), '');
  await inlineInput.pressSequentially('Inline ');
  await inlineInput.evaluate((input) => {
    input.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    input.value = 'Inline café';
    input.dispatchEvent(
      new InputEvent('input', {
        bubbles: true,
        inputType: 'insertCompositionText',
        isComposing: true,
      }),
    );
    input.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'é' }));
  });
  await inlineInput.press('End');
  await inlineInput.evaluate((input) => {
    const clipboardData = new DataTransfer();
    clipboardData.setData('text/plain', 'bad\nline');
    const event = new ClipboardEvent('paste', { clipboardData, bubbles: true, cancelable: true });
    input.dispatchEvent(event);
    if (!event.defaultPrevented) throw new Error('Multiline paste was not refused');
  });
  await inlineInput.press('Enter');
  assert.equal(await inlineInput.inputValue(), 'Inline café');
  await page.getByRole('button', { name: 'Preview inline edit' }).click();
  await page.getByRole('button', { name: 'Apply transaction' }).click();
  await page.locator('.editable-span').filter({ hasText: 'Inline café' }).waitFor();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await emptySpan.waitFor();
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await page.locator('.editable-span').filter({ hasText: 'Inline café' }).waitFor();
  const inlinePath = join(output, 'inline-edited.docx');
  await app.evaluate(({ dialog }, path) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: path });
  }, inlinePath);
  await page.getByRole('button', { name: 'Save As…' }).click();
  await page.waitForFunction(() =>
    document.querySelector('footer')?.textContent.includes('Saved and verified'),
  );
  const savedInline = unzipSync(await readFile(inlinePath));
  assert.ok(strFromU8(savedInline['word/document.xml']).includes('Inline café'));
  for (const [part, bytes] of Object.entries(crossRunParts)) {
    if (part !== 'word/document.xml') assert.deepEqual(savedInline[part], bytes, part);
  }
  await page.getByRole('button', { name: /tables.docx/ }).click();
  await page.locator('.editable-span').filter({ hasText: 'Signed confirmation' }).dblclick();
  await inlineInput.fill('Inline cell');
  await page.getByRole('button', { name: 'Preview inline edit' }).click();
  await page.getByRole('button', { name: 'Apply transaction' }).click();
  await page.locator('.document-table .editable-span').filter({ hasText: 'Inline cell' }).waitFor();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.locator('.editable-span').filter({ hasText: 'Signed confirmation' }).waitFor();
  await page.getByLabel('Reveal codes', { exact: true }).check();
  await page.locator('.editable-span').filter({ hasText: 'Signed confirmation' }).dblclick();
  await inlineInput.fill('Stale draft');
  await page.evaluate(async () => {
    const snapshot = await window.ritr.snapshot();
    const doc = snapshot.documents.find((d) => d.name === 'tables.docx');
    const token = doc.model.stories
      .flatMap((s) => s.tokens)
      .find((t) => t.kind === 'text' && t.span.text === 'Signed confirmation');
    const preview = await window.ritr.previewEdit({
      documentId: doc.id,
      spanId: token.span.id,
      from: 0,
      to: token.span.text.length,
      text: 'External edit',
    });
    await window.ritr.commit(preview.id);
  });
  await page.getByRole('button', { name: 'Preview inline edit' }).click();
  await page.waitForFunction(() =>
    document.querySelector('footer')?.textContent.includes('Inline draft is stale'),
  );
  assert.equal(await page.getByLabel('Change preview').count(), 0);
  await page.getByRole('button', { name: 'Cancel inline edit' }).click();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.screenshot({ path: join(output, 'inline-editing.png'), fullPage: true });
  await page.getByRole('button', { name: /cross-run.docx/ }).click();
  await page.getByLabel('Reveal codes', { exact: true }).uncheck();
  const joinedSpan = page.locator('.editable-span').filter({ hasText: 'Joined text' });
  await joinedSpan.click();
  await page.keyboard.press('x');
  await inlineInput.waitFor();
  assert.ok((await inlineInput.inputValue()).includes('x'));
  await inlineInput.press('Escape');
  await joinedSpan.waitFor();
  await joinedSpan.click();
  await page.keyboard.press('Delete');
  await inlineInput.waitFor();
  assert.equal((await inlineInput.inputValue()).length, 'Joined text'.length - 1);
  await inlineInput.press('Escape');
  await joinedSpan.click();
  await joinedSpan.evaluate((span) => {
    const clipboardData = new DataTransfer();
    clipboardData.setData('text/plain', 'Pasted café');
    span.dispatchEvent(
      new ClipboardEvent('paste', { clipboardData, bubbles: true, cancelable: true }),
    );
  });
  await inlineInput.waitFor();
  assert.ok((await inlineInput.inputValue()).includes('Pasted café'));
  await inlineInput.press('Escape');
  // A selection across adjacent source runs remains protected even with codes hidden.
  const editorContent = page.locator('.cm-content').first();
  await editorContent.press('Control+Home');
  await editorContent.press('Control+Shift+End');
  await editorContent.press('x');
  assert.equal(await inlineInput.count(), 0);
  await page.getByRole('button', { name: /review.docx/ }).click();
  await page.locator('.protected-span').first().dblclick();
  assert.equal(await inlineInput.count(), 0);
  await page.keyboard.press('Control+Shift+P');
  await searchCommands.fill('selected span inline');
  assert.equal(await palette.getByRole('option').getAttribute('aria-disabled'), 'true');
  await page.screenshot({ path: join(output, 'commands-protected.png'), fullPage: true });
  await searchCommands.press('Escape');
  await page.getByRole('button', { name: /cross-run.docx/ }).click();
  await joinedSpan.click();
  await page.keyboard.press('Control+Shift+P');
  await searchCommands.fill('selected span inline');
  await searchCommands.press('Enter');
  await inlineInput.waitFor();
  await inlineInput.fill('Command draft');
  await inlineInput.evaluate((input) => input.setSelectionRange(2, 5));
  await page.keyboard.press('Control+Shift+P');
  await searchCommands.fill('Save As');
  assert.ok(
    (await palette.getByRole('option').textContent()).includes(
      'Apply or cancel the inline draft first.',
    ),
  );
  await searchCommands.press('Escape');
  assert.equal(await inlineInput.inputValue(), 'Command draft');
  assert.deepEqual(
    await inlineInput.evaluate((input) => [
      input === document.activeElement,
      input.selectionStart,
      input.selectionEnd,
    ]),
    [true, 2, 5],
  );
  // Native text undo must never undo a committed workspace transaction.
  const beforeNativeUndo = await page.evaluate(() => window.ritr.snapshot());
  await inlineInput.press('End');
  await inlineInput.pressSequentially('!');
  await inlineInput.press('Control+z');
  assert.deepEqual(await page.evaluate(() => window.ritr.snapshot()), beforeNativeUndo);
  await inlineInput.fill('Command draft');
  await inlineInput.press('Control+Enter');
  await page.getByLabel('Change preview').waitFor();
  await page.keyboard.press('Control+Shift+Enter');
  await page.locator('.editable-span').filter({ hasText: 'Command draft' }).waitFor();
  await page.keyboard.press('Control+z');
  await joinedSpan.waitFor();
  await page.keyboard.press('Control+Shift+z');
  await page.locator('.editable-span').filter({ hasText: 'Command draft' }).waitFor();
  await page.keyboard.press('Control+f');
  assert.equal(
    await page
      .getByLabel('Find text', { exact: true })
      .evaluate((input) => input === document.activeElement),
    true,
  );
  await page.keyboard.press('Control+Shift+P');
  await searchCommands.fill('compare package');
  await searchCommands.press('Enter');
  await page.locator('.part-report').waitFor();
  await page.keyboard.press('Control+Shift+P');
  await searchCommands.press('ArrowDown');
  assert.equal(await searchCommands.getAttribute('aria-activedescendant'), 'command-file.saveAs');
  await page.screenshot({ path: join(output, 'commands.png'), fullPage: true });
  await searchCommands.press('Escape');
  assert.deepEqual(errors, []);
  console.log(
    `Desktop smoke passed: command palette, shortcuts, availability and focus, inspector and inline editing, Unicode drafts, stale/boundary refusal, preview, cross-run replacement, undo/redo, Save As, code inspection, lists and merged/nested tables. Artifacts: ${output}`,
  );
} catch (error) {
  console.error(error);
  throw error;
} finally {
  await app.close();
}
