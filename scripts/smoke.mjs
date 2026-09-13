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
    .replace(
      /(<w:p[^>]*>)/,
      '$1<w:pPr><w:pStyle w:val="SmokeHeading"/></w:pPr><w:bookmarkStart w:id="50" w:name="split_join_smoke"/><w:proofErr w:type="spellStart"/>',
    )
    .replace('</w:p>', '<w:proofErr w:type="spellEnd"/><w:bookmarkEnd w:id="50"/></w:p>')
    .replace('Replace this phrase', 'Unrelated text')
    .replace(
      '<w:sectPr',
      '<w:p><w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/><w:color w:val="1234AB"/><w:highlight w:val="yellow"/><w:rFonts w:ascii="Arial"/><w:sz w:val="72"/></w:rPr><w:t>Formatted preview</w:t></w:r><w:r><w:rPr><w:b w:val="0"/><w:i w:val="0"/></w:rPr><w:t> Plain preview</w:t></w:r></w:p><w:sectPr',
    ),
);
crossRunParts['word/styles.xml'] = strToU8(
  (crossRunParts['word/styles.xml']
    ? strFromU8(crossRunParts['word/styles.xml'])
    : '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>'
  ).replace(
    '</w:styles>',
    '<w:style w:type="paragraph" w:styleId="SmokeHeading"><w:next w:val="SmokeBody"/></w:style><w:style w:type="paragraph" w:styleId="SmokeBody"/></w:styles>',
  ),
);
const styleRelationships = crossRunParts['word/_rels/document.xml.rels']
  ? strFromU8(crossRunParts['word/_rels/document.xml.rels'])
  : '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
if (!styleRelationships.includes('/styles"'))
  crossRunParts['word/_rels/document.xml.rels'] = strToU8(
    styleRelationships.replace(
      '</Relationships>',
      '<Relationship Id="smokeStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    ),
  );
const contentTypes = strFromU8(crossRunParts['[Content_Types].xml']);
if (!contentTypes.includes('/word/styles.xml'))
  crossRunParts['[Content_Types].xml'] = strToU8(
    contentTypes.replace(
      '</Types>',
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    ),
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
  await page.locator('.keymap-settings summary').click();
  const keySettings = page.getByRole('textbox', { name: 'Keybindings JSON' });
  await keySettings.fill(
    JSON.stringify({
      'commands.open': ['Ctrl+k'],
      'view.codes': [],
      'text.deleteBackward': ['Alt+h'],
    }),
  );
  await page.getByRole('button', { name: 'Apply keybindings' }).click();
  assert.ok(
    (
      await page.getByRole('button', { name: 'Commands', exact: true }).getAttribute('title')
    ).includes('Ctrl+k'),
  );
  await page.keyboard.press('Control+k');
  await page.getByRole('dialog', { name: 'Commands', exact: true }).waitFor();
  await page.getByRole('combobox', { name: 'Search commands' }).press('Escape');
  await page.keyboard.press('Control+Shift+P');
  assert.equal(await page.getByRole('dialog', { name: 'Commands', exact: true }).count(), 0);
  await page.reload();
  await page.getByRole('heading', { name: 'plain.docx', exact: true }).waitFor();
  await page.keyboard.press('Control+k');
  await page.getByRole('dialog', { name: 'Commands', exact: true }).waitFor();
  await page.getByRole('combobox', { name: 'Search commands' }).press('Escape');
  const initialText = 'A small document with a clear beginning.';
  await page.locator('.editable-span').first().click();
  await page.keyboard.press('End');
  await page.keyboard.press('Backspace');
  assert.equal(await page.locator('.editable-span').first().textContent(), initialText);
  await page.keyboard.press('Alt+h');
  await page
    .locator('.editable-span')
    .filter({ hasText: /^A small document with a clear beginning$/ })
    .waitFor();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page
    .locator('.editable-span')
    .filter({ hasText: /^A small document with a clear beginning\.$/ })
    .waitFor();
  await page.locator('.keymap-settings summary').click();
  await keySettings.fill('{}');
  await page.getByRole('button', { name: 'Apply keybindings' }).click();
  await page.screenshot({ path: join(output, 'keybindings.png'), fullPage: true });
  await page.locator('.keymap-settings summary').click();
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
  for (const codes of [true, false]) {
    await page.getByLabel('Reveal codes', { exact: true }).setChecked(codes);
    const formatting = await page
      .locator('.editable-span')
      .filter({ hasText: /^Formatted preview$/ })
      .evaluate((span) => {
        const style = getComputedStyle(span);
        const plain = getComputedStyle(
          [...document.querySelectorAll('.editable-span')].find(
            (s) => s.textContent === ' Plain preview',
          ),
        );
        return {
          weight: style.fontWeight,
          italic: style.fontStyle,
          underline: style.textDecorationLine,
          color: style.color,
          highlight: style.backgroundColor,
          sameFont: style.fontFamily === plain.fontFamily && style.fontSize === plain.fontSize,
        };
      });
    assert.deepEqual(formatting, {
      weight: '700',
      italic: 'italic',
      underline: 'underline',
      color: 'rgb(18, 52, 171)',
      highlight: 'rgb(255, 255, 0)',
      sameFont: true,
    });
  }
  await page.getByLabel('Reveal codes', { exact: true }).check();
  const expectedCodes = await page.evaluate(async () => {
    const doc = (await window.ritr.snapshot()).documents.find((d) => d.name === 'cross-run.docx');
    return doc.model.stories[0].tokens.filter((t) => t.kind === 'code').map((t) => t.label);
  });
  assert.deepEqual(
    await page.locator('.code-token').allTextContents(),
    expectedCodes,
    'Code widgets must follow source order, including adjacent Run boundaries',
  );
  await page.screenshot({ path: join(output, 'text-formatting.png'), fullPage: true });
  // Native text input replaces selections without a draft field or apply step.
  for (const codes of [false, true]) {
    await page.getByLabel('Reveal codes', { exact: true }).setChecked(codes);
    const first = page.locator('.editable-span').filter({ hasText: /^Cross-run $/ });
    await first.click();
    await first.evaluate((span) => {
      const next = [...document.querySelectorAll('.editable-span')].find(
        (s) => s.textContent === 'needle',
      );
      const range = document.createRange();
      range.setStart(span.firstChild, 7);
      range.setEnd(next.firstChild, 3);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      document.dispatchEvent(new Event('selectionchange'));
    });
    await page.waitForTimeout(100);
    await page.keyboard.press('X');
    await page
      .locator('.editable-span')
      .filter({ hasText: /^Cross-rX$/ })
      .waitFor();
    await page.locator('.editable-span').filter({ hasText: /^dle$/ }).waitFor();
    assert.equal(await page.locator('.inline-text').count(), 0);
    assert.equal(await page.getByLabel('Change preview').count(), 0);
    await page.keyboard.press('Enter');
    await page.waitForFunction(() => document.querySelectorAll('.paragraph-line').length === 5);
    // The caret is at the start of the new paragraph. Rapid input is queued losslessly.
    await page.keyboard.type('fast typing');
    await page.waitForFunction(async () => {
      const snapshot = await window.ritr.snapshot();
      const doc = snapshot.documents.find((d) => d.name === 'cross-run.docx');
      return doc.model.stories[0].tokens
        .filter((t) => t.kind === 'text')
        .map((t) => t.span.text)
        .join('')
        .includes('fast typingdle');
    });
    await page
      .locator('.editable-span')
      .filter({ hasText: /fast typing/ })
      .waitFor();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await page.waitForFunction(async () => {
      const doc = (await window.ritr.snapshot()).documents.find((d) => d.name === 'cross-run.docx');
      return !doc.model.stories[0].tokens.some(
        (t) => t.kind === 'text' && t.span.text.includes('fast typing'),
      );
    });
    assert.equal(await page.locator('.paragraph-line').count(), 5, 'Undo typing keeps the split');
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await page
      .locator('.editable-span')
      .filter({ hasText: /fast typing/ })
      .waitFor();
    await page.keyboard.press('Home');
    await page.keyboard.press('Backspace');
    await page.waitForFunction(() => document.querySelectorAll('.paragraph-line').length === 4);
    const saved = join(output, `direct-${codes}.docx`);
    await app.evaluate(({ dialog }, path) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: path });
    }, saved);
    await page.getByRole('button', { name: 'Save As…' }).click();
    await page.waitForFunction(() =>
      document.querySelector('footer')?.textContent.includes('Saved and verified'),
    );
    const parts = unzipSync(await readFile(saved));
    for (const [name, bytes] of Object.entries(crossRunParts))
      if (name !== 'word/document.xml') assert.deepEqual(parts[name], bytes);
    // Undo the join, queued typing transactions, split, and replacement.
    let attempts = 0;
    while (
      await page.evaluate(async () => {
        const doc = (await window.ritr.snapshot()).documents.find(
          (d) => d.name === 'cross-run.docx',
        );
        return !doc.model.stories[0].tokens.some(
          (t) => t.kind === 'text' && t.span.text === 'needle',
        );
      })
    ) {
      assert.ok(attempts++ < 30);
      const revision = await page.evaluate(
        async () =>
          (await window.ritr.snapshot()).documents.find((d) => d.name === 'cross-run.docx')
            .revision,
      );
      await page.getByRole('button', { name: 'Undo', exact: true }).click();
      await page.waitForFunction(
        async (previous) =>
          (await window.ritr.snapshot()).documents.find((d) => d.name === 'cross-run.docx')
            .revision !== previous,
        revision,
      );
    }
    await page
      .locator('.editable-span')
      .filter({ hasText: /^needle$/ })
      .waitFor();
  }
  // End-of-document input must address the final source paragraph, not a phantom line.
  for (const codes of [false, true]) {
    await page.getByLabel('Reveal codes', { exact: true }).setChecked(codes);
    for (const backward of [true, false]) {
      await page.locator('.cm-content').first().press('Control+End');
      await page.keyboard.press('Enter');
      await page.waitForFunction(async () => {
        const doc = (await window.ritr.snapshot()).documents.find(
          (d) => d.name === 'cross-run.docx',
        );
        return doc.model.stories[0].paragraphs.length === 5;
      });
      await page.waitForFunction(() => document.querySelectorAll('.paragraph-line').length === 5);
      if (!backward) {
        await page
          .locator('.editable-span')
          .filter({ hasText: /^ Plain preview$/ })
          .evaluate((span) => {
            const range = document.createRange();
            range.selectNodeContents(span);
            range.collapse(false);
            const selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);
            document.dispatchEvent(new Event('selectionchange'));
          });
        await page.waitForTimeout(100);
      }
      await page.keyboard.press(backward ? 'Backspace' : 'Delete');
      await page.waitForFunction(async () => {
        const doc = (await window.ritr.snapshot()).documents.find(
          (d) => d.name === 'cross-run.docx',
        );
        return doc.model.stories[0].paragraphs.length === 4;
      });
      await page.waitForFunction(() => document.querySelectorAll('.paragraph-line').length === 4);
      assert.ok(
        !(await page.locator('footer').textContent()).includes('Select text inside a paragraph'),
      );
    }
  }
  // Paste may contain paragraph breaks; Unicode deletion remains grapheme-aware.
  await page.getByLabel('Reveal codes', { exact: true }).uncheck();
  const nativeEditor = page.locator('.cm-content').first();
  await nativeEditor.press('Control+Home');
  await page.keyboard.insertText('café 👩‍💻');
  await page
    .locator('.editable-span')
    .filter({ hasText: /^café 👩‍💻Cross-run $/ })
    .waitFor();
  await page.keyboard.press('Backspace');
  await page
    .locator('.editable-span')
    .filter({ hasText: /^café Cross-run $/ })
    .waitFor();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page
    .locator('.editable-span')
    .filter({ hasText: /^café 👩‍💻Cross-run $/ })
    .waitFor();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page
    .locator('.editable-span')
    .filter({ hasText: /^Cross-run $/ })
    .waitFor();
  await nativeEditor.press('Control+Home');
  await nativeEditor.evaluate((editor) =>
    editor.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })),
  );
  await page.keyboard.insertText('é');
  await page.waitForTimeout(50);
  assert.equal(await page.getByRole('button', { name: 'Save As…' }).isDisabled(), true);
  assert.equal(
    await page.evaluate(
      async () =>
        (await window.ritr.snapshot()).documents
          .find((d) => d.name === 'cross-run.docx')
          .model.stories[0].tokens.find((t) => t.kind === 'text').span.text,
    ),
    'Cross-run ',
  );
  await nativeEditor.evaluate((editor) =>
    editor.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'é' })),
  );
  await page
    .locator('.editable-span')
    .filter({ hasText: /^éCross-run $/ })
    .waitFor();
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page
    .locator('.editable-span')
    .filter({ hasText: /^Cross-run $/ })
    .waitFor();
  // A dedicated Enter at the end of a heading uses its declared next paragraph style.
  const headingEnd = page.locator('.editable-span').filter({ hasText: /^needle$/ });
  await headingEnd.click();
  await headingEnd.evaluate((span) => {
    const range = document.createRange();
    range.selectNodeContents(span);
    range.collapse(false);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    document.dispatchEvent(new Event('selectionchange'));
  });
  await page.waitForTimeout(100);
  await page.keyboard.press('Enter');
  await page.waitForFunction(async () => {
    const doc = (await window.ritr.snapshot()).documents.find((d) => d.name === 'cross-run.docx');
    return doc.model.stories[0].paragraphs[1]?.styleId === 'SmokeBody';
  });
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await page.waitForFunction(async () => {
    const doc = (await window.ritr.snapshot()).documents.find((d) => d.name === 'cross-run.docx');
    return doc.model.stories[0].paragraphs.length === 4;
  });
  // Selection formatting and caret typing use the same source-preserving commands.
  const formatSpan = page.locator('.editable-span').filter({ hasText: /^Cross-run $/ });
  await formatSpan.click();
  await formatSpan.evaluate((span) => {
    const range = document.createRange();
    range.setStart(span.firstChild, 0);
    range.setEnd(span.firstChild, 5);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
    document.dispatchEvent(new Event('selectionchange'));
  });
  await page.waitForTimeout(100);
  await page.getByRole('button', { name: 'Bold', exact: true }).click();
  const formattedSelection = page.locator('.editable-span').filter({ hasText: /^Cross$/ });
  await formattedSelection.waitFor();
  assert.equal(await formattedSelection.evaluate((s) => getComputedStyle(s).fontWeight), '700');
  await page.keyboard.press('Control+i');
  await page.waitForFunction(() =>
    [...document.querySelectorAll('.editable-span')].some(
      (s) => s.textContent === 'Cross' && getComputedStyle(s).fontStyle === 'italic',
    ),
  );
  await page.getByLabel('Text color', { exact: true }).fill('#123456');
  await page.waitForFunction(() =>
    [...document.querySelectorAll('.editable-span')].some(
      (s) => s.textContent === 'Cross' && getComputedStyle(s).color === 'rgb(18, 52, 86)',
    ),
  );
  await page.getByLabel('Text highlight', { exact: true }).selectOption('yellow');
  await page.waitForFunction(() =>
    [...document.querySelectorAll('.editable-span')].some(
      (s) =>
        s.textContent === 'Cross' && getComputedStyle(s).backgroundColor === 'rgb(255, 255, 0)',
    ),
  );
  await page.screenshot({ path: join(output, 'formatting-controls.png'), fullPage: true });
  for (let i = 0; i < 4; i++) {
    const revision = await page.evaluate(
      async () =>
        (await window.ritr.snapshot()).documents.find((d) => d.name === 'cross-run.docx').revision,
    );
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await page.waitForFunction(
      async (revision) =>
        (await window.ritr.snapshot()).documents.find((d) => d.name === 'cross-run.docx')
          .revision !== revision,
      revision,
    );
  }
  await formatSpan.waitFor();
  await page.locator('.cm-content').first().press('Control+Home');
  await page.keyboard.press('Control+b');
  await page.keyboard.type('format me');
  const typedFormat = page.locator('.editable-span').filter({ hasText: /^format me$/ });
  await typedFormat.waitFor();
  assert.equal(await typedFormat.evaluate((s) => getComputedStyle(s).fontWeight), '700');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await typedFormat.waitFor({ state: 'detached' });
  await formatSpan.waitFor();
  assert.equal(
    await typedFormat.count(),
    0,
    'One undo removes the complete formatted typing burst',
  );
  await page.screenshot({ path: join(output, 'direct-editing.png'), fullPage: true });
  // Edits inside ordinary table cells use the same native editor.
  await page.getByRole('button', { name: /tables.docx/ }).click();
  const cell = page.locator('.editable-span').filter({ hasText: 'Signed confirmation' });
  const tableParagraphs = await page.evaluate(
    async () =>
      (await window.ritr.snapshot()).documents.find((d) => d.name === 'tables.docx').model
        .stories[0].paragraphs.length,
  );
  await cell.click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Enter');
  await page.waitForFunction(async (count) => {
    const snapshot = await window.ritr.snapshot();
    const doc = snapshot.documents.find((d) => d.name === 'tables.docx');
    return doc.model.stories[0].paragraphs.length === count + 1;
  }, tableParagraphs);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await cell.waitFor();
  await page.getByRole('button', { name: /review.docx/ }).click();
  const protectedText = page.locator('.protected-span').first();
  const before = await protectedText.textContent();
  await protectedText.click();
  await page.keyboard.press('x');
  await page.waitForFunction(
    () =>
      document.querySelector('footer')?.textContent.includes('preserved') ||
      document.querySelector('footer')?.textContent.includes('protected'),
  );
  assert.equal(await protectedText.textContent(), before);
  assert.deepEqual(errors, []);
  console.log(
    `Desktop smoke passed: command palette, persisted remapping, availability and focus, native inline editing, queued typing, paragraph split/join, formatting, protected-boundary refusal, inspector preview, undo/redo, Save As, code inspection, lists and merged/nested tables. Artifacts: ${output}`,
  );
} catch (error) {
  console.error(error);
  throw error;
} finally {
  await app.close();
}
