import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { saveAs, exportWorkspace } from '../src/io/save';
import { DocxPackage } from '../src/package/docx';
import { fixture, paragraph, relation, relationships } from './helpers';

test('Save As verifies bytes, refuses existing destinations, and cleans staging files', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'ritr-test-'));
  try {
    const bytes = fixture();
    const pkg = DocxPackage.open(bytes);
    const target = join(directory, 'copy.docx');
    await saveAs(target, pkg);
    assert.deepEqual(new Uint8Array(await readFile(target)), bytes);
    await assert.rejects(saveAs(target, pkg), /EEXIST/);
    assert.deepEqual(await readdir(directory), ['copy.docx']);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test('batch export validates all files before publishing the new directory', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'ritr-test-'));
  try {
    const good = DocxPackage.open(fixture());
    const bad = DocxPackage.open(
      fixture(paragraph('bad'), {
        'word/_rels/document.xml.rels': relationships(relation('x', 'image', 'missing.bin')),
      }),
    );
    const target = join(directory, 'export');
    await assert.rejects(
      exportWorkspace(target, [
        { name: 'good.docx', pkg: good },
        { name: 'bad.docx', pkg: bad },
      ]),
      /validation/,
    );
    assert.deepEqual(await readdir(directory), []);
    await exportWorkspace(target, [{ name: 'good.docx', pkg: good }]);
    assert.deepEqual(await readdir(target), ['good.docx']);
    await assert.rejects(
      exportWorkspace(target, [{ name: 'other.docx', pkg: good }]),
      /already exists/,
    );
    await assert.rejects(
      exportWorkspace(join(directory, 'duplicate'), [
        { name: 'A.docx', pkg: good },
        { name: 'a.docx', pkg: good },
      ]),
      /duplicate/,
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
