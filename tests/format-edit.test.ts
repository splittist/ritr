import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workspace } from '../src/engine/workspace';
import { spans } from '../src/engine/document';
import { projectTokens, tokenId } from '../src/engine/projection';
import { descendants } from '../src/package/xml';
import { effectiveFormat, type TextFormat } from '../src/engine/format';
import { fixture, paragraph } from './helpers';
function setup(
  body = '<w:p><w:r><w:rPr><w:i/><w:lang w:val="en-GB"/><x:custom/></w:rPr><w:t>abcdef</w:t></w:r></w:p>',
) {
  const w = new Workspace(),
    id = w.open('format.docx', fixture(body, { 'custom.bin': new Uint8Array([1, 2, 3]) }));
  const request = () => {
    const d = w.document(id),
      s = d.model.stories[0]!;
    return {
      documentId: id,
      storyId: s.id,
      origin: tokenId(s.tokens[0]!),
      expectedRevision: d.revision,
    };
  };
  const format = (from: number, to: number, format: TextFormat) =>
    w.applyFormat({ ...request(), from, to, format });
  return { w, id, format, request };
}
test('partial formatting isolates only selected text and preserves source identities and unrelated properties', () => {
  const { w, id, format } = setup();
  const before = w.package(id),
    original = spans(w.document(id).model)[0]!.id;
  format(2, 4, { b: true, color: '1234AB', highlight: 'yellow' });
  const result = spans(w.document(id).model);
  assert.deepEqual(
    result.map((s) => s.text),
    ['ab', 'cd', 'ef'],
  );
  assert.equal(result[1]!.id, original);
  assert.equal(effectiveFormat(result[0]!).b, false);
  assert.deepEqual(effectiveFormat(result[1]!), {
    b: true,
    i: true,
    u: false,
    color: '1234AB',
    highlight: 'yellow',
  });
  const root = w.package(id).xml(before.mainPart);
  assert.equal(descendants(root).filter((n) => n.local === 'r').length, 3);
  assert.equal(descendants(root).filter((n) => n.local === 'custom').length, 3);
  for (const r of descendants(root).filter((n) => n.local === 'r'))
    assert.ok(
      !descendants(r)
        .slice(1)
        .some((n) => n.local === 'r'),
    );
  assert.deepEqual(
    w
      .report(id)
      .filter((p) => p.status !== 'unchanged')
      .map((p) => p.part),
    ['word/document.xml'],
  );
  w.undo();
  assert.equal(w.package(id), before);
  assert.equal(w.selection?.anchor, 2);
  assert.equal(w.selection?.head, 4);
  w.redo();
  assert.deepEqual(
    spans(w.document(id).model).map((s) => s.text),
    ['ab', 'cd', 'ef'],
  );
  const reopened = new Workspace();
  const copy = reopened.open('copy', w.package(id).save());
  assert.deepEqual(
    spans(reopened.document(copy).model).map((s) => effectiveFormat(s)),
    result.map((s) => effectiveFormat(s)),
  );
});
test('formatting whole runs and multiple paragraphs retains existing run boundaries, markers, and off values', () => {
  const { w, id, format } = setup(
    '<w:p><w:bookmarkStart w:id="1" w:name="m"/><w:r><w:rPr><w:b/></w:rPr><w:t>ab</w:t><w:t>cd</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
      paragraph('ef'),
  );
  const before = w.package(id),
    nodes = descendants(before.xml(before.mainPart));
  const ids = nodes
    .filter((n) => ['r', 't', 'bookmarkStart', 'bookmarkEnd'].includes(n.local))
    .map((n) => n.id);
  format(0, 7, { b: false, u: true });
  assert.equal(projectTokens(w.document(id).model.stories[0]!.tokens).text, 'abcd\nef');
  for (const s of spans(w.document(id).model)) {
    assert.equal(effectiveFormat(s).b, false);
    assert.equal(effectiveFormat(s).u, true);
  }
  for (const nodeId of ids)
    assert.ok(descendants(w.package(id).xml(before.mainPart)).some((n) => n.id === nodeId));
  assert.equal(
    descendants(w.package(id).xml(before.mainPart)).filter((n) => n.local === 'r').length,
    2,
  );
});
test('formatting refuses protected text, selected objects, invalid colors, stale revisions and surrogate boundaries atomically', () => {
  for (const body of [
    '<w:p><w:ins w:id="1"><w:r><w:t>abc</w:t></w:r></w:ins></w:p>',
    '<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>',
  ]) {
    const { w, id, format } = setup(body),
      before = w.package(id);
    assert.throws(() => format(0, 3, { b: true }));
    assert.equal(w.package(id), before);
  }
  const { w, id, format, request } = setup(paragraph('a😀b')),
    before = w.package(id),
    stale = request();
  assert.throws(() => format(0, 2, { b: true }));
  assert.throws(() => format(0, 1, { color: 'not-a-color' }));
  assert.throws(() => format(0, 1, { highlight: 'orange' }));
  assert.equal(w.package(id), before);
  format(0, 1, { b: true });
  assert.throws(() => w.applyFormat({ ...stale, from: 0, to: 1, format: { i: true } }), /stale/);
});
test('typing formatting is atomic with input, groups naturally, and does not make a new run per character', () => {
  const { w, id, request } = setup(paragraph('tail'));
  for (const [from, text] of ['X', 'Y', 'Z'].entries())
    w.applyProjection({
      ...request(),
      from,
      to: from,
      text,
      typingFormat: { b: true },
      history: { id: 'typing', kind: 'typing' },
    });
  assert.deepEqual(
    spans(w.document(id).model).map((s) => s.text),
    ['XYZ', 'tail'],
  );
  assert.equal(effectiveFormat(spans(w.document(id).model)[0]!).b, true);
  w.undo();
  assert.deepEqual(
    spans(w.document(id).model).map((s) => s.text),
    ['tail'],
  );
  w.redo();
  assert.equal(projectTokens(w.document(id).model.stories[0]!.tokens).text, 'XYZtail');
});
