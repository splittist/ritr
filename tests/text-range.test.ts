import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workspace } from '../src/engine/workspace';
import { spans } from '../src/engine/document';
import { positionAt, replacePieces, textSegments } from '../src/engine/text-range';
import {
  editDraft,
  inputDraft,
  selectDraft,
  deletionRange,
  type InlineDraft,
} from '../src/ui/inline-edit';
import { fixture } from './helpers';

const body =
  '<w:p><w:r><w:t>ab</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>CD</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>ef</w:t></w:r></w:p>';
function setup(xml = body) {
  const workspace = new Workspace();
  const documentId = workspace.open(
    'runs.docx',
    fixture(xml, { 'custom.bin': new Uint8Array([1, 2, 3]) }),
  );
  const document = workspace.document(documentId),
    story = document.model.stories[0]!;
  const source = spans(document.model);
  return { workspace, documentId, story, source };
}
function draft(texts = ['ab', 'CD', 'ef']): InlineDraft {
  const pieces = texts.map((text, i) => ({ spanId: String(i), text }));
  const start = positionAt(pieces, 0);
  return {
    spanId: '0',
    pieces,
    original: pieces,
    text: texts.join(''),
    anchor: 0,
    head: 0,
    typing: start,
    initialAnchor: start,
    initialHead: start,
  };
}

test('cross-run replacement preserves suffix formatting, XML properties, identities, and unrelated parts through undo and reopen', () => {
  const { workspace: w, documentId, story, source } = setup();
  const original = w.package(documentId);
  const preview = w.previewRange({
    documentId,
    storyId: story.id,
    expectedRevision: 0,
    anchor: { spanId: source[1]!.id, offset: 1 },
    head: { spanId: source[0]!.id, offset: 1 },
    text: 'xy',
  });
  assert.deepEqual(preview.caret, { spanId: source[0]!.id, offset: 3, affinity: 'right' });
  w.commit(preview.id);
  assert.deepEqual(
    spans(w.document(documentId).model).map((s) => s.text),
    ['axy', 'D', 'ef'],
  );
  assert.deepEqual(
    spans(w.document(documentId).model).map((s) => s.id),
    source.map((s) => s.id),
  );
  assert.match(
    w.package(documentId).text('word/document.xml'),
    /<w:rPr><w:b\/><\/w:rPr><w:t[^>]*>D<\/w:t>/,
  );
  assert.deepEqual(
    w
      .report(documentId)
      .filter((p) => p.status !== 'unchanged')
      .map((p) => p.part),
    ['word/document.xml'],
  );
  const reopened = new Workspace();
  const id = reopened.open('copy.docx', w.package(documentId).save());
  assert.deepEqual(
    spans(reopened.document(id).model).map((s) => s.text),
    ['axy', 'D', 'ef'],
  );
  w.undo();
  assert.equal(w.package(documentId), original);
  w.redo();
  assert.deepEqual(
    spans(w.document(documentId).model).map((s) => s.text),
    ['axy', 'D', 'ef'],
  );
});

test('range edits refuse structure, protected text, invalid offsets, and stale revisions atomically', () => {
  const {
    workspace: w,
    documentId,
    story,
    source,
  } = setup(body + '<w:p><w:r><w:t>next</w:t></w:r></w:p>');
  const edit = {
    documentId,
    storyId: story.id,
    expectedRevision: 0,
    anchor: { spanId: source[0]!.id, offset: 1 },
    head: { spanId: source[3]!.id, offset: 1 },
    text: 'x',
  };
  const original = w.package(documentId);
  assert.throws(() => w.previewRange(edit), /structural boundary/);
  assert.throws(
    () => w.previewRange({ ...edit, head: { spanId: source[0]!.id, offset: 99 } }),
    /Invalid/,
  );
  assert.throws(() => w.previewRange({ ...edit, expectedRevision: 2 }), /stale/);
  assert.throws(
    () => w.previewPieces({ ...edit, pieces: source.map((s) => ({ spanId: s.id, text: '' })) }),
    /structural boundary/,
  );
  assert.equal(w.package(documentId), original);
  const protectedCase = setup('<w:p><w:ins w:id="1"><w:r><w:t>tracked</w:t></w:r></w:ins></w:p>');
  const p = protectedCase;
  assert.throws(
    () =>
      p.workspace.previewRange({
        documentId: p.documentId,
        storyId: p.story.id,
        expectedRevision: 0,
        anchor: { spanId: p.source[0]!.id, offset: 0 },
        head: { spanId: p.source[0]!.id, offset: 2 },
        text: '',
      }),
    /revision|tracked|protected/i,
  );
});

test('shared segments stop at bookmarks, tabs, hyperlinks, and paragraphs but cross formatting', () => {
  const { story } = setup(
    body +
      '<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r><w:bookmarkStart w:id="0" w:name="x"/><w:r><w:t>c</w:t></w:r><w:hyperlink><w:r><w:t>d</w:t></w:r></w:hyperlink><w:r><w:t>e</w:t></w:r></w:p>',
  );
  assert.deepEqual(
    textSegments(story).map((s) => s.text),
    ['abCDef', 'a', 'b', 'c', 'd', 'e'],
  );
});

test('draft insertion uses affinity, replacement uses first selected character, and no-op retains all formatting', () => {
  const d = draft();
  assert.deepEqual(
    editDraft(selectDraft(d, 2, 2), 2, 2, 'x').pieces.map((p) => p.text),
    ['abx', 'CD', 'ef'],
  );
  const right = { ...d, typing: { spanId: '1', offset: 0, affinity: 'right' as const } };
  assert.deepEqual(
    editDraft(right, 2, 2, 'x').pieces.map((p) => p.text),
    ['ab', 'xCD', 'ef'],
  );
  assert.deepEqual(
    editDraft(d, 2, 4, 'xy').pieces.map((p) => p.text),
    ['ab', 'xy', 'ef'],
  );
  assert.deepEqual(editDraft(d, 1, 5, 'bCDe').pieces, d.pieces);
});

test('repeated draft edits preserve unrelated formatting and typing format after deletion, including empty segments', () => {
  let d = editDraft(draft(), 1, 5, '');
  assert.deepEqual(
    d.pieces.map((p) => p.text),
    ['a', '', 'f'],
  );
  d = editDraft(d, 1, 1, 'x');
  assert.deepEqual(
    d.pieces.map((p) => p.text),
    ['ax', '', 'f'],
  );
  d = editDraft(selectDraft(d, 3, 3), 3, 3, '!');
  assert.deepEqual(
    d.pieces.map((p) => p.text),
    ['ax', '', 'f!'],
  );
  d = editDraft(draft(), 2, 6, '');
  d = editDraft(d, 2, 2, 'x');
  assert.deepEqual(
    d.pieces.map((p) => p.text),
    ['ab', 'x', ''],
  );
  d = editDraft(draft(), 0, 6, '');
  d = editDraft(d, 0, 0, 'x');
  assert.deepEqual(
    d.pieces.map((p) => p.text),
    ['x', '', ''],
  );
});

test('browser input uses actual selection in repeated text and consumes graphemes across runs', () => {
  const d = draft(['aa', 'aa']);
  const changed = inputDraft(d, 'aaaaa', 3, 3, { from: 2, to: 2 });
  assert.deepEqual(
    changed.pieces.map((p) => p.text),
    ['aaa', 'aa'],
  );
  const emoji = draft(['a👩', '‍💻z']);
  const range = deletionRange(emoji.text, 1, 1, false);
  const deleted = editDraft(emoji, range.from, range.to, '');
  assert.deepEqual(
    deleted.pieces.map((p) => p.text),
    ['a', 'z'],
  );
  assert.throws(() => replacePieces(emoji.pieces, 2, 3, ''), /surrogate/);
  assert.equal(inputDraft(draft(['😀']), '😁', 2, 2).text, '😁');
});

test('editable neighbors of tracked formatting remain editable; drafts cannot cross protected text', () => {
  const {
    workspace: w,
    documentId,
    story,
    source,
  } = setup(
    '<w:p><w:r><w:t>safe</w:t></w:r><w:r><w:rPr><w:rPrChange w:id="1"/></w:rPr><w:t>protected</w:t></w:r><w:r><w:t>also safe</w:t></w:r></w:p>',
  );
  const edit = {
    documentId,
    storyId: story.id,
    expectedRevision: 0,
    pieces: [{ spanId: source[0]!.id, text: 'updated' }],
  };
  assert.throws(
    () =>
      w.previewPieces({ ...edit, pieces: source.map((s) => ({ spanId: s.id, text: 'changed' })) }),
    /structural boundary/,
  );
  w.commit(w.previewPieces(edit).id);
  assert.deepEqual(
    spans(w.document(documentId).model).map((s) => s.text),
    ['updated', 'protected', 'also safe'],
  );
});

test('empty-run typing ownership survives engine apply and subsequent range insertion', () => {
  const { workspace: w, documentId, story, source } = setup();
  const edit = {
    documentId,
    storyId: story.id,
    expectedRevision: 0,
    anchor: { spanId: source[1]!.id, offset: 0 },
    head: { spanId: source[2]!.id, offset: 2 },
    text: '',
  };
  const deletion = w.previewRange(edit);
  w.commit(deletion.id);
  const insertion = w.previewRange({
    ...edit,
    expectedRevision: 1,
    anchor: deletion.caret,
    head: deletion.caret,
    text: 'X',
  });
  w.commit(insertion.id);
  assert.deepEqual(
    spans(w.document(documentId).model).map((s) => s.text),
    ['ab', 'X', ''],
  );
});
