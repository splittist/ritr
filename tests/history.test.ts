import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workspace } from '../src/engine/workspace';
import { projectTokens, tokenId, type InputHistory } from '../src/engine/projection';
import { fixture, paragraph } from './helpers';
function setup() {
  const w = new Workspace();
  const id = w.open('history.docx', fixture(paragraph('abcd')));
  const edit = (from: number, to: number, text: string, history?: InputHistory) => {
    const doc = w.document(id),
      story = doc.model.stories[0]!;
    w.applyProjection({
      documentId: id,
      storyId: story.id,
      origin: tokenId(story.tokens[0]!),
      expectedRevision: doc.revision,
      from,
      to,
      text,
      history,
    });
  };
  const text = () => projectTokens(w.document(id).model.stories[0]!.tokens).text;
  return { w, id, edit, text };
}
test('typing bursts undo and redo atomically with original and final caret positions', () => {
  const { w, edit, text } = setup();
  const history: InputHistory = { id: 'burst', kind: 'typing' };
  edit(0, 0, 'X', history);
  edit(1, 1, 'Y', history);
  edit(2, 2, 'Z', history);
  assert.equal(text(), 'XYZabcd');
  w.undo();
  assert.equal(text(), 'abcd');
  assert.equal(w.canUndo, false);
  assert.equal(w.selection?.head, 0);
  w.redo();
  assert.equal(text(), 'XYZabcd');
  assert.equal(w.selection?.head, 3);
  edit(3, 3, '!', history);
  w.undo();
  assert.equal(text(), 'XYZabcd', 'Redo seals the old group');
});
test('repeated deletion groups by direction and preserves the original selection', () => {
  for (const kind of ['backspace', 'delete'] as const) {
    const { w, edit, text } = setup();
    const history = { id: 'delete', kind };
    edit(kind === 'backspace' ? 3 : 0, kind === 'backspace' ? 4 : 1, '', history);
    edit(kind === 'backspace' ? 2 : 0, kind === 'backspace' ? 3 : 1, '', history);
    assert.equal(text(), kind === 'backspace' ? 'ab' : 'cd');
    w.undo();
    assert.equal(text(), 'abcd');
    assert.equal(w.canUndo, false);
  }
});
test('save, explicit boundaries, nonadjacent edits, paste and paragraph changes seal history groups', () => {
  for (const boundary of ['save', 'new-group', 'nonadjacent', 'paste', 'split'] as const) {
    const { w, id, edit, text } = setup();
    const history: InputHistory = { id: 'burst', kind: 'typing' };
    edit(0, 0, 'X', history);
    if (boundary === 'save') w.markSaved(id, w.package(id));
    if (boundary === 'paste') edit(1, 1, 'paste');
    if (boundary === 'split') edit(1, 1, '\n', history);
    const before = text();
    const from =
      boundary === 'nonadjacent' ? 0 : boundary === 'paste' ? 6 : boundary === 'split' ? 2 : 1;
    edit(from, from, 'Y', boundary === 'new-group' ? { ...history, id: 'next' } : history);
    w.undo();
    assert.equal(text(), before, boundary);
    if (boundary === 'save') assert.equal(w.document(id).dirty, false);
  }
});
