import { test } from 'node:test';
import assert from 'node:assert/strict';
import { deletionRange, wordDeletionRange, validInlineText } from '../src/ui/inline-edit';

test('inline deletion consumes complete combining, surrogate, and ZWJ graphemes', () => {
  for (const grapheme of ['e\u0301', '😀', '👩‍💻', '🇨🇭']) {
    const text = `a${grapheme}z`;
    assert.deepEqual(deletionRange(text, 1 + grapheme.length, 1 + grapheme.length, true), {
      from: 1,
      to: 1 + grapheme.length,
    });
    assert.deepEqual(deletionRange(text, 1, 1, false), { from: 1, to: 1 + grapheme.length });
    assert.deepEqual(deletionRange(text, 2, 2, true), { from: 1, to: 1 + grapheme.length });
  }
  assert.deepEqual(deletionRange('', 0, 0, true), { from: 0, to: 0 });
  assert.deepEqual(deletionRange('abc', 0, 0, true), { from: 0, to: 0 });
  assert.deepEqual(deletionRange('abc', 3, 3, false), { from: 3, to: 3 });
});

test('inline input allows Unicode plain text and refuses structure and projection markers', () => {
  for (const text of ['', 'café 😀 e\u0301', '<a> & text']) assert.ok(validInlineText(text));
  for (const text of ['a\nb', '\t', '\r', '\0', '\ufffc', '\ud800', '\udfff'])
    assert.equal(validInlineText(text), false);
});

test('word deletion crosses formatting-independent text, skips horizontal space and respects paragraph/object edges', () => {
  assert.deepEqual(wordDeletionRange('one two  three', 9, 9, true), { from: 4, to: 9 });
  assert.deepEqual(wordDeletionRange('one  two three', 3, 3, false), { from: 3, to: 8 });
  assert.deepEqual(wordDeletionRange('one, two', 4, 4, true), { from: 3, to: 4 });
  assert.deepEqual(wordDeletionRange('one\ntwo', 4, 4, true), { from: 3, to: 4 });
  assert.deepEqual(wordDeletionRange('one\ntwo', 3, 3, false), { from: 3, to: 4 });
  assert.deepEqual(wordDeletionRange('one\n  two', 6, 6, true), { from: 4, to: 6 });
  assert.deepEqual(wordDeletionRange('one  \ntwo', 3, 3, false), { from: 3, to: 5 });
  assert.deepEqual(wordDeletionRange('one\ufffctwo', 4, 4, true), { from: 4, to: 4 });
  assert.deepEqual(wordDeletionRange('one\ufffctwo', 3, 3, false), { from: 3, to: 3 });
  assert.deepEqual(wordDeletionRange('', 0, 0, true), { from: 0, to: 0 });
  assert.deepEqual(wordDeletionRange('one', 0, 0, true), { from: 0, to: 0 });
  assert.deepEqual(wordDeletionRange('one', 3, 3, false), { from: 3, to: 3 });
});
test('word deletion retains complete Unicode graphemes and deletes a selection without expansion to words', () => {
  for (const word of ['café', 'e\u0301lan', '👩‍💻', '🇨🇭', "don't"]) {
    const text = `a ${word} z`;
    assert.deepEqual(wordDeletionRange(text, 2, 2, false), { from: 2, to: 2 + word.length });
    assert.deepEqual(wordDeletionRange(text, 2 + word.length, 2 + word.length, true), {
      from: 2,
      to: 2 + word.length,
    });
  }
  assert.deepEqual(wordDeletionRange('first second', 2, 8, true), { from: 2, to: 8 });
  assert.deepEqual(wordDeletionRange('a😀z', 2, 2, false), { from: 1, to: 3 });
});
