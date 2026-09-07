import { test } from 'node:test';
import assert from 'node:assert/strict';
import { deletionRange, validInlineText } from '../src/ui/inline-edit';

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
