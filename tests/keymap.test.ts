import { test } from 'node:test';
import assert from 'node:assert/strict';
import { resolveKeymap, commandShortcuts, shortcutCommand, editorCommand } from '../src/ui/keymap';

const event = {
  key: 'k',
  ctrlKey: true,
  metaKey: false,
  altKey: false,
  shiftKey: false,
  repeat: false,
  isComposing: false,
};
test('remapping updates resolution and presentation and supports explicit unbinding', () => {
  const map = resolveKeymap({ 'file.open': ['Ctrl+k'], 'file.saveAs': [] });
  assert.equal(shortcutCommand(event, false, map), 'file.open');
  assert.equal(shortcutCommand({ ...event, key: 'o' }, false, map), undefined);
  assert.deepEqual(commandShortcuts('file.open', map), ['Ctrl+k']);
  assert.deepEqual(commandShortcuts('file.saveAs', map), []);
  assert.equal(shortcutCommand({ ...event, key: 'k', isComposing: true }, false, map), undefined);
});
test('custom bindings win over conflicting defaults; ambiguous overrides and malformed input are refused', () => {
  const map = resolveKeymap({ 'file.open': ['Shift+Ctrl+P'] });
  assert.deepEqual(commandShortcuts('commands.open', map), []);
  assert.equal(shortcutCommand({ ...event, key: 'P', shiftKey: true }, false, map), 'file.open');
  assert.throws(
    () => resolveKeymap({ 'file.open': ['Ctrl+k'], 'file.saveAs': ['Mod+K'] }),
    /Conflicting/,
  );
  assert.throws(() => resolveKeymap({ 'file.open': ['Super+K'] }), /Invalid/);
  assert.throws(() => resolveKeymap({ unknown: ['Ctrl+k'] } as never), /Unknown/);
  assert.throws(() => resolveKeymap({ 'file.open': 'Ctrl+k' } as never), /array/);
});
test('workspace history defers to fields and paragraph editing commands can be remapped', () => {
  const map = resolveKeymap({ 'text.deleteBackward': ['Alt+h'] });
  const z = { ...event, key: 'z' };
  assert.equal(shortcutCommand(z, false, map), 'history.undo');
  assert.equal(shortcutCommand(z, true, map), undefined);
  assert.equal(editorCommand(z, map), undefined);
  assert.equal(editorCommand({ ...event, key: 'Enter', ctrlKey: false }, map), 'paragraph.split');
  assert.equal(
    editorCommand({ ...event, key: 'h', ctrlKey: false, altKey: true }, map),
    'text.deleteBackward',
  );
  assert.equal(editorCommand({ ...event, key: 'Backspace', ctrlKey: false }, map), undefined);
});
