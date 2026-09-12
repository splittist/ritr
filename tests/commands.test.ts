import { shortcutCommand, defaultKeymap } from '../src/ui/keymap';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  commands,
  commandReason,
  filterCommands,
  invokeCommand,
  type CommandContext,
} from '../src/ui/commands';

const ready: CommandContext = {
  connected: true,
  busy: false,
  hasDocument: true,
  canUndo: true,
  canRedo: false,
  hasInlineDraft: false,
  hasQuery: true,
  hasPreview: true,
  hasChanges: true,
};

test('commands enforce document, draft, preview, and protected-selection boundaries consistently', () => {
  assert.match(commandReason('file.saveAs', { ...ready, hasInlineDraft: true })!, /cancel/);
  assert.match(commandReason('search.replace', { ...ready, hasQuery: false })!, /Find text/);
  assert.match(commandReason('transaction.apply', { ...ready, hasPreview: false })!, /preview/);
  assert.match(
    commandReason('transaction.apply', { ...ready, hasChanges: false })!,
    /no editable changes/,
  );
  assert.equal(
    commandReason('edit.inline', { ...ready, selectionReason: 'Tracked text is protected.' }),
    'Tracked text is protected.',
  );
  assert.equal(
    commandReason('edit.preview', { ...ready, editReason: 'The draft has no changes.' }),
    'The draft has no changes.',
  );
  for (const command of commands) {
    if (command.id === 'commands.open') continue;
    assert.match(commandReason(command.id, { ...ready, busy: true })!, /finish/);
    assert.match(commandReason(command.id, { ...ready, connected: false })!, /desktop/);
  }
  assert.equal(
    commandReason('commands.open', { ...ready, busy: true, connected: false }),
    undefined,
  );
  assert.equal(commandReason('edit.cancelInline', { ...ready, hasInlineDraft: true }), undefined);
});

test('invocation rechecks current availability and never calls an unavailable operation', async () => {
  let context = ready;
  let called = 0;
  const execute = () =>
    invokeCommand(
      'transaction.apply',
      () => context,
      () => {
        called++;
      },
    );
  assert.equal(commandReason('transaction.apply', context), undefined);
  context = { ...ready, hasPreview: false };
  assert.match((await execute())!, /preview/);
  assert.equal(called, 0);
  context = ready;
  assert.equal(await execute(), undefined);
  assert.equal(called, 1);
});

test('shortcuts distinguish preview from commit and leave native text undo and composition alone', () => {
  const key = {
    key: 'Enter',
    ctrlKey: true,
    metaKey: false,
    altKey: false,
    shiftKey: false,
    repeat: false,
    isComposing: false,
  };
  assert.equal(shortcutCommand(key, true), 'edit.preview');
  assert.equal(shortcutCommand({ ...key, shiftKey: true }, true), 'transaction.apply');
  assert.equal(shortcutCommand({ ...key, key: 'z' }, false), 'history.undo');
  assert.equal(shortcutCommand({ ...key, key: 'z' }, true), undefined);
  assert.equal(shortcutCommand({ ...key, key: 'z', shiftKey: true }, false), 'history.redo');
  assert.equal(shortcutCommand({ ...key, key: 'p', shiftKey: true }, true), 'commands.open');
  assert.equal(shortcutCommand({ ...key, ctrlKey: false, metaKey: true }, true), 'edit.preview');
  for (const overrides of [
    { isComposing: true },
    { repeat: true },
    { altKey: true },
    { ctrlKey: false },
  ])
    assert.equal(shortcutCommand({ ...key, ...overrides }, false), undefined);
});

test('palette search matches multiple terms and command IDs and shortcut combinations stay unique', () => {
  assert.deepEqual(
    filterCommands(' PACKAGE  compare ').map((command) => command.id),
    ['document.compare'],
  );
  assert.equal(filterCommands('not-a-command').length, 0);
  assert.equal(filterCommands('').length, commands.length);
  assert.equal(new Set(commands.map((command) => command.id)).size, commands.length);
  const shortcuts = defaultKeymap.map((binding) => `${binding.scope}:${binding.key}`);
  assert.equal(new Set(shortcuts).size, shortcuts.length);
});
