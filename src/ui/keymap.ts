import { commands, type CommandId, type EditorCommandId } from './commands';

export interface KeyBinding {
  command: CommandId | EditorCommandId;
  key: string;
  scope: 'global' | 'projection' | 'editor';
}
export const defaultKeymap: readonly KeyBinding[] = [
  { command: 'commands.open', key: 'Ctrl+Shift+P', scope: 'global' },
  { command: 'file.open', key: 'Ctrl+O', scope: 'global' },
  { command: 'file.saveAs', key: 'Ctrl+Shift+S', scope: 'global' },
  { command: 'history.undo', key: 'Ctrl+Z', scope: 'projection' },
  { command: 'history.redo', key: 'Ctrl+Y', scope: 'projection' },
  { command: 'history.redo', key: 'Ctrl+Shift+Z', scope: 'projection' },
  { command: 'edit.preview', key: 'Ctrl+Enter', scope: 'global' },
  { command: 'search.focus', key: 'Ctrl+F', scope: 'global' },
  { command: 'transaction.apply', key: 'Ctrl+Shift+Enter', scope: 'global' },
  { command: 'view.codes', key: 'Ctrl+Shift+E', scope: 'global' },
  { command: 'text.deleteBackward', key: 'Backspace', scope: 'editor' },
  { command: 'text.deleteForward', key: 'Delete', scope: 'editor' },
  { command: 'paragraph.split', key: 'Enter', scope: 'editor' },
];
export type KeymapOverrides = Partial<Record<CommandId | EditorCommandId, readonly string[]>>;

function chord(key: string) {
  const parts = key.toLowerCase().split('+');
  const name = parts.pop()!;
  const modifiers = parts.map((p) => (p === 'mod' ? 'ctrl' : p)).sort();
  if (
    !name ||
    modifiers.some((p) => !['ctrl', 'meta', 'alt', 'shift'].includes(p)) ||
    new Set(modifiers).size !== modifiers.length ||
    (modifiers.includes('ctrl') && modifiers.includes('meta'))
  )
    throw new Error(`Invalid key chord: ${key}`);
  return [...modifiers, name].join('+');
}
/** An override replaces all bindings for a command; [] explicitly unbinds it. */
export function resolveKeymap(overrides: KeymapOverrides = {}): readonly KeyBinding[] {
  if (!overrides || typeof overrides !== 'object' || Array.isArray(overrides))
    throw new Error('Keybindings must be a JSON object');
  const result = defaultKeymap.filter((b) => overrides[b.command] === undefined);
  const custom: KeyBinding[] = [];
  for (const [command, keys] of Object.entries(overrides)) {
    const original = defaultKeymap.find((b) => b.command === command);
    if (!original && !commands.some((c) => c.id === command))
      throw new Error(`Unknown command: ${command}`);
    if (!Array.isArray(keys) || keys.some((k) => typeof k !== 'string'))
      throw new Error(`Expected an array of key chords for ${command}`);
    const scope = original?.scope ?? 'global';
    for (const key of keys) {
      const normalized = chord(key);
      if (custom.some((b) => chord(b.key) === normalized))
        throw new Error(`Conflicting key chord: ${key}`);
      for (let i = result.length - 1; i >= 0; i--)
        if (chord(result[i]!.key) === normalized) result.splice(i, 1);
      custom.push({ command: command as KeyBinding['command'], key, scope });
    }
  }
  return [...result, ...custom];
}
export function commandShortcuts(id: CommandId, keymap: readonly KeyBinding[] = defaultKeymap) {
  return keymap.filter((b) => b.command === id).map((b) => b.key);
}
type KeyEvent = Pick<
  KeyboardEvent,
  'key' | 'ctrlKey' | 'metaKey' | 'altKey' | 'shiftKey' | 'repeat' | 'isComposing'
>;
export function matchKey(event: KeyEvent, binding: KeyBinding): boolean {
  if (event.isComposing || (event.repeat && binding.scope !== 'editor')) return false;
  const parts = binding.key.toLowerCase().split('+');
  const mod = parts.includes('ctrl') || parts.includes('mod');
  return (
    event.key.toLowerCase() === parts.at(-1) &&
    (mod
      ? event.ctrlKey !== event.metaKey
      : !event.ctrlKey && event.metaKey === parts.includes('meta')) &&
    event.altKey === parts.includes('alt') &&
    event.shiftKey === parts.includes('shift')
  );
}
export function shortcutCommand(
  event: KeyEvent,
  nativeText: boolean,
  keymap: readonly KeyBinding[] = defaultKeymap,
): CommandId | undefined {
  return keymap.find(
    (b) => b.scope !== 'editor' && (!nativeText || b.scope !== 'projection') && matchKey(event, b),
  )?.command as CommandId | undefined;
}
export function editorCommand(
  event: KeyEvent,
  keymap: readonly KeyBinding[] = defaultKeymap,
): EditorCommandId | undefined {
  return keymap.find((b) => b.scope === 'editor' && matchKey(event, b))?.command as
    EditorCommandId | undefined;
}
