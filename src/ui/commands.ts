export interface CommandContext {
  connected: boolean;
  busy: boolean;
  hasDocument: boolean;
  canUndo: boolean;
  canRedo: boolean;
  hasInlineDraft: boolean;
  editReason?: string;
  selectionReason?: string;
  hasQuery: boolean;
  hasPreview: boolean;
  hasChanges: boolean;
}

interface Definition {
  id: string;
  label: string;
  description: string;
  shortcuts?: readonly string[];
  nativeUndo?: boolean;
  unavailable: (context: CommandContext) => string | undefined;
}
const documentRequired = (c: CommandContext) =>
  !c.hasDocument ? 'Open a document first.' : undefined;
const noInlineDraft = (c: CommandContext) =>
  c.hasInlineDraft ? 'Apply or cancel the inline draft first.' : undefined;
const searchRequired = (c: CommandContext) =>
  documentRequired(c) ??
  noInlineDraft(c) ??
  (!c.hasQuery ? 'Enter text in Find text first.' : undefined);

/** Shared presentation and availability for buttons, shortcuts, and the palette. */
export const commands = [
  {
    id: 'commands.open',
    label: 'Show commands',
    description: 'Search available actions and keyboard shortcuts.',
    shortcuts: ['Ctrl+Shift+P'],
    unavailable: () => undefined,
  },
  {
    id: 'file.open',
    label: 'Open documents',
    description: 'Open one or more DOCX files.',
    shortcuts: ['Ctrl+O'],
    unavailable: noInlineDraft,
  },
  {
    id: 'file.saveAs',
    label: 'Save As…',
    description: 'Save and verify a new copy of the current document.',
    shortcuts: ['Ctrl+Shift+S'],
    unavailable: (c: CommandContext) => documentRequired(c) ?? noInlineDraft(c),
  },
  {
    id: 'history.undo',
    label: 'Undo',
    description: 'Undo the last workspace transaction; clears any inline draft.',
    shortcuts: ['Ctrl+Z'],
    nativeUndo: true,
    unavailable: (c: CommandContext) =>
      !c.canUndo ? 'No workspace transaction to undo.' : undefined,
  },
  {
    id: 'history.redo',
    label: 'Redo',
    description: 'Redo the last undone workspace transaction; clears any inline draft.',
    shortcuts: ['Ctrl+Y', 'Ctrl+Shift+Z'],
    nativeUndo: true,
    unavailable: (c: CommandContext) =>
      !c.canRedo ? 'No workspace transaction to redo.' : undefined,
  },
  {
    id: 'edit.inline',
    label: 'Edit selected span inline',
    description: 'Start an inline draft at the selected text span.',
    unavailable: (c: CommandContext) =>
      documentRequired(c) ?? noInlineDraft(c) ?? c.selectionReason,
  },
  {
    id: 'edit.preview',
    label: 'Preview text edit',
    description: 'Preview the inline draft or inspector text changes.',
    shortcuts: ['Ctrl+Enter'],
    unavailable: (c: CommandContext) => documentRequired(c) ?? c.editReason,
  },
  {
    id: 'edit.cancelInline',
    label: 'Cancel inline edit',
    description: 'Discard the unapplied inline draft.',
    unavailable: (c: CommandContext) =>
      !c.hasInlineDraft ? 'There is no inline draft.' : undefined,
  },
  {
    id: 'search.focus',
    label: 'Focus workspace search',
    description: 'Enter a literal query to find text across open documents.',
    shortcuts: ['Ctrl+F'],
    unavailable: documentRequired,
  },
  {
    id: 'search.find',
    label: 'Find',
    description: 'Find the current literal query in all open documents and stories.',
    unavailable: searchRequired,
  },
  {
    id: 'search.replace',
    label: 'Preview replacement',
    description: 'Preview replacement of the current query across the workspace.',
    unavailable: searchRequired,
  },
  {
    id: 'transaction.apply',
    label: 'Apply transaction',
    description: 'Apply the currently displayed preview as one undoable change.',
    shortcuts: ['Ctrl+Shift+Enter'],
    unavailable: (c: CommandContext) =>
      !c.hasPreview
        ? 'Create a preview first.'
        : !c.hasChanges
          ? 'The preview has no editable changes.'
          : undefined,
  },
  {
    id: 'transaction.dismiss',
    label: 'Dismiss preview',
    description: 'Hide the current preview without applying it.',
    unavailable: (c: CommandContext) => (!c.hasPreview ? 'There is no preview.' : undefined),
  },
  {
    id: 'view.codes',
    label: 'Toggle Reveal Codes',
    description: 'Show or hide document formatting and structure codes.',
    shortcuts: ['Ctrl+Shift+E'],
    unavailable: documentRequired,
  },
  {
    id: 'document.compare',
    label: 'Compare package parts',
    description: 'Report changed and preserved parts against the original document.',
    unavailable: documentRequired,
  },
] as const satisfies readonly Definition[];

export type CommandId = (typeof commands)[number]['id'];
export type CommandActions = Record<CommandId, () => void | Promise<void>>;

export function commandReason(id: CommandId, context: CommandContext): string | undefined {
  if (id !== 'commands.open') {
    if (context.busy) return 'Wait for the current operation to finish.';
    if (!context.connected) return 'Open the desktop app to use this command.';
  }
  const command: Definition = commands.find((command) => command.id === id)!;
  return command.unavailable(context);
}

/** Recheck current context at invocation, even if a surface rendered earlier. */
export async function invokeCommand(
  id: CommandId,
  context: () => CommandContext,
  action: () => void | Promise<void>,
) {
  const reason = commandReason(id, context());
  if (reason) return reason;
  await action();
  return undefined;
}

export function filterCommands(query: string) {
  const terms = query.toLocaleLowerCase().trim().split(/\s+/).filter(Boolean);
  return commands.filter((command) =>
    terms.every((term) =>
      `${command.label} ${command.description}`.toLocaleLowerCase().includes(term),
    ),
  );
}

export function shortcutCommand(
  event: Pick<
    KeyboardEvent,
    'key' | 'ctrlKey' | 'metaKey' | 'altKey' | 'shiftKey' | 'repeat' | 'isComposing'
  >,
  nativeText: boolean,
): CommandId | undefined {
  if (
    event.repeat ||
    event.isComposing ||
    event.altKey ||
    !(event.ctrlKey || event.metaKey) ||
    (event.ctrlKey && event.metaKey)
  )
    return;
  return commands.find((entry) => {
    const command: Definition = entry;
    if (nativeText && command.nativeUndo) return false;
    return command.shortcuts?.some((shortcut) => {
      const keys = shortcut.toLowerCase().split('+');
      return event.key.toLowerCase() === keys.at(-1) && event.shiftKey === keys.includes('shift');
    });
  })?.id;
}
