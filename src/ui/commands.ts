export type EditorCommandId = 'text.deleteBackward' | 'text.deleteForward' | 'paragraph.split';

export interface CommandContext {
  connected: boolean;
  busy: boolean;
  hasDocument: boolean;
  canUndo: boolean;
  canRedo: boolean;
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
  unavailable: (context: CommandContext) => string | undefined;
}
const documentRequired = (c: CommandContext) =>
  !c.hasDocument ? 'Open a document first.' : undefined;
const searchRequired = (c: CommandContext) =>
  documentRequired(c) ?? (!c.hasQuery ? 'Enter text in Find text first.' : undefined);

/** Shared presentation and availability for buttons, shortcuts, and the palette. */
export const commands = [
  {
    id: 'commands.open',
    label: 'Show commands',
    description: 'Search available actions and keyboard shortcuts.',
    unavailable: () => undefined,
  },
  {
    id: 'file.open',
    label: 'Open documents',
    description: 'Open one or more DOCX files.',
    unavailable: () => undefined,
  },
  {
    id: 'file.saveAs',
    label: 'Save As…',
    description: 'Save and verify a new copy of the current document.',
    unavailable: (c: CommandContext) => documentRequired(c),
  },
  {
    id: 'history.undo',
    label: 'Undo',
    description: 'Undo the last workspace transaction.',
    unavailable: (c: CommandContext) =>
      !c.canUndo ? 'No workspace transaction to undo.' : undefined,
  },
  {
    id: 'history.redo',
    label: 'Redo',
    description: 'Redo the last undone workspace transaction.',
    unavailable: (c: CommandContext) =>
      !c.canRedo ? 'No workspace transaction to redo.' : undefined,
  },
  {
    id: 'edit.inline',
    label: 'Focus document text',
    description: 'Focus the selected source text in the document editor.',
    unavailable: (c: CommandContext) => documentRequired(c) ?? c.selectionReason,
  },
  {
    id: 'edit.preview',
    label: 'Preview text edit',
    description: 'Preview the inspector text changes.',
    unavailable: (c: CommandContext) => documentRequired(c) ?? c.editReason,
  },
  {
    id: 'search.focus',
    label: 'Focus workspace search',
    description: 'Enter a literal query to find text across open documents.',
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
