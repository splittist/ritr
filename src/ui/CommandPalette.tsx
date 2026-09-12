import { commandShortcuts, type KeyBinding } from './keymap';
import { useEffect, useRef, useState } from 'react';
import { commandReason, filterCommands, type CommandContext, type CommandId } from './commands';

/** Capture before React mounts/autofocuses the dialog; native close may lose field selection. */
export function captureCommandFocus(): () => void {
  const target = document.activeElement;
  if (!(target instanceof HTMLElement)) return () => {};
  const field =
    target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement
      ? target
      : undefined;
  const selection =
    field?.selectionStart !== null && field?.selectionStart !== undefined
      ? {
          start: field.selectionStart,
          end: field.selectionEnd!,
          direction: field.selectionDirection!,
        }
      : undefined;
  return () => {
    if (!target.isConnected) return;
    target.focus({ preventScroll: true });
    if (field && selection)
      field.setSelectionRange(selection.start, selection.end, selection.direction);
  };
}

export function CommandPalette({
  keymap,
  context,
  onClose,
  onExecute,
  restoreFocus,
}: {
  keymap: readonly KeyBinding[];
  context: CommandContext;
  onClose: () => void;
  onExecute: (id: CommandId) => void;
  restoreFocus: () => void;
}) {
  const dialog = useRef<HTMLDialogElement>(null);
  const closed = useRef(false);
  const [query, setQuery] = useState('');
  const [index, setIndex] = useState(0);
  const results = filterCommands(query).filter((command) => command.id !== 'commands.open');
  const active = results[Math.min(index, results.length - 1)];
  useEffect(() => {
    const element = dialog.current!;
    element.showModal();
    return () => {
      if (!closed.current) {
        element.close();
        restoreFocus();
      }
    };
  }, []);
  useEffect(() => {
    if (active)
      dialog.current
        ?.querySelector(`[data-command="${active.id}"]`)
        ?.scrollIntoView({ block: 'nearest' });
  }, [active?.id]);
  const close = () => {
    closed.current = true;
    dialog.current?.close();
    restoreFocus();
    onClose();
  };
  const execute = (id: CommandId) => {
    if (commandReason(id, context)) return;
    close();
    onExecute(id);
  };
  return (
    <dialog
      ref={dialog}
      className="command-palette"
      aria-labelledby="command-title"
      onCancel={(event) => {
        event.preventDefault();
        close();
      }}
    >
      <div className="command-heading">
        <h2 id="command-title">Commands</h2>
        <button onClick={close} aria-label="Close commands">
          Esc
        </button>
      </div>
      <input
        autoFocus
        role="combobox"
        aria-label="Search commands"
        aria-expanded="true"
        aria-controls="command-results"
        aria-autocomplete="list"
        aria-activedescendant={active ? `command-${active.id}` : undefined}
        placeholder="Search commands…"
        value={query}
        onChange={(event) => {
          setQuery(event.target.value);
          setIndex(0);
        }}
        onKeyDown={(event) => {
          if (event.nativeEvent.isComposing) return;
          if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
            event.preventDefault();
            setIndex((value) =>
              results.length
                ? (value + (event.key === 'ArrowDown' ? 1 : -1) + results.length) % results.length
                : 0,
            );
          } else if (event.key === 'Enter') {
            event.preventDefault();
            if (active) execute(active.id);
          }
        }}
      />
      <div id="command-results" role="listbox" aria-label="Commands" className="command-results">
        {results.map((command, position) => {
          const reason = commandReason(command.id, context);
          return (
            <div
              key={command.id}
              id={`command-${command.id}`}
              data-command={command.id}
              role="option"
              aria-selected={active?.id === command.id}
              aria-disabled={!!reason}
              className="command-option"
              onMouseMove={() => setIndex(position)}
              onClick={() => execute(command.id)}
            >
              <div>
                <strong>{command.label}</strong>
                {commandShortcuts(command.id, keymap).length > 0 && (
                  <kbd>{commandShortcuts(command.id, keymap)[0]}</kbd>
                )}
              </div>
              <small>{reason ?? command.description}</small>
            </div>
          );
        })}
      </div>
      {!results.length && <p role="status">No commands match “{query}”.</p>}
      <p className="command-hint">
        ↑ ↓ to choose · Enter to run · Esc to close. Unavailable commands explain why.
      </p>
    </dialog>
  );
}
