import { editListProjection } from './list-edit';
import { formatProjection, type FormatEdit } from './format-edit';
import { validateFormat } from './format';
import { editProjection } from './paragraph-edit';
import { projectStory, type ProjectionEdit, type ProjectionSelection } from './projection';
import {
  resolveTextRange,
  replacePieces,
  editingSegments,
  type TextPosition,
  type TextPiece,
} from './text-range';
import { DocxPackage, comparePackages, type PartDifference } from '../package/docx';
import { assertText, descendants, textPatch } from '../package/xml';
import { readDocument, spans, type DocumentModel } from './document';
import { searchStory, type SearchRange } from './search';

export interface OpenDocument {
  id: string;
  name: string;
  revision: number;
  dirty: boolean;
  model: DocumentModel;
}
export interface TextEdit {
  documentId: string;
  spanId: string;
  from: number;
  to: number;
  text: string;
}
export interface SearchMatch extends TextEdit {
  /** All matched slices; inherited spanId/from/to identify the first slice for navigation. */
  ranges: SearchRange[];
  storyId: string;
  before: string;
  editable: boolean;
  reason?: string;
}
export interface ChangePreview {
  id: string;
  label: string;
  edits: { documentId: string; name: string; spanId: string; before: string; after: string }[];
  skipped: number;
}
interface Entry {
  id: string;
  name: string;
  revision: number;
  initial: DocxPackage;
  saved: DocxPackage;
  current: DocxPackage;
}
interface Transaction {
  history?: { id: string; kind: string; barrier: number };
  selectionBefore?: ProjectionSelection;
  selectionAfter?: ProjectionSelection;
  label: string;
  before: Map<string, DocxPackage>;
  after: Map<string, DocxPackage>;
}
interface Pending {
  epoch: number;
  preview: ChangePreview;
  transaction: Transaction;
}
export interface WorkspaceEvent {
  type: 'open' | 'commit' | 'undo' | 'redo' | 'save';
  documentIds: string[];
}

/** The only mutable engine object. A transaction is staged in full before publication. */
export class Workspace {
  selection?: ProjectionSelection;
  private typing?: ProjectionSelection & { spanId?: string };
  private entries = new Map<string, Entry>();
  private pending = new Map<string, Pending>();
  private past: Transaction[] = [];
  private future: Transaction[] = [];
  private listeners = new Set<(event: WorkspaceEvent) => void>();
  private epoch = 0;
  private historyBarrier = 0;
  private nextId = 1;
  subscribe(listener: (event: WorkspaceEvent) => void): () => void {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }
  private publish(type: WorkspaceEvent['type'], ids: string[]) {
    if (!['undo', 'redo'].includes(type)) this.selection = undefined;
    if (type !== 'commit') {
      this.typing = undefined;
      this.historyBarrier++;
    }
    this.epoch++;
    this.pending.clear();
    // Observer failures cannot turn a committed transaction into an apparent failure.
    for (const listener of this.listeners) {
      try {
        listener({ type, documentIds: ids });
      } catch {
        /* observers own their failures */
      }
    }
  }
  private entry(id: string): Entry {
    const entry = this.entries.get(id);
    if (!entry) throw new Error(`Unknown document: ${id}`);
    return entry;
  }
  open(name: string, bytes: Uint8Array): string {
    const pkg = DocxPackage.open(bytes);
    readDocument(pkg); // fail before adding an unreadable document
    const id = `doc-${this.nextId++}`;
    this.entries.set(id, { id, name, revision: 0, initial: pkg, saved: pkg, current: pkg });
    this.publish('open', [id]);
    return id;
  }
  documents(): OpenDocument[] {
    return [...this.entries.keys()].map((id) => this.document(id));
  }
  document(id: string): OpenDocument {
    const e = this.entry(id);
    return {
      id,
      name: e.name,
      revision: e.revision,
      dirty: e.current !== e.saved,
      model: readDocument(e.current),
    };
  }
  package(id: string): DocxPackage {
    return this.entry(id).current;
  }
  report(id: string): PartDifference[] {
    const e = this.entry(id);
    return comparePackages(e.initial, e.current);
  }
  get canUndo(): boolean {
    return this.past.length > 0;
  }
  get canRedo(): boolean {
    return this.future.length > 0;
  }

  search(query: string, caseSensitive = true): SearchMatch[] {
    if (!query) return [];
    const result: SearchMatch[] = [];
    // RegExp's case folding retains UTF-16 offsets (lowercasing can change length).
    const expression = new RegExp(
      query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
      caseSensitive ? 'gu' : 'giu',
    );
    for (const document of this.documents())
      for (const story of document.model.stories)
        for (const match of searchStory(story, expression))
          result.push({
            ...match,
            ...match.ranges[0]!,
            documentId: document.id,
            storyId: story.id,
            text: '',
          });
    return result;
  }
  previewReplace(query: string, replacement: string, caseSensitive = true): ChangePreview {
    if (!query) throw new Error('Search text cannot be empty');
    assertText(replacement);
    const matches = this.search(query, caseSensitive);
    return this.preview(
      `Replace “${query}”`,
      matches
        .filter((m) => m.editable && m.before !== replacement)
        .flatMap((m) =>
          m.ranges.map((range, index) => ({
            ...range,
            documentId: m.documentId,
            text: index === 0 ? replacement : '',
          })),
        ),
      matches.filter((m) => !m.editable).length,
    );
  }
  applyProjection(edit: ProjectionEdit): void {
    const entry = this.entry(edit.documentId);
    if (entry.revision !== edit.expectedRevision) throw new Error('The editing snapshot is stale');
    const continuation =
      this.typing &&
      edit.from === edit.to &&
      edit.from === this.typing.head &&
      edit.documentId === this.typing.documentId &&
      edit.storyId === this.typing.storyId &&
      edit.origin === this.typing.origin;
    const inputStory = readDocument(entry.current).stories.find((s) => s.id === edit.storyId);
    const projection = inputStory && projectStory(inputStory, edit.origin);
    const selectedText = projection?.text.slice(edit.from, edit.to) ?? '';
    const emptyListEnter =
      edit.enter === true &&
      edit.text === '\n' &&
      edit.from === edit.to &&
      projection?.paragraphs.some((p) => p.from === edit.from && p.to === edit.from && p.numbering);
    const listChange = edit.listAction !== undefined || emptyListEnter;
    const history = edit.history;
    if (
      history &&
      (typeof history.id !== 'string' ||
        !history.id.length ||
        history.id.length > 200 ||
        !['typing', 'backspace', 'delete'].includes(history.kind))
    )
      throw new Error('Invalid input history group');
    const groupable =
      !listChange &&
      history &&
      !edit.text.includes('\n') &&
      !selectedText.includes('\n') &&
      (history.kind === 'typing'
        ? edit.from === edit.to && !!edit.text
        : !edit.text && edit.from < edit.to);
    const previous = this.past.at(-1);
    const previousCaret = previous?.selectionAfter;
    const merge =
      groupable &&
      previous?.history?.id === history.id &&
      previous.history.kind === history.kind &&
      previous.history.barrier === this.historyBarrier &&
      previous.after.get(entry.id) === entry.current &&
      previousCaret?.documentId === edit.documentId &&
      previousCaret.storyId === edit.storyId &&
      previousCaret.origin === edit.origin &&
      previousCaret.head === (history.kind === 'backspace' ? edit.to : edit.from);
    const result: {
      pkg: DocxPackage;
      typingSpan?: string;
      selection?: Pick<ProjectionSelection, 'anchor' | 'head'>;
      label?: string;
    } = listChange
      ? editListProjection(
          entry.current,
          emptyListEnter && edit.listAction === undefined
            ? { ...edit, text: '', listAction: 'enter' }
            : edit,
        )
      : editProjection(entry.current, {
          ...edit,
          typingSpan: edit.typingSpan ?? (continuation ? this.typing?.spanId : undefined),
        });
    const selectionAfter = result.selection ?? {
      anchor: edit.from + edit.text.length,
      head: edit.from + edit.text.length,
    };
    if (edit.typingFormat) validateFormat(edit.typingFormat);
    const after =
      edit.typingFormat && edit.text.replace(/\n/g, '')
        ? formatProjection(result.pkg, {
            ...edit,
            from: edit.from,
            to: edit.from + edit.text.length,
            format: edit.typingFormat,
          })
        : result.pkg;
    if (after === entry.current) return;
    DocxPackage.open(after.save());
    const id = `change-${this.nextId++}`;
    const label =
      result.label ?? (edit.text.includes('\n') ? 'Split paragraph' : 'Edit document text');
    this.pending.clear();
    this.pending.set(id, {
      epoch: this.epoch,
      preview: { id, label, edits: [], skipped: 0 },
      transaction: {
        history: groupable ? { ...history, barrier: this.historyBarrier } : undefined,
        label,
        selectionBefore: {
          documentId: edit.documentId,
          storyId: edit.storyId,
          origin: edit.origin,
          anchor: groupable && history.kind === 'backspace' ? edit.to : edit.from,
          head: groupable && history.kind === 'delete' ? edit.from : edit.to,
        },
        selectionAfter: {
          documentId: edit.documentId,
          storyId: edit.storyId,
          origin: edit.origin,
          ...selectionAfter,
        },
        before: new Map([[entry.id, entry.current]]),
        after: new Map([[entry.id, after]]),
      },
    });
    this.commitStaged(id, !!merge);
    this.typing = {
      documentId: edit.documentId,
      storyId: edit.storyId,
      origin: edit.origin,
      ...selectionAfter,
      spanId: result.typingSpan,
    };
  }
  applyFormat(edit: FormatEdit): void {
    const entry = this.entry(edit.documentId);
    if (entry.revision !== edit.expectedRevision)
      throw new Error('The formatting selection is stale');
    const after = formatProjection(entry.current, edit);
    if (after === entry.current) return;
    DocxPackage.open(after.save());
    const selection = {
      documentId: edit.documentId,
      storyId: edit.storyId,
      origin: edit.origin,
      anchor: edit.from,
      head: edit.to,
    };
    const id = `change-${this.nextId++}`;
    this.pending.clear();
    this.pending.set(id, {
      epoch: this.epoch,
      preview: { id, label: 'Format text', edits: [], skipped: 0 },
      transaction: {
        label: 'Format text',
        before: new Map([[entry.id, entry.current]]),
        after: new Map([[entry.id, after]]),
        selectionBefore: selection,
        selectionAfter: selection,
      },
    });
    this.commit(id);
    this.typing = undefined;
    this.selection = selection;
  }
  previewRange(edit: {
    documentId: string;
    storyId: string;
    anchor: TextPosition;
    head: TextPosition;
    text: string;
    expectedRevision: number;
  }): ChangePreview & { caret: TextPosition } {
    const document = this.document(edit.documentId);
    if (document.revision !== edit.expectedRevision)
      throw new Error('Inline draft is stale; cancel it and edit the current text.');
    const story = document.model.stories.find((s) => s.id === edit.storyId);
    if (!story) throw new Error('Unknown story');
    const range = resolveTextRange(story, edit.anchor, edit.head);
    const result = replacePieces(range.pieces, range.from, range.to, edit.text, edit.anchor);
    const preview = this.previewPieces({ ...edit, pieces: result.pieces });
    return { ...preview, caret: result.caret };
  }
  previewPieces(edit: {
    documentId: string;
    storyId: string;
    pieces: TextPiece[];
    expectedRevision: number;
  }): ChangePreview {
    const document = this.document(edit.documentId);
    if (document.revision !== edit.expectedRevision)
      throw new Error('Inline draft is stale; cancel it and edit the current text.');
    const story = document.model.stories.find((s) => s.id === edit.storyId);
    const segment =
      story &&
      editingSegments(story).find(
        (s) =>
          s.spans.length === edit.pieces.length &&
          s.spans.every((span, i) => span.id === edit.pieces[i]?.spanId),
      );
    if (!segment)
      throw new Error('Draft crosses a structural boundary or has invalid source spans');
    return this.preview(
      'Edit text range',
      edit.pieces.flatMap((piece, i) => {
        const span = segment.spans[i]!;
        return span.text === piece.text
          ? []
          : [
              {
                documentId: edit.documentId,
                spanId: span.id,
                from: 0,
                to: span.text.length,
                text: piece.text,
              },
            ];
      }),
    );
  }
  preview(label: string, edits: readonly TextEdit[], skipped = 0): ChangePreview {
    const before = new Map<string, DocxPackage>(),
      after = new Map<string, DocxPackage>();
    const preview: ChangePreview = { id: `change-${this.nextId++}`, label, edits: [], skipped };
    const byDocument = new Map<string, TextEdit[]>();
    for (const edit of edits) {
      const group = byDocument.get(edit.documentId) ?? [];
      group.push({ ...edit });
      byDocument.set(edit.documentId, group);
    }
    for (const [documentId, documentEdits] of byDocument) {
      const entry = this.entry(documentId);
      const index = new Map(spans(readDocument(entry.current)).map((span) => [span.id, span]));
      const bySpan = new Map<string, TextEdit[]>();
      for (const edit of documentEdits) {
        const group = bySpan.get(edit.spanId) ?? [];
        group.push(edit);
        bySpan.set(edit.spanId, group);
      }
      const partPatches = new Map<string, ReturnType<typeof textPatch>[]>();
      for (const [spanId, changes] of bySpan) {
        const span = index.get(spanId);
        if (!span) throw new Error(`Unknown text span: ${spanId}`);
        if (!span.editable) throw new Error(span.reason ?? 'Protected text');
        let value = span.text;
        let boundary = value.length + 1;
        const splitsPair = (offset: number) =>
          offset > 0 &&
          offset < span.text.length &&
          /[\ud800-\udbff]/.test(span.text[offset - 1]!) &&
          /[\udc00-\udfff]/.test(span.text[offset]!);
        for (const edit of [...changes].sort((a, b) => b.from - a.from || b.to - a.to)) {
          assertText(edit.text);
          if (
            !Number.isInteger(edit.from) ||
            !Number.isInteger(edit.to) ||
            edit.from < 0 ||
            edit.to > span.text.length ||
            edit.to < edit.from ||
            edit.to > boundary ||
            edit.from === boundary ||
            splitsPair(edit.from) ||
            splitsPair(edit.to)
          )
            throw new Error('Invalid, overlapping, or split-surrogate edit ranges');
          value = value.slice(0, edit.from) + edit.text + value.slice(edit.to);
          boundary = edit.from;
        }
        if (value === span.text) continue;
        const source = entry.current.text(span.source.part);
        const node = descendants(entry.current.xml(span.source.part)).find((n) => n.id === spanId)!;
        const patches = partPatches.get(span.source.part) ?? [];
        patches.push(textPatch(source, node, value));
        partPatches.set(span.source.part, patches);
        preview.edits.push({
          documentId,
          name: entry.name,
          spanId,
          before: span.text,
          after: value,
        });
      }
      if (!partPatches.size) continue;
      const patched = entry.current.withXmlPatches(partPatches);
      // Save/reopen validation happens before any document is changed.
      DocxPackage.open(patched.save());
      before.set(documentId, entry.current);
      after.set(documentId, patched);
    }
    this.pending.clear(); // one bounded, reviewable proposal at a time
    this.pending.set(preview.id, {
      epoch: this.epoch,
      preview: structuredClone(preview),
      transaction: { label, before, after },
    });
    return preview;
  }
  commit(previewId: string): void {
    this.commitStaged(previewId);
  }
  private commitStaged(previewId: string, merge = false): void {
    const staged = this.pending.get(previewId);
    if (!staged || staged.epoch !== this.epoch)
      throw new Error('Preview is stale; create a new preview');
    this.pending.clear();
    if (!staged.transaction.after.size) return;
    for (const [id, pkg] of staged.transaction.after) {
      const e = this.entry(id);
      e.current = pkg;
      e.revision++;
    }
    if (merge) {
      const previous = this.past.pop()!;
      staged.transaction.before = previous.before;
      staged.transaction.selectionBefore = previous.selectionBefore;
    }
    this.past.push(staged.transaction);
    this.future = [];
    this.publish('commit', [...staged.transaction.after.keys()]);
  }
  undo(): void {
    const transaction = this.past.pop();
    if (!transaction) return;
    for (const [id, pkg] of transaction.before) {
      const e = this.entry(id);
      e.current = pkg;
      e.revision++;
    }
    this.future.push(transaction);
    this.selection = transaction.selectionBefore;
    this.publish('undo', [...transaction.before.keys()]);
  }
  redo(): void {
    const transaction = this.future.pop();
    if (!transaction) return;
    for (const [id, pkg] of transaction.after) {
      const e = this.entry(id);
      e.current = pkg;
      e.revision++;
    }
    this.past.push(transaction);
    this.selection = transaction.selectionAfter;
    this.publish('redo', [...transaction.after.keys()]);
  }
  markSaved(id: string, savedPackage: DocxPackage): void {
    const e = this.entry(id);
    // A save completing after an edit must not clear the new edit's dirty flag.
    e.saved = savedPackage;
    this.publish('save', [id]);
  }
}
