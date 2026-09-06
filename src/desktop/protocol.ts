import type { ChangePreview, OpenDocument, SearchMatch, TextEdit } from '../engine/workspace';
import type { PartDifference } from '../package/docx';

export interface Snapshot {
  documents: OpenDocument[];
  canUndo: boolean;
  canRedo: boolean;
}
export interface DesktopApi {
  snapshot(): Promise<Snapshot>;
  open(): Promise<Snapshot>;
  previewEdit(edit: TextEdit): Promise<ChangePreview>;
  previewReplace(
    query: string,
    replacement: string,
    caseSensitive: boolean,
  ): Promise<ChangePreview>;
  search(query: string, caseSensitive: boolean): Promise<SearchMatch[]>;
  commit(id: string): Promise<Snapshot>;
  undo(): Promise<Snapshot>;
  redo(): Promise<Snapshot>;
  save(id: string): Promise<{ path?: string; snapshot: Snapshot; report: PartDifference[] }>;
  report(id: string): Promise<PartDifference[]>;
}
