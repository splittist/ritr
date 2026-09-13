import type { FormatEdit } from '../engine/format-edit';
import type { ProjectionEdit, ProjectionSelection } from '../engine/projection';
import type { TextPiece } from '../engine/text-range';
import type { ChangePreview, OpenDocument, SearchMatch, TextEdit } from '../engine/workspace';
import type { PartDifference } from '../package/docx';

export interface Snapshot {
  selection?: ProjectionSelection;
  documents: OpenDocument[];
  canUndo: boolean;
  canRedo: boolean;
}
export interface DesktopApi {
  formatProjection(edit: FormatEdit): Promise<Snapshot>;
  editProjection(edit: ProjectionEdit): Promise<Snapshot>;
  snapshot(): Promise<Snapshot>;
  open(): Promise<Snapshot>;
  previewPieces(edit: {
    documentId: string;
    storyId: string;
    pieces: TextPiece[];
    expectedRevision: number;
  }): Promise<ChangePreview>;
  previewEdit(edit: TextEdit & { expectedRevision?: number }): Promise<ChangePreview>;
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
