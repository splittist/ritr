import type { TextFormat } from '../engine/format';
import type { FormatEdit } from '../engine/format-edit';
import type { ProjectionEdit, ProjectionSelection } from '../engine/projection';
import type { KeyBinding } from './keymap';
export interface FormattingTarget {
  apply: (format: TextFormat | 'b' | 'i' | 'u') => void;
  values: TextFormat;
  editable: boolean;
}
export interface DirectEditing {
  activate: (target?: FormattingTarget) => void;
  format: (edit: Pick<FormatEdit, 'origin' | 'from' | 'to' | 'format'>) => void;
  busy: boolean;
  pending: boolean;
  composition: (active: boolean) => void;
  restore?: ProjectionSelection;
  keymap: readonly KeyBinding[];
  focus?: { spanId: string };
  change: (
    edit: Pick<
      ProjectionEdit,
      'origin' | 'from' | 'to' | 'text' | 'typingSpan' | 'history' | 'typingFormat' | 'enter'
    >,
  ) => void;
  message: (message: string) => void;
}
