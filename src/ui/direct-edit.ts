import type { ProjectionEdit, ProjectionSelection } from '../engine/projection';
import type { KeyBinding } from './keymap';
export interface DirectEditing {
  busy: boolean;
  pending: boolean;
  composition: (active: boolean) => void;
  restore?: ProjectionSelection;
  keymap: readonly KeyBinding[];
  focus?: { spanId: string };
  change: (edit: Pick<ProjectionEdit, 'origin' | 'from' | 'to' | 'text' | 'typingSpan'>) => void;
  message: (message: string) => void;
}
