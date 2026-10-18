import type {
  CapabilityReport,
  DocumentStats,
  OfficeKind,
} from '@mini-office/shared';
import type { DocModel } from './word';
import type { WorkbookModel } from './sheet';
import type { SlideDeck } from './slides';

export type AnyDocumentData = DocModel | WorkbookModel | SlideDeck;

export type LifecycleState =
  | 'UNOPENED'
  | 'VALIDATING'
  | 'LOADING'
  | 'VIEW_READY'
  | 'READ_ONLY'
  | 'EDIT_READY'
  | 'DIRTY'
  | 'SAVING'
  | 'SAVED'
  | 'SAVE_FAILED'
  | 'EXPORTING'
  | 'ERROR';

export type LifecycleEvent =
  | 'VALIDATE'
  | 'LOAD'
  | 'VIEW'
  | 'READONLY'
  | 'EDIT'
  | 'CHANGE'
  | 'SAVE'
  | 'SAVED'
  | 'SAVE_FAILED'
  | 'EXPORT'
  | 'EXPORT_DONE'
  | 'ERROR'
  | 'RESET';

const TRANSITIONS: Record<LifecycleState, Partial<Record<LifecycleEvent, LifecycleState>>> = {
  UNOPENED: { VALIDATE: 'VALIDATING' },
  VALIDATING: { LOAD: 'LOADING', ERROR: 'ERROR' },
  LOADING: { VIEW: 'VIEW_READY', READONLY: 'READ_ONLY', EDIT: 'EDIT_READY', ERROR: 'ERROR' },
  VIEW_READY: {
    EDIT: 'EDIT_READY',
    READONLY: 'READ_ONLY',
    CHANGE: 'DIRTY',
    EXPORT: 'EXPORTING',
    ERROR: 'ERROR',
  },
  READ_ONLY: { CHANGE: 'READ_ONLY', EXPORT: 'EXPORTING', ERROR: 'ERROR' },
  EDIT_READY: { CHANGE: 'DIRTY', SAVE: 'SAVING', EXPORT: 'EXPORTING', READONLY: 'READ_ONLY', ERROR: 'ERROR' },
  DIRTY: { SAVE: 'SAVING', CHANGE: 'DIRTY', EXPORT: 'EXPORTING', ERROR: 'ERROR' },
  SAVING: { SAVED: 'SAVED', SAVE_FAILED: 'SAVE_FAILED', ERROR: 'ERROR' },
  SAVED: { CHANGE: 'DIRTY', EDIT: 'EDIT_READY', SAVE: 'SAVING', EXPORT: 'EXPORTING', ERROR: 'ERROR' },
  SAVE_FAILED: { SAVE: 'SAVING', CHANGE: 'DIRTY', EDIT: 'EDIT_READY', ERROR: 'ERROR' },
  EXPORTING: { EXPORT_DONE: 'VIEW_READY', ERROR: 'ERROR' },
  ERROR: { RESET: 'VIEW_READY', VALIDATE: 'VALIDATING' },
};

export function transition(from: LifecycleState, event: LifecycleEvent): LifecycleState {
  return TRANSITIONS[from]?.[event] ?? from;
}

export function canEdit(state: LifecycleState): boolean {
  return ['VIEW_READY', 'EDIT_READY', 'DIRTY', 'SAVED', 'SAVE_FAILED', 'SAVING'].includes(state);
}

export function isBusyState(state: LifecycleState): boolean {
  return state === 'VALIDATING' || state === 'LOADING' || state === 'SAVING' || state === 'EXPORTING';
}

export interface SessionView {
  zoom?: number;
  activeSheetId?: string;
  slideIndex?: number;
  showNotes?: boolean;
  presenting?: boolean;
  mode?: 'pages' | 'continuous';
  page?: number;
}

export interface DocumentSession<T extends AnyDocumentData = AnyDocumentData> {
  id: string;
  kind: OfficeKind;
  displayName: string;
  path?: string;
  readonly: boolean;
  capability: CapabilityReport;
  state: LifecycleState;
  data: T;
  dirty: boolean;
  isNew: boolean;
  restoredFromSnapshot: boolean;
  canSaveInPlace: boolean;
  /** True when the file contains features that require saving a copy first. */
  saveAsCopyRequired: boolean;
  sourceSize: number;
  createdAt: number;
  updatedAt: number;
  lastSavedAt?: number;
  saveError?: string;
  view: SessionView;
}

export interface OpenDocumentInput {
  name: string;
  bytes: Uint8Array;
  path?: string;
  knownFonts?: string[];
}

export interface FormatOpenResult<T extends AnyDocumentData = AnyDocumentData> {
  data: T;
  capability: CapabilityReport;
  readonly?: boolean;
  displayName?: string;
}

export interface SerializeContext {
  sourceBytes?: Uint8Array;
  name: string;
}

export interface FormatHandler<T extends AnyDocumentData = AnyDocumentData> {
  readonly kind: OfficeKind;
  readonly extensions: string[];
  createBlank(): T;
  open(input: OpenDocumentInput): Promise<FormatOpenResult<T>>;
  serialize(data: T, context: SerializeContext): Promise<Uint8Array>;
  getStats?(data: T): DocumentStats;
}

export type AnyFormatHandler = FormatHandler<AnyDocumentData>;

export type FormatRegistry = Partial<Record<OfficeKind, AnyFormatHandler>>;

export function initialViewFor(kind: OfficeKind): SessionView {
  if (kind === 'xlsx') return { zoom: 1, activeSheetId: undefined };
  if (kind === 'pptx') return { slideIndex: 0, showNotes: false, presenting: false };
  return { zoom: 1, mode: 'pages', page: 0 };
}

export function stateAfterOpen(readonly: boolean, hasEdits = false): LifecycleState {
  if (readonly) return 'READ_ONLY';
  return hasEdits ? 'DIRTY' : 'EDIT_READY';
}

export function fileBaseName(name: string): string {
  const slash = Math.max(name.lastIndexOf('/'), name.lastIndexOf('\\'));
  return slash >= 0 ? name.slice(slash + 1) : name;
}
