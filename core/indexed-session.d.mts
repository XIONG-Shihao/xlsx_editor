import type {
  WorkbookCommand,
  WorkbookHistoryEdit,
  WorkbookObjectPage,
  WorkbookObjectPreview,
  WorkbookPrintDefinition,
  WorkbookPrintSettings,
  WorkbookPrintSnapshot,
  WorkbookConditionalDefinition,
  WorkbookValidationDefinition,
  WorkbookFormulaTrace,
  WorkbookTraceOptions,
  WorkbookGeometryProfile,
  WorkbookFormulaPreview,
  WorkbookFormulaPreviewCancellation,
  WorkbookDuplicatePreview,
  WorkbookUsedRangeInner,
  WorkbookNoteMembershipInner,
  WorkbookHyperlinkInner,
} from './contracts.js';
import type { PasteDescriptor, PastePage } from './paste-protocol.mjs';
export interface IndexedWorkbookContext {
  importPages?: (descriptor: PasteDescriptor) => AsyncIterable<PastePage>;
  importProgress?: (event: { phase: 'applying-paste'; completed: number; total: number }) => void;
}
export interface SelectionSummary {
  count: number;
  numericCount: number;
  sum: number | null;
  average: number | null;
  min: number | null;
  max: number | null;
  errors: number;
  unresolved: number;
  unsafeNumbers: number;
}
export interface FilterValuePage {
  values: Array<{ key: string; value: string | number | boolean | null; count: number }>;
  total: number;
  nextOffset: number | null;
  protected: { errors: number; unresolved: number; special: number; unsafeNumbers: number };
}
/** Private owner token; never serialized into renderer requests or recovery. */
export interface PublicationRebase {
  readonly id: string;
  readonly sourceRevision: string;
}
/** Trusted private checkpoint; never supplied by renderer or external workbook. */
export interface PublicationCalculationCheckpoint {
  readonly path: string;
  readonly revision: string;
  readonly bytes: number;
}
/** Projection bytes are serialized accounting estimates, never physical/retained heap. */
export interface IndexedWorkbookDiagnostics {
  retainedStates: number;
  storageRoots: number;
  closed: 0 | 1;
  calculationCalls: number;
  lastProjectedBytes: number;
  peakProjectedBytes: number;
  lastInputCells: number;
  peakInputCells: number;
  lastEvaluatedFormulaCells: number;
  lastExportStyleCount: number;
  lastExportMetadataBytes: number;
}
export class IndexedWorkbook {
  cacheHit?: boolean;
  cachePublication?: { published: boolean; reason?: string; bytes?: number; existing?: boolean };
  readonly id: string;
  constructor(
    native: (input: Record<string, unknown>) => Promise<any>,
    legacy: (
      source: string,
      commands: any[],
      output: string,
      signal?: AbortSignal
    ) => Promise<{ revision: string; edits: any[] }>,
    calculate: (input: Record<string, unknown>, signal?: AbortSignal) => Promise<{ version: number }>,
    context?: IndexedWorkbookContext
  );
  open(
    source: string,
    options?: {
      cacheIdentity?: string;
      cacheMode?: 'saved-only' | 'all';
      maxSessionBytes?: number;
      storageDirectory?: string;
      cacheOptions?: { directory?: string; maxBytes?: number; maxEntries?: number };
      initialCalculationPolicy?: { mode: 'auto' | 'manual'; recalculateBeforeSave: boolean };
      savedCalculationCheckpoint?: PublicationCalculationCheckpoint;
    }
  ): Promise<any>;
  view(state: string): any;
  diagnosticStats(): IndexedWorkbookDiagnostics;
  read(
    state: string,
    sheet: string,
    range: string
  ): Promise<{ sheet: string; range: string; cells: Record<string, any> }>;
  scan(
    state: string,
    sheet: string,
    after?: string,
    options?: { limit?: number }
  ): Promise<{ cells: Record<string, any>; nextCursor: string | null }>;
  trace(state: string, sheet: string, cell: string, options: WorkbookTraceOptions): Promise<WorkbookFormulaTrace>;
  usedRange(state: string, sheet: string): Promise<WorkbookUsedRangeInner>;
  noteMembership(state: string, sheet: string): Promise<WorkbookNoteMembershipInner>;
  accessibility(state: string): Promise<import('./accessibility.mjs').AccessibilityReport>;
  notes(state: string, sheet: string): Promise<import('./notes.mjs').NoteInventory>;
  hyperlink(state: string, sheet: string, cell: string): Promise<WorkbookHyperlinkInner>;
  duplicatePreview(
    state: string,
    sheet: string,
    range: string,
    hasHeaders: boolean,
    columns: string[],
    offset?: number
  ): Promise<WorkbookDuplicatePreview>;
  previewFormula(
    state: string,
    sheet: string,
    cell: string,
    formula: string,
    options: { requestId: string }
  ): Promise<Omit<WorkbookFormulaPreview, 'receipt'>>;
  cancelFormulaPreview(state: string, requestId: string): Promise<WorkbookFormulaPreviewCancellation>;
  objects(state: string, sheet: string, options?: { after?: string; limit?: number }): Promise<WorkbookObjectPage>;
  object(state: string, sheet: string, objectId: string): Promise<WorkbookObjectPreview>;
  printDefinition(state: string, sheet: string): Promise<WorkbookPrintDefinition>;
  conditionalRules(state: string, sheet: string): Promise<WorkbookConditionalDefinition>;
  validationRules(state: string, sheet: string): Promise<WorkbookValidationDefinition>;
  printSnapshot(
    state: string,
    sheet: string,
    settings?: WorkbookPrintSettings,
    geometryProfile?: WorkbookGeometryProfile
  ): Promise<WorkbookPrintSnapshot>;
  query(
    state: string,
    sheet: string,
    query: string,
    mode: 'display' | 'raw' | 'formula',
    after?: string,
    options?: { range?: string; matchLimit?: number }
  ): Promise<{ cells: Record<string, any>; matches: number; scanned: number; nextCursor: string | null }>;
  keys(
    state: string,
    sheet: string,
    range: string,
    column: string,
    values: any[],
    options?: Record<string, unknown>
  ): Promise<any>;
  selectionSummary(state: string, sheet: string, range: string): Promise<SelectionSummary>;
  filterValues(
    state: string,
    sheet: string,
    range: string,
    column: string,
    query?: string,
    offset?: number,
    limit?: number
  ): Promise<FilterValuePage>;
  calculationSummary(state: string): Promise<any>;
  affected(before: string, after: string): Promise<{ affected: any[]; total: number }>;
  apply(
    state: string,
    commands: WorkbookCommand[]
  ): Promise<{ view: any; edits: WorkbookHistoryEdit[]; commands?: WorkbookCommand[]; unchanged?: true }>;
  export(state: string, options?: { publication?: boolean }): Promise<{ path: string; revision: string }>;
  preparePublicationRebase(
    state: string,
    source: string,
    expectedRevision: string,
    options?: { calculationCheckpoint?: PublicationCalculationCheckpoint }
  ): Promise<PublicationRebase>;
  assertPublicationRebase(token: PublicationRebase): void;
  publicationCheckpoint(token: PublicationRebase): PublicationCalculationCheckpoint;
  adoptPublicationRebase(token: PublicationRebase): { view: any };
  discardPublicationRebase(token: PublicationRebase): Promise<void>;
  compact(retain: string[], options?: { retainBase?: boolean }): Promise<any>;
  close(): Promise<void>;
}
