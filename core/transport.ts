import type { LocalQueryDraft, LocalQueryConnection, LocalQueryList, LocalQueryPreview } from './local-query.mjs';
import type {
  WorkbookCommand,
  WorkbookHistoryEdit,
  WorkbookView,
  WorkbookCell,
  WorkbookFilterValuePage,
  WorkbookObjectPage,
  WorkbookObjectPreview,
  WorkbookConditionalDefinition,
  WorkbookValidationDefinition,
  WorkbookPrintDefinition,
  WorkbookPrintSettings,
  WorkbookPrintLayout,
  WorkbookPrintPage,
  WorkbookFormulaTrace,
  WorkbookGeometryProfile,
  WorkbookFormulaPreview,
  WorkbookFormulaPreviewCancellation,
  WorkbookDuplicatePreview,
  WorkbookUsedRangeResult,
  WorkbookNoteMembershipResult,
  WorkbookHyperlinkDefinition,
  WorkbookNotes,
} from './contracts.js';
import type { PasteShape, PastePlan, PastePage, PasteDescriptor, TypedPasteConversion } from './paste-protocol.mjs';
import type { CalculationAction, CalculationGroupReceipt } from '../adapters/workbook-calculation-group.mjs';
import type {
  WorkbookSpellingAvailability,
  WorkbookSpellingCheck,
  WorkbookSpellingCancellation,
} from '../adapters/workbook-spelling-jobs.mjs';
export type { WorkbookStorage } from './contracts.js';
export type WorkbookOwnerRequest =
  | { kind: 'indexed'; operation: 'open'; format?: 'xlsx' | 'csv' | 'xlsm' | 'xls' | 'xlsb' }
  | {
      kind: 'indexed';
      operation:
        | 'status'
        | 'spellingAvailability'
        | 'spellingCheck'
        | 'spellingCancel'
        | 'calculationGroup'
        | 'calculationCancel'
        | 'read'
        | 'apply'
        | 'save'
        | 'copy'
        | 'scan'
        | 'trace'
        | 'usedRange'
        | 'noteMembership'
        | 'notes'
        | 'accessibility'
        | 'hyperlink'
        | 'duplicatePreview'
        | 'previewFormula'
        | 'cancelFormulaPreview'
        | 'query'
        | 'filterValues'
        | 'selection'
        | 'objects'
        | 'object'
        | 'conditionalRules'
        | 'validationRules'
        | 'compact'
        | 'close'
        | 'undo'
        | 'redo'
        | 'adopt'
        | 'discard'
        | 'importBegin'
        | 'importUpload'
        | 'importSeal'
        | 'importCancel'
        | 'importPrune'
        | 'importReceipt'
        | 'queryConnections'
        | 'queryCreate'
        | 'queryUpdate'
        | 'queryRemove'
        | 'queryPreview'
        | 'queryRefresh'
        | 'queryCancel'
        | 'importFile'
        | 'importFileCancel'
        | 'cancelOperation'
        | 'printDefinition'
        | 'printPrepare'
        | 'printPage'
        | 'printRelease'
        | 'printPublish'
        | 'printCancel'
        | 'vbaModules'
        | 'vbaModule'
        | 'vbaFunctions'
        | 'evaluateVba'
        | 'dataOnlyCopy';
      sessionId: string;
      state: string;
      ownerId?: string;
      revision?: string;
      locale?: string;
      words?: string[];
      calculationAction?: CalculationAction;
      hasHeaders?: boolean;
      columns?: string[];
      formula?: string;
      cell?: string;
      direction?: 'precedents' | 'dependents';
      retain?: string[];
      query?: string;
      mode?: 'display' | 'raw' | 'formula';
      matchLimit?: number;
      name?: string;
      after?: string;
      offset?: number;
      column?: string;
      sheet?: string;
      range?: string;
      objectId?: string;
      limit?: number;
      commands?: WorkbookCommand[];
      encodedCommands?: string;
      expectedRevision?: string;
      documentId?: string;
      shape?: PasteShape;
      count?: number;
      policy?: PastePlan['policy'];
      conversion?: TypedPasteConversion;
      uploadId?: string;
      sequence?: number;
      cells?: PastePage['cells'];
      requestId?: string;
      descriptor?: PasteDescriptor;
      parentState?: string;
      connectionId?: string;
      connectionRevision?: number;
      queryDraft?: LocalQueryDraft;
      expectedConversionRevision?: string;
      replacementOffset?: number;
      expectedReplacementRevision?: string;
      settings?: WorkbookPrintSettings;
      geometryProfile?: WorkbookGeometryProfile;
      jobId?: string;
      pageNumber?: number;
      printKind?: 'pdf' | 'printer';
      start?: string;
      moduleId?: string;
      functionName?: string;
      args?: number[];
      dataFormat?: 'xls' | 'xlsb' | 'xlsx';
    };
export interface WorkbookOwnerHistory {
  documentId: string;
  dirty: boolean;
  canUndo: boolean;
  canRedo: boolean;
  recovered: boolean;
  head: string | null;
  tip: string | null;
  sourceRevision: string;
  sourceConflict?: boolean;
  notice?: string;
}
export interface WorkbookOwnerResult {
  usedRange?: WorkbookUsedRangeResult;
  accessibility?: import('./accessibility.mjs').AccessibilityReport & import('./notes.mjs').NoteReceipt;
  notes?: WorkbookNotes;
  hyperlink?: WorkbookHyperlinkDefinition;
  noteMembership?: WorkbookNoteMembershipResult;
  spellingAvailability?: WorkbookSpellingAvailability;
  spellingCheck?: WorkbookSpellingCheck;
  spellingCancellation?: WorkbookSpellingCancellation;
  duplicatePreview?: WorkbookDuplicatePreview;
  formulaPreview?: WorkbookFormulaPreview;
  encodedFormulaPreview?: string;
  formulaPreviewCancellation?: WorkbookFormulaPreviewCancellation;
  calculationGroup?: CalculationGroupReceipt;
  calculationCancellation?: { outcome: 'cancelled' | 'completed' | 'finishing' };
  applicationCalculation?: { mode: 'auto' | 'manual'; revision: number };
  formulaTrace?: WorkbookFormulaTrace;
  copied?: boolean;
  omissions?: string[];
  values?: number;
  formulaValues?: number;
  sourceRevision?: string;
  vbaModules?: import('../adapters/passive-workbook.mjs').PassiveVbaPage;
  vbaModule?: import('../adapters/passive-workbook.mjs').PassiveVbaSource;
  vbaFunctions?: import('../adapters/passive-workbook.mjs').PureVbaFunctionPage;
  vbaEvaluation?: import('../adapters/passive-workbook.mjs').PureVbaEvaluation;
  history?: WorkbookOwnerHistory;
  capabilities?: { durableHistory: true; stagedPaste: true; delimitedImport?: true };
  importSource?: {
    id: string;
    name: string;
    bytes: number;
    sha256: string;
    count: number;
    rows: number;
    columns: number;
    delimiter: ',' | ';' | '\t';
    encoding: 'utf-8' | 'utf-16le' | 'utf-16be';
    bom: boolean;
  };
  uploadId?: string;
  accepted?: number;
  nextSequence?: number;
  descriptor?: PasteDescriptor;
  cancelled?: boolean;
  granted?: boolean;
  reason?: 'completed' | 'finishing';
  pruned?: boolean;
  conditionalRules?: WorkbookConditionalDefinition;
  validationRules?: WorkbookValidationDefinition;
  printDefinition?: WorkbookPrintDefinition;
  printJob?: WorkbookPrintLayout & { id: string; pageCount: number };
  printPage?: WorkbookPrintPage;
  printReleased?: boolean;
  printPublished?: { entryId: string; name: string; relativePath: string; pageCount: number };
  queryConnections?: LocalQueryList;
  queryConnection?: LocalQueryConnection;
  queryRefreshed?: boolean;
  querySource?: { id: string; name: string; bytes: number; sha256: string };
  queryPreview?: LocalQueryPreview;
  warning?: string;
  outputRows?: number;
  outputColumns?: number;
  printDispatched?: { pageCount: number };
  importReceipt?: {
    found: boolean;
    head?: string;
    parentState?: string;
    descriptor?: PasteDescriptor;
    reason?: 'not-retained';
    currentState: string;
    tipState: string;
  };
  filterValues?: WorkbookFilterValuePage;
  encodedFilterValues?: string;
  objects?: WorkbookObjectPage;
  object?: WorkbookObjectPreview;
  encodedObject?: string;
  selection?: import('@office-editors/xlsx-editor/core/indexed-session.mjs').SelectionSummary;
  view?: WorkbookView;
  edits?: WorkbookHistoryEdit[];
  commands?: WorkbookCommand[];
  sheet?: string;
  range?: string;
  encodedCells?: string;
  cells?: Record<string, WorkbookCell>;
  nextOffset?: number | null;
  nextCursor?: string | null;
  matches?: number;
  scanned?: number;
  revision?: string;
  compacted?: boolean;
  closed?: boolean;
}

/** JSON transports must not turn overflow into null or negative zero into zero. */
function exceptionalJson(value: unknown): string | undefined {
  let exceptional = false;
  const json = JSON.stringify(value, (_key, item) => {
    if (typeof item === 'number' && (!Number.isFinite(item) || Object.is(item, -0))) {
      exceptional = true;
      return { $workbookNumber: Object.is(item, -0) ? '-0' : String(item) };
    }
    return item;
  });
  return exceptional ? json : undefined;
}
function restoreNumbers(json: string): any {
  return JSON.parse(json, (_key, item) =>
    item && typeof item === 'object' && Object.keys(item).length === 1 && typeof item.$workbookNumber === 'string'
      ? Number(item.$workbookNumber)
      : item
  );
}
export function encodeOwnerResult<T extends WorkbookOwnerResult>(result: T): T {
  const encodedCells = result.cells ? exceptionalJson(result.cells) : undefined;
  const encodedFilterValues = result.filterValues ? exceptionalJson(result.filterValues) : undefined;
  const encodedObject = result.object ? exceptionalJson(result.object) : undefined;
  const encodedFormulaPreview = result.formulaPreview ? exceptionalJson(result.formulaPreview) : undefined;
  return encodedCells || encodedFilterValues || encodedObject || encodedFormulaPreview
    ? {
        ...result,
        ...(encodedCells ? { cells: undefined, encodedCells } : {}),
        ...(encodedFilterValues ? { filterValues: undefined, encodedFilterValues } : {}),
        ...(encodedObject ? { object: undefined, encodedObject } : {}),
        ...(encodedFormulaPreview ? { formulaPreview: undefined, encodedFormulaPreview } : {}),
      }
    : result;
}
export function decodeOwnerResult<T extends WorkbookOwnerResult>(result: T): T {
  return result.encodedCells || result.encodedFilterValues || result.encodedObject || result.encodedFormulaPreview
    ? {
        ...result,
        ...(result.encodedCells ? { cells: restoreNumbers(result.encodedCells), encodedCells: undefined } : {}),
        ...(result.encodedFilterValues
          ? { filterValues: restoreNumbers(result.encodedFilterValues), encodedFilterValues: undefined }
          : {}),
        ...(result.encodedObject ? { object: restoreNumbers(result.encodedObject), encodedObject: undefined } : {}),
        ...(result.encodedFormulaPreview
          ? { formulaPreview: restoreNumbers(result.encodedFormulaPreview), encodedFormulaPreview: undefined }
          : {}),
      }
    : result;
}
export function ownerCommands(commands: WorkbookCommand[]): { commands?: WorkbookCommand[]; encodedCommands?: string } {
  const encodedCommands = exceptionalJson(commands);
  return encodedCommands ? { encodedCommands } : { commands };
}
export function decodedCommands(request: {
  commands?: WorkbookCommand[];
  encodedCommands?: string;
}): WorkbookCommand[] {
  return request.encodedCommands ? restoreNumbers(request.encodedCommands) : request.commands!;
}
