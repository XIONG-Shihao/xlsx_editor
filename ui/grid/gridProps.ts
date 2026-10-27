import type { ConditionalRulesProps, DataValidationProps } from '../WorkbookRules';
import type { DuplicateReceipt } from '../duplicateSelection';
import type { FilterValuePage } from './filterValues';
import type { FormulaPreviewControls } from '../formulaDraftPreview';
import type { FormulaTraceFrame, FormulaTraceReceipt } from '../formulaTraceAdapter';
import type { HyperlinkFollowResult } from '../hyperlinkController';
import type { PasteInspection } from '../largePaste';
import type { ReadRange } from '../rangeCells';
import type { SearchQuery } from '../searchPages';
import type { SelectionSummary } from '../../core/indexed-session.mjs';
import type {
  WorkbookAccessibilityControls,
  WorkbookCell,
  WorkbookCommand,
  WorkbookDuplicatePreview,
  WorkbookEdit,
  WorkbookHyperlinkDefinition,
  WorkbookHyperlinkReceipt,
  WorkbookNoteMembershipControls,
  WorkbookNotesControls,
  WorkbookSheet,
  WorkbookUsedRangeControls,
  WorkbookView,
} from '../../core/contracts';
import type { WorkbookActionsControls } from '../WorkbookActions';
import type { WorkbookCalculationControls } from '../calculationJobs';
import type { WorkbookObjectsProps } from '../WorkbookObjects';
import type { WorkbookPrintProps } from '../WorkbookPrint';
import type { WorkbookQueryControls } from '../WorkbookQueries';
import type { WorkbookReviewControls } from '../reviewHostAdapter';

export interface SpreadsheetGridProps {
  sheets: WorkbookSheet[];
  /** Native page-theme hints opt this host into the closed sample authoring profile. */
  pageTheme?: WorkbookView['pageTheme'];
  active?: boolean;
  scan?: (
    sheet: string,
    after?: string,
    limit?: number
  ) => Promise<{ cells: Record<string, WorkbookCell>; nextCursor: string | null }>;
  query?: SearchQuery;
  duplicatePreview?: (
    sheet: string,
    range: string,
    hasHeaders: boolean,
    columns: string[],
    offset: number
  ) => Promise<WorkbookDuplicatePreview>;
  duplicatePreviewReceipt?: DuplicateReceipt;
  filterValues?: (
    sheet: string,
    range: string,
    column: string,
    query?: string,
    offset?: number
  ) => Promise<FilterValuePage>;
  summarize?: (sheet: string, range: string) => Promise<SelectionSummary>;
  readRange?: ReadRange;
  storageKey?: string;
  /** Stable private document identity; accepted revisions belong in storageKey. */
  documentId?: string;
  /** Mutable owner lifetime, distinct from an accepted revision or profile key. */
  ownerIdentity?: string;
  /** Accepted local/portable head, without pretending it is native storage. */
  acceptedCommandState?: string;
  /** Stable file/profile identity may replace an untitled identity on first Save. */
  viewPreferenceId?: string;
  namedRanges?: WorkbookView['namedRanges'];
  tables?: WorkbookView['tables'];
  pivots?: WorkbookView['pivots'];
  passiveNotice?: string;
  queryControls?: WorkbookQueryControls;
  calculationControls?: WorkbookCalculationControls;
  usedRangeControls?: WorkbookUsedRangeControls;
  /** Shared source inventory for future Home and Review note consumers. */
  noteMembershipControls?: WorkbookNoteMembershipControls;
  traceControls?: { receipt: FormulaTraceReceipt; trace: FormulaTraceFrame['trace'] };
  formulaPreviewControls?: FormulaPreviewControls;
  reviewControls?: WorkbookReviewControls;
  actionsControls?: WorkbookActionsControls;
  nameCommands?: boolean;
  /** Explicit native XLSX profile; strict captured command acceptance is required. */
  cellInsertionControls?: boolean;
  accessibilityControls?: WorkbookAccessibilityControls;
  notesControls?: WorkbookNotesControls;
  hyperlinkControls?: {
    receipt: WorkbookHyperlinkReceipt;
    read: (sheet: string, cell: string) => Promise<WorkbookHyperlinkDefinition>;
    follow?: (sheet: string, cell: string, revision: string) => Promise<HyperlinkFollowResult>;
  };
  objectControls?: Pick<WorkbookObjectsProps, 'documentId' | 'revision' | 'objects' | 'object' | 'onBatch' | 'pickPng'>;
  printControls?: Omit<WorkbookPrintProps, 'sheet' | 'selection' | 'readOnly' | 'busy' | 'words' | 'preview'> & {
    preview: (
      sheet: string,
      ...args: Parameters<WorkbookPrintProps['preview']>
    ) => ReturnType<WorkbookPrintProps['preview']>;
  };
  conditionalControls?: Omit<ConditionalRulesProps, 'sheet' | 'selection' | 'readOnly' | 'busy' | 'words'>;
  validationControls?: Omit<DataValidationProps, 'sheet' | 'selection' | 'readOnly' | 'busy' | 'words'>;
  importPaste?: (
    text: string,
    inspection: PasteInspection,
    options: { signal: AbortSignal; onProgress: (completed: number, total: number) => void }
  ) => Promise<void>;
  importFile?: (sheet: string, start: string) => Promise<void>;
  readOnly: boolean;
  literal?: boolean;
  structuralReason?: string;
  onUndo?: () => void;
  onCommandUndo?: () => Promise<void>;
  onCommandRedo?: () => Promise<void>;
  onRedo?: () => void;
  canUndo?: boolean;
  canRedo?: boolean;
  focus?: { sheet: string; cell: string };
  status?: string;
  onEdit: (edit: WorkbookEdit) => void;
  onBatch?: (commands: WorkbookCommand[]) => void;
  /** Registered writes settle only after the owner accepts or refuses them. */
  onCommandBatch?: (commands: WorkbookCommand[]) => Promise<void>;
  busy?: boolean;
}
