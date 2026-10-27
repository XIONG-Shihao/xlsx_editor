import type { CalculationAction, CalculationJobSnapshot } from '../calculationJobs';
import type { CalculationGroupReceipt } from '../../adapters/workbook-calculation-group.mjs';
import type { ConditionalRulesProps, DataValidationProps } from '../WorkbookRules';
import type { FilterContext } from '../filterSelection';
import type { FilterValuePage } from './filterValues';
import type { FormulaDraftHandle, FormulaDraftSnapshot, NamedFormulaChoice } from '../namedFormulaDraft';
import type { FunctionCategory } from '../functionCatalog';
import type {
  HomeArgumentsById,
  HomeBinding,
  HomeCommandId,
  HomeCommandLease,
  NoArgumentHomeId,
} from '../homeCommands';
import type { HyperlinkFollowResult } from '../hyperlinkController';
import type { LocalExecutionAction } from '../../core/local-action-execution.mjs';
import type { PasteInspection } from '../largePaste';
import type { QueryRefreshAllControls } from '../queryRefreshAllController';
import type { RangeWindow, ReadRange } from '../rangeCells';
import type { ReviewCellPage, ReviewCellRead, ReviewOwnerReceipt } from '../reviewScanAdapter';
import type { RibbonBinding, WorkbookTool } from '../ribbonCommands';
import type {
  WorkbookCell,
  WorkbookCommand,
  WorkbookConditionalDefinition,
  WorkbookDuplicatePreview,
  WorkbookEdit,
  WorkbookHyperlinkDefinition,
  WorkbookHyperlinkReceipt,
  WorkbookSheet,
} from '../../core/contracts';
import type { WorkbookObjectsProps } from '../WorkbookObjects';
import type { WorkbookPrintProps } from '../WorkbookPrint';

export type GridCalculationController = {
  subscribe: (listener: () => void) => () => void;
  getSnapshot: () => CalculationJobSnapshot;
  visible: () => CalculationJobSnapshot;
  run(actionInput: CalculationAction, expectedKey: string): Promise<CalculationGroupReceipt>;
  cancel(expectedRequestId: string): Promise<void>;
  dispose(): void;
};

export type GridConditionalControls =
  | Omit<ConditionalRulesProps, 'sheet' | 'readOnly' | 'busy' | 'selection' | 'words'>
  | undefined;

export type GridHyperlinkControls =
  | {
      receipt: WorkbookHyperlinkReceipt;
      read: (sheet: string, cell: string) => Promise<WorkbookHyperlinkDefinition>;
      follow?: (sheet: string, cell: string, revision: string) => Promise<HyperlinkFollowResult>;
    }
  | undefined;

export type GridLocalExecutionFrame = {
  key: string;
  ownerKey: string;
  sheet: string;
  selection: string;
  active: boolean;
  reason: string | undefined;
  apply: (commands: WorkbookCommand[]) => Promise<void>;
} | null;

export type GridObjectControls =
  | Pick<WorkbookObjectsProps, 'object' | 'objects' | 'documentId' | 'revision' | 'onBatch' | 'pickPng'>
  | undefined;

export type GridPageTheme =
  | { revision: string | null; sample: string | null; canApply: boolean; protectedReason: string | null }
  | undefined;

export type GridPrintControls =
  | (Omit<WorkbookPrintProps, 'sheet' | 'readOnly' | 'busy' | 'selection' | 'words' | 'preview'> & {
      preview: (
        sheet: string,
        ...args: Parameters<WorkbookPrintProps['preview']>
      ) => ReturnType<WorkbookPrintProps['preview']>;
    })
  | undefined;

export type GridQueryAllFrame = {
  key: string;
  ownerKey: string;
  name: string;
  active: boolean;
  busy: boolean;
  reason: string | undefined;
  controls: QueryRefreshAllControls;
} | null;

export type GridScan =
  | ((
      sheet: string,
      after?: string,
      limit?: number
    ) => Promise<{ cells: Record<string, WorkbookCell>; nextCursor: string | null }>)
  | undefined;

export type GridSetAcceptedConditional = React.Dispatch<
  React.SetStateAction<{ scope: string; definition: WorkbookConditionalDefinition } | null>
>;

export type GridTables = {
  name: string;
  sheet: string;
  ref: string;
  columns: Array<{ name: string }>;
  headerRowCount?: number;
  totalsRowCount?: number;
}[];

export type GridToggleTool = (
  tool: WorkbookTool,
  initiatingBinding?: RibbonBinding,
  browse?: 'all' | 'recent' | FunctionCategory
) => void;

export type GridValidationControls =
  | Omit<DataValidationProps, 'sheet' | 'readOnly' | 'busy' | 'selection' | 'words'>
  | undefined;

export type GridFillDragging = React.MutableRefObject<{
  source: string;
  target: string | null;
  dispatch?: (range: string) => Promise<void>;
} | null>;

export type GridHomeInvoke = <K extends HomeCommandId>(
  id: K,
  args?: HomeArgumentsById[K],
  work?: (lease: HomeCommandLease, args: HomeArgumentsById[K]) => void | Promise<void>
) => Promise<void>;

export type GridInkLive = React.MutableRefObject<{
  key: string;
  active: false | 'draw' | 'select' | 'erase' | null;
  reason: string | undefined;
}>;

export type GridInkReason =
  | 'Drawing is unavailable in this host.'
  | 'Activate this workbook to draw.'
  | 'Return to Normal view to draw.'
  | 'This workbook is read-only.'
  | 'Wait for the workbook operation to finish.'
  | 'Unfreeze panes before drawing.'
  | 'Click a regular cell before drawing.'
  | undefined;

export type GridObjectControls2 =
  | Pick<WorkbookObjectsProps, 'object' | 'revision' | 'documentId' | 'objects' | 'onBatch' | 'pickPng'>
  | undefined;

export type GridSetViewport = React.Dispatch<
  React.SetStateAction<{ width: number; height: number; scrollTop: number; scrollLeft: number }>
>;

export type GridHomeProps = (id: HomeCommandId) => {
  'data-home-command':
    | NoArgumentHomeId
    | 'home.paste'
    | 'home.paste-text'
    | 'home.paint-format'
    | 'home.font-family'
    | 'home.font-size'
    | 'home.bold'
    | 'home.italic'
    | 'home.underline'
    | 'home.borders'
    | 'home.fill-color'
    | 'home.font-color'
    | 'home.wrap'
    | 'home.number-format'
    | 'home.conditional-draft'
    | 'home.conditional-apply'
    | 'home.table-draft'
    | 'home.create-table'
    | 'home.column-width'
    | 'home.row-height'
    | 'home.fill-range'
    | 'home.fill-across-preview'
    | 'home.fill-across-apply'
    | 'home.dimension-draft'
    | 'home.find-draft'
    | 'home.find-page'
    | 'home.find-match'
    | 'home.menu';
  disabled: boolean;
  'aria-description': string | undefined;
  title: string | undefined;
};

export type GridConditionalControls2 =
  | Omit<ConditionalRulesProps, 'sheet' | 'selection' | 'readOnly' | 'busy' | 'words'>
  | undefined;

export type GridImportPaste =
  | ((
      text: string,
      inspection: PasteInspection,
      options: { signal: AbortSignal; onProgress: (completed: number, total: number) => void }
    ) => Promise<void>)
  | undefined;

export type GridNamePicker = {
  key: string;
  target: string;
  snapshot: FormulaDraftSnapshot;
  choices: readonly NamedFormulaChoice[];
  handle: FormulaDraftHandle;
} | null;

export type GridSetNamePicker = React.Dispatch<
  React.SetStateAction<{
    key: string;
    target: string;
    snapshot: FormulaDraftSnapshot;
    choices: readonly NamedFormulaChoice[];
    handle: FormulaDraftHandle;
  } | null>
>;

export type GridFilterBounds = {
  first: { row: number; col: number };
  last: { row: number; col: number };
  rows: number;
  columns: number;
} | null;

export type GridFilterValues =
  | ((sheet: string, range: string, column: string, query?: string, offset?: number) => Promise<FilterValuePage>)
  | undefined;

export type GridInvokeFilter = (
  builder: (context: FilterContext) => WorkbookCommand,
  notice: string,
  requireRectangle?: boolean
) => Promise<void>;

export type GridSetFilterSelected = React.Dispatch<
  React.SetStateAction<Map<string, { key: string; value: string | number | boolean | null; count: number }>>
>;

export type GridSortBounds = {
  first: { row: number; col: number };
  last: { row: number; col: number };
  rows: number;
  columns: number;
} | null;

export type GridActionsParkedDraft = React.MutableRefObject<{
  key: string;
  handle: FormulaDraftHandle;
  snapshot: FormulaDraftSnapshot;
} | null>;

export type GridConditionalControls3 =
  | Omit<ConditionalRulesProps, 'sheet' | 'selection' | 'busy' | 'readOnly' | 'words'>
  | undefined;

export type GridCreateNamesParkedDraft = React.MutableRefObject<{
  key: string;
  handle: FormulaDraftHandle;
  snapshot: FormulaDraftSnapshot;
} | null>;

export type GridDuplicateParkedDraft = React.MutableRefObject<{
  key: string;
  handle: FormulaDraftHandle;
  snapshot: FormulaDraftSnapshot;
} | null>;

export type GridLiveFunctionFrame = React.MutableRefObject<{
  reader: ReadRange | undefined;
  sheet: WorkbookSheet;
  ownerIdentity: string;
  grantDocumentId: string;
}>;

export type GridLiveNameInventory = React.MutableRefObject<{
  namedRanges: { name: string; formula: string; sheet?: string; scopeInvalid?: true }[];
  tables: {
    name: string;
    sheet: string;
    ref: string;
    columns: Array<{ name: string }>;
    headerRowCount?: number;
    totalsRowCount?: number;
  }[];
}>;

export type GridLocalExecutionReason =
  | 'This workbook is read-only.'
  | 'Local action execution is unavailable in this host.'
  | 'Return to Normal view to run actions.'
  | 'Wait for the workbook operation to finish.'
  | 'Click a regular cell before running a local action.'
  | undefined;

export type GridReplacementParkedDraft = React.MutableRefObject<{
  key: string;
  handle: FormulaDraftHandle;
  snapshot: FormulaDraftSnapshot;
} | null>;

export type GridReviewParkedDraft = React.MutableRefObject<{
  key: string;
  handle: FormulaDraftHandle;
  snapshot: FormulaDraftSnapshot;
} | null>;

export type GridSeriesParkedDraft = React.MutableRefObject<{
  key: string;
  handle: FormulaDraftHandle;
  snapshot: FormulaDraftSnapshot;
} | null>;

export type GridSetActionsIntent = React.Dispatch<
  React.SetStateAction<{ serial: number; mode: 'new' | 'list' | 'run'; action?: LocalExecutionAction }>
>;

export type GridSetFunctionBrowse = React.Dispatch<
  React.SetStateAction<{ mode: 'all' | 'recent' | FunctionCategory; serial: number }>
>;

export type GridSortParkedDraft = React.MutableRefObject<{
  key: string;
  handle: FormulaDraftHandle;
  snapshot: FormulaDraftSnapshot;
} | null>;

export type GridSpecialParkedDraft = React.MutableRefObject<{
  key: string;
  handle: FormulaDraftHandle;
  snapshot: FormulaDraftSnapshot;
} | null>;

export type GridTextColumnsParkedDraft = React.MutableRefObject<{
  key: string;
  handle: FormulaDraftHandle;
  snapshot: FormulaDraftSnapshot;
} | null>;

export type GridValidationControls2 =
  | Omit<DataValidationProps, 'sheet' | 'selection' | 'busy' | 'readOnly' | 'words'>
  | undefined;

export type GridObjectControls3 =
  | Pick<WorkbookObjectsProps, 'object' | 'revision' | 'objects' | 'documentId' | 'onBatch' | 'pickPng'>
  | undefined;

export type GridConditionalControls4 =
  | Omit<ConditionalRulesProps, 'readOnly' | 'busy' | 'sheet' | 'selection' | 'words'>
  | undefined;

export type GridObjectControls4 =
  | Pick<WorkbookObjectsProps, 'object' | 'revision' | 'documentId' | 'onBatch' | 'objects' | 'pickPng'>
  | undefined;

export type GridPrintControls2 =
  | (Omit<WorkbookPrintProps, 'preview' | 'readOnly' | 'busy' | 'sheet' | 'selection' | 'words'> & {
      preview: (
        sheet: string,
        ...args: Parameters<WorkbookPrintProps['preview']>
      ) => ReturnType<WorkbookPrintProps['preview']>;
    })
  | undefined;

export type GridReviewReader = {
  scan(
    receipt: ReviewOwnerReceipt,
    identity: string,
    sheet: string,
    after: string | undefined,
    limit: number,
    signal: AbortSignal
  ): Promise<ReviewCellPage>;
  read(
    receipt: ReviewOwnerReceipt,
    identity: string,
    sheet: string,
    cell: string,
    signal: AbortSignal
  ): Promise<ReviewCellRead>;
};

export type GridValidationControls3 =
  | Omit<DataValidationProps, 'readOnly' | 'busy' | 'sheet' | 'selection' | 'words'>
  | undefined;

export type GridPrintControls3 =
  | (Omit<WorkbookPrintProps, 'sheet' | 'selection' | 'readOnly' | 'busy' | 'words' | 'preview'> & {
      preview: (
        sheet: string,
        ...args: Parameters<WorkbookPrintProps['preview']>
      ) => ReturnType<WorkbookPrintProps['preview']>;
    })
  | undefined;

export type GridValidationControls4 =
  | Omit<DataValidationProps, 'sheet' | 'selection' | 'readOnly' | 'busy' | 'words'>
  | undefined;

export type GridHomeReadCells = (
  lease: HomeCommandLease,
  name: string,
  target: string,
  signal?: AbortSignal
) => Promise<Record<string, WorkbookCell>>;

export type GridDuplicatePreview =
  | ((
      sheet: string,
      range: string,
      hasHeaders: boolean,
      columns: string[],
      offset: number
    ) => Promise<WorkbookDuplicatePreview>)
  | undefined;

export type GridPaintStyleRef = React.MutableRefObject<Pick<
  WorkbookEdit,
  | 'bold'
  | 'italic'
  | 'underline'
  | 'fontFamily'
  | 'fontSize'
  | 'fontColor'
  | 'fillColor'
  | 'horizontalAlign'
  | 'verticalAlign'
  | 'wrapText'
  | 'border'
  | 'styleId'
  | 'numberFormat'
> | null>;

export type GridRanges = {
  sheets: WorkbookSheet[];
  request: (windows: RangeWindow[]) => string;
  load: (key: string) => void;
  error: string;
  ready: (key: string) => boolean;
};

export type GridInternalClipboard = React.MutableRefObject<{
  sheet: string;
  range: string;
  text: string;
  cut: boolean;
  binding?: HomeBinding;
} | null>;

export type GridSetCalculationNotice = React.Dispatch<
  React.SetStateAction<{ key: string; phase: 'running' | 'completed' | 'failed'; text: string } | null>
>;

export type GridHomePaintRunner = React.MutableRefObject<
  ((cell: string, format: Extract<WorkbookCommand, { type: 'formatRange' }>['format']) => void) | null
>;

export type GridRunCalculationAction = (
  binding: RibbonBinding,
  action: CalculationAction,
  expectedKey: string | null
) => Promise<void>;

export type GridFindConfig = React.MutableRefObject<{
  text: string;
  mode: 'display' | 'raw' | 'formula';
  sheets: string[];
} | null>;

export type GridHomeFormatCommands = (
  lease: HomeCommandLease,
  format: Extract<WorkbookCommand, { type: 'formatRange' }>['format']
) => WorkbookCommand[];

export type GridConditionalControls5 =
  | Omit<ConditionalRulesProps, 'busy' | 'sheet' | 'selection' | 'readOnly' | 'words'>
  | undefined;

export type GridHomeInvoke2 = <K extends HomeCommandId>(
  id: K,
  args?: HomeArgumentsById[K],
  work?: (lease: HomeCommandLease, args: HomeArgumentsById[K]) => void | Promise<void>
) => Promise<void>;

export type GridInkLive2 = React.MutableRefObject<{
  key: string;
  active: false | 'select' | 'draw' | 'erase' | null;
  reason: string | undefined;
}>;

export type GridObjectControls5 =
  | Pick<WorkbookObjectsProps, 'object' | 'documentId' | 'revision' | 'onBatch' | 'objects' | 'pickPng'>
  | undefined;

export type GridPrintControls4 =
  | (Omit<WorkbookPrintProps, 'busy' | 'sheet' | 'selection' | 'readOnly' | 'words' | 'preview'> & {
      preview: (
        sheet: string,
        ...args: Parameters<WorkbookPrintProps['preview']>
      ) => ReturnType<WorkbookPrintProps['preview']>;
    })
  | undefined;

export type GridValidationControls5 =
  | Omit<DataValidationProps, 'busy' | 'sheet' | 'selection' | 'readOnly' | 'words'>
  | undefined;

export type GridCurrentNamePicker = {
  key: string;
  target: string;
  snapshot: FormulaDraftSnapshot;
  choices: readonly NamedFormulaChoice[];
  handle: FormulaDraftHandle;
} | null;

export type GridInkReason2 =
  | 'This workbook is read-only.'
  | 'Wait for the workbook operation to finish.'
  | 'Drawing is unavailable in this host.'
  | 'Activate this workbook to draw.'
  | 'Return to Normal view to draw.'
  | 'Unfreeze panes before drawing.'
  | 'Click a regular cell before drawing.'
  | undefined;

export type GridLocalExecutionReason2 =
  | 'Local action execution is unavailable in this host.'
  | 'Return to Normal view to run actions.'
  | 'This workbook is read-only.'
  | 'Wait for the workbook operation to finish.'
  | 'Click a regular cell before running a local action.'
  | undefined;

export type GridNameDraftReason =
  | 'Name insertion is unavailable in this editor.'
  | 'Show the formula bar to insert a name.'
  | 'Finish composing the formula first.'
  | 'Click a regular cell to insert a name.'
  | undefined;
