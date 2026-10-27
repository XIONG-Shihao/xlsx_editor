import type { AcceptedPrintPreview } from '../useAcceptedPrintPreview';
import type { AccessibilityFrame } from '../accessibilityController';
import type { BackgroundTile } from '../backgroundTiles';
import type { CalculationFrame, WorkbookCalculationControls } from '../calculationJobs';
import type { CellInsertionFrame } from '../cellInsertionController';
import type { ChartRecommendationFrame } from '../chartRecommendationsController';
import type { CreateNamesContext } from '../createNamesFromSelection';
import type { DuplicateContext } from '../duplicateSelection';
import type { ErrorScanScope, ErrorScanner } from '../errorChecking';
import type { FillAcrossContext } from '../WorkbookFillAcross';
import type { FillSeriesContext } from '../WorkbookFill';
import type { FormulaDraftHandle, FormulaDraftSnapshot } from '../namedFormulaDraft';
import type { FormulaDraftPreviewController, FormulaPreviewFrame } from '../formulaDraftPreview';
import type { FormulaTraceController, FormulaTraceScope, TraceEndpoint } from '../formulaTrace';
import type { FormulaTraceFrame, FormulaTraceReceipt } from '../formulaTraceAdapter';
import type { FunctionCategory } from '../functionCatalog';
import type { FunctionInsertionContext } from '../functionInsertion';
import type {
  GridCalculationController,
  GridConditionalControls5,
  GridFillDragging,
  GridHomeInvoke2,
  GridHomeProps,
  GridHyperlinkControls,
  GridInkLive2,
  GridInkReason,
  GridLocalExecutionFrame,
  GridObjectControls5,
  GridPageTheme,
  GridPrintControls4,
  GridQueryAllFrame,
  GridScan,
  GridSetAcceptedConditional,
  GridSetViewport,
  GridTables,
  GridToggleTool,
  GridValidationControls5,
} from './gridTypes';
import type { HomeCommandId, HomeFrame } from '../homeCommands';
import type { HyperlinkFrame } from '../hyperlinkController';
import type { IllustrationMode } from '../WorkbookShapes';
import type { InkPenId } from '../../core/ink-stroke.mjs';
import type { LocalExecutionAction } from '../../core/local-action-execution.mjs';
import type { NoteFrame } from '../noteController';
import type { ObjectAxes } from '../worksheetObjectGeometry';
import type { PageBoundaryModel } from '../pageBoundaries';
import type { PageStylesFrame } from '../pageStylesController';
import type { PivotContextFrame, PivotFieldDraft } from '../pivotSelection';
import type { PivotLayoutDraft } from '../pivotLayout';
import type { PrintFocusRequest } from '../WorkbookPrint';
import type { ReadRange } from '../rangeCells';
import type { ReplaceContext } from '../WorkbookReplace';
import type {
  RibbonBinding,
  RibbonCommandHandler,
  RibbonCommandId,
  RibbonDispatchScope,
  WorkbookTool,
} from '../ribbonCommands';
import type { SearchCursor } from '../searchPages';
import type { SelectionSummary } from '../../core/indexed-session.mjs';
import type { SortContext } from '../sortSelection';
import type { SpecialSelectionContext } from '../WorkbookSpecialSelection';
import type { TextColumnContext } from '../textColumns';
import type { ViewGeometry, ViewVirtualWindow } from '../viewGeometry';
import type { ViewPreferences } from '../viewPreferences';
import type { WatchReader, WatchScope } from '../watchWindow';
import type {
  WorkbookAccessibilityControls,
  WorkbookCell,
  WorkbookCommand,
  WorkbookEdit,
  WorkbookGeometryProfile,
  WorkbookNormalFont,
  WorkbookNotesControls,
  WorkbookPivotDefinition,
  WorkbookSheet,
  WorkbookValidationRule,
} from '../../core/contracts';
import type { WorkbookActionsControls } from '../WorkbookActions';
import type { WorkbookQueryControls } from '../WorkbookQueries';
import type { WorkbookReviewFrame, WorkbookReviewMode } from '../workbookReviewController';
import type { WorkbookRibbonTool } from '../SpreadsheetRibbon';
import type { WorksheetObjectFrame } from '../worksheetObjectController';

export interface GridWorkspaceInputs {
  accessibilityControls: WorkbookAccessibilityControls | undefined;
  accessibilityFrame: AccessibilityFrame | null;
  acrossContext: FillAcrossContext | null;
  actionsControls: WorkbookActionsControls | undefined;
  actionsFrame: { key: string; sheet: string; controls: WorkbookActionsControls | undefined };
  actionsIntent: { serial: number; mode: 'new' | 'list' | 'run'; action?: LocalExecutionAction };
  activationEpoch: React.MutableRefObject<number>;
  active: boolean;
  activeFrame: (work: () => void) => number;
  activeRange: () => string;
  activeTool: WorkbookTool | null;
  boundaryModel: PageBoundaryModel | null;
  busy: boolean;
  calculationController: GridCalculationController;
  calculationControls: WorkbookCalculationControls | undefined;
  calculationFrame: React.MutableRefObject<CalculationFrame | null>;
  calculationKey: string;
  calculationNotice: { key: string; phase: 'running' | 'completed' | 'failed'; text: string } | null;
  cancelCellFocus: () => void;
  canEditRef: React.MutableRefObject<boolean>;
  cellInsertionControls: boolean | undefined;
  cellInsertionFrame: CellInsertionFrame | null;
  cellInsertionMode: 'checkbox' | 'symbols';
  chartFrame: ChartRecommendationFrame | null;
  chartToolsMode: 'manage' | 'recommendations';
  closeTaskPane: () => void;
  columnOffsets: Float64Array<ArrayBufferLike>;
  columnWindow: ViewVirtualWindow;
  commandScope: RibbonDispatchScope;
  conditionalControls: GridConditionalControls5;
  conditionalScope: string;
  createNamesContext: CreateNamesContext | null;
  currentPivot: WorkbookPivotDefinition | null;
  documentId: string | undefined;
  dragEnter: (cell: string) => void;
  dragStart: (cell: string, extend: boolean) => void;
  drawMode: 'select' | 'draw' | 'erase' | null;
  duplicateContext: DuplicateContext | null;
  editRef: React.MutableRefObject<(edit: WorkbookEdit) => void>;
  errorScanner: ErrorScanner;
  errorScope: ErrorScanScope;
  fillDragging: GridFillDragging;
  fillPreview: string | null;
  fillToolMode: 'series' | 'across';
  find: () => Promise<void>;
  findBusy: boolean;
  findIndex: number;
  findMatches: { sheet: string; cell: string }[];
  findMode: 'formula' | 'display' | 'raw';
  findNextCursor: SearchCursor | null;
  findOpen: boolean;
  findPageIndex: number;
  findPages: React.MutableRefObject<{ cursor: SearchCursor; offset: number }[]>;
  findScope: 'sheet' | 'workbook';
  findText: string;
  findTotal: number;
  firstRenderedColumn: number;
  firstRenderedRow: number;
  formulaDraftHandle: React.MutableRefObject<FormulaDraftHandle | null>;
  formulaPreviewController: FormulaDraftPreviewController;
  formulaPreviewDraft: { scopeKey: string; snapshot: FormulaDraftSnapshot; active: boolean } | null;
  formulaPreviewFrame: React.MutableRefObject<FormulaPreviewFrame | null>;
  formulaPreviewKey: string | null;
  frozenColumnCount: number;
  frozenColumnGap: number;
  frozenRowCount: number;
  frozenRowGap: number;
  functionBrowse: { mode: 'all' | 'recent' | FunctionCategory; serial: number };
  functionContext: FunctionInsertionContext | null;
  functionViewMode: React.MutableRefObject<{ mode: 'normal' | 'pageLayout' | 'pageBreak'; epoch: number }>;
  geometry: ViewGeometry;
  geometryKey: string;
  geometryProfile: WorkbookGeometryProfile;
  goPivotSource: (key: string) => void;
  goToName: (formula: string) => void;
  gridError: string;
  gridRoot: React.RefObject<HTMLDivElement>;
  gridScroll: React.RefObject<HTMLDivElement>;
  homeDisabled: (id: HomeCommandId) => boolean;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homeInvoke: GridHomeInvoke2;
  homeProps: GridHomeProps;
  homeRenderFrame: HomeFrame;
  hyperlinkControls: GridHyperlinkControls;
  hyperlinkFrame: HyperlinkFrame | null;
  illustrationFrame: WorksheetObjectFrame | null;
  illustrationMode: IllustrationMode;
  inkKey: string;
  inkLive: GridInkLive2;
  inkPreset: InkPenId;
  inkReason: GridInkReason;
  invoke: (id: RibbonCommandId, handler: RibbonCommandHandler) => void;
  jump: string;
  jumpError: string;
  jumpToCell: () => void;
  lastRenderedColumn: number;
  lastRenderedRow: number;
  layoutReason: string | undefined;
  literal: boolean | undefined;
  liveConditionalScope: React.MutableRefObject<string>;
  liveValidationScope: React.MutableRefObject<string>;
  localExecutionAvailable: boolean;
  localExecutionFrame: GridLocalExecutionFrame;
  moveAfterCommit: (from: string, key: 'Enter' | 'Tab', shift: boolean) => void;
  nameCommands: boolean;
  namedDraftKey: string;
  namedRanges: { name: string; formula: string; sheet?: string; scopeInvalid?: true }[];
  namePage: number;
  nameScope: 'workbook' | 'worksheet';
  navigateMetadataRange: (name: string, target: string) => void;
  navigateTrace: (target: TraceEndpoint) => void;
  navigationGeneration: React.MutableRefObject<number>;
  newRangeName: string;
  newSheetName: string;
  normalFont: Readonly<WorkbookNormalFont>;
  normalMode: React.MutableRefObject<boolean>;
  notesControls: WorkbookNotesControls | undefined;
  notesFrame: NoteFrame | null;
  objectAxes: ObjectAxes;
  objectControls: GridObjectControls5;
  objectFamily: string;
  objectFamilyEpoch: React.MutableRefObject<{ family: string; epoch: number }>;
  onBatch: ((commands: WorkbookCommand[]) => void) | undefined;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  openCharts: (binding: RibbonBinding, mode: 'manage' | 'recommendations') => void;
  openNamePicker: () => void;
  openTracePane: () => void;
  ownerBusy: boolean;
  ownerIdentity: string | undefined;
  pageStylesFrame: PageStylesFrame | null;
  pageStylesMode: 'themes' | 'backgrounds';
  pageTheme: GridPageTheme;
  passiveNotice: string | undefined;
  pivotFrame: React.MutableRefObject<PivotContextFrame | null>;
  pivotNotice: { key: string; text: string } | null;
  pivotReason: string | undefined;
  pivotRunning: boolean;
  pivots: WorkbookPivotDefinition[];
  printControls: GridPrintControls4;
  printFocus: PrintFocusRequest | null;
  printPreview: AcceptedPrintPreview;
  queriesAllIntent: number;
  queriesMode: 'individual' | 'all';
  queryAllFrame: GridQueryAllFrame;
  queryControls: WorkbookQueryControls | undefined;
  range: string;
  readOnly: boolean;
  readRange: ReadRange | undefined;
  renderedColumnIndices: number[];
  renderedPivotKey: string | undefined;
  renderedRows: number[];
  replacementContext: ReplaceContext | null;
  replacementDisabledReason: (action?: 'preview' | 'apply') => string | undefined;
  replaceText: string;
  reportFormulaPreviewDraft: (scopeKey: string, snapshot: FormulaDraftSnapshot, active: boolean) => void;
  returnNormal: () => void;
  reviewFrame: WorkbookReviewFrame | null;
  reviewIntent: { serial: number; mode: WorkbookReviewMode };
  rowOffsets: Float64Array<ArrayBufferLike>;
  rowWindow: ViewVirtualWindow;
  runPivotLayout: (key: string, draft: PivotLayoutDraft) => Promise<void>;
  runPivotRefresh: (key: string) => Promise<void>;
  runPivotUpdate: (key: string, draft: PivotFieldDraft) => Promise<void>;
  scan: GridScan;
  scrollCell: (owner: WorkbookSheet, row: number, col: number, align?: boolean) => void;
  select: (cell: string, extend?: boolean) => void;
  selectColumn: (col: number) => void;
  selected: string;
  selectedCell: WorkbookCell;
  selectedFontProfile: WorkbookGeometryProfile;
  selectedRef: React.MutableRefObject<string>;
  selectedValidation: WorkbookValidationRule | undefined;
  selectionSummary: SelectionSummary | null;
  selectRow: (row: number) => void;
  send: (commands: WorkbookCommand[], allowSpecial?: boolean) => void;
  seriesContext: FillSeriesContext | null;
  setAcceptedConditional: GridSetAcceptedConditional;
  setFormulaComposing: React.Dispatch<React.SetStateAction<boolean>>;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setJump: React.Dispatch<React.SetStateAction<string>>;
  setNamePage: React.Dispatch<React.SetStateAction<number>>;
  setNameScope: React.Dispatch<React.SetStateAction<'workbook' | 'worksheet'>>;
  setNewRangeName: React.Dispatch<React.SetStateAction<string>>;
  setNewSheetName: React.Dispatch<React.SetStateAction<string>>;
  setRange: React.Dispatch<React.SetStateAction<string>>;
  setSelectionEnd: React.Dispatch<React.SetStateAction<string>>;
  setSheetAction: React.Dispatch<React.SetStateAction<'addSheet' | 'renameSheet' | 'duplicateSheet' | null>>;
  setSheetName: React.Dispatch<React.SetStateAction<string>>;
  setSortHeader: React.Dispatch<React.SetStateAction<boolean>>;
  setTablePage: React.Dispatch<React.SetStateAction<number>>;
  setValue: (cell: string, value: string) => void;
  setViewNotice: React.Dispatch<React.SetStateAction<string>>;
  setViewport: GridSetViewport;
  sheet: WorkbookSheet;
  sheetAction: 'addSheet' | 'renameSheet' | 'duplicateSheet' | null;
  sheets: WorkbookSheet[];
  sortContext: SortContext | null;
  specialContext: SpecialSelectionContext | null;
  specialSelection: Set<string> | null;
  specialSelectionMode: React.MutableRefObject<{ selection: Set<string> | null }>;
  status: string | undefined;
  storageKey: string | undefined;
  submitSheetAction: () => void;
  summaryError: string;
  tableColumnCount: number;
  tablePage: number;
  tables: GridTables;
  taskPaneHeading: React.RefObject<HTMLHeadingElement>;
  taskPaneId: string;
  taskPaneTitleId: string;
  textColumnContext: TextColumnContext | null;
  toggleTool: GridToggleTool;
  toolFocusGeneration: React.MutableRefObject<number>;
  traceController: FormulaTraceController;
  traceControls: { receipt: FormulaTraceReceipt; trace: FormulaTraceFrame['trace'] } | undefined;
  traceInventoryReason: 'Formula tracing supports up to 1,024 worksheets in a workbook.' | undefined;
  traceScope: FormulaTraceScope;
  updateObjectExtent: (family: string, extent: { rows: number; columns: number }) => void;
  validationControls: GridValidationControls5;
  validationScope: string;
  view: ViewPreferences;
  viewActive: React.MutableRefObject<boolean>;
  viewMode: 'normal' | 'pageLayout' | 'pageBreak';
  viewNotice: string;
  viewport: { width: number; height: number; scrollTop: number; scrollLeft: number };
  visibleColumns: number;
  visibleRows: number;
  visibleTool: WorkbookRibbonTool | undefined;
  visitedTools: WorkbookTool[];
  watchReader: WatchReader;
  watchScope: WatchScope;
  watchSelectionReason: string | undefined;
  words: (en: string) => string;
  worksheetBackground: BackgroundTile | undefined;
}
