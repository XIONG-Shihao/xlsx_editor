import type { AccessibilityFrame } from '../accessibilityController';
import type { CalculationFrame, CalculationJobSnapshot, WorkbookCalculationControls } from '../calculationJobs';
import type { ChartRecommendationFrame } from '../chartRecommendationsController';
import type { FormulaTraceController, TraceDirection } from '../formulaTrace';
import type { FormulaTraceFrame, FormulaTraceReceipt } from '../formulaTraceAdapter';
import type { FunctionCategory } from '../functionCatalog';
import type {
  GridConditionalControls5,
  GridCurrentNamePicker,
  GridHomeInvoke2,
  GridHomeProps,
  GridImportPaste,
  GridInkReason2,
  GridLocalExecutionReason2,
  GridNameDraftReason,
  GridObjectControls5,
  GridPrintControls4,
  GridRunCalculationAction,
  GridScan,
  GridToggleTool,
} from './gridTypes';
import type { HomeBinding, HomeCommandId, HomeFrame } from '../homeCommands';
import type { IllustrationMode } from '../WorkbookShapes';
import type { InkPenId } from '../../core/ink-stroke.mjs';
import type { LocalExecutionAction } from '../../core/local-action-execution.mjs';
import type { NamedFormulaChoice } from '../namedFormulaDraft';
import { type NamedViewPreset, ViewPresetStore } from '../viewPresets';
import type { NoteFrame } from '../noteController';
import type { PageStylesFrame } from '../pageStylesController';
import type { PivotContextFrame } from '../pivotSelection';
import type { PrintFacet } from '../WorkbookPrint';
import type { ReadRange } from '../rangeCells';
import type {
  RibbonBinding,
  RibbonCommandId,
  RibbonDispatchScope,
  WorkbookTool,
  dispatchRibbonCommand,
} from '../ribbonCommands';
import { ViewPreferenceStore, type ViewPreferences, type ViewPreferencesPatch } from '../viewPreferences';
import type { WorkbookActionsControls } from '../WorkbookActions';
import type {
  WorkbookCell,
  WorkbookCommand,
  WorkbookGeometryProfile,
  WorkbookNormalFont,
  WorkbookPivotDefinition,
  WorkbookSheet,
} from '../../core/contracts';
import type { WorkbookQueryControls } from '../WorkbookQueries';
import type { WorkbookReviewControls } from '../reviewHostAdapter';
import type { WorkbookReviewMode } from '../workbookReviewController';
import type { WorkbookRibbonTool } from '../SpreadsheetRibbon';
import type { WorksheetDimensions } from '../../core/worksheet-geometry.mjs';
import type { WorksheetObjectFrame } from '../worksheetObjectController';

export interface GridRibbonInputs {
  accessibilityFrame: AccessibilityFrame | null;
  actionsControls: WorkbookActionsControls | undefined;
  actionsIntent: { serial: number; mode: 'new' | 'list' | 'run'; action?: LocalExecutionAction };
  active: boolean;
  activeFrame: (work: () => void) => number;
  activeRange: () => string;
  activeTool: WorkbookTool | null;
  autoSum: (binding: RibbonBinding) => Promise<void>;
  busy: boolean;
  calculateCurrentWorkbook: (binding: RibbonBinding) => Promise<void>;
  calculationBusy: boolean;
  calculationControls: WorkbookCalculationControls | undefined;
  calculationFrame: React.MutableRefObject<CalculationFrame | null>;
  calculationJob: CalculationJobSnapshot;
  calculationOwnerAvailable: boolean;
  canRedo: boolean | undefined;
  canUndo: boolean | undefined;
  changeView: (patch: ViewPreferencesPatch) => void;
  chartFrame: ChartRecommendationFrame | null;
  chartToolsMode: 'manage' | 'recommendations';
  chooseFormulaName: (choice: NamedFormulaChoice) => void;
  closeNamePicker: () => void;
  commandScope: RibbonDispatchScope;
  conditionalColor: string;
  conditionalControls: GridConditionalControls5;
  conditionalRule: 'greaterThan' | 'lessThan' | 'duplicateValues';
  conditionalThreshold: string;
  currentNamePicker: GridCurrentNamePicker;
  currentPivot: WorkbookPivotDefinition | null;
  drawMode: 'draw' | 'select' | 'erase' | null;
  enterDraw: (binding: RibbonBinding, mode?: 'draw' | 'select' | 'erase') => void;
  fillToolMode: 'series' | 'across';
  functionBrowse: { mode: 'all' | 'recent' | FunctionCategory; serial: number };
  geometryProfile: WorkbookGeometryProfile;
  goPivotSource: (key: string) => void;
  gridRoot: React.RefObject<HTMLDivElement>;
  homeDisabled: (id: HomeCommandId) => boolean;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homeInvoke: GridHomeInvoke2;
  homeMenu: (event: React.MouseEvent, menu: string) => void;
  homeProps: GridHomeProps;
  homeReason: (id: HomeCommandId) => string | undefined;
  homeRenderFrame: HomeFrame;
  homeSameTarget: (binding: HomeBinding) => boolean;
  illustrationFrame: WorksheetObjectFrame | null;
  illustrationMode: IllustrationMode;
  importFile: ((sheet: string, start: string) => Promise<void>) | undefined;
  importPaste: GridImportPaste;
  inkPreferenceStore: { readonly warning: string | null; read(): InkPenId; write(pen: InkPenId): InkPenId };
  inkPreferenceWarning: string | null;
  inkPreset: InkPenId;
  inkReason: GridInkReason2;
  literal: boolean | undefined;
  localExecutionAvailable: boolean;
  localExecutionReason: GridLocalExecutionReason2;
  mixedStyle: (key: keyof WorkbookCell) => boolean;
  nameCommands: boolean;
  nameDraftReason: GridNameDraftReason;
  namedViewPresets: NamedViewPreset[];
  navigationGeneration: React.MutableRefObject<number>;
  normalFont: Readonly<WorkbookNormalFont>;
  normalMode: React.MutableRefObject<boolean>;
  notesFrame: NoteFrame | null;
  objectControls: GridObjectControls5;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  onCommandResult: (_id: RibbonCommandId, result: Awaited<ReturnType<typeof dispatchRibbonCommand>>) => void;
  openAccessibility: (binding: RibbonBinding) => void;
  openActions: (binding: RibbonBinding, mode: 'new' | 'list') => void;
  openCharts: (binding: RibbonBinding, mode: 'manage' | 'recommendations') => void;
  openIllustrations: (binding: RibbonBinding, mode: IllustrationMode) => void;
  openLocalExecution: (binding: RibbonBinding, action: LocalExecutionAction) => void;
  openNamePicker: () => void;
  openNotes: (binding: RibbonBinding) => void;
  openPageStyles: (binding: RibbonBinding, mode: 'themes' | 'backgrounds') => void;
  openPrintFacet: (facet: PrintFacet, binding: RibbonBinding) => void;
  openQueries: (binding: RibbonBinding, mode: 'individual' | 'all') => void;
  openReview: (binding: RibbonBinding, mode: WorkbookReviewMode) => void;
  ownerBusy: boolean;
  ownerIdentity: string | undefined;
  pageStylesFrame: PageStylesFrame | null;
  pageStylesMode: 'themes' | 'backgrounds';
  paintStyleId: number | null;
  passiveNotice: string | undefined;
  pivotFrame: React.MutableRefObject<PivotContextFrame | null>;
  pivotReason: string | undefined;
  preferenceDocumentId: string | undefined;
  preferenceStore: ViewPreferenceStore;
  presetAction: (action: () => void) => void;
  presetBindingKey: string;
  presetDocumentId: string;
  presetStore: ViewPresetStore;
  printControls: GridPrintControls4;
  queriesMode: 'individual' | 'all';
  queryAllReason: 'This workbook is read-only.' | 'Return to Normal view to refresh queries.' | undefined;
  queryControls: WorkbookQueryControls | undefined;
  readOnly: boolean;
  readRange: ReadRange | undefined;
  renderedCalculationKey: string | null;
  renderedPivotKey: string | undefined;
  returnNormal: () => void;
  reviewControls: WorkbookReviewControls | undefined;
  reviewIntent: { serial: number; mode: WorkbookReviewMode };
  runCalculationAction: GridRunCalculationAction;
  runPivotRefresh: (key: string) => Promise<void>;
  scan: GridScan;
  scrollCell: (owner: WorkbookSheet, row: number, col: number, align?: boolean) => void;
  selected: string;
  sendCommand: (commands: WorkbookCommand[], allowSpecial?: boolean) => Promise<void>;
  setDrawMode: React.Dispatch<React.SetStateAction<'draw' | 'select' | 'erase' | null>>;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setInkPreferenceWarning: React.Dispatch<React.SetStateAction<string | null>>;
  setInkPreset: React.Dispatch<React.SetStateAction<InkPenId>>;
  setViewNotice: React.Dispatch<React.SetStateAction<string>>;
  sheet: WorkbookSheet;
  showPageBreak: (binding: RibbonBinding) => Promise<void>;
  showPageLayout: (binding: RibbonBinding) => Promise<void>;
  sortFilterCommands: JSX.Element | undefined;
  specialSelection: Set<string> | null;
  startTrace: (binding: RibbonBinding, direction: TraceDirection) => Promise<void>;
  structuralBlock: string;
  styleCell: WorkbookCell;
  tableName: string;
  taskPaneId: string;
  toggleTool: GridToggleTool;
  tools: WorkbookRibbonTool[];
  traceController: FormulaTraceController;
  traceControls: { receipt: FormulaTraceReceipt; trace: FormulaTraceFrame['trace'] } | undefined;
  traceFrame: React.MutableRefObject<FormulaTraceFrame | null>;
  traceInventoryReason: 'Formula tracing supports up to 1,024 worksheets in a workbook.' | undefined;
  view: ViewPreferences;
  viewMode: 'normal' | 'pageLayout' | 'pageBreak';
  viewport: { width: number; height: number; scrollTop: number; scrollLeft: number };
  visibleColumns: number;
  visibleRows: number;
  visibleTool: WorkbookRibbonTool | undefined;
  words: (en: string) => string;
  worksheetDimensions: WorksheetDimensions;
}
