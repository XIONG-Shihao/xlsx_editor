import type { AcceptedPrintPreview } from '../useAcceptedPrintPreview';
import type { AccessibilityFrame } from '../accessibilityController';
import type { CalculationFrame } from '../calculationJobs';
import type { CellInsertionFrame } from '../cellInsertionController';
import type { ChartRecommendationFrame } from '../chartRecommendationsController';
import { ConditionalRules, DataValidation } from '../WorkbookRules';
import type { CreateNamesContext } from '../createNamesFromSelection';
import type { DuplicateContext } from '../duplicateSelection';
import type { ErrorScanScope, ErrorScanner } from '../errorChecking';
import { type FillAcrossContext, WorkbookFillAcross } from '../WorkbookFillAcross';
import { type FillSeriesContext, WorkbookFill } from '../WorkbookFill';
import type { FormulaDraftPreviewController, FormulaPreviewFrame } from '../formulaDraftPreview';
import type { FormulaTraceController, FormulaTraceScope, TraceEndpoint } from '../formulaTrace';
import type { FormulaTraceFrame, FormulaTraceReceipt } from '../formulaTraceAdapter';
import type { FunctionCategory } from '../functionCatalog';
import type { FunctionInsertionContext } from '../functionInsertion';
import type {
  GridCalculationController,
  GridConditionalControls,
  GridHyperlinkControls,
  GridLocalExecutionFrame,
  GridObjectControls,
  GridPageTheme,
  GridPrintControls,
  GridQueryAllFrame,
  GridScan,
  GridSetAcceptedConditional,
  GridTables,
  GridToggleTool,
  GridValidationControls,
} from './gridTypes';
import type { HyperlinkFrame } from '../hyperlinkController';
import type { IllustrationMode } from '../WorkbookShapes';
import type { LocalExecutionAction } from '../../core/local-action-execution.mjs';
import type { NoteFrame } from '../noteController';
import type { PageStylesFrame } from '../pageStylesController';
import type { PivotContextFrame, PivotFieldDraft } from '../pivotSelection';
import type { PivotLayoutDraft } from '../pivotLayout';
import { type PrintFocusRequest, WorkbookPrint } from '../WorkbookPrint';
import type { ReadRange } from '../rangeCells';
import { type ReplaceContext, WorkbookReplace } from '../WorkbookReplace';
import { RetainedWorkbookTool } from './RetainedWorkbookTool';
import type { RibbonBinding, RibbonDispatchScope } from '../ribbonCommands';
import type { SortContext } from '../sortSelection';
import { type SpecialSelectionContext, WorkbookSpecialSelection } from '../WorkbookSpecialSelection';
import type { TextColumnContext } from '../textColumns';
import type { WatchReader, WatchScope } from '../watchWindow';
import { WorkbookAccessibility } from '../WorkbookAccessibility';
import type {
  WorkbookAccessibilityControls,
  WorkbookCommand,
  WorkbookGeometryProfile,
  WorkbookNotesControls,
  WorkbookPivotDefinition,
  WorkbookSheet,
} from '../../core/contracts';
import { WorkbookActions, type WorkbookActionsControls } from '../WorkbookActions';
import { WorkbookCalculation } from '../WorkbookCalculation';
import { WorkbookCellInsertions } from '../WorkbookCellInsertions';
import { WorkbookCreateNames } from '../WorkbookCreateNames';
import { WorkbookDuplicates } from '../WorkbookDuplicates';
import { WorkbookErrors } from '../WorkbookErrors';
import { WorkbookFunctions } from '../WorkbookFunctions';
import { WorkbookHyperlink } from '../WorkbookHyperlink';
import { WorkbookIllustrations } from '../WorkbookIllustrations';
import { WorkbookLocalExecution } from '../WorkbookLocalExecution';
import { WorkbookNotes } from '../WorkbookNotes';
import { WorkbookObjects } from '../WorkbookObjects';
import { WorkbookPageStyles } from '../WorkbookPageStyles';
import { WorkbookPivotContext } from '../WorkbookPivotContext';
import { WorkbookPivotDesign } from '../WorkbookPivotDesign';
import { WorkbookQueries, type WorkbookQueryControls } from '../WorkbookQueries';
import { WorkbookReferences } from '../WorkbookReferences';
import { WorkbookRefreshAll } from '../WorkbookRefreshAll';
import { WorkbookReview } from '../WorkbookReview';
import type { WorkbookReviewFrame, WorkbookReviewMode } from '../workbookReviewController';
import type { WorkbookRibbonTool, WorkbookTool } from '../SpreadsheetRibbon';
import { WorkbookSort } from '../WorkbookSort';
import { WorkbookTextColumns } from '../WorkbookTextColumns';
import { WorkbookTrace } from '../WorkbookTrace';
import { WorkbookWatch } from '../WorkbookWatch';
import type { WorksheetObjectFrame } from '../worksheetObjectController';
import { NameManagerPane } from './NameManagerPane';

export interface GridTaskPaneInputs {
  accessibilityControls: WorkbookAccessibilityControls | undefined;
  accessibilityFrame: AccessibilityFrame | null;
  acrossContext: FillAcrossContext | null;
  actionsControls: WorkbookActionsControls | undefined;
  actionsFrame: { key: string; sheet: string; controls: WorkbookActionsControls | undefined };
  actionsIntent: { serial: number; mode: 'new' | 'list' | 'run'; action?: LocalExecutionAction };
  activationEpoch: React.MutableRefObject<number>;
  active: boolean;
  activeRange: () => string;
  activeTool: WorkbookTool | null;
  busy: boolean;
  calculationController: GridCalculationController;
  calculationFrame: React.MutableRefObject<CalculationFrame | null>;
  cancelCellFocus: () => void;
  cellInsertionControls: boolean | undefined;
  cellInsertionFrame: CellInsertionFrame | null;
  cellInsertionMode: 'checkbox' | 'symbols';
  chartFrame: ChartRecommendationFrame | null;
  chartToolsMode: 'manage' | 'recommendations';
  closeTaskPane: () => void;
  commandScope: RibbonDispatchScope;
  conditionalControls: GridConditionalControls;
  conditionalScope: string;
  createNamesContext: CreateNamesContext | null;
  currentPivot: WorkbookPivotDefinition | null;
  documentId: string | undefined;
  duplicateContext: DuplicateContext | null;
  errorScanner: ErrorScanner;
  errorScope: ErrorScanScope;
  fillToolMode: 'series' | 'across';
  formulaPreviewController: FormulaDraftPreviewController;
  formulaPreviewFrame: React.MutableRefObject<FormulaPreviewFrame | null>;
  functionBrowse: { mode: 'all' | 'recent' | FunctionCategory; serial: number };
  functionContext: FunctionInsertionContext | null;
  functionViewMode: React.MutableRefObject<{ mode: 'normal' | 'pageLayout' | 'pageBreak'; epoch: number }>;
  geometryProfile: WorkbookGeometryProfile;
  goPivotSource: (key: string) => void;
  goToName: (formula: string) => void;
  hyperlinkControls: GridHyperlinkControls;
  hyperlinkFrame: HyperlinkFrame | null;
  illustrationFrame: WorksheetObjectFrame | null;
  illustrationMode: IllustrationMode;
  layoutReason: string | undefined;
  literal: boolean | undefined;
  liveConditionalScope: React.MutableRefObject<string>;
  localExecutionAvailable: boolean;
  localExecutionFrame: GridLocalExecutionFrame;
  nameCommands: boolean;
  namedRanges: { name: string; formula: string; sheet?: string; scopeInvalid?: true }[];
  namePage: number;
  nameScope: 'workbook' | 'worksheet';
  navigateMetadataRange: (name: string, target: string) => void;
  navigateTrace: (target: TraceEndpoint) => void;
  navigationGeneration: React.MutableRefObject<number>;
  newRangeName: string;
  normalMode: React.MutableRefObject<boolean>;
  notesControls: WorkbookNotesControls | undefined;
  notesFrame: NoteFrame | null;
  objectControls: GridObjectControls;
  onBatch: ((commands: WorkbookCommand[]) => void) | undefined;
  openCharts: (binding: RibbonBinding, mode: 'manage' | 'recommendations') => void;
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
  printControls: GridPrintControls;
  printFocus: PrintFocusRequest | null;
  printPreview: AcceptedPrintPreview;
  queriesAllIntent: number;
  queriesMode: 'all' | 'individual';
  queryAllFrame: GridQueryAllFrame;
  queryControls: WorkbookQueryControls | undefined;
  readOnly: boolean;
  readRange: ReadRange | undefined;
  renderedPivotKey: string | undefined;
  replacementContext: ReplaceContext | null;
  reviewFrame: WorkbookReviewFrame | null;
  reviewIntent: { serial: number; mode: WorkbookReviewMode };
  runPivotLayout: (key: string, draft: PivotLayoutDraft) => Promise<void>;
  runPivotRefresh: (key: string) => Promise<void>;
  runPivotUpdate: (key: string, draft: PivotFieldDraft) => Promise<void>;
  scan: GridScan;
  selected: string;
  send: (commands: WorkbookCommand[], allowSpecial?: boolean) => void;
  seriesContext: FillSeriesContext | null;
  setAcceptedConditional: GridSetAcceptedConditional;
  setNamePage: React.Dispatch<React.SetStateAction<number>>;
  setNameScope: React.Dispatch<React.SetStateAction<'workbook' | 'worksheet'>>;
  setNewRangeName: React.Dispatch<React.SetStateAction<string>>;
  setSortHeader: React.Dispatch<React.SetStateAction<boolean>>;
  setTablePage: React.Dispatch<React.SetStateAction<number>>;
  sheet: WorkbookSheet;
  sheets: WorkbookSheet[];
  sortContext: SortContext | null;
  specialContext: SpecialSelectionContext | null;
  storageKey: string | undefined;
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
  validationControls: GridValidationControls;
  viewActive: React.MutableRefObject<boolean>;
  visibleTool: WorkbookRibbonTool | undefined;
  visitedTools: WorkbookTool[];
  watchReader: WatchReader;
  watchScope: WatchScope;
  watchSelectionReason: string | undefined;
  words: (en: string) => string;
}

export function GridTaskPane({
  accessibilityControls,
  accessibilityFrame,
  acrossContext,
  actionsControls,
  actionsFrame,
  actionsIntent,
  activationEpoch,
  active,
  activeRange,
  activeTool,
  busy,
  calculationController,
  calculationFrame,
  cancelCellFocus,
  cellInsertionControls,
  cellInsertionFrame,
  cellInsertionMode,
  chartFrame,
  chartToolsMode,
  closeTaskPane,
  commandScope,
  conditionalControls,
  conditionalScope,
  createNamesContext,
  currentPivot,
  documentId,
  duplicateContext,
  errorScanner,
  errorScope,
  fillToolMode,
  formulaPreviewController,
  formulaPreviewFrame,
  functionBrowse,
  functionContext,
  functionViewMode,
  geometryProfile,
  goPivotSource,
  goToName,
  hyperlinkControls,
  hyperlinkFrame,
  illustrationFrame,
  illustrationMode,
  layoutReason,
  literal,
  liveConditionalScope,
  localExecutionAvailable,
  localExecutionFrame,
  nameCommands,
  namedRanges,
  namePage,
  nameScope,
  navigateMetadataRange,
  navigateTrace,
  navigationGeneration,
  newRangeName,
  normalMode,
  notesControls,
  notesFrame,
  objectControls,
  onBatch,
  openCharts,
  ownerBusy,
  ownerIdentity,
  pageStylesFrame,
  pageStylesMode,
  pageTheme,
  passiveNotice,
  pivotFrame,
  pivotNotice,
  pivotReason,
  pivotRunning,
  pivots,
  printControls,
  printFocus,
  printPreview,
  queriesAllIntent,
  queriesMode,
  queryAllFrame,
  queryControls,
  readOnly,
  readRange,
  renderedPivotKey,
  replacementContext,
  reviewFrame,
  reviewIntent,
  runPivotLayout,
  runPivotRefresh,
  runPivotUpdate,
  scan,
  selected,
  send,
  seriesContext,
  setAcceptedConditional,
  setNamePage,
  setNameScope,
  setNewRangeName,
  setSortHeader,
  setTablePage,
  sheet,
  sheets,
  sortContext,
  specialContext,
  storageKey,
  tablePage,
  tables,
  taskPaneHeading,
  taskPaneId,
  taskPaneTitleId,
  textColumnContext,
  toggleTool,
  toolFocusGeneration,
  traceController,
  traceControls,
  traceInventoryReason,
  traceScope,
  validationControls,
  viewActive,
  visibleTool,
  visitedTools,
  watchReader,
  watchScope,
  watchSelectionReason,
  words,
}: GridTaskPaneInputs): JSX.Element {
  return (
    <aside
      id={taskPaneId}
      className="sheet-task-pane"
      aria-label={words('Workbook tools')}
      hidden={!visibleTool}
      ref={(element) => {
        if (element) element.inert = !visibleTool;
      }}
      onCopy={(event) => event.stopPropagation()}
      onCut={(event) => event.stopPropagation()}
      onPaste={(event) => event.stopPropagation()}
      onFocusCapture={cancelCellFocus}
      onKeyDown={(event) => {
        if (event.key === 'Escape' && !event.defaultPrevented && !event.nativeEvent.isComposing) {
          event.preventDefault();
          event.stopPropagation();
          closeTaskPane();
        }
      }}
    >
      <header className="sheet-task-pane-header">
        <div>
          <small>{words('Workbook tools')}</small>
          <h2 id={taskPaneTitleId} ref={taskPaneHeading} tabIndex={-1}>
            {words(visibleTool?.label ?? 'Workbook tools')}
          </h2>
        </div>
        <button
          type="button"
          aria-label={words('Close workbook tools')}
          title={words('Close workbook tools')}
          onClick={closeTaskPane}
        >
          ×
        </button>
      </header>
      <div className="sheet-task-pane-body">
        {visitedTools.some((tool) => tool === 'checkbox' || tool === 'symbols') &&
          cellInsertionControls &&
          !literal &&
          !passiveNotice && (
            <RetainedWorkbookTool
              tool={cellInsertionMode}
              visible={activeTool === 'checkbox' || activeTool === 'symbols'}
            >
              <WorkbookCellInsertions frame={cellInsertionFrame} mode={cellInsertionMode} words={words} />
            </RetainedWorkbookTool>
          )}
        {visitedTools.includes('accessibility') && accessibilityControls && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="accessibility" visible={activeTool === 'accessibility'}>
            <WorkbookAccessibility frame={accessibilityFrame} words={words} />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('notes') && notesControls && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="notes" visible={activeTool === 'notes'}>
            <WorkbookNotes frame={notesFrame} words={words} />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('hyperlink') && hyperlinkControls && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="hyperlink" visible={activeTool === 'hyperlink'}>
            <WorkbookHyperlink frame={hyperlinkFrame} sheets={sheets.map((s) => s.name)} words={words} />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('special') && !passiveNotice && (
          <RetainedWorkbookTool tool="special" visible={activeTool === 'special'}>
            <WorkbookSpecialSelection
              context={specialContext}
              visible={activeTool === 'special' && active}
              words={words}
            />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('replace') && !passiveNotice && (
          <RetainedWorkbookTool tool="replace" visible={activeTool === 'replace'}>
            <WorkbookReplace context={replacementContext} visible={activeTool === 'replace' && active} words={words} />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('fill') && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="fill" visible={activeTool === 'fill'}>
            {fillToolMode === 'across' ? (
              <WorkbookFillAcross context={acrossContext} visible={activeTool === 'fill' && active} words={words} />
            ) : (
              <WorkbookFill context={seriesContext} visible={activeTool === 'fill' && active} words={words} />
            )}
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('actions') && (actionsControls || localExecutionAvailable) && (
          <RetainedWorkbookTool tool="actions" visible={activeTool === 'actions'}>
            {actionsControls && (
              <div hidden={actionsIntent.mode === 'run'}>
                <WorkbookActions
                  frame={actionsFrame}
                  visible={activeTool === 'actions' && active && actionsIntent.mode !== 'run'}
                  openIntent={
                    actionsIntent.mode === 'run'
                      ? undefined
                      : { serial: actionsIntent.serial, mode: actionsIntent.mode }
                  }
                  words={words}
                />
              </div>
            )}
            {localExecutionAvailable && (
              <WorkbookLocalExecution
                frame={localExecutionFrame}
                intent={
                  actionsIntent.mode === 'run' && actionsIntent.action
                    ? { serial: actionsIntent.serial, action: actionsIntent.action }
                    : null
                }
                words={words}
              />
            )}
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('review') && reviewFrame && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="review" visible={activeTool === 'review'}>
            <WorkbookReview
              frame={reviewFrame}
              visible={activeTool === 'review'}
              openIntent={reviewIntent}
              words={words}
            />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('duplicates') && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="duplicates" visible={activeTool === 'duplicates'}>
            <WorkbookDuplicates context={duplicateContext} visible={activeTool === 'duplicates'} words={words} />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('sort') && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="sort" visible={activeTool === 'sort'}>
            <WorkbookSort
              context={sortContext}
              visible={activeTool === 'sort'}
              words={words}
              onHeaderChange={setSortHeader}
            />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('calculation') && calculationFrame.current && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="calculation" visible={activeTool === 'calculation'}>
            <WorkbookCalculation frame={calculationFrame.current} controller={calculationController} words={words} />
          </RetainedWorkbookTool>
        )}
        <NameManagerPane
          activeRange={activeRange}
          activeTool={activeTool}
          busy={busy}
          goToName={goToName}
          literal={literal}
          nameCommands={nameCommands}
          namedRanges={namedRanges}
          namePage={namePage}
          nameScope={nameScope}
          navigateMetadataRange={navigateMetadataRange}
          newRangeName={newRangeName}
          readOnly={readOnly}
          send={send}
          setNamePage={setNamePage}
          setNameScope={setNameScope}
          setNewRangeName={setNewRangeName}
          setTablePage={setTablePage}
          sheet={sheet}
          tablePage={tablePage}
          tables={tables}
          visitedTools={visitedTools}
          words={words}
        />
        {visitedTools.includes('createNames') && createNamesContext && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="createNames" visible={activeTool === 'createNames'}>
            <WorkbookCreateNames context={createNamesContext} visible={activeTool === 'createNames'} words={words} />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('textColumns') && textColumnContext && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="textColumns" visible={activeTool === 'textColumns'}>
            <WorkbookTextColumns context={textColumnContext} visible={activeTool === 'textColumns'} words={words} />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('errors') && scan && ownerIdentity && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="errors" visible={activeTool === 'errors'}>
            <WorkbookErrors
              scope={errorScope}
              scan={errorScanner}
              goToIssue={(reference) => {
                if (viewActive.current && normalMode.current && activeTool === 'errors')
                  navigateMetadataRange(sheet.name, reference);
              }}
              words={words}
            />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('trace') && traceControls && ownerIdentity && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="trace" visible={activeTool === 'trace'}>
            <WorkbookTrace
              controller={traceController}
              scope={traceScope}
              onNavigate={navigateTrace}
              disabledReason={traceInventoryReason}
              words={words}
            />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('watch') && readRange && ownerIdentity && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="watch" visible={activeTool === 'watch'}>
            <WorkbookWatch
              scope={watchScope}
              read={watchReader}
              selected={sheet ? { sheet: sheet.name, reference: selected } : null}
              selectionReason={watchSelectionReason}
              words={words}
            />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('functions') && functionContext && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="functions" visible={activeTool === 'functions'}>
            <WorkbookFunctions
              context={functionContext}
              visible={activeTool === 'functions'}
              browse={functionBrowse}
              words={words}
              formulaPreview={
                formulaPreviewFrame.current
                  ? { controller: formulaPreviewController, scope: formulaPreviewFrame.current.scope }
                  : undefined
              }
              onRestart={() => toggleTool('functions', commandScope.captureBinding(), 'all')}
            />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('pivot') && currentPivot && pivotFrame.current && (
          <RetainedWorkbookTool tool="pivot" visible={activeTool === 'pivot'}>
            <WorkbookPivotContext
              pivot={currentPivot}
              reason={pivotReason}
              running={pivotRunning}
              bindingKey={renderedPivotKey}
              sourceSheets={sheets.map((sheet) => sheet.name)}
              onUpdate={(draft) => runPivotUpdate(renderedPivotKey!, draft)}
              onRefresh={() => runPivotRefresh(renderedPivotKey!)}
              onSource={
                currentPivot.sourceSheet && currentPivot.sourceRange
                  ? () => goPivotSource(renderedPivotKey!)
                  : undefined
              }
              words={words}
            />
            {pivotNotice?.key === pivotFrame.current.key && <p role="status">{words(pivotNotice.text)}</p>}
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('pivotDesign') && currentPivot && pivotFrame.current && (
          <RetainedWorkbookTool tool="pivotDesign" visible={activeTool === 'pivotDesign'}>
            <WorkbookPivotDesign
              pivot={currentPivot}
              reason={layoutReason}
              running={pivotRunning}
              bindingKey={renderedPivotKey!}
              visible={activeTool === 'pivotDesign'}
              onApply={(draft) => runPivotLayout(renderedPivotKey!, draft)}
              words={words}
            />
            {pivotNotice?.key === pivotFrame.current.key && <p role="status">{words(pivotNotice.text)}</p>}
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('references') && !literal && !passiveNotice && onBatch && (
          <RetainedWorkbookTool tool="references" visible={activeTool === 'references'}>
            <WorkbookReferences
              sheet={sheet.name}
              selection={activeRange()}
              revision={JSON.stringify([
                documentId,
                ownerIdentity,
                storageKey,
                activationEpoch.current,
                navigationGeneration.current,
                functionViewMode.current.epoch,
                toolFocusGeneration.current,
              ])}
              tables={tables}
              pivots={pivots}
              readOnly={readOnly}
              busy={!!busy}
              active={active && activeTool === 'references' && normalMode.current}
              onBatch={onBatch}
            />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('queries') && queryControls && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="queries" visible={activeTool === 'queries'}>
            <div hidden={queriesMode === 'all'}>
              <WorkbookQueries
                {...queryControls}
                sheet={sheet.name}
                selection={activeRange()}
                visible={activeTool === 'queries' && active && queriesMode === 'individual'}
                readOnly={readOnly || !active}
                busy={!!ownerBusy}
              />
            </div>
            {queryControls.refreshAll && (
              <WorkbookRefreshAll frame={queryAllFrame} intent={queriesAllIntent} words={words} />
            )}
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('illustrations') && illustrationFrame && (
          <RetainedWorkbookTool tool="illustrations" visible={activeTool === 'illustrations'}>
            <WorkbookIllustrations
              frame={illustrationFrame}
              mode={illustrationMode}
              selection={activeRange()}
              words={words}
            />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('objects') && objectControls && !literal && (
          <RetainedWorkbookTool tool="objects" visible={activeTool === 'objects'}>
            <WorkbookObjects
              {...objectControls}
              mode={chartToolsMode}
              recommendationFrame={chartFrame ?? undefined}
              onRecommendationOpen={
                chartFrame ? () => openCharts(commandScope.captureBinding(), 'recommendations') : undefined
              }
              ownerIdentity={ownerIdentity}
              sheet={sheet.name}
              selection={activeRange()}
              readOnly={readOnly}
              busy={!!busy}
              words={words}
            />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('pageStyles') && pageTheme && !literal && !passiveNotice && (
          <RetainedWorkbookTool tool="pageStyles" visible={activeTool === 'pageStyles'}>
            <WorkbookPageStyles frame={pageStylesFrame} mode={pageStylesMode} words={words} />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('print') && printControls && !literal && (
          <RetainedWorkbookTool tool="print" visible={activeTool === 'print'}>
            <WorkbookPrint
              {...printControls}
              previewSession={printPreview}
              focusRequest={printFocus}
              preview={(settings, requestId) => printControls.preview(sheet.name, settings, requestId, geometryProfile)}
              sheet={sheet.name}
              selection={activeRange()}
              readOnly={readOnly}
              busy={!!busy}
              words={words}
            />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('conditional') && conditionalControls && !literal && (
          <RetainedWorkbookTool tool="conditional" visible={activeTool === 'conditional'}>
            <ConditionalRules
              {...conditionalControls}
              sheet={sheet.name}
              selection={activeRange()}
              readOnly={readOnly}
              busy={!!busy}
              words={words}
              onInventory={(definition) => {
                if (liveConditionalScope.current === conditionalScope)
                  setAcceptedConditional({ scope: conditionalScope, definition });
                conditionalControls.onInventory?.(definition);
              }}
            />
          </RetainedWorkbookTool>
        )}
        {visitedTools.includes('validation') && validationControls && !literal && (
          <RetainedWorkbookTool tool="validation" visible={activeTool === 'validation'}>
            <DataValidation
              {...validationControls}
              sheet={sheet.name}
              selection={activeRange()}
              readOnly={readOnly}
              busy={!!busy}
              words={words}
            />
          </RetainedWorkbookTool>
        )}
      </div>
    </aside>
  );
}
