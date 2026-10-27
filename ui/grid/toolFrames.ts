import type { ErrorScanFrame } from '../errorScanAdapter';
import type { ErrorScanScope } from '../errorChecking';
import type { FormulaTraceController, FormulaTraceScope, TraceDirection, TraceEndpoint } from '../formulaTrace';
import type { FormulaTraceFrame, FormulaTraceReceipt } from '../formulaTraceAdapter';
import type {
  GridConditionalControls4,
  GridHyperlinkControls,
  GridObjectControls4,
  GridPageTheme,
  GridPrintControls2,
  GridReviewReader,
  GridScan,
  GridValidationControls3,
} from './gridTypes';
import type { IllustrationMode } from '../WorkbookShapes';
import type { LocalExecutionAction } from '../../core/local-action-execution.mjs';
import type { ReadRange } from '../rangeCells';
import type { ReviewOwnerFrame } from '../reviewScanAdapter';
import type { RibbonBinding, RibbonDispatchScope } from '../ribbonCommands';
import type { WatchReadFrame } from '../watchAdapter';
import type { WatchScope } from '../watchWindow';
import type {
  WorkbookAccessibilityControls,
  WorkbookCommand,
  WorkbookNotesControls,
  WorkbookPivotDefinition,
  WorkbookSheet,
} from '../../core/contracts';
import type { WorkbookActionsControls } from '../WorkbookActions';
import type { WorkbookCalculationControls } from '../calculationJobs';
import type { WorkbookQueryControls } from '../WorkbookQueries';
import type { WorkbookReviewControls } from '../reviewHostAdapter';
import { type WorkbookReviewFrame, workbookReviewScopeKey } from '../workbookReviewController';
import type { WorkbookRibbonTool, WorkbookTool } from '../SpreadsheetRibbon';
import { useRef } from 'react';

export interface ToolFramesInputs {
  accessibilityControls: WorkbookAccessibilityControls | undefined;
  acrossDisabledReason: () => string | undefined;
  actionsControls: WorkbookActionsControls | undefined;
  actionsIntent: { serial: number; mode: 'new' | 'list' | 'run'; action?: LocalExecutionAction };
  activationEpoch: React.MutableRefObject<number>;
  active: boolean;
  activeRange: () => string;
  activeTool: WorkbookTool | null;
  busy: boolean;
  calculationControls: WorkbookCalculationControls | undefined;
  calculationOwnerAvailable: boolean;
  cancelCellFocus: () => void;
  cellInsertionControls: boolean | undefined;
  chartToolsMode: 'manage' | 'recommendations';
  commandBinding: React.MutableRefObject<RibbonBinding | null>;
  commandScope: RibbonDispatchScope;
  conditionalControls: GridConditionalControls4;
  currentPivot: WorkbookPivotDefinition | null;
  documentId: string | undefined;
  duplicateDisabledReason: (action?: 'preview' | 'remove') => string | undefined;
  errorScanFrame: React.MutableRefObject<ErrorScanFrame | null>;
  fallbackDocumentId: string;
  fillToolMode: 'series' | 'across';
  functionViewMode: React.MutableRefObject<{ mode: 'normal' | 'pageLayout' | 'pageBreak'; epoch: number }>;
  hyperlinkControls: GridHyperlinkControls;
  illustrationMode: IllustrationMode;
  literal: boolean | undefined;
  nameCommands: boolean;
  navigateMetadataRange: (name: string, target: string) => void;
  navigationGeneration: React.MutableRefObject<number>;
  normalMode: React.MutableRefObject<boolean>;
  notesControls: WorkbookNotesControls | undefined;
  objectControls: GridObjectControls4;
  onBatch: ((commands: WorkbookCommand[]) => void) | undefined;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  ownerBusy: boolean;
  ownerIdentity: string | undefined;
  pageTheme: GridPageTheme;
  passiveNotice: string | undefined;
  printControls: GridPrintControls2;
  queriesMode: 'individual' | 'all';
  queryControls: WorkbookQueryControls | undefined;
  readOnly: boolean;
  readRange: ReadRange | undefined;
  replacementDisabledReason: (action?: 'preview' | 'apply') => string | undefined;
  reviewControls: WorkbookReviewControls | undefined;
  reviewReader: GridReviewReader;
  reviewReadFrame: React.MutableRefObject<ReviewOwnerFrame | null>;
  scan: GridScan;
  selected: string;
  seriesDisabledReason: (action?: 'preview' | 'apply') => string | undefined;
  setActiveTool: React.Dispatch<React.SetStateAction<WorkbookTool | null>>;
  setTraceRoot: React.Dispatch<React.SetStateAction<{ sheet: string; cell: string } | null>>;
  setVisitedTools: React.Dispatch<React.SetStateAction<WorkbookTool[]>>;
  sheet: WorkbookSheet;
  sheets: WorkbookSheet[];
  sortDisabledReason: () => string | undefined;
  sortPending: boolean;
  specialDisabledReason: () => string | undefined;
  specialSelection: Set<string> | null;
  storageKey: string | undefined;
  toolFocusGeneration: React.MutableRefObject<number>;
  traceController: FormulaTraceController;
  traceControls: { receipt: FormulaTraceReceipt; trace: FormulaTraceFrame['trace'] } | undefined;
  traceFrame: React.MutableRefObject<FormulaTraceFrame | null>;
  traceRoot: { sheet: string; cell: string } | null;
  validationControls: GridValidationControls3;
  viewActive: React.MutableRefObject<boolean>;
  viewMode: 'normal' | 'pageLayout' | 'pageBreak';
  watchReadFrame: React.MutableRefObject<WatchReadFrame | null>;
}

export function useToolFrames({
  accessibilityControls,
  acrossDisabledReason,
  actionsControls,
  actionsIntent,
  activationEpoch,
  active,
  activeRange,
  activeTool,
  busy,
  calculationControls,
  calculationOwnerAvailable,
  cancelCellFocus,
  cellInsertionControls,
  chartToolsMode,
  commandBinding,
  commandScope,
  conditionalControls,
  currentPivot,
  documentId,
  duplicateDisabledReason,
  errorScanFrame,
  fallbackDocumentId,
  fillToolMode,
  functionViewMode,
  hyperlinkControls,
  illustrationMode,
  literal,
  nameCommands,
  navigateMetadataRange,
  navigationGeneration,
  normalMode,
  notesControls,
  objectControls,
  onBatch,
  onCommandBatch,
  ownerBusy,
  ownerIdentity,
  pageTheme,
  passiveNotice,
  printControls,
  queriesMode,
  queryControls,
  readOnly,
  readRange,
  replacementDisabledReason,
  reviewControls,
  reviewReader,
  reviewReadFrame,
  scan,
  selected,
  seriesDisabledReason,
  setActiveTool,
  setTraceRoot,
  setVisitedTools,
  sheet,
  sheets,
  sortDisabledReason,
  sortPending,
  specialDisabledReason,
  specialSelection,
  storageKey,
  toolFocusGeneration,
  traceController,
  traceControls,
  traceFrame,
  traceRoot,
  validationControls,
  viewActive,
  viewMode,
  watchReadFrame,
}: ToolFramesInputs) {
  const tools: WorkbookRibbonTool[] = [];
  if (!passiveNotice)
    tools.push({ id: 'replace', label: 'Replacement preview', disabled: !!replacementDisabledReason() });
  if (!passiveNotice) tools.push({ id: 'special', label: 'Go To Special', disabled: !!specialDisabledReason() });
  const localExecutionAvailable = !!onCommandBatch && !!ownerIdentity && !literal && !passiveNotice;
  if (actionsControls || localExecutionAvailable)
    tools.push({ id: 'actions', label: 'Local Actions', disabled: !active || !actionsControls });
  if (!literal) {
    if (!passiveNotice)
      tools.push({
        id: 'fill',
        label: fillToolMode === 'across' ? 'Across Worksheets' : 'Fill series',
        disabled: !!(fillToolMode === 'across' ? acrossDisabledReason() : seriesDisabledReason()),
      });
    if (!passiveNotice)
      tools.push({ id: 'duplicates', label: 'Remove Duplicates', disabled: !!duplicateDisabledReason() });
    if (!passiveNotice) tools.push({ id: 'sort', label: 'Sort', disabled: !!sortDisabledReason() || sortPending });
    tools.push({ id: 'names', label: 'Names and tables', disabled: !active });
    if (!passiveNotice && calculationControls && calculationOwnerAvailable)
      tools.push({ id: 'calculation', label: 'Calculation Options', disabled: !active });
    if (!passiveNotice && nameCommands && readRange && ownerIdentity && onCommandBatch)
      tools.push({
        id: 'createNames',
        label: 'Create from Selection',
        disabled: !active || readOnly || !!busy || !!specialSelection,
      });
    if (!passiveNotice) tools.push({ id: 'functions', label: 'Insert Function', disabled: !active });
    if (!passiveNotice && readRange && ownerIdentity && sheets.length <= 1000)
      tools.push({ id: 'watch', label: 'Watch Window', disabled: !active });
    if (!passiveNotice && scan && ownerIdentity)
      tools.push({ id: 'errors', label: 'Error Checking', disabled: !active });
    if (!passiveNotice && scan && readRange && ownerIdentity && reviewControls)
      tools.push({ id: 'review', label: 'Review', disabled: !active });
    if (!passiveNotice && traceControls && ownerIdentity)
      tools.push({ id: 'trace', label: 'Formula tracing', disabled: !active });
    if (
      !passiveNotice &&
      readRange &&
      ownerIdentity &&
      nameCommands &&
      onCommandBatch &&
      conditionalControls &&
      validationControls
    )
      tools.push({ id: 'textColumns', label: 'Text to Columns', disabled: !active });
    if (!passiveNotice && cellInsertionControls && readRange && onCommandBatch)
      tools.push(
        { id: 'checkbox', label: 'Checkbox', disabled: !active || viewMode !== 'normal' },
        { id: 'symbols', label: 'Symbols', disabled: !active || viewMode !== 'normal' }
      );
    if (!passiveNotice && accessibilityControls)
      tools.push({
        id: 'accessibility',
        label: 'Check Accessibility',
        disabled: !active || viewMode !== 'normal' || !!busy,
      });
    if (!passiveNotice && notesControls)
      tools.push({ id: 'notes', label: 'Notes', disabled: !active || viewMode !== 'normal' });
    if (!passiveNotice && hyperlinkControls)
      tools.push({ id: 'hyperlink', label: 'Link', disabled: !active || viewMode !== 'normal' });
    if (objectControls && ownerIdentity && onCommandBatch && !passiveNotice)
      tools.push({
        id: 'illustrations',
        label:
          illustrationMode === 'text'
            ? 'Text Box'
            : illustrationMode === 'icon'
              ? 'Icons'
              : illustrationMode === 'wordart'
                ? 'WordArt'
                : illustrationMode === 'smartart'
                  ? 'SmartArt'
                  : illustrationMode === 'model3d'
                    ? '3D Models'
                    : 'Shapes',
        disabled: !active || viewMode !== 'normal',
      });
    if (objectControls)
      tools.push({
        id: 'objects',
        label: chartToolsMode === 'recommendations' ? 'Recommended Charts' : 'Charts and images',
        disabled: !active,
      });
    if (pageTheme && ownerIdentity && onCommandBatch && !passiveNotice)
      tools.push({ id: 'pageStyles', label: 'Themes and backgrounds', disabled: !active || viewMode !== 'normal' });
    if (printControls) tools.push({ id: 'print', label: 'Print and PDF', disabled: !active });
    if (conditionalControls) tools.push({ id: 'conditional', label: 'Conditional formatting', disabled: !active });
    if (validationControls) tools.push({ id: 'validation', label: 'Data validation', disabled: !active });
    if (!passiveNotice && onBatch) tools.push({ id: 'references', label: 'Tables and pivots', disabled: !active });
    if (currentPivot)
      tools.push(
        { id: 'pivot', label: 'PivotTable Fields', disabled: !active },
        { id: 'pivotDesign', label: 'PivotTable Styles', disabled: !active }
      );
    if (!passiveNotice && queryControls) tools.push({ id: 'queries', label: 'Queries and refresh', disabled: !active });
  }
  const actionsFrame = {
    key: JSON.stringify([
      documentId ?? fallbackDocumentId,
      ownerIdentity ?? storageKey ?? 'local',
      commandBinding.current!.state,
      activationEpoch.current,
      active,
      viewMode,
    ]),
    sheet: sheet.name,
    controls: active && normalMode.current ? actionsControls : undefined,
  };
  const localExecutionOwnerKey = JSON.stringify([
    documentId ?? fallbackDocumentId,
    ownerIdentity ?? '',
    sheet.name,
    activationEpoch.current,
    navigationGeneration.current,
    functionViewMode.current.epoch,
    toolFocusGeneration.current,
  ]);
  const localExecutionKey = JSON.stringify([localExecutionOwnerKey, commandBinding.current!.state, activeRange()]);
  const localExecutionReason = !localExecutionAvailable
    ? 'Local action execution is unavailable in this host.'
    : !normalMode.current
      ? 'Return to Normal view to run actions.'
      : readOnly
        ? 'This workbook is read-only.'
        : busy
          ? 'Wait for the workbook operation to finish.'
          : specialSelection
            ? 'Click a regular cell before running a local action.'
            : undefined;
  const localExecutionActive = active && activeTool === 'actions' && actionsIntent.mode === 'run' && normalMode.current;
  const localExecutionLive = useRef({
    key: localExecutionKey,
    active: localExecutionActive,
    reason: localExecutionReason,
  });
  localExecutionLive.current = { key: localExecutionKey, active: localExecutionActive, reason: localExecutionReason };
  const localExecutionFrame = localExecutionAvailable
    ? {
        key: localExecutionKey,
        ownerKey: localExecutionOwnerKey,
        sheet: sheet.name,
        selection: activeRange(),
        active: localExecutionActive,
        reason: localExecutionReason,
        apply: async (commands: WorkbookCommand[]) => {
          if (
            localExecutionLive.current.key !== localExecutionKey ||
            !localExecutionLive.current.active ||
            localExecutionLive.current.reason ||
            !viewActive.current ||
            !onCommandBatch
          )
            throw Error('The original action worksheet or selection changed. Preview again.');
          await onCommandBatch(commands);
        },
      }
    : null;
  const queryAllOwnerKey = JSON.stringify([
    documentId ?? fallbackDocumentId,
    ownerIdentity ?? '',
    activationEpoch.current,
    functionViewMode.current.epoch,
    toolFocusGeneration.current,
  ]);
  const queryAllReason = !normalMode.current
    ? 'Return to Normal view to refresh queries.'
    : readOnly
      ? 'This workbook is read-only.'
      : undefined;
  const queryAllFrame = queryControls?.refreshAll
    ? {
        key: JSON.stringify([queryAllOwnerKey, queryControls.revision]),
        ownerKey: queryAllOwnerKey,
        name: queryControls.refreshAll.name ?? 'Current workbook',
        active: active && activeTool === 'queries' && queriesMode === 'all' && normalMode.current,
        busy: !!ownerBusy,
        reason: queryAllReason,
        controls: queryControls.refreshAll,
      }
    : null;
  const errorScope: ErrorScanScope = {
    documentId: documentId ?? fallbackDocumentId,
    ownerIdentity: ownerIdentity ?? storageKey ?? 'local',
    state: commandBinding.current?.state ?? 'local',
    sheet: sheet.name,
    activationEpoch: activationEpoch.current,
    active: active && activeTool === 'errors' && normalMode.current,
  };
  errorScanFrame.current = scan ? { scope: errorScope, scan } : null;
  const watchScope: WatchScope = {
    documentId: documentId ?? fallbackDocumentId,
    ownerIdentity: ownerIdentity ?? storageKey ?? 'local',
    state: commandBinding.current?.state ?? 'local',
    activationEpoch: activationEpoch.current,
    active: active && activeTool === 'watch' && normalMode.current,
    sheets: sheets.map((item) => item.name),
  };
  watchReadFrame.current = readRange ? { scope: watchScope, read: readRange } : null;
  const reviewScope = reviewControls
    ? {
        ...reviewControls.receipt,
        ownerIdentity: ownerIdentity ?? '',
        sheet: sheet.name,
        selection: activeRange(),
        activationEpoch: activationEpoch.current,
        navigationGeneration: navigationGeneration.current,
        modeEpoch: functionViewMode.current.epoch,
        active:
          active && activeTool === 'review' && normalMode.current && !literal && !passiveNotice && !specialSelection,
        writable: !readOnly && !busy && !!onCommandBatch,
        sheets: sheets.map((item) => ({ name: item.name, state: item.state ?? 'visible' })),
      }
    : null;
  const reviewIdentity = reviewScope ? workbookReviewScopeKey(reviewScope) : null;
  if (reviewScope && reviewIdentity && scan && readRange)
    reviewReadFrame.current = {
      receipt: reviewControls!.receipt,
      identity: reviewIdentity,
      active: reviewScope.active,
      sheets: sheets.map((item) => item.name),
      scan,
      read: async (sheet, cell) => ({ cells: await readRange(sheet, cell) }),
    };
  const reviewFrame: WorkbookReviewFrame | null =
    reviewScope && reviewIdentity && reviewControls && scan && readRange
      ? {
          scope: reviewScope,
          callbacks: {
            scan: (sheet, after, limit, signal) =>
              reviewReader.scan(reviewControls.receipt, reviewIdentity, sheet, after, limit, signal),
            read: (sheet, cell, signal) =>
              reviewReader.read(reviewControls.receipt, reviewIdentity, sheet, cell, signal),
            availability: reviewControls.availability,
            check: reviewControls.check,
            cancel: reviewControls.cancel,
            apply: async (commands) => {
              if (
                !reviewReadFrame.current?.active ||
                reviewReadFrame.current.identity !== reviewIdentity ||
                !reviewScope.writable ||
                !onCommandBatch
              )
                throw Error('The original Review workbook changed or is read-only.');
              await onCommandBatch(commands);
            },
            navigate: (sheet, cell) => {
              if (reviewReadFrame.current?.active && reviewReadFrame.current.identity === reviewIdentity)
                navigateMetadataRange(sheet, cell);
            },
          },
        }
      : null;
  const traceScope: FormulaTraceScope = {
    documentId: documentId ?? fallbackDocumentId,
    ownerIdentity: ownerIdentity ?? storageKey ?? 'local',
    state: JSON.stringify([
      commandBinding.current?.state ?? 'local',
      traceControls?.receipt.documentId ?? null,
      traceControls?.receipt.ownerId ?? null,
      traceControls?.receipt.state ?? null,
      traceControls?.receipt.revision ?? null,
    ]),
    sheet: sheet.name,
    cell: traceRoot?.sheet === sheet.name ? traceRoot.cell : selected,
    activationEpoch: activationEpoch.current,
    active: active && normalMode.current && !literal && !passiveNotice && sheets.length <= 1024,
    sheets: sheets.map((item) => item.name),
  };
  const traceInventoryReason =
    sheets.length > 1024 ? 'Formula tracing supports up to 1,024 worksheets in a workbook.' : undefined;
  traceFrame.current = traceControls ? { scope: traceScope, ...traceControls } : null;
  const openTracePane = (): void => {
    if (!viewActive.current || !normalMode.current || !traceControls) return;
    cancelCellFocus();
    toolFocusGeneration.current++;
    setVisitedTools((previous) => (previous.includes('trace') ? previous : [...previous, 'trace']));
    setActiveTool('trace');
  };
  const navigateTrace = (target: TraceEndpoint): void => {
    if (viewActive.current && normalMode.current && traceFrame.current?.scope.active)
      navigateMetadataRange(target.sheet, target.range);
  };
  const startTrace = async (binding: RibbonBinding, direction: TraceDirection): Promise<void> => {
    if (!traceControls || !commandScope.isCurrent(binding) || !normalMode.current)
      throw Error('Return to the active Normal worksheet to trace formulas.');
    if (specialSelection) throw Error('Click a regular cell to trace formulas.');
    const scope = { ...traceScope, cell: selected };
    setTraceRoot({ sheet: scope.sheet, cell: scope.cell });
    traceFrame.current = { scope, ...traceControls };
    traceController.setScope(scope);
    openTracePane();
    await traceController.trace(direction);
  };
  return {
    tools,
    localExecutionAvailable,
    actionsFrame,
    localExecutionReason,
    localExecutionFrame,
    queryAllReason,
    queryAllFrame,
    errorScope,
    watchScope,
    reviewFrame,
    traceScope,
    traceInventoryReason,
    openTracePane,
    navigateTrace,
    startTrace,
  } as const;
}
