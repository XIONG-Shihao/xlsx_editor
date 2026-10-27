import { useRangeCells } from './rangeCells';
import { useDocumentWords } from './words';
import { useEffect, useState, useRef, useMemo, useId } from 'react';
import { createInkPreferenceStore, INK_PREFERENCE_KEY } from './inkPreferences';
import type { InkPenId } from '../core/ink-stroke.mjs';
import { boundedRangeRead } from './boundedRangeRead';
import { createWorksheetDimensions, normalGeometryFont } from '../core/worksheet-geometry.mjs';
import { useWorksheetGeometryProfile } from './worksheetFontMetrics';
import { homeBindingMatches } from './homeCommands';
import {
  BooleanCellControls,
  CalculationRevision,
  ConditionalDefinition,
  OwnedPivotPresentation,
} from './grid/gridContexts';

export { acceptGridCommandBatch, normalizeGridCommandBatch } from './grid/commandBatch';
import type { SpreadsheetGridProps } from './grid/gridProps';
import { buildCalculationActions } from './grid/calculationActions';
import { buildFillSeriesContext } from './grid/fillSeriesContext';
import { buildFilterActions } from './grid/filterActions';
import { buildHomeDataCommands } from './grid/homeDataCommands';
import { buildNamePickerActions } from './grid/namePickerActions';
import { buildNavigationActions } from './grid/navigationActions';
import { buildPageAndPivotActions } from './grid/pageAndPivotActions';
import { buildReplaceContext } from './grid/replaceContext';
import { buildSortDuplicateContext } from './grid/sortDuplicateContext';
import { buildSortFilterCommands } from './grid/sortFilterCommands';
import { buildTaskPaneActions } from './grid/taskPaneActions';
import { buildToolPaneFrames } from './grid/toolPaneFrames';
import { useCommandPlumbing } from './grid/commandPlumbing';
import { useFillAcrossContext } from './grid/fillAcrossContext';
import { useFindFilterState } from './grid/findFilterState';
import { useGridGeometry } from './grid/gridGeometry';
import { useGridWindows } from './grid/gridWindows';
import { usePrintAndViewMode } from './grid/printAndViewMode';
import { useRenderFrames } from './grid/renderFrames';
import { useRuleInventories } from './grid/ruleInventories';
import { useScrollToCell } from './grid/scrollToCell';
import { useSelectionDerived } from './grid/selectionDerived';
import { useSelectionHandlers } from './grid/selectionHandlers';
import { useSelectionState } from './grid/selectionState';
import { useSpecialContext } from './grid/specialContext';
import { useToolFrames } from './grid/toolFrames';
import { useToolPaneState } from './grid/toolPaneState';
import { buildCommandSend } from './grid/commandSend';
import { buildHomeHandle } from './grid/homeHandle';
import { buildKeyboardNavigation } from './grid/keyboardNavigation';
import { useHomeDispatch } from './grid/homeDispatch';
import { useViewPreferences } from './grid/viewPreferences';
import { useSelectionSummary } from './grid/selectionSummary';
import { GridRibbon } from './grid/GridRibbon';
import { GridWorkspace } from './grid/GridWorkspace';
export { applyGridEdit, displayValue, addressPosition, colName } from './sheetModel';
export function SpreadsheetGrid({
  sheets: metadata,
  pageTheme,
  active = true,
  readRange: rawReadRange,
  summarize,
  scan,
  query,
  filterValues,
  duplicatePreview,
  duplicatePreviewReceipt,
  storageKey,
  documentId,
  ownerIdentity,
  acceptedCommandState,
  viewPreferenceId,
  namedRanges = [],
  tables = [],
  pivots = [],
  passiveNotice,
  queryControls,
  traceControls,
  formulaPreviewControls,
  reviewControls,
  actionsControls,
  calculationControls,
  usedRangeControls,
  nameCommands = false,
  cellInsertionControls,
  accessibilityControls,
  notesControls,
  hyperlinkControls,
  objectControls,
  printControls,
  conditionalControls,
  validationControls,
  importPaste,
  importFile,
  readOnly,
  literal,
  structuralReason,
  onEdit,
  onBatch,
  onCommandBatch,
  busy: hostBusy,
  onUndo,
  onCommandUndo,
  onCommandRedo,
  onRedo,
  canUndo,
  canRedo,
  focus,
  status,
}: SpreadsheetGridProps): JSX.Element {
  const readRange = useMemo(
    () => (rawReadRange && passiveNotice ? boundedRangeRead(rawReadRange) : rawReadRange),
    [rawReadRange, !!passiveNotice]
  );
  const [pasteProgress, setPasteProgress] = useState<{ completed: number; total: number } | null>(null);
  const pasteAbort = useRef<AbortController | null>(null);
  const ownerBusy = hostBusy || pasteProgress !== null;
  let busy = ownerBusy;
  useEffect(
    () => () => {
      pasteAbort.current?.abort();
    },
    []
  );
  const ranges = useRangeCells(metadata, readRange, storageKey);
  const sheets = ranges.sheets;
  const activeStorage = useRef(storageKey);
  activeStorage.current = storageKey;
  const viewActive = useRef(active),
    activationEpoch = useRef(0);
  if (viewActive.current !== active) {
    viewActive.current = active;
    activationEpoch.current++;
  }
  const activeFrame = (work: () => void): number => {
    const epoch = activationEpoch.current;
    return requestAnimationFrame(() => {
      if (viewActive.current && epoch === activationEpoch.current) work();
    });
  };
  const navigationGeneration = useRef(0);
  const [, setHomeNavigationVersion] = useState(0);
  const gridRoot = useRef<HTMLDivElement>(null);
  const gridScroll = useRef<HTMLDivElement>(null);
  const inkPreferenceStore = useMemo(() => {
    try {
      return createInkPreferenceStore(window.localStorage);
    } catch {
      return createInkPreferenceStore(null);
    }
  }, []);
  const [inkPreset, setInkPreset] = useState<InkPenId>(() => inkPreferenceStore.read());
  const [inkPreferenceWarning, setInkPreferenceWarning] = useState<string | null>(() => inkPreferenceStore.warning);
  useEffect(() => {
    const refresh = () => {
      setInkPreset(inkPreferenceStore.read());
      setInkPreferenceWarning(inkPreferenceStore.warning);
    };
    refresh();
    const changed = (event: StorageEvent) => {
      if (event.key === INK_PREFERENCE_KEY || event.key === null) refresh();
    };
    window.addEventListener('storage', changed);
    return () => window.removeEventListener('storage', changed);
  }, [inkPreferenceStore, active]);
  const [drawMode, setDrawMode] = useState<'draw' | 'select' | 'erase' | null>(null);
  const pendingCellFocus = useRef<number | null>(null);
  const cancelCellFocus = (): void => {
    if (pendingCellFocus.current !== null) cancelAnimationFrame(pendingCellFocus.current);
    pendingCellFocus.current = null;
  };
  const words = useDocumentWords();
  const [sheetName, setSheetName] = useState(sheets.find((s) => s.state === 'visible')?.name ?? sheets[0]?.name ?? '');
  const sheet = sheets.find((s) => s.name === sheetName) ?? sheets[0];
  const geometryProfile = useWorksheetGeometryProfile(sheet),
    geometryKey = JSON.stringify(geometryProfile);
  const normalFont = normalGeometryFont(sheet);
  const worksheetDimensions = useMemo(
    () => createWorksheetDimensions(sheet, geometryProfile),
    [
      sheet?.sheetFormat,
      sheet?.normalFont,
      sheet?.rowHeights,
      sheet?.hiddenRows,
      sheet?.visibleRows,
      sheet?.widths,
      geometryKey,
    ]
  );
  const fallbackDocumentId = useId();
  const liveFunctionFrame = useRef({
    reader: readRange,
    sheet,
    ownerIdentity: ownerIdentity ?? storageKey ?? 'local',
    grantDocumentId: documentId ?? fallbackDocumentId,
  });
  liveFunctionFrame.current = {
    reader: readRange,
    sheet,
    ownerIdentity: ownerIdentity ?? storageKey ?? 'local',
    grantDocumentId: documentId ?? fallbackDocumentId,
  };
  const [viewMode, setViewMode] = useState<'normal' | 'pageLayout' | 'pageBreak'>('normal'),
    modeGeneration = useRef(0),
    normalMode = useRef(true);
  const functionViewMode = useRef({ mode: viewMode, epoch: 0 });
  if (functionViewMode.current.mode !== viewMode)
    functionViewMode.current = { mode: viewMode, epoch: functionViewMode.current.epoch + 1 };
  normalMode.current = viewMode === 'normal';
  readOnly = readOnly || !normalMode.current;
  const {
    printFocus,
    setPrintFocus,
    printFocusSerial,
    printScope,
    printPreview,
    normalViewportKey,
    normalViewportBookmark,
    normalReturnPending,
  } = usePrintAndViewMode({
    activationEpoch,
    active,
    documentId,
    fallbackDocumentId,
    geometryKey,
    geometryProfile,
    gridScroll,
    printControls,
    sheet,
    storageKey,
    viewMode,
  });
  busy = busy || printPreview.preparing || printPreview.publishing;
  useEffect(() => {
    modeGeneration.current++;
    normalViewportBookmark.current = null;
    normalReturnPending.current = false;
    setViewMode('normal');
  }, [
    printScope.documentId,
    printScope.grantDocumentId,
    printScope.state,
    printScope.sheet,
    printScope.ownerKey,
    printScope.activationEpoch,
    active,
  ]);

  const {
    preferenceDocumentId,
    preferenceStore,
    presetStore,
    view,
    viewport,
    setViewport,
    changeView,
    presetDocumentId,
    namedViewPresets,
    presetBindingKey,
    presetAction,
  } = useViewPreferences({
    activationEpoch,
    activeFrame,
    documentId,
    fallbackDocumentId,
    gridScroll,
    navigationGeneration,
    ownerIdentity,
    sheet,
    viewActive,
    viewPreferenceId,
    late: {
      get scrollCell() {
        return scrollCell;
      },
      get selectedRef() {
        return selectedRef;
      },
      get setGridError() {
        return setGridError;
      },
      get sheetRef() {
        return sheetRef;
      },
    },
  });
  const {
    sheetAction,
    setSheetAction,
    newSheetName,
    setNewSheetName,
    pageStylesMode,
    setPageStylesMode,
    pageStylesLive,
    illustrationMode,
    setIllustrationMode,
    illustrationLive,
    chartToolsMode,
    setChartToolsMode,
    chartLive,
    accessibilityLive,
    notesLive,
    activeTool,
    setActiveTool,
    visitedTools,
    setVisitedTools,
    reviewIntent,
    setReviewIntent,
    queriesMode,
    setQueriesMode,
    queriesAllIntent,
    setQueriesAllIntent,
    actionsIntent,
    setActionsIntent,
    actionsParkedDraft,
    reviewParkedDraft,
    duplicateContext,
    setDuplicateContext,
    duplicateSerial,
    duplicateWritePhysical,
    fillToolMode,
    setFillToolMode,
    acrossContext,
    setAcrossContext,
    acrossSerial,
    acrossWritePhysical,
    acrossLive,
    seriesContext,
    setSeriesContext,
    seriesSerial,
    seriesWritePhysical,
    seriesLive,
    seriesParkedDraft,
    replacementContext,
    setReplacementContext,
    replacementSerial,
    replacementWritePhysical,
    replacementLive,
    replacementParkedDraft,
    specialContext,
    setSpecialContext,
    specialSerial,
    specialSelectPhysical,
    specialLive,
    specialParkedDraft,
    textColumnsParkedDraft,
    duplicateParkedDraft,
    liveDuplicateReceipt,
    sortContext,
    setSortContext,
    sortSerial,
    sortPhysical,
    sortPending,
    setSortPending,
    sortParkedDraft,
    createNamesContext,
    setCreateNamesContext,
    liveNameInventory,
    createNamesParkedDraft,
    functionContext,
    setFunctionContext,
    functionBrowse,
    setFunctionBrowse,
    taskPaneId,
    taskPaneTitleId,
    taskPaneHeading,
    booleanWritePending,
    toolFocusGeneration,
  } = useToolPaneState({ active, duplicatePreviewReceipt, namedRanges, tables, viewActive });
  const {
    conditionalScope,
    liveConditionalScope,
    acceptedConditional,
    setAcceptedConditional,
    validationScope,
    liveValidationScope,
    acceptedValidation,
    namePage,
    setNamePage,
    tablePage,
    setTablePage,
    newRangeName,
    setNewRangeName,
    nameScope,
    setNameScope,
    pendingSheetSelection,
  } = useRuleInventories({
    conditionalControls,
    namedRanges,
    ownerBusy,
    setSheetName,
    sheet,
    sheets,
    tables,
    validationControls,
  });
  const {
    jump,
    setJump,
    jumpError,
    setJumpError,
    gridError,
    setGridError,
    viewNotice,
    setViewNotice,
    calculationRun,
    calculationBusy,
    setCalculationBusy,
    calculationNotice,
    setCalculationNotice,
    calculationMounted,
    calculationOwnerAvailable,
    calculationKey,
    calculationLive,
    calculationFrame,
    calculationController,
    calculationJob,
    commandBinding,
    commandAvailability,
    textColumnContext,
    setTextColumnContext,
    errorScanFrame,
    errorScanner,
    watchReadFrame,
    watchReader,
    reviewReadFrame,
    reviewReader,
    traceFrame,
    traceController,
    traceRoot,
    setTraceRoot,
    pivotFrame,
    pivotFlight,
    pivotRunning,
    setPivotRunning,
    pivotNotice,
    setPivotNotice,
    commandScope,
    onCommandResult,
    invoke,
  } = useCommandPlumbing({
    activationEpoch,
    active,
    documentId,
    fallbackDocumentId,
    functionViewMode,
    navigationGeneration,
    normalMode,
    onCommandBatch,
    ownerIdentity,
    sheet,
    viewActive,
  });
  const {
    findOpen,
    setFindOpen,
    findText,
    setFindText,
    replaceText,
    setReplaceText,
    findScope,
    setFindScope,
    findMode,
    setFindMode,
    findMatches,
    setFindMatches,
    findTotal,
    setFindTotal,
    findPageIndex,
    setFindPageIndex,
    findNextCursor,
    setFindNextCursor,
    findIndex,
    setFindIndex,
    findBusy,
    setFindBusy,
    replacementConfigurationKey,
    replacementConfigurationLive,
    sortPrimary,
    setSortPrimary,
    conditionalRule,
    setConditionalRule,
    conditionalThreshold,
    setConditionalThreshold,
    conditionalColor,
    setConditionalColor,
    tableName,
    setTableName,
    sortHeader,
    setSortHeader,
    filterChoices,
    filterText,
    setFilterText,
    filterPage,
    setFilterPage,
    filterOffset,
    setFilterOffset,
    filterBusy,
    setFilterBusy,
    filterReadPending,
    setFilterReadPending,
    filterApplying,
    setFilterApplying,
    filterPageScope,
    setFilterPageScope,
    filterPhysical,
    filterWritePhysical,
    filterMounted,
    filterNotice,
    setFilterNotice,
    filterSelected,
    setFilterSelected,
    filterGeneration,
    filterScope,
    liveFilterDefinition,
    resetFilter,
    findGeneration,
    foundValues,
    findConfig,
    findPages,
    resetFind,
  } = useFindFilterState({ metadata, sheet, storageKey });
  const {
    range,
    setRange,
    specialSelection,
    setSpecialSelection,
    specialSelectionMode,
    selectionEnd,
    setSelectionEnd,
    objectExtent,
    updateObjectExtent,
    dragging,
    fillDragging,
    fillPreview,
    setFillPreview,
    homeFrameRef,
    homePhysical,
    homePending,
    setHomePending,
    homeMetadata,
    homeSpecial,
    findReceipt,
    paintHomeBinding,
    homePaintTarget,
    homePaintRunner,
    paintStyleId,
    setPaintStyleId,
    paintStyleRef,
    internalClipboard,
    selected,
    setSelected,
    formulaDraftHandle,
    formulaComposing,
    setFormulaComposing,
    formulaPreviewDraft,
    formulaPreviewFrame,
    formulaPreviewController,
    reportFormulaPreviewDraft,
    namePicker,
    setNamePicker,
    namePickerSerial,
    namedDraftLive,
    selectedRef,
  } = useSelectionState({ metadata, setJump, setJumpError, sheet });
  const { scrollCell } = useScrollToCell({
    activeFrame,
    filterGeneration,
    focus,
    geometryProfile,
    gridRoot,
    gridScroll,
    literal,
    preferenceDocumentId,
    preferenceStore,
    setRange,
    setSelected,
    setSelectionEnd,
    setSheetName,
    setViewport,
    sheet,
    sheets,
  });
  const sheetRef = useRef(sheet);
  sheetRef.current = sheet;
  const batchRef = useRef(onBatch);
  batchRef.current = onBatch;
  const canEditRef = useRef(active && !readOnly && !busy);
  canEditRef.current = active && !readOnly && !busy;
  useEffect(() => {
    if (active) return;
    navigationGeneration.current++;
    setHomeNavigationVersion(navigationGeneration.current);
    cancelCellFocus();
    dragging.current = false;
    fillDragging.current = null;
    setFillPreview(null);
  }, [active]);
  const literalRef = useRef(literal);
  literalRef.current = literal;
  const editRef = useRef(onEdit);
  editRef.current = onEdit;
  const { setValue, select, selectRow, selectColumn, dragStart, dragEnter } = useSelectionHandlers({
    dragging,
    editRef,
    fillDragging,
    homeFrameRef,
    homePaintRunner,
    literal,
    navigationGeneration,
    normalMode,
    paintHomeBinding,
    paintStyleRef,
    selectedRef,
    setFillPreview,
    setGridError,
    setHomeNavigationVersion,
    setPaintStyleId,
    setRange,
    setSelected,
    setSelectionEnd,
    setSpecialSelection,
    sheet,
    sheetRef,
    specialSelection,
    viewActive,
  });
  const { send, sendCommand } = buildCommandSend({
    literal,
    normalMode,
    onBatch,
    onCommandBatch,
    readOnly,
    setGridError,
    specialSelection,
    viewActive,
    late: {
      get busy() {
        return busy;
      },
    },
  });
  const { calculateCurrentWorkbook, runCalculationAction } = buildCalculationActions({
    calculationController,
    calculationFrame,
    calculationKey,
    calculationLive,
    calculationMounted,
    calculationOwnerAvailable,
    calculationRun,
    cancelCellFocus,
    commandScope,
    normalMode,
    onCommandBatch,
    passiveNotice,
    setActiveTool,
    setCalculationBusy,
    setCalculationNotice,
    setGridError,
    setVisitedTools,
    specialSelectionMode,
  });
  const { selectionSummary, summaryError } = useSelectionSummary({
    range,
    selected,
    sheet,
    specialSelection,
    storageKey,
    summarize,
    late: {
      get activeRange() {
        return activeRange;
      },
      get busy() {
        return busy;
      },
    },
  });
  const {
    activeRange,
    homeSameTarget,
    homeReadCells,
    homeFormatCommands,
    clearHyperlinksDisabledReason,
    homeClipboardText,
    homePasteAsText,
    homePaste,
    homePasteSpecial,
  } = buildHomeDataCommands({
    homeFrameRef,
    importPaste,
    internalClipboard,
    literal,
    metadata,
    pasteAbort,
    range,
    readRange,
    selected,
    setPasteProgress,
    setRange,
    sheets,
    storageKey,
  });
  const { homeHandle } = buildHomeHandle({
    conditionalThreshold,
    findBusy,
    findConfig,
    findGeneration,
    findIndex,
    findMatches,
    findMode,
    findNextCursor,
    findPageIndex,
    findPages,
    findReceipt,
    findScope,
    findText,
    findTotal,
    foundValues,
    homeClipboardText,
    homeFormatCommands,
    homePaintTarget,
    homePaste,
    homePasteAsText,
    homePasteSpecial,
    homeReadCells,
    homeSameTarget,
    internalClipboard,
    literal,
    normalFont,
    onCommandRedo,
    onCommandUndo,
    paintHomeBinding,
    paintStyleRef,
    pasteAbort,
    query,
    replaceText,
    resetFind,
    scan,
    select,
    setConditionalColor,
    setConditionalRule,
    setConditionalThreshold,
    setFindBusy,
    setFindIndex,
    setFindMatches,
    setFindMode,
    setFindNextCursor,
    setFindOpen,
    setFindPageIndex,
    setFindScope,
    setFindText,
    setFindTotal,
    setPaintStyleId,
    setRange,
    setReplaceText,
    setSpecialSelection,
    setTableName,
    sheet,
    sheets,
    storageKey,
    late: {
      get goToMatch() {
        return goToMatch;
      },
      get homeWorkbookKey() {
        return homeWorkbookKey;
      },
      get openAcross() {
        return openAcross;
      },
      get openReplacement() {
        return openReplacement;
      },
      get openSeries() {
        return openSeries;
      },
      get openSpecial() {
        return openSpecial;
      },
      get structuralBlock() {
        return structuralBlock;
      },
      get styleCell() {
        return styleCell;
      },
      get toggleTool() {
        return toggleTool;
      },
    },
  });
  const {
    homeInvoke,
    homeReason,
    homeDisabled,
    homeProps,
    homeMenu,
    copy,
    clearSelected,
    autoSum,
    find,
    goToMatch,
    paste,
  } = useHomeDispatch({
    activeFrame,
    clearHyperlinksDisabledReason,
    commandScope,
    dragging,
    fillDragging,
    formulaDraftHandle,
    homeFrameRef,
    homeHandle,
    homePaintRunner,
    homePaintTarget,
    homePending,
    homePhysical,
    navigationGeneration,
    onCommandBatch,
    onCommandRedo,
    onCommandUndo,
    readRange,
    scrollCell,
    select,
    setFillPreview,
    setGridError,
    setHomePending,
    setSheetName,
    sheet,
    sheets,
    specialSelection,
    specialSelectionMode,
    viewActive,
    late: {
      get acrossDisabledReason() {
        return acrossDisabledReason;
      },
      get homeRenderFrame() {
        return homeRenderFrame;
      },
      get seriesDisabledReason() {
        return seriesDisabledReason;
      },
      get specialDisabledReason() {
        return specialDisabledReason;
      },
    },
  });
  const { navigate } = buildKeyboardNavigation({
    activeFrame,
    activeStorage,
    calculationController,
    calculationControls,
    calculationFrame,
    cancelCellFocus,
    clearSelected,
    geometryProfile,
    gridRoot,
    homeInvoke,
    invoke,
    literal,
    navigationGeneration,
    normalMode,
    paintStyleRef,
    passiveNotice,
    readOnly,
    readRange,
    runCalculationAction,
    scrollCell,
    select,
    selected,
    selectionEnd,
    setGridError,
    setHomeNavigationVersion,
    setPaintStyleId,
    sheet,
    sheetRef,
    storageKey,
    viewActive,
    late: {
      get busy() {
        return busy;
      },
      get renderedCalculationKey() {
        return renderedCalculationKey;
      },
    },
  });
  const {
    moveAfterCommit,
    objectFamily,
    objectFamilyEpoch,
    objectAxes,
    boundaryModel,
    visibleRows,
    visibleColumns,
    worksheetBackground,
    geometry,
    inkReason,
    inkKey,
    inkLive,
    enterDraw,
    rowOffsets,
    columnOffsets,
  } = useGridGeometry({
    activationEpoch,
    active,
    activeFrame,
    busy,
    cancelCellFocus,
    commandScope,
    drawMode,
    formulaDraftHandle,
    functionViewMode,
    geometryKey,
    geometryProfile,
    gridRoot,
    inkPreset,
    literal,
    normalMode,
    objectControls,
    objectExtent,
    onCommandBatch,
    ownerIdentity,
    pendingCellFocus,
    preferenceDocumentId,
    preferenceStore,
    printPreview,
    readOnly,
    scrollCell,
    select,
    selected,
    setActiveTool,
    setDrawMode,
    sheet,
    sheetRef,
    specialSelection,
    storageKey,
    view,
    viewActive,
    viewMode,
    viewport,
    worksheetDimensions,
  });
  const {
    rowWindow,
    columnWindow,
    firstRenderedRow,
    lastRenderedRow,
    frozenRowCount,
    renderedRows,
    firstRenderedColumn,
    lastRenderedColumn,
    frozenColumnCount,
    renderedColumnIndices,
    frozenRowGap,
    frozenColumnGap,
    rangeKey,
  } = useGridWindows({
    activeRange,
    columnOffsets,
    geometry,
    gridScroll,
    internalClipboard,
    normalReturnPending,
    normalViewportBookmark,
    normalViewportKey,
    ranges,
    rowOffsets,
    selected,
    setViewport,
    sheet,
    view,
    viewMode,
    viewport,
  });
  busy = busy || calculationJob.physicalBusy || !ranges.ready(rangeKey);
  const {
    homeRenderFrame,
    filterRenderScope,
    currentFilterPage,
    filterFormGeneration,
    filterRuntimeKey,
    homeWorkbookKey,
    namedDraftKey,
    formulaPreviewKey,
    renderedCalculationKey,
  } = useRenderFrames({
    acceptedCommandState,
    activationEpoch,
    active,
    activeRange,
    busy,
    calculationControls,
    calculationFrame,
    canEditRef,
    commandAvailability,
    commandBinding,
    documentId,
    fallbackDocumentId,
    filterGeneration,
    filterPage,
    filterPageScope,
    filterScope,
    formulaPreviewControls,
    formulaPreviewFrame,
    functionViewMode,
    homeFrameRef,
    homeMetadata,
    homeSpecial,
    literal,
    namedDraftLive,
    namedRanges,
    namePickerSerial,
    navigationGeneration,
    normalMode,
    onBatch,
    onCommandBatch,
    ownerIdentity,
    paintHomeBinding,
    paintStyleRef,
    passiveNotice,
    ranges,
    readOnly,
    resetFilter,
    resetFind,
    reviewReadFrame,
    selected,
    setGridError,
    setNamePicker,
    setPaintStyleId,
    sheet,
    specialSelection,
    storageKey,
    traceFrame,
    viewMode,
  });
  if (!sheet) return <p>{words('This workbook has no worksheets.')}</p>;
  const {
    structuralBlock,
    selectedCell,
    selectedValidation,
    styleCell,
    selectedFontProfile,
    sortBounds,
    filterBounds,
    sortColumns,
    filterColumn,
    distinctFilterValues,
  } = useSelectionDerived({
    acceptedValidation,
    activeRange,
    filterValues,
    normalFont,
    selected,
    sheet,
    sheets,
    specialSelection,
    structuralReason,
    validationScope,
    words,
  });
  const { sortDisabledReason, captureSortContext, sortSelected, duplicateDisabledReason, captureDuplicateContext } =
    buildSortDuplicateContext({
      duplicatePreview,
      duplicatePreviewReceipt,
      duplicateSerial,
      duplicateWritePhysical,
      formulaDraftHandle,
      homeFrameRef,
      homeRenderFrame,
      homeSameTarget,
      liveDuplicateReceipt,
      onCommandBatch,
      readRange,
      setGridError,
      setSortPending,
      sheet,
      sortBounds,
      sortHeader,
      sortPhysical,
      sortPrimary,
      sortSerial,
      tables,
    });
  const { seriesDisabledReason, openSeries } = buildFillSeriesContext({
    cancelCellFocus,
    formulaDraftHandle,
    homeFrameRef,
    homePhysical,
    homeReadCells,
    homeRenderFrame,
    namedDraftKey,
    readRange,
    seriesLive,
    seriesParkedDraft,
    seriesSerial,
    seriesWritePhysical,
    setActiveTool,
    setFillToolMode,
    setHomePending,
    setSeriesContext,
    setVisitedTools,
    sheet,
    storageKey,
    tables,
    toolFocusGeneration,
  });
  const { acrossDisabledReason, openAcross } = useFillAcrossContext({
    acrossLive,
    acrossSerial,
    acrossWritePhysical,
    cancelCellFocus,
    formulaDraftHandle,
    homeFrameRef,
    homePhysical,
    homeRenderFrame,
    metadata,
    namedDraftKey,
    seriesParkedDraft,
    setAcrossContext,
    setActiveTool,
    setFillToolMode,
    setHomePending,
    setVisitedTools,
    toolFocusGeneration,
  });
  const { specialDisabledReason, openSpecial } = useSpecialContext({
    calculationControls,
    cancelCellFocus,
    conditionalControls,
    conditionalScope,
    documentId,
    formulaDraftHandle,
    homeFrameRef,
    homePhysical,
    homeRenderFrame,
    liveConditionalScope,
    liveValidationScope,
    metadata,
    namedDraftKey,
    ownerIdentity,
    readRange,
    scrollCell,
    selectedRef,
    setActiveTool,
    setHomePending,
    setJump,
    setRange,
    setSelected,
    setSelectionEnd,
    setSpecialContext,
    setSpecialSelection,
    setVisitedTools,
    sheet,
    specialLive,
    specialParkedDraft,
    specialSelectPhysical,
    specialSerial,
    toolFocusGeneration,
    usedRangeControls,
    validationControls,
    validationScope,
  });
  const { replacementDisabledReason, openReplacement } = buildReplaceContext({
    cancelCellFocus,
    findMode,
    findOpen,
    findScope,
    findText,
    formulaDraftHandle,
    homeFrameRef,
    homePhysical,
    homeRenderFrame,
    literal,
    namedDraftKey,
    query,
    replacementConfigurationKey,
    replacementConfigurationLive,
    replacementLive,
    replacementParkedDraft,
    replacementSerial,
    replacementWritePhysical,
    replaceText,
    scan,
    setActiveTool,
    setHomePending,
    setReplacementContext,
    setVisitedTools,
    sheet,
    sheets,
    storageKey,
    toolFocusGeneration,
    words,
  });
  const {
    filterDisabledReason,
    filterActionReason,
    invokeFilter,
    applyAcceptedFilter,
    applyFilterChoices,
    loadFilterValues,
    protectedFilter,
    applyPagedFilter,
    submitSheetAction,
  } = buildFilterActions({
    busy,
    currentFilterPage,
    distinctFilterValues,
    filterApplying,
    filterBounds,
    filterChoices,
    filterColumn,
    filterFormGeneration,
    filterGeneration,
    filterMounted,
    filterPageScope,
    filterPhysical,
    filterReadPending,
    filterRenderScope,
    filterRuntimeKey,
    filterScope,
    filterSelected,
    filterText,
    filterValues,
    filterWritePhysical,
    formulaDraftHandle,
    homeFrameRef,
    homeRenderFrame,
    homeSameTarget,
    literal,
    liveFilterDefinition,
    newSheetName,
    onCommandBatch,
    pendingSheetSelection,
    readOnly,
    resetFilter,
    send,
    setFilterApplying,
    setFilterBusy,
    setFilterNotice,
    setFilterOffset,
    setFilterPage,
    setFilterPageScope,
    setFilterReadPending,
    setGridError,
    setNewSheetName,
    setSheetAction,
    sheet,
    sheetAction,
    sheets,
    sortBounds,
  });
  const { mixedStyle, tableColumnCount, navigateMetadataRange, goToName, jumpToCell, returnNormal } =
    buildNavigationActions({
      activeFrame,
      activeRange,
      activeStorage,
      cancelCellFocus,
      firstRenderedColumn,
      frozenColumnCount,
      frozenColumnGap,
      gridRoot,
      gridScroll,
      jump,
      lastRenderedColumn,
      modeGeneration,
      navigationGeneration,
      normalFont,
      normalReturnPending,
      pendingCellFocus,
      printPreview,
      renderedColumnIndices,
      scrollCell,
      select,
      setGridError,
      setJump,
      setJumpError,
      setRange,
      setSelectionEnd,
      setSheetName,
      setViewMode,
      sheet,
      sheets,
      specialSelection,
      view,
      viewActive,
      visibleColumns,
    });
  const {
    showPageLayout,
    showPageBreak,
    openPrintFacet,
    currentPivot,
    renderedPivotKey,
    pivotReason,
    runPivotRefresh,
    runPivotUpdate,
    layoutReason,
    runPivotLayout,
    goPivotSource,
  } = buildPageAndPivotActions({
    activationEpoch,
    active,
    activeRange,
    activeTool,
    busy,
    cancelCellFocus,
    commandBinding,
    commandScope,
    documentId,
    fallbackDocumentId,
    functionViewMode,
    gridScroll,
    literal,
    modeGeneration,
    navigateMetadataRange,
    navigationGeneration,
    normalMode,
    normalViewportBookmark,
    normalViewportKey,
    onCommandBatch,
    ownerIdentity,
    passiveNotice,
    pivotFlight,
    pivotFrame,
    pivotRunning,
    pivots,
    printControls,
    printFocusSerial,
    printPreview,
    readOnly,
    returnNormal,
    setActiveTool,
    setPivotNotice,
    setPivotRunning,
    setPrintFocus,
    setViewMode,
    setVisitedTools,
    sheet,
    specialSelection,
    toolFocusGeneration,
    view,
    viewMode,
  });
  const {
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
  } = useToolFrames({
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
  });
  const {
    watchSelectionReason,
    hyperlinkFrame,
    accessibilityFrame,
    openAccessibility,
    notesFrame,
    openNotes,
    cellInsertionMode,
    cellInsertionFrame,
    booleanCellControls,
    pageStylesFrame,
    openPageStyles,
    illustrationFrame,
    openIllustrations,
    chartFrame,
    openCharts,
    visibleTool,
    closeTaskPane,
  } = buildToolPaneFrames({
    accessibilityControls,
    accessibilityLive,
    actionsParkedDraft,
    activationEpoch,
    active,
    activeFrame,
    activeRange,
    activeTool,
    booleanWritePending,
    busy,
    cancelCellFocus,
    cellInsertionControls,
    chartLive,
    chartToolsMode,
    commandScope,
    createNamesParkedDraft,
    documentId,
    duplicateParkedDraft,
    fillToolMode,
    formulaDraftHandle,
    functionViewMode,
    gridRoot,
    homeFrameRef,
    homeRenderFrame,
    hyperlinkControls,
    illustrationLive,
    illustrationMode,
    literal,
    namedDraftLive,
    navigateMetadataRange,
    navigationGeneration,
    normalMode,
    notesControls,
    notesLive,
    objectControls,
    objectFamily,
    onCommandBatch,
    ownerIdentity,
    pageStylesLive,
    pageStylesMode,
    pageTheme,
    passiveNotice,
    pendingCellFocus,
    readOnly,
    readRange,
    replacementParkedDraft,
    reviewParkedDraft,
    selected,
    seriesParkedDraft,
    setActiveTool,
    setChartToolsMode,
    setGridError,
    setIllustrationMode,
    setPageStylesMode,
    setVisitedTools,
    sheet,
    sheetRef,
    sheets,
    sortParkedDraft,
    specialParkedDraft,
    specialSelection,
    textColumnsParkedDraft,
    toolFocusGeneration,
    tools,
    viewActive,
    viewMode,
  });
  const { openActions, openQueries, openLocalExecution, openReview, toggleTool } = buildTaskPaneActions({
    actionsControls,
    actionsIntent,
    actionsParkedDraft,
    active,
    activeTool,
    busy,
    cancelCellFocus,
    captureDuplicateContext,
    captureSortContext,
    closeTaskPane,
    commandBinding,
    commandScope,
    conditionalControls,
    createNamesParkedDraft,
    documentId,
    duplicateDisabledReason,
    duplicateParkedDraft,
    fallbackDocumentId,
    formulaDraftHandle,
    functionViewMode,
    homeFrameRef,
    homeRenderFrame,
    literal,
    liveFunctionFrame,
    liveNameInventory,
    localExecutionReason,
    modeGeneration,
    namedDraftKey,
    navigationGeneration,
    normalMode,
    onCommandBatch,
    openAccessibility,
    openCharts,
    openIllustrations,
    openNotes,
    openPageStyles,
    openReplacement,
    openSeries,
    openSpecial,
    ownerBusy,
    ownerIdentity,
    passiveNotice,
    pivots,
    queriesMode,
    queryAllReason,
    queryControls,
    readOnly,
    readRange,
    reviewFrame,
    reviewParkedDraft,
    selected,
    setActionsIntent,
    setActiveTool,
    setCreateNamesContext,
    setDuplicateContext,
    setFunctionBrowse,
    setFunctionContext,
    setGridError,
    setQueriesAllIntent,
    setQueriesMode,
    setReviewIntent,
    setSortContext,
    setTextColumnContext,
    setVisitedTools,
    sheet,
    sortDisabledReason,
    sortParkedDraft,
    sortPending,
    specialSelection,
    specialSelectionMode,
    startTrace,
    storageKey,
    tables,
    textColumnsParkedDraft,
    toolFocusGeneration,
    tools,
    validationControls,
    viewActive,
  });
  const { sortFilterCommands } = buildSortFilterCommands({
    applyAcceptedFilter,
    applyFilterChoices,
    applyPagedFilter,
    currentFilterPage,
    distinctFilterValues,
    filterActionReason,
    filterApplying,
    filterBounds,
    filterBusy,
    filterChoices,
    filterColumn,
    filterDisabledReason,
    filterFormGeneration,
    filterGeneration,
    filterNotice,
    filterOffset,
    filterReadPending,
    filterRenderScope,
    filterRuntimeKey,
    filterScope,
    filterSelected,
    filterText,
    filterValues,
    homeFrameRef,
    homeRenderFrame,
    invoke,
    invokeFilter,
    literal,
    loadFilterValues,
    protectedFilter,
    selectedCell,
    setFilterBusy,
    setFilterOffset,
    setFilterPage,
    setFilterPageScope,
    setFilterSelected,
    setFilterText,
    setGridError,
    setSortHeader,
    setSortPrimary,
    sheet,
    sortBounds,
    sortColumns,
    sortDisabledReason,
    sortHeader,
    sortPending,
    sortPrimary,
    sortSelected,
    taskPaneId,
    toggleTool,
    words,
  });
  const { currentNamePicker, closeNamePicker, nameDraftReason, openNamePicker, chooseFormulaName } =
    buildNamePickerActions({
      activeFrame,
      cancelCellFocus,
      formulaComposing,
      formulaDraftHandle,
      namedDraftKey,
      namedDraftLive,
      namedRanges,
      namePicker,
      namePickerSerial,
      onCommandBatch,
      ownerIdentity,
      readRange,
      selected,
      setGridError,
      setNamePicker,
      sheet,
      sheets,
      specialSelection,
      view,
    });
  return (
    <OwnedPivotPresentation.Provider value={pivots}>
      <CalculationRevision.Provider value={sheet.calculationRevision}>
        <BooleanCellControls.Provider value={booleanCellControls}>
          <ConditionalDefinition.Provider
            value={
              conditionalControls
                ? acceptedConditional?.scope === conditionalScope
                  ? acceptedConditional.definition
                  : null
                : undefined
            }
          >
            <div
              className="spreadsheet-editor"
              data-home-pending={homePending}
              ref={gridRoot}
              onPointerDownCapture={(event) => {
                const target = event.target as HTMLElement;
                if (
                  target.closest('[data-ribbon-command^="review."],[data-ribbon-command^="actions."]') &&
                  homeFrameRef.current?.active
                ) {
                  const handle = formulaDraftHandle.current;
                  if (handle?.ownsFocus()) {
                    event.preventDefault();
                    const snapshot = handle.capture();
                    if (!snapshot.composing) {
                      handle.park();
                      if (target.closest('[data-ribbon-command^="actions."]'))
                        actionsParkedDraft.current = { key: namedDraftKey, handle, snapshot };
                      else reviewParkedDraft.current = { key: namedDraftKey, handle, snapshot };
                    }
                  }
                }
                if (
                  target.closest('[data-home-command="home.replace-preview"]') &&
                  homeFrameRef.current?.active &&
                  homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding)
                ) {
                  const handle = formulaDraftHandle.current;
                  if (handle?.ownsFocus()) {
                    event.preventDefault();
                    const snapshot = handle.capture();
                    if (!snapshot.composing) {
                      handle.park();
                      replacementParkedDraft.current = { key: namedDraftKey, handle, snapshot };
                    }
                  }
                }
                if (
                  target.closest('[data-home-command="home.fill-series"],[data-home-command="home.fill-across"]') &&
                  homeFrameRef.current?.active &&
                  homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding)
                ) {
                  const handle = formulaDraftHandle.current;
                  if (handle?.ownsFocus()) {
                    event.preventDefault();
                    const snapshot = handle.capture();
                    if (!snapshot.composing) {
                      handle.park();
                      seriesParkedDraft.current = { key: namedDraftKey, handle, snapshot };
                    }
                  }
                }
                if (
                  target.closest(
                    '[data-home-command="home.special-open"],[data-home-command="home.select-blanks"],[data-home-command="home.select-formulas"]'
                  ) &&
                  homeFrameRef.current?.active &&
                  homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding)
                ) {
                  const handle = formulaDraftHandle.current;
                  if (handle?.ownsFocus()) {
                    event.preventDefault();
                    const snapshot = handle.capture();
                    if (!snapshot.composing) {
                      handle.park();
                      specialParkedDraft.current = { key: namedDraftKey, handle, snapshot };
                    }
                  }
                }
                if (target.closest('[data-ribbon-command="names.create-from-selection"]')) event.preventDefault();
                if (
                  target.closest('[data-ribbon-command="data.text-to-columns"]') &&
                  homeFrameRef.current?.active &&
                  homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding)
                ) {
                  const handle = formulaDraftHandle.current;
                  if (handle?.ownsFocus()) {
                    event.preventDefault();
                    const snapshot = handle.capture();
                    if (!snapshot.composing) {
                      handle.park();
                      textColumnsParkedDraft.current = { key: namedDraftKey, handle, snapshot };
                    }
                  }
                }
                if (
                  target.closest(
                    '[data-home-command],[data-ribbon-command="data.sort"],[data-ribbon-command="data.filter"],[data-ribbon-command="data.remove-duplicates"]'
                  )
                ) {
                  const current = document.activeElement as HTMLElement | null,
                    handle = formulaDraftHandle.current;
                  if (handle?.ownsFocus()) {
                    if (handle.capture().composing) event.preventDefault();
                    else handle.park();
                  } else if (current?.closest('.sheet-grid') && current.getAttribute('data-editing') === 'true')
                    event.preventDefault();
                }
              }}
              onCopy={(event) => copy(event)}
              onCut={(event) => copy(event, true)}
              onPaste={paste}
              onKeyDown={navigate}
              aria-busy={busy}
            >
              <div className="sheet-a11y-status" aria-live="polite" aria-atomic="true">
                {sheet.name} {selected}.{' '}
                {specialSelection
                  ? `${specialSelection.size} matching cells selected.`
                  : range.includes(':')
                    ? `Range ${range} selected.`
                    : 'Cell selected.'}
                {selectedCell?.formula
                  ? ` Formula ${selectedCell.formula}; calculation ${selectedCell.calculation?.status ?? 'pending'}.`
                  : ''}
              </div>
              {passiveNotice && <p role="status">{passiveNotice}</p>}
              <GridRibbon
                accessibilityFrame={accessibilityFrame}
                actionsControls={actionsControls}
                actionsIntent={actionsIntent}
                active={active}
                activeFrame={activeFrame}
                activeRange={activeRange}
                activeTool={activeTool}
                autoSum={autoSum}
                busy={busy}
                calculateCurrentWorkbook={calculateCurrentWorkbook}
                calculationBusy={calculationBusy}
                calculationControls={calculationControls}
                calculationFrame={calculationFrame}
                calculationJob={calculationJob}
                calculationOwnerAvailable={calculationOwnerAvailable}
                canRedo={canRedo}
                canUndo={canUndo}
                changeView={changeView}
                chartFrame={chartFrame}
                chartToolsMode={chartToolsMode}
                chooseFormulaName={chooseFormulaName}
                closeNamePicker={closeNamePicker}
                commandScope={commandScope}
                conditionalColor={conditionalColor}
                conditionalControls={conditionalControls}
                conditionalRule={conditionalRule}
                conditionalThreshold={conditionalThreshold}
                currentNamePicker={currentNamePicker}
                currentPivot={currentPivot}
                drawMode={drawMode}
                enterDraw={enterDraw}
                fillToolMode={fillToolMode}
                functionBrowse={functionBrowse}
                geometryProfile={geometryProfile}
                goPivotSource={goPivotSource}
                gridRoot={gridRoot}
                homeDisabled={homeDisabled}
                homeFrameRef={homeFrameRef}
                homeInvoke={homeInvoke}
                homeMenu={homeMenu}
                homeProps={homeProps}
                homeReason={homeReason}
                homeRenderFrame={homeRenderFrame}
                homeSameTarget={homeSameTarget}
                illustrationFrame={illustrationFrame}
                illustrationMode={illustrationMode}
                importFile={importFile}
                importPaste={importPaste}
                inkPreferenceStore={inkPreferenceStore}
                inkPreferenceWarning={inkPreferenceWarning}
                inkPreset={inkPreset}
                inkReason={inkReason}
                literal={literal}
                localExecutionAvailable={localExecutionAvailable}
                localExecutionReason={localExecutionReason}
                mixedStyle={mixedStyle}
                nameCommands={nameCommands}
                nameDraftReason={nameDraftReason}
                namedViewPresets={namedViewPresets}
                navigationGeneration={navigationGeneration}
                normalFont={normalFont}
                normalMode={normalMode}
                notesFrame={notesFrame}
                objectControls={objectControls}
                onCommandBatch={onCommandBatch}
                onCommandResult={onCommandResult}
                openAccessibility={openAccessibility}
                openActions={openActions}
                openCharts={openCharts}
                openIllustrations={openIllustrations}
                openLocalExecution={openLocalExecution}
                openNamePicker={openNamePicker}
                openNotes={openNotes}
                openPageStyles={openPageStyles}
                openPrintFacet={openPrintFacet}
                openQueries={openQueries}
                openReview={openReview}
                ownerBusy={ownerBusy}
                ownerIdentity={ownerIdentity}
                pageStylesFrame={pageStylesFrame}
                pageStylesMode={pageStylesMode}
                paintStyleId={paintStyleId}
                passiveNotice={passiveNotice}
                pivotFrame={pivotFrame}
                pivotReason={pivotReason}
                preferenceDocumentId={preferenceDocumentId}
                preferenceStore={preferenceStore}
                presetAction={presetAction}
                presetBindingKey={presetBindingKey}
                presetDocumentId={presetDocumentId}
                presetStore={presetStore}
                printControls={printControls}
                queriesMode={queriesMode}
                queryAllReason={queryAllReason}
                queryControls={queryControls}
                readOnly={readOnly}
                readRange={readRange}
                renderedCalculationKey={renderedCalculationKey}
                renderedPivotKey={renderedPivotKey}
                returnNormal={returnNormal}
                reviewControls={reviewControls}
                reviewIntent={reviewIntent}
                runCalculationAction={runCalculationAction}
                runPivotRefresh={runPivotRefresh}
                scan={scan}
                scrollCell={scrollCell}
                selected={selected}
                sendCommand={sendCommand}
                setDrawMode={setDrawMode}
                setGridError={setGridError}
                setInkPreferenceWarning={setInkPreferenceWarning}
                setInkPreset={setInkPreset}
                setViewNotice={setViewNotice}
                sheet={sheet}
                showPageBreak={showPageBreak}
                showPageLayout={showPageLayout}
                sortFilterCommands={sortFilterCommands}
                specialSelection={specialSelection}
                startTrace={startTrace}
                structuralBlock={structuralBlock}
                styleCell={styleCell}
                tableName={tableName}
                taskPaneId={taskPaneId}
                toggleTool={toggleTool}
                tools={tools}
                traceController={traceController}
                traceControls={traceControls}
                traceFrame={traceFrame}
                traceInventoryReason={traceInventoryReason}
                view={view}
                viewMode={viewMode}
                viewport={viewport}
                visibleColumns={visibleColumns}
                visibleRows={visibleRows}
                visibleTool={visibleTool}
                words={words}
                worksheetDimensions={worksheetDimensions}
              />
              {pasteProgress && (
                <div role="status" aria-live="polite">
                  {words('Preparing paste')} {pasteProgress.completed} / {pasteProgress.total}{' '}
                  <button {...homeProps('home.cancel-paste')} onClick={() => void homeInvoke('home.cancel-paste')}>
                    {words('Cancel paste')}
                  </button>
                </div>
              )}
              <GridWorkspace
                accessibilityControls={accessibilityControls}
                accessibilityFrame={accessibilityFrame}
                acrossContext={acrossContext}
                actionsControls={actionsControls}
                actionsFrame={actionsFrame}
                actionsIntent={actionsIntent}
                activationEpoch={activationEpoch}
                active={active}
                activeFrame={activeFrame}
                activeRange={activeRange}
                activeTool={activeTool}
                boundaryModel={boundaryModel}
                busy={busy}
                calculationController={calculationController}
                calculationControls={calculationControls}
                calculationFrame={calculationFrame}
                calculationKey={calculationKey}
                calculationNotice={calculationNotice}
                cancelCellFocus={cancelCellFocus}
                canEditRef={canEditRef}
                cellInsertionControls={cellInsertionControls}
                cellInsertionFrame={cellInsertionFrame}
                cellInsertionMode={cellInsertionMode}
                chartFrame={chartFrame}
                chartToolsMode={chartToolsMode}
                closeTaskPane={closeTaskPane}
                columnOffsets={columnOffsets}
                columnWindow={columnWindow}
                commandScope={commandScope}
                conditionalControls={conditionalControls}
                conditionalScope={conditionalScope}
                createNamesContext={createNamesContext}
                currentPivot={currentPivot}
                documentId={documentId}
                dragEnter={dragEnter}
                dragStart={dragStart}
                drawMode={drawMode}
                duplicateContext={duplicateContext}
                editRef={editRef}
                errorScanner={errorScanner}
                errorScope={errorScope}
                fillDragging={fillDragging}
                fillPreview={fillPreview}
                fillToolMode={fillToolMode}
                find={find}
                findBusy={findBusy}
                findIndex={findIndex}
                findMatches={findMatches}
                findMode={findMode}
                findNextCursor={findNextCursor}
                findOpen={findOpen}
                findPageIndex={findPageIndex}
                findPages={findPages}
                findScope={findScope}
                findText={findText}
                findTotal={findTotal}
                firstRenderedColumn={firstRenderedColumn}
                firstRenderedRow={firstRenderedRow}
                formulaDraftHandle={formulaDraftHandle}
                formulaPreviewController={formulaPreviewController}
                formulaPreviewDraft={formulaPreviewDraft}
                formulaPreviewFrame={formulaPreviewFrame}
                formulaPreviewKey={formulaPreviewKey}
                frozenColumnCount={frozenColumnCount}
                frozenColumnGap={frozenColumnGap}
                frozenRowCount={frozenRowCount}
                frozenRowGap={frozenRowGap}
                functionBrowse={functionBrowse}
                functionContext={functionContext}
                functionViewMode={functionViewMode}
                geometry={geometry}
                geometryKey={geometryKey}
                geometryProfile={geometryProfile}
                goPivotSource={goPivotSource}
                goToName={goToName}
                gridError={gridError}
                gridRoot={gridRoot}
                gridScroll={gridScroll}
                homeDisabled={homeDisabled}
                homeFrameRef={homeFrameRef}
                homeInvoke={homeInvoke}
                homeProps={homeProps}
                homeRenderFrame={homeRenderFrame}
                hyperlinkControls={hyperlinkControls}
                hyperlinkFrame={hyperlinkFrame}
                illustrationFrame={illustrationFrame}
                illustrationMode={illustrationMode}
                inkKey={inkKey}
                inkLive={inkLive}
                inkPreset={inkPreset}
                inkReason={inkReason}
                invoke={invoke}
                jump={jump}
                jumpError={jumpError}
                jumpToCell={jumpToCell}
                lastRenderedColumn={lastRenderedColumn}
                lastRenderedRow={lastRenderedRow}
                layoutReason={layoutReason}
                literal={literal}
                liveConditionalScope={liveConditionalScope}
                liveValidationScope={liveValidationScope}
                localExecutionAvailable={localExecutionAvailable}
                localExecutionFrame={localExecutionFrame}
                moveAfterCommit={moveAfterCommit}
                nameCommands={nameCommands}
                namedDraftKey={namedDraftKey}
                namedRanges={namedRanges}
                namePage={namePage}
                nameScope={nameScope}
                navigateMetadataRange={navigateMetadataRange}
                navigateTrace={navigateTrace}
                navigationGeneration={navigationGeneration}
                newRangeName={newRangeName}
                newSheetName={newSheetName}
                normalFont={normalFont}
                normalMode={normalMode}
                notesControls={notesControls}
                notesFrame={notesFrame}
                objectAxes={objectAxes}
                objectControls={objectControls}
                objectFamily={objectFamily}
                objectFamilyEpoch={objectFamilyEpoch}
                onBatch={onBatch}
                onCommandBatch={onCommandBatch}
                openCharts={openCharts}
                openNamePicker={openNamePicker}
                openTracePane={openTracePane}
                ownerBusy={ownerBusy}
                ownerIdentity={ownerIdentity}
                pageStylesFrame={pageStylesFrame}
                pageStylesMode={pageStylesMode}
                pageTheme={pageTheme}
                passiveNotice={passiveNotice}
                pivotFrame={pivotFrame}
                pivotNotice={pivotNotice}
                pivotReason={pivotReason}
                pivotRunning={pivotRunning}
                pivots={pivots}
                printControls={printControls}
                printFocus={printFocus}
                printPreview={printPreview}
                queriesAllIntent={queriesAllIntent}
                queriesMode={queriesMode}
                queryAllFrame={queryAllFrame}
                queryControls={queryControls}
                range={range}
                readOnly={readOnly}
                readRange={readRange}
                renderedColumnIndices={renderedColumnIndices}
                renderedPivotKey={renderedPivotKey}
                renderedRows={renderedRows}
                replacementContext={replacementContext}
                replacementDisabledReason={replacementDisabledReason}
                replaceText={replaceText}
                reportFormulaPreviewDraft={reportFormulaPreviewDraft}
                returnNormal={returnNormal}
                reviewFrame={reviewFrame}
                reviewIntent={reviewIntent}
                rowOffsets={rowOffsets}
                rowWindow={rowWindow}
                runPivotLayout={runPivotLayout}
                runPivotRefresh={runPivotRefresh}
                runPivotUpdate={runPivotUpdate}
                scan={scan}
                scrollCell={scrollCell}
                select={select}
                selectColumn={selectColumn}
                selected={selected}
                selectedCell={selectedCell}
                selectedFontProfile={selectedFontProfile}
                selectedRef={selectedRef}
                selectedValidation={selectedValidation}
                selectionSummary={selectionSummary}
                selectRow={selectRow}
                send={send}
                seriesContext={seriesContext}
                setAcceptedConditional={setAcceptedConditional}
                setFormulaComposing={setFormulaComposing}
                setGridError={setGridError}
                setJump={setJump}
                setNamePage={setNamePage}
                setNameScope={setNameScope}
                setNewRangeName={setNewRangeName}
                setNewSheetName={setNewSheetName}
                setRange={setRange}
                setSelectionEnd={setSelectionEnd}
                setSheetAction={setSheetAction}
                setSheetName={setSheetName}
                setSortHeader={setSortHeader}
                setTablePage={setTablePage}
                setValue={setValue}
                setViewNotice={setViewNotice}
                setViewport={setViewport}
                sheet={sheet}
                sheetAction={sheetAction}
                sheets={sheets}
                sortContext={sortContext}
                specialContext={specialContext}
                specialSelection={specialSelection}
                specialSelectionMode={specialSelectionMode}
                status={status}
                storageKey={storageKey}
                submitSheetAction={submitSheetAction}
                summaryError={summaryError}
                tableColumnCount={tableColumnCount}
                tablePage={tablePage}
                tables={tables}
                taskPaneHeading={taskPaneHeading}
                taskPaneId={taskPaneId}
                taskPaneTitleId={taskPaneTitleId}
                textColumnContext={textColumnContext}
                toggleTool={toggleTool}
                toolFocusGeneration={toolFocusGeneration}
                traceController={traceController}
                traceControls={traceControls}
                traceInventoryReason={traceInventoryReason}
                traceScope={traceScope}
                updateObjectExtent={updateObjectExtent}
                validationControls={validationControls}
                validationScope={validationScope}
                view={view}
                viewActive={viewActive}
                viewMode={viewMode}
                viewNotice={viewNotice}
                viewport={viewport}
                visibleColumns={visibleColumns}
                visibleRows={visibleRows}
                visibleTool={visibleTool}
                visitedTools={visitedTools}
                watchReader={watchReader}
                watchScope={watchScope}
                watchSelectionReason={watchSelectionReason}
                words={words}
                worksheetBackground={worksheetBackground}
              />
            </div>
          </ConditionalDefinition.Provider>
        </BooleanCellControls.Provider>
      </CalculationRevision.Provider>
    </OwnedPivotPresentation.Provider>
  );
}
