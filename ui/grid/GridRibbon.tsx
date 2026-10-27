import { COMMAND_BATCH_UNAVAILABLE } from './commandBatch';
import { type CalculationAction, calculationDisabledReason } from '../calculationJobs';
import { HomeAlignmentGroup } from './home/HomeAlignmentGroup';
import { HomeCellsGroup } from './home/HomeCellsGroup';
import { HomeClipboardGroup } from './home/HomeClipboardGroup';
import { HomeEditingGroup } from './home/HomeEditingGroup';
import { HomeFontGroup } from './home/HomeFontGroup';
import { HomeNumberGroup } from './home/HomeNumberGroup';
import { HomeStylesGroup } from './home/HomeStylesGroup';
import type { InkPenId } from '../../core/ink-stroke.mjs';
import { type RibbonBinding, type RibbonCommandId, ribbonFunctionCategories } from '../ribbonCommands';
import { SpreadsheetRibbon } from '../SpreadsheetRibbon';
import { WorkbookNamePicker } from '../WorkbookNamePicker';
import { addressPosition } from '../sheetModel';
import { createViewGeometry, zoomToSelection } from '../viewGeometry';
import { selectedBounds } from './rangeAddress';
import type { GridRibbonInputs } from './gridRibbonInputs';

export function GridRibbon({
  accessibilityFrame,
  actionsControls,
  actionsIntent,
  active,
  activeFrame,
  activeRange,
  activeTool,
  autoSum,
  busy,
  calculateCurrentWorkbook,
  calculationBusy,
  calculationControls,
  calculationFrame,
  calculationJob,
  calculationOwnerAvailable,
  canRedo,
  canUndo,
  changeView,
  chartFrame,
  chartToolsMode,
  chooseFormulaName,
  closeNamePicker,
  commandScope,
  conditionalColor,
  conditionalControls,
  conditionalRule,
  conditionalThreshold,
  currentNamePicker,
  currentPivot,
  drawMode,
  enterDraw,
  fillToolMode,
  functionBrowse,
  geometryProfile,
  goPivotSource,
  gridRoot,
  homeDisabled,
  homeFrameRef,
  homeInvoke,
  homeMenu,
  homeProps,
  homeReason,
  homeRenderFrame,
  homeSameTarget,
  illustrationFrame,
  illustrationMode,
  importFile,
  importPaste,
  inkPreferenceStore,
  inkPreferenceWarning,
  inkPreset,
  inkReason,
  literal,
  localExecutionAvailable,
  localExecutionReason,
  mixedStyle,
  nameCommands,
  nameDraftReason,
  namedViewPresets,
  navigationGeneration,
  normalFont,
  normalMode,
  notesFrame,
  objectControls,
  onCommandBatch,
  onCommandResult,
  openAccessibility,
  openActions,
  openCharts,
  openIllustrations,
  openLocalExecution,
  openNamePicker,
  openNotes,
  openPageStyles,
  openPrintFacet,
  openQueries,
  openReview,
  ownerBusy,
  ownerIdentity,
  pageStylesFrame,
  pageStylesMode,
  paintStyleId,
  passiveNotice,
  pivotFrame,
  pivotReason,
  preferenceDocumentId,
  preferenceStore,
  presetAction,
  presetBindingKey,
  presetDocumentId,
  presetStore,
  printControls,
  queriesMode,
  queryAllReason,
  queryControls,
  readOnly,
  readRange,
  renderedCalculationKey,
  renderedPivotKey,
  returnNormal,
  reviewControls,
  reviewIntent,
  runCalculationAction,
  runPivotRefresh,
  scan,
  scrollCell,
  selected,
  sendCommand,
  setDrawMode,
  setGridError,
  setInkPreferenceWarning,
  setInkPreset,
  setViewNotice,
  sheet,
  showPageBreak,
  showPageLayout,
  sortFilterCommands,
  specialSelection,
  startTrace,
  structuralBlock,
  styleCell,
  tableName,
  taskPaneId,
  toggleTool,
  tools,
  traceController,
  traceControls,
  traceFrame,
  traceInventoryReason,
  view,
  viewMode,
  viewport,
  visibleColumns,
  visibleRows,
  visibleTool,
  words,
  worksheetDimensions,
}: GridRibbonInputs): JSX.Element {
  return (
    <SpreadsheetRibbon
      canUndo={!homeDisabled('home.undo') && !!canUndo}
      canRedo={!homeDisabled('home.redo') && !!canRedo}
      onUndo={() => void homeInvoke('home.undo')}
      onRedo={() => void homeInvoke('home.redo')}
      busy={busy}
      tools={tools.filter(
        (tool) =>
          tool.id !== 'conditional' &&
          tool.id !== 'sort' &&
          tool.id !== 'fill' &&
          tool.id !== 'replace' &&
          tool.id !== 'special'
      )}
      activeTool={visibleTool?.id}
      onToolToggle={toggleTool}
      taskPaneId={taskPaneId}
      commandScope={commandScope}
      onCommandResult={onCommandResult}
      onTabChange={(tab) => {
        if (tab !== 'Draw') setDrawMode(null);
      }}
      ink={
        objectControls && !literal && !passiveNotice
          ? {
              selected: inkPreset,
              disabled: !!inkReason,
              reason: inkReason,
              warning: inkPreferenceWarning,
              onChoose: (binding, args) => {
                enterDraw(binding);
                const pen = (args as { pen: InkPenId }).pen;
                setInkPreset(inkPreferenceStore.write(pen));
                setInkPreferenceWarning(inkPreferenceStore.warning);
              },
            }
          : undefined
      }
      dataCommands={sortFilterCommands}
      commands={[
        ...(objectControls && !literal && !passiveNotice
          ? [
              {
                id: 'draw.pen' as const,
                pressed: drawMode === 'draw',
                disabled: !!inkReason,
                reason: inkReason,
                handler: (binding: RibbonBinding) => enterDraw(binding),
              },
              {
                id: 'draw.select-strokes' as const,
                pressed: drawMode === 'select',
                disabled: !!inkReason,
                reason: inkReason,
                handler: (binding: RibbonBinding) => enterDraw(binding, 'select'),
              },
              {
                id: 'draw.eraser' as const,
                pressed: drawMode === 'erase',
                disabled: !!inkReason,
                reason: inkReason,
                handler: (binding: RibbonBinding) => enterDraw(binding, 'erase'),
              },
              { id: 'draw.select-cells' as const, pressed: !drawMode, handler: () => setDrawMode(null) },
            ]
          : []),
        ...(currentPivot && pivotFrame.current
          ? [
              {
                id: 'pivots.refresh' as const,
                disabled: !!pivotReason,
                reason: pivotReason,
                handler: () => runPivotRefresh(renderedPivotKey!),
              },
              {
                id: 'pivots.go-source' as const,
                disabled: !currentPivot.sourceSheet || !currentPivot.sourceRange,
                reason: !currentPivot.sourceSheet ? 'This pivot source is unavailable.' : undefined,
                handler: () => goPivotSource(renderedPivotKey!),
              },
            ]
          : []),
        ...(objectControls
          ? [
              {
                id: 'objects.manage' as const,
                pressed: activeTool === 'objects' && chartToolsMode === 'manage',
                handler: (binding: RibbonBinding) => openCharts(binding, 'manage'),
              },
            ]
          : []),
        ...(chartFrame
          ? [
              {
                id: 'insert.recommended-charts' as const,
                pressed: activeTool === 'objects' && chartToolsMode === 'recommendations',
                disabled: !normalMode.current,
                reason: !normalMode.current ? 'Return to Normal view to preview charts.' : undefined,
                handler: (binding: RibbonBinding) => openCharts(binding, 'recommendations'),
              },
            ]
          : []),
        ...(illustrationFrame
          ? (['shape', 'text', 'icon', 'wordart', 'smartart', 'model3d'] as const).map((mode) => ({
              id: (mode === 'shape'
                ? 'insert.shapes'
                : mode === 'icon'
                  ? 'insert.icons'
                  : mode === 'wordart'
                    ? 'insert.wordart'
                    : mode === 'smartart'
                      ? 'insert.smartart'
                      : mode === 'model3d'
                        ? 'insert.model3d'
                        : 'insert.text-box') as RibbonCommandId,
              pressed: activeTool === 'illustrations' && illustrationMode === mode,
              disabled: !normalMode.current,
              reason: !normalMode.current ? 'Return to Normal view to use illustrations.' : undefined,
              handler: (binding: RibbonBinding) => openIllustrations(binding, mode),
            }))
          : []),
        ...(pageStylesFrame
          ? (['themes', 'backgrounds'] as const).map((mode) => ({
              id: (mode === 'themes' ? 'page.themes' : 'page.background') as RibbonCommandId,
              pressed: activeTool === 'pageStyles' && pageStylesMode === mode,
              disabled: !normalMode.current,
              reason: !normalMode.current ? 'Return to Normal view to use page styles.' : undefined,
              handler: (binding: RibbonBinding) => openPageStyles(binding, mode),
            }))
          : []),
        ...(localExecutionAvailable
          ? (['unhide-all', 'freeze-selection'] as const).map((action) => ({
              id: ('actions.' + action) as RibbonCommandId,
              pressed: activeTool === 'actions' && actionsIntent.mode === 'run' && actionsIntent.action === action,
              disabled: !!localExecutionReason,
              reason: localExecutionReason,
              handler: (binding: RibbonBinding) => openLocalExecution(binding, action),
            }))
          : []),
        ...(queryControls
          ? [
              {
                id: 'data.queries' as const,
                pressed: activeTool === 'queries' && queriesMode === 'individual',
                disabled: !normalMode.current,
                reason: !normalMode.current ? 'Return to Normal view to use queries.' : undefined,
                handler: (binding: RibbonBinding) => openQueries(binding, 'individual'),
              },
            ]
          : []),
        ...(queryControls?.refreshAll
          ? [
              {
                id: 'data.refresh-all' as const,
                pressed: activeTool === 'queries' && queriesMode === 'all',
                disabled: !!queryAllReason || !!ownerBusy,
                reason: queryAllReason ?? (ownerBusy ? 'Wait for the workbook operation to finish.' : undefined),
                handler: (binding: RibbonBinding) => openQueries(binding, 'all'),
              },
            ]
          : []),
        ...(actionsControls
          ? (['new', 'library'] as const).map((mode) => ({
              id: ('actions.' + mode) as RibbonCommandId,
              pressed: activeTool === 'actions' && actionsIntent.mode === (mode === 'new' ? 'new' : 'list'),
              disabled: !normalMode.current,
              reason: !normalMode.current ? 'Return to Normal view to manage local actions.' : undefined,
              handler: (binding: RibbonBinding) => openActions(binding, mode === 'new' ? 'new' : 'list'),
            }))
          : []),
        ...(accessibilityFrame
          ? [
              {
                id: 'review.accessibility' as const,
                pressed: activeTool === 'accessibility',
                disabled: !normalMode.current || !!busy,
                reason: !normalMode.current
                  ? 'Return to Normal view to check accessibility.'
                  : busy
                    ? 'Wait for the workbook operation to finish.'
                    : undefined,
                handler: (binding: RibbonBinding) => openAccessibility(binding),
              },
            ]
          : []),
        ...(notesFrame
          ? (['insert.note', 'review.notes'] as const).map((id) => ({
              id,
              pressed: activeTool === 'notes',
              disabled: !normalMode.current,
              reason: !normalMode.current ? 'Return to Normal view to use notes.' : undefined,
              handler: (binding: RibbonBinding) => openNotes(binding),
            }))
          : []),
        ...(reviewControls && scan && readRange && ownerIdentity && !literal && !passiveNotice
          ? (['statistics', 'spelling'] as const).map((mode) => ({
              id: ('review.' + mode) as RibbonCommandId,
              pressed: activeTool === 'review' && reviewIntent.mode === mode,
              disabled: !normalMode.current || !!specialSelection,
              reason: !normalMode.current
                ? 'Return to Normal view to review the workbook.'
                : specialSelection
                  ? 'Click a regular cell before opening Review.'
                  : undefined,
              handler: (binding: RibbonBinding) => openReview(binding, mode),
            }))
          : []),
        ...(!literal && !passiveNotice
          ? [
              {
                id: 'names.use-in-formula' as const,
                pressed: null,
                disabled: !!nameDraftReason,
                reason: nameDraftReason,
                handler: openNamePicker,
                popover: currentNamePicker ? (
                  <WorkbookNamePicker
                    choices={currentNamePicker.choices}
                    target={currentNamePicker.target}
                    words={words}
                    onChoose={chooseFormulaName}
                    onClose={closeNamePicker}
                  />
                ) : null,
              },
            ]
          : []),
        ...(!literal && !passiveNotice
          ? calculationControls
            ? [
                {
                  id: 'formulas.calculate-now' as const,
                  action: { type: 'calculateAllOpen' } as CalculationAction,
                  hint: 'Calculate all open editable workbooks (F9). Each accepted workbook has its own Undo step.',
                },
                {
                  id: 'formulas.calculate-sheet' as const,
                  action: { type: 'calculateSheet', sheet: sheet.name } as CalculationAction,
                  hint: 'Calculate the active worksheet and required inputs (Shift+F9). This creates one Undo step.',
                },
              ].map((item) => ({
                id: item.id,
                pressed: null,
                disabled:
                  calculationJob.physicalBusy || !!calculationDisabledReason(calculationFrame.current, item.action),
                reason: calculationJob.physicalBusy
                  ? 'Wait for the current calculation to finish.'
                  : calculationDisabledReason(calculationFrame.current, item.action),
                hint: 'hint' in item ? item.hint : undefined,
                handler: (binding: RibbonBinding) => runCalculationAction(binding, item.action, renderedCalculationKey),
              }))
            : [
                {
                  id: 'formulas.calculate-now' as const,
                  pressed: null,
                  disabled: !calculationOwnerAvailable || calculationBusy || !!specialSelection,
                  reason: !calculationOwnerAvailable
                    ? 'Calculation is unavailable in this editor.'
                    : calculationBusy
                      ? 'Wait for the current calculation to finish.'
                      : specialSelection
                        ? 'Click a regular cell before calculating this workbook.'
                        : undefined,
                  handler: (binding: RibbonBinding) => calculateCurrentWorkbook(binding),
                },
              ]
          : []),
        ...(traceControls && !literal && !passiveNotice
          ? [
              {
                id: 'formulas.trace-precedents' as const,
                pressed: null,
                disabled: !!traceInventoryReason,
                reason: traceInventoryReason,
                handler: (binding: RibbonBinding) => startTrace(binding, 'precedents'),
              },
              {
                id: 'formulas.trace-dependents' as const,
                pressed: null,
                disabled: !!traceInventoryReason,
                reason: traceInventoryReason,
                handler: (binding: RibbonBinding) => startTrace(binding, 'dependents'),
              },
              {
                id: 'formulas.remove-arrows' as const,
                handler: () => traceController.clear(traceFrame.current?.scope),
              },
            ]
          : []),
        ...(printControls
          ? [
              {
                id: 'view.page-break-preview' as const,
                pressed: viewMode === 'pageBreak',
                handler: showPageBreak,
              },
              {
                id: 'view.page-layout' as const,
                pressed: viewMode === 'pageLayout',
                handler: showPageLayout,
              },
              ...(['margins', 'orientation', 'paper', 'area', 'titles', 'scale'] as const).map((facet) => ({
                id: ('print.' + facet) as RibbonCommandId,
                handler: (binding: RibbonBinding) => openPrintFacet(facet, binding),
              })),
            ]
          : []),
        ...(!literal
          ? [{ id: 'formulas.auto-sum' as const, handler: (binding: RibbonBinding) => autoSum(binding) }]
          : []),
        ...(!literal && !passiveNotice
          ? [
              {
                id: 'formulas.recently-used' as const,
                pressed: activeTool === 'functions' && functionBrowse.mode === 'recent',
                handler: (binding: RibbonBinding) => toggleTool('functions', binding, 'recent'),
              },
              ...Object.entries(ribbonFunctionCategories).map(([id, category]) => ({
                id: id as RibbonCommandId,
                pressed: activeTool === 'functions' && functionBrowse.mode === category,
                handler: (binding: RibbonBinding) => toggleTool('functions', binding, category),
              })),
            ]
          : []),
        ...(importFile && !literal
          ? [
              {
                id: 'data.import-file' as const,
                handler: (binding: RibbonBinding) => importFile(binding.sheet, binding.selection),
              },
            ]
          : []),
        ...(nameCommands && !literal
          ? [
              {
                id: 'names.define' as const,
                handler: () => {
                  if (activeTool !== 'names') toggleTool('names');
                  activeFrame(() =>
                    gridRoot.current
                      ?.querySelector<HTMLInputElement>('.sheet-name-manager input')
                      ?.focus({ preventScroll: true })
                  );
                },
              },
            ]
          : []),
      ]}
      view={{
        ...view,
        presets: {
          key: presetBindingKey,
          active: active && !busy,
          preferences: view,
          named: namedViewPresets,
          warning: presetStore.warning ?? preferenceStore.warning,
          apply: changeView,
          save: (name) => presetAction(() => presetStore.create(presetDocumentId, sheet.name, name, view)),
          update: (name) => presetAction(() => presetStore.update(presetDocumentId, sheet.name, name, view)),
          remove: (name) => presetAction(() => presetStore.delete(presetDocumentId, sheet.name, name)),
          reset: () => presetAction(() => preferenceStore.reset(preferenceDocumentId, sheet.name)),
        },
        mode: viewMode,
        onNormal: returnNormal,
        onChange: changeView,
        onZoomToSelection: () => {
          try {
            const navigation = navigationGeneration.current;
            const bounds = selectedBounds(activeRange(), 16384 * 1048576);
            const fitting = createViewGeometry(sheet, {
              rows: Math.max(visibleRows, bounds.last.row + 1),
              columns: Math.max(visibleColumns, bounds.last.col + 1),
              zoom: view.zoom,
              headings: view.headings,
              geometryProfile,
              literal,
            });
            const percent = zoomToSelection(fitting, bounds, viewport);
            const width =
              (((fitting.rowHeaderWidth +
                fitting.frozenWidth +
                Math.max(
                  0,
                  fitting.columnOffsets[bounds.last.col + 1] -
                    Math.max(fitting.columnOffsets[bounds.first.col], fitting.frozenWidth)
                )) /
                fitting.scale) *
                percent) /
              100;
            const height =
              (((fitting.columnHeaderHeight +
                fitting.frozenHeight +
                Math.max(
                  0,
                  fitting.rowOffsets[bounds.last.row + 1] -
                    Math.max(fitting.rowOffsets[bounds.first.row], fitting.frozenHeight)
                )) /
                fitting.scale) *
                percent) /
              100;
            changeView({ zoom: percent });
            setGridError('');
            setViewNotice(
              width > viewport.width + 1 || height > viewport.height + 1
                ? 'The selection is larger than the window at the minimum zoom of 10%.'
                : ''
            );
            activeFrame(() => {
              if (navigationGeneration.current !== navigation) return;
              scrollCell(sheet, bounds.first.row, bounds.first.col, true);
            });
          } catch (error) {
            setGridError(String(error));
          }
        },
        canFreeze: !readOnly && !busy && !literal && !!onCommandBatch,
        freezeReason: !onCommandBatch ? COMMAND_BATCH_UNAVAILABLE : undefined,
        onFreeze: (mode) => {
          const position = addressPosition(selected);
          return sendCommand([
            {
              sheet: sheet.name,
              operation: 'freeze',
              rows: mode === 'selection' ? position.row : mode === 'row' ? 1 : 0,
              columns: mode === 'selection' ? position.col : mode === 'column' ? 1 : 0,
            },
          ]);
        },
      }}
    >
      <HomeClipboardGroup
        busy={busy}
        homeDisabled={homeDisabled}
        homeInvoke={homeInvoke}
        homeMenu={homeMenu}
        homeProps={homeProps}
        homeReason={homeReason}
        importPaste={importPaste}
        literal={literal}
        paintStyleId={paintStyleId}
        readOnly={readOnly}
        words={words}
      />
      <HomeFontGroup
        busy={busy}
        homeDisabled={homeDisabled}
        homeInvoke={homeInvoke}
        homeProps={homeProps}
        homeReason={homeReason}
        literal={literal}
        mixedStyle={mixedStyle}
        normalFont={normalFont}
        readOnly={readOnly}
        styleCell={styleCell}
        words={words}
      />
      <HomeAlignmentGroup
        activeRange={activeRange}
        busy={busy}
        homeDisabled={homeDisabled}
        homeInvoke={homeInvoke}
        homeMenu={homeMenu}
        homeProps={homeProps}
        homeReason={homeReason}
        literal={literal}
        mixedStyle={mixedStyle}
        readOnly={readOnly}
        sheet={sheet}
        styleCell={styleCell}
        words={words}
      />
      <HomeNumberGroup
        busy={busy}
        homeDisabled={homeDisabled}
        homeInvoke={homeInvoke}
        homeProps={homeProps}
        homeReason={homeReason}
        literal={literal}
        mixedStyle={mixedStyle}
        readOnly={readOnly}
        styleCell={styleCell}
        words={words}
      />
      <HomeStylesGroup
        active={active}
        activeTool={activeTool}
        busy={busy}
        conditionalColor={conditionalColor}
        conditionalControls={conditionalControls}
        conditionalRule={conditionalRule}
        conditionalThreshold={conditionalThreshold}
        homeDisabled={homeDisabled}
        homeInvoke={homeInvoke}
        homeMenu={homeMenu}
        homeProps={homeProps}
        literal={literal}
        readOnly={readOnly}
        tableName={tableName}
        taskPaneId={taskPaneId}
        words={words}
      />
      <HomeCellsGroup
        busy={busy}
        homeDisabled={homeDisabled}
        homeFrameRef={homeFrameRef}
        homeInvoke={homeInvoke}
        homeMenu={homeMenu}
        homeProps={homeProps}
        homeRenderFrame={homeRenderFrame}
        homeSameTarget={homeSameTarget}
        literal={literal}
        readOnly={readOnly}
        selected={selected}
        setGridError={setGridError}
        sheet={sheet}
        structuralBlock={structuralBlock}
        words={words}
        worksheetDimensions={worksheetDimensions}
      />
      <HomeEditingGroup
        activeTool={activeTool}
        busy={busy}
        fillToolMode={fillToolMode}
        homeDisabled={homeDisabled}
        homeInvoke={homeInvoke}
        homeMenu={homeMenu}
        homeProps={homeProps}
        literal={literal}
        readOnly={readOnly}
        specialSelection={specialSelection}
        taskPaneId={taskPaneId}
        words={words}
      />
    </SpreadsheetRibbon>
  );
}
