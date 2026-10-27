import type { AccessibilityFrame } from '../accessibilityController';
import type { CellInsertionFrame } from '../cellInsertionController';
import type { ChartRecommendationFrame } from '../chartRecommendationsController';
import type { FormulaDraftHandle } from '../namedFormulaDraft';
import type {
  GridActionsParkedDraft,
  GridCreateNamesParkedDraft,
  GridDuplicateParkedDraft,
  GridHyperlinkControls,
  GridObjectControls3,
  GridPageTheme,
  GridReplacementParkedDraft,
  GridReviewParkedDraft,
  GridSeriesParkedDraft,
  GridSortParkedDraft,
  GridSpecialParkedDraft,
  GridTextColumnsParkedDraft,
} from './gridTypes';
import { type HomeFrame, homeBindingMatches } from '../homeCommands';
import type { HyperlinkFrame } from '../hyperlinkController';
import type { IllustrationMode } from '../WorkbookShapes';
import type { NoteFrame } from '../noteController';
import type { PageStylesFrame } from '../pageStylesController';
import type { ReadRange } from '../rangeCells';
import type { RibbonBinding, RibbonDispatchScope } from '../ribbonCommands';
import type {
  WorkbookAccessibilityControls,
  WorkbookCommand,
  WorkbookNotesControls,
  WorkbookSheet,
} from '../../core/contracts';
import type { WorkbookRibbonTool, WorkbookTool } from '../SpreadsheetRibbon';
import type { WorksheetObjectFrame } from '../worksheetObjectController';
import { booleanCellCommand } from '../../core/cell-insertions.mjs';
import { inRange, selectedBounds } from './rangeAddress';

export interface ToolPaneFramesInputs {
  accessibilityControls: WorkbookAccessibilityControls | undefined;
  accessibilityLive: React.MutableRefObject<AccessibilityFrame | null>;
  actionsParkedDraft: GridActionsParkedDraft;
  activationEpoch: React.MutableRefObject<number>;
  active: boolean;
  activeFrame: (work: () => void) => number;
  activeRange: () => string;
  activeTool: WorkbookTool | null;
  booleanWritePending: React.MutableRefObject<boolean>;
  busy: boolean;
  cancelCellFocus: () => void;
  cellInsertionControls: boolean | undefined;
  chartLive: React.MutableRefObject<ChartRecommendationFrame | null>;
  chartToolsMode: 'manage' | 'recommendations';
  commandScope: RibbonDispatchScope;
  createNamesParkedDraft: GridCreateNamesParkedDraft;
  documentId: string | undefined;
  duplicateParkedDraft: GridDuplicateParkedDraft;
  fillToolMode: 'series' | 'across';
  formulaDraftHandle: React.MutableRefObject<FormulaDraftHandle | null>;
  functionViewMode: React.MutableRefObject<{ mode: 'normal' | 'pageLayout' | 'pageBreak'; epoch: number }>;
  gridRoot: React.RefObject<HTMLDivElement>;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homeRenderFrame: HomeFrame;
  hyperlinkControls: GridHyperlinkControls;
  illustrationLive: React.MutableRefObject<WorksheetObjectFrame | null>;
  illustrationMode: IllustrationMode;
  literal: boolean | undefined;
  namedDraftLive: React.MutableRefObject<{ key: string; editable: boolean }>;
  navigateMetadataRange: (name: string, target: string) => void;
  navigationGeneration: React.MutableRefObject<number>;
  normalMode: React.MutableRefObject<boolean>;
  notesControls: WorkbookNotesControls | undefined;
  notesLive: React.MutableRefObject<NoteFrame | null>;
  objectControls: GridObjectControls3;
  objectFamily: string;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  ownerIdentity: string | undefined;
  pageStylesLive: React.MutableRefObject<PageStylesFrame | null>;
  pageStylesMode: 'themes' | 'backgrounds';
  pageTheme: GridPageTheme;
  passiveNotice: string | undefined;
  pendingCellFocus: React.MutableRefObject<number | null>;
  readOnly: boolean;
  readRange: ReadRange | undefined;
  replacementParkedDraft: GridReplacementParkedDraft;
  reviewParkedDraft: GridReviewParkedDraft;
  selected: string;
  seriesParkedDraft: GridSeriesParkedDraft;
  setActiveTool: React.Dispatch<React.SetStateAction<WorkbookTool | null>>;
  setChartToolsMode: React.Dispatch<React.SetStateAction<'manage' | 'recommendations'>>;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setIllustrationMode: React.Dispatch<React.SetStateAction<IllustrationMode>>;
  setPageStylesMode: React.Dispatch<React.SetStateAction<'themes' | 'backgrounds'>>;
  setVisitedTools: React.Dispatch<React.SetStateAction<WorkbookTool[]>>;
  sheet: WorkbookSheet;
  sheetRef: React.MutableRefObject<WorkbookSheet>;
  sheets: WorkbookSheet[];
  sortParkedDraft: GridSortParkedDraft;
  specialParkedDraft: GridSpecialParkedDraft;
  specialSelection: Set<string> | null;
  textColumnsParkedDraft: GridTextColumnsParkedDraft;
  toolFocusGeneration: React.MutableRefObject<number>;
  tools: WorkbookRibbonTool[];
  viewActive: React.MutableRefObject<boolean>;
  viewMode: 'normal' | 'pageLayout' | 'pageBreak';
}

export function buildToolPaneFrames({
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
}: ToolPaneFramesInputs) {
  let watchSelectionReason: string | undefined;
  try {
    const bounds = selectedBounds(activeRange(), 1);
    if (bounds.rows !== 1 || bounds.columns !== 1) watchSelectionReason = 'Select one cell to add a watch.';
  } catch {
    watchSelectionReason = 'Select one cell to add a watch.';
  }
  if (specialSelection) watchSelectionReason = 'Click a regular cell to leave the special selection.';
  if (sheet?.merges.some((range) => inRange(range, selected)))
    watchSelectionReason = 'Choose an unmerged cell to add a watch.';
  let hyperlinkReason: string | undefined;
  try {
    const bounds = selectedBounds(activeRange(), 1);
    if (bounds.rows !== 1 || bounds.columns !== 1) hyperlinkReason = 'Select one regular cell for a hyperlink.';
  } catch {
    hyperlinkReason = 'Select one regular cell for a hyperlink.';
  }
  if (specialSelection) hyperlinkReason = 'Click a regular cell to leave the special selection.';
  if (sheet.merges.some((range) => inRange(range, selected)))
    hyperlinkReason = 'Choose an unmerged cell for a hyperlink.';
  const hyperlinkScopeKey = JSON.stringify([
    hyperlinkControls?.receipt,
    homeRenderFrame.binding,
    viewMode,
    toolFocusGeneration.current,
  ]);
  const hyperlinkFrame: HyperlinkFrame | null = hyperlinkControls
    ? {
        key: hyperlinkScopeKey,
        family: JSON.stringify([
          hyperlinkControls.receipt.documentId,
          hyperlinkControls.receipt.ownerId,
          sheet.name,
          selected,
          activationEpoch.current,
          navigationGeneration.current,
          functionViewMode.current.epoch,
          toolFocusGeneration.current,
        ]),
        sheet: sheet.name,
        cell: selected,
        receipt: hyperlinkControls.receipt,
        active: active && activeTool === 'hyperlink' && viewMode === 'normal' && !hyperlinkReason,
        writable: !readOnly && !busy && !!onCommandBatch && !hyperlinkReason,
        reason:
          hyperlinkReason ??
          (!onCommandBatch
            ? 'Link authoring is unavailable in this host.'
            : readOnly
              ? 'This workbook is read-only.'
              : busy
                ? 'Wait for the workbook operation to finish.'
                : undefined),
        read: () => hyperlinkControls.read(sheet.name, selected),
        readCell: readRange
          ? async () => {
              const cells = await readRange(sheet.name, selected);
              if (
                !cells ||
                typeof cells !== 'object' ||
                Array.isArray(cells) ||
                Object.keys(cells).some((key) => key !== selected)
              )
                throw Error('The cell text reply changed its target.');
              return cells[selected] ?? null;
            }
          : undefined,
        apply: async (commands) => {
          if (
            !viewActive.current ||
            !normalMode.current ||
            !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current!.binding) ||
            !onCommandBatch ||
            readOnly ||
            busy ||
            hyperlinkReason
          )
            throw Error('The original hyperlink cell or workbook changed.');
          await onCommandBatch(commands);
        },
        ...(hyperlinkControls.follow
          ? { follow: (revision: string) => hyperlinkControls.follow!(sheet.name, selected, revision) }
          : {}),
        navigate: (name, cell) => navigateMetadataRange(name, cell),
      }
    : null;
  let noteSelectionReason: string | undefined;
  try {
    const bounds = selectedBounds(activeRange(), 1);
    if (bounds.rows !== 1 || bounds.columns !== 1) noteSelectionReason = 'Select one regular cell to edit a note.';
  } catch {
    noteSelectionReason = 'Select one regular cell to edit a note.';
  }
  if (specialSelection) noteSelectionReason = 'Click a regular cell before editing a note.';
  // This workbook-wide report survives ordinary cell/sheet navigation, but not owner/head or pane replacement.
  const accessibilityKey = JSON.stringify([
    accessibilityControls?.receipt,
    activationEpoch.current,
    viewMode,
    toolFocusGeneration.current,
  ]);
  const accessibilityFrame: AccessibilityFrame | null =
    accessibilityControls && !literal && !passiveNotice
      ? {
          key: accessibilityKey,
          receipt: accessibilityControls.receipt,
          active: active && activeTool === 'accessibility' && viewMode === 'normal' && !busy,
          visibleSheets: sheets.filter((s) => (s.state ?? 'visible') === 'visible').map((s) => s.name),
          read: () => accessibilityControls.read(),
          navigate: (name, cell) => {
            if (
              accessibilityLive.current?.key !== accessibilityKey ||
              !accessibilityLive.current.active ||
              !viewActive.current ||
              !normalMode.current
            )
              throw Error('The original accessibility view changed.');
            navigateMetadataRange(name, cell);
          },
        }
      : null;
  accessibilityLive.current = accessibilityFrame;
  const openAccessibility = (binding: RibbonBinding): void => {
    if (!commandScope.isCurrent(binding) || !normalMode.current || !accessibilityFrame || busy)
      throw Error('Activate this workbook in Normal view to check accessibility.');
    const focused = document.activeElement as HTMLElement | null;
    if (focused?.getAttribute('data-editing') === 'true' || formulaDraftHandle.current?.ownsFocus())
      throw Error('Finish or cancel the cell edit before checking accessibility.');
    if (activeTool === 'accessibility') {
      closeTaskPane();
      return;
    }
    cancelCellFocus();
    toolFocusGeneration.current++;
    setVisitedTools((previous) => (previous.includes('accessibility') ? previous : [...previous, 'accessibility']));
    setActiveTool('accessibility');
  };
  const notesKey = JSON.stringify([
    notesControls?.receipt,
    homeRenderFrame.binding,
    viewMode,
    toolFocusGeneration.current,
  ]);
  const notesFrame: NoteFrame | null =
    notesControls && !literal && !passiveNotice
      ? {
          key: notesKey,
          sheet: sheet.name,
          cell: selected,
          receipt: notesControls.receipt,
          active: active && activeTool === 'notes' && viewMode === 'normal',
          writable: !readOnly && !busy && !!onCommandBatch && !noteSelectionReason,
          reason:
            noteSelectionReason ??
            (readOnly
              ? 'This workbook is read-only.'
              : busy
                ? 'Wait for the workbook operation to finish.'
                : !onCommandBatch
                  ? 'Note authoring is unavailable in this host.'
                  : undefined),
          read: () => notesControls.read(sheet.name),
          apply: async (commands) => {
            if (
              notesLive.current?.key !== notesKey ||
              !notesLive.current.active ||
              !notesLive.current.writable ||
              !viewActive.current ||
              !normalMode.current ||
              !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current!.binding) ||
              !onCommandBatch
            )
              throw Error('The original note cell or workbook changed.');
            await onCommandBatch(commands);
          },
          navigate: (name, cell) => {
            if (
              notesLive.current?.key !== notesKey ||
              !notesLive.current.active ||
              !viewActive.current ||
              !normalMode.current ||
              !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current!.binding)
            )
              throw Error('The original note view changed.');
            navigateMetadataRange(name, cell);
          },
        }
      : null;
  notesLive.current = notesFrame;
  const openNotes = (binding: RibbonBinding): void => {
    if (!commandScope.isCurrent(binding) || !normalMode.current || !notesFrame)
      throw Error('Activate this workbook in Normal view to use notes.');
    const focused = document.activeElement as HTMLElement | null;
    if (focused?.getAttribute('data-editing') === 'true' || formulaDraftHandle.current?.ownsFocus())
      throw Error('Finish or cancel the cell edit before opening notes.');
    if (activeTool === 'notes') {
      closeTaskPane();
      return;
    }
    cancelCellFocus();
    toolFocusGeneration.current++;
    setVisitedTools((previous) => (previous.includes('notes') ? previous : [...previous, 'notes']));
    setActiveTool('notes');
  };
  const cellInsertionMode = activeTool === 'symbols' ? 'symbols' : 'checkbox';
  const cellInsertionKey = JSON.stringify([
    homeRenderFrame.binding,
    viewMode,
    toolFocusGeneration.current,
    cellInsertionMode,
  ]);
  const cellInsertionFrame: CellInsertionFrame | null =
    cellInsertionControls && readRange
      ? {
          key: cellInsertionKey,
          family: JSON.stringify([
            documentId,
            ownerIdentity,
            sheet.name,
            selected,
            activationEpoch.current,
            navigationGeneration.current,
            functionViewMode.current.epoch,
            toolFocusGeneration.current,
          ]),
          sheet: sheet.name,
          cell: selected,
          active:
            active && ['checkbox', 'symbols'].includes(activeTool ?? '') && viewMode === 'normal' && !hyperlinkReason,
          writable: !readOnly && !busy && !!onCommandBatch && !hyperlinkReason,
          reason:
            hyperlinkReason?.replace(/hyperlink/g, 'cell insertion') ??
            (readOnly
              ? 'This workbook is read-only.'
              : busy
                ? 'Wait for the workbook operation to finish.'
                : undefined),
          read: async () => {
            const cells = await readRange(sheet.name, selected);
            if (
              !cells ||
              typeof cells !== 'object' ||
              Array.isArray(cells) ||
              Object.keys(cells).some((key) => key !== selected)
            )
              throw Error('The selected cell reply changed its target.');
            return cells[selected] ?? null;
          },
          apply: async (commands) => {
            if (
              !viewActive.current ||
              !normalMode.current ||
              !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current!.binding) ||
              !onCommandBatch ||
              readOnly ||
              busy ||
              hyperlinkReason
            )
              throw Error('The original insertion cell or workbook changed.');
            await onCommandBatch(commands);
          },
        }
      : null;
  const booleanCellControls =
    cellInsertionControls && onCommandBatch
      ? {
          disabled: readOnly || !!busy || !active || viewMode !== 'normal',
          toggle: async (name: string, cell: string, value: boolean) => {
            if (
              booleanWritePending.current ||
              readOnly ||
              busy ||
              !viewActive.current ||
              !normalMode.current ||
              !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current!.binding) ||
              name !== sheet.name
            )
              return;
            const focused = document.activeElement as HTMLElement | null,
              handle = formulaDraftHandle.current;
            if (focused?.getAttribute('data-editing') === 'true' || handle?.ownsFocus()) {
              setGridError('Finish or cancel the cell/formula edit before toggling a checkbox.');
              return;
            }
            booleanWritePending.current = true;
            try {
              await onCommandBatch([
                booleanCellCommand({ type: 'booleanCell', action: 'toggle', sheet: name, cell, expectedValue: value }),
              ]);
            } catch (e) {
              if (viewActive.current && homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current!.binding))
                setGridError(String(e instanceof Error ? e.message : e));
            } finally {
              booleanWritePending.current = false;
            }
          },
        }
      : null;
  const pageStylesKey = JSON.stringify([
    homeRenderFrame.binding,
    viewMode,
    toolFocusGeneration.current,
    pageStylesMode,
    pageTheme,
    sheet.background,
  ]);
  const pageStylesFrame: PageStylesFrame | null =
    pageTheme && ownerIdentity && onCommandBatch && !literal && !passiveNotice
      ? {
          key: pageStylesKey,
          family: JSON.stringify([
            documentId,
            ownerIdentity,
            sheet.name,
            activationEpoch.current,
            navigationGeneration.current,
            functionViewMode.current.epoch,
            toolFocusGeneration.current,
            pageStylesMode,
          ]),
          sheet: sheet.name,
          theme: pageTheme,
          background: sheet.background,
          active: active && activeTool === 'pageStyles' && viewMode === 'normal',
          writable: !readOnly && !busy,
          reason: readOnly
            ? 'This workbook is read-only.'
            : busy
              ? 'Wait for the workbook operation to finish.'
              : undefined,
          apply: async (commands) => {
            if (
              pageStylesLive.current?.key !== pageStylesKey ||
              !pageStylesLive.current.active ||
              !pageStylesLive.current.writable ||
              !viewActive.current ||
              !normalMode.current ||
              !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current!.binding)
            )
              throw Error('The original page-style workbook changed.');
            await onCommandBatch(commands);
          },
        }
      : null;
  pageStylesLive.current = pageStylesFrame;
  const openPageStyles = (binding: RibbonBinding, mode: 'themes' | 'backgrounds'): void => {
    if (!commandScope.isCurrent(binding) || !normalMode.current || !pageStylesFrame)
      throw Error('Activate this workbook in Normal view to use page styles.');
    const focused = document.activeElement as HTMLElement | null,
      handle = formulaDraftHandle.current;
    if ((focused?.closest('.sheet-grid') && focused.getAttribute('data-editing') === 'true') || handle?.ownsFocus())
      throw Error('Finish or cancel the cell/formula edit before opening page styles.');
    if (activeTool === 'pageStyles' && pageStylesMode === mode) {
      closeTaskPane();
      return;
    }
    cancelCellFocus();
    toolFocusGeneration.current++;
    setPageStylesMode(mode);
    setVisitedTools((previous) => (previous.includes('pageStyles') ? previous : [...previous, 'pageStyles']));
    setActiveTool('pageStyles');
  };
  const illustrationKey = JSON.stringify([
    homeRenderFrame.binding,
    viewMode,
    toolFocusGeneration.current,
    illustrationMode,
    objectControls?.revision,
  ]);
  const illustrationFrame: WorksheetObjectFrame | null =
    objectControls && ownerIdentity && onCommandBatch && !literal && !passiveNotice
      ? {
          key: illustrationKey,
          family: objectFamily,
          sheet: sheet.name,
          active: active && activeTool === 'illustrations' && viewMode === 'normal',
          reason: readOnly
            ? 'This workbook is read-only.'
            : busy
              ? 'Wait for the workbook operation to finish.'
              : undefined,
          controls: {
            ...objectControls,
            onBatch: async (commands) => {
              if (
                illustrationLive.current?.key !== illustrationKey ||
                !illustrationLive.current.active ||
                illustrationLive.current.reason ||
                !viewActive.current ||
                !normalMode.current ||
                !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current!.binding)
              )
                throw Error('The original illustration workbook or selection changed.');
              await onCommandBatch(commands);
            },
          },
        }
      : null;
  illustrationLive.current = illustrationFrame;
  const openIllustrations = (binding: RibbonBinding, mode: IllustrationMode): void => {
    if (!commandScope.isCurrent(binding) || !normalMode.current || !illustrationFrame)
      throw Error('Activate this workbook in Normal view to use illustrations.');
    const focused = document.activeElement as HTMLElement | null,
      handle = formulaDraftHandle.current;
    if ((focused?.closest('.sheet-grid') && focused.getAttribute('data-editing') === 'true') || handle?.ownsFocus())
      throw Error('Finish or cancel the cell/formula edit before opening illustrations.');
    if (activeTool === 'illustrations' && illustrationMode === mode) {
      closeTaskPane();
      return;
    }
    cancelCellFocus();
    toolFocusGeneration.current++;
    setIllustrationMode(mode);
    setVisitedTools((previous) => (previous.includes('illustrations') ? previous : [...previous, 'illustrations']));
    setActiveTool('illustrations');
  };
  const chartKey = JSON.stringify([
    homeRenderFrame.binding,
    viewMode,
    toolFocusGeneration.current,
    chartToolsMode,
    objectControls?.revision,
    !!specialSelection,
  ]);
  const chartFrame: ChartRecommendationFrame | null =
    objectControls && ownerIdentity && onCommandBatch && readRange && !literal && !passiveNotice
      ? {
          key: chartKey,
          sheet: sheet.name,
          range: activeRange(),
          view: { sheets: [{ name: sheet.name, hiddenRows: [...(sheet.hiddenRows ?? [])] }] },
          active: active && activeTool === 'objects' && chartToolsMode === 'recommendations' && viewMode === 'normal',
          reason: readOnly
            ? 'This workbook is read-only.'
            : busy
              ? 'Wait for the workbook operation to finish.'
              : specialSelection
                ? 'Select one regular chart rectangle.'
                : undefined,
          read: async (name, range) => {
            if (
              chartLive.current?.key !== chartKey ||
              !chartLive.current.active ||
              specialSelection ||
              !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current!.binding)
            )
              throw Error('The original chart selection changed.');
            const cells = await readRange(name, range);
            if (
              chartLive.current?.key !== chartKey ||
              !chartLive.current.active ||
              !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current!.binding)
            )
              throw Error('The original chart selection changed.');
            return cells;
          },
          apply: async (command) => {
            if (
              chartLive.current?.key !== chartKey ||
              !chartLive.current.active ||
              chartLive.current.reason ||
              !viewActive.current ||
              !normalMode.current ||
              !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current!.binding)
            )
              throw Error('The original chart preview changed.');
            await onCommandBatch([command]);
          },
        }
      : null;
  chartLive.current = chartFrame;
  const openCharts = (binding: RibbonBinding, mode: 'manage' | 'recommendations'): void => {
    if (!commandScope.isCurrent(binding) || (mode === 'recommendations' && (!normalMode.current || !chartFrame)))
      throw Error('Activate this workbook in Normal view to preview charts.');
    const focused = document.activeElement as HTMLElement | null;
    if (focused?.getAttribute('data-editing') === 'true' || formulaDraftHandle.current?.ownsFocus())
      throw Error('Finish or cancel the cell edit before opening charts.');
    if (activeTool === 'objects' && chartToolsMode === mode) {
      closeTaskPane();
      return;
    }
    cancelCellFocus();
    toolFocusGeneration.current++;
    setChartToolsMode(mode);
    setVisitedTools((previous) => (previous.includes('objects') ? previous : [...previous, 'objects']));
    setActiveTool('objects');
  };
  const visibleTool = tools.find((tool) => tool.id === activeTool);
  const closeTaskPane = (): void => {
    const previous = activeTool;
    const generation = ++toolFocusGeneration.current;
    cancelCellFocus();
    setActiveTool(null);
    pendingCellFocus.current = activeFrame(() => {
      pendingCellFocus.current = null;
      if (generation !== toolFocusGeneration.current) return;
      const parked =
        previous === 'special'
          ? specialParkedDraft.current
          : previous === 'replace'
            ? replacementParkedDraft.current
            : previous === 'fill'
              ? seriesParkedDraft.current
              : previous === 'actions'
                ? actionsParkedDraft.current
                : previous === 'review'
                  ? reviewParkedDraft.current
                  : previous === 'createNames'
                    ? createNamesParkedDraft.current
                    : previous === 'sort'
                      ? sortParkedDraft.current
                      : previous === 'duplicates'
                        ? duplicateParkedDraft.current
                        : previous === 'textColumns'
                          ? textColumnsParkedDraft.current
                          : null;
      if (parked && namedDraftLive.current.key === parked.key) {
        const current = parked.handle.capture();
        if (
          current.text === parked.snapshot.text &&
          current.generation === parked.snapshot.generation &&
          !current.composing
        ) {
          parked.handle.focus();
          return;
        }
      }
      const selector =
        previous === 'fill'
          ? `[data-home-command="home.fill-${fillToolMode === 'across' ? 'across' : 'series'}"]`
          : `[data-workbook-tool="${previous}"]`;
      const trigger = gridRoot.current?.querySelector<HTMLButtonElement>(selector);
      // Closed details can retain rectangles while refusing focus in Chromium.
      if (
        trigger &&
        !trigger.disabled &&
        !trigger.closest('details:not([open]),[hidden],[inert]') &&
        trigger.getClientRects().length
      ) {
        trigger.focus({ preventScroll: true });
        if (document.activeElement === trigger) return;
      }
      // Focusing a single cell changes selection. Retain sparse and multi-cell
      // selections when closing their inspector.
      const target =
        specialSelection || activeRange() !== selected
          ? null
          : [...(gridRoot.current?.querySelectorAll<HTMLInputElement>('.sheet-grid input') ?? [])].find(
              (input) => input.getAttribute('aria-label') === `${sheetRef.current.name} ${selected}`
            );
      (target ?? gridRoot.current?.querySelector<HTMLInputElement>('.sheet-cell-reference'))?.focus({
        preventScroll: true,
      });
    });
  };
  return {
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
  } as const;
}
