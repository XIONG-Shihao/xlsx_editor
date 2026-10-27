import { type AcceptedPrintPreview, isPrintPreviewRetired } from '../useAcceptedPrintPreview';
import { BULK_CELL_LIMIT } from './gridLimits';
import type { GridPrintControls3 } from './gridTypes';
import {
  type PivotContextFrame,
  type PivotFieldDraft,
  pivotContextKey,
  pivotRefreshReason,
  refreshPivotContext,
  selectedPivot,
  updatePivotContext,
} from '../pivotSelection';
import { type PivotLayoutDraft, pivotLayoutReason, setPivotLayoutContext } from '../pivotLayout';
import type { PrintFacet, PrintFocusRequest } from '../WorkbookPrint';
import type { RibbonBinding, RibbonDispatchScope } from '../ribbonCommands';
import type { ViewPreferences } from '../viewPreferences';
import type { WorkbookCommand, WorkbookPivotDefinition, WorkbookSheet } from '../../core/contracts';
import type { WorkbookTool } from '../SpreadsheetRibbon';
import { pageBoundaryModel } from '../pageBoundaries';
import { selectedBounds } from './rangeAddress';

export interface PageAndPivotActionsInputs {
  activationEpoch: React.MutableRefObject<number>;
  active: boolean;
  activeRange: () => string;
  activeTool: WorkbookTool | null;
  busy: boolean;
  cancelCellFocus: () => void;
  commandBinding: React.MutableRefObject<RibbonBinding | null>;
  commandScope: RibbonDispatchScope;
  documentId: string | undefined;
  fallbackDocumentId: string;
  functionViewMode: React.MutableRefObject<{ mode: 'normal' | 'pageLayout' | 'pageBreak'; epoch: number }>;
  gridScroll: React.RefObject<HTMLDivElement>;
  literal: boolean | undefined;
  modeGeneration: React.MutableRefObject<number>;
  navigateMetadataRange: (name: string, target: string) => void;
  navigationGeneration: React.MutableRefObject<number>;
  normalMode: React.MutableRefObject<boolean>;
  normalViewportBookmark: React.MutableRefObject<{ key: string; left: number; top: number; zoom: number } | null>;
  normalViewportKey: string;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  ownerIdentity: string | undefined;
  passiveNotice: string | undefined;
  pivotFlight: React.MutableRefObject<boolean>;
  pivotFrame: React.MutableRefObject<PivotContextFrame | null>;
  pivotRunning: boolean;
  pivots: WorkbookPivotDefinition[];
  printControls: GridPrintControls3;
  printFocusSerial: React.MutableRefObject<number>;
  printPreview: AcceptedPrintPreview;
  readOnly: boolean;
  returnNormal: () => void;
  setActiveTool: React.Dispatch<React.SetStateAction<WorkbookTool | null>>;
  setPivotNotice: React.Dispatch<React.SetStateAction<{ key: string; text: string } | null>>;
  setPivotRunning: React.Dispatch<React.SetStateAction<boolean>>;
  setPrintFocus: React.Dispatch<React.SetStateAction<PrintFocusRequest | null>>;
  setViewMode: React.Dispatch<React.SetStateAction<'normal' | 'pageLayout' | 'pageBreak'>>;
  setVisitedTools: React.Dispatch<React.SetStateAction<WorkbookTool[]>>;
  sheet: WorkbookSheet;
  specialSelection: Set<string> | null;
  toolFocusGeneration: React.MutableRefObject<number>;
  view: ViewPreferences;
  viewMode: 'normal' | 'pageLayout' | 'pageBreak';
}

export function buildPageAndPivotActions({
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
}: PageAndPivotActionsInputs) {
  const showPageMode = async (binding: RibbonBinding, mode: 'pageLayout' | 'pageBreak'): Promise<void> => {
    if (!printControls) throw Error('Page Layout preview is unavailable in this editor.');
    const generation = ++modeGeneration.current,
      navigation = navigationGeneration.current;
    const current = () =>
      generation === modeGeneration.current &&
      navigation === navigationGeneration.current &&
      commandScope.isCurrent(binding);
    try {
      if (commandScope.availability('view.page-layout').busy) throw Error('Wait for the current print operation.');
      const definition = await printControls.definition(binding.sheet);
      if (!current()) return;
      if (commandScope.availability('view.page-layout').busy) throw Error('Wait for the current print operation.');
      if (!definition.canEdit || !definition.settings)
        throw Error(definition.protectedReasons.join('; ') || 'These print settings are protected.');
      const saved = !!definition.settings.area;
      if (!saved) selectedBounds(binding.selection, BULK_CELL_LIMIT);
      const settings = { ...definition.settings, area: definition.settings.area ?? binding.selection };
      const options = {
        guard: current,
        origin: { kind: saved ? ('saved' as const) : ('selection' as const), range: settings.area },
      };
      const job =
        printPreview.reuseCurrent(settings, options) ??
        (await printPreview.prepare(settings, crypto.randomUUID(), options));
      if (mode === 'pageBreak') pageBoundaryModel(job);
      if (current()) {
        if (viewMode === 'normal' && gridScroll.current)
          normalViewportBookmark.current = {
            key: normalViewportKey,
            left: gridScroll.current.scrollLeft,
            top: gridScroll.current.scrollTop,
            zoom: view.zoom,
          };
        cancelCellFocus();
        setActiveTool(null);
        setViewMode(mode);
      }
    } catch (error) {
      if (!isPrintPreviewRetired(error) && current()) throw error;
    }
  };
  const showPageLayout = (binding: RibbonBinding) => showPageMode(binding, 'pageLayout');
  const showPageBreak = (binding: RibbonBinding) => showPageMode(binding, 'pageBreak');
  const openPrintFacet = (facet: PrintFacet, binding: RibbonBinding): void => {
    if (!printControls || !commandScope.isCurrent(binding)) return;
    returnNormal();
    const navigation = navigationGeneration.current;
    setVisitedTools((previous) => (previous.includes('print') ? previous : [...previous, 'print']));
    setActiveTool('print');
    setPrintFocus({
      facet,
      serial: ++printFocusSerial.current,
      documentId: printControls.documentId,
      revision: printControls.revision,
      sheet: binding.sheet,
      guard: () => commandScope.isCurrent(binding) && navigationGeneration.current === navigation,
    });
  };
  const currentPivot = selectedPivot(
    pivots,
    sheet.name,
    activeRange(),
    active && normalMode.current && !literal && !passiveNotice && !specialSelection
  );
  pivotFrame.current = currentPivot
    ? {
        key: pivotContextKey(
          [
            documentId ?? fallbackDocumentId,
            ownerIdentity ?? '',
            commandBinding.current!.state,
            sheet.name,
            activeRange(),
            navigationGeneration.current,
            activationEpoch.current,
            functionViewMode.current.epoch,
            activeTool,
            toolFocusGeneration.current,
          ],
          currentPivot
        ),
        pivot: structuredClone(currentPivot),
        active,
        normal: normalMode.current,
        readOnly,
        busy: !!busy || pivotRunning,
        accept: onCommandBatch,
      }
    : null;
  const renderedPivotKey = pivotFrame.current?.key;
  const pivotReason = pivotRefreshReason(pivotFrame.current);
  const runPivotRefresh = async (key: string): Promise<void> => {
    const captured = pivotFrame.current;
    if (!captured || captured.key !== key) return;
    if (pivotFlight.current) return;
    pivotFlight.current = true;
    setPivotRunning(true);
    setPivotNotice(null);
    try {
      const result = await refreshPivotContext(captured, () => pivotFrame.current);
      if (result.current) setPivotNotice({ key, text: 'Pivot refreshed.' });
    } catch (error) {
      if (pivotFrame.current?.key === key)
        setPivotNotice({ key, text: error instanceof Error ? error.message : String(error) });
    } finally {
      pivotFlight.current = false;
      setPivotRunning(false);
    }
  };
  const runPivotUpdate = async (key: string, draft: PivotFieldDraft): Promise<void> => {
    const captured = pivotFrame.current;
    if (!captured || captured.key !== key || pivotFlight.current) return;
    pivotFlight.current = true;
    setPivotRunning(true);
    setPivotNotice(null);
    try {
      const result = await updatePivotContext(captured, () => pivotFrame.current, draft);
      if (result.current) setPivotNotice({ key, text: 'Pivot fields updated.' });
    } catch (error) {
      if (pivotFrame.current?.key === key)
        setPivotNotice({ key, text: error instanceof Error ? error.message : String(error) });
    } finally {
      pivotFlight.current = false;
      setPivotRunning(false);
    }
  };
  const layoutReason = pivotLayoutReason(pivotFrame.current);
  const runPivotLayout = async (key: string, draft: PivotLayoutDraft): Promise<void> => {
    const captured = pivotFrame.current;
    if (!captured || captured.key !== key || pivotFlight.current) return;
    pivotFlight.current = true;
    setPivotRunning(true);
    setPivotNotice(null);
    try {
      const result = await setPivotLayoutContext(captured, () => pivotFrame.current, draft);
      if (result.current) setPivotNotice({ key, text: 'Pivot layout applied.' });
    } catch (error) {
      if (pivotFrame.current?.key === key)
        setPivotNotice({ key, text: error instanceof Error ? error.message : String(error) });
    } finally {
      pivotFlight.current = false;
      setPivotRunning(false);
    }
  };
  const goPivotSource = (key: string): void => {
    const captured = pivotFrame.current;
    if (!captured || captured.key !== key || !captured.pivot.sourceSheet || !captured.pivot.sourceRange) return;
    navigateMetadataRange(captured.pivot.sourceSheet, captured.pivot.sourceRange);
  };
  return {
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
  } as const;
}
