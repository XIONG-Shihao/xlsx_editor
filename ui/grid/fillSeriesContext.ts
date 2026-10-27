import { COMMAND_BATCH_UNAVAILABLE } from './commandBatch';
import type { FillSeriesContext } from '../WorkbookFill';
import {
  type FillSeriesKind,
  type FillSeriesPreview,
  createFillSeriesPreview,
  fillSeriesCommands,
} from '../fillSeriesPreview';
import type { FormulaDraftHandle } from '../namedFormulaDraft';
import type { GridHomeReadCells, GridSeriesParkedDraft, GridTables } from './gridTypes';
import {
  type HomeFrame,
  type HomeCommandLease as HomeLease,
  captureHomeBinding,
  dispatchHomeCommand,
  homeBindingMatches,
} from '../homeCommands';
import type { ReadRange } from '../rangeCells';
import type { WorkbookSheet } from '../../core/contracts';
import type { WorkbookTool } from '../SpreadsheetRibbon';

export interface FillSeriesContextInputs {
  cancelCellFocus: () => void;
  formulaDraftHandle: React.MutableRefObject<FormulaDraftHandle | null>;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homePhysical: React.MutableRefObject<boolean>;
  homeReadCells: GridHomeReadCells;
  homeRenderFrame: HomeFrame;
  namedDraftKey: string;
  readRange: ReadRange | undefined;
  seriesLive: React.MutableRefObject<{ key: string; visible: boolean }>;
  seriesParkedDraft: GridSeriesParkedDraft;
  seriesSerial: React.MutableRefObject<number>;
  seriesWritePhysical: React.MutableRefObject<boolean>;
  setActiveTool: React.Dispatch<React.SetStateAction<WorkbookTool | null>>;
  setFillToolMode: React.Dispatch<React.SetStateAction<'series' | 'across'>>;
  setHomePending: React.Dispatch<React.SetStateAction<boolean>>;
  setSeriesContext: React.Dispatch<React.SetStateAction<FillSeriesContext | null>>;
  setVisitedTools: React.Dispatch<React.SetStateAction<WorkbookTool[]>>;
  sheet: WorkbookSheet;
  storageKey: string | undefined;
  tables: GridTables;
  toolFocusGeneration: React.MutableRefObject<number>;
}

export function buildFillSeriesContext({
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
}: FillSeriesContextInputs) {
  const seriesDisabledReason = (action: 'preview' | 'apply' = 'preview'): string | undefined => {
    const frame = homeFrameRef.current;
    if (!frame?.active) return 'Activate this workbook before previewing a series.';
    if (frame.profile !== 'xlsx') return 'Fill series requires an XLSX workbook.';
    if (frame.binding.mode !== 'normal' || frame.binding.specialSelection)
      return 'Select a regular range in Normal view.';
    if (!frame.binding.ownerIdentity || (storageKey && !readRange))
      return 'The original worksheet reader is unavailable.';
    if (frame.busy) return 'Wait for the workbook operation to finish.';
    if (action === 'apply' && !frame.canEdit) return 'This workbook is read-only.';
    if (action === 'apply' && !frame.acceptBatch) return COMMAND_BATCH_UNAVAILABLE;
  };
  const captureSeriesContext = (): FillSeriesContext => {
    if (!homeFrameRef.current?.active || !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding))
      throw Error('The workbook or selection changed. Reopen Fill series.');
    const reason = seriesDisabledReason();
    if (reason) throw Error(reason);
    const binding = captureHomeBinding(homeRenderFrame.binding),
      frame: HomeFrame = { ...homeRenderFrame, binding },
      key = JSON.stringify([++seriesSerial.current, binding]);
    const originalSheet = structuredClone({ ...sheet, cells: {} }),
      originalTables = structuredClone(tables.filter((table) => table.sheet === binding.sheet));
    const scope = { captureFrame: () => frame, currentFrame: () => homeFrameRef.current };
    const current = () =>
      seriesLive.current.visible &&
      seriesLive.current.key === key &&
      !!homeFrameRef.current?.active &&
      homeFrameRef.current.profile === 'xlsx' &&
      homeBindingMatches(binding, homeFrameRef.current.binding);
    const runtime = () =>
      seriesLive.current.visible &&
      seriesLive.current.key === key &&
      !!homeFrameRef.current?.active &&
      homeFrameRef.current.profile === 'xlsx' &&
      homeBindingMatches(binding, homeFrameRef.current.binding, true);
    let prepared: FillSeriesPreview | null = null;
    const read = async (
      lease: HomeLease,
      signal?: AbortSignal,
      kind: FillSeriesKind = 'number'
    ): Promise<FillSeriesPreview> => {
      if (!current() || signal?.aborted) throw Error('The workbook or selection changed. Reopen Fill series.');
      const cells = await homeReadCells(lease, binding.sheet, binding.selection, signal);
      if (!current() || signal?.aborted) throw Error('The workbook or selection changed. Reopen Fill series.');
      const result = createFillSeriesPreview({
        binding,
        sheet: originalSheet,
        cells,
        readRange: binding.selection,
        tables: originalTables,
        kind,
      });
      if (result.status === 'refused') throw Error(result.reason);
      return result.preview;
    };
    return {
      key,
      binding,
      caption: `${binding.sheet}!${binding.selection}`,
      isCurrent: current,
      isRuntimeCurrent: runtime,
      disabledReason: seriesDisabledReason,
      preview: async (signal, kind = 'number') => {
        prepared = null;
        const result = await dispatchHomeCommand('home.fill-series', scope, (lease) => read(lease, signal, kind));
        if (result.outcome === 'failed') throw result.error;
        if (result.outcome === 'refused') throw Error(result.reason);
        if (result.outcome !== 'completed' || !result.current || !current() || signal.aborted)
          throw Error('The workbook or selection changed. Reopen Fill series.');
        prepared = result.value;
        return prepared;
      },
      apply: async (receipt) => {
        if (
          !current() ||
          !prepared ||
          JSON.stringify(receipt) !== JSON.stringify(prepared) ||
          prepared.seeds.some((seed, index) => !Object.is(seed.value, receipt.seeds[index]?.value))
        )
          throw Error('Preview the current selected series before applying it.');
        const blocked = seriesDisabledReason('apply');
        if (blocked) throw Error(blocked);
        if (seriesWritePhysical.current || homePhysical.current)
          throw Error('Wait for the current Home command to finish.');
        const captured = prepared;
        seriesWritePhysical.current = true;
        homePhysical.current = true;
        setHomePending(true);
        try {
          const result = await dispatchHomeCommand('home.fill-series-apply', scope, async (lease) => {
            const fresh = await read(lease, undefined, captured.date ? 'date' : 'number');
            if (!current() || JSON.stringify(fresh) !== JSON.stringify(captured))
              throw Error('The selected cells changed. Preview the series again.');
            await lease.accept(fillSeriesCommands(captured));
          });
          if (result.outcome === 'failed') throw result.error;
          if (result.outcome === 'refused') throw Error(result.reason);
          if (result.outcome !== 'completed') throw Error('The workbook or selection changed. Reopen Fill series.');
          prepared = null;
        } finally {
          seriesWritePhysical.current = false;
          homePhysical.current = false;
          setHomePending(false);
        }
      },
    };
  };
  const openSeries = (): void => {
    if (!homeFrameRef.current?.active || !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding))
      throw Error('The workbook or selection changed. Reopen Fill series.');
    const reason = seriesDisabledReason();
    if (reason) throw Error(reason);
    const handle = formulaDraftHandle.current;
    if (handle && (handle.ownsFocus() || handle.isParked?.())) {
      const snapshot = handle.capture();
      if (snapshot.composing) throw Error('Finish the current input before opening Fill series.');
      handle.park();
      seriesParkedDraft.current = { key: namedDraftKey, handle, snapshot };
    } else seriesParkedDraft.current = null;
    const context = captureSeriesContext();
    cancelCellFocus();
    toolFocusGeneration.current++;
    setFillToolMode('series');
    setSeriesContext(context);
    setVisitedTools((previous) => (previous.includes('fill') ? previous : [...previous, 'fill']));
    setActiveTool('fill');
  };
  return { seriesDisabledReason, openSeries } as const;
}
