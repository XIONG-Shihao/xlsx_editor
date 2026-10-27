import { COMMAND_BATCH_UNAVAILABLE } from './commandBatch';
import type { FillAcrossContext } from '../WorkbookFillAcross';
import { type FillAcrossPreview, captureFillAcrossSheets, createFillAcrossPreview } from '../fillAcrossPreview';
import type { FormulaDraftHandle } from '../namedFormulaDraft';
import type { GridSeriesParkedDraft } from './gridTypes';
import { type HomeFrame, captureHomeBinding, dispatchHomeCommand, homeBindingMatches } from '../homeCommands';
import type { WorkbookSheet } from '../../core/contracts';
import type { WorkbookTool } from '../SpreadsheetRibbon';
import { selectedBounds } from './rangeAddress';
import { useRef } from 'react';

export interface FillAcrossContextInputs {
  acrossLive: React.MutableRefObject<{ key: string; visible: boolean }>;
  acrossSerial: React.MutableRefObject<number>;
  acrossWritePhysical: React.MutableRefObject<boolean>;
  cancelCellFocus: () => void;
  formulaDraftHandle: React.MutableRefObject<FormulaDraftHandle | null>;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homePhysical: React.MutableRefObject<boolean>;
  homeRenderFrame: HomeFrame;
  metadata: WorkbookSheet[];
  namedDraftKey: string;
  seriesParkedDraft: GridSeriesParkedDraft;
  setAcrossContext: React.Dispatch<React.SetStateAction<FillAcrossContext | null>>;
  setActiveTool: React.Dispatch<React.SetStateAction<WorkbookTool | null>>;
  setFillToolMode: React.Dispatch<React.SetStateAction<'series' | 'across'>>;
  setHomePending: React.Dispatch<React.SetStateAction<boolean>>;
  setVisitedTools: React.Dispatch<React.SetStateAction<WorkbookTool[]>>;
  toolFocusGeneration: React.MutableRefObject<number>;
}

export function useFillAcrossContext({
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
}: FillAcrossContextInputs) {
  const acrossSheetsKey = JSON.stringify(metadata.map((item) => [item.name, item.state ?? 'visible']));
  const acrossSheetsLive = useRef(acrossSheetsKey);
  acrossSheetsLive.current = acrossSheetsKey;
  const acrossDisabledReason = (): string | undefined => {
    const frame = homeFrameRef.current;
    if (!frame?.active) return 'Activate this workbook before filling across worksheets.';
    if (frame.profile !== 'xlsx') return 'Across Worksheets requires an XLSX workbook.';
    if (frame.binding.mode !== 'normal' || frame.binding.specialSelection)
      return 'Select a regular range in Normal view.';
    if (!frame.canEdit) return 'This workbook is read-only.';
    if (frame.busy || homePhysical.current) return 'Wait for the current workbook operation to finish.';
    if (!frame.binding.ownerIdentity) return 'The original worksheet owner is unavailable.';
    if (!frame.acceptBatch) return COMMAND_BATCH_UNAVAILABLE;
    try {
      selectedBounds(frame.binding.selection, 10000);
      const captured = captureFillAcrossSheets(metadata, frame.binding.sheet);
      if (!captured.some((item) => item.name !== frame.binding.sheet && item.state === 'visible'))
        return 'Add another visible worksheet before filling across worksheets.';
    } catch (error) {
      return error instanceof Error ? error.message : String(error);
    }
  };
  const captureAcrossContext = (): FillAcrossContext => {
    if (!homeFrameRef.current?.active || !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding))
      throw Error('The workbook or selection changed. Reopen Across Worksheets.');
    const blocked = acrossDisabledReason();
    if (blocked) throw Error(blocked);
    const binding = captureHomeBinding(homeRenderFrame.binding),
      frame: HomeFrame = { ...homeRenderFrame, binding },
      key = JSON.stringify([++acrossSerial.current, binding]),
      inventory = captureFillAcrossSheets(metadata, binding.sheet),
      inventoryKey = acrossSheetsKey;
    const scope = { captureFrame: () => frame, currentFrame: () => homeFrameRef.current };
    const current = () =>
      acrossLive.current.visible &&
      acrossLive.current.key === key &&
      acrossSheetsLive.current === inventoryKey &&
      !!homeFrameRef.current?.active &&
      homeFrameRef.current.profile === 'xlsx' &&
      homeBindingMatches(binding, homeFrameRef.current.binding);
    const runtime = () =>
      acrossLive.current.visible &&
      acrossLive.current.key === key &&
      acrossSheetsLive.current === inventoryKey &&
      !!homeFrameRef.current?.active &&
      homeFrameRef.current.profile === 'xlsx' &&
      homeBindingMatches(binding, homeFrameRef.current.binding, true);
    let prepared: FillAcrossPreview | null = null;
    return {
      key,
      binding,
      caption: `${binding.sheet}!${binding.selection}`,
      sheets: inventory,
      isCurrent: current,
      isRuntimeCurrent: runtime,
      disabledReason: acrossDisabledReason,
      preview: async (options, signal) => {
        prepared = null;
        const result = await dispatchHomeCommand(
          'home.fill-across-preview',
          scope,
          (_lease, args) => {
            if (!current() || signal.aborted)
              throw Error('The workbook or selection changed. Reopen Across Worksheets.');
            return createFillAcrossPreview(binding, inventory, args);
          },
          options
        );
        if (result.outcome === 'failed') throw result.error;
        if (result.outcome === 'refused') throw Error(result.reason);
        if (result.outcome !== 'completed' || !result.current || !current() || signal.aborted)
          throw Error('The workbook or selection changed. Reopen Across Worksheets.');
        prepared = result.value;
        return prepared;
      },
      apply: async (receipt) => {
        if (!current() || !prepared || JSON.stringify(receipt) !== JSON.stringify(prepared))
          throw Error('Preview the current destination worksheets before applying it.');
        const blocked = acrossDisabledReason();
        if (blocked) throw Error(blocked);
        if (acrossWritePhysical.current || homePhysical.current)
          throw Error('Wait for the current Home command to finish.');
        const captured = prepared;
        acrossWritePhysical.current = true;
        homePhysical.current = true;
        setHomePending(true);
        try {
          const result = await dispatchHomeCommand(
            'home.fill-across-apply',
            scope,
            async (lease, args) => {
              if (!current()) throw Error('The workbook or selection changed. Reopen Across Worksheets.');
              const fresh = createFillAcrossPreview(binding, inventory, args);
              if (JSON.stringify(fresh) !== JSON.stringify(captured))
                throw Error('The destination choices changed. Preview them again.');
              await lease.accept([...captured.commands]);
            },
            { targets: captured.targets, mode: captured.mode }
          );
          if (result.outcome === 'failed') throw result.error;
          if (result.outcome === 'refused') throw Error(result.reason);
          if (result.outcome !== 'completed')
            throw Error('The workbook or selection changed. Reopen Across Worksheets.');
          prepared = null;
        } finally {
          acrossWritePhysical.current = false;
          homePhysical.current = false;
          setHomePending(false);
        }
      },
    };
  };
  const openAcross = (): void => {
    if (!homeFrameRef.current?.active || !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding))
      throw Error('The workbook or selection changed. Reopen Across Worksheets.');
    const reason = acrossDisabledReason();
    if (reason) throw Error(reason);
    const handle = formulaDraftHandle.current;
    if (handle && (handle.ownsFocus() || handle.isParked?.())) {
      const snapshot = handle.capture();
      if (snapshot.composing) throw Error('Finish the current input before opening Across Worksheets.');
      handle.park();
      seriesParkedDraft.current = { key: namedDraftKey, handle, snapshot };
    } else seriesParkedDraft.current = null;
    const context = captureAcrossContext();
    cancelCellFocus();
    toolFocusGeneration.current++;
    setFillToolMode('across');
    setAcrossContext(context);
    setVisitedTools((previous) => (previous.includes('fill') ? previous : [...previous, 'fill']));
    setActiveTool('fill');
  };
  return { acrossDisabledReason, openAcross } as const;
}
