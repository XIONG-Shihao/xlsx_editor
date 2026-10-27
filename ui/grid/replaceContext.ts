import { COMMAND_BATCH_UNAVAILABLE } from './commandBatch';
import type { FormulaDraftHandle } from '../namedFormulaDraft';
import type { GridReplacementParkedDraft, GridScan } from './gridTypes';
import {
  type HomeFrame,
  type HomeCommandLease as HomeLease,
  captureHomeBinding,
  dispatchHomeCommand,
  homeBindingMatches,
} from '../homeCommands';
import { type ReplaceConfiguration, type ReplacePreview, collectReplacementPreview } from '../replacePreview';
import type { ReplaceContext } from '../WorkbookReplace';
import { type SearchQuery, scanSearchQuery } from '../searchPages';
import type { WorkbookSheet } from '../../core/contracts';
import type { WorkbookTool } from '../SpreadsheetRibbon';
import { displayValue } from '../sheetModel';

export interface ReplaceContextInputs {
  cancelCellFocus: () => void;
  findMode: 'display' | 'raw' | 'formula';
  findOpen: boolean;
  findScope: 'sheet' | 'workbook';
  findText: string;
  formulaDraftHandle: React.MutableRefObject<FormulaDraftHandle | null>;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homePhysical: React.MutableRefObject<boolean>;
  homeRenderFrame: HomeFrame;
  literal: boolean | undefined;
  namedDraftKey: string;
  query: SearchQuery | undefined;
  replacementConfigurationKey: string;
  replacementConfigurationLive: React.MutableRefObject<string>;
  replacementLive: React.MutableRefObject<{ key: string; visible: boolean }>;
  replacementParkedDraft: GridReplacementParkedDraft;
  replacementSerial: React.MutableRefObject<number>;
  replacementWritePhysical: React.MutableRefObject<boolean>;
  replaceText: string;
  scan: GridScan;
  setActiveTool: React.Dispatch<React.SetStateAction<WorkbookTool | null>>;
  setHomePending: React.Dispatch<React.SetStateAction<boolean>>;
  setReplacementContext: React.Dispatch<React.SetStateAction<ReplaceContext | null>>;
  setVisitedTools: React.Dispatch<React.SetStateAction<WorkbookTool[]>>;
  sheet: WorkbookSheet;
  sheets: WorkbookSheet[];
  storageKey: string | undefined;
  toolFocusGeneration: React.MutableRefObject<number>;
  words: (en: string) => string;
}

export function buildReplaceContext({
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
}: ReplaceContextInputs) {
  const replacementDisabledReason = (action: 'preview' | 'apply' = 'preview'): string | undefined => {
    const frame = homeFrameRef.current;
    if (!frame?.active) return 'Activate this workbook before previewing replacements.';
    if (frame.profile === 'passive') return 'Replacement requires an XLSX or delimited workbook.';
    if (frame.binding.mode !== 'normal' || frame.binding.specialSelection)
      return 'Select a regular range in Normal view.';
    if (storageKey && !query && !scan) return 'The complete original search reader is unavailable.';
    if (frame.busy) return 'Wait for the workbook operation to finish.';
    if (!findOpen || !findText) return 'Enter the text to find before previewing replacements.';
    if (action === 'apply' && !frame.canEdit) return 'This workbook is read-only.';
    if (action === 'apply' && !frame.acceptBatch) return COMMAND_BATCH_UNAVAILABLE;
  };
  const captureReplacementContext = (): ReplaceContext => {
    if (!homeFrameRef.current?.active || !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding))
      throw Error('The workbook or search changed. Reopen replacement preview.');
    const reason = replacementDisabledReason();
    if (reason) throw Error(reason);
    const binding = captureHomeBinding(homeRenderFrame.binding),
      frame: HomeFrame = { ...homeRenderFrame, binding },
      key = JSON.stringify([++replacementSerial.current, binding, replacementConfigurationKey]);
    const configuration: ReplaceConfiguration = Object.freeze({
      text: findText,
      replacement: replaceText,
      mode: findMode,
      sheets: Object.freeze((findScope === 'sheet' ? [sheet] : sheets).map((item) => item.name)),
      literal: !!literal,
      date1904: !!sheet.date1904,
    });
    const configurationKey = replacementConfigurationKey,
      originalSheets = sheets,
      originalQuery =
        query ??
        scanSearchQuery({
          scan,
          cells: (name) => originalSheets.find((item) => item.name === name)?.cells ?? {},
          content: (name, ref, value, mode) =>
            mode === 'formula'
              ? value.formula
                ? '=' + value.formula
                : ''
              : mode === 'raw'
                ? String(value.rawValue ?? value.value ?? '')
                : displayValue({ ...originalSheets.find((item) => item.name === name)!, cells: { [ref]: value } }, ref),
        });
    const scope = { captureFrame: () => frame, currentFrame: () => homeFrameRef.current };
    const current = () =>
      replacementLive.current.visible &&
      replacementLive.current.key === key &&
      replacementConfigurationLive.current === configurationKey &&
      !!homeFrameRef.current?.active &&
      homeFrameRef.current.profile === frame.profile &&
      homeBindingMatches(binding, homeFrameRef.current.binding);
    const runtime = () =>
      replacementLive.current.visible &&
      replacementLive.current.key === key &&
      replacementConfigurationLive.current === configurationKey &&
      !!homeFrameRef.current?.active &&
      homeFrameRef.current.profile === frame.profile &&
      homeBindingMatches(binding, homeFrameRef.current.binding, true);
    let prepared: ReplacePreview | null = null;
    const read = (lease: HomeLease, signal?: AbortSignal): Promise<ReplacePreview> =>
      collectReplacementPreview({
        binding,
        configuration,
        query: (...args) => lease.read(() => originalQuery(...args)),
        isCurrent: () => current() && lease.isCurrent(),
        ...(signal ? { signal } : {}),
        yieldWork: () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())),
      });
    return {
      key,
      binding,
      caption: findScope === 'sheet' ? sheet.name : words('All sheets'),
      isCurrent: current,
      isRuntimeCurrent: runtime,
      disabledReason: replacementDisabledReason,
      preview: async (signal) => {
        prepared = null;
        const result = await dispatchHomeCommand('home.replace-preview', scope, (lease) => read(lease, signal));
        if (result.outcome === 'failed') throw result.error;
        if (result.outcome === 'refused') throw Error(result.reason);
        if (result.outcome !== 'completed' || !result.current || !current() || signal.aborted)
          throw Error('The workbook or search changed. Reopen replacement preview.');
        prepared = result.value;
        return prepared;
      },
      apply: async (receipt) => {
        if (!current() || !prepared || !prepared.editable || JSON.stringify(receipt) !== JSON.stringify(prepared))
          throw Error('Preview all current replacements before applying them.');
        const blocked = replacementDisabledReason('apply');
        if (blocked) throw Error(blocked);
        if (replacementWritePhysical.current || homePhysical.current)
          throw Error('Wait for the current Home command to finish.');
        const captured = prepared;
        replacementWritePhysical.current = true;
        homePhysical.current = true;
        setHomePending(true);
        try {
          const result = await dispatchHomeCommand('home.replace-all', scope, async (lease) => {
            const fresh = await read(lease);
            if (!current() || JSON.stringify(fresh) !== JSON.stringify(captured))
              throw Error('The matching cells changed. Preview all replacements again.');
            await lease.accept([...captured.commands]);
          });
          if (result.outcome === 'failed') throw result.error;
          if (result.outcome === 'refused') throw Error(result.reason);
          if (result.outcome !== 'completed')
            throw Error('The workbook or search changed. Reopen replacement preview.');
          prepared = null;
        } finally {
          replacementWritePhysical.current = false;
          homePhysical.current = false;
          setHomePending(false);
        }
      },
    };
  };
  const openReplacement = (): void => {
    const reason = replacementDisabledReason();
    if (reason) throw Error(reason);
    const handle = formulaDraftHandle.current;
    if (handle && (handle.ownsFocus() || handle.isParked?.())) {
      const snapshot = handle.capture();
      if (snapshot.composing) throw Error('Finish the current input before previewing replacements.');
      handle.park();
      replacementParkedDraft.current = { key: namedDraftKey, handle, snapshot };
    } else replacementParkedDraft.current = null;
    const context = captureReplacementContext();
    cancelCellFocus();
    toolFocusGeneration.current++;
    setReplacementContext(context);
    setVisitedTools((previous) => (previous.includes('replace') ? previous : [...previous, 'replace']));
    setActiveTool('replace');
  };
  return { replacementDisabledReason, openReplacement } as const;
}
