import { COMMAND_BATCH_UNAVAILABLE } from './commandBatch';
import {
  type DuplicateContext,
  type DuplicateReceipt,
  duplicateColumns,
  duplicateReceiptKey,
} from '../duplicateSelection';
import type { FormulaDraftHandle } from '../namedFormulaDraft';
import type { GridDuplicatePreview, GridSortBounds, GridTables } from './gridTypes';
import { type HomeBinding, type HomeFrame, captureHomeBinding, homeBindingMatches } from '../homeCommands';
import type { ReadRange } from '../rangeCells';
import { type SortContext, applySelectionSort, sortSelectionColumns } from '../sortSelection';
import type { WorkbookCell, WorkbookCommand, WorkbookSheet } from '../../core/contracts';
import { colName } from '../sheetModel';
import { inRange, selectedBounds } from './rangeAddress';

export interface SortDuplicateContextInputs {
  duplicatePreview: GridDuplicatePreview;
  duplicatePreviewReceipt: DuplicateReceipt | undefined;
  duplicateSerial: React.MutableRefObject<number>;
  duplicateWritePhysical: React.MutableRefObject<boolean>;
  formulaDraftHandle: React.MutableRefObject<FormulaDraftHandle | null>;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homeRenderFrame: HomeFrame;
  homeSameTarget: (binding: HomeBinding) => boolean;
  liveDuplicateReceipt: React.MutableRefObject<DuplicateReceipt | undefined>;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  readRange: ReadRange | undefined;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setSortPending: React.Dispatch<React.SetStateAction<boolean>>;
  sheet: WorkbookSheet;
  sortBounds: GridSortBounds;
  sortHeader: boolean;
  sortPhysical: React.MutableRefObject<boolean>;
  sortPrimary: string;
  sortSerial: React.MutableRefObject<number>;
  tables: GridTables;
}

export function buildSortDuplicateContext({
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
}: SortDuplicateContextInputs) {
  const sortDisabledReason = (): string | undefined => {
    const frame = homeFrameRef.current;
    if (!frame?.active) return 'Activate this workbook before sorting.';
    if (frame.profile !== 'xlsx') return 'Sorting requires an editable XLSX workbook.';
    if (frame.binding.mode !== 'normal') return 'Return to Normal view to sort rows.';
    if (!frame.canEdit) return 'This workbook is read-only.';
    if (frame.busy) return 'Wait for the workbook operation to finish.';
    if (frame.binding.specialSelection) return 'Select a rectangular range before sorting.';
    if (!frame.binding.ownerIdentity) return 'Sorting is unavailable in this editor.';
    if (!frame.acceptBatch) return COMMAND_BATCH_UNAVAILABLE;
    if (!sortBounds || sortBounds.rows < 2) return 'Select at least two rows to sort.';
  };
  const captureSortContext = (): SortContext => {
    if (!homeFrameRef.current?.active || !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding))
      throw Error('The workbook or selection changed. Reopen Sort.');
    const reason = sortDisabledReason();
    if (reason) throw Error(reason);
    const binding = captureHomeBinding(homeRenderFrame.binding),
      reader = readRange,
      originalSheet = sheet,
      batch = onCommandBatch!;
    const columns = sortSelectionColumns(binding.selection),
      initialColumn = columns.includes(sortPrimary)
        ? sortPrimary
        : columns.includes(binding.selectedCell.replace(/[0-9]+$/, ''))
          ? binding.selectedCell.replace(/[0-9]+$/, '')
          : columns[0];
    const table = tables.find(
      (item) =>
        item.sheet === binding.sheet &&
        inRange(item.ref, binding.selection.split(':')[0]) &&
        inRange(item.ref, binding.selection.split(':').at(-1)!)
    );
    return {
      binding,
      serial: ++sortSerial.current,
      caption: `${binding.sheet}!${binding.selection}${table ? ' · ' + table.name : ''}`,
      initialColumn,
      initialHeader: sortHeader,
      isCurrent: (value) => !!homeFrameRef.current?.active && homeBindingMatches(value, homeFrameRef.current.binding),
      isRuntimeCurrent: homeSameTarget,
      disabledReason: sortDisabledReason,
      read: (range) =>
        reader
          ? reader(binding.sheet, range)
          : Promise.resolve(
              Object.fromEntries(
                (() => {
                  const b = selectedBounds(range, 1000),
                    values: Array<[string, WorkbookCell]> = [];
                  for (let col = b.first.col; col <= b.last.col; col++) {
                    const ref = `${colName(col)}${b.first.row + 1}`,
                      cell = originalSheet.cells[ref];
                    if (cell) values.push([ref, cell]);
                  }
                  return values;
                })()
              )
            ),
      accept: async (commands) => {
        if (!homeFrameRef.current?.active || !homeBindingMatches(binding, homeFrameRef.current.binding))
          throw Error('The workbook or selection changed. Reopen Sort.');
        const blocked = sortDisabledReason();
        if (blocked) throw Error(blocked);
        if (sortPhysical.current) throw Error('Wait for the current sort to finish.');
        sortPhysical.current = true;
        setSortPending(true);
        try {
          const pending = batch(commands);
          if (!pending || typeof pending.then !== 'function')
            throw Error('Sorting requires an acknowledged workbook command.');
          await pending;
        } finally {
          sortPhysical.current = false;
          setSortPending(false);
        }
      },
    };
  };
  const sortSelected = async (direction: 'asc' | 'desc'): Promise<void> => {
    if (!homeFrameRef.current?.active || !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding))
      return;
    let context: SortContext | undefined;
    try {
      context = captureSortContext();
      const focused = document.activeElement as HTMLElement | null,
        handle = formulaDraftHandle.current;
      if (
        (focused?.closest('.sheet-grid') && focused.getAttribute('data-editing') === 'true') ||
        (handle?.ownsFocus() && handle.capture().composing)
      ) {
        setGridError('Finish or cancel the current input before sorting.');
        return;
      }
      if (handle?.ownsFocus()) handle.park();
      await applySelectionSort(context, sortHeader, [{ column: context.initialColumn, direction }]);
      if (context.isRuntimeCurrent(context.binding)) setGridError('');
    } catch (error) {
      if (
        context?.isCurrent(context.binding) ||
        (!context &&
          homeFrameRef.current?.active &&
          homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding))
      )
        setGridError(error instanceof Error ? error.message : String(error));
    }
  };
  const duplicateDisabledReason = (action: 'preview' | 'remove' = 'preview'): string | undefined => {
    const frame = homeFrameRef.current;
    if (!frame?.active) return 'Activate this workbook before comparing rows.';
    if (frame.profile !== 'xlsx') return 'Duplicate removal requires an XLSX workbook.';
    if (frame.binding.mode !== 'normal') return 'Return to Normal view to compare rows.';
    if (frame.busy) return 'Wait for the workbook operation to finish.';
    if (frame.binding.specialSelection) return 'Select a rectangular range to compare.';
    if (!frame.binding.ownerIdentity || !duplicatePreview || !duplicatePreviewReceipt)
      return 'Duplicate preview is unavailable in this editor.';
    try {
      duplicateColumns(frame.binding.selection);
    } catch (error) {
      return error instanceof Error ? error.message : String(error);
    }
    if (action === 'remove' && !frame.canEdit) return 'This workbook is read-only.';
    if (action === 'remove' && !frame.acceptBatch) return COMMAND_BATCH_UNAVAILABLE;
  };
  const captureDuplicateContext = (): DuplicateContext => {
    if (!homeFrameRef.current?.active || !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding))
      throw Error('The workbook or selection changed. Reopen Remove Duplicates.');
    const reason = duplicateDisabledReason();
    if (reason) throw Error(reason);
    const binding = captureHomeBinding(homeRenderFrame.binding),
      receipt = structuredClone(duplicatePreviewReceipt!),
      receiptKey = duplicateReceiptKey(receipt),
      reader = duplicatePreview!,
      headers = readRange,
      originalSheet = sheet,
      batch = onCommandBatch;
    const isCurrent = () =>
      !!homeFrameRef.current?.active &&
      homeBindingMatches(binding, homeFrameRef.current.binding) &&
      duplicateReceiptKey(liveDuplicateReceipt.current) === receiptKey;
    const runtimeCurrent = () =>
      homeSameTarget(binding) &&
      liveDuplicateReceipt.current?.documentId === receipt.documentId &&
      liveDuplicateReceipt.current.ownerId === receipt.ownerId;
    return {
      binding,
      receipt,
      serial: ++duplicateSerial.current,
      caption: `${binding.sheet}!${binding.selection}`,
      initialHeader: true,
      isCurrent,
      isRuntimeCurrent: runtimeCurrent,
      disabledReason: duplicateDisabledReason,
      read: async (config, offset) => {
        if (!isCurrent()) throw Error('The workbook or selection changed. Reopen Remove Duplicates.');
        const result = await reader(config.sheet, config.range, config.hasHeaders, [...config.columns], offset);
        if (!isCurrent()) throw Error('The workbook or selection changed. Reopen Remove Duplicates.');
        return result;
      },
      readHeaders: (range) =>
        headers
          ? headers(binding.sheet, range)
          : Promise.resolve(
              Object.fromEntries(
                (() => {
                  const b = selectedBounds(range, 100),
                    values: Array<[string, WorkbookCell]> = [];
                  for (let col = b.first.col; col <= b.last.col; col++) {
                    const ref = `${colName(col)}${b.first.row + 1}`,
                      cell = originalSheet.cells[ref];
                    if (cell) values.push([ref, cell]);
                  }
                  return values;
                })()
              )
            ),
      accept: async (commands) => {
        if (!isCurrent()) throw Error('The workbook or selection changed. Reopen Remove Duplicates.');
        const blocked = duplicateDisabledReason('remove');
        if (blocked) throw Error(blocked);
        if (duplicateWritePhysical.current) throw Error('Wait for the current duplicate removal to finish.');
        duplicateWritePhysical.current = true;
        try {
          const pending = batch!(commands);
          if (!pending || typeof pending.then !== 'function')
            throw Error('Removing duplicates requires an acknowledged workbook command.');
          await pending;
        } finally {
          duplicateWritePhysical.current = false;
        }
      },
    };
  };
  return {
    sortDisabledReason,
    captureSortContext,
    sortSelected,
    duplicateDisabledReason,
    captureDuplicateContext,
  } as const;
}
