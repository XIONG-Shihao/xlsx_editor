import {
  DEFAULT_SPECIAL_SELECTION_OPTIONS,
  type SpecialSelectionKind,
  type SpecialSelectionOptions,
  type SpecialSelectionPreview,
  type SpecialSelectionRuleReceipt,
  collectSpecialSelectionPreview,
  specialSelectionNeedsUsedRange,
} from '../specialSelection';
import type { FormulaDraftHandle } from '../namedFormulaDraft';
import type { GridConditionalControls2, GridSpecialParkedDraft, GridValidationControls4 } from './gridTypes';
import {
  type HomeFrame,
  type HomeCommandLease as HomeLease,
  captureHomeBinding,
  dispatchHomeCommand,
  homeBindingMatches,
} from '../homeCommands';
import type { ReadRange } from '../rangeCells';
import type { SpecialSelectionContext } from '../WorkbookSpecialSelection';
import type { WorkbookCalculationControls } from '../calculationJobs';
import type {
  WorkbookConditionalDefinition,
  WorkbookSheet,
  WorkbookUsedRangeControls,
  WorkbookValidationDefinition,
} from '../../core/contracts';
import type { WorkbookTool } from '../SpreadsheetRibbon';
import { addressPosition } from '../sheetModel';
import { flushSync } from 'react-dom';
import { selectedBounds } from './rangeAddress';
import { useRef } from 'react';
import { validateUsedRange } from '../../core/used-range.mjs';

export interface SpecialContextInputs {
  calculationControls: WorkbookCalculationControls | undefined;
  cancelCellFocus: () => void;
  conditionalControls: GridConditionalControls2;
  conditionalScope: string;
  documentId: string | undefined;
  formulaDraftHandle: React.MutableRefObject<FormulaDraftHandle | null>;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homePhysical: React.MutableRefObject<boolean>;
  homeRenderFrame: HomeFrame;
  liveConditionalScope: React.MutableRefObject<string>;
  liveValidationScope: React.MutableRefObject<string>;
  metadata: WorkbookSheet[];
  namedDraftKey: string;
  ownerIdentity: string | undefined;
  readRange: ReadRange | undefined;
  scrollCell: (owner: WorkbookSheet, row: number, col: number, align?: boolean) => void;
  selectedRef: React.MutableRefObject<string>;
  setActiveTool: React.Dispatch<React.SetStateAction<WorkbookTool | null>>;
  setHomePending: React.Dispatch<React.SetStateAction<boolean>>;
  setJump: React.Dispatch<React.SetStateAction<string>>;
  setRange: React.Dispatch<React.SetStateAction<string>>;
  setSelected: React.Dispatch<React.SetStateAction<string>>;
  setSelectionEnd: React.Dispatch<React.SetStateAction<string>>;
  setSpecialContext: React.Dispatch<React.SetStateAction<SpecialSelectionContext | null>>;
  setSpecialSelection: React.Dispatch<React.SetStateAction<Set<string> | null>>;
  setVisitedTools: React.Dispatch<React.SetStateAction<WorkbookTool[]>>;
  sheet: WorkbookSheet;
  specialLive: React.MutableRefObject<{ key: string; visible: boolean }>;
  specialParkedDraft: GridSpecialParkedDraft;
  specialSelectPhysical: React.MutableRefObject<boolean>;
  specialSerial: React.MutableRefObject<number>;
  toolFocusGeneration: React.MutableRefObject<number>;
  usedRangeControls: WorkbookUsedRangeControls | undefined;
  validationControls: GridValidationControls4;
  validationScope: string;
}

export function useSpecialContext({
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
}: SpecialContextInputs) {
  const specialSourceSheet = metadata.find((item) => item.name === sheet?.name);
  const specialCalculationReceipt = calculationControls?.receipt;
  const specialCalculationRevision =
    specialCalculationReceipt &&
    specialCalculationReceipt.documentId === documentId &&
    specialCalculationReceipt.ownerId === ownerIdentity
      ? specialCalculationReceipt.revision
      : specialSourceSheet?.calculationRevision;
  const specialSourceKey = JSON.stringify([
    specialSourceSheet?.name,
    specialCalculationRevision,
    specialCalculationReceipt,
    usedRangeControls?.receipt,
    specialSourceSheet?.merges,
    specialSourceSheet?.sheetFormat,
    specialSourceSheet?.hiddenRows,
    specialSourceSheet?.visibleRows,
    specialSourceSheet?.rowHeights,
    specialSourceSheet?.widths,
    specialSourceSheet?.geometryIssues,
    conditionalScope,
    validationScope,
  ]);
  const specialSourceLive = useRef(specialSourceKey);
  specialSourceLive.current = specialSourceKey;
  const specialDisabledReason = (): string | undefined => {
    const frame = homeFrameRef.current;
    if (!frame?.active) return 'Activate this workbook before opening Go To Special.';
    if (frame.profile === 'passive') return 'Go To Special requires an owned worksheet reader.';
    if (frame.binding.mode !== 'normal' || frame.binding.specialSelection)
      return 'Select a regular range in Normal view.';
    if (!frame.binding.ownerIdentity || !readRange) return 'The complete original worksheet reader is unavailable.';
    if (frame.busy || homePhysical.current) return 'Wait for the current workbook operation to finish.';
    // The family collector limits materialized positions. Last Cell only
    // needs a corner, even when the current rectangle is very large.
    try {
      selectedBounds(frame.binding.selection, Number.MAX_SAFE_INTEGER);
    } catch (error) {
      return error instanceof Error ? error.message : String(error);
    }
  };
  const captureSpecialContext = (kind: SpecialSelectionKind): SpecialSelectionContext => {
    if (!homeFrameRef.current?.active || !homeBindingMatches(homeRenderFrame.binding, homeFrameRef.current.binding))
      throw Error('The workbook or selection changed. Reopen Go To Special.');
    const reason = specialDisabledReason();
    if (reason) throw Error(reason);
    const binding = captureHomeBinding(homeRenderFrame.binding),
      frame: HomeFrame = { ...homeRenderFrame, binding },
      key = JSON.stringify([++specialSerial.current, binding]),
      sourceKey = specialSourceKey,
      calculationRevision = specialCalculationRevision;
    if (!specialSourceSheet) throw Error('The original worksheet metadata is unavailable.');
    const originalSheet = structuredClone({ ...specialSourceSheet, cells: {} }),
      reader = readRange!,
      conditionalInventory = conditionalControls?.inventory,
      validationInventory = validationControls?.inventory;
    const originalUsedRangeRead = usedRangeControls?.read,
      originalUsedRangeReceipt = usedRangeControls?.receipt ? structuredClone(usedRangeControls.receipt) : undefined;
    const originalConditionalScope = conditionalScope,
      originalValidationScope = validationScope;
    const scope = { captureFrame: () => frame, currentFrame: () => homeFrameRef.current };
    const current = () =>
      specialLive.current.visible &&
      specialLive.current.key === key &&
      specialSourceLive.current === sourceKey &&
      !!homeFrameRef.current?.active &&
      homeFrameRef.current.profile === frame.profile &&
      homeBindingMatches(binding, homeFrameRef.current.binding);
    let selectedReceipt: { cell: string; references: readonly string[]; epoch: number; regular?: string } | null = null;
    const runtime = () => {
      const live = homeFrameRef.current;
      if (
        !specialLive.current.visible ||
        specialLive.current.key !== key ||
        specialSourceLive.current !== sourceKey ||
        !live?.active ||
        live.profile !== frame.profile ||
        !(
          [
            'documentId',
            'ownerIdentity',
            'state',
            'sheet',
            'activationEpoch',
            'navigationGeneration',
            'modeGeneration',
            'mode',
          ] as const
        ).every((field) => binding[field] === live.binding[field])
      )
        return false;
      const receipt = selectedReceipt;
      if (!receipt) return homeBindingMatches(binding, live.binding);
      if (receipt.regular)
        return (
          live.binding.selection === receipt.regular &&
          live.binding.selectedCell === receipt.cell &&
          live.binding.specialSelectionEpoch === receipt.epoch &&
          !live.binding.specialSelection
        );
      if (live.binding.selection !== binding.selection) return false;
      const targets = live.binding.specialSelection;
      return (
        live.binding.selectedCell === receipt.cell &&
        live.binding.specialSelectionEpoch === receipt.epoch &&
        !!targets &&
        targets.length === receipt.references.length &&
        [...targets].sort().every((reference, index) => reference === receipt.references[index])
      );
    };
    let prepared: SpecialSelectionPreview | null = null;
    const read = async (
      lease: HomeLease,
      options: SpecialSelectionOptions,
      signal?: AbortSignal
    ): Promise<SpecialSelectionPreview> => {
      if (!current() || signal?.aborted) throw Error('The workbook or selection changed. Reopen Go To Special.');
      let usedRange;
      if (specialSelectionNeedsUsedRange(binding.selection, originalSheet, options.kind)) {
        if (!originalUsedRangeRead || !originalUsedRangeReceipt)
          throw Error('A complete original-owner UsedRange reader is unavailable.');
        const result = await lease.read(() => originalUsedRangeRead(binding.sheet));
        if (!current() || signal?.aborted) throw Error('The worksheet extent changed. Reopen Go To Special.');
        usedRange = validateUsedRange(result, { sheet: binding.sheet, receipt: originalUsedRangeReceipt });
      }
      let rules: SpecialSelectionRuleReceipt | undefined;
      if (options.kind === 'conditional-formats' || options.kind === 'data-validation') {
        const conditional = options.kind === 'conditional-formats',
          inventory = conditional ? conditionalInventory : validationInventory,
          expectedScope = conditional ? originalConditionalScope : originalValidationScope,
          liveScope = conditional ? liveConditionalScope : liveValidationScope;
        if (!inventory || !expectedScope || liveScope.current !== expectedScope)
          throw Error('A complete original authored rule-range inventory is unavailable.');
        const definition = await lease.read<WorkbookConditionalDefinition | WorkbookValidationDefinition>(() =>
          inventory(binding.sheet)
        );
        if (!current() || liveScope.current !== expectedScope)
          throw Error('The original rule-range inventory changed. Reopen Go To Special.');
        rules = { complete: true, definition: structuredClone(definition) };
      }
      return await collectSpecialSelectionPreview({
        binding,
        options,
        geometry: originalSheet,
        isCurrent: () => current() && lease.isCurrent(),
        ...(signal ? { signal } : {}),
        ...(calculationRevision ? { calculationRevision } : {}),
        visibility: originalSheet,
        ...(rules ? { rules } : {}),
        ...(usedRange ? { usedRange, usedRangeReceipt: originalUsedRangeReceipt! } : {}),
        read: async (name, target) => {
          if (name !== binding.sheet) throw Error('The Go To Special read belongs to another worksheet.');
          selectedBounds(target, 1000);
          const cells = await lease.read(() => reader(name, target));
          if (
            !cells ||
            typeof cells !== 'object' ||
            Array.isArray(cells) ||
            ![Object.prototype, null].includes(Object.getPrototypeOf(cells))
          )
            throw Error('The original worksheet did not return a complete cell range.');
          const bounds = selectedBounds(target, 1000);
          for (const [reference, cell] of Object.entries(cells)) {
            const position = addressPosition(reference);
            if (
              !/^[A-Z]{1,3}[1-9]\d*$/.test(reference) ||
              position.row < bounds.first.row ||
              position.row > bounds.last.row ||
              position.col < bounds.first.col ||
              position.col > bounds.last.col ||
              !cell ||
              typeof cell !== 'object' ||
              Array.isArray(cell) ||
              !Object.hasOwn(cell, 'value') ||
              (cell.value !== null && !['number', 'string', 'boolean'].includes(typeof cell.value))
            )
              throw Error('The original worksheet returned a foreign or malformed source cell.');
          }
          // A successful trusted owner read covers this entire bounded range.
          // The collector also checks exact keys/types/bytes before classifying.
          return { range: target, complete: true, cells: structuredClone(cells) };
        },
        yieldWork: () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())),
      });
    };
    return {
      key,
      binding,
      caption: `${binding.sheet}!${binding.selection}`,
      initialOptions: { ...DEFAULT_SPECIAL_SELECTION_OPTIONS, kind },
      isCurrent: current,
      isRuntimeCurrent: runtime,
      disabledReason: specialDisabledReason,
      preview: async (options, signal) => {
        prepared = null;
        const result = await dispatchHomeCommand('home.special-preview', scope, (lease) =>
          read(lease, options, signal)
        );
        if (result.outcome === 'failed') throw result.error;
        if (result.outcome === 'refused') throw Error(result.reason);
        if (result.outcome !== 'completed' || !result.current || !current() || signal.aborted)
          throw Error('The workbook or selection changed. Reopen Go To Special.');
        prepared = result.value;
        return prepared;
      },
      apply: async (receipt) => {
        if (
          !current() ||
          !prepared ||
          !prepared.references.length ||
          JSON.stringify(receipt) !== JSON.stringify(prepared)
        )
          throw Error('Preview the complete current selection before selecting matching cells.');
        const blocked = specialDisabledReason();
        if (blocked) throw Error(blocked);
        if (specialSelectPhysical.current || homePhysical.current)
          throw Error('Wait for the current selection to finish.');
        const captured = prepared;
        specialSelectPhysical.current = true;
        homePhysical.current = true;
        setHomePending(true);
        try {
          const result = await dispatchHomeCommand('home.special-select', scope, async (lease) => {
            const fresh = await read(lease, captured.options);
            if (!current() || JSON.stringify(fresh) !== JSON.stringify(captured))
              throw Error('The matching cells changed. Preview the selection again.');
            lease.assertCurrent();
            specialParkedDraft.current = null;
            const cursor = captured.activeCell;
            if (!cursor || !captured.references.includes(cursor))
              throw Error('The accepted Special selection has no current active cell.');
            if (captured.options.kind === 'current-region') {
              selectedReceipt = {
                cell: cursor,
                references: [],
                epoch: binding.specialSelectionEpoch,
                regular: captured.range,
              };
              const last = captured.range.split(':')[1] ?? captured.range;
              flushSync(() => {
                selectedRef.current = cursor;
                setSelected(cursor);
                setSelectionEnd(last);
                setRange(captured.range);
                setJump('');
                setSpecialSelection(null);
              });
            } else {
              selectedReceipt = {
                cell: cursor,
                references: Object.freeze([...captured.references].sort()),
                epoch: binding.specialSelectionEpoch + 1,
              };
              // Commit this exact local selection receipt before acknowledging it.
              // select() would clear the sparse set and collapse the source range.
              flushSync(() => {
                selectedRef.current = cursor;
                setSelected(cursor);
                setSelectionEnd(cursor);
                setJump('');
                setSpecialSelection(new Set(captured.references));
              });
            }
            if (!runtime()) throw Error('The accepted Special selection view changed.');
            const position = addressPosition(cursor);
            scrollCell(originalSheet, position.row, position.col);
            if (!runtime()) throw Error('The accepted Special selection view changed during reveal.');
          });
          if (result.outcome === 'failed') throw result.error;
          if (result.outcome === 'refused') throw Error(result.reason);
          if (result.outcome !== 'completed') throw Error('The workbook or selection changed. Reopen Go To Special.');
          prepared = null;
        } finally {
          specialSelectPhysical.current = false;
          homePhysical.current = false;
          setHomePending(false);
        }
      },
    };
  };
  const openSpecial = (kind: SpecialSelectionKind = 'constants'): void => {
    const reason = specialDisabledReason();
    if (reason) throw Error(reason);
    const handle = formulaDraftHandle.current;
    if (handle && (handle.ownsFocus() || handle.isParked?.())) {
      const snapshot = handle.capture();
      if (snapshot.composing) throw Error('Finish the current input before opening Go To Special.');
      handle.park();
      specialParkedDraft.current = { key: namedDraftKey, handle, snapshot };
    } else specialParkedDraft.current = null;
    const context = captureSpecialContext(kind);
    cancelCellFocus();
    toolFocusGeneration.current++;
    setSpecialContext(context);
    setVisitedTools((previous) => (previous.includes('special') ? previous : [...previous, 'special']));
    setActiveTool('special');
  };
  return { specialDisabledReason, openSpecial } as const;
}
