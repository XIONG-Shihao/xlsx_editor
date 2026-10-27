import { BULK_CELL_LIMIT } from './gridLimits';
import { COMMAND_BATCH_UNAVAILABLE, acceptGridCommandBatch } from './commandBatch';
import type { FormulaDraftHandle } from '../namedFormulaDraft';
import type { GridFillDragging, GridHomePaintRunner } from './gridTypes';
import {
  type HomeArgumentsById,
  type HomeCommandId,
  type HomeFrame,
  type HomeCommandLease as HomeLease,
  dispatchHomeCommand,
  homeBindingMatches,
  homeCommandCapability,
} from '../homeCommands';
import type { ReadRange } from '../rangeCells';
import type { RibbonBinding, RibbonDispatchScope } from '../ribbonCommands';
import type { WorkbookCell, WorkbookCommand, WorkbookSheet } from '../../core/contracts';
import { addressPosition, colName } from '../sheetModel';
import { selectedBounds } from './rangeAddress';
import { useEffect } from 'react';

export interface HomeDispatchInputs {
  activeFrame: (work: () => void) => number;
  clearHyperlinksDisabledReason: () => string | undefined;
  commandScope: RibbonDispatchScope;
  dragging: React.MutableRefObject<boolean>;
  fillDragging: GridFillDragging;
  formulaDraftHandle: React.MutableRefObject<FormulaDraftHandle | null>;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homeHandle: (id: HomeCommandId, lease: HomeLease, args: unknown) => Promise<void>;
  homePaintRunner: GridHomePaintRunner;
  homePaintTarget: React.MutableRefObject<string | null>;
  homePending: boolean;
  homePhysical: React.MutableRefObject<boolean>;
  navigationGeneration: React.MutableRefObject<number>;
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  onCommandRedo: (() => Promise<void>) | undefined;
  onCommandUndo: (() => Promise<void>) | undefined;
  readRange: ReadRange | undefined;
  scrollCell: (owner: WorkbookSheet, row: number, col: number, align?: boolean) => void;
  select: (cell: string, extend?: boolean) => void;
  setFillPreview: React.Dispatch<React.SetStateAction<string | null>>;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setHomePending: React.Dispatch<React.SetStateAction<boolean>>;
  setSheetName: React.Dispatch<React.SetStateAction<string>>;
  sheet: WorkbookSheet;
  sheets: WorkbookSheet[];
  specialSelection: Set<string> | null;
  specialSelectionMode: React.MutableRefObject<{ selection: Set<string> | null }>;
  viewActive: React.MutableRefObject<boolean>;
  /** Read when the code runs, like the closed-over variables they stand for. */
  late: {
    readonly acrossDisabledReason: () => string | undefined;
    readonly homeRenderFrame: HomeFrame;
    readonly seriesDisabledReason: (action?: 'preview' | 'apply') => string | undefined;
    readonly specialDisabledReason: () => string | undefined;
  };
}

export function useHomeDispatch({
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
  late,
}: HomeDispatchInputs) {
  const homeInvoke = async <K extends HomeCommandId>(
    id: K,
    args?: HomeArgumentsById[K],
    work?: (lease: HomeLease, args: HomeArgumentsById[K]) => void | Promise<void>
  ): Promise<void> => {
    if (
      !homeFrameRef.current?.active ||
      !homeBindingMatches(late.homeRenderFrame.binding, homeFrameRef.current.binding)
    )
      return;
    const local = [
      'home.fill-series',
      'home.fill-across',
      'home.special-open',
      'home.select-blanks',
      'home.select-formulas',
      'home.menu',
      'styles.conditional',
      'home.find-draft',
      'home.find-open',
      'home.find-close',
      'home.find-cancel',
      'home.find-next',
      'home.find-match',
      'home.replace-preview',
      'home.table-draft',
      'home.conditional-draft',
      'home.dimension-draft',
      'home.clear-special-selection',
      'home.cancel-paste',
    ].includes(id);
    if (!local && homePhysical.current) {
      if (homeFrameRef.current?.binding.ownerIdentity === late.homeRenderFrame.binding.ownerIdentity)
        setGridError('Wait for the current Home command to finish.');
      return;
    }
    const focused = document.activeElement as HTMLElement | null,
      formula = formulaDraftHandle.current;
    if (
      (focused?.closest('.sheet-grid') && focused.getAttribute('data-editing') === 'true') ||
      (formula?.ownsFocus() && formula.capture().composing)
    ) {
      setGridError('Finish or cancel the current input before using Home commands.');
      return;
    }
    if (formula?.ownsFocus()) formula.park();
    if (!local) {
      homePhysical.current = true;
      setHomePending(true);
    }
    try {
      const result = await dispatchHomeCommand(
        id,
        { captureFrame: () => late.homeRenderFrame, currentFrame: () => homeFrameRef.current },
        work ?? ((lease, args) => homeHandle(id, lease, args)),
        args
      );
      if (result.outcome === 'failed' && result.current)
        setGridError(String(result.error instanceof Error ? result.error.message : result.error));
      else if (
        result.outcome === 'refused' &&
        homeFrameRef.current?.binding.ownerIdentity === late.homeRenderFrame.binding.ownerIdentity &&
        homeFrameRef.current?.binding.state === late.homeRenderFrame.binding.state
      )
        setGridError(result.reason);
      else if (result.outcome === 'completed' && result.current) setGridError('');
    } catch (error) {
      if (
        homeFrameRef.current?.active &&
        homeBindingMatches(late.homeRenderFrame.binding, homeFrameRef.current.binding)
      )
        setGridError(String(error instanceof Error ? error.message : error));
    } finally {
      if (!local) {
        homePhysical.current = false;
        setHomePending(false);
      }
    }
  };
  homePaintRunner.current = (cell, format) => {
    homePaintTarget.current = cell;
    void homeInvoke('home.paint-format', {
      format: Object.fromEntries(Object.entries(format).filter(([, value]) => value !== undefined)),
    }).finally(() => {
      homePaintTarget.current = null;
    });
  };
  const homeLocalAction = (id: HomeCommandId): boolean =>
    [
      'home.fill-series',
      'home.fill-across',
      'home.special-open',
      'home.select-blanks',
      'home.select-formulas',
      'home.menu',
      'styles.conditional',
      'home.find-draft',
      'home.find-open',
      'home.find-close',
      'home.find-cancel',
      'home.find-next',
      'home.find-match',
      'home.replace-preview',
      'home.table-draft',
      'home.conditional-draft',
      'home.dimension-draft',
      'home.clear-special-selection',
      'home.cancel-paste',
    ].includes(id);
  const homeReason = (id: HomeCommandId): string | undefined =>
    homePending && !homeLocalAction(id)
      ? 'Wait for the current Home command to finish.'
      : (homeCommandCapability(id, late.homeRenderFrame).reason ??
        (id === 'home.clear-hyperlinks'
          ? clearHyperlinksDisabledReason()
          : id === 'home.fill-series'
            ? late.seriesDisabledReason()
            : id === 'home.fill-across'
              ? late.acrossDisabledReason()
              : ['home.special-open', 'home.select-blanks', 'home.select-formulas'].includes(id)
                ? late.specialDisabledReason()
                : undefined) ??
        ((id === 'home.undo' && !onCommandUndo) || (id === 'home.redo' && !onCommandRedo)
          ? 'Workbook history commands are unavailable in this editor.'
          : undefined));
  const homeDisabled = (id: HomeCommandId): boolean =>
    !!homeReason(id) || homeCommandCapability(id, late.homeRenderFrame).state !== 'enabled';
  const homeProps = (id: HomeCommandId) => ({
    'data-home-command': id,
    disabled: homeDisabled(id),
    'aria-description': homeReason(id),
    title: homeReason(id),
  });
  const homeMenu = (event: React.MouseEvent, menu: string): void => {
    if (
      !homeFrameRef.current?.active ||
      !homeBindingMatches(late.homeRenderFrame.binding, homeFrameRef.current.binding)
    )
      return;
    if (homeDisabled('home.menu')) event.preventDefault();
    void homeInvoke('home.menu', { value: menu as never });
  };

  useEffect(() => {
    const release = (): void => {
      if (!viewActive.current) return;
      dragging.current = false;
      const fill = fillDragging.current;
      fillDragging.current = null;
      setFillPreview(null);
      if (fill?.target && fill.dispatch) void fill.dispatch(fill.target);
    };
    window.addEventListener('mouseup', release);
    return () => window.removeEventListener('mouseup', release);
  }, [select]);
  const copy = (event: React.ClipboardEvent, cut = false): void => {
    const target = event.target as HTMLElement;
    if (
      (target.closest('input,textarea,select') && !target.closest('.sheet-grid')) ||
      target.getAttribute('data-editing') === 'true'
    )
      return;
    event.preventDefault();
    void homeInvoke(cut ? 'home.cut' : 'home.copy');
  };
  const clearSelected = (): void => void homeInvoke('home.clear-contents');
  const autoSum = async (binding: RibbonBinding): Promise<void> => {
    const originalSheet = sheet,
      reader = readRange,
      batch = onCommandBatch;
    const navigation = navigationGeneration.current,
      selectionMode = specialSelectionMode.current;
    const check = (): void => {
      if (
        !commandScope.isCurrent(binding) ||
        navigationGeneration.current !== navigation ||
        specialSelectionMode.current !== selectionMode
      )
        throw Error('The initiating workbook, revision or selection changed before AutoSum could be applied.');
    };
    check();
    const bounds = selectedBounds(binding.selection, BULK_CELL_LIMIT);
    if (bounds.rows < 2) throw Error('Select at least two rows to insert totals below them.');
    if (bounds.last.row >= 1048575) throw Error('The selected range reaches the last worksheet row.');
    if (specialSelection)
      throw Error('Special selection supports formatting and Clear. Click a cell for other actions.');
    const targets: Record<string, WorkbookCell | undefined> = {};
    const targetRow = bounds.last.row + 2;
    for (let col = bounds.first.col; col <= bounds.last.col; col++) {
      const target = `${colName(col)}${targetRow}`;
      targets[target] = originalSheet.cells[target];
    }
    // A rendered/selected tile does not cover the next row. Read the complete
    // output row at the captured owner head without adopting it into UI caches.
    if (reader)
      for (let first = bounds.first.col; first <= bounds.last.col; first += 1000) {
        const last = Math.min(bounds.last.col, first + 999);
        const found = await reader(binding.sheet, `${colName(first)}${targetRow}:${colName(last)}${targetRow}`);
        check();
        for (let col = first; col <= last; col++) {
          const target = `${colName(col)}${targetRow}`;
          if (Object.prototype.hasOwnProperty.call(found, target)) targets[target] = found[target];
        }
      }
    const cells = Array.from({ length: bounds.columns }, (_, index) => {
      const col = colName(bounds.first.col + index);
      const target = `${col}${targetRow}`,
        existing = targets[target];
      if (existing?.formula || existing?.formulaType || (existing?.value !== undefined && existing?.value !== null))
        throw Error(`AutoSum would overwrite ${target}; choose a range with an empty row below.`);
      return { cell: target, formula: `SUM(${col}${bounds.first.row + 1}:${col}${bounds.last.row + 1})` };
    });
    check();
    const available = commandScope.availability('formulas.auto-sum');
    if (available.capability?.state === 'disabled')
      throw Error(available.capability.reason ?? COMMAND_BATCH_UNAVAILABLE);
    await acceptGridCommandBatch([{ type: 'setCells', sheet: binding.sheet, cells }], {
      active: viewActive.current,
      readOnly: !available.canEdit,
      busy: available.busy,
      literal: available.profile !== 'xlsx',
      specialSelection: !!specialSelection,
      onCommandBatch: batch,
    });
  };
  const find = (): Promise<void> => homeInvoke('home.find-search');
  const goToMatch = (match: { sheet: string; cell: string }): void => {
    setSheetName(match.sheet);
    select(match.cell);
    const position = addressPosition(match.cell);
    const owner = sheets.find((item) => item.name === match.sheet);
    if (owner) activeFrame(() => scrollCell(owner, position.row, position.col));
  };
  const pasteText = (text: string): Promise<void> => homeInvoke('home.paste', { text });
  const paste = (event: React.ClipboardEvent): void => {
    const target = event.target as HTMLElement;
    if (
      (target.closest('input,textarea,select') && !target.closest('.sheet-grid')) ||
      target.getAttribute('data-editing') === 'true'
    )
      return;
    const text = event.clipboardData.getData('text/plain');
    if (!text) return;
    event.preventDefault();
    void pasteText(text);
  };
  return {
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
  } as const;
}
