import {
  type CalculationAction,
  type CalculationFrame,
  type WorkbookCalculationControls,
  calculationDisabledReason,
} from '../calculationJobs';
import type {
  GridCalculationController,
  GridHomeInvoke,
  GridPaintStyleRef,
  GridRunCalculationAction,
} from './gridTypes';
import { type ReadRange, indexedBoundary } from '../rangeCells';
import type { RibbonCommandHandler, RibbonCommandId } from '../ribbonCommands';
import type { WorkbookGeometryProfile, WorkbookSheet } from '../../core/contracts';
import { addressPosition, colName } from '../sheetModel';
import { flushSync } from 'react-dom';
import { nextNavigableCell } from './cellNavigation';

export interface KeyboardNavigationInputs {
  activeFrame: (work: () => void) => number;
  activeStorage: React.MutableRefObject<string | undefined>;
  calculationController: GridCalculationController;
  calculationControls: WorkbookCalculationControls | undefined;
  calculationFrame: React.MutableRefObject<CalculationFrame | null>;
  cancelCellFocus: () => void;
  clearSelected: () => void;
  geometryProfile: WorkbookGeometryProfile;
  gridRoot: React.RefObject<HTMLDivElement>;
  homeInvoke: GridHomeInvoke;
  invoke: (id: RibbonCommandId, handler: RibbonCommandHandler) => void;
  literal: boolean | undefined;
  navigationGeneration: React.MutableRefObject<number>;
  normalMode: React.MutableRefObject<boolean>;
  paintStyleRef: GridPaintStyleRef;
  passiveNotice: string | undefined;
  readOnly: boolean;
  readRange: ReadRange | undefined;
  runCalculationAction: GridRunCalculationAction;
  scrollCell: (owner: WorkbookSheet, row: number, col: number, align?: boolean) => void;
  select: (cell: string, extend?: boolean) => void;
  selected: string;
  selectionEnd: string;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setHomeNavigationVersion: React.Dispatch<React.SetStateAction<number>>;
  setPaintStyleId: React.Dispatch<React.SetStateAction<number | null>>;
  sheet: WorkbookSheet;
  sheetRef: React.MutableRefObject<WorkbookSheet>;
  storageKey: string | undefined;
  viewActive: React.MutableRefObject<boolean>;
  /** Read when the code runs, like the closed-over variables they stand for. */
  late: {
    readonly busy: boolean;
    readonly renderedCalculationKey: string | null;
  };
}

export function buildKeyboardNavigation({
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
  late,
}: KeyboardNavigationInputs) {
  const navigate = async (event: React.KeyboardEvent): Promise<void> => {
    if (!normalMode.current) return;
    if (!viewActive.current) return;
    if (event.key === 'F9') {
      const target = event.target as HTMLElement,
        field = target.closest('input,textarea,select,[contenteditable="true"]');
      // Readonly cell-display inputs are a navigation surface until F2/type
      // enters editing. Never blur or commit a live text/formula/IME draft.
      const liveDraft =
        !!field &&
        !(
          field.matches('input') &&
          (['checkbox', 'radio', 'button', 'submit'].includes((field as HTMLInputElement).type) ||
            (!!field.closest('.sheet-grid') && field.getAttribute('data-editing') === 'false'))
        );
      if (
        event.defaultPrevented ||
        event.nativeEvent.isComposing ||
        event.keyCode === 229 ||
        target.closest('[data-composing="true"],[data-editing="true"]') ||
        liveDraft ||
        event.metaKey ||
        event.ctrlKey ||
        event.altKey ||
        !calculationControls ||
        readOnly ||
        late.busy ||
        literal ||
        passiveNotice
      )
        return;
      const action: CalculationAction = event.shiftKey
          ? { type: 'calculateSheet', sheet: sheet.name }
          : { type: 'calculateAllOpen' },
        id = event.shiftKey ? 'formulas.calculate-sheet' : 'formulas.calculate-now';
      if (
        calculationDisabledReason(calculationFrame.current, action) ||
        calculationController.getSnapshot().physicalBusy
      )
        return;
      event.preventDefault();
      invoke(id, (binding) => runCalculationAction(binding, action, late.renderedCalculationKey));
      return;
    }
    if (event.key === 'Escape') {
      const openMenus = gridRoot.current?.querySelectorAll<HTMLDetailsElement>('details.sheet-ribbon-menu[open]');
      if (openMenus?.length) {
        event.preventDefault();
        const menu = openMenus[openMenus.length - 1];
        openMenus.forEach((item) => {
          item.open = false;
        });
        activeFrame(() => menu.querySelector<HTMLElement>('summary')?.focus());
        return;
      }
      if (paintStyleRef.current !== null) {
        paintStyleRef.current = null;
        setPaintStyleId(null);
        return;
      }
    }
    if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'f') {
      event.preventDefault();
      void homeInvoke('home.find-open');
      return;
    }
    if (
      (event.metaKey || event.ctrlKey) &&
      ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key) &&
      (event.target as HTMLElement).closest('.sheet-grid') &&
      (event.target as HTMLElement).getAttribute('data-editing') !== 'true'
    ) {
      event.preventDefault();
      const current = sheetRef.current,
        position = addressPosition(event.shiftKey ? selectionEnd : selected);
      const vertical = event.key === 'ArrowUp' || event.key === 'ArrowDown';
      const step = event.key === 'ArrowUp' || event.key === 'ArrowLeft' ? -1 : 1;
      const limit = vertical ? Math.max(29, current.rows - 1) : Math.max(11, current.columns - 1);
      const filled = (index: number): boolean => {
        const ref = vertical ? `${colName(position.col)}${index + 1}` : `${colName(index)}${position.row + 1}`;
        const cell = current.cells[ref];
        return !!cell?.formula || (cell?.value !== undefined && cell.value !== null && cell.value !== '');
      };
      let index = vertical ? position.row : position.col;
      if (readRange) {
        const generation = ++navigationGeneration.current,
          revision = storageKey;
        setHomeNavigationVersion(generation);
        try {
          index = await indexedBoundary(readRange, current, position, vertical, step);
          if (!viewActive.current || generation !== navigationGeneration.current || revision !== activeStorage.current)
            return;
        } catch (error) {
          setGridError(String(error));
          return;
        }
      }
      const next = index + step;
      if (readRange) {
        /* The owner supplied the complete line boundary. */
      } else if (next >= 0 && next <= limit && filled(index) && filled(next)) {
        while (index + step >= 0 && index + step <= limit && filled(index + step)) index += step;
      } else {
        while (index + step >= 0 && index + step <= limit) {
          index += step;
          if (filled(index)) break;
        }
      }
      const row = vertical ? index : position.row,
        col = vertical ? position.col : index;
      const ref = `${colName(col)}${row + 1}`;
      cancelCellFocus();
      flushSync(() => select(ref, event.shiftKey));
      scrollCell(current, row, col);
      activeFrame(() =>
        [...(gridRoot.current?.querySelectorAll<HTMLInputElement>('input') ?? [])]
          .find((input) => input.getAttribute('aria-label') === `${current.name} ${ref}`)
          ?.focus({ preventScroll: true })
      );
      return;
    }
    if (event.nativeEvent.isComposing || event.metaKey || event.ctrlKey || event.altKey) return;
    if (
      ['Delete', 'Backspace'].includes(event.key) &&
      (event.target as HTMLElement).closest('.sheet-grid') &&
      (event.target as HTMLElement).getAttribute('data-editing') !== 'true'
    ) {
      event.preventDefault();
      clearSelected();
      return;
    }
    const movement: Record<string, [number, number]> = {
      ArrowUp: [-1, 0],
      ArrowDown: [1, 0],
      ArrowLeft: [0, -1],
      ArrowRight: [0, 1],
      Enter: [event.shiftKey ? -1 : 1, 0],
      Tab: [0, event.shiftKey ? -1 : 1],
    };
    if (!movement[event.key] || !(event.target as HTMLElement).closest('.sheet-grid')) return;
    if ((event.target as HTMLElement).getAttribute('data-editing') === 'true') return;
    event.preventDefault();
    const position = addressPosition(event.shiftKey ? selectionEnd : selected);
    const [row, col] = movement[event.key],
      destination = nextNavigableCell(sheet, position, row, col, geometryProfile);
    const { row: r, col: c, ref } = destination;
    if (event.shiftKey) select(ref, true);
    else {
      (event.target as HTMLInputElement).blur();
      flushSync(() => select(ref));
      scrollCell(sheet, r, c);
      activeFrame(() =>
        [...(gridRoot.current?.querySelectorAll<HTMLInputElement>('input') ?? [])]
          .find((input) => input.getAttribute('aria-label') === `${sheet.name} ${ref}`)
          ?.focus({ preventScroll: true })
      );
    }
  };
  return { navigate } as const;
}
