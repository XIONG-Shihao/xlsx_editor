import type { AcceptedPrintPreview } from '../useAcceptedPrintPreview';
import type { ViewPreferences } from '../viewPreferences';
import type { WorkbookCell, WorkbookNormalFont, WorkbookSheet } from '../../core/contracts';
import { addressPosition, colName } from '../sheetModel';
import { mergeAnchor } from './cellNavigation';
import { rectangle, selectedBounds } from './rangeAddress';

export interface NavigationActionsInputs {
  activeFrame: (work: () => void) => number;
  activeRange: () => string;
  activeStorage: React.MutableRefObject<string | undefined>;
  cancelCellFocus: () => void;
  firstRenderedColumn: number;
  frozenColumnCount: number;
  frozenColumnGap: number;
  gridRoot: React.RefObject<HTMLDivElement>;
  gridScroll: React.RefObject<HTMLDivElement>;
  jump: string;
  lastRenderedColumn: number;
  modeGeneration: React.MutableRefObject<number>;
  navigationGeneration: React.MutableRefObject<number>;
  normalFont: Readonly<WorkbookNormalFont>;
  normalReturnPending: React.MutableRefObject<boolean>;
  pendingCellFocus: React.MutableRefObject<number | null>;
  printPreview: AcceptedPrintPreview;
  renderedColumnIndices: number[];
  scrollCell: (owner: WorkbookSheet, row: number, col: number, align?: boolean) => void;
  select: (cell: string, extend?: boolean) => void;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setJump: React.Dispatch<React.SetStateAction<string>>;
  setJumpError: React.Dispatch<React.SetStateAction<string>>;
  setRange: React.Dispatch<React.SetStateAction<string>>;
  setSelectionEnd: React.Dispatch<React.SetStateAction<string>>;
  setSheetName: React.Dispatch<React.SetStateAction<string>>;
  setViewMode: React.Dispatch<React.SetStateAction<'normal' | 'pageLayout' | 'pageBreak'>>;
  sheet: WorkbookSheet;
  sheets: WorkbookSheet[];
  specialSelection: Set<string> | null;
  view: ViewPreferences;
  viewActive: React.MutableRefObject<boolean>;
  visibleColumns: number;
}

export function buildNavigationActions({
  activeFrame,
  activeRange,
  activeStorage,
  cancelCellFocus,
  firstRenderedColumn,
  frozenColumnCount,
  frozenColumnGap,
  gridRoot,
  gridScroll,
  jump,
  lastRenderedColumn,
  modeGeneration,
  navigationGeneration,
  normalFont,
  normalReturnPending,
  pendingCellFocus,
  printPreview,
  renderedColumnIndices,
  scrollCell,
  select,
  setGridError,
  setJump,
  setJumpError,
  setRange,
  setSelectionEnd,
  setSheetName,
  setViewMode,
  sheet,
  sheets,
  specialSelection,
  view,
  viewActive,
  visibleColumns,
}: NavigationActionsInputs) {
  const mixedStyle = (key: keyof WorkbookCell): boolean => {
    try {
      const bounds = selectedBounds(activeRange(), 1000);
      const refs = specialSelection
        ? [...specialSelection]
        : Array.from(
            { length: bounds.rows * bounds.columns },
            (_, index) =>
              `${colName(bounds.first.col + (index % bounds.columns))}${bounds.first.row + Math.floor(index / bounds.columns) + 1}`
          );
      if (refs.length < 2) return false;
      const styleValue = (cell: WorkbookCell | undefined): unknown => {
        const value = cell?.[key];
        if (key === 'format') return value ?? 0;
        if (key === 'fontFamily') return value ?? normalFont.fontFamily;
        if (key === 'fontSize') return value ?? normalFont.fontSize;
        if (key === 'bold' || key === 'italic') return value ?? normalFont[key];
        if (['underline', 'wrapText'].includes(key)) return value ?? false;
        return value;
      };
      const first = styleValue(sheet.cells[refs[0]]);
      for (const ref of refs.slice(1)) {
        const value = styleValue(sheet.cells[ref]);
        if (!Object.is(value, first)) return true;
      }
      return false;
    } catch {
      return true;
    }
  };
  const tableColumnCount =
    renderedColumnIndices.length +
    Number(view.headings) +
    Number(firstRenderedColumn > frozenColumnCount) +
    Number(lastRenderedColumn < visibleColumns) +
    Number(frozenColumnGap > 0);
  const navigateMetadataRange = (name: string, target: string): void => {
    try {
      if (!viewActive.current) return;
      const owner = sheets.find((item) => item.name === name);
      if (!owner) throw Error('The referenced worksheet is unavailable.');
      const bounds = selectedBounds(target, 16384 * 1048576),
        ref = target.split(':')[0],
        trigger = document.activeElement,
        storage = activeStorage.current;
      cancelCellFocus();
      setSheetName(name);
      select(ref);
      setRange(target);
      setSelectionEnd(target.split(':')[1] ?? target);
      setGridError('');
      const navigation = navigationGeneration.current;
      pendingCellFocus.current = activeFrame(() => {
        pendingCellFocus.current = null;
        if (navigationGeneration.current !== navigation || activeStorage.current !== storage) return;
        const focused = document.activeElement;
        if (
          focused !== trigger &&
          focused instanceof HTMLElement &&
          focused.matches('input,textarea,select,[contenteditable="true"]') &&
          !gridScroll.current?.contains(focused)
        )
          return;
        scrollCell(owner, bounds.first.row, bounds.first.col);
        // Focusing a cell selects it. Keep a metadata rectangle intact by
        // returning keyboard focus to Go To, as rectangular Go To already does.
        const target =
          bounds.rows * bounds.columns === 1
            ? [...(gridRoot.current?.querySelectorAll<HTMLInputElement>('input') ?? [])].find(
                (input) => input.getAttribute('aria-label') === `${owner.name} ${ref}`
              )
            : gridRoot.current?.querySelector<HTMLInputElement>('.sheet-cell-reference');
        target?.focus({ preventScroll: true });
      });
    } catch (error) {
      setGridError(String(error));
    }
  };
  const goToName = (formula: string): void => {
    const match = /^(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!([\$A-Za-z0-9:]+)$/.exec(formula);
    if (!match) {
      setGridError('This name has a complex reference. Its formula is preserved and shown read-only.');
      return;
    }
    navigateMetadataRange((match[1] ?? match[2]).replaceAll("''", "'"), match[3].replaceAll('$', ''));
  };
  const jumpToCell = (): void => {
    try {
      const target = selectedBounds(jump.trim().toUpperCase(), Number.MAX_SAFE_INTEGER),
        ref = mergeAnchor(sheet, `${colName(target.first.col)}${target.first.row + 1}`),
        end = `${colName(target.last.col)}${target.last.row + 1}`,
        position = addressPosition(ref);
      cancelCellFocus();
      select(ref);
      setSelectionEnd(end);
      setRange(rectangle(ref, end));
      setJump('');
      setJumpError('');
      pendingCellFocus.current = activeFrame(() => {
        pendingCellFocus.current = null;
        scrollCell(sheet, position.row, position.col);
        const active = document.activeElement;
        if (
          active instanceof HTMLElement &&
          !active.classList.contains('sheet-cell-reference') &&
          active.matches('input,textarea,select,[contenteditable="true"]') &&
          !gridScroll.current?.contains(active)
        )
          return;
        if (target.rows * target.columns === 1)
          [...(gridRoot.current?.querySelectorAll<HTMLInputElement>('input') ?? [])]
            .find((input) => input.getAttribute('aria-label') === `${sheet.name} ${ref}`)
            ?.focus({ preventScroll: true });
      });
    } catch {
      setJumpError('Enter a cell or rectangular range, such as A129972 or A1:C100, within worksheet bounds.');
    }
  };
  const returnNormal = (): void => {
    modeGeneration.current++;
    printPreview.abandonPreparation();
    normalReturnPending.current = true;
    setViewMode('normal');
    cancelCellFocus();
  };
  return { mixedStyle, tableColumnCount, navigateMetadataRange, goToName, jumpToCell, returnNormal } as const;
}
