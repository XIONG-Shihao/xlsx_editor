import type { GridSetViewport } from './gridTypes';
import { ViewPreferenceStore } from '../viewPreferences';
import type { WorkbookGeometryProfile, WorkbookSheet } from '../../core/contracts';
import { addressPosition } from '../sheetModel';
import { createViewGeometry, revealCellScroll } from '../viewGeometry';
import { flushSync } from 'react-dom';
import { measureGeometryProfile } from '../worksheetFontMetrics';
import { mergeAnchor } from './cellNavigation';
import { normalGeometryFont } from '../../core/worksheet-geometry.mjs';
import { useEffect } from 'react';

export interface ScrollToCellInputs {
  activeFrame: (work: () => void) => number;
  filterGeneration: React.MutableRefObject<number>;
  focus: { sheet: string; cell: string } | undefined;
  geometryProfile: WorkbookGeometryProfile;
  gridRoot: React.RefObject<HTMLDivElement>;
  gridScroll: React.RefObject<HTMLDivElement>;
  literal: boolean | undefined;
  preferenceDocumentId: string | undefined;
  preferenceStore: ViewPreferenceStore;
  setRange: React.Dispatch<React.SetStateAction<string>>;
  setSelected: React.Dispatch<React.SetStateAction<string>>;
  setSelectionEnd: React.Dispatch<React.SetStateAction<string>>;
  setSheetName: React.Dispatch<React.SetStateAction<string>>;
  setViewport: GridSetViewport;
  sheet: WorkbookSheet;
  sheets: WorkbookSheet[];
}

export function useScrollToCell({
  activeFrame,
  filterGeneration,
  focus,
  geometryProfile,
  gridRoot,
  gridScroll,
  literal,
  preferenceDocumentId,
  preferenceStore,
  setRange,
  setSelected,
  setSelectionEnd,
  setSheetName,
  setViewport,
  sheet,
  sheets,
}: ScrollToCellInputs) {
  const scrollCell = (owner: WorkbookSheet, row: number, col: number, align = false): void => {
    const scroll = gridScroll.current;
    if (!scroll) return;
    const preferences = preferenceStore.get(preferenceDocumentId, owner.name, owner.sourceView ?? {});
    const geometry = createViewGeometry(owner, {
      rows: Math.min(1048576, Math.max(30, owner.rows + 1, row + 1)),
      columns: Math.min(16384, Math.max(12, owner.columns + 1, col + 1)),
      zoom: preferences.zoom,
      headings: preferences.headings,
      geometryProfile: owner === sheet ? geometryProfile : measureGeometryProfile(normalGeometryFont(owner)),
      literal,
    });
    const current = {
      width: scroll.clientWidth,
      height: scroll.clientHeight,
      scrollLeft: scroll.scrollLeft,
      scrollTop: scroll.scrollTop,
    };
    const next = align
      ? {
          scrollTop: row < geometry.frozenRows ? 0 : Math.max(0, geometry.rowOffsets[row] - geometry.frozenHeight),
          scrollLeft:
            col < geometry.frozenColumns ? 0 : Math.max(0, geometry.columnOffsets[col] - geometry.frozenWidth),
          // Overlay scrollbars can cover cells without reducing clientWidth/Height.
          // Keep the selected corner and its fill handle clear of their hit region.
        }
      : revealCellScroll(
          geometry,
          { ...current, width: Math.max(0, current.width - 20), height: Math.max(0, current.height - 20) },
          row,
          col,
          owner
        );
    scroll.scrollTop = next.scrollTop;
    scroll.scrollLeft = next.scrollLeft;
    // Publish the actual browser-clamped offsets before a caller focuses a
    // virtualized target. A synthetic row index is not a screen coordinate.
    flushSync(() => setViewport({ ...current, scrollTop: scroll.scrollTop, scrollLeft: scroll.scrollLeft }));
  };
  useEffect(
    () => () => {
      filterGeneration.current++;
    },
    []
  );
  useEffect(() => {
    if (focus) {
      const targetSheet = sheets.find((item) => item.name === focus.sheet);
      const targetCell = mergeAnchor(targetSheet, focus.cell);
      setSheetName(focus.sheet);
      setSelected(targetCell);
      setRange(targetCell);
      setSelectionEnd(targetCell);
      const position = addressPosition(targetCell);
      const timer = activeFrame(() => {
        // Undo/redo can finish after the user starts typing in a toolbar field.
        // Do not redirect the remaining keystrokes into the restored cell.
        const active = document.activeElement;
        if (
          active instanceof HTMLElement &&
          active.matches('input,textarea,select,[contenteditable="true"]') &&
          !gridScroll.current?.contains(active)
        )
          return;
        if (targetSheet) scrollCell(targetSheet, position.row, position.col);
        const target = [...(gridRoot.current?.querySelectorAll<HTMLInputElement>('input') ?? [])].find(
          (input) => input.getAttribute('aria-label') === `${focus.sheet} ${targetCell}`
        );
        target?.focus({ preventScroll: true });
      });
      return () => cancelAnimationFrame(timer);
    }
  }, [focus]);
  return { scrollCell } as const;
}
