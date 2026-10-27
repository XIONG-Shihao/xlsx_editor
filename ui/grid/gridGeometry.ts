import type { AcceptedPrintPreview } from '../useAcceptedPrintPreview';
import type { FormulaDraftHandle } from '../namedFormulaDraft';
import type { GridObjectControls2 } from './gridTypes';
import type { InkPenId } from '../../core/ink-stroke.mjs';
import type { RibbonBinding, RibbonDispatchScope } from '../ribbonCommands';
import { ViewPreferenceStore, type ViewPreferences } from '../viewPreferences';
import type { WorkbookCommand, WorkbookGeometryProfile, WorkbookSheet } from '../../core/contracts';
import type { WorkbookTool } from '../SpreadsheetRibbon';
import type { WorksheetDimensions } from '../../core/worksheet-geometry.mjs';
import { addressPosition } from '../sheetModel';
import { backgroundTile } from '../backgroundTiles';
import { createObjectAxes } from '../worksheetObjectGeometry';
import { createViewGeometry } from '../viewGeometry';
import { nextNavigableCell } from './cellNavigation';
import { pageBoundaryModel } from '../pageBoundaries';
import { useCallback, useEffect, useMemo, useRef } from 'react';

export interface GridGeometryInputs {
  activationEpoch: React.MutableRefObject<number>;
  active: boolean;
  activeFrame: (work: () => void) => number;
  busy: boolean;
  cancelCellFocus: () => void;
  commandScope: RibbonDispatchScope;
  drawMode: 'draw' | 'select' | 'erase' | null;
  formulaDraftHandle: React.MutableRefObject<FormulaDraftHandle | null>;
  functionViewMode: React.MutableRefObject<{ mode: 'normal' | 'pageLayout' | 'pageBreak'; epoch: number }>;
  geometryKey: string;
  geometryProfile: WorkbookGeometryProfile;
  gridRoot: React.RefObject<HTMLDivElement>;
  inkPreset: InkPenId;
  literal: boolean | undefined;
  normalMode: React.MutableRefObject<boolean>;
  objectControls: GridObjectControls2;
  objectExtent: { family: string; rows: number; columns: number };
  onCommandBatch: ((commands: WorkbookCommand[]) => Promise<void>) | undefined;
  ownerIdentity: string | undefined;
  pendingCellFocus: React.MutableRefObject<number | null>;
  preferenceDocumentId: string | undefined;
  preferenceStore: ViewPreferenceStore;
  printPreview: AcceptedPrintPreview;
  readOnly: boolean;
  scrollCell: (owner: WorkbookSheet, row: number, col: number, align?: boolean) => void;
  select: (cell: string, extend?: boolean) => void;
  selected: string;
  setActiveTool: React.Dispatch<React.SetStateAction<WorkbookTool | null>>;
  setDrawMode: React.Dispatch<React.SetStateAction<'draw' | 'select' | 'erase' | null>>;
  sheet: WorkbookSheet;
  sheetRef: React.MutableRefObject<WorkbookSheet>;
  specialSelection: Set<string> | null;
  storageKey: string | undefined;
  view: ViewPreferences;
  viewActive: React.MutableRefObject<boolean>;
  viewMode: 'normal' | 'pageLayout' | 'pageBreak';
  viewport: { width: number; height: number; scrollTop: number; scrollLeft: number };
  worksheetDimensions: WorksheetDimensions;
}

export function useGridGeometry({
  activationEpoch,
  active,
  activeFrame,
  busy,
  cancelCellFocus,
  commandScope,
  drawMode,
  formulaDraftHandle,
  functionViewMode,
  geometryKey,
  geometryProfile,
  gridRoot,
  inkPreset,
  literal,
  normalMode,
  objectControls,
  objectExtent,
  onCommandBatch,
  ownerIdentity,
  pendingCellFocus,
  preferenceDocumentId,
  preferenceStore,
  printPreview,
  readOnly,
  scrollCell,
  select,
  selected,
  setActiveTool,
  setDrawMode,
  sheet,
  sheetRef,
  specialSelection,
  storageKey,
  view,
  viewActive,
  viewMode,
  viewport,
  worksheetDimensions,
}: GridGeometryInputs) {
  const moveAfterCommit = useCallback(
    (from: string, key: 'Enter' | 'Tab', shift: boolean): void => {
      if (!viewActive.current) return;
      const position = addressPosition(from);
      const destination = nextNavigableCell(
        sheetRef.current,
        position,
        key === 'Enter' ? (shift ? -1 : 1) : 0,
        key === 'Tab' ? (shift ? -1 : 1) : 0,
        geometryProfile
      );
      const { row, col, ref } = destination;
      select(ref);
      cancelCellFocus();
      pendingCellFocus.current = activeFrame(() => {
        pendingCellFocus.current = null;
        scrollCell(sheetRef.current, row, col);
        [...(gridRoot.current?.querySelectorAll<HTMLInputElement>('input') ?? [])]
          .find((input) => input.getAttribute('aria-label') === `${sheetRef.current.name} ${ref}`)
          ?.focus({ preventScroll: true });
      });
    },
    [select, preferenceDocumentId, preferenceStore, geometryKey, literal]
  );
  const visibleOverrideEnd = sheet?.visibleRows?.reduce((last, row) => Math.max(last, row + 1), 0) ?? 0;
  const defaultRowPixels = literal ? 58 : (worksheetDimensions.defaultRowHeight * 96) / 72;
  const defaultColumnPixels = literal ? 100 : worksheetDimensions.defaultColumnPixels;
  const objectFamily = JSON.stringify([objectControls?.documentId ?? '', ownerIdentity ?? '', sheet?.name ?? '']);
  const objectFamilyEpoch = useRef({ family: objectFamily, epoch: 0 });
  if (objectFamilyEpoch.current.family !== objectFamily)
    objectFamilyEpoch.current = { family: objectFamily, epoch: objectFamilyEpoch.current.epoch + 1 };
  const drawingExtent = objectExtent.family === objectFamily ? objectExtent : { rows: 0, columns: 0 };
  const objectAxes = useMemo(
    () => createObjectAxes(sheet ?? {}, geometryProfile),
    [
      sheet?.sheetFormat,
      sheet?.normalFont,
      sheet?.rowHeights,
      sheet?.hiddenRows,
      sheet?.visibleRows,
      sheet?.widths,
      geometryKey,
    ]
  );
  const boundaryModel = useMemo(() => {
    if (viewMode !== 'pageBreak' || !printPreview.job) return null;
    try {
      return pageBoundaryModel(printPreview.job);
    } catch {
      return null;
    }
  }, [viewMode, printPreview.job]);
  const visibleRows = Math.min(
    1048576,
    Math.max(
      30,
      defaultRowPixels > 0 ? Math.ceil(viewport.height / ((defaultRowPixels * view.zoom) / 100)) + 12 : 0,
      (sheet?.rows ?? 0) + 1,
      drawingExtent.rows,
      boundaryModel?.rows ?? 0,
      visibleOverrideEnd,
      addressPosition(selected).row + 1
    )
  );
  const visibleColumns = Math.min(
    16384,
    Math.max(
      12,
      defaultColumnPixels > 0 ? Math.ceil(viewport.width / ((defaultColumnPixels * view.zoom) / 100)) + 6 : 0,
      (sheet?.columns ?? 0) + 1,
      drawingExtent.columns,
      boundaryModel?.columns ?? 0,
      addressPosition(selected).col + 1
    )
  );
  const worksheetBackground = backgroundTile(sheet, view.zoom / 100);
  const geometry = useMemo(
    () =>
      createViewGeometry(viewMode === 'pageBreak' ? { ...sheet, freeze: { rows: 0, columns: 0 } } : sheet, {
        rows: visibleRows,
        columns: visibleColumns,
        zoom: view.zoom,
        headings: view.headings,
        geometryProfile,
        literal,
      }),
    [
      sheet?.sheetFormat,
      sheet?.normalFont,
      sheet?.rowHeights,
      sheet?.hiddenRows,
      sheet?.visibleRows,
      sheet?.widths,
      sheet?.freeze,
      geometryKey,
      literal,
      visibleRows,
      visibleColumns,
      view.zoom,
      view.headings,
      viewMode,
    ]
  );
  const inkReason = !onCommandBatch
    ? 'Drawing is unavailable in this host.'
    : !active
      ? 'Activate this workbook to draw.'
      : !normalMode.current
        ? 'Return to Normal view to draw.'
        : readOnly
          ? 'This workbook is read-only.'
          : busy
            ? 'Wait for the workbook operation to finish.'
            : sheet.freeze?.rows || sheet.freeze?.columns
              ? 'Unfreeze panes before drawing.'
              : specialSelection
                ? 'Click a regular cell before drawing.'
                : undefined;
  const inkKey = JSON.stringify([
    objectFamily,
    objectFamilyEpoch.current.epoch,
    objectControls?.revision,
    storageKey ?? '',
    activationEpoch.current,
    functionViewMode.current.epoch,
    geometryKey,
    sheet.sheetFormat,
    sheet.widths,
    sheet.rowHeights,
    sheet.hiddenRows,
    sheet.visibleRows,
    sheet.normalFont,
    view.zoom,
    view.headings,
    viewport.scrollLeft,
    viewport.scrollTop,
    viewport.width,
    viewport.height,
    drawMode,
    inkPreset,
  ]);
  const inkLive = useRef({ key: inkKey, active: active && drawMode, reason: inkReason });
  inkLive.current = { key: inkKey, active: active && drawMode, reason: inkReason };
  useEffect(() => {
    setDrawMode(null);
  }, [objectFamily, active, viewMode]);
  const enterDraw = (binding: RibbonBinding, mode: 'draw' | 'select' | 'erase' = 'draw'): void => {
    if (!commandScope.isCurrent(binding) || inkReason) throw Error(inkReason ?? 'The original worksheet changed.');
    const focus = document.activeElement as HTMLElement | null;
    if (focus?.getAttribute('data-editing') === 'true' || formulaDraftHandle.current?.ownsFocus())
      throw Error('Finish or cancel the cell edit before drawing.');
    cancelCellFocus();
    setActiveTool(null);
    setDrawMode(mode);
  };
  const { rowOffsets, columnOffsets } = geometry;
  return {
    moveAfterCommit,
    objectFamily,
    objectFamilyEpoch,
    objectAxes,
    boundaryModel,
    visibleRows,
    visibleColumns,
    worksheetBackground,
    geometry,
    inkReason,
    inkKey,
    inkLive,
    enterDraw,
    rowOffsets,
    columnOffsets,
  } as const;
}
