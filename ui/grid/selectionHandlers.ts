import type { GridFillDragging, GridHomePaintRunner, GridPaintStyleRef } from './gridTypes';
import type { HomeBinding, HomeFrame } from '../homeCommands';
import type { WorkbookEdit, WorkbookSheet } from '../../core/contracts';
import { addressPosition, colName, parseGridInput } from '../sheetModel';
import { rectangle, selectedBounds } from './rangeAddress';
import { useCallback } from 'react';

export interface SelectionHandlersInputs {
  dragging: React.MutableRefObject<boolean>;
  editRef: React.MutableRefObject<(edit: WorkbookEdit) => void>;
  fillDragging: GridFillDragging;
  homeFrameRef: React.MutableRefObject<HomeFrame | null>;
  homePaintRunner: GridHomePaintRunner;
  literal: boolean | undefined;
  navigationGeneration: React.MutableRefObject<number>;
  normalMode: React.MutableRefObject<boolean>;
  paintHomeBinding: React.MutableRefObject<HomeBinding | null>;
  paintStyleRef: GridPaintStyleRef;
  selectedRef: React.MutableRefObject<string>;
  setFillPreview: React.Dispatch<React.SetStateAction<string | null>>;
  setGridError: React.Dispatch<React.SetStateAction<string>>;
  setHomeNavigationVersion: React.Dispatch<React.SetStateAction<number>>;
  setPaintStyleId: React.Dispatch<React.SetStateAction<number | null>>;
  setRange: React.Dispatch<React.SetStateAction<string>>;
  setSelected: React.Dispatch<React.SetStateAction<string>>;
  setSelectionEnd: React.Dispatch<React.SetStateAction<string>>;
  setSpecialSelection: React.Dispatch<React.SetStateAction<Set<string> | null>>;
  sheet: WorkbookSheet;
  sheetRef: React.MutableRefObject<WorkbookSheet>;
  specialSelection: Set<string> | null;
  viewActive: React.MutableRefObject<boolean>;
}

export function useSelectionHandlers({
  dragging,
  editRef,
  fillDragging,
  homeFrameRef,
  homePaintRunner,
  literal,
  navigationGeneration,
  normalMode,
  paintHomeBinding,
  paintStyleRef,
  selectedRef,
  setFillPreview,
  setGridError,
  setHomeNavigationVersion,
  setPaintStyleId,
  setRange,
  setSelected,
  setSelectionEnd,
  setSpecialSelection,
  sheet,
  sheetRef,
  specialSelection,
  viewActive,
}: SelectionHandlersInputs) {
  const setValue = useCallback(
    (cell: string, value: string): void => {
      if (!viewActive.current || !normalMode.current) return;
      if (specialSelection) {
        setGridError('Click a cell to leave the special selection before entering a value.');
        return;
      }
      const current = sheetRef.current;
      editRef.current({
        sheet: current.name,
        cell,
        ...parseGridInput(
          value,
          !!literal || current.cells[cell]?.format === 49 || current.cells[cell]?.format === '@'
        ),
      });
    },
    [literal, specialSelection]
  );
  const select = useCallback((cell: string, extend = false): void => {
    navigationGeneration.current++;
    setHomeNavigationVersion(navigationGeneration.current);
    setSpecialSelection(null);
    setSelectionEnd(cell);
    if (extend) setRange(rectangle(selectedRef.current, cell));
    else {
      selectedRef.current = cell;
      setSelected(cell);
      setRange(cell);
    }
  }, []);
  const selectRow = useCallback((row: number): void => {
    navigationGeneration.current++;
    setHomeNavigationVersion(navigationGeneration.current);
    setSpecialSelection(null);
    const from = `A${row + 1}`;
    const to = `${colName(Math.max(0, sheetRef.current.columns - 1))}${row + 1}`;
    selectedRef.current = from;
    setSelected(from);
    setSelectionEnd(to);
    setRange(rectangle(from, to));
  }, []);
  const selectColumn = (col: number): void => {
    navigationGeneration.current++;
    setHomeNavigationVersion(navigationGeneration.current);
    setSpecialSelection(null);
    const from = `${colName(col)}1`;
    const to = `${colName(col)}${Math.max(sheet.rows, 30)}`;
    selectedRef.current = from;
    setSelected(from);
    setSelectionEnd(to);
    setRange(rectangle(from, to));
  };
  const dragStart = useCallback(
    (cell: string, extend: boolean): void => {
      const paintedFormat = paintStyleRef.current;
      if (paintedFormat !== null) {
        const expected = paintHomeBinding.current,
          current = homeFrameRef.current?.binding;
        if (
          expected &&
          current &&
          expected.documentId === current.documentId &&
          expected.ownerIdentity === current.ownerIdentity &&
          expected.state === current.state &&
          expected.sheet === current.sheet &&
          expected.navigationGeneration === current.navigationGeneration &&
          expected.activationEpoch === current.activationEpoch &&
          expected.modeGeneration === current.modeGeneration &&
          expected.specialSelectionEpoch === current.specialSelectionEpoch
        )
          homePaintRunner.current?.(cell, paintedFormat);
        else setGridError('The formatting source changed. Use Format Painter again.');
        paintStyleRef.current = null;
        setPaintStyleId(null);
        select(cell);
        return;
      }
      dragging.current = true;
      select(cell, extend);
    },
    [select]
  );
  const dragEnter = useCallback(
    (cell: string): void => {
      const fill = fillDragging.current;
      if (fill) {
        const source = selectedBounds(fill.source, 1000),
          position = addressPosition(cell);
        const target =
          position.row > source.last.row && position.col >= source.first.col && position.col <= source.last.col
            ? `${colName(source.first.col)}${source.last.row + 2}:${colName(source.last.col)}${position.row + 1}`
            : position.col > source.last.col && position.row >= source.first.row && position.row <= source.last.row
              ? `${colName(source.last.col + 1)}${source.first.row + 1}:${colName(position.col)}${source.last.row + 1}`
              : null;
        fill.target = target;
        setFillPreview(target);
        return;
      }
      if (dragging.current) select(cell, true);
    },
    [select]
  );
  return { setValue, select, selectRow, selectColumn, dragStart, dragEnter } as const;
}
