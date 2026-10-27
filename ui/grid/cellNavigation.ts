import { type WorkbookGeometryProfile, type WorkbookSheet } from '../../core/contracts';
import { addressPosition, colName } from '../sheetModel';
import { createWorksheetDimensions } from '../../core/worksheet-geometry.mjs';

export function mergeAnchor(sheet: WorkbookSheet | undefined, ref: string): string {
  const position = addressPosition(ref);
  for (const range of sheet?.merges ?? []) {
    const [start, end = start] = range.split(':');
    const first = addressPosition(start),
      last = addressPosition(end);
    if (position.row >= first.row && position.row <= last.row && position.col >= first.col && position.col <= last.col)
      return start;
  }
  return ref;
}
export function nextNavigableCell(
  sheet: WorkbookSheet,
  from: { row: number; col: number },
  rowStep: number,
  colStep: number,
  geometryProfile?: WorkbookGeometryProfile
): { row: number; col: number; ref: string } {
  const merges = sheet.merges.map((range) => {
    const [start, end] = range.split(':');
    return { start: addressPosition(start), end: addressPosition(end ?? start) };
  });
  const containing = (row: number, col: number) =>
    merges.find(
      (merge) => row >= merge.start.row && row <= merge.end.row && col >= merge.start.col && col <= merge.end.col
    );
  const dimensions = createWorksheetDimensions(sheet, geometryProfile);
  const hiddenRow = (row: number) => dimensions.rowHeightPt(row) === 0;
  const hiddenColumn = (col: number) => dimensions.columnPixels(col) === 0;
  const visibleOverrides = sheet.sheetFormat?.rowsHiddenByDefault
    ? [...(sheet.visibleRows ?? [])].filter((row) => !hiddenRow(row)).sort((a, b) => a - b)
    : null;
  let { row, col } = from;
  const currentMerge = containing(row, col);
  if (currentMerge) {
    if (rowStep > 0) row = currentMerge.end.row;
    else if (rowStep < 0) row = currentMerge.start.row;
    if (colStep > 0) col = currentMerge.end.col;
    else if (colStep < 0) col = currentMerge.start.col;
  }
  for (;;) {
    row += rowStep;
    col += colStep;
    if (row < 0 || row > 1_048_575 || col < 0 || col > 16_383)
      return { ...from, ref: `${colName(from.col)}${from.row + 1}` };
    if (hiddenRow(row)) {
      if (rowStep && visibleOverrides) {
        const next =
          rowStep > 0
            ? visibleOverrides.find((index) => index > row)
            : [...visibleOverrides].reverse().find((index) => index < row);
        if (next === undefined) return { ...from, ref: `${colName(from.col)}${from.row + 1}` };
        row = next - rowStep;
      }
      continue;
    }
    if (hiddenColumn(col)) continue;
    const merge = containing(row, col);
    if (merge) {
      if (hiddenRow(merge.start.row) || hiddenColumn(merge.start.col)) {
        if (rowStep > 0) row = merge.end.row;
        else if (rowStep < 0) row = merge.start.row;
        if (colStep > 0) col = merge.end.col;
        else if (colStep < 0) col = merge.start.col;
        continue;
      }
      row = merge.start.row;
      col = merge.start.col;
    }
    return { row, col, ref: `${colName(col)}${row + 1}` };
  }
}
