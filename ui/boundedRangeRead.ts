import type { WorkbookCell } from '../core/contracts';
import type { ReadRange } from './rangeCells';
import { addressPosition, colName } from './sheetModel';
/** Adapt viewport tiles and navigation to a reader's strict per-call budget. */
export function boundedRangeRead(read: ReadRange, positions = 1000): ReadRange {
  return async (sheet, range) => {
    const [first, last = first] = range.split(':'),
      a = addressPosition(first),
      b = addressPosition(last),
      width = b.col - a.col + 1,
      height = b.row - a.row + 1;
    if (
      !Number.isInteger(positions) ||
      positions < 1 ||
      positions > 1000 ||
      width < 1 ||
      height < 1 ||
      width * height > 10000
    )
      throw Error('Passive grid read exceeds its 10,000-position viewport/selection budget');
    const result: Record<string, WorkbookCell> = {};
    for (let c = a.col; c <= b.col; c += positions) {
      const endColumn = Math.min(b.col, c + positions - 1),
        rows = Math.max(1, Math.floor(positions / (endColumn - c + 1)));
      for (let r = a.row; r <= b.row; r += rows) {
        const endRow = Math.min(b.row, r + rows - 1),
          cells = await read(sheet, colName(c) + (r + 1) + ':' + colName(endColumn) + (endRow + 1));
        Object.assign(result, cells);
        if (new TextEncoder().encode(JSON.stringify(result)).length > 1024 * 1024)
          throw Error('Passive grid read exceeds its 1 MiB combined reply budget');
      }
    }
    return result;
  };
}
