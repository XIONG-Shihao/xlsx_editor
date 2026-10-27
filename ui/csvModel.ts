import { csvCellInfo, type CsvIndex } from '../core/csv-workbook.mjs';
import type { WorkbookCell, WorkbookSheet } from '../core/contracts';
import { addressPosition, colName, rememberFormulaCount } from './sheetModel';
/** A viewport-sized cache over the shared indexed source. No complete cell map. */
export function csvSheet(index: CsvIndex): WorkbookSheet {
  const cache = new Map<string, WorkbookCell>();
  const cell = (key: string): WorkbookCell | undefined => {
    if (cache.has(key)) return cache.get(key);
    if (!/^[A-Z]+[1-9]\d*$/.test(key)) return undefined;
    const p = addressPosition(key),
      value = csvCellInfo(index, p.row, p.col);
    if (!value) return undefined;
    cache.set(key, value);
    while (cache.size > 4096) cache.delete(cache.keys().next().value!);
    return value;
  };
  const cells = new Proxy(Object.create(null) as Record<string, WorkbookCell>, {
    get: (_target, key) => (typeof key === 'string' ? cell(key) : undefined),
    // Enumeration is complete if explicitly requested; ordinary grid/history
    // operations only request addressed cells and never enumerate the source.
    ownKeys: () => {
      const keys: string[] = [];
      for (let row = 0; row < index.rows; row++)
        for (let col = 0; col < (index.widthOverrides[row] ?? index.rowWidths[row] ?? 1); col++)
          keys.push(`${colName(col)}${row + 1}`);
      return keys;
    },
    getOwnPropertyDescriptor: (_target, key) =>
      typeof key === 'string' && cell(key) !== undefined
        ? { configurable: true, enumerable: true, value: cell(key) }
        : undefined,
  });
  const sheet: WorkbookSheet = {
    name: 'CSV',
    state: 'visible',
    path: '',
    cells,
    rows: index.rows,
    columns: index.columns,
    merges: [],
    widths: [],
  };
  rememberFormulaCount(sheet, 0);
  return sheet;
}
