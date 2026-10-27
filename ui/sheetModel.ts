import { displayCell, displayCellDetails, type WorkbookDisplayedText } from '../core/display.mjs';
import { hasExactNumericValue } from '../core/calculation.mjs';
import type { WorkbookCell, WorkbookEdit, WorkbookSheet } from '../core/contracts';
import { parseCellInput } from '../core/cell-input.mjs';
export function colName(index: number): string {
  let name = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}
export function addressPosition(ref: string): {
  row: number;
  col: number;
} {
  const match = /^([A-Z]+)([1-9]\d*)$/i.exec(ref);
  if (!match) throw Error('Invalid cell address');
  const position = {
    col: [...match[1].toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1,
    row: Number(match[2]) - 1,
  };
  if (!Number.isSafeInteger(position.row) || !Number.isSafeInteger(position.col)) throw Error('Invalid cell address');
  return position;
}
export function maxColumnCount(rows: readonly (readonly unknown[])[]): number {
  let maximum = 1;
  for (const row of rows) maximum = Math.max(maximum, row.length);
  return maximum;
}
export function applyGridEdit(sheet: WorkbookSheet, edit: WorkbookEdit): WorkbookSheet {
  if (edit.type === 'clearHyperlinks')
    throw Error('Clear Hyperlinks requires its accepted native XLSX owner view; browser projection is unsupported');
  const cells = edit.operation ? { ...sheet.cells } : patchCells(sheet.cells, edit.cell!);
  if (edit.cell && edit.styleId === 0) {
    const cleaned = { ...cells[edit.cell] };
    for (const key of [
      'bold',
      'italic',
      'underline',
      'fontFamily',
      'fontSize',
      'fontColor',
      'fillColor',
      'horizontalAlign',
      'verticalAlign',
      'wrapText',
      'border',
    ] as const)
      delete cleaned[key];
    cells[edit.cell] = { ...cleaned, format: 0 };
  }
  if (edit.cell)
    cells[edit.cell] = {
      ...cells[edit.cell],
      ...(edit.value !== undefined ? { value: edit.value, formula: undefined, calculation: undefined } : {}),
      ...(edit.formula !== undefined ? { formula: edit.formula, calculation: undefined } : {}),
      ...(edit.bold !== undefined ? { bold: edit.bold } : {}),
      ...(edit.italic !== undefined ? { italic: edit.italic } : {}),
      ...(edit.underline !== undefined ? { underline: edit.underline } : {}),
      ...(edit.fontFamily !== undefined ? { fontFamily: edit.fontFamily } : {}),
      ...(edit.fontSize !== undefined ? { fontSize: edit.fontSize } : {}),
      ...(edit.fontColor !== undefined ? { fontColor: edit.fontColor } : {}),
      ...(edit.fillColor !== undefined ? { fillColor: edit.fillColor } : {}),
      ...(edit.horizontalAlign !== undefined ? { horizontalAlign: edit.horizontalAlign } : {}),
      ...(edit.verticalAlign !== undefined ? { verticalAlign: edit.verticalAlign } : {}),
      ...(edit.wrapText !== undefined ? { wrapText: edit.wrapText } : {}),
      ...(edit.border !== undefined ? { border: edit.border } : {}),
      ...(edit.styleId !== undefined ? { styleId: edit.styleId } : {}),
      ...(edit.numberFormat !== undefined
        ? {
            format: {
              general: 0,
              number: 2,
              percent: 10,
              date: 14,
              text: 49,
              currency: '$#,##0.00',
              accounting: 44,
              thousands: 3,
              decimal0: 1,
              decimal1: '0.0',
              decimal2: 2,
            }[edit.numberFormat],
          }
        : {}),
    };
  if (edit.operation) {
    const row = edit.operation.endsWith('Row'),
      remove = edit.operation.startsWith('delete'),
      index = edit.index ?? 0;
    for (const key of Object.keys(cells)) delete cells[key];
    for (const [key, value] of Object.entries(sheet.cells)) {
      const p = addressPosition(key),
        n = row ? p.row : p.col;
      if (remove && n === index) continue;
      if (n >= index) {
        if (row) p.row += remove ? -1 : 1;
        else p.col += remove ? -1 : 1;
      }
      cells[`${colName(p.col)}${p.row + 1}`] = value;
    }
  }
  let rows = sheet.rows,
    columns = sheet.columns;
  if (edit.operation) {
    for (const ref of Object.keys(cells)) {
      const p = addressPosition(ref);
      rows = Math.max(rows, p.row + 1);
      columns = Math.max(columns, p.col + 1);
    }
  } else if (edit.cell) {
    const p = addressPosition(edit.cell);
    rows = Math.max(rows, p.row + 1);
    columns = Math.max(columns, p.col + 1);
  }
  const next = { ...sheet, cells, rows, columns };
  if (!edit.operation && edit.cell) {
    formulaCounts.set(
      next,
      formulaCount(sheet) - Number(!!sheet.cells[edit.cell]?.formula) + Number(!!cells[edit.cell]?.formula)
    );
  }
  return next;
}
export function literalDisplay(cell?: WorkbookCell): string | number | boolean {
  if ((cell?.scenarioResult || cell?.formulaType) && cell.calculation?.status === 'calculated')
    return cell.calculation.value ?? '';
  if (typeof cell?.value === 'number' && !hasExactNumericValue(cell.value, cell.rawValue)) return cell.rawValue!;
  return cell?.value ?? '';
}
export function inputDisplay(cell?: WorkbookCell): string {
  return cell?.formula ? '=' + cell.formula : String(literalDisplay(cell));
}
export function displayDetails(sheet: WorkbookSheet, ref: string): WorkbookDisplayedText {
  return displayCellDetails(sheet.cells[ref], sheet.date1904);
}
export function displayValue(sheet: WorkbookSheet, ref: string): string {
  return displayDetails(sheet, ref).text;
}
// Sparse immutable overlays share the original record. Preparing a large sheet
// never creates another dictionary containing every original cell.
interface CellOverlay {
  base: Record<string, WorkbookCell>;
  buckets: Record<string, WorkbookCell>[];
}
const overlays = new WeakMap<Record<string, WorkbookCell>, CellOverlay>();
const proxyRecords = new WeakSet<Record<string, WorkbookCell>>();
function bucket(ref: string): number {
  let hash = 0;
  for (let i = 0; i < ref.length; i++) hash = (hash * 31 + ref.charCodeAt(i)) | 0;
  return hash & 255;
}
export function prepareCellRecord(previous: Record<string, WorkbookCell>): void {
  if (!overlays.has(previous))
    overlays.set(previous, { base: previous, buckets: Array.from({ length: 256 }, () => Object.create(null)) });
}
export function patchCellRecord(
  previous: Record<string, WorkbookCell>,
  patches: Record<string, WorkbookCell>
): Record<string, WorkbookCell> {
  prepareCellRecord(previous);
  const prior = overlays.get(previous)!,
    next = prior.buckets.slice(),
    changed = new Set<number>();
  for (const [ref, value] of Object.entries(patches)) {
    const index = bucket(ref);
    if (!changed.has(index)) {
      next[index] = { ...next[index] };
      changed.add(index);
    }
    next[index][ref] = value;
  }
  const value = (key: string) => (Object.hasOwn(next[bucket(key)], key) ? next[bucket(key)][key] : prior.base[key]);
  const cells = new Proxy(Object.create(null) as Record<string, WorkbookCell>, {
    get: (_target, key) => (typeof key === 'string' ? value(key) : undefined),
    ownKeys: () => {
      const keys = Object.keys(prior.base);
      for (const part of next) for (const key of Object.keys(part)) if (!Object.hasOwn(prior.base, key)) keys.push(key);
      return keys;
    },
    getOwnPropertyDescriptor: (_target, key) =>
      typeof key === 'string' && (Object.hasOwn(next[bucket(key)], key) || Object.hasOwn(prior.base, key))
        ? { configurable: true, enumerable: true, value: value(key) }
        : undefined,
  });
  overlays.set(cells, { base: prior.base, buckets: next });
  proxyRecords.add(cells);
  return cells;
}
function patchCells(previous: Record<string, WorkbookCell>, ref: string): Record<string, WorkbookCell> {
  const cells = patchCellRecord(previous, {}),
    state = overlays.get(cells)!,
    index = bucket(ref);
  state.buckets[index] = { ...state.buckets[index] };
  const writable = new Proxy(cells, {
    set: (_target, key, value) => {
      if (key !== ref) throw Error('Invalid cell patch');
      state.buckets[index][ref] = value;
      return true;
    },
  });
  overlays.set(writable, state);
  proxyRecords.add(writable);
  return writable;
}
export function applyCalculatedPatches(
  sheet: WorkbookSheet,
  patches: Record<string, WorkbookCell>,
  revision: string
): WorkbookSheet {
  let count = formulaCount(sheet),
    rows = sheet.rows,
    columns = sheet.columns;
  for (const [ref, cell] of Object.entries(patches)) {
    count +=
      Number(!!(cell.formula || cell.formulaType)) -
      Number(!!(sheet.cells[ref]?.formula || sheet.cells[ref]?.formulaType));
    const p = addressPosition(ref);
    rows = Math.max(rows, p.row + 1);
    columns = Math.max(columns, p.col + 1);
  }
  const next = { ...sheet, cells: patchCellRecord(sheet.cells, patches), rows, columns, calculationRevision: revision };
  formulaCounts.set(next, count);
  return next;
}
export function rememberFormulaCount(sheet: WorkbookSheet, count: number): void {
  formulaCounts.set(sheet, count);
}
const formulaCounts = new WeakMap<WorkbookSheet, number>();
export function formulaCount(sheet: WorkbookSheet): number {
  if (sheet.storedFormulaCount !== undefined) return sheet.storedFormulaCount;
  let count = formulaCounts.get(sheet);
  if (count === undefined) {
    count = 0;
    for (const ref in sheet.cells) if (sheet.cells[ref].formula || sheet.cells[ref].formulaType) count++;
    formulaCounts.set(sheet, count);
  }
  return count;
}
/** Calculation is authoritative only when supplied by the shared worker/service. */
export function formulaDisplay(sheet: WorkbookSheet, ref: string): string | number | boolean {
  const cell = sheet.cells[ref];
  if (!cell?.formula && !cell?.formulaType && !cell?.scenarioResult) return literalDisplay(cell);
  const result = cell.calculation;
  if (result?.pending) return displayCell(cell, sheet.date1904);
  if (!result) return 'Calculating…';
  if (result.status === 'calculated' || result.status === 'stored') return result.value ?? result.error ?? '';
  if (result.status === 'circular') return '#CYCLE!';
  if (result.status === 'error') return result.error ?? '#ERROR!';
  return 'Unresolved';
}
export function formulaRowVersion(_sheet: WorkbookSheet, _row: number): number {
  return 0;
}
/** Materialize the structural-sharing facade only at the structured-clone boundary. */
export function plainSheets(sheets: WorkbookSheet[]): WorkbookSheet[] {
  return sheets.map((sheet) => ({ ...sheet, cells: proxyRecords.has(sheet.cells) ? { ...sheet.cells } : sheet.cells }));
}
export const parseGridInput = parseCellInput;
