import { createWorkbook, editWorkbook } from '../../core/workbook-package.mjs';
import { zipSync, unzipSync } from 'fflate';
export const initialPivotCommand = {
  operation: 'createPivot',
  sheet: 'Sheet1',
  name: 'SalesPivot',
  range: 'A1:D5',
  rowField: 0,
  valueField: 1,
  aggregate: 'sum',
  pivotSheet: 'Summary',
  expectedPivotRevision: null,
};
export function pivotFieldFixture() {
  let bytes = editWorkbook(createWorkbook(), [{ operation: 'addSheet', sheet: 'Alternate' }]);
  const first = [
    ['Category', 'Amount', 'Region', 'Units'],
    ['A', 2, 'North', 10],
    ['B', 5, 'South', 20],
    ['A', -1, 'North', 30],
    ['C', 4, 'West', null],
    ['C', 8, 'East', 40],
  ];
  const other = [
    ['Category', 'Amount', 'Region', 'Units'],
    ['Z', 9, 'Central', 1],
    ['Y', 3, 'Central', 2],
    ['Y', 2, 'Rural', null],
    ['Z', 1, 'Rural', 4],
  ];
  bytes = editWorkbook(bytes, [
    ...first.flatMap((row, r) => row.map((value, c) => ({ sheet: 'Sheet1', cell: 'ABCD'[c] + (r + 1), value }))),
    ...other.flatMap((row, r) => row.map((value, c) => ({ sheet: 'Alternate', cell: 'BCDE'[c] + (r + 2), value }))),
    { sheet: 'Sheet1', cell: 'J10', value: 'Outside' },
  ]);
  const parts = unzipSync(bytes);
  parts['custom/opaque.bin'] = new Uint8Array([1, 3, 5, 255]);
  return Buffer.from(zipSync(parts));
}
