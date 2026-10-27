import { performance } from 'node:perf_hooks';
import { formulaDisplay, applyGridEdit } from '@office-editors/xlsx-editor/ui/sheetModel';
import { calculateWorkbook } from '@office-editors/xlsx-editor/core/calculation-node.mjs';
import * as before from '../../fixtures/spreadsheet-model-before.ts';
const sheet = (cells) => ({
  name: 'Sheet1',
  state: 'visible',
  path: 'xl/worksheets/sheet1.xml',
  cells,
  rows: 10000,
  columns: 8,
  merges: [],
  widths: [],
});
const cells = { A1: { value: 1 } };
for (let i = 2; i <= 17; i++) cells[`A${i}`] = { formula: `A${i - 1}+A${i - 1}` };
const s = sheet(cells),
  times = { before: [], after: [] };
for (let n = 0; n < 11; n++)
  for (const kind of n % 2 ? ['after', 'before'] : ['before', 'after']) {
    const start = performance.now();
    const result =
      kind === 'before'
        ? before.formulaDisplay(s, 'A17')
        : formulaDisplay((await calculateWorkbook({ revision: `bench-${n}`, sheets: [s] })).sheets[0], 'A17');
    if (result !== 65536) throw Error('Parity');
    const ms = performance.now() - start;
    times[kind].push(ms);
  }
let big = sheet(Object.fromEntries(Array.from({ length: 10000 }, (_, i) => [`A${i + 1}`, { value: i }])));
big = applyGridEdit(big, { cell: 'A1', value: 0 });
const edits = [];
for (let i = 0; i < 100; i++) {
  const t = performance.now();
  big = applyGridEdit(big, { cell: 'A1', value: i });
  edits.push(performance.now() - t);
}
console.log(
  JSON.stringify(
    {
      kind: 'Historical recursive evaluation vs current headless calculation plus result display; same 17-cell chain and result. Includes current calculation-engine setup; excludes React paint and IPC',
      node: process.version,
      formula: times,
      edit10000CellsMs: edits,
    },
    null,
    2
  )
);
