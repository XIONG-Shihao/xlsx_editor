// How often does the general (Univer) calculator agree with the results Excel
// stored? Reads each workbook, calculates it, and compares every formula that
// has a stored result. A measurement, not a pass/fail test.
//
//   node --import ./scripts/ts-resolve-register.mjs scripts/audit/calculator-agreement.mjs FILE...
import fs from 'node:fs/promises';
import path from 'node:path';
import { readWorkbook } from '../../core/workbook-package.mjs';
import * as engine from '../../core/calculation-node.mjs';

let same = 0,
  different = 0,
  unresolved = 0;
for (const file of process.argv.slice(2)) {
  let view;
  try {
    view = readWorkbook(await fs.readFile(file));
  } catch (error) {
    console.log(path.basename(file).padEnd(36), 'could not be read:', String(error.message).slice(0, 80));
    continue;
  }
  let calculated;
  try {
    calculated = await engine.calculateWorkbook(view);
  } catch (error) {
    console.log(path.basename(file).padEnd(36), 'could not be calculated:', String(error.message).slice(0, 80));
    continue;
  }
  const row = { same: 0, different: 0, unresolved: 0, examples: [] };
  for (const sheet of calculated.sheets)
    for (const [reference, cell] of Object.entries(sheet.cells)) {
      if (!cell.formula || cell.cachedValue === undefined || cell.cachedValue === null) continue;
      const status = cell.calculation?.status,
        got = status === 'error' ? cell.calculation.error : cell.calculation?.value,
        want = cell.cachedValue;
      if (status !== 'calculated' && status !== 'error') {
        row.unresolved += 1;
        if (row.examples.length < 2)
          row.examples.push(
            `${sheet.name}!${reference} =${cell.formula.slice(0, 40)} not calculated: ${cell.calculation?.error ?? status}`
          );
      } else if (
        typeof got === 'number' && typeof want === 'number'
          ? Math.abs(got - want) <= 1e-9 * Math.max(1, Math.abs(want))
          : String(got) === String(want)
      )
        row.same += 1;
      else {
        row.different += 1;
        if (row.examples.length < 2)
          row.examples.push(
            `${sheet.name}!${reference} =${cell.formula.slice(0, 40)} excel=${String(want).slice(0, 20)} got=${String(got).slice(0, 20)}`
          );
      }
    }
  same += row.same;
  different += row.different;
  unresolved += row.unresolved;
  if (row.same + row.different + row.unresolved)
    console.log(
      path.basename(file).padEnd(36),
      `agree ${String(row.same).padStart(5)}  differ ${String(row.different).padStart(4)}  not calculated ${String(row.unresolved).padStart(4)}`
    );
  for (const example of row.examples) console.log('      ', example);
}
const total = same + different + unresolved;
console.log(
  `\n${total} formulas: ${same} agree (${((same / Math.max(1, total)) * 100).toFixed(1)}%), ${different} differ, ${unresolved} not calculated`
);
process.exit(0);
