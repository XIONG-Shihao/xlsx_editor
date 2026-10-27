import { createRequire } from 'node:module';
import {
  createCalculationSessionWith,
  calculateWorkbookWith,
  shiftFormulaWith,
  transformWorkbookWith,
  shiftFormulasWith,
  formulaReferencesWith,
  calculationInputRangesWith,
  calculationDependenciesWith,
} from './calculation.mjs';
const require = createRequire(import.meta.url);
let libraries;
function bindings() {
  if (!libraries) {
    libraries = Object.assign(
      {},
      ...['core', 'engine-formula', 'sheets', 'sheets-formula', 'sheets-table'].map((name) =>
        require('@univerjs/' + name)
      ),
      require('@univerjs/core/facade')
    );
    for (const name of ['engine-formula', 'sheets', 'sheets-formula', 'sheets-table'])
      require('@univerjs/' + name + '/facade');
  }
  return libraries;
}
export const calculateWorkbook = (view, options) => calculateWorkbookWith(bindings(), view, options);
export const shiftFormula = (formula, columnOffset, rowOffset) =>
  shiftFormulaWith(bindings(), formula, columnOffset, rowOffset);

export const transformWorkbook = (view, edits) => transformWorkbookWith(bindings(), view, edits);

export const shiftFormulas = (requests) => shiftFormulasWith(bindings(), requests);
export const formulaReferences = (formula) => formulaReferencesWith(bindings(), formula);

export const createCalculationSession = (view, options) => createCalculationSessionWith(bindings(), view, options);
export const calculationInputRanges = (view) => calculationInputRangesWith(bindings(), view);

export const calculationDependencies = (view, options) => calculationDependenciesWith(bindings(), view, options);
