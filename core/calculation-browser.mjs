import * as core from '@univerjs/core';
import * as formula from '@univerjs/engine-formula';
import * as sheets from '@univerjs/sheets';
import * as sheetsFormula from '@univerjs/sheets-formula';
import * as tables from '@univerjs/sheets-table';
import { FUniver } from '@univerjs/core/facade';
import '@univerjs/engine-formula/facade';
import '@univerjs/sheets/facade';
import '@univerjs/sheets-formula/facade';
import '@univerjs/sheets-table/facade';
import {
  createCalculationSessionWith,
  calculateWorkbookWith,
  shiftFormulaWith,
  transformWorkbookWith,
  shiftFormulasWith,
  formulaReferencesWith,
} from './calculation.mjs';
const libraries = { ...core, ...formula, ...sheets, ...sheetsFormula, ...tables, FUniver };
export const calculateWorkbook = (view, options) => calculateWorkbookWith(libraries, view, options);
export const shiftFormula = (formula, columnOffset, rowOffset) =>
  shiftFormulaWith(libraries, formula, columnOffset, rowOffset);

export const transformWorkbook = (view, edits) => transformWorkbookWith(libraries, view, edits);

export const shiftFormulas = (requests) => shiftFormulasWith(libraries, requests);
export const formulaReferences = (formula) => formulaReferencesWith(libraries, formula);

export const createCalculationSession = (view, options) => createCalculationSessionWith(libraries, view, options);
