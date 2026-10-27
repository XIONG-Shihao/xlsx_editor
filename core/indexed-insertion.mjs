/** Inserting and deleting rows and columns. The engine moves the cells, the
 * formulas, every record inside a worksheet and everything beside the
 * worksheets (names, tables, notes, pictures, charts). Here the request is
 * checked for admission and handed over: sheet, rows or columns, how many. */
import fs from 'node:fs/promises';
import path from 'node:path';
import { assertStructuralEdit } from './workbook-package.mjs';
import { refusal } from './refusal-codes.mjs';
export async function prepareIndexedInsertion(input, native) {
  const { root, version, command } = input;
  assertStructuralEdit(await fs.readFile(path.join(root, 'metadata.xlsx')), command, { objects: true, scoped: true });
  const stored = await native({
    action: 'workbook',
    operation: 'plan-references',
    root,
    version,
    command,
    ...(input.rewriteInputId === undefined ? {} : { source: input.source, rewriteInputId: input.rewriteInputId }),
  });
  if (!stored.supported)
    throw refusal(
      'workbook-unsupported',
      `Structural reference syntax is not qualified for native preservation: ${stored.reason ?? 'metadata binding'}`
    );
  return {
    headers: stored.headers ?? [],
    prepared: true,
    nativeReferences: true,
    nativeStructuralPlan: true,
    formulaCount: stored.formulaCount ?? 0,
    inputCells: 0,
  };
}
