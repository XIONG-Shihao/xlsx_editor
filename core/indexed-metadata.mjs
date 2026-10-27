import { structuralReasonsFromEngine } from './workbook-package.mjs';
import { refusal } from './refusal-codes.mjs';

/** Calendar flags are semantics, never truthy/coerced cache identifiers. */
export function indexedCalendarGuards(view) {
  for (const sheet of view.sheets)
    if (sheet.date1904 !== undefined && typeof sheet.date1904 !== 'boolean')
      throw refusal('internal-invalid-field', 'Invalid workbook date-system metadata.');
}

/** Apply the same bounded UI capability projection to imported and saved
 * checkpoints. Metadata keeps object protections; current native counts add
 * formula protections without inheriting stale flags from an earlier version. */
export function indexedViewGuards(view, answer, sheets, metadataView = view) {
  view.structuralReason = metadataView.structuralReason;
  let reasons = null;
  try {
    reasons = structuralReasonsFromEngine(answer, { objects: true, scoped: true });
  } catch {
    // A package that cannot be inspected offers no structural editing.
  }
  for (const sheet of view.sheets) {
    // Shared formulas are written out in full when the structure changes, so
    // only this sheet's own extension records hold it back. Array formulas and
    // extension records elsewhere that name this sheet are checked when the
    // edit is made, where the answer can be exact.
    const special = sheets.some((item) => item.path === sheet.path && item.extended);
    const baseline = metadataView.sheets.find((item) => item.path === sheet.path);
    if (baseline?.capabilities) sheet.capabilities = { ...baseline.capabilities };
    else delete sheet.capabilities;
    if (reasons?.get(sheet.name) === null) sheet.capabilities = { insertRows: true, insertRowsReason: null };
    if (special)
      sheet.capabilities = {
        insertRows: false,
        insertRowsReason:
          'This sheet has extension records that cannot yet follow inserted or deleted rows and columns',
      };
    if (sheet.capabilities)
      Object.assign(sheet.capabilities, {
        deleteRows: sheet.capabilities.insertRows,
        referenceColumns: sheet.capabilities.insertRows,
      });
    const native = sheets.find((item) => item.path === sheet.path);
    if (native?.formulaCount || native?.storedFormulaCount)
      view.structuralReason ??= 'Row/column changes require qualified formula-reference updates.';
  }
  return view;
}
