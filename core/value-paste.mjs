/** Current Values snapshots. A caller must supply the captured accepted owner
 * read; this pure helper does not certify a native spill or manufacture caches. */
import { hasExactNumericValue } from './calculation.mjs';
import { cellPosition, rangeBounds } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';
export const VALUE_PASTE_ERRORS = Object.freeze([
  '#DIV/0!',
  '#VALUE!',
  '#NUM!',
  '#REF!',
  '#N/A',
  '#NAME?',
  '#NULL!',
  '#SPILL!',
  '#CALC!',
]);
const sourceOwners = Symbol('accepted Values source geometry');
/** Call once with the complete accepted source sheet, never a viewport cache. */
export function createValuePasteSourceContext(sheet) {
  if (!sheet || typeof sheet.name !== 'string' || !sheet.cells || typeof sheet.cells !== 'object')
    throw refusal('paste-bad-input', 'Values needs complete accepted source metadata');
  const entries = Object.entries(sheet.cells),
    formulaCount = entries.filter(([, cell]) => cell.formula || cell.formulaType).length;
  if (Number.isInteger(sheet.storedFormulaCount) && sheet.storedFormulaCount !== formulaCount)
    throw refusal('paste-unsupported', 'Values source metadata is incomplete');
  const owners = entries
    .filter(([, cell]) => cell.formulaType && cell.formulaReference)
    .map(([master, cell]) =>
      Object.freeze({
        master,
        range: cell.formulaReference,
        type: cell.formulaType,
        bounds: rangeBounds(cell.formulaReference),
        spill: cell.spill ? structuredClone(cell.spill) : null,
        scenario: cell.scenarioResult === true,
      })
    );
  return Object.freeze({ sheet, [sourceOwners]: Object.freeze(owners) });
}
export function validateValueType(edit) {
  if (
    Object.hasOwn(edit, 'valueType') &&
    (edit.valueType !== 'error' ||
      typeof edit.value !== 'string' ||
      !VALUE_PASTE_ERRORS.includes(edit.value) ||
      Object.hasOwn(edit, 'formula'))
  )
    throw refusal('internal-invalid-field', 'Invalid declared error value');
  return edit;
}
export function normalizeValuePasteCell(cell, context) {
  if (context !== undefined) {
    if (!context || typeof context.cell !== 'string')
      throw refusal('internal-invalid-field', 'Invalid Values source context');
    const prepared = context[sourceOwners] ? context : createValuePasteSourceContext(context.sheet),
      at = cellPosition(context.cell);
    for (const owner of prepared[sourceOwners]) {
      const { first, last } = owner.bounds;
      if (at.row < first.row || at.row > last.row || at.col < first.col || at.col > last.col) continue;
      const owned =
        owner.spill?.schema === 1 &&
        owner.spill.master === owner.master &&
        owner.spill.range === owner.range &&
        cell?.spill?.schema === 1 &&
        cell.spill.master === owner.master &&
        cell.spill.range === owner.range;
      const scenario = owner.type === 'dataTable' && owner.scenario && cell?.scenarioResult === true;
      if (!owned && !scenario)
        throw refusal('paste-bad-input', 'Paste Values source output has no current scalar ownership proof');
    }
  }
  if (cell === undefined || cell === null) return Object.freeze({ value: null });
  if (typeof cell !== 'object' || Array.isArray(cell))
    throw refusal('internal-invalid-field', 'Invalid Values source cell');
  const derived = Boolean(cell.formula || cell.formulaType || cell.formulaOutput || cell.spill || cell.scenarioResult),
    calculation = cell.calculation;
  if (calculation?.pending === true)
    throw refusal('paste-bad-input', 'Paste Values needs a current calculated result; source is pending');
  let value, type;
  if (derived) {
    if (calculation?.status === 'calculated' && Object.hasOwn(calculation, 'value')) value = calculation.value;
    else if (calculation?.status === 'error' && VALUE_PASTE_ERRORS.includes(calculation.error)) {
      value = calculation.error;
      type = 'error';
    } else throw refusal('paste-bad-input', 'Paste Values needs a current calculated result');
  } else {
    value = cell.value ?? null;
    if (cell.dataType === 'error') {
      type = 'error';
      if (!VALUE_PASTE_ERRORS.includes(value)) throw refusal('internal-invalid-field', 'Invalid typed Values error');
    }
  }
  if (value !== null && !['string', 'number', 'boolean'].includes(typeof value))
    throw refusal('paste-bad-input', 'Values source is not a scalar');
  if (
    typeof value === 'number' &&
    (!Number.isFinite(value) || (!derived && !hasExactNumericValue(value, cell.rawValue)))
  )
    throw refusal('paste-refused', 'Paste Values would lose stored numeric precision');
  if (
    typeof value === 'string' &&
    (value.length > 32767 || /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/u.test(value))
  )
    throw refusal('internal-invalid-field', 'Invalid Values text');
  return Object.freeze({ value, ...(type ? { valueType: type } : {}) });
}
