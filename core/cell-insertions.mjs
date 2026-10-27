/** Explicit editor Boolean-cell profile and finite literal Symbols samples. */
import { refusal } from './refusal-codes.mjs';

export const symbolSamples = Object.freeze(
  [
    { id: 'omega', label: 'Omega', value: 'Ω' },
    { id: 'plus-minus', label: 'Plus/minus', value: '±' },
    { id: 'degree', label: 'Degree', value: '°' },
  ].map(Object.freeze)
);
export function booleanCellCommand(value) {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).sort().join(',') !== 'action,cell,expectedValue,sheet,type' ||
    value.type !== 'booleanCell' ||
    !['insert', 'toggle', 'remove'].includes(value.action)
  )
    throw refusal('internal-invalid-field', 'Invalid Boolean-cell command.');
  if (
    typeof value.sheet !== 'string' ||
    !value.sheet.trim() ||
    value.sheet.length > 31 ||
    /[\\/?*\[\]:\u0000-\u001f]/u.test(value.sheet)
  )
    throw refusal('internal-invalid-field', 'Invalid Boolean-cell worksheet.');
  const m = typeof value.cell === 'string' && /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(value.cell);
  let column = 0;
  if (m) for (const c of m[1]) column = column * 26 + c.charCodeAt(0) - 64;
  if (!m || column > 16384 || Number(m[2]) > 1048576)
    throw refusal('workbook-bad-input', 'Boolean-cell requires one canonical cell.');
  if (
    (value.expectedValue !== null && typeof value.expectedValue !== 'boolean') ||
    (value.action === 'insert' && value.expectedValue !== null) ||
    (value.action === 'toggle' && typeof value.expectedValue !== 'boolean')
  )
    throw refusal('workbook-bad-input', 'Boolean-cell action requires a blank or literal Boolean target.');
  return {
    type: 'booleanCell',
    action: value.action,
    sheet: value.sheet,
    cell: value.cell,
    expectedValue: value.expectedValue,
  };
}
export function normalizeBooleanCellCommands(commands) {
  if (!Array.isArray(commands) || !commands.length || commands.length > 10000)
    throw refusal('internal-invalid-field', 'Invalid Boolean-cell batch.');
  const captured = commands.map((c) => (c?.type === 'booleanCell' ? booleanCellCommand(c) : structuredClone(c)));
  if (new TextEncoder().encode(JSON.stringify(captured)).byteLength > 2 * 1024 * 1024)
    throw refusal('workbook-limit', 'Boolean-cell batch is too large.');
  return captured;
}
export function symbolCellEdit(sheet, cell, source, id) {
  const sample = symbolSamples.find((s) => s.id === id);
  if (!sample) throw refusal('workbook-bad-input', 'Choose one of the three supported Symbols samples.');
  booleanCellCommand({ type: 'booleanCell', action: 'insert', sheet, cell, expectedValue: null });
  if (
    source &&
    (source.formula != null ||
      source.formulaType != null ||
      source.spill != null ||
      source.formulaOutput != null ||
      source.calculation != null ||
      !['text', 'blank', undefined].includes(source.dataType) ||
      (source.value != null && typeof source.value !== 'string'))
  )
    throw refusal(
      'workbook-bad-input',
      'Symbols append to a blank or plain-text cell; use the worksheet to edit other values.'
    );
  const original = source?.value ?? null,
    value = (original ?? '') + sample.value;
  if (value.length > 32767) throw refusal('workbook-limit', 'Symbols would exceed the cell text limit.');
  return { sheet, cell, value, expectedValue: original, expectedFormula: null };
}

export function booleanCellValue(cell) {
  if (
    cell &&
    (['formula', 'formulaType', 'spill', 'formulaOutput', 'calculation'].some((key) => cell[key] != null) ||
      !['blank', 'boolean', undefined, null].includes(cell.dataType) ||
      (cell.value != null && typeof cell.value !== 'boolean'))
  )
    throw refusal(
      'workbook-bad-input',
      'Choose an unmerged blank or literal TRUE/FALSE cell for this checkbox profile.'
    );
  return cell?.value ?? null;
}
