import { rangeBounds } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';
const own = (value, key) => Object.hasOwn(value, key);
const styles = [
  'styleId',
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
  'numberFormat',
];
/** Editing a value or format grants no ownership of an unsupported formula. */
export function preservationPolicy(policy, reason) {
  return {
    ...policy,
    preserveSource: true,
    protectedSource: true,
    canRecalculate: false,
    canCalculateSheet: false,
    canChangeOptions: false,
    reason: reason ?? 'Formula results are stored; recalculate in Microsoft Excel after editing.',
  };
}
export function assertPreservationCommands(commands) {
  if (!Array.isArray(commands) || !commands.length || commands.length > 256)
    throw refusal('formula-unsupported', 'Preserved-calculation editing requires 1–256 commands');
  const footprints = [];
  const area = (sheet, range, content) => {
    const bounds = rangeBounds(range);
    if (bounds.rows * bounds.columns > 10000)
      throw refusal(
        'paste-limit',
        'Preserved-calculation range edits are limited to 10,000 positions; use bounded pasted pages for larger inputs'
      );
    footprints.push({ sheet, range, content });
  };
  const point = (sheet, cell) => {
    if (
      !cell ||
      typeof cell !== 'object' ||
      Array.isArray(cell) ||
      Object.keys(cell).some(
        (key) =>
          !['sheet', 'cell', 'value', 'expectedValue', 'expectedFormula', 'expectedRawValue', ...styles].includes(key)
      ) ||
      (!own(cell, 'value') && !styles.some((key) => own(cell, key)))
    )
      throw refusal(
        'formula-protected',
        'Formula and structural changes remain protected while source calculation is preserved'
      );
    area(sheet, cell.cell, own(cell, 'value'));
  };
  for (const command of commands) {
    if (!command || typeof command !== 'object' || Array.isArray(command))
      throw refusal('internal-invalid-field', 'Invalid preserved-calculation command');
    if (command.type === 'importCells') {
      if (commands.length !== 1) throw refusal('paste-bad-input', 'Staged paste must be one command');
      continue;
    }
    if (command.type === 'setCells') {
      if (
        Object.keys(command).some((key) => !['type', 'sheet', 'cells'].includes(key)) ||
        !Array.isArray(command.cells) ||
        !command.cells.length ||
        command.cells.length > 1000
      )
        throw refusal('workbook-unsupported', 'Preserved setCells requires 1–1000 cells');
      for (const cell of command.cells) point(command.sheet, cell);
    } else if (command.type === 'formatRange') {
      if (Object.keys(command).some((key) => !['type', 'sheet', 'range', 'format'].includes(key)))
        throw refusal('internal-invalid-field', 'Invalid preserved range format');
      area(command.sheet, command.range, false);
    } else if (command.type === 'clearRange' || command.type === 'pasteValues') {
      if (Object.keys(command).some((key) => !['type', 'sheet', 'range', 'source', 'sourceSheet'].includes(key)))
        throw refusal('internal-invalid-field', 'Invalid preserved range edit');
      area(command.sheet, command.range, true);
    } else point(command.sheet, command);
  }
  return footprints;
}
