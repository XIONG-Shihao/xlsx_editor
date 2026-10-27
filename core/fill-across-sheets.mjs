/** One native-owned Across Worksheets action. Ordinary commands separate
 * actions during compact replay; their valid limits do not accumulate. */
import { rangeBounds, columnName } from './workbook-commands.mjs';
import { refusal } from './refusal-codes.mjs';
export const FILL_ACROSS_SHEETS_LIMITS = Object.freeze({ destinationPositions: 10000, actionBytes: 2 * 1024 * 1024 });
const fields = ['type', 'sheet', 'sourceSheet', 'range', 'mode'];
const modes = ['all', 'contents', 'formats'];
const plain = (value) =>
  !!value &&
  typeof value === 'object' &&
  !Array.isArray(value) &&
  (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null);
const validSheet = (value) =>
  typeof value === 'string' &&
  value.length > 0 &&
  value.length <= 31 &&
  !!value.trim() &&
  !/[\\/*?:\[\]\u0000-\u001f\uD800-\uDFFF\uFFFE\uFFFF]/u.test(value) &&
  !value.startsWith("'") &&
  !value.endsWith("'");
export function fillAcrossSheetsCommand(value) {
  if (
    !plain(value) ||
    Object.keys(value).length !== fields.length ||
    fields.some((key) => !Object.hasOwn(value, key)) ||
    value.type !== 'fillAcrossSheets' ||
    !validSheet(value.sheet) ||
    !validSheet(value.sourceSheet) ||
    value.sheet === value.sourceSheet ||
    !modes.includes(value.mode) ||
    typeof value.range !== 'string' ||
    !/^([A-Z]{1,3}[1-9]\d{0,6})(?::([A-Z]{1,3}[1-9]\d{0,6}))?$/.test(value.range)
  )
    throw refusal(
      'internal-invalid-field',
      'Invalid Across Worksheets command: use distinct worksheet names, a canonical rectangle and All, Contents or Formats.'
    );
  const bounds = rangeBounds(value.range),
    first = columnName(bounds.first.col) + (bounds.first.row + 1),
    last = columnName(bounds.last.col) + (bounds.last.row + 1),
    canonical = first === last ? first : first + ':' + last;
  if (value.range !== canonical || bounds.size > FILL_ACROSS_SHEETS_LIMITS.destinationPositions)
    throw refusal('fill-limit', 'Across Worksheets requires a canonical rectangle of at most 10,000 positions.');
  return {
    type: 'fillAcrossSheets',
    sheet: value.sheet,
    sourceSheet: value.sourceSheet,
    range: value.range,
    mode: value.mode,
  };
}
export function normalizeFillAcrossSheetsRuns(commands, view) {
  if (!Array.isArray(commands) || !commands.length)
    throw refusal('internal-invalid-field', 'Invalid workbook command batch');
  const normalized = [],
    runs = [];
  let run = null;
  for (const command of commands) {
    if (command?.type !== 'fillAcrossSheets') {
      normalized.push(command);
      run = null;
      continue;
    }
    const accepted = fillAcrossSheetsCommand(command);
    if (view) {
      for (const name of [accepted.sourceSheet, accepted.sheet]) {
        const sheets = view.sheets.filter((sheet) => sheet.name === name);
        if (sheets.length !== 1 || !sheets[0].path || (sheets[0].state ?? 'visible') !== 'visible')
          throw refusal(
            'fill-bad-input',
            'Across Worksheets requires distinct visible source and destination worksheets.'
          );
      }
    }
    if (!run) {
      run = {
        commands: [],
        sourceSheet: accepted.sourceSheet,
        range: accepted.range,
        mode: accepted.mode,
        targets: new Set(),
        positions: 0,
      };
      runs.push(run);
    }
    if (
      run.sourceSheet !== accepted.sourceSheet ||
      run.range !== accepted.range ||
      run.mode !== accepted.mode ||
      run.targets.has(accepted.sheet)
    )
      throw refusal(
        'fill-bad-input',
        'One Across Worksheets action requires a common source, rectangle and policy with distinct destinations.'
      );
    run.targets.add(accepted.sheet);
    run.positions += rangeBounds(accepted.range).size;
    if (run.positions > FILL_ACROSS_SHEETS_LIMITS.destinationPositions)
      throw refusal('fill-limit', 'Across Worksheets supports at most 10,000 total destination positions per action.');
    run.commands.push(accepted);
    normalized.push(accepted);
  }
  for (const item of runs)
    if (
      new TextEncoder().encode(JSON.stringify({ operation: 'apply', commands: item.commands })).length >
      FILL_ACROSS_SHEETS_LIMITS.actionBytes
    )
      throw refusal('fill-limit', 'Across Worksheets action exceeds its 2 MiB bound.');
  return normalized;
}
