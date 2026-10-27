/** Portable print policy shared by owner, controls and physical compositor. */
import { rangeBounds, columnName, cellPosition } from './workbook-commands.mjs';
import { validatePrintHeaderFooter } from './print-header-footer.mjs';
import { refusal } from './refusal-codes.mjs';
export const PRINT_LIMITS = Object.freeze({
  sourcePositions: 10000,
  pages: 100,
  renderedPositions: 100000,
  snapshotBytes: 16 * 1024 ** 2,
  layoutBytes: 16 * 1024 ** 2,
  htmlBytes: 16 * 1024 ** 2,
  objects: 32,
});
export function defaultPrintSettings() {
  return {
    area: null,
    paper: 'A4',
    orientation: 'portrait',
    marginsPt: { top: 36, right: 36, bottom: 36, left: 36, header: 18, footer: 18 },
    scaling: { mode: 'scale', percent: 100 },
    repeatRows: null,
    repeatColumns: null,
    gridlines: false,
  };
}
export function printPaperSize(settings) {
  const size = settings.paper === 'A4' ? [(210 / 25.4) * 72, (297 / 25.4) * 72] : [612, 792];
  return settings.orientation === 'landscape' ? size.reverse() : size;
}
function exact(value, keys, label) {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).some((key) => !keys.includes(key))
  )
    throw refusal('internal-invalid-field', 'Invalid ' + label);
  return value;
}
function finite(value, min, max, label) {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max)
    throw refusal('internal-invalid-field', 'Invalid ' + label);
  return value;
}
function integer(value, min, max, label) {
  finite(value, min, max, label);
  if (!Number.isSafeInteger(value)) throw refusal('internal-invalid-field', 'Invalid ' + label);
  return value;
}
export function validatePrintSettings(input, { allowAbsentArea = false } = {}) {
  exact(
    input,
    [
      'area',
      'paper',
      'orientation',
      'marginsPt',
      'scaling',
      'repeatRows',
      'repeatColumns',
      'gridlines',
      'headerFooter',
    ],
    'print settings'
  );
  const settings = { ...defaultPrintSettings(), ...input };
  if (
    !['A4', 'Letter'].includes(settings.paper) ||
    !['portrait', 'landscape'].includes(settings.orientation) ||
    typeof settings.gridlines !== 'boolean'
  )
    throw refusal('print-unsupported', 'Unsupported paper/orientation/gridlines');
  if (settings.headerFooter !== undefined) settings.headerFooter = validatePrintHeaderFooter(settings.headerFooter);
  const margins = exact(settings.marginsPt, ['top', 'right', 'bottom', 'left', 'header', 'footer'], 'print margins');
  settings.marginsPt = {};
  for (const key of ['top', 'right', 'bottom', 'left', 'header', 'footer'])
    settings.marginsPt[key] = finite(margins[key], 0, 144, key + ' margin');
  const paper = printPaperSize(settings);
  if (margins.left + margins.right >= paper[0] - 18 || margins.top + margins.bottom >= paper[1] - 18)
    throw refusal('print-bad-input', 'Print margins leave no printable body');
  const scaling = exact(settings.scaling, ['mode', 'percent', 'width', 'height'], 'print scaling');
  if (scaling.mode === 'scale') {
    exact(scaling, ['mode', 'percent'], 'percent scaling');
    settings.scaling = { mode: 'scale', percent: integer(scaling.percent, 10, 400, 'print scale') };
  } else if (scaling.mode === 'fit') {
    exact(scaling, ['mode', 'width', 'height'], 'fit scaling');
    settings.scaling = {
      mode: 'fit',
      width: integer(scaling.width, 1, 10, 'fit width'),
      height: integer(scaling.height, 0, 10, 'fit height'),
    };
  } else throw refusal('print-unsupported', 'Unsupported print scaling');
  if (settings.area === null && allowAbsentArea) {
    if (settings.repeatRows !== null || settings.repeatColumns !== null)
      throw refusal('print-bad-input', 'Print titles require an area');
    return settings;
  }
  if (typeof settings.area !== 'string') throw refusal('print-bad-input', 'Choose one explicit print area');
  const bounds = rangeBounds(settings.area);
  if (bounds.size > PRINT_LIMITS.sourcePositions)
    throw refusal('print-limit', 'This print slice supports at most 10,000 source positions');
  settings.area =
    columnName(bounds.first.col) + (bounds.first.row + 1) + ':' + columnName(bounds.last.col) + (bounds.last.row + 1);
  if (settings.repeatRows !== null) {
    const repeat = exact(settings.repeatRows, ['first', 'last'], 'repeated rows');
    integer(repeat.first, 1, 1048576, 'repeat first row');
    integer(repeat.last, repeat.first, Math.min(1048576, repeat.first + 99), 'repeat last row');
    if (repeat.first !== bounds.first.row + 1 || repeat.last >= bounds.last.row + 1)
      throw refusal('print-bad-input', 'Repeat rows must be a contiguous prefix with a remaining body row');
    settings.repeatRows = { first: repeat.first, last: repeat.last };
  }
  if (settings.repeatColumns !== null) {
    const repeat = exact(settings.repeatColumns, ['first', 'last'], 'repeated columns');
    if (
      typeof repeat.first !== 'string' ||
      typeof repeat.last !== 'string' ||
      !/^([A-Z]{1,3})$/.test(repeat.first) ||
      !/^([A-Z]{1,3})$/.test(repeat.last)
    )
      throw refusal('internal-invalid-field', 'Invalid repeated column letters');
    const first = cellPosition(repeat.first + '1').col,
      last = cellPosition(repeat.last + '1').col;
    if (first !== bounds.first.col || last < first || last >= bounds.last.col || last - first > 99)
      throw refusal('print-bad-input', 'Repeat columns must be a contiguous prefix with a remaining body column');
    settings.repeatColumns = { first: columnName(first), last: columnName(last) };
  }
  return settings;
}
