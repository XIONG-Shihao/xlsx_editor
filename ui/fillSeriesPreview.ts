import type { WorkbookCell, WorkbookCommand, WorkbookSheet, WorkbookView } from '../core/contracts';
import { formatDisplayValue } from '../core/display.mjs';
import { hasExactNumericValue } from '../core/calculation.mjs';
import { addressPosition, colName } from './sheetModel';
import { captureHomeBinding, type HomeBinding } from './homeCommands';

export const FILL_SERIES_LIMITS = Object.freeze({ positions: 10_000, page: 100 });
export const FILL_SERIES_POLICY = 'linear-two-seeds-15-significant-digits' as const;
export type FillSeriesKind = 'number' | 'date';
export interface FillSeriesSeed {
  readonly cell: string;
  readonly value: number;
  readonly rawValue: string | null;
}
export interface FillSeriesDestination {
  readonly cell: string;
  readonly value: number;
}
export interface FillSeriesPreview {
  readonly schema: 1;
  readonly policy: typeof FILL_SERIES_POLICY;
  readonly binding: HomeBinding;
  readonly sheet: string;
  readonly range: string;
  readonly orientation: 'vertical' | 'horizontal';
  readonly seeds: readonly [FillSeriesSeed, FillSeriesSeed];
  readonly step: number;
  readonly positions: number;
  readonly destinationCount: number;
  readonly destinations: readonly FillSeriesDestination[];
  readonly date?: Readonly<{ date1904: boolean }>;
  readonly command: Extract<WorkbookCommand, { type: 'clearRange' | 'fillSeries' }> & { readonly type: 'fillSeries' };
}
export type FillSeriesRefusalCode =
  | 'invalid-input'
  | 'retired-target'
  | 'unsupported-selection'
  | 'read-incomplete'
  | 'protected-target'
  | 'invalid-seeds'
  | 'occupied-target'
  | 'number-overflow';
export type FillSeriesPreviewResult =
  | { readonly status: 'ready'; readonly preview: FillSeriesPreview }
  | { readonly status: 'refused'; readonly code: FillSeriesRefusalCode; readonly reason: string };
export interface FillSeriesPreviewInput {
  binding: HomeBinding;
  sheet: WorkbookSheet;
  cells: Record<string, WorkbookCell>;
  readRange: string;
  tables?: WorkbookView['tables'];
  protectionReason?: string | null;
  kind?: FillSeriesKind;
}
export interface FillSeriesPreviewPage {
  readonly offset: number;
  readonly total: number;
  readonly rows: readonly FillSeriesDestination[];
  readonly nextOffset: number | null;
}

class Refusal extends Error {
  constructor(
    readonly code: FillSeriesRefusalCode,
    reason: string
  ) {
    super(reason);
  }
}
const refuse = (code: FillSeriesRefusalCode, reason: string): never => {
  throw new Refusal(code, reason);
};
function rectangle(range: string) {
  if (typeof range !== 'string' || !/^[A-Z]{1,3}[1-9]\d{0,6}(?::[A-Z]{1,3}[1-9]\d{0,6})?$/.test(range))
    return refuse('unsupported-selection', 'Select a canonical worksheet range.');
  const [first, last = first] = range.split(':'),
    a = addressPosition(first),
    b = addressPosition(last);
  if (a.col > 16383 || b.col > 16383 || a.row > 1048575 || b.row > 1048575 || a.col > b.col || a.row > b.row)
    return refuse('unsupported-selection', 'The selected range exceeds worksheet bounds.');
  return { first: a, last: b, rows: b.row - a.row + 1, columns: b.col - a.col + 1 };
}
const intersects = (a: ReturnType<typeof rectangle>, b: ReturnType<typeof rectangle>) =>
  a.first.row <= b.last.row && a.last.row >= b.first.row && a.first.col <= b.last.col && a.last.col >= b.first.col;
function plain(value: unknown): value is Record<string, unknown> {
  return (
    !!value &&
    typeof value === 'object' &&
    !Array.isArray(value) &&
    (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null)
  );
}
function protectedCell(cell: WorkbookCell | undefined): boolean {
  return (
    !!cell &&
    (cell.formula !== undefined ||
      cell.formulaType !== undefined ||
      cell.formulaOutput !== undefined ||
      cell.spill !== undefined ||
      cell.scenarioResult !== undefined ||
      cell.calculation !== undefined ||
      cell.cachedResult !== undefined ||
      cell.error !== undefined ||
      cell.dataType === 'error')
  );
}

/** Pure UI preview of the existing native numeric profile, never an admission
 * certificate. The captured host must provide a complete accepted read of
 * readRange; an absent record then means an empty destination. Native Apply
 * still checks actual source groups, spill ownership and accepted version. */
export function createFillSeriesPreview(input: FillSeriesPreviewInput): FillSeriesPreviewResult {
  try {
    if (
      !plain(input) ||
      Object.keys(input).some(
        (key) => !['binding', 'sheet', 'cells', 'readRange', 'tables', 'protectionReason', 'kind'].includes(key)
      ) ||
      !input.sheet ||
      !plain(input.cells)
    )
      return refuse('invalid-input', 'The fill preview input is unavailable.');
    const binding = captureHomeBinding(input.binding),
      sheet = input.sheet,
      kind = input.kind ?? 'number';
    if (kind !== 'number' && kind !== 'date') return refuse('invalid-input', 'Choose Numbers or Dates for the series.');
    if (kind === 'date' && sheet.date1904 !== undefined && typeof sheet.date1904 !== 'boolean')
      return refuse('invalid-input', 'The workbook date system is unavailable.');
    if (!binding.ownerIdentity || binding.sheet !== sheet.name)
      return refuse('retired-target', 'The original worksheet owner is unavailable.');
    if (binding.specialSelection || binding.mode !== 'normal')
      return refuse('unsupported-selection', 'Fill Series needs one regular range in Normal view.');
    if (input.readRange !== binding.selection)
      return refuse('read-incomplete', 'Read the complete original selected range before previewing.');
    const bounds = rectangle(binding.selection),
      positions = bounds.rows * bounds.columns;
    const vertical = bounds.columns === 1 && bounds.rows >= 3,
      horizontal = bounds.rows === 1 && bounds.columns >= 3;
    if ((!vertical && !horizontal) || positions > FILL_SERIES_LIMITS.positions)
      return refuse(
        'unsupported-selection',
        'Select one row or column with two numeric seeds and a destination, up to 10,000 positions.'
      );
    if (input.protectionReason !== undefined && input.protectionReason !== null) {
      if (typeof input.protectionReason !== 'string' || !input.protectionReason)
        return refuse('invalid-input', 'The protection result is unavailable.');
      return refuse('protected-target', input.protectionReason);
    }
    if (!Array.isArray(sheet.merges) || sheet.merges.some((range) => intersects(bounds, rectangle(range))))
      return refuse('protected-target', 'The fill range intersects merged cells or has unavailable merge metadata.');
    if (sheet.filter && intersects(bounds, rectangle(sheet.filter.range)))
      return refuse('protected-target', 'Fill Series in a filtered range is not qualified.');
    if (
      (sheet.tableRanges !== undefined && !Array.isArray(sheet.tableRanges)) ||
      (input.tables !== undefined && !Array.isArray(input.tables))
    )
      return refuse('invalid-input', 'The table protection inventory is unavailable.');
    for (const table of sheet.tableRanges ?? []) {
      if (!table || typeof table.ref !== 'string')
        return refuse('invalid-input', 'The table protection inventory is unavailable.');
      if (intersects(bounds, rectangle(table.ref)))
        return refuse('protected-target', 'The fill range intersects a table.');
    }
    for (const table of input.tables ?? []) {
      if (!table || typeof table.sheet !== 'string' || typeof table.ref !== 'string')
        return refuse('invalid-input', 'The table protection inventory is unavailable.');
      if (table.sheet === sheet.name && intersects(bounds, rectangle(table.ref)))
        return refuse('protected-target', 'The fill range intersects a table.');
    }
    for (const [reference, cell] of Object.entries(input.cells)) {
      const at = rectangle(reference);
      if (
        reference.includes(':') ||
        !intersects(bounds, at) ||
        !plain(cell) ||
        !Object.hasOwn(cell, 'value') ||
        (cell.value !== null && !['string', 'number', 'boolean'].includes(typeof cell.value))
      )
        return refuse('read-incomplete', 'The selected read contains a foreign or malformed cell.');
    }
    const at = (offset: number) =>
      colName(bounds.first.col + (horizontal ? offset : 0)) + (bounds.first.row + (vertical ? offset : 0) + 1);
    const seeds = [0, 1].map((offset) => {
      const cell = at(offset),
        source = input.cells[cell];
      if (
        !source ||
        protectedCell(source) ||
        (source.cachedValue !== undefined && source.cachedValue !== null) ||
        typeof source.value !== 'number' ||
        !Number.isFinite(source.value) ||
        (source.rawValue !== undefined && source.rawValue !== null && typeof source.rawValue !== 'string') ||
        !hasExactNumericValue(source.value, source.rawValue)
      )
        return refuse(
          'invalid-seeds',
          'Fill Series needs two exact literal numeric seeds, without formulas or calculated/stored results.'
        );
      return Object.freeze({ cell, value: source.value, rawValue: source.rawValue ?? null });
    }) as [FillSeriesSeed, FillSeriesSeed];
    const calendar = !!sheet.date1904;
    if (kind === 'date' && seeds.some((seed) => !validDateSerial(seed.value, calendar)))
      return refuse(
        'invalid-seeds',
        'Date Series needs two whole-day numeric Excel dates, within the workbook date system. Date text and times are not supported.'
      );
    const step = seeds[1].value - seeds[0].value;
    if (!Number.isFinite(step))
      return refuse('number-overflow', 'The numeric seed difference exceeds workbook limits.');
    const destinations: FillSeriesDestination[] = [];
    for (let offset = 2; offset < positions; offset++) {
      const cell = at(offset),
        previous = input.cells[cell];
      if (protectedCell(previous) || (previous?.cachedValue !== undefined && previous.cachedValue !== null))
        return refuse(
          'protected-target',
          `The destination ${cell} contains a formula, stored result or protected output.`
        );
      if (previous?.value !== undefined && previous.value !== null)
        return refuse('occupied-target', `Fill Series would overwrite ${cell}.`);
      const raw = seeds[0].value + step * offset,
        value = Number(raw.toPrecision(15));
      if (!Number.isFinite(raw) || !Number.isFinite(value))
        return refuse('number-overflow', 'Fill Series produced a number outside workbook limits.');
      if (kind === 'date' && !validDateSerial(value, calendar))
        return refuse('number-overflow', 'Date Series would leave the supported date range.');
      // Native series_precision rounds exact binary64 to 15 significant digits
      // and returns positive zero. Keep both original seeds untouched.
      destinations.push(Object.freeze({ cell, value: value === 0 ? 0 : value }));
    }
    const command = Object.freeze({ type: 'fillSeries' as const, sheet: binding.sheet, range: binding.selection });
    return Object.freeze({
      status: 'ready',
      preview: Object.freeze({
        schema: 1,
        policy: FILL_SERIES_POLICY,
        binding,
        sheet: binding.sheet,
        range: binding.selection,
        orientation: vertical ? 'vertical' : 'horizontal',
        ...(kind === 'date' ? { date: Object.freeze({ date1904: calendar }) } : {}),
        seeds: Object.freeze(seeds),
        step,
        positions,
        destinationCount: destinations.length,
        destinations: Object.freeze(destinations),
        command,
      }),
    });
  } catch (error) {
    return Object.freeze({
      status: 'refused',
      code: error instanceof Refusal ? error.code : 'invalid-input',
      reason: error instanceof Refusal ? error.message : 'The original fill metadata is invalid or unavailable.',
    });
  }
}

// Whole-day serials through 9999-12-31; the existing display contract retains
// Excel's 1900 serial-60 compatibility date. No local-time arithmetic is used.
function validDateSerial(value: number, date1904: boolean): boolean {
  return Number.isInteger(value) && value >= (date1904 ? 0 : 1) && value <= (date1904 ? 2_957_003 : 2_958_465);
}
export function fillSeriesPreviewValue(preview: FillSeriesPreview, value: number): string {
  return preview.date
    ? formatDisplayValue(value, 14, preview.date.date1904)
    : Object.is(value, -0)
      ? '-0'
      : String(value);
}
/** Fill content and date-format only the generated tail in one owner batch.
 * Seed values/styles and every other destination style property stay intact. */
export function fillSeriesCommands(preview: FillSeriesPreview): WorkbookCommand[] {
  fillSeriesPreviewPage(preview);
  if (!preview.date) return [preview.command];
  const first = preview.destinations[0].cell,
    last = preview.destinations.at(-1)!.cell;
  return [
    preview.command,
    {
      type: 'formatRange',
      sheet: preview.sheet,
      range: first === last ? first : first + ':' + last,
      format: { numberFormat: 'date' },
    },
  ];
}

export function fillSeriesPreviewPage(preview: FillSeriesPreview, offset = 0, limit = 100): FillSeriesPreviewPage {
  if (
    !preview ||
    preview.schema !== 1 ||
    preview.policy !== FILL_SERIES_POLICY ||
    !Array.isArray(preview.destinations) ||
    preview.destinationCount !== preview.destinations.length ||
    preview.positions !== preview.destinationCount + 2 ||
    preview.positions > FILL_SERIES_LIMITS.positions ||
    !Number.isInteger(offset) ||
    offset < 0 ||
    offset > preview.destinationCount ||
    !Number.isInteger(limit) ||
    limit < 1 ||
    limit > FILL_SERIES_LIMITS.page
  )
    throw Error('Invalid Fill Series preview page.');
  if (
    preview.date &&
    (!plain(preview.date) ||
      Object.keys(preview.date).join(',') !== 'date1904' ||
      typeof preview.date.date1904 !== 'boolean' ||
      preview.seeds.some((seed) => !validDateSerial(seed.value, preview.date!.date1904)) ||
      preview.destinations.some((row) => !validDateSerial(row.value, preview.date!.date1904)))
  )
    throw Error('Invalid Date Series preview.');
  const end = Math.min(offset + limit, preview.destinationCount);
  return Object.freeze({
    offset,
    total: preview.destinationCount,
    rows: Object.freeze(preview.destinations.slice(offset, end)),
    nextOffset: end < preview.destinationCount ? end : null,
  });
}
