/** Explicit typed field conversion; no formula parsing or structural edit. */
import type { WorkbookCell, WorkbookCommand } from '../core/contracts';
import { convertTextField, type TextConversionPolicy, type TextConversionResult } from '../core/text-conversion.mjs';
import { hasExactNumericValue } from '../core/calculation.mjs';
export const TEXT_COLUMN_LIMITS = Object.freeze({
  rows: 1000,
  positions: 10000,
  readPositions: 1000,
  bytes: 1048576,
  textLength: 32767,
});
export interface TextColumnBinding {
  documentId: string;
  grantDocumentId: string;
  ownerIdentity: string;
  ownerKey: string;
  state: string;
  sheet: string;
  selection: string;
  navigationGeneration: number;
  activationEpoch: number;
}
export type TextColumnOptions = { destination: string; policies?: readonly TextConversionPolicy[] } & (
  | { mode?: 'delimited'; delimiter: string; quotes: 'none' | 'double' }
  | { mode: 'fixed-width'; breaks: readonly number[] }
);
export interface TextColumnContext {
  binding: TextColumnBinding;
  /** Captured accepted worksheet calendar; undefined is explicitly unknown. */
  date1904?: boolean;
  isCurrent: (binding: TextColumnBinding) => boolean;
  isRuntimeCurrent?: (binding: TextColumnBinding) => boolean;
  availability: () => {
    active: boolean;
    readOnly: boolean;
    busy: boolean;
    literal: boolean;
    specialSelection: boolean;
  };
  /** Captured reader/inspection/strict callback, never a mutable latest owner. */
  read: (range: string) => Promise<Record<string, WorkbookCell>>;
  inspect: (source: string, destination: string) => Promise<{ reason: string | null }>;
  accept?: (commands: WorkbookCommand[]) => Promise<void>;
}
export interface TextColumnRow {
  source: string;
  text: string;
  fields: readonly string[];
  padded?: readonly boolean[];
}
export interface TextColumnOutput {
  reference: string;
  value: string | number | boolean;
  padded: boolean;
  original: string;
  conversion: TextConversionResult;
  previous?: WorkbookCell;
  occupied: boolean;
}
export interface TextColumnPreview {
  binding: TextColumnBinding;
  options: TextColumnOptions;
  source: string;
  destination: string;
  rows: readonly TextColumnRow[];
  outputs: readonly TextColumnOutput[];
  columns: number;
  collisions: number;
  rejections: number;
  notices: number;
  date1904: boolean | undefined;
  fingerprint: string;
  commands: WorkbookCommand[];
}
export class TextColumnsRetired extends Error {
  readonly code = 'TEXT_COLUMNS_RETIRED';
  constructor() {
    super('The workbook or selection changed. Reopen Text to Columns.');
  }
}
export class TextColumnsCancelled extends Error {
  readonly code = 'TEXT_COLUMNS_CANCELLED';
  constructor() {
    super('Text to Columns cancelled before the workbook command.');
  }
}
const bytes = (value: string) => new TextEncoder().encode(value).length;
function xmlText(value: string): void {
  if (value.length > 32767 || /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/u.test(value))
    throw Error('Text contains an unsupported character or exceeds cell text limits.');
}
export function textColumnContextKey(context: TextColumnContext): string {
  return JSON.stringify([textColumnBindingKey(context.binding), context.date1904 ?? null]);
}
export function textColumnBindingKey(binding: TextColumnBinding): string {
  return JSON.stringify([
    binding.documentId,
    binding.grantDocumentId,
    binding.ownerIdentity,
    binding.ownerKey,
    binding.state,
    binding.sheet,
    binding.selection,
    binding.navigationGeneration,
    binding.activationEpoch,
  ]);
}
function position(ref: string): { row: number; col: number } {
  const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(ref);
  if (!match) throw Error('Enter a canonical cell such as C2.');
  const col = [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1,
    row = Number(match[2]) - 1;
  if (col >= 16384 || row >= 1048576) throw Error('The destination exceeds worksheet limits.');
  return { row, col };
}
function column(index: number): string {
  let result = '';
  for (let n = index + 1; n; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + ((n - 1) % 26)) + result;
  return result;
}
function ref(row: number, col: number): string {
  return column(col) + (row + 1);
}
function bounds(range: string) {
  const parts = range.split(':');
  if (parts.length > 2) throw Error('Select one column of text.');
  const first = position(parts[0]),
    last = position(parts[1] ?? parts[0]);
  if (last.row < first.row || last.col < first.col) throw Error('Select one column of text.');
  return { first, last, rows: last.row - first.row + 1, columns: last.col - first.col + 1 };
}
function checkedBreaks(breaks: readonly number[]): readonly number[] {
  if (
    !Array.isArray(breaks) ||
    breaks.length < 1 ||
    breaks.length >= TEXT_COLUMN_LIMITS.positions ||
    Array.from(breaks).some(
      (value, index) =>
        !Number.isInteger(value) ||
        value < 1 ||
        value > TEXT_COLUMN_LIMITS.textLength ||
        (index > 0 && value <= breaks[index - 1])
    )
  )
    throw Error('Enter increasing, unique split positions from 1 to 32,767.');
  return Object.freeze([...breaks]);
}
function checkedOptions(value: TextColumnOptions): TextColumnOptions {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw Error('Choose Delimited or Fixed width splitting.');
  if (
    value.policies !== undefined &&
    (!Array.isArray(value.policies) ||
      value.policies.length > 10000 ||
      Array.from(value.policies).some((policy) => !['text', 'number', 'boolean', 'iso-date'].includes(policy)))
  )
    throw Error('Choose a supported type for each output field.');
  position(value.destination);
  const policies = value.policies ? { policies: Object.freeze([...value.policies]) } : {};
  if (value.mode === 'fixed-width') {
    if (Object.keys(value).some((key) => !['mode', 'breaks', 'destination', 'policies'].includes(key)))
      throw Error('Fixed width uses split positions, without delimiters or quote parsing.');
    return Object.freeze({ ...value, breaks: checkedBreaks(value.breaks), ...policies });
  }
  if (
    (value.mode !== undefined && value.mode !== 'delimited') ||
    Object.keys(value).some((key) => !['mode', 'delimiter', 'quotes', 'destination', 'policies'].includes(key)) ||
    typeof value.delimiter !== 'string' ||
    [...value.delimiter].length !== 1 ||
    /[\r\n"]/u.test(value.delimiter) ||
    !['none', 'double'].includes(value.quotes)
  )
    throw Error('Choose one delimiter and a supported quote policy.');
  xmlText(value.delimiter);
  return Object.freeze({ ...value, ...policies });
}
/** Positions count Unicode code points. Combining marks are separate characters;
 * source whitespace/quotes/newlines are literal, not formatting instructions. */
export function splitFixedWidthRecord(
  text: string,
  breaks: readonly number[]
): { fields: readonly string[]; padded: readonly boolean[] } {
  const positions = checkedBreaks(breaks);
  xmlText(text);
  const chars = [...text],
    fields: string[] = [],
    padded: boolean[] = [];
  let start = 0;
  for (const end of [...positions, chars.length]) {
    fields.push(chars.slice(start, end).join(''));
    padded.push(start > chars.length);
    start = end;
  }
  return Object.freeze({ fields: Object.freeze(fields), padded: Object.freeze(padded) });
}
export function splitTextRecord(text: string, delimiter: string, quotes: 'none' | 'double'): readonly string[] {
  checkedOptions({ delimiter, quotes, destination: 'A1' });
  xmlText(text);
  if (quotes === 'none') return Object.freeze(text.split(delimiter));
  const fields: string[] = [];
  let value = '',
    quoted = false,
    closed = false,
    started = false;
  const chars = [...text];
  for (let i = 0; i < chars.length; i++) {
    const char = chars[i];
    if (quoted) {
      if (char === '"') {
        if (chars[i + 1] === '"') {
          value += '"';
          i++;
        } else {
          quoted = false;
          closed = true;
        }
      } else value += char;
    } else if (char === delimiter) {
      fields.push(value);
      value = '';
      closed = false;
      started = false;
    } else if (char === '"' && !started && !closed) {
      quoted = true;
      started = true;
    } else {
      if (closed || char === '"' || char === '\n' || char === '\r')
        throw Error('Use one correctly quoted record per cell, or choose No quotes.');
      value += char;
      started = true;
    }
  }
  if (quoted) throw Error('A quoted field is unfinished.');
  fields.push(value);
  fields.forEach(xmlText);
  return Object.freeze(fields);
}
function check(context: TextColumnContext, signal?: AbortSignal, editable = false): void {
  if (signal?.aborted) throw new TextColumnsCancelled();
  if (!context.isCurrent(context.binding) || !context.availability().active) throw new TextColumnsRetired();
  const state = context.availability();
  if (state.literal) throw Error('This command requires an editable XLSX workbook.');
  if (state.specialSelection) throw Error('Click a regular cell to leave the special selection.');
  if (state.busy) throw Error('Wait for the workbook operation to finish.');
  if (editable && state.readOnly) throw Error('This workbook is read-only.');
  if (editable && !context.accept) throw Error('Workbook commands are unavailable in this editor.');
}
function protect(cell: WorkbookCell | undefined): void {
  if (!cell) return;
  if (
    cell.formula !== undefined ||
    cell.formulaType ||
    cell.spill ||
    cell.scenarioResult ||
    cell.calculation ||
    cell.dataType === 'error' ||
    cell.error
  )
    throw Error('A source or destination contains a protected formula, array or error.');
  if (
    typeof cell.value === 'number' &&
    (!Number.isFinite(cell.value) ||
      (Number.isInteger(cell.value) && !Number.isSafeInteger(cell.value)) ||
      !hasExactNumericValue(cell.value, cell.rawValue))
  )
    throw Error('A destination number has protected stored precision.');
  if (cell.value !== undefined && cell.value !== null && !['string', 'number', 'boolean'].includes(typeof cell.value))
    throw Error('A source or destination is not an ordinary scalar cell.');
  const type =
    cell.value === undefined || cell.value === null
      ? 'blank'
      : typeof cell.value === 'string'
        ? 'text'
        : typeof cell.value === 'number'
          ? 'number'
          : 'boolean';
  if (cell.dataType && cell.dataType !== type) throw Error('A source or destination has an unqualified value type.');
}
async function read(
  context: TextColumnContext,
  range: string,
  signal?: AbortSignal
): Promise<Record<string, WorkbookCell>> {
  const area = bounds(range),
    cells: Record<string, WorkbookCell> = {};
  for (let left = area.first.col; left <= area.last.col; left += 1000) {
    const width = Math.min(1000, area.last.col - left + 1),
      step = Math.floor(1000 / width);
    for (let row = area.first.row; row <= area.last.row; row += step) {
      check(context, signal);
      const last = Math.min(area.last.row, row + step - 1),
        window = ref(row, left) + ':' + ref(last, left + width - 1),
        reply = await context.read(window);
      check(context, signal);
      if (!reply || typeof reply !== 'object' || Array.isArray(reply))
        throw Error('The cell preview returned an invalid rectangle.');
      for (const [address, cell] of Object.entries(reply)) {
        const p = position(address);
        if (
          p.row < row ||
          p.row > last ||
          p.col < left ||
          p.col >= left + width ||
          !cell ||
          typeof cell !== 'object' ||
          Array.isArray(cell)
        )
          throw Error('The cell preview returned a foreign or invalid cell.');
        cells[address] = structuredClone(cell);
      }
      if (bytes(JSON.stringify(cells)) > TEXT_COLUMN_LIMITS.bytes)
        throw Error('The cell preview exceeds the supported size.');
    }
  }
  return cells;
}
const cellFingerprint = (cell: WorkbookCell | undefined) => [
  cell?.value === undefined
    ? null
    : typeof cell.value === 'number' && Object.is(cell.value, -0)
      ? ['number', '-0']
      : cell.value,
  cell?.formula ?? null,
  cell?.formulaType ?? null,
  cell?.spill ?? null,
  cell?.scenarioResult ?? null,
  cell?.dataType ?? null,
  cell?.rawValue ?? null,
  cell?.error ?? null,
];
export async function previewTextColumns(
  context: TextColumnContext,
  input: TextColumnOptions,
  { signal }: { signal?: AbortSignal } = {}
): Promise<TextColumnPreview> {
  check(context, signal);
  const options = checkedOptions(input),
    source = bounds(context.binding.selection);
  if (source.columns !== 1 || source.rows > 1000) throw Error('Select one text column with at most 1,000 rows.');
  const original = await read(context, context.binding.selection, signal),
    rows: TextColumnRow[] = [];
  for (let row = source.first.row; row <= source.last.row; row++) {
    const address = ref(row, source.first.col),
      cell = original[address];
    protect(cell);
    if (typeof cell?.value !== 'string' || (cell.dataType && cell.dataType !== 'text'))
      throw Error('Every selected source cell must contain literal text.');
    rows.push({
      source: address,
      text: cell.value,
      ...(options.mode === 'fixed-width'
        ? splitFixedWidthRecord(cell.value, options.breaks)
        : { fields: splitTextRecord(cell.value, options.delimiter, options.quotes) }),
    });
  }
  const width = Math.max(...rows.map((row) => row.fields.length)),
    start = position(options.destination);
  if (rows.length * width > 10000 || start.row + rows.length > 1048576 || start.col + width > 16384)
    throw Error('The split exceeds 10,000 output cells or worksheet limits.');
  if (options.policies && options.policies.length > width)
    throw Error('The output fields changed. Choose their types again.');
  const destination = ref(start.row, start.col) + ':' + ref(start.row + rows.length - 1, start.col + width - 1);
  const admission = await context.inspect(context.binding.selection, destination);
  check(context, signal);
  if (!admission || !(admission.reason === null || typeof admission.reason === 'string'))
    throw Error('The workbook protection inventory is unavailable.');
  if (admission.reason) throw Error(admission.reason);
  const prior = await read(context, destination, signal),
    outputs: TextColumnOutput[] = [],
    cells: Array<{
      cell: string;
      value: string | number | boolean;
      numberFormat?: 'date';
      expectedValue: string | number | boolean | null;
      expectedFormula: null;
      expectedRawValue: string | null;
    }> = [];
  rows.forEach((row, index) => {
    for (let col = 0; col < width; col++) {
      const reference = ref(start.row + index, start.col + col),
        previous = prior[reference];
      protect(previous);
      const original = row.fields[col] ?? '',
        padded = row.padded?.[col] ?? col >= row.fields.length,
        conversion = convertTextField(original, options.policies?.[col] ?? 'text', {
          date1904: context.date1904,
          padded,
        }),
        value = conversion.status === 'converted' ? conversion.value : original,
        occupied = previous?.value !== undefined && previous.value !== null;
      outputs.push({ reference, value, padded, original, conversion, previous, occupied });
      if (conversion.status === 'converted')
        cells.push({
          cell: reference,
          value,
          ...(conversion.kind === 'date' ? { numberFormat: conversion.numberFormat } : {}),
          expectedValue: previous?.value ?? null,
          expectedFormula: null,
          expectedRawValue: previous?.rawValue ?? null,
        });
    }
  });
  const rejections = outputs.filter((output) => output.conversion.status === 'rejected').length,
    notices = outputs.reduce((count, output) => count + output.conversion.notices.length, 0);
  const commands: WorkbookCommand[] = rejections ? [] : [{ type: 'setCells', sheet: context.binding.sheet, cells }];
  if (bytes(JSON.stringify(commands)) > 1048576) throw Error('The guarded split command exceeds the supported size.');
  const fingerprint = JSON.stringify([
    textColumnContextKey(context),
    options,
    rows.map((row) => [row.source, row.text]),
    outputs.map((output) => [
      output.reference,
      output.value,
      output.padded,
      output.conversion,
      cellFingerprint(output.previous),
    ]),
  ]);
  check(context, signal);
  return Object.freeze({
    binding: context.binding,
    options,
    source: context.binding.selection,
    destination,
    rows: Object.freeze(rows),
    outputs: Object.freeze(outputs),
    columns: width,
    rejections,
    notices,
    date1904: context.date1904,
    collisions: outputs.filter((output) => output.occupied).length,
    fingerprint,
    commands,
  });
}
export async function applyTextColumns(
  context: TextColumnContext,
  preview: TextColumnPreview,
  {
    signal,
    replaceFingerprint,
    beforeAcceptance,
  }: { signal?: AbortSignal; replaceFingerprint?: string; beforeAcceptance?: () => void } = {}
): Promise<{
  outcome: 'accepted';
  source: string;
  destination: string;
  positions: number;
  current: boolean;
  runtimeCurrent: boolean;
}> {
  check(context, signal, true);
  if (preview.rejections) throw Error('Resolve the rejected conversions before splitting.');
  if (preview.date1904 !== context.date1904) throw new TextColumnsRetired();
  if (textColumnBindingKey(preview.binding) !== textColumnBindingKey(context.binding)) throw new TextColumnsRetired();
  const fresh = await previewTextColumns(context, preview.options, { signal });
  check(context, signal, true);
  if (fresh.rejections) throw Error('Resolve the rejected conversions before splitting.');
  if (fresh.fingerprint !== preview.fingerprint)
    throw Error('The source or destination changed. Preview the split again.');
  if (fresh.collisions && replaceFingerprint !== fresh.fingerprint)
    throw Error('Review the occupied destination cells before replacing them.');
  beforeAcceptance?.();
  check(context, signal, true);
  const pending = context.accept!(fresh.commands);
  if (!pending || typeof pending.then !== 'function')
    throw Error('Text to Columns requires an acknowledged workbook command.');
  await pending;
  return {
    outcome: 'accepted',
    source: fresh.source,
    destination: fresh.destination,
    positions: fresh.outputs.length,
    current: context.isCurrent(context.binding),
    runtimeCurrent: context.isRuntimeCurrent?.(context.binding) ?? context.isCurrent(context.binding),
  };
}
