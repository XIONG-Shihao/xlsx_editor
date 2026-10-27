import type { WorkbookFillAcrossSheetsCommand, WorkbookFillAcrossSheetsMode } from '../core/contracts';
import { FILL_ACROSS_SHEETS_LIMITS, normalizeFillAcrossSheetsRuns } from '../core/fill-across-sheets.mjs';
import {
  captureHomeBinding,
  normalizeHomeArguments,
  type HomeBinding,
  type HomeFillAcrossOptions,
} from './homeCommands';

export const FILL_ACROSS_LIMITS = Object.freeze({
  positions: FILL_ACROSS_SHEETS_LIMITS.destinationPositions,
  actionBytes: FILL_ACROSS_SHEETS_LIMITS.actionBytes,
  page: 100,
});
export interface FillAcrossSheet {
  readonly name: string;
  readonly state?: string;
}
export interface FillAcrossPreview {
  readonly schema: 1;
  readonly binding: HomeBinding;
  readonly sourceSheet: string;
  readonly range: string;
  readonly mode: WorkbookFillAcrossSheetsMode;
  readonly targets: readonly string[];
  readonly positionsPerSheet: number;
  readonly destinationPositions: number;
  readonly commands: readonly WorkbookFillAcrossSheetsCommand[];
}
const name = (value: unknown): value is string =>
  typeof value === 'string' &&
  value.length > 0 &&
  value.length <= 31 &&
  !/[\u0000-\u001f\uD800-\uDFFF\uFFFE\uFFFF\[\]:*?\/\\]/u.test(value);
function positions(range: string): number {
  const address = (value: string) => {
    const match = /^([A-Z]{1,3})([1-9]\d*)$/.exec(value);
    if (!match) throw Error('Across Worksheets needs a canonical rectangular range.');
    const row = Number(match[2]),
      column = [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
    if (row > 1048576 || column > 16384) throw Error('Across Worksheets range is outside the worksheet.');
    return { row, column };
  };
  const parts = range.split(':');
  if (parts.length > 2) throw Error('Across Worksheets needs one rectangular range.');
  const first = address(parts[0]),
    last = address(parts[1] ?? parts[0]),
    rows = last.row - first.row + 1,
    columns = last.column - first.column + 1;
  if (rows < 1 || columns < 1) throw Error('Across Worksheets range must run from its first to last cell.');
  return rows * columns;
}
/** Capture the accepted worksheet list, never worksheet cells or rendered styles. */
export function captureFillAcrossSheets(
  sheets: readonly FillAcrossSheet[],
  sourceSheet: string
): readonly FillAcrossSheet[] {
  if (!Array.isArray(sheets) || !sheets.length || !name(sourceSheet))
    throw Error('The original worksheet inventory is unavailable.');
  const names = new Set<string>(),
    captured = sheets.map((sheet) => {
      if (
        !sheet ||
        !name(sheet.name) ||
        names.has(sheet.name) ||
        !['visible', 'hidden', 'veryHidden'].includes(sheet.state ?? 'visible')
      )
        throw Error('The original worksheet inventory is invalid.');
      names.add(sheet.name);
      return Object.freeze({ name: sheet.name, state: sheet.state ?? 'visible' });
    });
  if (!captured.some((sheet) => sheet.name === sourceSheet && sheet.state === 'visible'))
    throw Error('The original source worksheet is not visible.');
  return Object.freeze(captured);
}
/** This preview describes an exact native command plan. Native admission owns
 * contents, formulas, styles, protections and the atomic before-state. */
export function createFillAcrossPreview(
  binding: HomeBinding,
  sheets: readonly FillAcrossSheet[],
  options: HomeFillAcrossOptions
): FillAcrossPreview {
  const captured = captureHomeBinding(binding),
    inventory = captureFillAcrossSheets(sheets, captured.sheet),
    normalized = normalizeHomeArguments('home.fill-across-preview', options);
  if (!captured.ownerIdentity || captured.mode !== 'normal' || captured.specialSelection)
    throw Error('Across Worksheets needs one regular original-owner range in Normal view.');
  const chosen = new Set(normalized.targets),
    available = inventory.filter((sheet) => sheet.name !== captured.sheet && sheet.state === 'visible');
  if (chosen.has(captured.sheet) || [...chosen].some((target) => !available.some((sheet) => sheet.name === target)))
    throw Error('Choose visible destination worksheets other than the source worksheet.');
  const targets = available.filter((sheet) => chosen.has(sheet.name)).map((sheet) => sheet.name),
    positionsPerSheet = positions(captured.selection),
    destinationPositions = positionsPerSheet * targets.length;
  if (destinationPositions > FILL_ACROSS_LIMITS.positions)
    throw Error('Across Worksheets supports up to 10,000 destination positions in total.');
  const commands = targets.map((sheet) =>
    Object.freeze({
      type: 'fillAcrossSheets' as const,
      sheet,
      sourceSheet: captured.sheet,
      range: captured.selection,
      mode: normalized.mode,
    })
  );
  normalizeFillAcrossSheetsRuns(commands); // The shared native command contract also validates canonical addresses and action bytes.
  return Object.freeze({
    schema: 1,
    binding: captured,
    sourceSheet: captured.sheet,
    range: captured.selection,
    mode: normalized.mode,
    targets: Object.freeze(targets),
    positionsPerSheet,
    destinationPositions,
    commands: Object.freeze(commands),
  });
}
