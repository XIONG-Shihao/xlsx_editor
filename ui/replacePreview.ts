import type { WorkbookCell, WorkbookCommand, WorkbookEdit } from '../core/contracts';
import { hasExactNumericValue } from '../core/calculation.mjs';
import { displayCell } from '../core/display.mjs';
import { VALUE_PASTE_ERRORS } from '../core/value-paste.mjs';
import { captureHomeBinding, HomeCommandRetired, normalizeHomeBatch, type HomeBinding } from './homeCommands';
import { loadSearchPage, type SearchQuery, type SearchCursor, type SearchMatch, type SearchMode } from './searchPages';

export const REPLACE_PREVIEW_LIMITS = Object.freeze({
  matches: 10_000,
  page: 100,
  ownerReply: 100,
  actionBytes: 2 * 1024 * 1024,
  rowsBytes: 16 * 1024 * 1024,
  workPages: 100_000,
});
export interface ReplaceConfiguration {
  readonly text: string;
  readonly replacement: string;
  readonly mode: SearchMode;
  readonly sheets: readonly string[];
  readonly literal: boolean;
  readonly date1904?: boolean;
}
export interface ReplaceRow {
  readonly sheet: string;
  readonly cell: string;
  readonly before: string;
  readonly after: string | null;
  readonly edit: WorkbookEdit | null;
  readonly reason: string | null;
}
export interface ReplacePreview {
  readonly schema: 1;
  readonly binding: HomeBinding;
  readonly configuration: ReplaceConfiguration;
  readonly rows: readonly ReplaceRow[];
  readonly matched: number;
  readonly editable: number;
  readonly skipped: number;
  readonly commands: readonly WorkbookCommand[];
}
export interface ReplaceProgress {
  readonly matched: number;
  readonly editable: number;
  readonly skipped: number;
}
export interface ReplacePreviewPage {
  readonly offset: number;
  readonly total: number;
  readonly rows: readonly ReplaceRow[];
  readonly nextOffset: number | null;
}
export interface CollectReplacementInput {
  binding: HomeBinding;
  configuration: ReplaceConfiguration;
  query: SearchQuery;
  isCurrent: () => boolean;
  signal?: AbortSignal;
  yieldWork?: () => Promise<void>;
  progress?: (value: ReplaceProgress) => void;
}
const completePreviews = new WeakSet<object>();

const own = (value: object, key: string) => Object.hasOwn(value, key);
const plain = (value: unknown): value is Record<string, unknown> =>
  !!value &&
  typeof value === 'object' &&
  !Array.isArray(value) &&
  (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null);
function exact(value: unknown, required: string[], optional: string[] = []): value is Record<string, unknown> {
  return (
    plain(value) &&
    required.every((key) => own(value, key)) &&
    Object.keys(value).every((key) => required.includes(key) || optional.includes(key))
  );
}
function xmlText(value: unknown, max = 32767): value is string {
  return (
    typeof value === 'string' &&
    value.length <= max &&
    !/[\u0000-\u0008\u000b\u000c\u000e-\u001f\uD800-\uDFFF\uFFFE\uFFFF]/u.test(value)
  );
}
const utf8 = (value: unknown) => new TextEncoder().encode(JSON.stringify(value)).length;
function freeze<T>(value: T): T {
  if (value && typeof value === 'object') {
    for (const child of Object.values(value)) freeze(child);
    Object.freeze(value);
  }
  return value;
}
function position(ref: string): number {
  const match = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(ref);
  if (!match) throw Error('Invalid replacement cell reference.');
  const column = [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0),
    row = Number(match[2]);
  if (column > 16384 || row > 1048576) throw Error('Replacement cell is outside the worksheet.');
  return (row - 1) * 16384 + column - 1;
}
function configuration(value: ReplaceConfiguration): ReplaceConfiguration {
  if (
    !exact(value, ['text', 'replacement', 'mode', 'sheets', 'literal'], ['date1904']) ||
    !xmlText(value.text) ||
    !value.text ||
    !xmlText(value.replacement) ||
    !['display', 'raw', 'formula'].includes(value.mode) ||
    typeof value.literal !== 'boolean' ||
    (own(value, 'date1904') && typeof value.date1904 !== 'boolean') ||
    !Array.isArray(value.sheets) ||
    !value.sheets.length ||
    !value.sheets.every((sheet) => xmlText(sheet, 4096) && !!sheet) ||
    new Set(value.sheets).size !== value.sheets.length
  )
    throw Error('Invalid replacement configuration.');
  return freeze({ ...value, sheets: [...value.sheets] }) as ReplaceConfiguration;
}
function cellSnapshot(value: WorkbookCell): WorkbookCell {
  if (
    !plain(value) ||
    !own(value, 'value') ||
    (value.value !== null && !['string', 'number', 'boolean'].includes(typeof value.value)) ||
    (typeof value.value === 'number' && !Number.isFinite(value.value)) ||
    (typeof value.value === 'string' && !xmlText(value.value)) ||
    (value.rawValue !== undefined && value.rawValue !== null && !xmlText(value.rawValue)) ||
    (value.formula !== undefined && (!xmlText(value.formula, 8192) || !value.formula)) ||
    ['sheet', 'cell', 'documentId', 'state'].some((key) => own(value, key))
  )
    throw Error('Search owner returned a foreign or malformed replacement cell.');
  try {
    return freeze(structuredClone(value));
  } catch {
    throw Error('Search owner returned an invalid replacement cell snapshot.');
  }
}
function content(value: WorkbookCell, config: ReplaceConfiguration): string {
  return config.mode === 'formula'
    ? value.formula
      ? '=' + value.formula
      : ''
    : config.mode === 'raw'
      ? String(value.rawValue ?? value.value ?? '')
      : displayCell(value, config.date1904 ?? false);
}
function replaceLiteral(source: string, text: string, replacement: string, max = 32767): string | null {
  const pattern = new RegExp(text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
  let size = source.length,
    match: RegExpExecArray | null;
  while ((match = pattern.exec(source)) !== null) {
    size += replacement.length - match[0].length;
    if (size > max) return null;
  }
  return source.replace(pattern, () => replacement);
}
function skip(match: SearchMatch, before: string, reason: string, after: string | null = null): ReplaceRow {
  return freeze({ sheet: match.sheet, cell: match.cell, before, after, edit: null, reason });
}
function proposal(match: SearchMatch, before: string, after: string, edit: WorkbookEdit): ReplaceRow {
  return freeze({ sheet: match.sheet, cell: match.cell, before, after, edit, reason: null });
}
function replacement(match: SearchMatch, config: ReplaceConfiguration): ReplaceRow {
  if (
    !exact(match, ['sheet', 'cell', 'value']) ||
    typeof match.sheet !== 'string' ||
    !config.sheets.includes(match.sheet) ||
    typeof match.cell !== 'string'
  )
    throw Error('Replacement match belongs to a foreign worksheet.');
  position(match.cell);
  const cell = cellSnapshot(match.value),
    before = content(cell, config);
  if (!xmlText(before)) throw Error('Replacement searched text exceeds cell/XML limits.');
  const record = cell as WorkbookCell & { protected?: boolean; readOnly?: boolean; protection?: unknown };
  if (record.protected === true || record.readOnly === true || record.protection !== undefined)
    return skip(match, before, 'Known protected source cannot be replaced.');
  const ownedMaster =
    cell.spill?.schema === 1 &&
    cell.spill.role === 'master' &&
    cell.spill.master === match.cell &&
    typeof cell.spill.range === 'string';
  if (cell.formulaOutput || cell.scenarioResult || (cell.spill && !ownedMaster) || (cell.formulaType && !ownedMaster))
    return skip(match, before, 'Special or derived output cannot be replaced directly.');
  if (config.mode === 'formula') {
    if (config.literal) return skip(match, before, 'Literal workbook text is not an authored formula.');
    if (!cell.formula) return skip(match, before, 'No authored formula matches this policy.');
    const after = replaceLiteral(before, config.text, config.replacement, 8193);
    if (after === null) return skip(match, before, 'Replacement exceeds formula/XML limits.');
    if (after === before) return skip(match, before, 'No effective change.', after);
    if (
      !after.startsWith('=') ||
      !after.slice(1) ||
      !xmlText(after.slice(1), 8192) ||
      /https?:|file:|(?:DDE|WEBSERVICE|HYPERLINK)\s*\(|\[[^\]]*\.(?:xlsx?|xlsm)\]/i.test(after.slice(1))
    )
      return skip(match, before, 'Replacement must retain a bounded supported authored formula.', after);
    return proposal(match, before, after, {
      sheet: match.sheet,
      cell: match.cell,
      formula: after.slice(1),
      expectedFormula: cell.formula,
      expectedValue: cell.value,
      expectedRawValue: cell.rawValue ?? null,
    });
  }
  if (cell.formula || cell.calculation || cell.cachedResult)
    return skip(match, before, 'Formula, calculated or stored results cannot be replaced directly.');
  if (cell.value === null) return skip(match, before, 'Blank cells have no editable source text.');
  const after = replaceLiteral(before, config.text, config.replacement);
  if (after === null) return skip(match, before, 'Replacement exceeds cell text/XML limits.');
  if (after === before) return skip(match, before, 'No effective change.', after);
  if (!xmlText(after)) return skip(match, before, 'Replacement exceeds cell text/XML limits.', after);
  const guard = {
    sheet: match.sheet,
    cell: match.cell,
    expectedValue: cell.value,
    expectedFormula: null,
    expectedRawValue: cell.rawValue ?? null,
  };
  if (cell.dataType === 'error') {
    if (
      typeof cell.value !== 'string' ||
      !VALUE_PASTE_ERRORS.includes(cell.value) ||
      !VALUE_PASTE_ERRORS.includes(after)
    )
      return skip(match, before, 'A genuine error may change only to a canonical typed error.', after);
    return proposal(match, before, after, { ...guard, value: after, valueType: 'error' });
  }
  if (typeof cell.value === 'string') return proposal(match, before, after, { ...guard, value: after });
  if (typeof cell.value === 'boolean') {
    if (!/^(true|false)$/i.test(after))
      return skip(match, before, 'Boolean replacement must remain TRUE or FALSE.', after);
    const value = after.toLowerCase() === 'true';
    return value === cell.value
      ? skip(match, before, 'No effective Boolean change.', after)
      : proposal(match, before, after, { ...guard, value });
  }
  if (typeof cell.value === 'number') {
    if (!hasExactNumericValue(cell.value, cell.rawValue))
      return skip(match, before, 'Stored numeric precision is not qualified for replacement.', after);
    if (config.mode === 'display' && before !== String(cell.value) && before !== cell.rawValue)
      return skip(match, before, 'Formatted numeric display is not an exact raw numeric source.', after);
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(after))
      return skip(match, before, 'Numeric replacement must remain an exact finite number.', after);
    const value = Number(after);
    if (!Number.isFinite(value) || !hasExactNumericValue(value, after))
      return skip(match, before, 'Numeric replacement would lose precision or overflow.', after);
    return Object.is(value, cell.value)
      ? skip(match, before, 'No effective numeric change.', after)
      : proposal(match, before, after, { ...guard, value });
  }
  return skip(match, before, 'Source type is not qualified for replacement.');
}
/** A proposal is not a native admission certificate; Apply must re-collect
 * under the captured owner and native source/protection guards. */
export function replacementForMatch(match: SearchMatch, config: ReplaceConfiguration): ReplaceRow {
  return replacement(match, configuration(config));
}
function commandsFor(rows: readonly ReplaceRow[]): WorkbookCommand[] {
  const grouped = new Map<string, Array<Omit<WorkbookEdit, 'sheet'>>>();
  for (const row of rows) {
    if (!row.edit) continue;
    const { sheet, ...edit } = row.edit;
    const group = grouped.get(sheet) ?? [];
    group.push(structuredClone(edit));
    grouped.set(sheet, group);
  }
  const commands = [...grouped].map(([sheet, cells]) => ({ type: 'setCells' as const, sheet, cells }));
  return commands.length ? normalizeHomeBatch(commands) : [];
}
function actionBytes(binding: HomeBinding, commands: readonly WorkbookCommand[]): number {
  return utf8({ operation: 'apply', documentId: binding.documentId, state: binding.state, commands });
}

export async function collectReplacementPreview(input: CollectReplacementInput): Promise<ReplacePreview> {
  if (
    !exact(input, ['binding', 'configuration', 'query', 'isCurrent'], ['signal', 'yieldWork', 'progress']) ||
    typeof input.query !== 'function' ||
    typeof input.isCurrent !== 'function' ||
    (input.yieldWork !== undefined && typeof input.yieldWork !== 'function') ||
    (input.progress !== undefined && typeof input.progress !== 'function')
  )
    throw Error('Invalid replacement preview input.');
  const binding = captureHomeBinding(input.binding),
    config = configuration(input.configuration);
  const current = () => {
    if (input.signal?.aborted) {
      const error = new Error('Replacement preview cancelled.');
      error.name = 'AbortError';
      throw error;
    }
    if (!input.isCurrent()) throw new HomeCommandRetired();
  };
  current();
  let workPages = 0,
    candidateBytes = 2;
  const candidates = new Map<string, ReplaceRow>();
  const query: SearchQuery = async (sheet, text, mode, after, options) => {
    current();
    if (!config.sheets.includes(sheet) || text !== config.text || mode !== config.mode)
      throw Error('Replacement query changed its captured configuration.');
    if (++workPages > REPLACE_PREVIEW_LIMITS.workPages)
      throw Error('Replacement preview exceeded 100,000 owner pages. Narrow the search scope.');
    const limit = Math.min(REPLACE_PREVIEW_LIMITS.ownerReply, options?.matchLimit ?? REPLACE_PREVIEW_LIMITS.ownerReply);
    const reply = await input.query(sheet, text, mode, after, { matchLimit: limit });
    current();
    if (
      !exact(reply, ['cells', 'matches', 'nextCursor']) ||
      !plain(reply.cells) ||
      !Number.isInteger(reply.matches) ||
      reply.matches < 0 ||
      reply.matches > limit ||
      reply.matches !== Object.keys(reply.cells).length ||
      (reply.nextCursor !== null && typeof reply.nextCursor !== 'string') ||
      utf8(reply) > 16 * 1024 * 1024
    )
      throw Error('Search owner returned an incomplete or oversized replacement page.');
    const cells: Record<string, WorkbookCell> = {};
    for (const [ref, value] of Object.entries(reply.cells)) {
      position(ref);
      const snapshot = cellSnapshot(value);
      if (!content(snapshot, config).toLocaleLowerCase().includes(config.text.toLocaleLowerCase()))
        throw Error('Search owner returned a foreign non-matching replacement cell.');
      const key = sheet + '\0' + ref;
      if (candidates.has(key)) throw Error('Search owner returned duplicate replacement matches.');
      if (candidates.size === REPLACE_PREVIEW_LIMITS.matches)
        throw Error('Replacement preview exceeds 10,000 matches. Narrow the search scope.');
      const row = replacement({ sheet, cell: ref, value: snapshot }, config);
      candidateBytes += utf8(row) + (candidates.size ? 1 : 0);
      if (candidateBytes > REPLACE_PREVIEW_LIMITS.rowsBytes)
        throw Error('Replacement preview exceeds its 16 MiB captured-row limit. Narrow the search scope.');
      candidates.set(key, row);
      cells[ref] = snapshot;
    }
    return { cells, matches: reply.matches, nextCursor: reply.nextCursor };
  };
  const rows: ReplaceRow[] = [],
    seen = new Set<string>();
  let cursor: SearchCursor | undefined,
    rowsBytes = 2,
    editable = 0,
    lastSheet = -1,
    lastPosition = -1;
  const cancelled = () => input.signal?.aborted === true || !input.isCurrent();
  const yieldWork = async () => {
    current();
    if (input.yieldWork) await input.yieldWork();
    current();
  };
  do {
    current();
    const page = await loadSearchPage({
      sheets: [...config.sheets],
      text: config.text,
      mode: config.mode,
      query,
      cursor,
      cancelled,
      yieldWork,
    });
    current();
    if (!page) throw new HomeCommandRetired();
    for (const match of page.matches) {
      const sheetIndex = config.sheets.indexOf(match.sheet),
        at = position(match.cell),
        key = match.sheet + '\0' + match.cell;
      if (sheetIndex < 0 || sheetIndex < lastSheet || (sheetIndex === lastSheet && at <= lastPosition) || seen.has(key))
        throw Error('Replacement owner returned duplicate, foreign or unordered matches.');
      if (rows.length === REPLACE_PREVIEW_LIMITS.matches)
        throw Error('Replacement preview exceeds 10,000 matches. Narrow the search scope.');
      const row = candidates.get(key);
      if (!row) throw Error('Replacement match has no captured proposal.');
      rowsBytes += utf8(row) + (rows.length ? 1 : 0);
      if (rowsBytes > REPLACE_PREVIEW_LIMITS.rowsBytes)
        throw Error('Replacement preview exceeds its 16 MiB captured-row limit. Narrow the search scope.');
      rows.push(row);
      seen.add(key);
      lastSheet = sheetIndex;
      lastPosition = at;
      if (row.edit) editable++;
    }
    current();
    input.progress?.(freeze({ matched: rows.length, editable, skipped: rows.length - editable }));
    current();
    cursor = page.next ?? undefined;
  } while (cursor);
  // At the inclusive match cap, obtain a further accepted end-of-search read.
  if (rows.length === REPLACE_PREVIEW_LIMITS.matches) {
    const last = rows.at(-1)!;
    const end = await loadSearchPage({
      sheets: [...config.sheets],
      text: config.text,
      mode: config.mode,
      query,
      cursor: { sheet: config.sheets.indexOf(last.sheet), after: last.cell },
      cancelled,
      yieldWork,
    });
    current();
    if (!end || end.matches.length || end.next !== null)
      throw Error('Replacement preview exceeds 10,000 matches or lacks an end certificate.');
  }
  const commands = commandsFor(rows);
  if (actionBytes(binding, commands) > REPLACE_PREVIEW_LIMITS.actionBytes)
    throw Error('Replacement action exceeds its 2 MiB escaped JSON limit. Narrow the search scope.');
  current();
  const preview = freeze({
    schema: 1 as const,
    binding,
    configuration: config,
    rows,
    matched: rows.length,
    editable,
    skipped: rows.length - editable,
    commands,
  });
  completePreviews.add(preview);
  return preview;
}

export function replacementPreviewPage(preview: ReplacePreview, offset = 0, limit = 100): ReplacePreviewPage {
  if (
    !exact(preview, ['schema', 'binding', 'configuration', 'rows', 'matched', 'editable', 'skipped', 'commands']) ||
    preview.schema !== 1 ||
    !Array.isArray(preview.rows) ||
    !Array.isArray(preview.commands) ||
    !Number.isInteger(preview.matched) ||
    preview.matched !== preview.rows.length ||
    preview.matched > REPLACE_PREVIEW_LIMITS.matches ||
    !Number.isInteger(preview.editable) ||
    !Number.isInteger(preview.skipped) ||
    preview.editable < 0 ||
    preview.skipped < 0 ||
    preview.editable + preview.skipped !== preview.matched ||
    !Number.isInteger(offset) ||
    offset < 0 ||
    offset > preview.matched ||
    !Number.isInteger(limit) ||
    limit < 1 ||
    limit > REPLACE_PREVIEW_LIMITS.page
  )
    throw Error('Invalid replacement preview page.');
  // Only the immutable complete collector result bypasses repeated full-plan
  // validation. Copied/caller-built objects receive the complete schema audit.
  if (completePreviews.has(preview)) {
    const end = Math.min(offset + limit, preview.matched);
    return freeze({
      offset,
      total: preview.matched,
      rows: structuredClone(preview.rows.slice(offset, end)),
      nextOffset: end < preview.matched ? end : null,
    });
  }
  const binding = captureHomeBinding(preview.binding),
    config = configuration(preview.configuration);
  let editable = 0,
    bytes = 2,
    lastSheet = -1,
    lastPosition = -1,
    rowIndex = 0;
  for (const row of preview.rows) {
    if (
      !exact(row, ['sheet', 'cell', 'before', 'after', 'edit', 'reason']) ||
      typeof row.sheet !== 'string' ||
      typeof row.cell !== 'string' ||
      !config.sheets.includes(row.sheet) ||
      !xmlText(row.before) ||
      (row.after !== null && !xmlText(row.after)) ||
      (row.reason !== null && typeof row.reason !== 'string')
    )
      throw Error('Invalid replacement preview row.');
    const index = config.sheets.indexOf(row.sheet),
      at = position(row.cell);
    if (index < lastSheet || (index === lastSheet && at <= lastPosition))
      throw Error('Invalid replacement preview row order.');
    lastSheet = index;
    lastPosition = at;
    if (row.edit) {
      const edit = row.edit;
      if (
        !exact(
          edit,
          ['sheet', 'cell', 'expectedValue', 'expectedRawValue', 'expectedFormula'],
          ['value', 'formula', 'valueType']
        ) ||
        edit.sheet !== row.sheet ||
        edit.cell !== row.cell ||
        row.reason !== null ||
        row.after === null ||
        own(edit, 'value') === own(edit, 'formula') ||
        (edit.expectedValue !== null && !['string', 'number', 'boolean'].includes(typeof edit.expectedValue)) ||
        (typeof edit.expectedValue === 'number' && !Number.isFinite(edit.expectedValue)) ||
        (edit.expectedRawValue !== null && !xmlText(edit.expectedRawValue)) ||
        (edit.expectedFormula !== null && !xmlText(edit.expectedFormula, 8192))
      )
        throw Error('Invalid replacement edit or foreign style field.');
      if (own(edit, 'formula')) {
        if (
          config.mode !== 'formula' ||
          config.literal ||
          !edit.expectedFormula ||
          !xmlText(edit.formula, 8192) ||
          !edit.formula ||
          row.after !== '=' + edit.formula ||
          row.before !== '=' + edit.expectedFormula ||
          own(edit, 'valueType')
        )
          throw Error('Invalid replacement formula proposal.');
      } else {
        if (
          config.mode === 'formula' ||
          edit.expectedFormula !== null ||
          edit.expectedValue === null ||
          (edit.value !== null && !['string', 'number', 'boolean'].includes(typeof edit.value)) ||
          (typeof edit.value === 'number' && !Number.isFinite(edit.value)) ||
          typeof edit.value !== typeof edit.expectedValue
        )
          throw Error('Replacement proposal changes its source type.');
        if (own(edit, 'valueType')) {
          if (
            edit.valueType !== 'error' ||
            !VALUE_PASTE_ERRORS.includes(edit.value as string) ||
            row.after !== edit.value
          )
            throw Error('Invalid typed replacement error.');
        } else if (
          (typeof edit.value === 'string' && row.after !== edit.value) ||
          (typeof edit.value === 'boolean' && row.after.toLowerCase() !== String(edit.value)) ||
          (typeof edit.value === 'number' && !hasExactNumericValue(edit.value, row.after))
        )
          throw Error('Replacement proposal disagrees with its displayed change.');
      }
      editable++;
    } else if (!row.reason) throw Error('Skipped replacement row lacks a reason.');
    bytes += utf8(row) + (rowIndex++ ? 1 : 0);
  }
  if (
    editable !== preview.editable ||
    bytes > REPLACE_PREVIEW_LIMITS.rowsBytes ||
    JSON.stringify(commandsFor(preview.rows)) !== JSON.stringify(preview.commands) ||
    actionBytes(binding, preview.commands) > REPLACE_PREVIEW_LIMITS.actionBytes
  )
    throw Error('Invalid replacement preview command/count/budget.');
  const end = Math.min(offset + limit, preview.matched);
  return freeze({
    offset,
    total: preview.matched,
    rows: structuredClone(preview.rows.slice(offset, end)),
    nextOffset: end < preview.matched ? end : null,
  });
}
