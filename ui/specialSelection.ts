import type {
  WorkbookCell,
  WorkbookSheet,
  WorkbookConditionalDefinition,
  WorkbookValidationDefinition,
  WorkbookUsedRangeResult,
  WorkbookUsedRangeReceipt,
} from '../core/contracts';
import { collectCurrentRegion } from './currentRegion';
import { validateUsedRange } from '../core/used-range.mjs';
import { captureHomeBinding, HomeCommandRetired, type HomeBinding } from './homeCommands';

export const SPECIAL_SELECTION_LIMITS = Object.freeze({
  positions: 10_000,
  readPositions: 1000,
  page: 100,
  replyBytes: 16 * 1024 * 1024,
});
export const SPECIAL_SELECTION_KINDS = Object.freeze([
  { id: 'notes', label: 'Notes', supported: false },
  { id: 'constants', label: 'Constants', supported: true },
  { id: 'formulas', label: 'Formulas', supported: true },
  { id: 'blanks', label: 'Blanks', supported: true },
  { id: 'current-region', label: 'Current region', supported: true },
  { id: 'current-array', label: 'Current array', supported: false },
  { id: 'row-differences', label: 'Row differences', supported: false },
  { id: 'column-differences', label: 'Column differences', supported: false },
  { id: 'precedents', label: 'Precedents', supported: false },
  { id: 'dependents', label: 'Dependents', supported: false },
  { id: 'last-cell', label: 'Last cell', supported: true },
  { id: 'visible', label: 'Visible cells only', supported: true },
  { id: 'objects', label: 'Objects', supported: false },
  { id: 'conditional-formats', label: 'Conditional formats', supported: true },
  { id: 'data-validation', label: 'Data validation', supported: true },
] as const);
export type SpecialSelectionKind = (typeof SPECIAL_SELECTION_KINDS)[number]['id'];
export interface SpecialSelectionTypes {
  readonly number: boolean;
  readonly text: boolean;
  readonly logical: boolean;
  readonly error: boolean;
}
export interface SpecialSelectionOptions {
  readonly kind: SpecialSelectionKind;
  readonly types: SpecialSelectionTypes;
}
export const DEFAULT_SPECIAL_SELECTION_OPTIONS: SpecialSelectionOptions = Object.freeze({
  kind: 'constants',
  types: Object.freeze({ number: true, text: true, logical: true, error: true }),
});
/** Complete means the trusted original-owner read returned every authored cell
 * in this exact bounded range. Only then may absent records mean true blanks. */
export interface SpecialSelectionRangeReply {
  readonly range: string;
  readonly complete: true;
  readonly cells: Record<string, WorkbookCell>;
}
export type SpecialSelectionReader = (
  sheet: string,
  range: string,
  signal?: AbortSignal
) => Promise<SpecialSelectionRangeReply>;
/** Accepted worksheet merge metadata, independent of viewport cells. */
export type SpecialSelectionGeometry = Pick<WorkbookSheet, 'name' | 'merges'>;
export type SpecialSelectionVisibility = Pick<
  WorkbookSheet,
  'name' | 'hiddenRows' | 'visibleRows' | 'sheetFormat' | 'widths' | 'rowHeights' | 'geometryIssues'
>;
export interface SpecialSelectionRuleReceipt {
  readonly complete: true;
  readonly definition: WorkbookConditionalDefinition | WorkbookValidationDefinition;
}
export interface SpecialSelectionRow {
  readonly cell: string;
  readonly classification: string;
}
export interface SpecialSelectionScope {
  readonly kind: 'selection' | 'used-range' | 'last-cell' | 'current-region';
  readonly usedRange: WorkbookUsedRangeResult | null;
  /** Original provider authority; Home's view aliases do not replace it. */
  readonly receipt: WorkbookUsedRangeReceipt | null;
}
export interface SpecialSelectionPreview {
  readonly schema: 1;
  readonly binding: HomeBinding;
  readonly options: SpecialSelectionOptions;
  readonly geometry: SpecialSelectionGeometry;
  readonly scope: SpecialSelectionScope;
  readonly activeCell: string | null;
  readonly sheet: string;
  readonly range: string;
  readonly positions: number;
  readonly rows: readonly SpecialSelectionRow[];
  readonly references: readonly string[];
  readonly readRanges: readonly string[];
}
export interface CollectSpecialSelectionInput {
  binding: HomeBinding;
  options: SpecialSelectionOptions;
  geometry: SpecialSelectionGeometry;
  read: SpecialSelectionReader;
  isCurrent: () => boolean;
  signal?: AbortSignal;
  yieldWork?: () => Promise<void>;
  calculationRevision?: string;
  visibility?: SpecialSelectionVisibility;
  rules?: SpecialSelectionRuleReceipt;
  usedRange?: WorkbookUsedRangeResult;
  usedRangeReceipt?: WorkbookUsedRangeReceipt;
}
export interface SpecialSelectionPage {
  readonly offset: number;
  readonly total: number;
  readonly rows: readonly SpecialSelectionRow[];
  readonly nextOffset: number | null;
}
const plain = (value: unknown): value is Record<string, unknown> =>
  !!value &&
  typeof value === 'object' &&
  !Array.isArray(value) &&
  (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null);
const exact = (value: unknown, required: string[], optional: string[] = []): value is Record<string, unknown> =>
  plain(value) &&
  required.every((k) => Object.hasOwn(value, k)) &&
  Object.keys(value).every((k) => required.includes(k) || optional.includes(k));
const text = (value: unknown, max = 32767): value is string =>
  typeof value === 'string' &&
  value.length <= max &&
  !/[\u0000-\u0008\u000b\u000c\u000e-\u001f\uD800-\uDFFF\uFFFE\uFFFF]/u.test(value);
function freeze<T>(value: T): T {
  if (value && typeof value === 'object') {
    for (const child of Object.values(value)) freeze(child);
    Object.freeze(value);
  }
  return value;
}
const bytes = (value: unknown) => new TextEncoder().encode(JSON.stringify(value)).length;
function address(value: string) {
  const match = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(value);
  if (!match) throw Error('Go To Special needs canonical worksheet references.');
  const column = [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0),
    row = Number(match[2]);
  if (column > 16384 || row > 1048576) throw Error('Go To Special range is outside the worksheet.');
  return { row: row - 1, column: column - 1 };
}
function rectangle(value: string) {
  const parts = value.split(':');
  if (parts.length > 2) throw Error('Go To Special needs one rectangular range.');
  const first = address(parts[0]),
    last = address(parts[1] ?? parts[0]);
  if (first.row > last.row || first.column > last.column)
    throw Error('Go To Special range must run from its first to last cell.');
  return { first, last, rows: last.row - first.row + 1, columns: last.column - first.column + 1 };
}
function columnName(value: number) {
  let name = '';
  for (let n = value + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}
const ref = (row: number, column: number) => columnName(column) + (row + 1);
const inside = (at: ReturnType<typeof address>, range: ReturnType<typeof rectangle>) =>
  at.row >= range.first.row &&
  at.row <= range.last.row &&
  at.column >= range.first.column &&
  at.column <= range.last.column;
const intersects = (first: ReturnType<typeof rectangle>, last: ReturnType<typeof rectangle>) =>
  first.first.row <= last.last.row &&
  first.last.row >= last.first.row &&
  first.first.column <= last.last.column &&
  first.last.column >= last.first.column;
const order = (value: string) => {
  const at = address(value);
  return at.row * 16384 + at.column;
};
function mergeGeometry(value: SpecialSelectionGeometry, sheet: string) {
  if (
    !plain(value) ||
    value.name !== sheet ||
    !Array.isArray(value.merges) ||
    value.merges.length > SPECIAL_SELECTION_LIMITS.positions
  )
    throw Error('Complete original worksheet merge geometry is unavailable.');
  const merges = value.merges
    .map((range) => {
      if (typeof range !== 'string' || range.length > 32) throw Error('Original worksheet merge geometry is invalid.');
      const parsed = rectangle(range);
      if (parsed.rows * parsed.columns < 2) throw Error('Original worksheet merge geometry needs a merged area.');
      return { ...parsed, anchor: ref(parsed.first.row, parsed.first.column), range };
    })
    .sort((a, b) => a.first.row - b.first.row || a.first.column - b.first.column);
  const active: typeof merges = [];
  for (const merge of merges) {
    for (let i = active.length - 1; i >= 0; i--) if (active[i].last.row < merge.first.row) active.splice(i, 1);
    if (active.some((prior) => intersects(prior, merge))) throw Error('Original worksheet merged areas overlap.');
    active.push(merge);
  }
  return merges;
}
type SpecialArea = ReturnType<typeof rectangle>;
function enclosingArea(first: SpecialArea, last: SpecialArea): SpecialArea {
  const a = { row: Math.min(first.first.row, last.first.row), column: Math.min(first.first.column, last.first.column) },
    b = { row: Math.max(first.last.row, last.last.row), column: Math.max(first.last.column, last.last.column) };
  return { first: a, last: b, rows: b.row - a.row + 1, columns: b.column - a.column + 1 };
}
function closedArea(value: string, merges: ReturnType<typeof mergeGeometry>): SpecialArea {
  let bounds = rectangle(value),
    changed = true;
  while (changed) {
    changed = false;
    for (const merge of merges)
      if (intersects(bounds, merge) && (!inside(merge.first, bounds) || !inside(merge.last, bounds))) {
        bounds = enclosingArea(bounds, merge);
        changed = true;
        if (bounds.rows * bounds.columns > SPECIAL_SELECTION_LIMITS.positions)
          throw Error('Merged Go To Special scope exceeds 10,000 source positions.');
      }
  }
  if (bounds.rows * bounds.columns > SPECIAL_SELECTION_LIMITS.positions)
    throw Error('Go To Special supports up to 10,000 selected positions.');
  return bounds;
}
function areaReference(area: SpecialArea): string {
  const first = ref(area.first.row, area.first.column),
    last = ref(area.last.row, area.last.column);
  return first === last ? first : first + ':' + last;
}
/** A single merged area behaves as one cell for default Special scope. */
export function specialSelectionNeedsUsedRange(
  selection: string,
  geometry: SpecialSelectionGeometry,
  kind: SpecialSelectionKind
): boolean {
  if (kind === 'last-cell') return true;
  if (!['constants', 'formulas', 'blanks'].includes(kind)) return false;
  const merges = mergeGeometry(geometry, geometry.name),
    selected = closedArea(selection, merges);
  return (
    selected.rows * selected.columns === 1 || merges.some((merge) => areaReference(merge) === areaReference(selected))
  );
}
function selectionScope(
  binding: HomeBinding,
  options: SpecialSelectionOptions,
  geometry: SpecialSelectionGeometry,
  usedRange?: WorkbookUsedRangeResult | null,
  receipt?: WorkbookUsedRangeReceipt | null
): { scope: SpecialSelectionScope; bounds: SpecialArea } {
  const merges = mergeGeometry(geometry, binding.sheet),
    required = specialSelectionNeedsUsedRange(binding.selection, geometry, options.kind);
  if (!required) {
    if (usedRange != null || receipt != null)
      throw Error('An explicit Special selection cannot borrow worksheet extent authority.');
    return {
      scope: { kind: 'selection', usedRange: null, receipt: null },
      bounds: closedArea(binding.selection, merges),
    };
  }
  if (!usedRange || !receipt)
    throw Error('A complete original-owner UsedRange receipt is required for this selection.');
  const captured = validateUsedRange(structuredClone(usedRange), {
    sheet: binding.sheet,
    receipt: structuredClone(receipt),
  });
  if (captured.receipt.documentId !== binding.documentId) throw Error('The UsedRange belongs to another document.');
  if (captured.status !== 'complete')
    throw Error('Worksheet UsedRange is incomplete: ' + captured.reasons.map((reason) => reason.message).join(' '));
  const kind = options.kind === 'last-cell' ? 'last-cell' : 'used-range';
  return {
    scope: { kind, usedRange: captured, receipt: { ...receipt } },
    bounds: closedArea(kind === 'last-cell' ? captured.lastCell : captured.range, merges),
  };
}
/** Preserve areas as whole rectangles; joining an L-shaped union would select
 * unmatched positions. Merge areas are coalesced before ordinary areas join. */
function coalescedAreas(input: SpecialArea[]): SpecialArea[] {
  const areas = input.slice();
  let changed = true;
  while (changed) {
    changed = false;
    outer: for (let i = 0; i < areas.length; i++)
      for (let j = i + 1; j < areas.length; j++) {
        const a = areas[i],
          b = areas[j],
          bounds = enclosingArea(a, b),
          overlap =
            Math.max(0, Math.min(a.last.row, b.last.row) - Math.max(a.first.row, b.first.row) + 1) *
            Math.max(0, Math.min(a.last.column, b.last.column) - Math.max(a.first.column, b.first.column) + 1);
        if (a.rows * a.columns + b.rows * b.columns - overlap === bounds.rows * bounds.columns) {
          areas[i] = bounds;
          areas.splice(j, 1);
          changed = true;
          break outer;
        }
      }
  }
  return areas;
}
function blankActiveCell(rows: SpecialSelectionRow[], merges: ReturnType<typeof mergeGeometry>): string | null {
  const ordinary: SpecialArea[] = [],
    matchedMerges = new Set<string>();
  for (const row of rows) {
    const at = address(row.cell),
      merge = merges.find((area) => inside(at, area));
    if (merge) {
      matchedMerges.add(merge.range);
      continue;
    }
    const prior = ordinary.at(-1);
    if (prior && prior.first.row === at.row && prior.last.row === at.row && prior.last.column + 1 === at.column) {
      prior.last.column = at.column;
      prior.columns++;
    } else ordinary.push({ first: { ...at }, last: { ...at }, rows: 1, columns: 1 });
  }
  // Excel's ordinary blank areas precede merged areas. Among ordinary areas,
  // the leftmost area's first row supplies the initial active cell.
  const regular = coalescedAreas(ordinary).sort((a, b) => a.first.column - b.first.column || a.first.row - b.first.row);
  const merged = coalescedAreas(merges.filter((area) => matchedMerges.has(area.range))),
    areas = coalescedAreas([...regular, ...merged]);
  return areas.length ? ref(areas[0].first.row, areas[0].first.column) : null;
}
function mergedMatches(rows: SpecialSelectionRow[], merges: ReturnType<typeof mergeGeometry>): SpecialSelectionRow[] {
  const matches = new Map(rows.map((row) => [row.cell, row])),
    expanded = new Set<string>();
  for (const row of rows) {
    const at = address(row.cell),
      merge = merges.find((area) => inside(at, area));
    if (!merge || expanded.has(merge.range)) continue;
    expanded.add(merge.range);
    for (let r = merge.first.row; r <= merge.last.row; r++)
      for (let c = merge.first.column; c <= merge.last.column; c++) {
        const cell = ref(r, c);
        if (!matches.has(cell))
          matches.set(cell, { cell, classification: `Matched merged area (${row.classification})` });
      }
  }
  return [...matches.values()];
}
export function normalizeSpecialSelectionOptions(value: SpecialSelectionOptions): SpecialSelectionOptions {
  if (
    !exact(value, ['kind', 'types']) ||
    !SPECIAL_SELECTION_KINDS.some((k) => k.id === value.kind) ||
    !exact(value.types, ['number', 'text', 'logical', 'error']) ||
    Object.values(value.types).some((v) => typeof v !== 'boolean')
  )
    throw Error('Invalid Go To Special options.');
  if (['constants', 'formulas'].includes(value.kind) && !Object.values(value.types).some(Boolean))
    throw Error('Choose at least one cell type.');
  return freeze({ kind: value.kind, types: { ...value.types } }) as SpecialSelectionOptions;
}
function windows(bounds: ReturnType<typeof rectangle>): string[] {
  const result: string[] = [];
  if (bounds.columns <= SPECIAL_SELECTION_LIMITS.readPositions) {
    const height = Math.max(1, Math.floor(SPECIAL_SELECTION_LIMITS.readPositions / bounds.columns));
    for (let row = bounds.first.row; row <= bounds.last.row; row += height)
      result.push(
        ref(row, bounds.first.column) + ':' + ref(Math.min(row + height - 1, bounds.last.row), bounds.last.column)
      );
  } else
    for (let row = bounds.first.row; row <= bounds.last.row; row++)
      for (
        let column = bounds.first.column;
        column <= bounds.last.column;
        column += SPECIAL_SELECTION_LIMITS.readPositions
      )
        result.push(
          ref(row, column) +
            ':' +
            ref(row, Math.min(column + SPECIAL_SELECTION_LIMITS.readPositions - 1, bounds.last.column))
        );
  return result;
}
function sourceCell(value: WorkbookCell): WorkbookCell {
  if (
    !plain(value) ||
    !Object.hasOwn(value, 'value') ||
    (value.value !== null && !['string', 'number', 'boolean'].includes(typeof value.value)) ||
    (typeof value.value === 'number' && !Number.isFinite(value.value)) ||
    (typeof value.value === 'string' && !text(value.value)) ||
    (value.formula !== undefined && (!text(value.formula, 8192) || !value.formula)) ||
    (value.formulaType !== undefined && (!text(value.formulaType, 100) || !value.formulaType)) ||
    (value.dataType !== undefined && !['number', 'text', 'boolean', 'blank', 'error'].includes(value.dataType)) ||
    ['sheet', 'cell', 'state', 'documentId', 'unsupported', 'truncated', 'incomplete'].some((k) =>
      Object.hasOwn(value, k)
    )
  )
    throw Error('Go To Special owner returned an unsupported or malformed source cell.');
  try {
    return freeze(structuredClone(value));
  } catch {
    throw Error('Go To Special source cell cannot be captured.');
  }
}
const authoredFormula = (cell: WorkbookCell) => !!cell.formula || !!cell.formulaType;
const derived = (cell: WorkbookCell) =>
  !!cell.formulaOutput || !!cell.spill || !!cell.scenarioResult || !!cell.calculation || !!cell.cachedResult;
function literalType(cell: WorkbookCell): keyof SpecialSelectionTypes | null {
  if (authoredFormula(cell) || derived(cell)) return null;
  if (cell.value === null) {
    if (cell.dataType !== undefined && cell.dataType !== 'blank')
      throw Error('Go To Special source type is incomplete.');
    return null;
  }
  if (cell.dataType === 'error') {
    if (typeof cell.value !== 'string' || !cell.value) throw Error('Go To Special error source is not qualified.');
    return 'error';
  }
  const type = typeof cell.value === 'number' ? 'number' : typeof cell.value === 'boolean' ? 'logical' : 'text',
    declared = type === 'logical' ? 'boolean' : type;
  if (cell.dataType !== undefined && cell.dataType !== declared)
    throw Error('Go To Special source type disagrees with its literal value.');
  if (cell.error !== undefined) throw Error('Go To Special source error authority is incomplete.');
  return type;
}
function formulaType(cell: WorkbookCell, revision: string | undefined): keyof SpecialSelectionTypes {
  const calculation = cell.calculation;
  if (
    !revision ||
    !calculation ||
    calculation.revision !== revision ||
    calculation.pending === true ||
    !['calculated', 'error'].includes(calculation.status)
  )
    throw Error(
      'Formula type filtering needs a current defined result; stored or pending results are not type authority.'
    );
  if (calculation.status === 'error') {
    if (!text(calculation.error) || !calculation.error) throw Error('The current formula error type is not qualified.');
    return 'error';
  }
  if (typeof calculation.value === 'number' && Number.isFinite(calculation.value)) return 'number';
  if (typeof calculation.value === 'boolean') return 'logical';
  if (typeof calculation.value === 'string' && text(calculation.value)) return 'text';
  throw Error('The current formula result type is unavailable.');
}
function visibility(value: SpecialSelectionVisibility | undefined, sheet: string) {
  if (!value || value.name !== sheet || !Array.isArray(value.widths) || value.geometryIssues?.length)
    throw Error('Complete source row and column visibility is unavailable.');
  const hidden = new Set<number>(),
    visible = new Set<number>();
  for (const [rows, target] of [
    [value.hiddenRows ?? [], hidden],
    [value.visibleRows ?? [], visible],
  ] as const) {
    if (!Array.isArray(rows)) throw Error('Source row visibility is unavailable.');
    for (const row of rows) {
      if (!Number.isSafeInteger(row) || row < 0 || row >= 1048576) throw Error('Source row visibility is invalid.');
      target.add(row);
    }
  }
  const hiddenDefault = value.sheetFormat?.rowsHiddenByDefault ?? false;
  if (typeof hiddenDefault !== 'boolean') throw Error('Source default row visibility is unavailable.');
  const zeroHeight = new Set<number>();
  for (const row of value.rowHeights ?? []) {
    if (
      !Number.isSafeInteger(row.row) ||
      row.row < 0 ||
      row.row >= 1048576 ||
      !Number.isFinite(row.height) ||
      row.height < 0 ||
      row.height > 409.5
    )
      throw Error('Source row height is invalid.');
    if (row.height === 0) zeroHeight.add(row.row);
  }
  const intervals = value.widths.map((interval) => {
    if (
      !Number.isSafeInteger(interval.min) ||
      !Number.isSafeInteger(interval.max) ||
      interval.min < 1 ||
      interval.max < interval.min ||
      interval.max > 16384 ||
      typeof interval.hidden !== 'boolean' ||
      !Number.isFinite(interval.width) ||
      interval.width < 0 ||
      interval.width > 255
    )
      throw Error('Source column visibility is invalid.');
    return { ...interval };
  });
  return (at: ReturnType<typeof address>) => {
    const interval = intervals.find((v) => at.column + 1 >= v.min && at.column + 1 <= v.max);
    return (
      !hidden.has(at.row) &&
      !zeroHeight.has(at.row) &&
      (!hiddenDefault || visible.has(at.row)) &&
      !interval?.hidden &&
      interval?.width !== 0
    );
  };
}
function ruleMembership(receipt: SpecialSelectionRuleReceipt | undefined, sheet: string) {
  if (
    !receipt ||
    receipt.complete !== true ||
    !receipt.definition ||
    receipt.definition.sheet !== sheet ||
    !Array.isArray(receipt.definition.rules) ||
    !Array.isArray(receipt.definition.protectedReasons)
  )
    throw Error('A complete authored rule-range inventory is required.');
  const definition = receipt.definition,
    protectedRanges = definition.protectedRanges ?? [];
  if (!Array.isArray(protectedRanges) || (definition.protectedReasons.length && !protectedRanges.length))
    throw Error('An opaque rule has no qualified source range; Go To Special cannot omit it.');
  const ranges = [...definition.rules.map((rule) => rule.range), ...protectedRanges];
  if (ranges.length > 10_000) throw Error('Rule-range inventory exceeds the Go To Special limit.');
  const parsed = ranges.flatMap((range) => {
    if (typeof range !== 'string' || !range.trim()) throw Error('An authored rule has no qualified source range.');
    return range
      .trim()
      .split(/\s+/)
      .map((part) => rectangle(part.replaceAll('$', '')));
  });
  return (at: ReturnType<typeof address>) => parsed.some((range) => inside(at, range));
}

/** Bounded complete reads of the original rectangle. This function never
 * submits workbook edits or updates selection; only a reviewed host Apply can. */
export async function collectSpecialSelectionPreview(
  input: CollectSpecialSelectionInput
): Promise<SpecialSelectionPreview> {
  if (
    !exact(
      input,
      ['binding', 'options', 'geometry', 'read', 'isCurrent'],
      ['signal', 'yieldWork', 'calculationRevision', 'visibility', 'rules', 'usedRange', 'usedRangeReceipt']
    ) ||
    typeof input.read !== 'function' ||
    typeof input.isCurrent !== 'function' ||
    (input.yieldWork !== undefined && typeof input.yieldWork !== 'function')
  )
    throw Error('Go To Special needs a complete original-owner reader and merge geometry.');
  const binding = captureHomeBinding(input.binding),
    options = normalizeSpecialSelectionOptions(input.options),
    read = input.read,
    isCurrent = input.isCurrent,
    signal = input.signal,
    yieldWork = input.yieldWork,
    calculationRevision = input.calculationRevision;
  if (calculationRevision !== undefined && (!text(calculationRevision, 4096) || !calculationRevision))
    throw Error('Go To Special calculation receipt is invalid.');
  if (!binding.ownerIdentity || binding.mode !== 'normal' || binding.specialSelection)
    throw Error('Go To Special needs one regular original-owner range in Normal view.');
  if (!SPECIAL_SELECTION_KINDS.find((k) => k.id === options.kind)?.supported)
    throw Error('This Go To Special family is not available yet.');
  const check = () => {
    if (signal?.aborted) {
      const error = new Error('Go To Special preview cancelled.');
      error.name = 'AbortError';
      throw error;
    }
    if (!isCurrent()) throw new HomeCommandRetired();
  };
  check();
  if (options.kind === 'current-region') {
    const geometry = freeze({ name: input.geometry.name, merges: [...input.geometry.merges] });
    mergeGeometry(geometry, binding.sheet);
    if (input.usedRange || input.usedRangeReceipt)
      throw Error('Current Region uses original source discovery, not UsedRange authority.');
    const found = await collectCurrentRegion({
      sheet: binding.sheet,
      anchor: binding.selectedCell,
      merges: geometry.merges,
      read,
      check,
      validateCell: sourceCell,
      ...(yieldWork ? { yieldWork } : {}),
    });
    check();
    return freeze({
      schema: 1 as const,
      binding,
      options,
      geometry,
      scope: { kind: 'current-region' as const, usedRange: null, receipt: null },
      activeCell: found.anchor,
      sheet: binding.sheet,
      range: found.range,
      positions: found.positions,
      rows: found.references.map((cell) => ({ cell, classification: 'Current region' })),
      references: found.references,
      readRanges: found.readRanges,
    });
  }
  const capturedGeometry = structuredClone(input.geometry),
    allMerges = mergeGeometry(capturedGeometry, binding.sheet),
    geometry = freeze({ name: capturedGeometry.name, merges: [...capturedGeometry.merges] }),
    resolved = selectionScope(binding, options, geometry, input.usedRange, input.usedRangeReceipt),
    scope = freeze(resolved.scope),
    bounds = resolved.bounds,
    range = areaReference(bounds),
    positions = bounds.rows * bounds.columns,
    merges = allMerges.filter((merge) => intersects(merge, bounds));
  if (options.kind === 'last-cell') {
    const references: string[] = [];
    for (let row = bounds.first.row; row <= bounds.last.row; row++)
      for (let column = bounds.first.column; column <= bounds.last.column; column++) references.push(ref(row, column));
    check();
    return freeze({
      schema: 1 as const,
      binding,
      options,
      geometry,
      scope,
      activeCell: references[0],
      sheet: binding.sheet,
      range,
      positions,
      rows: references.map((cell) => ({
        cell,
        classification: positions === 1 ? 'Worksheet last cell' : 'Worksheet last cell (merged area)',
      })),
      references,
      readRanges: [],
    });
  }
  const visible = options.kind === 'visible' ? visibility(structuredClone(input.visibility), binding.sheet) : null;
  const belongs = ['conditional-formats', 'data-validation'].includes(options.kind)
    ? ruleMembership(structuredClone(input.rules), binding.sheet)
    : null;
  const allTypes = Object.values(options.types).every(Boolean),
    rows: SpecialSelectionRow[] = [],
    readRanges = windows(bounds);
  for (const range of readRanges) {
    check();
    const result = await read(binding.sheet, range, signal);
    check();
    if (
      !exact(result, ['range', 'complete', 'cells']) ||
      result.range !== range ||
      result.complete !== true ||
      !plain(result.cells) ||
      bytes(result) > SPECIAL_SELECTION_LIMITS.replyBytes
    )
      throw Error('Go To Special owner returned an incomplete, foreign or oversized range.');
    const window = rectangle(range),
      cells: Record<string, WorkbookCell> = {};
    for (const [reference, value] of Object.entries(result.cells)) {
      if (!inside(address(reference), window))
        throw Error('Go To Special owner returned a cell outside its accepted range.');
      cells[reference] = sourceCell(value);
    }
    for (let row = window.first.row; row <= window.last.row; row++)
      for (let column = window.first.column; column <= window.last.column; column++) {
        const reference = ref(row, column),
          cell = cells[reference],
          at = { row, column };
        let classification: string | null = null;
        if (options.kind === 'visible') {
          if (visible!(at)) classification = 'Visible cell';
        } else if (options.kind === 'conditional-formats' || options.kind === 'data-validation') {
          if (belongs!(at))
            classification =
              options.kind === 'conditional-formats' ? 'Conditional-format range' : 'Data-validation range';
        } else if (options.kind === 'constants') {
          if (cell) {
            const type = literalType(cell);
            if (type && options.types[type])
              classification =
                type === 'logical' ? 'Logical constant' : type[0].toUpperCase() + type.slice(1) + ' constant';
          }
        } else if (options.kind === 'formulas') {
          if (cell && authoredFormula(cell)) {
            if (allTypes) classification = 'Authored formula';
            else {
              const type = formulaType(cell, calculationRevision);
              if (options.types[type])
                classification =
                  type === 'logical'
                    ? 'Logical formula result'
                    : type[0].toUpperCase() + type.slice(1) + ' formula result';
            }
          }
        } else if (options.kind === 'blanks') {
          if (!cell) classification = 'Absent source cell';
          // Canonical empty text remains a literal Text Constant. Only a null
          // blank record or an absent complete-source position is a Blank.
          else if (!authoredFormula(cell) && !derived(cell)) {
            if (cell.value === null && (cell.dataType === undefined || cell.dataType === 'blank'))
              classification = 'Authored blank cell';
            else literalType(cell);
          }
        }
        if (classification) rows.push({ cell: reference, classification });
      }
    check();
    if (yieldWork) await yieldWork();
    check();
  }
  const expands = ['constants', 'formulas', 'blanks', 'visible'].includes(options.kind),
    matched = expands ? mergedMatches(rows, merges) : rows;
  matched.sort((a, b) => order(a.cell) - order(b.cell));
  const activeCell = options.kind === 'blanks' ? blankActiveCell(rows, merges) : (matched[0]?.cell ?? null);
  check();
  return freeze({
    schema: 1 as const,
    binding,
    options,
    geometry,
    scope,
    activeCell,
    sheet: binding.sheet,
    range,
    positions,
    rows: matched,
    references: matched.map((row) => row.cell),
    readRanges,
  });
}

/** Copies/caller-built plans receive a full shape, ordering and scope audit. */
export function specialSelectionPreviewPage(
  preview: SpecialSelectionPreview,
  offset = 0,
  limit = 100
): SpecialSelectionPage {
  if (
    !exact(preview, [
      'schema',
      'binding',
      'options',
      'geometry',
      'scope',
      'activeCell',
      'sheet',
      'range',
      'positions',
      'rows',
      'references',
      'readRanges',
    ]) ||
    preview.schema !== 1 ||
    !Array.isArray(preview.rows) ||
    !Array.isArray(preview.references) ||
    !Array.isArray(preview.readRanges) ||
    !exact(preview.geometry, ['name', 'merges']) ||
    !exact(preview.scope, ['kind', 'usedRange', 'receipt']) ||
    !Number.isSafeInteger(offset) ||
    offset < 0 ||
    offset > preview.rows.length ||
    !Number.isSafeInteger(limit) ||
    limit < 1 ||
    limit > SPECIAL_SELECTION_LIMITS.page
  )
    throw Error('Invalid Go To Special preview.');
  if (preview.options.kind === 'current-region') {
    const binding = captureHomeBinding(preview.binding),
      bounds = rectangle(preview.range);
    normalizeSpecialSelectionOptions(preview.options);
    const merges = mergeGeometry(preview.geometry, binding.sheet);
    if (
      !binding.ownerIdentity ||
      binding.mode !== 'normal' ||
      binding.specialSelection ||
      preview.sheet !== binding.sheet ||
      preview.scope.kind !== 'current-region' ||
      preview.scope.usedRange !== null ||
      preview.scope.receipt !== null ||
      preview.activeCell !== binding.selectedCell ||
      !inside(address(binding.selectedCell), bounds) ||
      preview.positions !== bounds.rows * bounds.columns ||
      preview.positions > SPECIAL_SELECTION_LIMITS.positions ||
      preview.references.length !== preview.positions ||
      preview.rows.length !== preview.positions ||
      merges.some((merge) => intersects(bounds, merge)) ||
      preview.readRanges.length > 60_000
    )
      throw Error('Invalid Current Region preview.');
    let i = 0;
    for (let r = bounds.first.row; r <= bounds.last.row; r++)
      for (let c = bounds.first.column; c <= bounds.last.column; c++, i++) {
        const cell = ref(r, c);
        if (
          preview.references[i] !== cell ||
          !exact(preview.rows[i], ['cell', 'classification']) ||
          preview.rows[i].cell !== cell ||
          preview.rows[i].classification !== 'Current region'
        )
          throw Error('Current Region must keep its complete regular rectangle.');
      }
    let readPositions = 0;
    for (const target of preview.readRanges) {
      const window = rectangle(target);
      if (window.rows * window.columns > 1000) throw Error('Oversized Current Region read window.');
      readPositions += window.rows * window.columns;
    }
    if (readPositions > 60_000) throw Error('Current Region source discovery exceeded its limit.');
    const end = Math.min(offset + limit, preview.rows.length);
    return freeze({
      offset,
      total: preview.rows.length,
      rows: structuredClone(preview.rows.slice(offset, end)),
      nextOffset: end < preview.rows.length ? end : null,
    });
  }
  const binding = captureHomeBinding(preview.binding),
    options = normalizeSpecialSelectionOptions(preview.options),
    resolved = selectionScope(binding, options, preview.geometry, preview.scope.usedRange, preview.scope.receipt),
    bounds = resolved.bounds,
    last = options.kind === 'last-cell';
  if (
    !binding.ownerIdentity ||
    binding.mode !== 'normal' ||
    binding.specialSelection ||
    preview.scope.kind !== resolved.scope.kind ||
    preview.sheet !== binding.sheet ||
    preview.range !== areaReference(bounds) ||
    preview.positions !== bounds.rows * bounds.columns ||
    preview.positions > SPECIAL_SELECTION_LIMITS.positions ||
    !SPECIAL_SELECTION_KINDS.find((k) => k.id === options.kind)?.supported ||
    preview.rows.length > preview.positions ||
    preview.references.length !== preview.rows.length ||
    JSON.stringify(preview.readRanges) !== JSON.stringify(last ? [] : windows(bounds)) ||
    (preview.rows.length
      ? typeof preview.activeCell !== 'string' || !preview.references.includes(preview.activeCell)
      : preview.activeCell !== null) ||
    (last &&
      (preview.rows.length !== preview.positions || preview.activeCell !== ref(bounds.first.row, bounds.first.column)))
  )
    throw Error('Go To Special preview does not match its complete captured range.');
  let prior = -1;
  for (let index = 0; index < preview.rows.length; index++) {
    const row = preview.rows[index];
    if (
      !exact(row, ['cell', 'classification']) ||
      typeof row.cell !== 'string' ||
      !text(row.classification, 120) ||
      !row.classification ||
      !inside(address(row.cell), bounds) ||
      order(row.cell) <= prior ||
      preview.references[index] !== row.cell
    )
      throw Error('Invalid Go To Special target order or range.');
    prior = order(row.cell);
  }
  const end = Math.min(offset + limit, preview.rows.length);
  return freeze({
    offset,
    total: preview.rows.length,
    rows: structuredClone(preview.rows.slice(offset, end)),
    nextOffset: end < preview.rows.length ? end : null,
  });
}
