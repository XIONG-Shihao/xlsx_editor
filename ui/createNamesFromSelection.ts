/** Readonly label planning; only an explicit acknowledged batch creates names. */
import type { WorkbookCell, WorkbookCommand, WorkbookView } from '../core/contracts';
import { validateDefinedName, definedNameConflict } from '../core/defined-name.mjs';
export type NameEdge = 'top' | 'left' | 'bottom' | 'right';
export type NameEdges = Record<NameEdge, boolean>;
export type NameInventory = Required<Pick<WorkbookView, 'namedRanges' | 'tables'>>;
export interface CreateNamesBinding {
  documentId: string;
  ownerIdentity: string;
  state: string;
  sheet: string;
  selection: string;
  navigationGeneration: number;
  activationEpoch: number;
  modeEpoch: number;
}
export interface CreateNamesContext {
  binding: CreateNamesBinding;
  isCurrent: (binding: CreateNamesBinding) => boolean;
  isRuntimeCurrent: (binding: CreateNamesBinding) => boolean;
  availability: () => {
    active: boolean;
    readOnly: boolean;
    busy: boolean;
    literal: boolean;
    specialSelection: boolean;
  };
  read: (range: string) => Promise<Record<string, WorkbookCell>>;
  inventory: () => Promise<NameInventory>;
  merges: readonly string[];
  accept: (commands: WorkbookCommand[]) => Promise<void>;
}
export interface NameProposal {
  id: string;
  edge: NameEdge;
  labelCell: string;
  label: string;
  name: string;
  range: string;
  formula: string;
  sourceReason?: string;
  converted: boolean;
}
export interface NameChoice {
  name: string;
  included: boolean;
}
export interface CreateNamesPreview {
  binding: CreateNamesBinding;
  edges: NameEdges;
  dataRange: string;
  proposals: readonly NameProposal[];
  corners: readonly string[];
  inventory: NameInventory;
  fingerprint: string;
}
export const CREATE_NAMES_LIMITS = Object.freeze({
  positions: 10000,
  labels: 1000,
  readPositions: 1000,
  bytes: 1048576,
  names: 1000,
});
export class CreateNamesRetired extends Error {
  constructor() {
    super('The workbook or selection changed. Reopen Create from Selection.');
    this.name = 'CreateNamesRetired';
  }
}
export class CreateNamesCancelled extends Error {
  constructor() {
    super('Name creation cancelled before acceptance.');
    this.name = 'CreateNamesCancelled';
  }
}
export const createNamesBindingKey = (value: CreateNamesBinding): string =>
  JSON.stringify([
    value.documentId,
    value.ownerIdentity,
    value.state,
    value.sheet,
    value.selection,
    value.navigationGeneration,
    value.activationEpoch,
    value.modeEpoch,
  ]);
const bytes = (value: unknown) => new TextEncoder().encode(JSON.stringify(value)).length;
function position(ref: string) {
  const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(ref);
  if (!match) throw Error('Select a canonical rectangular range.');
  const col = [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1,
    row = Number(match[2]) - 1;
  if (col > 16383 || row > 1048575) throw Error('The selection exceeds worksheet limits.');
  return { row, col };
}
function column(index: number) {
  let value = '';
  for (let n = index + 1; n; n = Math.floor((n - 1) / 26)) value = String.fromCharCode(65 + ((n - 1) % 26)) + value;
  return value;
}
const reference = (row: number, col: number) => column(col) + (row + 1);
function bounds(range: string) {
  const parts = range.split(':');
  if (parts.length > 2) throw Error('Select one rectangular range.');
  const first = position(parts[0]),
    last = position(parts[1] ?? parts[0]);
  if (first.row > last.row || first.col > last.col) throw Error('Select one rectangular range.');
  return { first, last, positions: (last.row - first.row + 1) * (last.col - first.col + 1) };
}
function edges(input: NameEdges): NameEdges {
  if (
    !input ||
    Object.keys(input).length !== 4 ||
    (['top', 'left', 'bottom', 'right'] as const).some((key) => typeof input[key] !== 'boolean')
  )
    throw Error('Choose label locations.');
  if (!Object.values(input).some(Boolean)) throw Error('Choose at least one label location.');
  return { ...input };
}
function layout(selection: string, input: NameEdges) {
  const area = bounds(selection),
    chosen = edges(input);
  if (area.positions > CREATE_NAMES_LIMITS.positions) throw Error('Select at most 10,000 cells, including the labels.');
  const first = { row: area.first.row + Number(chosen.top), col: area.first.col + Number(chosen.left) },
    last = { row: area.last.row - Number(chosen.bottom), col: area.last.col - Number(chosen.right) };
  if (first.row > last.row || first.col > last.col)
    throw Error('Select labels and at least one data cell after excluding the chosen edges.');
  const labelCount =
    (last.col - first.col + 1) * (Number(chosen.top) + Number(chosen.bottom)) +
    (last.row - first.row + 1) * (Number(chosen.left) + Number(chosen.right));
  if (labelCount > CREATE_NAMES_LIMITS.labels) throw Error('This selection contains more than 1,000 label proposals.');
  return {
    area,
    chosen,
    first,
    last,
    dataRange: reference(first.row, first.col) + ':' + reference(last.row, last.col),
  };
}
export function proposedRangeName(label: string): string {
  return label.trim().replace(/ +/g, '_');
}
function label(cell: WorkbookCell | undefined): { text: string; reason?: string } {
  if (!cell) return { text: '' };
  if (cell.formulaType || cell.spill || cell.scenarioResult || cell.calculation?.pending)
    return {
      text: cell.rawValue ?? '',
      reason: 'This label is a pending or protected result. Use ordinary text or calculate it first.',
    };
  if (cell.formula) {
    if (cell.calculation?.status === 'calculated' && typeof cell.calculation.value === 'string')
      return { text: cell.calculation.value };
    return { text: cell.rawValue ?? '', reason: 'This formula label does not have a current text result.' };
  }
  if (cell.value === null || cell.value === undefined) return { text: '' };
  if (typeof cell.value === 'string' && (!cell.dataType || cell.dataType === 'text') && !cell.error)
    return { text: cell.value };
  return {
    text: cell.rawValue ?? String(cell.value),
    reason: 'This label is not ordinary text. Exclude it or use a text label.',
  };
}
function checkedInventory(value: NameInventory): NameInventory {
  if (
    !value ||
    !Array.isArray(value.namedRanges) ||
    !Array.isArray(value.tables) ||
    value.namedRanges.length > 1000 ||
    value.tables.length > 4096 ||
    value.namedRanges.some((item) => !item || typeof item.name !== 'string' || typeof item.formula !== 'string') ||
    value.tables.some((item) => !item || typeof item.name !== 'string') ||
    bytes(value) > CREATE_NAMES_LIMITS.bytes
  )
    throw Error('The accepted name inventory is unavailable or exceeds its supported size.');
  return structuredClone(value);
}
export function planNamesFromSelection(
  binding: CreateNamesBinding,
  input: NameEdges,
  cells: Record<string, WorkbookCell>,
  inventory: NameInventory,
  merges: readonly string[] = []
): CreateNamesPreview {
  const shape = layout(binding.selection, input),
    items: NameProposal[] = [],
    corners: string[] = [];
  const add = (edge: NameEdge, row: number, col: number, range: string) => {
    const labelCell = reference(row, col),
      value = label(cells[labelCell]),
      name = proposedRangeName(value.text);
    let sourceReason = value.reason;
    if (
      merges.some((range) => {
        const area = bounds(range);
        return row >= area.first.row && row <= area.last.row && col >= area.first.col && col <= area.last.col;
      })
    )
      sourceReason = 'Choose unmerged label cells.';
    const [a, b = a] = range.split(':'),
      absolute = (ref: string) => {
        const p = position(ref);
        return '$' + column(p.col) + '$' + (p.row + 1);
      };
    items.push(
      Object.freeze({
        id: edge + ':' + labelCell,
        edge,
        labelCell,
        label: value.text,
        name,
        range,
        formula: "'" + binding.sheet.replaceAll("'", "''") + "'!" + absolute(a) + ':' + absolute(b),
        sourceReason,
        converted: name !== value.text,
      })
    );
  };
  for (let col = shape.first.col; col <= shape.last.col; col++) {
    const range = reference(shape.first.row, col) + ':' + reference(shape.last.row, col);
    if (shape.chosen.top) add('top', shape.area.first.row, col, range);
    if (shape.chosen.bottom) add('bottom', shape.area.last.row, col, range);
  }
  for (let row = shape.first.row; row <= shape.last.row; row++) {
    const range = reference(row, shape.first.col) + ':' + reference(row, shape.last.col);
    if (shape.chosen.left) add('left', row, shape.area.first.col, range);
    if (shape.chosen.right) add('right', row, shape.area.last.col, range);
  }
  for (const [row, rowEnabled] of [
    [shape.area.first.row, shape.chosen.top],
    [shape.area.last.row, shape.chosen.bottom],
  ] as const)
    for (const [col, colEnabled] of [
      [shape.area.first.col, shape.chosen.left],
      [shape.area.last.col, shape.chosen.right],
    ] as const)
      if (rowEnabled && colEnabled) corners.push(reference(row, col));
  if (items.length > CREATE_NAMES_LIMITS.labels)
    throw Error('This selection contains more than 1,000 label proposals.');
  const acceptedInventory = checkedInventory(inventory),
    fingerprint = JSON.stringify([
      createNamesBindingKey(binding),
      shape.chosen,
      items,
      Object.entries(cells).sort(([a], [b]) => a.localeCompare(b)),
      acceptedInventory,
      merges,
    ]);
  if (bytes(fingerprint) > CREATE_NAMES_LIMITS.bytes) throw Error('The name preview exceeds its supported size.');
  return Object.freeze({
    binding: Object.freeze({ ...binding }),
    edges: Object.freeze(shape.chosen),
    dataRange: shape.dataRange,
    proposals: Object.freeze(items),
    corners: Object.freeze(corners),
    inventory: acceptedInventory,
    fingerprint,
  });
}
export function createNameChoiceReason(
  preview: CreateNamesPreview,
  proposal: NameProposal,
  choices: Record<string, NameChoice>
): string | null {
  const choice = choices[proposal.id] ?? { name: proposal.name, included: true };
  if (typeof choice.name !== 'string' || typeof choice.included !== 'boolean')
    throw Error('Invalid proposed name choice.');
  if (!choice.included) return null;
  if (proposal.sourceReason) return proposal.sourceReason;
  try {
    validateDefinedName(choice.name);
  } catch (error) {
    return error instanceof Error ? error.message : String(error);
  }
  const conflict = definedNameConflict(choice.name, preview.inventory, { visibleSheet: preview.binding.sheet });
  if (conflict) return conflict;
  const folded = choice.name.toUpperCase();
  if (
    preview.proposals.some(
      (other) =>
        other.id !== proposal.id &&
        (choices[other.id]?.included ?? true) &&
        (choices[other.id]?.name ?? other.name).toUpperCase() === folded
    )
  )
    return 'Another selected label proposes this name. Rename or exclude one.';
  return null;
}
function check(context: CreateNamesContext, signal?: AbortSignal) {
  if (signal?.aborted) throw new CreateNamesCancelled();
  if (!context.isCurrent(context.binding)) throw new CreateNamesRetired();
  const access = context.availability();
  if (!access.active) throw new CreateNamesRetired();
  if (access.readOnly || access.literal) throw Error('This command requires an editable XLSX workbook.');
  if (access.specialSelection) throw Error('Click a regular cell to leave the special selection.');
  if (access.busy) throw Error('Wait for the workbook operation to finish.');
}
async function readLabels(context: CreateNamesContext, input: NameEdges, signal?: AbortSignal) {
  const shape = layout(context.binding.selection, input),
    cells: Record<string, WorkbookCell> = {},
    windows: string[] = [];
  const band = (firstRow: number, lastRow: number, firstCol: number, lastCol: number) => {
    for (let col = firstCol; col <= lastCol; col += 1000) {
      const width = Math.min(1000, lastCol - col + 1),
        step = Math.floor(1000 / width);
      for (let row = firstRow; row <= lastRow; row += step)
        windows.push(reference(row, col) + ':' + reference(Math.min(lastRow, row + step - 1), col + width - 1));
    }
  };
  if (shape.chosen.top) band(shape.area.first.row, shape.area.first.row, shape.first.col, shape.last.col);
  if (shape.chosen.bottom) band(shape.area.last.row, shape.area.last.row, shape.first.col, shape.last.col);
  if (shape.chosen.left) band(shape.first.row, shape.last.row, shape.area.first.col, shape.area.first.col);
  if (shape.chosen.right) band(shape.first.row, shape.last.row, shape.area.last.col, shape.area.last.col);
  for (const range of windows) {
    check(context, signal);
    const reply = await context.read(range);
    check(context, signal);
    const window = bounds(range);
    if (!reply || typeof reply !== 'object' || Array.isArray(reply))
      throw Error('The label read returned an invalid rectangle.');
    for (const [ref, cell] of Object.entries(reply)) {
      const p = position(ref);
      if (
        p.row < window.first.row ||
        p.row > window.last.row ||
        p.col < window.first.col ||
        p.col > window.last.col ||
        !cell ||
        typeof cell !== 'object' ||
        Array.isArray(cell)
      )
        throw Error('The label read returned a foreign or malformed cell.');
      cells[ref] = structuredClone(cell);
    }
    if (bytes(cells) > CREATE_NAMES_LIMITS.bytes) throw Error('The label preview exceeds its supported size.');
  }
  return cells;
}
export async function previewCreateNames(
  context: CreateNamesContext,
  input: NameEdges,
  { signal }: { signal?: AbortSignal } = {}
): Promise<CreateNamesPreview> {
  check(context, signal);
  const chosen = edges(input),
    cells = await readLabels(context, chosen, signal);
  check(context, signal);
  const inventory = await context.inventory();
  check(context, signal);
  return planNamesFromSelection(context.binding, chosen, cells, inventory, context.merges);
}
export async function acceptCreateNames(
  context: CreateNamesContext,
  preview: CreateNamesPreview,
  choices: Record<string, NameChoice>,
  { signal, beforeAcceptance }: { signal?: AbortSignal; beforeAcceptance?: () => void } = {}
) {
  check(context, signal);
  if (createNamesBindingKey(context.binding) !== createNamesBindingKey(preview.binding)) throw new CreateNamesRetired();
  const fresh = await previewCreateNames(context, preview.edges, { signal });
  check(context, signal);
  if (fresh.fingerprint !== preview.fingerprint)
    throw Error('The labels or existing names changed. Preview names again.');
  const commands: Array<{
    operation: 'defineName';
    sheet: string;
    name: string;
    range: string;
    nameScope: 'workbook';
    expectedNameFormula: null;
  }> = [];
  for (const proposal of fresh.proposals) {
    const choice = choices[proposal.id] ?? { name: proposal.name, included: true };
    if (typeof choice.name !== 'string' || typeof choice.included !== 'boolean')
      throw Error('Invalid proposed name choice.');
    if (!choice.included) continue;
    const reason = createNameChoiceReason(fresh, proposal, choices);
    if (reason) throw Error(reason);
    commands.push({
      operation: 'defineName',
      sheet: context.binding.sheet,
      name: choice.name,
      range: proposal.range,
      nameScope: 'workbook',
      expectedNameFormula: null,
    });
  }
  if (!commands.length) throw Error('Include at least one valid name.');
  if (commands.length + fresh.inventory.namedRanges.length > CREATE_NAMES_LIMITS.names)
    throw Error('Creating these names would exceed the 1,000-name limit.');
  if (bytes(commands) > CREATE_NAMES_LIMITS.bytes) throw Error('The name batch exceeds its supported size.');
  check(context, signal);
  beforeAcceptance?.();
  check(context, signal);
  await context.accept(commands);
  return {
    outcome: 'accepted' as const,
    count: commands.length,
    names: commands.map((command) => command.name!),
    binding: context.binding,
    current: context.isCurrent(context.binding),
    runtimeCurrent: context.isRuntimeCurrent(context.binding),
  };
}
