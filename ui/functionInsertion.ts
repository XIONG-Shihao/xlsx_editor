/** One closed function call, one scoped target and one owner batch. */
import type { WorkbookCell, WorkbookCommand } from '../core/contracts';
import { serializeFunction, type FunctionArgument, type FunctionId } from './functionCatalog';

export interface FunctionBinding {
  readonly documentId: string;
  readonly grantDocumentId: string;
  readonly ownerKey: string;
  readonly ownerIdentity?: string;
  readonly state: string;
  readonly sheet: string;
  readonly target: string;
  readonly selection: string;
  readonly navigationGeneration: number;
  readonly activationEpoch: number;
}
export interface FunctionTarget {
  cell?: WorkbookCell;
  protectedReason?: string;
}
export interface FunctionInsertionContext {
  readonly binding: FunctionBinding;
  readonly isCurrent: (binding: FunctionBinding) => boolean;
  readonly isRuntimeCurrent?: (binding: FunctionBinding) => boolean;
  readonly currentState?: () => string;
  readonly readAcceptedTarget?: (formula: string) => Promise<{ state: string; cell: WorkbookCell }>;
  readonly availability: () => {
    active: boolean;
    readOnly: boolean;
    busy: boolean;
    literal: boolean;
    specialSelection: boolean;
  };
  /** Captured at pane opening; reads exactly one position at the captured owner. */
  readonly readTarget: () => Promise<FunctionTarget>;
  /** Captured strict callback; success means owner acceptance, not queueing. */
  readonly accept?: (commands: WorkbookCommand[]) => Promise<void>;
}
export class FunctionInsertionRetired extends Error {
  readonly code = 'FUNCTION_INSERTION_RETIRED';
  constructor() {
    super('The initiating workbook or target changed. Reopen Insert Function.');
    this.name = 'FunctionInsertionRetired';
  }
}
export class FunctionInsertionCancelled extends Error {
  readonly code = 'FUNCTION_INSERTION_CANCELLED';
  constructor() {
    super('Insertion cancelled before the workbook command.');
    this.name = 'FunctionInsertionCancelled';
  }
}
export function functionBindingKey(value: FunctionBinding): string {
  return JSON.stringify([
    value.documentId,
    value.grantDocumentId,
    value.ownerKey,
    value.ownerIdentity,
    value.state,
    value.sheet,
    value.target,
    value.selection,
    value.navigationGeneration,
    value.activationEpoch,
  ]);
}
function check(context: FunctionInsertionContext, signal?: AbortSignal, editable = false): void {
  if (signal?.aborted) throw new FunctionInsertionCancelled();
  if (!context.isCurrent(context.binding)) throw new FunctionInsertionRetired();
  const scope = context.availability();
  if (!scope.active) throw new FunctionInsertionRetired();
  if (scope.specialSelection) throw Error('Click a cell to leave the special selection before inserting a function.');
  if (scope.literal) throw Error('This command requires an editable XLSX workbook.');
  if (editable) {
    if (scope.readOnly) throw Error('This workbook is read-only.');
    if (scope.busy) throw Error('Wait for the workbook operation to finish.');
    if (!context.accept) throw Error('Workbook commands are unavailable in this editor.');
  }
}
export function targetOccupied(target: FunctionTarget): boolean {
  return (
    !!target.cell?.formula ||
    !!target.cell?.formulaType ||
    (target.cell?.value !== undefined && target.cell.value !== null)
  );
}
export function targetFingerprint(target: FunctionTarget): string {
  const cell = target.cell;
  return JSON.stringify([
    target.protectedReason ?? null,
    cell?.formula ?? null,
    cell?.formulaType ?? null,
    cell?.spill ?? null,
    cell?.formulaOutput ?? null,
    cell?.scenarioResult ?? null,
    cell?.dataType ?? null,
    cell?.value ?? null,
    cell?.rawValue ?? null,
  ]);
}
export function targetDescription(target: FunctionTarget): string {
  const cell = target.cell;
  if (cell?.formula) return 'Formula: =' + cell.formula;
  if (cell?.dataType === 'error' || cell?.error) return 'Error: ' + (cell.error ?? String(cell.value));
  if (cell?.value === undefined || cell.value === null) return 'Empty cell';
  if (typeof cell.value === 'string') return 'Text: ' + JSON.stringify(cell.value);
  return (
    (typeof cell.value === 'boolean' ? 'Boolean: ' : 'Number: ') +
    (typeof cell.value === 'number' && cell.rawValue ? cell.rawValue : String(cell.value))
  );
}
export function targetProtection(target: FunctionTarget): string | null {
  if (target.protectedReason) return target.protectedReason;
  if (target.cell?.spill) return 'This cell belongs to an array output. Choose a regular cell.';
  if (target.cell?.formulaOutput) return 'This cell belongs to a formula output. Choose a regular cell.';
  if (target.cell?.formulaType || target.cell?.scenarioResult)
    return 'This cell has a protected formula. Choose a regular cell.';
  return null;
}
export async function inspectFunctionTarget(
  context: FunctionInsertionContext,
  signal?: AbortSignal
): Promise<FunctionTarget> {
  check(context, signal);
  const target = await context.readTarget();
  check(context, signal);
  return target;
}
export interface FunctionInsertionResult {
  outcome: 'accepted';
  binding: FunctionBinding;
  functionId: FunctionId;
  formula: string;
  current: boolean;
}
export async function insertFunction(
  context: FunctionInsertionContext,
  input: { id: FunctionId; args: readonly FunctionArgument[]; replaceFingerprint?: string },
  options: { signal?: AbortSignal; beforeAcceptance?: () => void } = {}
): Promise<FunctionInsertionResult> {
  check(context, options.signal, true);
  const serialized = serializeFunction(input.id, input.args);
  const target = await inspectFunctionTarget(context, options.signal);
  check(context, options.signal, true);
  const reason = targetProtection(target);
  if (reason) throw Error(reason);
  if (targetOccupied(target) && input.replaceFingerprint !== targetFingerprint(target))
    throw Error('Review the existing cell and confirm its replacement.');
  options.beforeAcceptance?.();
  // There is no UI cancellation after this boundary: the owner decides whether
  // this single command was durably accepted, even if the view later moves on.
  await context.accept!([
    {
      type: 'setCells',
      sheet: context.binding.sheet,
      cells: [{ cell: context.binding.target, formula: serialized.formula }],
    },
  ]);
  return {
    outcome: 'accepted',
    binding: context.binding,
    functionId: serialized.functionId,
    formula: serialized.formula,
    current: context.isCurrent(context.binding),
  };
}
