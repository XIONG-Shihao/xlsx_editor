import type { WorkbookCalculationPolicy } from '../core/contracts';
import type {
  CalculationAction,
  CalculationBinding,
  CalculationProgress,
  CalculationGroupReceipt,
} from '../adapters/workbook-calculation-group.mjs';
export type { CalculationAction, CalculationBinding, CalculationProgress, CalculationGroupReceipt };
export interface WorkbookCalculationControls {
  receipt: CalculationBinding;
  policy: WorkbookCalculationPolicy;
  optionsScope: 'workbook' | 'application';
  applicationMode?: { mode: 'auto' | 'manual'; revision: number };
  run: (
    action: CalculationAction,
    requestId: string,
    onProgress: (value: CalculationProgress) => void
  ) => Promise<CalculationGroupReceipt>;
  cancel: (requestId: string) => Promise<{ outcome: 'cancelled' | 'completed' | 'finishing' }>;
}
/** UI state is an opaque acknowledged token, independent of the trusted owner
 * receipt. Neither state is decoded or derived from the other. */
export interface CalculationUIScope {
  documentId: string;
  ownerIdentity: string;
  state: string;
  sheet: string;
  activationEpoch: number;
  modeEpoch: number;
  navigationGeneration: number;
  active: boolean;
  normal: boolean;
  readOnly: boolean;
  busy: boolean;
  profile: 'xlsx' | 'delimited' | 'passive';
}
export interface CalculationFrame {
  scope: CalculationUIScope;
  controls: WorkbookCalculationControls;
}
export type CalculationJobPhase =
  | 'idle'
  | 'running'
  | 'stopping'
  | 'finishing'
  | 'completed'
  | 'partial'
  | 'cancelled'
  | 'failed';
export interface CalculationJobSnapshot {
  phase: CalculationJobPhase;
  physicalBusy: boolean;
  runtimeKey: string | null;
  requestId: string | null;
  action: CalculationAction | null;
  progress: CalculationProgress | null;
  receipt: CalculationGroupReceipt | null;
  error: string | null;
  canCancel: boolean;
}
export class CalculationJobRetired extends Error {
  constructor(cause?: unknown) {
    super('The initiating calculation view changed.', { cause });
  }
}
const object = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);
const text = (value: unknown): value is string => typeof value === 'string' && value.length > 0 && value.length <= 4096;
function binding(value: unknown): value is CalculationBinding {
  return object(value) && ['documentId', 'ownerId', 'state', 'revision'].every((key) => text(value[key]));
}
const bindingKey = (value: CalculationBinding) =>
  JSON.stringify([value.documentId, value.ownerId, value.state, value.revision]);
export const calculationRuntimeKey = (scope: CalculationUIScope): string =>
  JSON.stringify([scope.documentId, scope.ownerIdentity, scope.activationEpoch, scope.modeEpoch]);
const runtimeFrameKey = (frame: CalculationFrame): string =>
  JSON.stringify([
    calculationRuntimeKey(frame.scope),
    frame.controls.receipt.documentId,
    frame.controls.receipt.ownerId,
  ]);
export const calculationFrameKey = (frame: CalculationFrame): string =>
  JSON.stringify([
    runtimeFrameKey(frame),
    frame.scope.sheet,
    frame.scope.navigationGeneration,
    frame.scope.state,
    bindingKey(frame.controls.receipt),
  ]);
function actionKey(value: CalculationAction): string {
  if (value.type === 'calculateAllOpen') return JSON.stringify([value.type]);
  if (value.type === 'calculateSheet') return JSON.stringify([value.type, value.sheet]);
  return JSON.stringify([value.type, value.mode, value.recalculateBeforeSave ?? null]);
}
export function calculationPendingObjectCount(policy: WorkbookCalculationControls['policy']): number | undefined {
  const count = policy.pendingCacheObjects;
  return Number.isSafeInteger(count) && Number(count) >= 0 ? count : undefined;
}
export function calculationDisabledReason(
  frame: CalculationFrame | null,
  action: CalculationAction
): string | undefined {
  if (
    !frame ||
    !binding(frame.controls.receipt) ||
    !text(frame.scope.ownerIdentity) ||
    !frame.scope.ownerIdentity.trim()
  )
    return 'Calculation is unavailable in this editor.';
  const { scope, controls } = frame;
  if (scope.profile !== 'xlsx') return 'This command requires an editable XLSX workbook.';
  if (!scope.active) return 'Activate this workbook first.';
  if (!scope.normal) return 'Return to Normal view before calculating.';
  if (scope.readOnly) return 'This workbook is read-only.';
  if (scope.busy) return 'Wait for the workbook operation to finish.';
  if (
    controls.policy.schema !== 1 ||
    !['auto', 'manual'].includes(controls.policy.mode) ||
    typeof controls.policy.recalculateBeforeSave !== 'boolean' ||
    !Number.isSafeInteger(controls.policy.pendingFormulaCells) ||
    controls.policy.pendingFormulaCells < 0 ||
    typeof controls.policy.canCalculateSheet !== 'boolean' ||
    typeof controls.policy.canChangeOptions !== 'boolean' ||
    (controls.policy.canRecalculate !== undefined && typeof controls.policy.canRecalculate !== 'boolean') ||
    (controls.policy.protectedSource !== undefined && typeof controls.policy.protectedSource !== 'boolean') ||
    (controls.policy.pendingCacheObjects !== undefined &&
      (!Number.isSafeInteger(controls.policy.pendingCacheObjects) || controls.policy.pendingCacheObjects < 0))
  )
    return 'Calculation settings are unavailable.';
  if (action.type === 'calculateAllOpen' && controls.policy.canRecalculate === false)
    return controls.policy.reason ?? 'Calculation is protected for this workbook.';
  if (action.type === 'setOptions' && !controls.policy.canChangeOptions)
    return controls.policy.reason ?? 'Calculation options are protected for this workbook.';
  if (action.type === 'calculateSheet' && (!controls.policy.canCalculateSheet || action.sheet !== scope.sheet))
    return controls.policy.reason ?? 'Calculation of this worksheet is unavailable.';
}
function validateAction(value: CalculationAction): CalculationAction {
  if (!object(value)) throw Error('Invalid calculation action.');
  if (value.type === 'calculateAllOpen' && Object.keys(value).length === 1) return Object.freeze({ type: value.type });
  if (
    value.type === 'calculateSheet' &&
    text(value.sheet) &&
    Object.keys(value).every((k) => ['type', 'sheet'].includes(k))
  )
    return Object.freeze({ type: value.type, sheet: value.sheet });
  if (
    value.type === 'setOptions' &&
    ['auto', 'manual'].includes(value.mode as string) &&
    Object.keys(value).every((k) => ['type', 'mode', 'recalculateBeforeSave'].includes(k)) &&
    (value.recalculateBeforeSave === undefined || typeof value.recalculateBeforeSave === 'boolean')
  )
    return Object.freeze({ ...value }) as CalculationAction;
  throw Error('Invalid calculation action.');
}
export function validateCalculationReceipt(
  value: unknown,
  captured: { requestId: string; receipt: CalculationBinding; action: CalculationAction }
): CalculationGroupReceipt {
  if (
    !object(value) ||
    value.schema !== 1 ||
    value.requestId !== captured.requestId ||
    !binding(value.initiator) ||
    bindingKey(value.initiator) !== bindingKey(captured.receipt) ||
    value.consistency !== 'independent-workbook-commits' ||
    !['completed', 'partial', 'cancelled'].includes(value.outcome as string) ||
    !Array.isArray(value.results) ||
    value.results.length < 1 ||
    value.results.length > 128
  )
    throw Error('The calculation reply did not match the initiating workbook.');
  if (
    !object(value.action) ||
    actionKey(validateAction(value.action as unknown as CalculationAction)) !== actionKey(captured.action)
  )
    throw Error('The calculation reply did not match the requested action.');
  const seen = new Set<string>();
  for (const row of value.results) {
    if (
      !object(row) ||
      !['documentId', 'ownerId', 'name', 'beforeState'].every((key) => text(row[key])) ||
      !['accepted', 'skipped', 'refused', 'failed', 'cancelled', 'notRun'].includes(row.status as string) ||
      (row.reason !== undefined && typeof row.reason !== 'string') ||
      (row.afterState !== undefined && !text(row.afterState)) ||
      (row.status === 'accepted' && !text(row.afterState)) ||
      (row.status !== 'accepted' && row.afterState !== undefined)
    )
      throw Error('The calculation reply contained an invalid workbook result.');
    const key = JSON.stringify([row.documentId, row.ownerId]);
    if (seen.has(key)) throw Error('The calculation reply repeated a workbook result.');
    seen.add(key);
  }
  const first = value.results[0];
  if (
    first.documentId !== captured.receipt.documentId ||
    first.ownerId !== captured.receipt.ownerId ||
    first.beforeState !== captured.receipt.state ||
    (captured.action.type === 'calculateSheet' && value.results.length !== 1)
  )
    throw Error('The calculation reply did not retain its initiating owner.');
  if (
    value.outcome === 'completed' &&
    value.results.some((row) => ['refused', 'failed', 'cancelled', 'notRun'].includes(row.status))
  )
    throw Error('The calculation reply hid an unfinished workbook result.');
  if (
    value.applicationMode !== undefined &&
    (!object(value.applicationMode) ||
      !['auto', 'manual'].includes(value.applicationMode.mode as string) ||
      !Number.isSafeInteger(value.applicationMode.revision) ||
      Number(value.applicationMode.revision) < 0)
  )
    throw Error('The calculation reply contained an invalid application mode.');
  // Keep a private copy: later host mutations cannot change the accepted summary.
  return JSON.parse(JSON.stringify(value)) as CalculationGroupReceipt;
}
const idle = (): CalculationJobSnapshot =>
  Object.freeze({
    phase: 'idle',
    physicalBusy: false,
    runtimeKey: null,
    requestId: null,
    action: null,
    progress: null,
    receipt: null,
    error: null,
    canCancel: false,
  });
export function createCalculationJobController(frame: () => CalculationFrame | null) {
  let snapshot = idle(),
    disposed = false,
    serial = 0;
  const listeners = new Set<() => void>();
  let pending: {
    id: string;
    frame: CalculationFrame;
    key: string;
    runtime: string;
    action: CalculationAction;
    promise: Promise<CalculationGroupReceipt>;
    cancel: Promise<unknown> | null;
  } | null = null;
  const publish = (patch: Partial<CalculationJobSnapshot>) => {
    if (disposed) return;
    snapshot = Object.freeze({ ...snapshot, ...patch });
    for (const listener of listeners) listener();
  };
  const current = (run: NonNullable<typeof pending>) => {
    const now = frame();
    return !disposed && !!now && now.scope.active && now.scope.normal && runtimeFrameKey(now) === run.runtime;
  };
  const visible = (): CalculationJobSnapshot => {
    const now = frame();
    return now && now.scope.active && now.scope.normal && snapshot.runtimeKey === runtimeFrameKey(now)
      ? snapshot
      : Object.freeze({ ...idle(), physicalBusy: snapshot.physicalBusy });
  };
  return {
    subscribe: (listener: () => void) => {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    getSnapshot: () => snapshot,
    visible,
    async run(actionInput: CalculationAction, expectedKey: string): Promise<CalculationGroupReceipt> {
      const actual = frame(),
        action = validateAction(actionInput);
      if (disposed || !actual || calculationFrameKey(actual) !== expectedKey) throw new CalculationJobRetired();
      if (pending) {
        if (pending.key === expectedKey && actionKey(pending.action) === actionKey(action)) return pending.promise;
        throw Error('Wait for the current calculation to finish.');
      }
      const reason = calculationDisabledReason(actual, action);
      if (reason) throw Error(reason);
      const captured: CalculationFrame = {
        scope: { ...actual.scope },
        controls: { ...actual.controls, receipt: { ...actual.controls.receipt } },
      };
      const id = globalThis.crypto?.randomUUID?.() ?? `calculation-${Date.now()}-${++serial}`;
      const run = {
        id,
        frame: captured,
        key: expectedKey,
        runtime: runtimeFrameKey(captured),
        action,
        promise: null as unknown as Promise<CalculationGroupReceipt>,
        cancel: null as Promise<unknown> | null,
      };
      pending = run;
      publish({
        phase: 'running',
        physicalBusy: true,
        runtimeKey: run.runtime,
        requestId: id,
        action,
        progress: null,
        receipt: null,
        error: null,
        canCancel: true,
      });
      const onProgress = (value: CalculationProgress) => {
        if (
          pending !== run ||
          !current(run) ||
          !object(value) ||
          value.requestId !== id ||
          !text(value.phase) ||
          !Number.isSafeInteger(value.completed) ||
          !Number.isSafeInteger(value.total) ||
          value.completed < 0 ||
          value.total < value.completed ||
          value.total > 128 ||
          typeof value.cancellable !== 'boolean' ||
          (value.currentName !== undefined && typeof value.currentName !== 'string')
        )
          return;
        publish({ progress: { ...value }, canCancel: snapshot.phase === 'running' && value.cancellable });
      };
      // Assign the physical promise before entering any host callback, so a
      // same-tick duplicate cannot launch another owner job.
      run.promise = Promise.resolve()
        .then(() => {
          const before = frame();
          if (!before || disposed || calculationFrameKey(before) !== run.key) throw new CalculationJobRetired();
          return captured.controls.run(action, id, onProgress);
        })
        .then((value) => {
          const receipt = validateCalculationReceipt(value, {
            requestId: id,
            receipt: captured.controls.receipt,
            action,
          });
          if (current(run)) publish({ phase: receipt.outcome, receipt, canCancel: false, progress: null });
          return receipt;
        })
        .catch((error) => {
          const failedFrame = frame();
          if (!current(run) || !failedFrame || calculationFrameKey(failedFrame) !== run.key) {
            publish({ phase: 'idle', runtimeKey: null, receipt: null, error: null, progress: null });
            throw new CalculationJobRetired(error);
          }
          publish({
            phase: 'failed',
            error: error instanceof Error ? error.message : String(error),
            canCancel: false,
            progress: null,
          });
          throw error;
        })
        .finally(() => {
          if (pending === run) {
            pending = null;
            publish({ physicalBusy: false, canCancel: false });
          }
        });
      return run.promise;
    },
    async cancel(expectedRequestId: string): Promise<void> {
      const run = pending;
      if (!run || run.id !== expectedRequestId || !current(run) || !snapshot.canCancel) return;
      if (run.cancel) {
        await run.cancel;
        return;
      }
      publish({ canCancel: false });
      run.cancel = Promise.resolve()
        .then(() => run.frame.controls.cancel(run.id))
        .then((reply) => {
          if (!object(reply) || !['cancelled', 'completed', 'finishing'].includes(reply.outcome as string))
            throw Error('Invalid calculation cancellation reply.');
          if (pending === run && current(run))
            publish({ phase: reply.outcome === 'cancelled' ? 'stopping' : 'finishing', canCancel: false });
        })
        .catch((error) => {
          if (pending === run && current(run))
            publish({ error: error instanceof Error ? error.message : String(error), canCancel: true });
          throw error;
        });
      await run.cancel;
    },
    dispose() {
      disposed = true;
      listeners.clear();
      snapshot = idle();
    },
  };
}
export type CalculationJobController = ReturnType<typeof createCalculationJobController>;
