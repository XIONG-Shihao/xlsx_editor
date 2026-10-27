/** Readonly preview requests share one physical slot; no workbook command lives here. */
import type { WorkbookFormulaPreview, WorkbookFormulaPreviewReceipt } from '../core/contracts';

export type FormulaPreviewReceipt = WorkbookFormulaPreviewReceipt;
export interface FormulaPreviewControls {
  receipt: FormulaPreviewReceipt;
  preview: (sheet: string, cell: string, formula: string, requestId: string) => Promise<WorkbookFormulaPreview>;
  cancel: (requestId: string) => Promise<void>;
}
export interface FormulaPreviewScope {
  documentId: string;
  ownerIdentity: string;
  state: string;
  sheet: string;
  cell: string;
  navigationGeneration: number;
  activationEpoch: number;
  modeEpoch: number;
  active: boolean;
  normal: boolean;
  busy: boolean;
  profile: 'xlsx' | 'delimited' | 'passive';
  receipt: FormulaPreviewReceipt;
}
export interface FormulaPreviewFrame {
  scope: FormulaPreviewScope;
  controls: FormulaPreviewControls;
}
export type FormulaPreviewConsumer = 'bar' | 'functions';
export interface FormulaPreviewInput {
  scope: FormulaPreviewScope;
  formula: string;
  generation: number;
  active: boolean;
  composing?: boolean;
  invalidReason?: string;
}
export type FormulaPreviewPhase =
  | 'idle'
  | 'waiting'
  | 'running'
  | 'stopping'
  | 'result'
  | 'invalid'
  | 'unavailable'
  | 'cancelled';
export interface FormulaPreviewSnapshot {
  key: string | null;
  phase: FormulaPreviewPhase;
  requestId: string | null;
  reply: WorkbookFormulaPreview | null;
  reason: string | null;
}
export interface FormulaPreviewSnapshotPair {
  bar: FormulaPreviewSnapshot;
  functions: FormulaPreviewSnapshot;
}
export interface FormulaDraftPreviewController {
  getSnapshot: () => FormulaPreviewSnapshotPair;
  subscribe: (listener: () => void) => () => void;
  setInput: (consumer: FormulaPreviewConsumer, input: FormulaPreviewInput | null) => void;
  cancel: (consumer: FormulaPreviewConsumer, expectedKey: string) => void;
  dispose: () => void;
}
const blank = (): FormulaPreviewSnapshot => ({ key: null, phase: 'idle', requestId: null, reply: null, reason: null });
const object = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);
const errorCodes = new Set(['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#NULL!', '#SPILL!', '#CALC!']);
const receiptValid = (value: FormulaPreviewReceipt): boolean =>
  object(value) &&
  Object.keys(value).sort().join(',') === 'documentId,ownerId,revision,state' &&
  Object.values(value).every((item) => typeof item === 'string' && item.length > 0 && item.length <= 4096);
const receiptKey = (value: FormulaPreviewReceipt) =>
  JSON.stringify([value.documentId, value.ownerId, value.state, value.revision]);
export const formulaPreviewScopeKey = (value: FormulaPreviewScope): string =>
  JSON.stringify([
    value.documentId,
    value.ownerIdentity,
    value.state,
    value.sheet,
    value.cell,
    value.navigationGeneration,
    value.activationEpoch,
    value.modeEpoch,
    value.active,
    value.normal,
    value.busy,
    value.profile,
    receiptKey(value.receipt),
  ]);
export const formulaPreviewInputKey = (value: FormulaPreviewInput): string =>
  JSON.stringify([
    formulaPreviewScopeKey(value.scope),
    value.formula,
    value.generation,
    value.active,
    !!value.composing,
    value.invalidReason ?? null,
  ]);
function available(scope: FormulaPreviewScope): boolean {
  return (
    scope.active &&
    scope.normal &&
    !scope.busy &&
    scope.profile === 'xlsx' &&
    !!scope.documentId &&
    !!scope.ownerIdentity &&
    receiptValid(scope.receipt)
  );
}
/** Echoed public receipt is independent of the opaque UI state token. */
export function validateFormulaPreviewReply(
  value: unknown,
  input: FormulaPreviewInput,
  requestId: string
): WorkbookFormulaPreview {
  if (
    !object(value) ||
    Object.keys(value).some(
      (key) => !['schema', 'requestId', 'receipt', 'sheet', 'cell', 'formula', 'authority', 'result'].includes(key)
    ) ||
    value.schema !== 1 ||
    value.authority !== 'draft-preview' ||
    value.requestId !== requestId ||
    value.sheet !== input.scope.sheet ||
    value.cell !== input.scope.cell ||
    value.formula !== input.formula ||
    !object(value.receipt) ||
    Object.keys(value.receipt).sort().join(',') !== 'documentId,ownerId,revision,state' ||
    receiptKey(value.receipt as unknown as FormulaPreviewReceipt) !== receiptKey(input.scope.receipt) ||
    !object(value.result)
  )
    throw Error('The formula preview did not match this draft and workbook revision.');
  const result = value.result;
  if (result.status === 'value') {
    if (
      Object.keys(result).sort().join(',') !== 'status,type,value' ||
      !['number', 'text', 'boolean', 'error'].includes(String(result.type)) ||
      (result.type === 'number' && (typeof result.value !== 'number' || !Number.isFinite(result.value))) ||
      (result.type === 'boolean' && typeof result.value !== 'boolean') ||
      (result.type === 'text' &&
        (typeof result.value !== 'string' || new TextEncoder().encode(result.value).length > 32768)) ||
      (result.type === 'error' && (typeof result.value !== 'string' || !errorCodes.has(result.value)))
    )
      throw Error('The formula preview returned an invalid typed value.');
  } else if (
    Object.keys(result).sort().join(',') !== 'reason,status' ||
    !['invalid', 'unavailable', 'cancelled'].includes(String(result.status)) ||
    typeof result.reason !== 'string' ||
    !result.reason ||
    new TextEncoder().encode(result.reason).length > 2048
  )
    throw Error('The formula preview returned an invalid result.');
  return value as unknown as WorkbookFormulaPreview;
}
interface Entry {
  input: FormulaPreviewInput;
  key: string;
  serial: number;
  cancelled: boolean;
}
interface Physical {
  consumer: FormulaPreviewConsumer;
  entry: Entry;
  frame: FormulaPreviewFrame;
  requestId: string;
  retired: boolean;
  cancelSent: boolean;
}
export function createFormulaDraftPreviewController(
  current: () => FormulaPreviewFrame | null,
  { debounceMs = 200 }: { debounceMs?: number } = {}
): FormulaDraftPreviewController {
  if (!Number.isFinite(debounceMs) || debounceMs < 0 || debounceMs > 1000)
    throw Error('Invalid formula preview delay.');
  let snapshot: FormulaPreviewSnapshotPair = { bar: blank(), functions: blank() },
    timer: ReturnType<typeof setTimeout> | null = null,
    physical: Physical | null = null,
    serial = 0,
    disposed = false;
  const entries = new Map<FormulaPreviewConsumer, Entry>(),
    listeners = new Set<() => void>();
  const emit = () => {
    for (const listener of listeners) listener();
  };
  const put = (consumer: FormulaPreviewConsumer, value: FormulaPreviewSnapshot) => {
    snapshot = { ...snapshot, [consumer]: value };
    emit();
  };
  const valid = (entry: Entry) => {
    const frame = current();
    return (
      !disposed &&
      !entry.cancelled &&
      entry.input.active &&
      !entry.input.composing &&
      available(entry.input.scope) &&
      !!frame &&
      available(frame.scope) &&
      formulaPreviewScopeKey(frame.scope) === formulaPreviewScopeKey(entry.input.scope) &&
      receiptKey(frame.controls.receipt) === receiptKey(entry.input.scope.receipt)
    );
  };
  const chosen = (): [FormulaPreviewConsumer, Entry] | null => {
    for (const consumer of ['functions', 'bar'] as const) {
      const entry = entries.get(consumer);
      if (entry?.input.active) return [consumer, entry];
    }
    return null;
  };
  const retire = () => {
    if (!physical || physical.retired) return;
    physical.retired = true;
    if (!physical.cancelSent) {
      physical.cancelSent = true;
      void physical.frame.controls.cancel(physical.requestId).catch(() => {});
    }
  };
  const run = async (consumer: FormulaPreviewConsumer, entry: Entry) => {
    const frame = current();
    if (physical || !frame || !valid(entry) || entries.get(consumer) !== entry) return;
    const captured = {
      scope: { ...frame.scope, receipt: { ...frame.scope.receipt } },
      controls: { ...frame.controls, receipt: { ...frame.controls.receipt } },
    };
    const requestId = globalThis.crypto.randomUUID(),
      job: Physical = { consumer, entry, frame: captured, requestId, retired: false, cancelSent: false };
    physical = job;
    put(consumer, { key: entry.key, phase: 'running', requestId, reply: null, reason: null });
    try {
      const raw = await captured.controls.preview(
        entry.input.scope.sheet,
        entry.input.scope.cell,
        entry.input.formula,
        requestId
      );
      if (job.retired || entries.get(consumer) !== entry || !valid(entry) || chosen()?.[1] !== entry) return;
      const reply = validateFormulaPreviewReply(raw, entry.input, requestId);
      put(consumer, { key: entry.key, phase: 'result', requestId, reply, reason: null });
    } catch (error) {
      if (!job.retired && entries.get(consumer) === entry && valid(entry))
        put(consumer, {
          key: entry.key,
          phase: 'unavailable',
          requestId,
          reply: null,
          reason: error instanceof Error ? error.message : String(error),
        });
    } finally {
      if (physical === job) physical = null;
      schedule();
    }
  };
  const schedule = () => {
    if (timer !== null) {
      clearTimeout(timer);
      timer = null;
    }
    const next = chosen();
    if (!next) {
      retire();
      return;
    }
    const [consumer, entry] = next;
    for (const other of ['bar', 'functions'] as const)
      if (other !== consumer && snapshot[other].phase !== 'idle') put(other, blank());
    if (physical && physical.entry !== entry) retire();
    if (!valid(entry)) {
      if (!entry.cancelled)
        put(consumer, { key: entry.key, phase: 'idle', requestId: null, reply: null, reason: null });
      return;
    }
    if (
      !entry.input.formula.trim() ||
      entry.input.invalidReason ||
      new TextEncoder().encode(entry.input.formula).length > 8192
    ) {
      put(consumer, {
        key: entry.key,
        phase: 'invalid',
        requestId: null,
        reply: null,
        reason:
          entry.input.invalidReason ??
          (!entry.input.formula.trim()
            ? 'Enter a complete formula to preview its result.'
            : 'This formula is too long to preview.'),
      });
      return;
    }
    if (
      physical ||
      (snapshot[consumer].key === entry.key &&
        ['running', 'result', 'unavailable', 'invalid', 'cancelled'].includes(snapshot[consumer].phase))
    )
      return;
    put(consumer, { key: entry.key, phase: 'waiting', requestId: null, reply: null, reason: null });
    timer = setTimeout(() => {
      timer = null;
      void run(consumer, entry);
    }, debounceMs);
  };
  return {
    getSnapshot: () => snapshot,
    subscribe: (listener) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    setInput: (consumer, input) => {
      if (disposed) return;
      const key = input ? formulaPreviewInputKey(input) : null,
        prior = entries.get(consumer);
      if (prior?.key === key) return;
      if (input)
        entries.set(consumer, {
          input: { ...input, scope: { ...input.scope, receipt: { ...input.scope.receipt } } },
          key: key!,
          serial: ++serial,
          cancelled: false,
        });
      else entries.delete(consumer);
      put(consumer, blank());
      schedule();
    },
    cancel: (consumer, expectedKey) => {
      const entry = entries.get(consumer);
      if (!entry || entry.key !== expectedKey) return;
      entry.cancelled = true;
      if (physical?.entry === entry) retire();
      put(consumer, {
        key: entry.key,
        phase: 'cancelled',
        requestId: physical?.entry === entry ? physical.requestId : null,
        reply: null,
        reason: 'Preview cancelled. The workbook was not changed.',
      });
      schedule();
    },
    dispose: () => {
      if (disposed) return;
      disposed = true;
      if (timer !== null) clearTimeout(timer);
      timer = null;
      retire();
      entries.clear();
      snapshot = { bar: blank(), functions: blank() };
      listeners.clear();
    },
  };
}
