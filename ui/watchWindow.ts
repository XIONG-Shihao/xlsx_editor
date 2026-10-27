import type { WorkbookCell } from '../core/contracts';
import { hasExactNumericValue } from '../core/calculation.mjs';

export const MAX_CELL_WATCHES = 32;
export interface WatchScope {
  readonly documentId: string;
  readonly ownerIdentity: string;
  readonly state: string;
  readonly activationEpoch: number;
  readonly active: boolean;
  readonly sheets: readonly string[];
}
export interface WatchTarget {
  readonly sheet: string;
  readonly reference: string;
}
export type WatchValue =
  | { readonly kind: 'blank' }
  | { readonly kind: 'number'; readonly value: number }
  | { readonly kind: 'text'; readonly value: string }
  | { readonly kind: 'boolean'; readonly value: boolean }
  | { readonly kind: 'error'; readonly error: string }
  | { readonly kind: 'protected'; readonly raw: string; readonly reason: string }
  | { readonly kind: 'circular' | 'unresolved' | 'stored'; readonly reason: string };
export interface WatchRow extends WatchTarget {
  readonly id: string;
  readonly status: 'waiting' | 'reading' | 'ready' | 'failed' | 'paused';
  readonly formula?: string;
  readonly value?: WatchValue;
  readonly failure?: string;
}
export interface WatchSnapshot {
  readonly phase: 'idle' | 'reading' | 'paused' | 'disposed';
  readonly scope: WatchScope | null;
  readonly rows: readonly WatchRow[];
}
export interface WatchReply {
  readonly scope: WatchScope;
  readonly target: WatchTarget;
  readonly cell?: WorkbookCell;
}
/** The adapter must send this captured owner/head, never its latest owner. */
export type WatchReader = (scope: WatchScope, target: WatchTarget, signal: AbortSignal) => Promise<WatchReply>;
export interface WatchRefresh {
  readonly outcome: 'read' | 'retired' | 'paused';
  readonly read: number;
  readonly failed: number;
}
const scopeFields = ['documentId', 'ownerIdentity', 'state', 'activationEpoch', 'active', 'sheets'];
function nonempty(value: unknown, limit: number): value is string {
  return typeof value === 'string' && value.length > 0 && value.length <= limit && !/[\x00-\x1f]/.test(value);
}
function sheetName(value: unknown): value is string {
  return nonempty(value, 31) && !/[\\/?*\[\]:]/.test(value);
}
function reference(value: unknown): value is string {
  if (typeof value !== 'string' || !/^([A-Z]{1,3})([1-9][0-9]{0,6})$/.test(value)) return false;
  const parts = /^([A-Z]+)([0-9]+)$/.exec(value)!,
    column = [...parts[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  return column <= 16384 && Number(parts[2]) <= 1048576;
}
function captureScope(value: WatchScope): WatchScope {
  if (
    !value ||
    Object.keys(value).some((key) => !scopeFields.includes(key)) ||
    !nonempty(value.documentId, 512) ||
    !nonempty(value.ownerIdentity, 512) ||
    !nonempty(value.state, 512) ||
    !Number.isSafeInteger(value.activationEpoch) ||
    value.activationEpoch < 0 ||
    typeof value.active !== 'boolean' ||
    !Array.isArray(value.sheets) ||
    value.sheets.length > 1000 ||
    value.sheets.some((name) => !sheetName(name)) ||
    new Set(value.sheets).size !== value.sheets.length
  )
    throw Error('Invalid watch scope.');
  return Object.freeze({ ...value, sheets: Object.freeze([...value.sheets]) });
}
function scopeKey(scope: WatchScope): string {
  return JSON.stringify([
    scope.documentId,
    scope.ownerIdentity,
    scope.state,
    scope.activationEpoch,
    scope.active,
    scope.sheets,
  ]);
}
function runtimeEqual(a: WatchScope, b: WatchScope): boolean {
  return a.documentId === b.documentId && a.ownerIdentity === b.ownerIdentity;
}
function targetKey(target: WatchTarget): string {
  return JSON.stringify([target.sheet, target.reference]);
}
function captureTarget(target: WatchTarget, scope: WatchScope): WatchTarget {
  if (
    !target ||
    Object.keys(target).some((key) => !['sheet', 'reference'].includes(key)) ||
    !sheetName(target.sheet) ||
    !scope.sheets.includes(target.sheet) ||
    !reference(target.reference)
  )
    throw Error('Choose one canonical cell on this workbook.');
  return Object.freeze({ sheet: target.sheet, reference: target.reference });
}
function text(value: unknown, limit = 32768): string {
  if (typeof value !== 'string' || value.length > limit) throw Error('Watch text exceeds the bounded scalar profile.');
  return value;
}
function typed(value: unknown): WatchValue {
  if (value === null || value === undefined) return { kind: 'blank' };
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw Error('Non-finite watch value.');
    return { kind: 'number', value };
  }
  if (typeof value === 'string') return { kind: 'text', value: text(value) };
  if (typeof value === 'boolean') return { kind: 'boolean', value };
  throw Error('Watch result is not a scalar.');
}
/** Formula status, not error-looking text or a saved cache, is authoritative. */
export function classifyWatchCell(cell: WorkbookCell | undefined): { formula?: string; value: WatchValue } {
  if (!cell) return { value: { kind: 'blank' } };
  const formula = cell.formula !== undefined ? text(cell.formula) : undefined,
    calculation = cell.calculation;
  if (calculation?.pending)
    return { formula, value: { kind: 'stored', reason: 'Stored result; pending calculation.' } };
  if (formula !== undefined || cell.formulaType) {
    if (calculation?.status === 'calculated') {
      if (calculation.value === undefined || calculation.value === null)
        throw Error('Calculated watch value is missing.');
      return { formula, value: typed(calculation.value) };
    }
    if (calculation?.status === 'error')
      return { formula, value: { kind: 'error', error: text(calculation.error, 128) } };
    if (calculation?.status === 'circular')
      return { formula, value: { kind: 'circular', reason: text(calculation.error ?? 'Circular formula.', 512) } };
    if (calculation?.status === 'unresolved')
      return {
        formula,
        value: { kind: 'unresolved', reason: text(calculation.error ?? 'Formula is unresolved.', 512) },
      };
    return {
      formula,
      value: { kind: 'stored', reason: 'Stored formula value has not been qualified at this accepted head.' },
    };
  }
  if (cell.dataType === 'error') return { value: { kind: 'error', error: text(cell.error ?? cell.value, 128) } };
  if (typeof cell.value === 'number' && Number.isFinite(cell.value) && !hasExactNumericValue(cell.value, cell.rawValue))
    return {
      value: {
        kind: 'protected',
        raw: text(cell.rawValue),
        reason: 'Stored numeric spelling is more precise than this numeric value.',
      },
    };
  const value = typed(cell.value);
  if (
    cell.dataType &&
    cell.dataType !==
      ({ blank: 'blank', number: 'number', text: 'text', boolean: 'boolean' } as const)[
        value.kind as 'blank' | 'number' | 'text' | 'boolean'
      ]
  )
    throw Error('Watch type and scalar value disagree.');
  return { value };
}
function frozenRow(row: WatchRow): WatchRow {
  return Object.freeze({ ...row, ...(row.value ? { value: Object.freeze({ ...row.value }) } : {}) });
}
export interface WatchWindowController {
  setScope(scope: WatchScope | null): void;
  getSnapshot(): WatchSnapshot;
  subscribe(listener: () => void): () => void;
  add(target: WatchTarget): boolean;
  remove(target: WatchTarget): boolean;
  clear(): void;
  refresh(): Promise<WatchRefresh>;
  cancelRefresh(): void;
  dispose(): void;
}
export function createWatchWindowController(reader: WatchReader): WatchWindowController {
  if (typeof reader !== 'function') throw Error('A bounded watch reader is required.');
  let read: WatchReader | undefined = reader,
    scope: WatchScope | null = null,
    rows: WatchRow[] = [],
    generation = 0,
    disposed = false,
    abort: AbortController | undefined,
    pending: Promise<WatchRefresh> | undefined;
  let snapshot: WatchSnapshot = Object.freeze({ phase: 'paused', scope: null, rows: Object.freeze([]) });
  const listeners = new Set<() => void>();
  function publish(phase: WatchSnapshot['phase']): void {
    snapshot = Object.freeze({ phase, scope, rows: Object.freeze(rows.map(frozenRow)) });
    for (const listener of listeners)
      try {
        listener();
      } catch {
        /* Observers cannot change read acceptance. */
      }
  }
  function live(): WatchScope {
    if (disposed) throw Error('Watch controller is closed.');
    if (!scope?.active) throw Error('Activate the initiating workbook.');
    return scope;
  }
  function invalidate(): void {
    generation++;
    abort?.abort();
    rows = rows.map(({ id, sheet, reference }) => ({
      id,
      sheet,
      reference,
      status: scope?.active ? 'waiting' : 'paused',
    }));
  }
  const api: WatchWindowController = {
    setScope(value) {
      if (disposed) throw Error('Watch controller is closed.');
      const next = value === null ? null : captureScope(value);
      if (scope && next && scopeKey(scope) === scopeKey(next)) return;
      if (!scope || !next || !runtimeEqual(scope, next)) rows = [];
      scope = next;
      rows = rows.filter((row) => scope?.sheets.includes(row.sheet));
      invalidate();
      publish(scope?.active ? 'idle' : 'paused');
    },
    getSnapshot: () => snapshot,
    subscribe(listener) {
      if (disposed) return () => {};
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    add(target) {
      const current = live(),
        value = captureTarget(target, current),
        id = targetKey(value);
      if (rows.some((row) => row.id === id)) return false;
      if (rows.length >= MAX_CELL_WATCHES) throw Error('The watch window is limited to 32 cells.');
      invalidate();
      rows.push({ ...value, id, status: 'waiting' });
      publish('idle');
      return true;
    },
    remove(target) {
      const current = live(),
        id = targetKey(captureTarget(target, current));
      if (!rows.some((row) => row.id === id)) return false;
      invalidate();
      rows = rows.filter((row) => row.id !== id);
      publish('idle');
      return true;
    },
    clear() {
      if (disposed) return;
      invalidate();
      rows = [];
      publish(scope?.active ? 'idle' : 'paused');
    },
    refresh() {
      if (disposed || !scope?.active) return Promise.resolve({ outcome: 'paused', read: 0, failed: 0 });
      // Ignored AbortSignals cannot multiply outstanding owner reads. A stale
      // physical read retires before any new refresh may start.
      if (pending) return pending;
      const captured = scope,
        capturedGeneration = generation,
        targets = rows.map((row) => ({ sheet: row.sheet, reference: row.reference })),
        capturedRead = read!;
      abort = new AbortController();
      const signal = abort.signal,
        values: WatchRow[] = [],
        isCurrent = () =>
          !disposed &&
          !signal.aborted &&
          generation === capturedGeneration &&
          !!scope &&
          scopeKey(scope) === scopeKey(captured);
      rows = rows.map((row) => ({
        ...row,
        status: 'reading',
        value: undefined,
        formula: undefined,
        failure: undefined,
      }));
      let count = 0,
        failures = 0;
      const work = Promise.resolve().then(async (): Promise<WatchRefresh> => {
        for (const target of targets) {
          if (!isCurrent()) return { outcome: 'retired', read: count, failed: failures };
          try {
            count++;
            const reply = await capturedRead(captured, Object.freeze({ ...target }), signal);
            if (!isCurrent()) return { outcome: 'retired', read: count, failed: failures };
            if (
              !reply ||
              scopeKey(captureScope(reply.scope)) !== scopeKey(captured) ||
              targetKey(captureTarget(reply.target, captured)) !== targetKey(target)
            )
              throw Error('Watch reply belongs to another owner, head or cell.');
            values.push({ ...target, id: targetKey(target), status: 'ready', ...classifyWatchCell(reply.cell) });
          } catch (error) {
            if (!isCurrent()) return { outcome: 'retired', read: count, failed: failures };
            failures++;
            values.push({
              ...target,
              id: targetKey(target),
              status: 'failed',
              failure: String(error instanceof Error ? error.message : error).slice(0, 512),
            });
          }
        }
        if (!isCurrent()) return { outcome: 'retired', read: count, failed: failures };
        rows = values;
        publish('idle');
        return { outcome: 'read', read: count, failed: failures };
      });
      const final = work.finally(() => {
        if (pending === final) {
          pending = undefined;
          abort = undefined;
        }
      });
      pending = final;
      publish('reading');
      return final;
    },
    cancelRefresh() {
      if (disposed) return;
      invalidate();
      publish(scope?.active ? 'idle' : 'paused');
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      invalidate();
      scope = null;
      rows = [];
      read = undefined;
      publish('disposed');
      listeners.clear();
    },
  };
  return Object.freeze(api);
}
