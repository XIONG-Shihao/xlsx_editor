/** Trusted local coordination. A renderer sees summaries/tokens, never grants.
 * One accepted connection is one Undo unit; this is not an atomic batch. */
import { randomUUID } from 'node:crypto';
import { LocalQueryHost } from './local-query-host.mjs';
import { rangeBounds, columnName } from '../core/workbook-commands.mjs';
import { isRefusal, refusal } from '../core/refusal-codes.mjs';

export const REFRESH_ALL_LIMITS = Object.freeze({
  plans: 4,
  connections: 32,
  lifetimeMs: 10 * 60 * 1000,
  responseBytes: 1024 * 1024,
});
const key = (value) => JSON.stringify(value),
  message = (error) =>
    error?.code && !isRefusal(error)
      ? 'Local query could not access its granted source or registry (' + error.code + ').'
      : String(error?.message ?? error);
const scopeKey = (scope) => key([scope.documentId, scope.ownerId ?? null, scope.state]);
const sameOwner = (a, b) => a.documentId === b.documentId && (a.ownerId ?? null) === (b.ownerId ?? null);
const rectangle = (shape) => {
  const start = rangeBounds(shape.start).first;
  return {
    sheet: shape.sheet,
    left: start.col,
    top: start.row,
    right: start.col + shape.columns - 1,
    bottom: start.row + shape.rows - 1,
  };
};
const overlaps = (a, b) =>
  a && b && a.sheet === b.sheet && a.left <= b.right && b.left <= a.right && a.top <= b.bottom && b.top <= a.bottom;
const range = (shape) => {
  const b = rectangle(shape);
  return shape.start + ':' + columnName(b.right) + (b.bottom + 1);
};
const sample = (rows) =>
  rows.slice(0, 3).map((row) => row.slice(0, 4).map((value) => Array.from(String(value)).slice(0, 128).join('')));

export class LocalQueryRefreshAll {
  #store;
  #document;
  #directory;
  #context;
  #now;
  #plans = new Map();
  #busy = false;
  #disposed = false;
  /** Context.current enforces the original actor/activation, but returns the
   * current native head. This coordinator fences that head itself; only the
   * actual apply reply can advance its private expected state. */
  constructor({ store, document, directory, context, now = Date.now }) {
    this.#store = store;
    this.#document = document;
    this.#directory = directory;
    this.#context = context;
    this.#now = now;
  }
  #current() {
    if (this.#disposed) throw refusal('query-stale', 'Refresh All owner retired');
    const scope = structuredClone(this.#context.current());
    if (
      !scope ||
      typeof scope.documentId !== 'string' ||
      !scope.documentId ||
      typeof scope.state !== 'string' ||
      !scope.state
    )
      throw refusal('query-bad-input', 'Refresh All requires its original workbook owner');
    return scope;
  }
  #assert(scope) {
    const live = this.#current();
    if (scopeKey(live) !== scopeKey(scope))
      throw refusal('query-stale', 'Refresh All workbook owner or accepted head changed. Preview again.');
    return live;
  }
  #owner(scope) {
    const live = this.#current();
    if (!sameOwner(live, scope)) throw refusal('query-stale', 'Refresh All original owner retired');
    return live;
  }
  async #membership(units) {
    const found = await this.#store.list(this.#document);
    if (key(found.entries.map((row) => row.id).sort()) !== key(units.map((unit) => unit.entry.connectionId).sort()))
      throw refusal('query-stale', 'The saved connection set changed. Preview Refresh All again.');
  }
  #prune() {
    for (const [id, plan] of this.#plans) if (plan.expires <= this.#now()) this.#plans.delete(id);
  }
  #host(scope, options) {
    return new LocalQueryHost({
      store: this.#store,
      document: this.#document,
      directory: this.#directory,
      context: {
        current: (fenced) => (fenced === false ? this.#owner(scope) : this.#assert(scope)),
        calendar: (sheet) => {
          this.#assert(scope);
          return this.#context.calendar?.(sheet);
        },
        request: async (operation, input, fenced = true) => {
          const check = () => (fenced ? this.#assert(scope) : this.#owner(scope));
          check();
          const result = await this.#context.request(operation, input, fenced);
          check();
          return result;
        },
        apply: async (parent, descriptor) => {
          this.#assert(scope);
          if (parent !== scope.state) throw refusal('query-stale', 'Refresh All original parent changed');
          const result = await this.#context.apply(parent, descriptor),
            live = this.#current();
          if (!sameOwner(live, scope) || result?.state !== live.state || live.state === parent)
            throw refusal('query-bad-input', 'Refresh All acceptance could not be verified');
          scope.state = live.state;
          await options.afterApply?.({ parentState: parent, state: live.state });
          return result;
        },
      },
    });
  }
  #entry(record, prepared) {
    const p = prepared.preview,
      conversion = p.conversionInspection,
      replacement = p.replacementInspection;
    let reason = conversion.rejected
      ? 'Rejected conversions require source/settings correction.'
      : replacement.blocked
        ? 'Protected destination cells block refresh.'
        : replacement.occupied
          ? 'New occupied destination cells require the individual query replacement review.'
          : undefined;
    return {
      connectionId: record.id,
      revision: record.revision,
      name: record.name,
      target: structuredClone(record.target),
      source: structuredClone(p.source),
      outputRows: p.outputRows,
      outputColumns: p.outputColumns,
      range: range(prepared.shape),
      sample: sample(p.rows),
      sampleTruncated:
        p.truncated ||
        p.rows.length > 3 ||
        p.rows.some((row) => row.length > 4 || row.some((value) => Array.from(String(value)).length > 128)),
      conversions: {
        rejected: conversion.rejected,
        changed: conversion.changed,
        policies: conversion.policies,
        date1904: conversion.date1904,
      },
      replacements: { occupied: replacement.occupied, blocked: replacement.blocked },
      status: reason ? 'refused' : 'ready',
      ...(reason ? { reason } : {}),
    };
  }
  async preview(options = {}) {
    if (this.#busy) throw refusal('query-bad-input', 'Wait for the current Refresh All operation');
    this.#busy = true;
    try {
      const scope = this.#current(),
        host = this.#host(scope, options),
        list = await host.list();
      this.#assert(scope);
      if (list.entries.length > REFRESH_ALL_LIMITS.connections)
        throw refusal('query-limit', 'Refresh All exceeds 32 local connections');
      const units = [];
      for (const record of list.entries) {
        options.signal?.throwIfAborted();
        this.#assert(scope);
        let unit;
        try {
          const prepared = await host.prepareRefreshAllUnit(record.id, record.revision, options);
          unit = {
            entry: this.#entry(record, prepared),
            proof: structuredClone(prepared.proof),
            rectangle: rectangle(prepared.shape),
          };
        } catch (error) {
          options.signal?.throwIfAborted();
          this.#assert(scope);
          unit = {
            entry: {
              connectionId: record.id,
              revision: record.revision,
              name: record.name,
              target: structuredClone(record.target),
              status: 'refused',
              reason: message(error),
            },
            rectangle: rectangle(record.last?.shape ?? { ...record.target, rows: 1, columns: 1 }),
          };
        }
        units.push(unit);
      }
      for (let i = 0; i < units.length; i++)
        for (let j = i + 1; j < units.length; j++)
          if (overlaps(units[i].rectangle, units[j].rectangle))
            for (const unit of [units[i], units[j]]) {
              unit.entry.status = 'refused';
              unit.entry.reason = 'Connection output areas overlap. Choose separate targets before Refresh All.';
            }
      options.signal?.throwIfAborted();
      this.#assert(scope);
      this.#prune();
      while (this.#plans.size >= REFRESH_ALL_LIMITS.plans) this.#plans.delete(this.#plans.keys().next().value);
      const id = randomUUID(),
        expires = this.#now() + REFRESH_ALL_LIMITS.lifetimeMs,
        preview = {
          schema: 1,
          id,
          scope: structuredClone(scope),
          entries: units.map((unit) => structuredClone(unit.entry)),
          expiresAt: expires,
          undoPolicy: 'One Undo step for each accepted connection',
          ...(list.warning ? { warning: list.warning } : {}),
        };
      if (Buffer.byteLength(JSON.stringify(preview)) > REFRESH_ALL_LIMITS.responseBytes)
        throw refusal('query-limit', 'Refresh All preview exceeds 1 MiB');
      this.#plans.set(id, { scope: structuredClone(scope), units, expires });
      return { refreshAllPreview: preview };
    } finally {
      this.#busy = false;
    }
  }
  async refresh(id, options = {}) {
    if (this.#busy) throw refusal('query-bad-input', 'Wait for the current Refresh All operation');
    this.#prune();
    const plan = this.#plans.get(id);
    if (!plan) throw refusal('query-stale', 'Refresh All preview expired or was consumed. Preview again.');
    this.#assert(plan.scope);
    this.#busy = true;
    const scope = structuredClone(plan.scope),
      outcomes = [],
      host = this.#host(scope, options);
    let stopped = false,
      cancelled = false;
    try {
      await this.#membership(plan.units);
      this.#assert(scope);
      this.#plans.delete(id);
      for (const unit of plan.units) {
        const entry = unit.entry,
          base = { connectionId: entry.connectionId, name: entry.name, target: entry.target };
        if (stopped) {
          outcomes.push({
            ...base,
            status: 'not-run',
            undoUnits: 0,
            reason: cancelled
              ? 'Refresh All cancelled.'
              : 'An uncertain or retired operation stopped the remaining connections.',
          });
          continue;
        }
        if (options.signal?.aborted) {
          cancelled = true;
          stopped = true;
          outcomes.push({
            ...base,
            status: 'cancelled',
            undoUnits: 0,
            reason: 'Refresh All cancelled before this connection.',
          });
          continue;
        }
        if (entry.status !== 'ready') {
          outcomes.push({ ...base, status: 'refused', undoUnits: 0, reason: entry.reason });
          continue;
        }
        const before = scope.state;
        try {
          this.#assert(scope);
          await this.#membership(plan.units);
          const fresh = await host.prepareRefreshAllUnit(entry.connectionId, entry.revision, options);
          if (key(fresh.proof) !== key(unit.proof))
            throw refusal(
              'query-stale',
              'Connection source, settings or output changed after the full preview. Preview again.'
            );
          const review = fresh.preview.conversionInspection.revision;
          const result = await host.refresh(entry.connectionId, entry.revision, {
            ...options,
            expectedConversionRevision: review,
          });
          if (result.refreshed !== true || scope.state === before)
            throw refusal('query-bad-input', 'Refresh All did not obtain an accepted connection receipt');
          outcomes.push({
            ...base,
            status: 'refreshed',
            undoUnits: 1,
            parentState: before,
            state: scope.state,
            outputRows: result.outputRows,
            outputColumns: result.outputColumns,
            ...(result.warning ? { warning: result.warning } : {}),
          });
        } catch (error) {
          const retained = await this.#store.granted(this.#document, entry.connectionId).catch(() => null);
          let current = false;
          try {
            this.#assert(scope);
            current = true;
          } catch {}
          const acknowledged = scope.state !== before,
            uncertain = !!retained?.pending;
          if (acknowledged) {
            outcomes.push({
              ...base,
              status: 'refreshed',
              undoUnits: 1,
              parentState: before,
              state: scope.state,
              warning: 'Data acceptance was acknowledged; receipt/finalization needs reconciliation. ' + message(error),
            });
            stopped = true;
          } else if (uncertain) {
            outcomes.push({
              ...base,
              status: 'uncertain',
              undoUnits: null,
              reason:
                'Acceptance may be retained. Reconcile the original query receipt before retrying. ' + message(error),
            });
            stopped = true;
          } else if (options.signal?.aborted) {
            outcomes.push({
              ...base,
              status: 'cancelled',
              undoUnits: 0,
              reason: 'Refresh All cancelled before acceptance.',
            });
            cancelled = true;
            stopped = true;
          } else {
            outcomes.push({ ...base, status: 'refused', undoUnits: 0, reason: message(error) });
            if (!current) stopped = true;
          }
        }
      }
      return {
        refreshAllResult: {
          schema: 1,
          previewId: id,
          scope: structuredClone(plan.scope),
          lastAcknowledgedState: scope.state,
          cancelled,
          stopped,
          outcomes,
          accepted: outcomes.filter((row) => row.status === 'refreshed').length,
          uncertain: outcomes.filter((row) => row.status === 'uncertain').length,
          undoUnits: outcomes.some((row) => row.undoUnits === null)
            ? null
            : outcomes.reduce((n, row) => n + row.undoUnits, 0),
        },
      };
    } finally {
      this.#busy = false;
    }
  }
  dispose() {
    this.#disposed = true;
    this.#plans.clear();
  }
}
