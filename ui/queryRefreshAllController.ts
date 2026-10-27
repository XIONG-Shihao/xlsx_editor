import type {
  LocalQueryRefreshAllScope,
  LocalQueryRefreshAllPreview,
  LocalQueryRefreshAllResult,
} from '../core/local-query.mjs';
export interface QueryRefreshAllControls {
  name?: string;
  scope: LocalQueryRefreshAllScope;
  preview: (requestId: string) => Promise<{ refreshAllPreview?: LocalQueryRefreshAllPreview; cancelled?: boolean }>;
  refresh: (previewId: string, requestId: string) => Promise<{ refreshAllResult: LocalQueryRefreshAllResult }>;
  cancel: (requestId: string) => Promise<{ outcome?: string; reason?: string; granted?: boolean }>;
}
export interface QueryRefreshAllFrame {
  key: string;
  ownerKey: string;
  name: string;
  active: boolean;
  busy?: boolean;
  reason?: string;
  controls: QueryRefreshAllControls;
}
export interface QueryRefreshAllSnapshot {
  key: string;
  ownerKey: string;
  phase: 'idle' | 'previewing' | 'preview' | 'running' | 'completed' | 'failed';
  preview: LocalQueryRefreshAllPreview | null;
  result: LocalQueryRefreshAllResult | null;
  message: string;
  pending: boolean;
  cancelPending: boolean;
}
const uuid = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;
const scopeKey = (value: LocalQueryRefreshAllScope) =>
  JSON.stringify([value.documentId, value.ownerId ?? null, value.state]);
function validatePreview(
  value: LocalQueryRefreshAllPreview,
  scope: LocalQueryRefreshAllScope
): LocalQueryRefreshAllPreview {
  const p = structuredClone(value);
  if (
    !p ||
    p.schema !== 1 ||
    !uuid.test(p.id) ||
    !p.scope ||
    scopeKey(p.scope) !== scopeKey(scope) ||
    !Array.isArray(p.entries) ||
    p.entries.length > 32 ||
    !Number.isSafeInteger(p.expiresAt) ||
    p.expiresAt <= Date.now()
  )
    throw Error('The Refresh All preview belongs to another workbook or has expired. Preview again.');
  const ids = new Set<string>();
  for (const e of p.entries) {
    if (
      !uuid.test(e.connectionId) ||
      ids.has(e.connectionId) ||
      !Number.isSafeInteger(e.revision) ||
      e.revision < 1 ||
      typeof e.name !== 'string' ||
      !e.name ||
      !e.target ||
      typeof e.target.sheet !== 'string' ||
      typeof e.target.start !== 'string' ||
      !['ready', 'refused'].includes(e.status) ||
      (e.status === 'refused' && typeof e.reason !== 'string')
    )
      throw Error('The Refresh All connection summary is invalid.');
    ids.add(e.connectionId);
    if (
      (e.source &&
        (typeof e.source.name !== 'string' ||
          !Number.isSafeInteger(e.source.bytes) ||
          e.source.bytes < 0 ||
          typeof e.source.sha256 !== 'string')) ||
      (e.outputRows !== undefined && (!Number.isSafeInteger(e.outputRows) || e.outputRows < 0)) ||
      (e.outputColumns !== undefined &&
        (!Number.isInteger(e.outputColumns) || e.outputColumns < 1 || e.outputColumns > 256)) ||
      (e.range !== undefined && typeof e.range !== 'string')
    )
      throw Error('The Refresh All source or output summary is invalid.');
    if (
      e.sample &&
      (!Array.isArray(e.sample) ||
        e.sample.length > 3 ||
        e.sample.some((row) => !Array.isArray(row) || row.length > 4 || row.some((value) => typeof value !== 'string')))
    )
      throw Error('The Refresh All source sample is invalid.');
    if (
      e.conversions &&
      (!Array.isArray(e.conversions.policies) ||
        e.conversions.policies.length > 256 ||
        e.conversions.policies.some((v) => !['text', 'number', 'boolean', 'iso-date'].includes(v)) ||
        !Number.isSafeInteger(e.conversions.rejected) ||
        !Number.isSafeInteger(e.conversions.changed) ||
        e.conversions.rejected < 0 ||
        e.conversions.changed < 0)
    )
      throw Error('The Refresh All conversion summary is invalid.');
  }
  return p;
}
function validateResult(
  value: LocalQueryRefreshAllResult,
  preview: LocalQueryRefreshAllPreview
): LocalQueryRefreshAllResult {
  const r = structuredClone(value);
  if (
    !r ||
    r.schema !== 1 ||
    r.previewId !== preview.id ||
    scopeKey(r.scope) !== scopeKey(preview.scope) ||
    !Array.isArray(r.outcomes) ||
    r.outcomes.length !== preview.entries.length ||
    typeof r.cancelled !== 'boolean' ||
    typeof r.stopped !== 'boolean'
  )
    throw Error('The Refresh All outcome belongs to another captured preview.');
  let head = preview.scope.state,
    accepted = 0,
    uncertain = 0;
  for (const [index, row] of r.outcomes.entries()) {
    const e = preview.entries[index];
    if (
      row.connectionId !== e.connectionId ||
      row.name !== e.name ||
      row.target.sheet !== e.target.sheet ||
      row.target.start !== e.target.start ||
      !['refreshed', 'refused', 'cancelled', 'not-run', 'uncertain'].includes(row.status)
    )
      throw Error('The Refresh All outcome targets a different connection.');
    if (row.status === 'refreshed') {
      if (row.undoUnits !== 1 || row.parentState !== head || typeof row.state !== 'string' || row.state === head)
        throw Error('The Refresh All accepted history chain is invalid.');
      head = row.state;
      accepted++;
    } else if (row.status === 'uncertain') {
      if (row.undoUnits !== null) throw Error('An uncertain connection cannot claim an exact Undo count.');
      uncertain++;
    } else if (row.undoUnits !== 0) throw Error('An unaccepted connection cannot claim an Undo step.');
  }
  if (
    r.accepted !== accepted ||
    r.uncertain !== uncertain ||
    r.lastAcknowledgedState !== head ||
    r.undoUnits !== (uncertain ? null : accepted)
  )
    throw Error('The Refresh All outcome counts are inconsistent.');
  return r;
}
/** Captured controls/preview stay private. One physical request owns cancellation
 * even if the pane, accepted head or initiating runtime changes. */
export function createQueryRefreshAllController() {
  let frame: QueryRefreshAllFrame | null = null,
    prepared: { frame: QueryRefreshAllFrame; preview: LocalQueryRefreshAllPreview } | null = null,
    flight: { frame: QueryRefreshAllFrame; id: string; cancelled: boolean } | null = null,
    disposed = false,
    cancelPending = false;
  let snapshot: QueryRefreshAllSnapshot = {
    key: '',
    ownerKey: '',
    phase: 'idle',
    preview: null,
    result: null,
    message: '',
    pending: false,
    cancelPending: false,
  };
  const listeners = new Set<() => void>();
  const emit = (
    phase: QueryRefreshAllSnapshot['phase'],
    preview: LocalQueryRefreshAllPreview | null = null,
    result: LocalQueryRefreshAllResult | null = null,
    message = ''
  ) => {
    snapshot = {
      key: frame?.key ?? '',
      ownerKey: frame?.ownerKey ?? '',
      phase,
      preview: preview ? structuredClone(preview) : null,
      result: result ? structuredClone(result) : null,
      message,
      pending: !!flight,
      cancelPending,
    };
    for (const listener of listeners) listener();
  };
  const sameOwner = (captured: QueryRefreshAllFrame) =>
    !disposed && !!frame?.active && frame.ownerKey === captured.ownerKey;
  const sameTarget = (captured: QueryRefreshAllFrame) =>
    sameOwner(captured) && frame?.key === captured.key && !frame.reason;
  const capture = (): QueryRefreshAllFrame => ({
    ...frame!,
    controls: { ...frame!.controls, scope: structuredClone(frame!.controls.scope) },
  });
  return {
    subscribe: (listener: () => void) => {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    getSnapshot: () => snapshot,
    setFrame(next: QueryRefreshAllFrame | null) {
      if (disposed) return;
      const before = frame;
      frame = next;
      if (prepared && !sameTarget(prepared.frame)) {
        prepared = null;
        emit(
          flight ? 'running' : 'idle',
          null,
          null,
          next?.active ? 'The workbook changed. Preview Refresh All again.' : ''
        );
        return;
      }
      if (before?.ownerKey !== next?.ownerKey || before?.active !== next?.active) emit(flight ? 'running' : 'idle');
    },
    async preview() {
      if (flight || cancelPending || !frame?.active || frame.reason || frame.busy) return false;
      const captured = capture(),
        request = { frame: captured, id: crypto.randomUUID(), cancelled: false };
      flight = request;
      prepared = null;
      emit('previewing', null, null, 'Inspecting saved local connections…');
      try {
        const response = await captured.controls.preview(request.id);
        if (!sameTarget(captured) || request.cancelled || response.cancelled) {
          if (sameOwner(captured))
            emit('idle', null, null, 'Preview cancelled or retired. No connection was refreshed.');
          return false;
        }
        const preview = validatePreview(response.refreshAllPreview!, captured.controls.scope);
        prepared = { frame: captured, preview: structuredClone(preview) };
        emit('preview', preview);
        return true;
      } catch (error) {
        if (sameTarget(captured)) emit('failed', null, null, error instanceof Error ? error.message : String(error));
        return false;
      } finally {
        if (flight === request) flight = null;
        if (sameOwner(captured)) emit(snapshot.phase, snapshot.preview, snapshot.result, snapshot.message);
        else emit('idle');
      }
    },
    async run() {
      const target = prepared;
      if (flight || cancelPending || !target || !sameTarget(target.frame) || frame?.busy) return false;
      prepared = null;
      const request = { frame: target.frame, id: crypto.randomUUID(), cancelled: false };
      flight = request;
      emit('running', target.preview, null, 'Refreshing ready local connections…');
      try {
        const response = await target.frame.controls.refresh(target.preview.id, request.id),
          result = validateResult(response.refreshAllResult, target.preview);
        if (sameOwner(target.frame))
          emit(
            'completed',
            null,
            result,
            result.cancelled
              ? 'Refresh All stopped. Accepted connection results were kept.'
              : result.uncertain
                ? 'Some acceptance is uncertain. Reconcile the original query receipts before retrying.'
                : 'Refresh All finished.'
          );
        return true;
      } catch (error) {
        if (sameOwner(target.frame)) emit('failed', null, null, error instanceof Error ? error.message : String(error));
        return false;
      } finally {
        if (flight === request) flight = null;
        if (sameOwner(target.frame)) emit(snapshot.phase, snapshot.preview, snapshot.result, snapshot.message);
        else emit('idle');
      }
    },
    async cancel() {
      const request = flight;
      if (!request) {
        prepared = null;
        emit('idle', null, null, 'Refresh All preview cancelled.');
        return true;
      }
      if (cancelPending) return false;
      cancelPending = true;
      request.cancelled = true;
      emit(
        snapshot.phase,
        snapshot.preview,
        snapshot.result,
        'Cancellation requested. The current acceptance can finish; later connections will stop.'
      );
      try {
        const result = await request.frame.controls.cancel(request.id);
        if (
          sameOwner(request.frame) &&
          flight === request &&
          (result.outcome === 'finishing' || result.reason === 'finishing')
        )
          emit(
            snapshot.phase,
            snapshot.preview,
            snapshot.result,
            'The current connection is finishing. Try Cancel again to stop remaining connections.'
          );
        return result.granted !== false;
      } catch (error) {
        if (sameOwner(request.frame) && flight === request)
          emit(
            snapshot.phase,
            snapshot.preview,
            snapshot.result,
            error instanceof Error ? error.message : String(error)
          );
        return false;
      } finally {
        cancelPending = false;
        if (!disposed) emit(snapshot.phase, snapshot.preview, snapshot.result, snapshot.message);
      }
    },
    dispose() {
      disposed = true;
      frame = null;
      prepared = null;
      emit(flight ? 'running' : 'idle');
      listeners.clear();
    },
  };
}
