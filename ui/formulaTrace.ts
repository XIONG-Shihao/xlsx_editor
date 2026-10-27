import type {
  WorkbookTracePosition,
  WorkbookTraceEdge,
  WorkbookTraceMarker,
  WorkbookFormulaTrace,
} from '../core/contracts';
/** Readonly graph expansion. UI state is opaque; the captured adapter supplies
 * the original actual document/owner/head receipt, without parsing this token. */
export type TraceEndpoint = WorkbookTracePosition;
export type TraceDirection = 'precedents' | 'dependents';
export interface FormulaTraceScope {
  documentId: string;
  ownerIdentity: string;
  state: string;
  sheet: string;
  cell: string;
  activationEpoch: number;
  active: boolean;
  sheets: readonly string[];
}
export type FormulaTraceEdge = WorkbookTraceEdge;
export type FormulaTraceMarker = WorkbookTraceMarker;
export type FormulaTracePage = WorkbookFormulaTrace;
export interface ScopedFormulaTracePage {
  scope: FormulaTraceScope;
  target: { sheet: string; cell: string };
  direction: TraceDirection;
  after: string | null;
  page: FormulaTracePage;
}
export type FormulaTraceQuery = (
  scope: FormulaTraceScope,
  target: { sheet: string; cell: string },
  options: { direction: TraceDirection; after: string | null; limit: number },
  signal: AbortSignal
) => Promise<ScopedFormulaTracePage>;
export interface TraceEdgeRow extends FormulaTraceEdge {
  id: string;
}
export interface TraceMarkerRow extends FormulaTraceMarker {
  id: string;
}
export interface FormulaTraceSnapshot {
  scope: FormulaTraceScope | null;
  direction: TraceDirection | null;
  phase: 'idle' | 'reading' | 'failed';
  depth: number;
  edges: readonly TraceEdgeRow[];
  markers: readonly TraceMarkerRow[];
  requests: number;
  queriedCells: number;
  complete: boolean;
  canExpand: boolean;
  message: string | null;
}
export const TRACE_LIMITS = Object.freeze({ depth: 32, nodes: 1024, edges: 1024, requests: 64, pageItems: 100 });
const utf8 = (s: string) => new TextEncoder().encode(s).length;
function text(value: unknown, max: number): value is string {
  return typeof value === 'string' && value.length > 0 && utf8(value) <= max && !/[\u0000-\u001f]/u.test(value);
}
export function tracePosition(reference: string): { row: number; col: number } {
  const m = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(reference);
  if (!m) throw Error('Trace requires a canonical cell reference.');
  const col = [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1,
    row = Number(m[2]) - 1;
  if (col >= 16384 || row >= 1048576) throw Error('Trace cell is outside the worksheet.');
  return { row, col };
}
export function traceBounds(range: string): {
  first: { row: number; col: number };
  last: { row: number; col: number };
} {
  const parts = range.split(':');
  if (parts.length > 2) throw Error('Invalid trace range.');
  const first = tracePosition(parts[0]),
    last = tracePosition(parts[1] ?? parts[0]);
  if (last.row < first.row || last.col < first.col) throw Error('Invalid trace range.');
  return { first, last };
}
export function formulaTraceScopeKey(scope: FormulaTraceScope): string {
  return JSON.stringify([
    scope.documentId,
    scope.ownerIdentity,
    scope.state,
    scope.sheet,
    scope.cell,
    scope.activationEpoch,
    scope.active,
    scope.sheets,
  ]);
}
export function validateFormulaTraceScope(scope: FormulaTraceScope): void {
  if (
    !scope ||
    !text(scope.documentId, 4096) ||
    !text(scope.ownerIdentity, 4096) ||
    !text(scope.state, 4096) ||
    !text(scope.sheet, 128) ||
    !Number.isSafeInteger(scope.activationEpoch) ||
    scope.activationEpoch < 0 ||
    typeof scope.active !== 'boolean' ||
    !Array.isArray(scope.sheets) ||
    scope.sheets.length > 16384 ||
    scope.sheets.some((s) => !text(s, 128)) ||
    new Set(scope.sheets).size !== scope.sheets.length ||
    !scope.sheets.includes(scope.sheet)
  )
    throw Error('Invalid trace workbook scope.');
  tracePosition(scope.cell);
}
export function validateFormulaTracePage(
  value: unknown,
  scope: FormulaTraceScope,
  target: { sheet: string; cell: string },
  direction: TraceDirection
): FormulaTracePage {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw Error('Trace returned no valid page.');
  const p = value as FormulaTracePage;
  if (
    p.schema !== 1 ||
    (p.documentId !== undefined && !text(p.documentId, 4096)) ||
    !text(p.ownerId, 4096) ||
    !text(p.state, 4096) ||
    !text(p.revision, 4096) ||
    p.sheet !== target.sheet ||
    p.cell !== target.cell ||
    p.direction !== direction ||
    !Array.isArray(p.edges) ||
    !Array.isArray(p.markers) ||
    p.edges.length + p.markers.length > TRACE_LIMITS.pageItems ||
    !Object.hasOwn(p, 'nextCursor') ||
    (p.nextCursor !== null && !text(p.nextCursor, 4096)) ||
    !p.coverage ||
    p.coverage.mode !== 'direct-static-references' ||
    typeof p.coverage.complete !== 'boolean' ||
    typeof p.coverage.limited !== 'boolean' ||
    !Number.isSafeInteger(p.coverage.scannedFormulas) ||
    p.coverage.scannedFormulas < 0 ||
    p.coverage.scannedFormulas > 131072 ||
    !Number.isSafeInteger(p.coverage.pages) ||
    p.coverage.pages < 0 ||
    p.coverage.pages > 64 ||
    (p.coverage.complete && (p.coverage.limited || p.nextCursor !== null || p.markers.length > 0))
  )
    throw Error('Trace returned an invalid relationship page.');
  const endpoint = (e: TraceEndpoint) => {
    if (
      !e ||
      typeof e !== 'object' ||
      Array.isArray(e) ||
      !scope.sheets.includes(e.sheet) ||
      typeof e.range !== 'string'
    )
      throw Error('Invalid trace endpoint.');
    return traceBounds(e.range);
  };
  const targetPosition = tracePosition(target.cell);
  for (const edge of p.edges) {
    if (
      !edge ||
      typeof edge !== 'object' ||
      Array.isArray(edge) ||
      !['reference', 'spill-origin', 'spill-range'].includes(edge.kind) ||
      !['name', 'table', 'formula'].includes(edge.via)
    )
      throw Error('Invalid trace relationship.');
    const a = endpoint(edge.from),
      b = endpoint(edge.to),
      bound = direction === 'precedents' ? b : a,
      root = direction === 'precedents' ? edge.to : edge.from;
    if (
      root.sheet !== target.sheet ||
      targetPosition.row < bound.first.row ||
      targetPosition.row > bound.last.row ||
      targetPosition.col < bound.first.col ||
      targetPosition.col > bound.last.col
    )
      throw Error('Trace relationship does not contain its requested cell.');
  }
  for (const marker of p.markers) {
    if (
      !marker ||
      typeof marker !== 'object' ||
      Array.isArray(marker) ||
      !scope.sheets.includes(marker.sheet) ||
      !['dynamic', 'external', 'unsupported', 'special', 'limit'].includes(marker.kind) ||
      !text(marker.reason, 2048)
    )
      throw Error('Invalid trace coverage marker.');
    if (marker.cell !== undefined) tracePosition(marker.cell);
  }
  return p;
}
const cellKey = (target: { sheet: string; cell: string }) => JSON.stringify([target.sheet, target.cell]);
const edgeKey = (e: FormulaTraceEdge) =>
  JSON.stringify([e.from.sheet, e.from.range, e.to.sheet, e.to.range, e.kind, e.via ?? null]);
const markerKey = (m: FormulaTraceMarker) => JSON.stringify([m.sheet, m.cell ?? null, m.kind, m.reason]);
function column(col: number): string {
  let s = '';
  for (let n = col + 1; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
export interface FormulaTraceController {
  getSnapshot: () => FormulaTraceSnapshot;
  subscribe: (listener: () => void) => () => void;
  setScope: (scope: FormulaTraceScope) => void;
  trace: (
    direction: TraceDirection
  ) => Promise<{ outcome: 'completed' | 'retired' | 'paused' | 'limited'; depth: number }>;
  cancel: (scope?: FormulaTraceScope) => void;
  clear: (scope?: FormulaTraceScope) => void;
  dispose: () => void;
}
export function createFormulaTraceController(query: FormulaTraceQuery): FormulaTraceController {
  const listeners = new Set<() => void>();
  let generation = 0,
    disposed = false,
    abort: AbortController | null = null,
    pending: ReturnType<FormulaTraceController['trace']> | null = null,
    frontier: Array<{ sheet: string; cell: string }> = [],
    visited = new Set<string>(),
    queued = new Set<string>();
  let snapshot: FormulaTraceSnapshot = Object.freeze({
    scope: null,
    direction: null,
    phase: 'idle',
    depth: 0,
    edges: [],
    markers: [],
    requests: 0,
    queriedCells: 0,
    complete: false,
    canExpand: false,
    message: null,
  });
  const publish = (patch: Partial<FormulaTraceSnapshot>) => {
    snapshot = Object.freeze({ ...snapshot, ...patch });
    for (const fn of listeners) fn();
  };
  const scopeMatches = (scope?: FormulaTraceScope) =>
    !scope || (!!snapshot.scope && formulaTraceScopeKey(scope) === formulaTraceScopeKey(snapshot.scope));
  const reset = (scope?: FormulaTraceScope) => {
    frontier = [];
    visited = new Set();
    queued = new Set();
    publish({
      ...(scope ? { scope } : {}),
      direction: null,
      phase: 'idle',
      depth: 0,
      edges: [],
      markers: [],
      requests: 0,
      queriedCells: 0,
      complete: false,
      canExpand: false,
      message: null,
    });
  };
  const cancel = (scope?: FormulaTraceScope) => {
    if (disposed || !scopeMatches(scope)) return;
    generation++;
    abort?.abort();
    publish({ phase: 'idle', message: 'Trace cancelled. Previously shown arrows are kept.' });
  };
  const controller: FormulaTraceController = {
    getSnapshot: () => snapshot,
    subscribe: (fn) => {
      if (disposed) return () => {};
      listeners.add(fn);
      return () => listeners.delete(fn);
    },
    setScope: (scope) => {
      if (disposed) throw Error('Trace is closed.');
      validateFormulaTraceScope(scope);
      if (snapshot.scope && formulaTraceScopeKey(snapshot.scope) === formulaTraceScopeKey(scope)) return;
      generation++;
      abort?.abort();
      reset(Object.freeze({ ...scope, sheets: Object.freeze([...scope.sheets]) }));
    },
    cancel,
    clear: (scope) => {
      if (disposed || !scopeMatches(scope)) return;
      generation++;
      abort?.abort();
      reset();
    },
    dispose: () => {
      if (disposed) return;
      generation++;
      abort?.abort();
      disposed = true;
      publish({
        scope: null,
        direction: null,
        phase: 'idle',
        depth: 0,
        edges: [],
        markers: [],
        requests: 0,
        queriedCells: 0,
        complete: false,
        canExpand: false,
        message: null,
      });
      listeners.clear();
      frontier = [];
      visited.clear();
      queued.clear();
    },
    trace: (direction) => {
      if (disposed) return Promise.reject(Error('Trace is closed.'));
      if (!['precedents', 'dependents'].includes(direction)) return Promise.reject(Error('Invalid trace direction.'));
      if (pending) {
        if (snapshot.direction === direction) return pending;
        // A new explicit root/direction request retires the old query, then waits
        // for its physical read to settle. Duplicate latest requests share this tail.
        const previous = pending,
          bound = snapshot.scope;
        if (!bound?.active) return pending;
        cancel();
        reset();
        const key = formulaTraceScopeKey(bound),
          epoch = generation;
        frontier = [{ sheet: bound.sheet, cell: bound.cell }];
        queued.add(cellKey(frontier[0]));
        publish({ direction, canExpand: true, phase: 'reading' });
        let redirected: ReturnType<FormulaTraceController['trace']>;
        redirected = previous
          .then(() => {
            if (disposed || generation !== epoch || !snapshot.scope || formulaTraceScopeKey(snapshot.scope) !== key)
              return { outcome: 'retired' as const, depth: snapshot.depth };
            if (pending === redirected) pending = null;
            return controller.trace(direction);
          })
          .finally(() => {
            if (pending === redirected) pending = null;
          });
        pending = redirected;
        return redirected;
      }
      const scope = snapshot.scope;
      if (!scope?.active) return Promise.resolve({ outcome: 'paused', depth: snapshot.depth });
      if (snapshot.direction !== direction) {
        reset();
        frontier = [{ sheet: scope.sheet, cell: scope.cell }];
        queued.add(cellKey(frontier[0]));
        publish({ direction, canExpand: true });
      }
      if (!snapshot.canExpand)
        return Promise.resolve({ outcome: snapshot.complete ? 'completed' : 'limited', depth: snapshot.depth });
      const captured = formulaTraceScopeKey(scope),
        epoch = ++generation;
      abort = new AbortController();
      const signal = abort.signal;
      const current = () =>
        !disposed &&
        !signal.aborted &&
        generation === epoch &&
        !!snapshot.scope &&
        formulaTraceScopeKey(snapshot.scope) === captured &&
        snapshot.scope.active;
      const targets = frontier.slice(),
        workingVisited = new Set(visited),
        workingQueued = new Set(queued),
        edges = new Map(snapshot.edges.map((e) => [edgeKey(e), e])),
        markers = new Map(snapshot.markers.map((m) => [markerKey(m), m])),
        next: Array<{ sheet: string; cell: string }> = [];
      let requests = snapshot.requests,
        depth = snapshot.depth,
        limited = false;
      const addMarker = (m: FormulaTraceMarker) => {
        const key = markerKey(m);
        if (!markers.has(key) && markers.size < TRACE_LIMITS.edges)
          markers.set(key, Object.freeze({ ...m, id: 'marker-' + markers.size }));
      };
      const limit = (target: { sheet: string; cell: string }, reason: string) => {
        limited = true;
        addMarker({ ...target, kind: 'limit', reason });
      };
      const enqueue = (endpoint: TraceEndpoint) => {
        const b = traceBounds(endpoint.range),
          count = (b.last.row - b.first.row + 1) * (b.last.col - b.first.col + 1);
        if (count > TRACE_LIMITS.nodes) {
          limit(
            { sheet: endpoint.sheet, cell: endpoint.range.split(':')[0] },
            'This range is too large to expand. Its known relationship is still shown.'
          );
          return;
        }
        for (let row = b.first.row; row <= b.last.row; row++)
          for (let col = b.first.col; col <= b.last.col; col++) {
            const target = { sheet: endpoint.sheet, cell: column(col) + (row + 1) },
              key = cellKey(target);
            if (workingQueued.has(key)) continue;
            if (workingQueued.size >= TRACE_LIMITS.nodes) {
              limit(target, 'Trace reached its cell limit. Remove arrows and trace another cell.');
              return;
            }
            workingQueued.add(key);
            next.push(target);
          }
      };
      publish({ phase: 'reading', message: null });
      const work = async (): Promise<{ outcome: 'completed' | 'retired' | 'paused' | 'limited'; depth: number }> => {
        try {
          if (depth >= TRACE_LIMITS.depth) {
            limit(
              targets[0] ?? { sheet: scope.sheet, cell: scope.cell },
              'Trace reached 32 levels. Remove arrows and trace another cell.'
            );
          } else
            for (const target of targets) {
              if (!current()) return { outcome: 'retired', depth: snapshot.depth };
              let after: string | null = null;
              const cursors = new Set<string>();
              do {
                if (requests >= TRACE_LIMITS.requests) {
                  limit(target, 'Trace reached its read limit. Remove arrows and trace another cell.');
                  break;
                }
                const receipt = await query(scope, target, { direction, after, limit: TRACE_LIMITS.pageItems }, signal);
                requests++;
                if (!current()) return { outcome: 'retired', depth: snapshot.depth };
                if (
                  !receipt ||
                  formulaTraceScopeKey(receipt.scope) !== captured ||
                  receipt.target.sheet !== target.sheet ||
                  receipt.target.cell !== target.cell ||
                  receipt.direction !== direction ||
                  receipt.after !== after
                )
                  throw Error('Trace reply belongs to another workbook or target.');
                const page = validateFormulaTracePage(receipt.page, scope, target, direction);
                for (const edge of page.edges) {
                  const key = edgeKey(edge);
                  if (edges.has(key)) continue;
                  if (edges.size >= TRACE_LIMITS.edges) {
                    limit(target, 'Trace reached its arrow limit. Remove arrows and trace another cell.');
                    break;
                  }
                  edges.set(
                    key,
                    Object.freeze({
                      ...edge,
                      from: Object.freeze({ ...edge.from }),
                      to: Object.freeze({ ...edge.to }),
                      id: 'edge-' + edges.size,
                    })
                  );
                  enqueue(direction === 'precedents' ? edge.from : edge.to);
                }
                for (const marker of page.markers) addMarker(marker);
                if (page.coverage.limited && !page.nextCursor) limited = true;
                if (!page.coverage.complete && !page.nextCursor && !page.markers.length)
                  addMarker({
                    ...target,
                    kind: 'unsupported',
                    reason: 'Some relationships are not covered by this trace.',
                  });
                after = page.nextCursor;
                if (after) {
                  if (cursors.has(after)) throw Error('Trace continuation repeated.');
                  cursors.add(after);
                }
              } while (after && !limited);
              workingVisited.add(cellKey(target));
              if (limited) break;
            }
          if (!current()) return { outcome: 'retired', depth: snapshot.depth };
          visited = workingVisited;
          queued = workingQueued;
          frontier = next.filter((t) => !visited.has(cellKey(t)));
          depth++;
          const canExpand =
            !limited && frontier.length > 0 && depth < TRACE_LIMITS.depth && requests < TRACE_LIMITS.requests;
          if (frontier.length && !canExpand && !limited)
            limit(frontier[0], 'Trace reached its expansion limit. Remove arrows and trace another cell.');
          const complete = !frontier.length && !markers.size && !limited;
          publish({
            phase: 'idle',
            depth,
            edges: Object.freeze([...edges.values()]),
            markers: Object.freeze([...markers.values()]),
            requests,
            queriedCells: visited.size,
            complete,
            canExpand,
            message: limited ? 'Known arrows are shown; this trace reached a limit.' : null,
          });
          return { outcome: limited ? 'limited' : 'completed', depth };
        } catch (error) {
          if (!current()) return { outcome: 'retired', depth: snapshot.depth };
          publish({
            phase: 'failed',
            message: 'Relationships could not be read at this workbook state. Trace again to retry.',
          });
          throw error;
        } finally {
          if (generation === epoch) abort = null;
        }
      };
      const run = work().finally(() => {
        if (pending === run) pending = null;
      });
      pending = run;
      return run;
    },
  };
  return controller;
}
