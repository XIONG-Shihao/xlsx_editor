import type { WorkbookCell } from '../core/contracts';
export const ERROR_SCAN_LIMITS = Object.freeze({ pages: 64, issues: 100, cellsPerNativePage: 2048 });
export interface ErrorScanScope {
  readonly documentId: string;
  readonly ownerIdentity: string;
  readonly state: string;
  readonly sheet: string;
  readonly activationEpoch: number;
  readonly active: boolean;
}
export interface CellIssue {
  readonly reference: string;
  readonly kind: 'error' | 'circular' | 'unresolved' | 'stored';
  readonly formula?: string;
  readonly reason: string;
}
export interface ErrorScanReply {
  readonly scope: ErrorScanScope;
  readonly after: string | null;
  readonly cells: Record<string, WorkbookCell>;
  readonly nextCursor: string | null;
}
export type ErrorScanner = (
  scope: ErrorScanScope,
  after: string | null,
  signal: AbortSignal
) => Promise<ErrorScanReply>;
export interface ErrorScanSnapshot {
  readonly scope: ErrorScanScope | null;
  readonly phase: 'idle' | 'scanning' | 'paused' | 'disposed';
  readonly issues: readonly CellIssue[];
  readonly cursor: string | null;
  readonly complete: boolean;
  readonly incompleteReason: 'issue-limit' | 'page-limit' | 'cancelled' | 'scan-failed' | null;
  readonly processedCells: number;
  readonly nativePages: number;
  readonly issuesFound: number;
  readonly failure?: string;
}
export interface ErrorScanResult {
  readonly outcome: 'page' | 'retired' | 'paused';
  readonly issues: number;
  readonly processedCells: number;
  readonly nativePages: number;
  readonly complete: boolean;
  readonly limit: 'issue-limit' | 'page-limit' | null;
}
const fields = ['documentId', 'ownerIdentity', 'state', 'sheet', 'activationEpoch', 'active'];
function string(value: unknown, max: number): value is string {
  return typeof value === 'string' && value.length > 0 && value.length <= max && !/[\x00-\x1f]/.test(value);
}
function capture(scope: ErrorScanScope): ErrorScanScope {
  if (
    !scope ||
    Object.keys(scope).some((key) => !fields.includes(key)) ||
    !string(scope.documentId, 512) ||
    !string(scope.ownerIdentity, 512) ||
    !string(scope.state, 512) ||
    !string(scope.sheet, 31) ||
    /[\\/?*\[\]:]/.test(scope.sheet) ||
    !Number.isSafeInteger(scope.activationEpoch) ||
    scope.activationEpoch < 0 ||
    typeof scope.active !== 'boolean'
  )
    throw Error('Invalid error-checking scope.');
  return Object.freeze({ ...scope });
}
export const errorScanScopeKey = (scope: ErrorScanScope): string =>
  JSON.stringify([
    scope.documentId,
    scope.ownerIdentity,
    scope.state,
    scope.sheet,
    scope.activationEpoch,
    scope.active,
  ]);
function position(reference: string): readonly [number, number] {
  const match = typeof reference === 'string' && /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(reference);
  if (!match) throw Error('Invalid error-checking cursor.');
  const col = [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0),
    row = Number(match[2]);
  if (col > 16384 || row > 1048576) throw Error('Invalid error-checking cursor.');
  return [row, col];
}
function compare(a: string, b: string): number {
  const x = position(a),
    y = position(b);
  return x[0] - y[0] || x[1] - y[1];
}
function bounded(value: unknown, max: number, label: string): string {
  if (typeof value !== 'string' || value.length > max) throw Error('Unqualified ' + label + ' in owner reply.');
  return value;
}
/** Error-looking text never becomes an error. Unqualified formula caches are
 * separate issues and their cached value is deliberately not published. */
export function classifyCellIssue(reference: string, cell: WorkbookCell | undefined): CellIssue | null {
  position(reference);
  if (!cell) return null;
  const formula = cell.formula !== undefined ? bounded(cell.formula, 32768, 'formula') : undefined,
    calculation = cell.calculation;
  if (formula !== undefined || cell.formulaType) {
    if (calculation?.status === 'error')
      return { reference, formula, kind: 'error', reason: bounded(calculation.error, 128, 'typed error') };
    if (calculation?.status === 'circular')
      return {
        reference,
        formula,
        kind: 'circular',
        reason: bounded(calculation.error ?? 'Circular formula reported by the owner.', 512, 'circular reason'),
      };
    if (calculation?.status === 'unresolved')
      return {
        reference,
        formula,
        kind: 'unresolved',
        reason: bounded(
          calculation.error ?? 'Formula cannot be qualified at this accepted head.',
          512,
          'unresolved reason'
        ),
      };
    if (calculation?.status === 'calculated')
      return !['string', 'number', 'boolean'].includes(typeof calculation.value) ||
        (typeof calculation.value === 'number' && !Number.isFinite(calculation.value))
        ? { reference, formula, kind: 'unresolved', reason: 'Qualified finite scalar result is missing.' }
        : null;
    return {
      reference,
      formula,
      kind: 'stored',
      reason: 'Stored formula cache has not been qualified at this accepted head.',
    };
  }
  if (cell.dataType === 'error')
    return { reference, kind: 'error', reason: bounded(cell.error ?? cell.value, 128, 'typed error') };
  return null;
}
export interface ErrorCheckingController {
  setScope(scope: ErrorScanScope | null): void;
  getSnapshot(): ErrorScanSnapshot;
  subscribe(listener: () => void): () => void;
  scanNext(): Promise<ErrorScanResult>;
  restart(): void;
  cancel(): void;
  dispose(): void;
}
export function createErrorCheckingController(
  scanner: ErrorScanner,
  limits: { pages?: number; issues?: number } = {}
): ErrorCheckingController {
  if (typeof scanner !== 'function' || Object.keys(limits).some((key) => !['pages', 'issues'].includes(key)))
    throw Error('Invalid bounded scanner.');
  const pages = limits.pages ?? ERROR_SCAN_LIMITS.pages,
    issueLimit = limits.issues ?? ERROR_SCAN_LIMITS.issues;
  if (
    !Number.isSafeInteger(pages) ||
    pages < 1 ||
    pages > 64 ||
    !Number.isSafeInteger(issueLimit) ||
    issueLimit < 1 ||
    issueLimit > 100
  )
    throw Error('Error Checking is limited to64 owner pages and100 visible issues per action.');
  let scan: ErrorScanner | undefined = scanner,
    scope: ErrorScanScope | null = null,
    generation = 0,
    disposed = false,
    abort: AbortController | undefined,
    pending: Promise<ErrorScanResult> | undefined;
  let issues: CellIssue[] = [],
    cursor: string | null = null,
    complete = false,
    totalProcessed = 0,
    totalPages = 0,
    totalIssues = 0,
    incompleteReason: ErrorScanSnapshot['incompleteReason'] = null,
    failure: string | undefined;
  let snapshot: ErrorScanSnapshot = Object.freeze({
    scope: null,
    phase: 'paused',
    issues: Object.freeze([]),
    cursor: null,
    complete: false,
    incompleteReason: null,
    processedCells: 0,
    nativePages: 0,
    issuesFound: 0,
  });
  const listeners = new Set<() => void>();
  function publish(phase: ErrorScanSnapshot['phase']): void {
    snapshot = Object.freeze({
      scope,
      phase,
      issues: Object.freeze(issues.map((item) => Object.freeze({ ...item }))),
      cursor,
      complete,
      incompleteReason,
      processedCells: totalProcessed,
      nativePages: totalPages,
      issuesFound: totalIssues,
      ...(failure ? { failure } : {}),
    });
    for (const listener of listeners)
      try {
        listener();
      } catch {
        /* Observers cannot change owner-read acceptance. */
      }
  }
  function invalidate(): void {
    generation++;
    abort?.abort();
  }
  function reset(): void {
    issues = [];
    cursor = null;
    complete = false;
    totalProcessed = 0;
    totalPages = 0;
    totalIssues = 0;
    incompleteReason = null;
    failure = undefined;
  }
  const api: ErrorCheckingController = {
    setScope(value) {
      if (disposed) throw Error('Error Checking is closed.');
      const next = value === null ? null : capture(value);
      if (scope && next && errorScanScopeKey(scope) === errorScanScopeKey(next)) return;
      invalidate();
      scope = next;
      reset();
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
    scanNext() {
      if (disposed || !scope?.active)
        return Promise.resolve({
          outcome: 'paused',
          issues: 0,
          processedCells: 0,
          nativePages: 0,
          complete: false,
          limit: null,
        });
      // Never overlap physical owner reads, even when an old read ignores abort.
      if (pending) return pending;
      if (complete)
        return Promise.resolve({
          outcome: 'page',
          issues: 0,
          processedCells: 0,
          nativePages: 0,
          complete: true,
          limit: null,
        });
      const captured = scope,
        capturedGeneration = generation,
        capturedScan = scan!;
      abort = new AbortController();
      const signal = abort.signal,
        current = () =>
          !disposed &&
          !signal.aborted &&
          generation === capturedGeneration &&
          !!scope &&
          errorScanScopeKey(scope) === errorScanScopeKey(captured);
      let after = cursor,
        processed = 0,
        readPages = 0,
        reachedEnd = false,
        limit: ErrorScanResult['limit'] = null;
      const found: CellIssue[] = [];
      const work = Promise.resolve()
        .then(async (): Promise<ErrorScanResult> => {
          while (readPages < pages) {
            if (!current())
              return {
                outcome: 'retired',
                issues: 0,
                processedCells: processed,
                nativePages: readPages,
                complete: false,
                limit: null,
              };
            const reply = await capturedScan(captured, after, signal);
            readPages++;
            if (!current())
              return {
                outcome: 'retired',
                issues: 0,
                processedCells: processed,
                nativePages: readPages,
                complete: false,
                limit: null,
              };
            if (
              !reply ||
              errorScanScopeKey(capture(reply.scope)) !== errorScanScopeKey(captured) ||
              reply.after !== after ||
              !reply.cells ||
              typeof reply.cells !== 'object' ||
              Array.isArray(reply.cells) ||
              !Object.hasOwn(reply, 'nextCursor')
            )
              throw Error('Error Checking reply belongs to another owner, head, sheet or page.');
            const references = Object.keys(reply.cells).sort(compare);
            if (references.length > 2048) throw Error('Owner scan exceeded2048 stored cells.');
            if (after && references.some((ref) => compare(ref, after!) <= 0))
              throw Error('Owner scan repeated an earlier cell.');
            if (reply.nextCursor !== null) {
              position(reply.nextCursor);
              if (!references.length || reply.nextCursor !== references.at(-1))
                throw Error('Owner scan cursor did not advance to its actual last cell.');
            }
            let last = after;
            for (let index = 0; index < references.length; index++) {
              const reference = references[index],
                issue = classifyCellIssue(reference, reply.cells[reference]);
              processed++;
              last = reference;
              if (issue) found.push(issue);
              if (found.length === issueLimit) {
                after = last;
                reachedEnd = index === references.length - 1 && reply.nextCursor === null;
                limit = reachedEnd ? null : 'issue-limit';
                break;
              }
            }
            if (found.length === issueLimit) break;
            after = last;
            if (reply.nextCursor === null) {
              reachedEnd = true;
              break;
            }
            after = reply.nextCursor;
          }
          if (!reachedEnd && limit === null) limit = 'page-limit';
          if (!current())
            return {
              outcome: 'retired',
              issues: 0,
              processedCells: processed,
              nativePages: readPages,
              complete: false,
              limit: null,
            };
          issues = found;
          cursor = after;
          complete = reachedEnd;
          totalProcessed += processed;
          totalPages += readPages;
          totalIssues += found.length;
          incompleteReason = limit;
          failure = undefined;
          publish('idle');
          return {
            outcome: 'page',
            issues: found.length,
            processedCells: processed,
            nativePages: readPages,
            complete: reachedEnd,
            limit,
          };
        })
        .catch((error) => {
          if (!current())
            return {
              outcome: 'retired' as const,
              issues: 0,
              processedCells: processed,
              nativePages: readPages,
              complete: false,
              limit: null,
            };
          incompleteReason = 'scan-failed';
          failure = String(error instanceof Error ? error.message : error).slice(0, 512);
          publish('idle');
          throw error;
        });
      const final = work.finally(() => {
        if (pending === final) {
          pending = undefined;
          abort = undefined;
        }
      });
      pending = final;
      publish('scanning');
      return final;
    },
    restart() {
      if (disposed) return;
      invalidate();
      reset();
      publish(scope?.active ? 'idle' : 'paused');
    },
    cancel() {
      if (disposed) return;
      invalidate();
      incompleteReason = 'cancelled';
      failure = undefined;
      publish(scope?.active ? 'idle' : 'paused');
    },
    dispose() {
      if (disposed) return;
      disposed = true;
      invalidate();
      reset();
      scope = null;
      scan = undefined;
      publish('disposed');
      listeners.clear();
    },
  };
  return Object.freeze(api);
}
