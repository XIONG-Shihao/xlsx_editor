import type { WorkbookCell, WorkbookCommand } from '../core/contracts';
import {
  reviewStatistics,
  addReviewStatistics,
  classifyReviewCell,
  reviewTokens,
  captureSpellingFinding,
  planSpellingReplacement,
  validateReviewReceipt,
  WORKBOOK_REVIEW_LIMITS,
  type ReviewStatistics,
  type ReviewReceipt,
  type ReviewTokenCursor,
  type SpellingFinding,
} from '../core/workbook-review.mjs';
import type {
  OfflineSpellingAvailability,
  OfflineSpellingReply,
  OfflineSpellingRequest,
} from '../adapters/offline-spelling.mjs';
import type { ReviewOwnerReceipt, ReviewCellPage, ReviewCellRead } from './reviewScanAdapter';

export type { ReviewOwnerReceipt } from './reviewScanAdapter';
export interface WorkbookReviewScope extends ReviewOwnerReceipt {
  ownerIdentity: string;
  sheet: string;
  selection: string;
  activationEpoch: number;
  navigationGeneration: number;
  modeEpoch: number;
  active: boolean;
  writable: boolean;
  sheets: readonly { name: string; state: string }[];
}
export type WorkbookReviewScan = ReviewCellPage;
export type WorkbookReviewRead = ReviewCellRead;
export interface WorkbookReviewCallbacks {
  scan: (sheet: string, after: string | undefined, limit: number, signal: AbortSignal) => Promise<WorkbookReviewScan>;
  read: (sheet: string, cell: string, signal: AbortSignal) => Promise<WorkbookReviewRead>;
  availability: () => Promise<OfflineSpellingAvailability>;
  check: (request: OfflineSpellingRequest) => Promise<OfflineSpellingReply>;
  cancel: (requestId: string) => Promise<void>;
  apply: (commands: WorkbookCommand[]) => Promise<void>;
  navigate: (sheet: string, cell: string) => void;
}
export interface WorkbookReviewFrame {
  scope: WorkbookReviewScope;
  callbacks: WorkbookReviewCallbacks;
}
export type WorkbookReviewMode = 'statistics' | 'spelling';
export interface WorkbookSpellingFinding {
  id: string;
  scopeKey: string;
  finding: SpellingFinding;
  suggestions: readonly string[];
}
export interface WorkbookReviewSnapshot {
  scope: WorkbookReviewScope | null;
  mode: WorkbookReviewMode;
  locale: 'en-US' | 'en-GB';
  phase: 'idle' | 'scanning' | 'loading-provider' | 'checking' | 'reading-target' | 'accepting' | 'paused' | 'disposed';
  physicalBusy: boolean;
  statistics: ReviewStatistics | null;
  statisticsSheets: number;
  statisticsPages: number;
  statisticsComplete: boolean;
  findings: readonly WorkbookSpellingFinding[];
  spellingSheets: number;
  spellingRecords: number;
  spellingTextCells: number;
  wordsChecked: number;
  findingsFound: number;
  excludedTokens: number;
  excludedCells: number;
  spellingComplete: boolean;
  availability: OfflineSpellingAvailability | null;
  incompleteReason: string | null;
  error: string | null;
  notice: string | null;
}
export const WORKBOOK_REVIEW_UI_LIMITS = Object.freeze({
  pagesPerAction: 64,
  findingsPerPage: 100,
  tokenPagesPerAction: 64,
  scanCells: 256,
  scanBytes: 1048576,
});
const plain = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);
const text = (value: unknown, max = 512): value is string =>
  typeof value === 'string' && value.length > 0 && value.length <= max && !/[\x00-\x1f]/.test(value);
const bytes = (value: unknown): number => new TextEncoder().encode(JSON.stringify(value)).length;
const copy = <T>(value: T): T => structuredClone(value);
const canonical = (value: unknown): string =>
  JSON.stringify(value, (_key, item) =>
    plain(item)
      ? Object.fromEntries(
          Object.keys(item)
            .sort()
            .map((key) => [key, item[key]])
        )
      : item
  );
const owner = (scope: ReviewOwnerReceipt): ReviewOwnerReceipt => ({
  documentId: scope.documentId,
  ownerId: scope.ownerId,
  state: scope.state,
  revision: scope.revision,
});
function sameOwner(a: WorkbookReviewScope, b: WorkbookReviewScope): boolean {
  return a.documentId === b.documentId && a.ownerId === b.ownerId && a.ownerIdentity === b.ownerIdentity;
}
export function workbookReviewScopeKey(scope: WorkbookReviewScope): string {
  return canonical(scope);
}
function captureScope(scope: WorkbookReviewScope): WorkbookReviewScope {
  const fields = [
    'documentId',
    'ownerId',
    'state',
    'revision',
    'ownerIdentity',
    'sheet',
    'selection',
    'activationEpoch',
    'navigationGeneration',
    'modeEpoch',
    'active',
    'writable',
    'sheets',
  ];
  if (
    !plain(scope) ||
    Object.keys(scope).length !== fields.length ||
    fields.some((key) => !Object.hasOwn(scope, key)) ||
    ['documentId', 'ownerId', 'state', 'revision', 'ownerIdentity', 'sheet', 'selection'].some(
      (key) => !text(scope[key as keyof WorkbookReviewScope])
    ) ||
    ['activationEpoch', 'navigationGeneration', 'modeEpoch'].some(
      (key) =>
        !Number.isSafeInteger(scope[key as keyof WorkbookReviewScope]) ||
        Number(scope[key as keyof WorkbookReviewScope]) < 0
    ) ||
    typeof scope.active !== 'boolean' ||
    typeof scope.writable !== 'boolean' ||
    !Array.isArray(scope.sheets) ||
    !scope.sheets.length ||
    scope.sheets.some(
      (sheet) =>
        !plain(sheet) ||
        Object.keys(sheet).sort().join(',') !== 'name,state' ||
        !text(sheet.name, 31) ||
        /[\\/?*\[\]:]/.test(sheet.name) ||
        typeof sheet.state !== 'string' ||
        !['visible', 'hidden', 'veryHidden'].includes(sheet.state)
    ) ||
    new Set(scope.sheets.map((sheet) => sheet.name)).size !== scope.sheets.length ||
    !scope.sheets.some((sheet) => sheet.name === scope.sheet)
  )
    throw Error('Workbook Review needs a captured accepted owner and worksheet inventory.');
  return Object.freeze({ ...scope, sheets: Object.freeze(scope.sheets.map((sheet) => Object.freeze({ ...sheet }))) });
}
function position(reference: string): [number, number] {
  const match = typeof reference === 'string' && /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(reference);
  if (!match) throw Error('Invalid Review cell cursor.');
  const col = [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0),
    row = Number(match[2]);
  if (col > 16384 || row > 1048576) throw Error('Invalid Review cell cursor.');
  return [row, col];
}
const compare = (a: string, b: string): number => {
  const x = position(a),
    y = position(b);
  return x[0] - y[0] || x[1] - y[1];
};
function receiptMatches(actual: ReviewOwnerReceipt, scope: WorkbookReviewScope): void {
  if (
    !plain(actual) ||
    Object.keys(actual).sort().join(',') !== 'documentId,ownerId,revision,state' ||
    canonical(actual) !== canonical(owner(scope))
  )
    throw Error('The Review reply belongs to another accepted owner or head.');
}
class PageTooLarge extends Error {}
function scanPage(
  raw: WorkbookReviewScan,
  scope: WorkbookReviewScope,
  sheet: string,
  after: string | null,
  limit: number
): Array<[string, WorkbookCell]> {
  if (
    !plain(raw) ||
    Object.keys(raw).some((key) => !['after', 'cells', 'nextCursor', 'receipt', 'sheet'].includes(key)) ||
    !Object.hasOwn(raw, 'nextCursor') ||
    raw.sheet !== sheet ||
    (raw.after ?? null) !== after ||
    !plain(raw.cells)
  )
    throw Error('Invalid captured Review page.');
  receiptMatches(raw.receipt, scope);
  const references = Object.keys(raw.cells).sort(compare);
  if (
    references.length > limit ||
    references.some((ref) => after !== null && compare(ref, after) <= 0) ||
    references.some((ref) => !plain(raw.cells[ref]))
  )
    throw Error('Review page repeated a cell or exceeded its requested bound.');
  if (raw.nextCursor !== null && (!references.length || raw.nextCursor !== references.at(-1)))
    throw Error('Review page cursor did not advance to its last accepted record.');
  if (bytes(raw) > WORKBOOK_REVIEW_UI_LIMITS.scanBytes) throw new PageTooLarge('Review page exceeds its byte bound.');
  return references.map((ref) => [ref, raw.cells[ref]]);
}
function providerReceipt(
  availability: OfflineSpellingAvailability,
  scope: WorkbookReviewScope,
  locale: 'en-US' | 'en-GB'
): ReviewReceipt {
  if (
    !plain(availability) ||
    availability.schema !== 1 ||
    availability.provider !== 'nspell' ||
    availability.providerVersion !== '2.1.5' ||
    !availability.verifiedProvider ||
    canonical(availability.expectedProvider) !== canonical(availability.verifiedProvider) ||
    !Array.isArray(availability.locales)
  )
    throw Error('The offline spelling provider is not verified or available.');
  const dictionary = availability.locales.find((item) => item.locale === locale);
  if (
    !dictionary ||
    dictionary.status !== 'available' ||
    !dictionary.verified ||
    canonical(dictionary.expected) !== canonical(dictionary.verified) ||
    dictionary.verified.locale !== locale
  )
    throw Error(dictionary?.reason ?? 'The selected offline dictionary is unavailable.');
  return validateReviewReceipt({
    ...owner(scope),
    locale,
    providerId: availability.providerId,
    dictionaryId: dictionary.verified.dictionaryId,
  });
}
function spellingReply(reply: OfflineSpellingReply, request: OfflineSpellingRequest, receipt: ReviewReceipt): void {
  if (
    !plain(reply) ||
    reply.schema !== 1 ||
    reply.requestId !== request.requestId ||
    reply.locale !== receipt.locale ||
    reply.provider !== 'nspell' ||
    reply.providerVersion !== '2.1.5' ||
    reply.providerId !== receipt.providerId
  )
    throw Error('The spelling reply belongs to another request or provider.');
  if (reply.status !== 'completed') throw Error(reply.reason ?? 'Offline spelling did not complete.');
  if (
    reply.dictionaryId !== receipt.dictionaryId ||
    reply.dictionary?.dictionaryId !== receipt.dictionaryId ||
    reply.results.length !== request.words.length ||
    bytes(reply) > 1048576
  )
    throw Error('The spelling dictionary or result page changed.');
  for (const [index, result] of reply.results.entries())
    if (
      !plain(result) ||
      Object.keys(result).sort().join(',') !== 'correct,suggestions,word' ||
      result.word !== request.words[index] ||
      typeof result.correct !== 'boolean' ||
      !Array.isArray(result.suggestions) ||
      result.suggestions.length > 10 ||
      result.suggestions.some(
        (value) => typeof value !== 'string' || !value.length || Array.from(value).length > 256
      ) ||
      (result.correct && result.suggestions.length)
    )
      throw Error('Invalid typed spelling result.');
}
type Stream = {
  sheetIndex: number;
  after: string | null;
  buffer: Array<[string, WorkbookCell]>;
  index: number;
  next: string | null;
  token?: ReviewTokenCursor;
  started: boolean;
};
const stream = (): Stream => ({ sheetIndex: 0, after: null, buffer: [], index: 0, next: null, started: false });
export interface WorkbookReviewController {
  setFrame(frame: WorkbookReviewFrame | null): void;
  setMode(mode: WorkbookReviewMode): void;
  setLocale(locale: 'en-US' | 'en-GB'): void;
  getSnapshot(): WorkbookReviewSnapshot;
  subscribe(listener: () => void): () => void;
  scanStatistics(): Promise<void>;
  scanSpelling(): Promise<void>;
  changeWord(
    entry: WorkbookSpellingFinding,
    replacement: string
  ): Promise<{ outcome: 'accepted' | 'retired'; runtimeCurrent: boolean }>;
  navigate(entry: WorkbookSpellingFinding): void;
  restart(): void;
  cancel(): Promise<void>;
  dispose(): Promise<void>;
}
export function createWorkbookReviewController(): WorkbookReviewController {
  let frame: WorkbookReviewFrame | null = null,
    generation = 0,
    disposed = false,
    physical: Promise<unknown> | null = null,
    abort: AbortController | null = null,
    providerToken: string | null = null,
    cancelWork: Promise<void> | null = null;
  let mode: WorkbookReviewMode = 'statistics',
    locale: 'en-US' | 'en-GB' = 'en-US',
    stats: ReviewStatistics | null = null,
    statsSheet = 0,
    statsAfter: string | null = null,
    statsPages = 0,
    sp = stream(),
    spSheets = 0,
    spRecords = 0,
    spText = 0,
    words = 0,
    found = 0,
    excludedTokens = 0,
    excludedCells = 0,
    findings: WorkbookSpellingFinding[] = [],
    availability: OfflineSpellingAvailability | null = null,
    incomplete: string | null = null,
    error: string | null = null,
    notice: string | null = null;
  const listeners = new Set<() => void>();
  let phase: WorkbookReviewSnapshot['phase'] = 'paused',
    snapshot: WorkbookReviewSnapshot;
  const publish = (next = phase): void => {
    phase = next;
    snapshot = Object.freeze({
      scope: frame?.scope ?? null,
      mode,
      locale,
      phase,
      physicalBusy: !!physical,
      statistics: stats ? copy(stats) : null,
      statisticsSheets: statsSheet,
      statisticsPages: statsPages,
      statisticsComplete: !!frame && statsSheet === frame.scope.sheets.length,
      findings: Object.freeze([...findings]),
      spellingSheets: spSheets,
      spellingRecords: spRecords,
      spellingTextCells: spText,
      wordsChecked: words,
      findingsFound: found,
      excludedTokens,
      excludedCells,
      spellingComplete: !!frame && sp.sheetIndex === frame.scope.sheets.length,
      availability: availability ? copy(availability) : null,
      incompleteReason: incomplete,
      error,
      notice,
    });
    for (const listener of listeners)
      try {
        listener();
      } catch {
        /* View observers never authorize owner operations. */
      }
  };
  const resetSpelling = (): void => {
    sp = stream();
    spSheets = 0;
    spRecords = 0;
    spText = 0;
    words = 0;
    found = 0;
    excludedTokens = 0;
    excludedCells = 0;
    findings = [];
    availability = null;
  };
  const reset = (): void => {
    stats = frame ? reviewStatistics([], { revision: frame.scope.revision }) : null;
    statsSheet = 0;
    statsAfter = null;
    statsPages = 0;
    resetSpelling();
    incomplete = null;
    error = null;
    notice = null;
  };
  const invalidate = (): void => {
    generation++;
    abort?.abort();
    if (providerToken && frame) {
      const token = providerToken,
        callback = frame.callbacks.cancel;
      providerToken = null;
      cancelWork = Promise.resolve()
        .then(() => callback(token))
        .catch(() => {});
    }
  };
  const run = (
    operation: (captured: WorkbookReviewFrame, current: () => boolean, signal: AbortSignal) => Promise<unknown>
  ): Promise<unknown> => {
    if (disposed || !frame?.scope.active) return Promise.resolve();
    if (physical) return physical;
    const captured = frame,
      key = workbookReviewScopeKey(captured.scope),
      serial = generation;
    abort = new AbortController();
    const signal = abort.signal,
      current = () =>
        !disposed &&
        !signal.aborted &&
        generation === serial &&
        !!frame?.scope.active &&
        workbookReviewScopeKey(frame.scope) === key;
    const promise = Promise.resolve()
      .then(() => operation(captured, current, signal))
      .catch((failure) => {
        if (current()) {
          error = failure instanceof Error ? failure.message : 'Review could not complete.';
          incomplete = 'Scan incomplete. Retry or restart.';
          publish('idle');
        }
        throw failure;
      })
      .finally(async () => {
        await cancelWork;
        cancelWork = null;
        if (physical === promise) {
          physical = null;
          abort = null;
          providerToken = null;
          if (!disposed) publish(frame?.scope.active ? 'idle' : 'paused');
        }
      });
    physical = promise;
    publish();
    return promise;
  };
  const page = async (
    captured: WorkbookReviewFrame,
    sheet: string,
    after: string | null,
    current: () => boolean,
    signal: AbortSignal
  ): Promise<{ entries: Array<[string, WorkbookCell]>; next: string | null } | null> => {
    let limit: number = WORKBOOK_REVIEW_UI_LIMITS.scanCells;
    for (;;) {
      if (!current()) return null;
      try {
        const reply = await captured.callbacks.scan(sheet, after ?? undefined, limit, signal);
        if (!current()) return null;
        return { entries: scanPage(reply, captured.scope, sheet, after, limit), next: reply.nextCursor };
      } catch (failure) {
        if (!current()) return null;
        if (
          !(failure instanceof PageTooLarge) &&
          !(failure instanceof Error && failure.message === 'Native workbook reply exceeds transport bound')
        )
          throw failure;
        if (limit === 1) throw Error('One accepted record exceeds the inspection byte limit. This scan is incomplete.');
        limit = Math.max(1, Math.floor(limit / 2));
      }
    }
  };
  const advanceCell = (): void => {
    sp.after = sp.buffer[sp.index][0];
    sp.index++;
    sp.token = undefined;
    sp.started = false;
    if (sp.index === sp.buffer.length) {
      if (sp.next === null) {
        sp.sheetIndex++;
        spSheets++;
        sp.after = null;
      } else sp.after = sp.next;
      sp.buffer = [];
      sp.index = 0;
      sp.next = null;
    }
  };
  const entryCurrent = (entry: WorkbookSpellingFinding): boolean =>
    !!frame?.scope.active &&
    entry.scopeKey === workbookReviewScopeKey(frame.scope) &&
    findings.some((value) => value.id === entry.id) &&
    entry.finding.receipt.locale === locale;
  const api: WorkbookReviewController = {
    setFrame(value) {
      if (disposed) throw Error('Workbook Review is closed.');
      const next = value ? { scope: captureScope(value.scope), callbacks: value.callbacks } : null;
      if (frame && next && workbookReviewScopeKey(frame.scope) === workbookReviewScopeKey(next.scope)) {
        frame = next;
        return;
      }
      invalidate();
      frame = next;
      reset();
      publish(frame?.scope.active ? 'idle' : 'paused');
    },
    setMode(value) {
      if (!['statistics', 'spelling'].includes(value)) throw Error('Invalid Review mode.');
      if (mode === value || phase === 'accepting') return;
      invalidate();
      mode = value;
      error = null;
      notice = null;
      incomplete = null;
      publish(frame?.scope.active ? 'idle' : 'paused');
    },
    setLocale(value) {
      if (!['en-US', 'en-GB'].includes(value)) throw Error('Select an installed spelling locale.');
      if (locale === value || phase === 'accepting') return;
      invalidate();
      locale = value;
      resetSpelling();
      error = null;
      notice = null;
      incomplete = null;
      publish(frame?.scope.active ? 'idle' : 'paused');
    },
    getSnapshot: () => snapshot,
    subscribe(listener) {
      if (disposed) return () => {};
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    scanStatistics() {
      if (mode !== 'statistics') return Promise.resolve();
      return run(async (captured, current, signal) => {
        if (!current()) return;
        error = null;
        notice = null;
        incomplete = null;
        publish('scanning');
        let pages = 0;
        while (statsSheet < captured.scope.sheets.length && pages < 64) {
          const reply = await page(captured, captured.scope.sheets[statsSheet].name, statsAfter, current, signal);
          if (!reply) return;
          stats = addReviewStatistics(
            stats!,
            reply.entries.map(([, cell]) => cell),
            { revision: captured.scope.revision }
          );
          statsPages++;
          pages++;
          if (reply.next === null) {
            statsSheet++;
            statsAfter = null;
          } else statsAfter = reply.next;
          publish('scanning');
        }
        if (current()) {
          incomplete =
            statsSheet < captured.scope.sheets.length
              ? 'This step reached 64 accepted-cell pages. Continue to inspect the remaining records.'
              : null;
          publish('idle');
        }
      }).then(() => {});
    },
    scanSpelling() {
      if (mode !== 'spelling') return Promise.resolve();
      return run(async (captured, current, signal) => {
        if (!current()) return;
        error = null;
        notice = null;
        incomplete = null;
        publish('loading-provider');
        const loaded = await captured.callbacks.availability();
        if (!current()) return;
        availability = loaded;
        const receipt = providerReceipt(loaded, captured.scope, locale);
        findings = [];
        publish('scanning');
        let pages = 0,
          tokenPages = 0;
        while (sp.sheetIndex < captured.scope.sheets.length && findings.length < 100 && pages < 64 && tokenPages < 64) {
          const sheet = captured.scope.sheets[sp.sheetIndex].name;
          if (!sp.buffer.length) {
            const reply = await page(captured, sheet, sp.after, current, signal);
            if (!reply) return;
            pages++;
            if (!reply.entries.length) {
              sp.sheetIndex++;
              spSheets++;
              sp.after = null;
              continue;
            }
            sp.buffer = reply.entries;
            sp.index = 0;
            sp.next = reply.next;
          }
          const [cellRef, cell] = sp.buffer[sp.index],
            classification = classifyReviewCell(cell, { revision: captured.scope.revision });
          if (!sp.started) {
            spRecords++;
            sp.started = true;
            if (classification.spellingEligible) spText++;
            else excludedCells++;
          }
          if (!classification.spellingEligible) {
            advanceCell();
            continue;
          }
          const tokenPage = reviewTokens(cell.value as string, sp.token, {
            maxTokens: Math.max(1, 100 - findings.length),
          });
          tokenPages++;
          if (tokenPage.tokens.length) {
            const request: OfflineSpellingRequest = {
              requestId: crypto.randomUUID(),
              locale,
              words: tokenPage.tokens.map((token) => token.text),
            };
            providerToken = request.requestId;
            publish('checking');
            const result = await captured.callbacks.check(request);
            providerToken = null;
            if (!current()) return;
            spellingReply(result, request, receipt);
            if (result.status !== 'completed') return;
            for (const [index, word] of result.results.entries())
              if (!word.correct) {
                const finding = captureSpellingFinding({
                    receipt,
                    target: { sheet, cell: cellRef },
                    cell,
                    token: tokenPage.tokens[index],
                  }),
                  scopeKey = workbookReviewScopeKey(captured.scope),
                  id = canonical([scopeKey, finding]);
                findings.push(
                  Object.freeze({
                    id,
                    scopeKey,
                    finding: Object.freeze(finding),
                    suggestions: Object.freeze([...word.suggestions]),
                  })
                );
                found++;
              }
            words += result.results.length;
          }
          if (!current()) return;
          excludedTokens += tokenPage.excluded.length;
          sp.token = tokenPage.nextCursor ?? undefined;
          if (tokenPage.nextCursor === null) advanceCell();
          publish('scanning');
        }
        if (current()) {
          incomplete =
            sp.sheetIndex === captured.scope.sheets.length
              ? null
              : findings.length === 100
                ? 'This page shows 100 findings. Continue at the exact next word.'
                : 'This step reached its page or text-work bound. Continue the review.';
          publish('idle');
        }
      }).then(() => {});
    },
    changeWord(entry, replacement) {
      if (disposed || !frame || !entryCurrent(entry) || !frame.scope.writable || physical)
        return Promise.resolve({ outcome: 'retired', runtimeCurrent: false });
      let outcome: { outcome: 'accepted' | 'retired'; runtimeCurrent: boolean } = {
        outcome: 'retired',
        runtimeCurrent: false,
      };
      return run(async (captured, current, signal) => {
        if (!current() || !entryCurrent(entry)) return;
        publish('reading-target');
        const raw = await captured.callbacks.read(entry.finding.target.sheet, entry.finding.target.cell, signal);
        if (!current() || !entryCurrent(entry)) return;
        receiptMatches(raw.receipt, captured.scope);
        if (raw.sheet !== entry.finding.target.sheet || raw.cell !== entry.finding.target.cell || !raw.value)
          throw Error('The original spelling target is unavailable.');
        const verified = providerReceipt(availability!, captured.scope, locale);
        const command = planSpellingReplacement(entry.finding, {
          receipt: verified,
          cell: raw.value,
          replacement,
          writable: captured.scope.writable,
        });
        if (!current()) return;
        publish('accepting');
        await captured.callbacks.apply([command]);
        outcome = {
          outcome: 'accepted',
          runtimeCurrent: !!frame?.scope.active && sameOwner(frame.scope, captured.scope),
        };
        if (outcome.runtimeCurrent) {
          notice = 'The word was changed. One Undo restores the original text. Scan again at the new accepted head.';
          findings = [];
          publish('idle');
        }
      }).then(() => outcome);
    },
    navigate(entry) {
      if (!entryCurrent(entry) || !frame) return;
      frame.callbacks.navigate(entry.finding.target.sheet, entry.finding.target.cell);
    },
    restart() {
      if (disposed || phase === 'accepting') return;
      invalidate();
      reset();
      publish(frame?.scope.active ? 'idle' : 'paused');
    },
    async cancel() {
      if (disposed || phase === 'accepting') return;
      invalidate();
      incomplete = 'Review cancelled. Counts and findings already shown are incomplete.';
      error = null;
      publish(frame?.scope.active ? 'idle' : 'paused');
      await cancelWork;
    },
    async dispose() {
      if (disposed) return;
      invalidate();
      disposed = true;
      frame = null;
      reset();
      publish('disposed');
      listeners.clear();
      await cancelWork;
      await physical?.catch(() => {});
    },
  };
  publish();
  return Object.freeze(api);
}
