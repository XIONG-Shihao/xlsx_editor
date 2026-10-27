import {
  formulaTraceScopeKey,
  validateFormulaTracePage,
  type FormulaTraceScope,
  type FormulaTracePage,
  type FormulaTraceQuery,
  type TraceDirection,
} from './formulaTrace';
/** Actual backend receipt fields stay separate from the opaque UI state token. */
export interface FormulaTraceReceipt {
  documentId: string;
  ownerId: string;
  state: string;
  revision: string;
}
export interface FormulaTraceFrame {
  scope: FormulaTraceScope;
  receipt: FormulaTraceReceipt;
  trace: (
    sheet: string,
    cell: string,
    direction: TraceDirection,
    after?: string,
    limit?: number
  ) => Promise<FormulaTracePage>;
}
const receiptKey = (r: FormulaTraceReceipt) => JSON.stringify([r.documentId, r.ownerId, r.state, r.revision]);
export function createScopedFormulaTraceQuery(current: () => FormulaTraceFrame | null): FormulaTraceQuery {
  return async (scope, target, options, signal) => {
    const frame = current(),
      key = formulaTraceScopeKey(scope);
    if (
      signal.aborted ||
      !scope.active ||
      !frame ||
      formulaTraceScopeKey(frame.scope) !== key ||
      !scope.sheets.includes(target.sheet)
    )
      throw Error('The trace workbook or accepted state changed.');
    const capturedReceipt = receiptKey(frame.receipt);
    // Capture the callback before await. It owns the real grant/session/head;
    // neither an updated callback nor a parsed UI token can retarget this query.
    const raw = await frame.trace(
        target.sheet,
        target.cell,
        options.direction,
        options.after ?? undefined,
        options.limit
      ),
      latest = current();
    if (
      signal.aborted ||
      !latest ||
      formulaTraceScopeKey(latest.scope) !== key ||
      receiptKey(latest.receipt) !== capturedReceipt
    )
      throw Error('The trace workbook or accepted state changed.');
    const page = validateFormulaTracePage(raw, scope, target, options.direction);
    if (
      page.documentId !== frame.receipt.documentId ||
      page.ownerId !== frame.receipt.ownerId ||
      page.state !== frame.receipt.state ||
      page.revision !== frame.receipt.revision
    )
      throw Error('Trace returned another document, owner or accepted state.');
    return { scope, target, direction: options.direction, after: options.after, page };
  };
}
