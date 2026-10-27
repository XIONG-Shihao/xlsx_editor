/** Captured readonly owner pages shared by the standalone and embedded Review. */
import type { WorkbookCell } from '../core/contracts';
import { validateErrorScanPage } from './errorScanAdapter';

export interface ReviewOwnerReceipt {
  documentId: string;
  ownerId: string;
  state: string;
  revision: string;
}
export interface ReviewOwnerFrame {
  receipt: ReviewOwnerReceipt;
  identity: string;
  active: boolean;
  sheets: readonly string[];
  scan: (sheet: string, after: string | undefined, limit: number) => Promise<unknown>;
  read: (sheet: string, cell: string) => Promise<unknown>;
}
export interface ReviewCellPage {
  receipt: ReviewOwnerReceipt;
  sheet: string;
  after?: string;
  cells: Record<string, WorkbookCell>;
  nextCursor: string | null;
}
export interface ReviewCellRead {
  receipt: ReviewOwnerReceipt;
  sheet: string;
  cell: string;
  value: WorkbookCell | undefined;
}
const receiptKey = (r: ReviewOwnerReceipt) => JSON.stringify([r.documentId, r.ownerId, r.state, r.revision]);
const validReceipt = (r: ReviewOwnerReceipt) =>
  !!r &&
  [r.documentId, r.ownerId, r.state, r.revision].every(
    (value) => typeof value === 'string' && value.length > 0 && value.length <= 256
  );
function coordinate(ref: string): [number, number] {
  const match = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(ref);
  if (!match) throw Error('Invalid Review cell coordinate.');
  const col = [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0),
    row = Number(match[2]);
  if (col > 16384 || row > 1048576) throw Error('Invalid Review cell coordinate.');
  return [row, col];
}
const afterCoordinate = (ref: string, after: string) => {
  const a = coordinate(ref),
    b = coordinate(after);
  return a[0] > b[0] || (a[0] === b[0] && a[1] > b[1]);
};

export function createScopedReviewReader(current: () => ReviewOwnerFrame | null) {
  function capture(
    receipt: ReviewOwnerReceipt,
    identity: string,
    sheet: string,
    signal: AbortSignal
  ): ReviewOwnerFrame {
    const frame = current();
    if (
      !validReceipt(receipt) ||
      signal.aborted ||
      !frame?.active ||
      !validReceipt(frame.receipt) ||
      frame.identity !== identity ||
      receiptKey(frame.receipt) !== receiptKey(receipt) ||
      !frame.sheets.includes(sheet)
    )
      throw Error('The Review workbook or worksheet changed.');
    return frame;
  }
  return {
    async scan(
      receipt: ReviewOwnerReceipt,
      identity: string,
      sheet: string,
      after: string | undefined,
      limit: number,
      signal: AbortSignal
    ): Promise<ReviewCellPage> {
      if (!Number.isInteger(limit) || limit < 1 || limit > 2048) throw Error('Invalid Review scan page size.');
      if (after !== undefined) coordinate(after);
      const frame = capture(receipt, identity, sheet, signal),
        result = await frame.scan(sheet, after, limit);
      capture(receipt, identity, sheet, signal);
      const page = validateErrorScanPage(result),
        references = Object.keys(page.cells);
      if (
        references.length > limit ||
        (after !== undefined && references.some((ref) => !afterCoordinate(ref, after))) ||
        (page.nextCursor !== null &&
          ((after !== undefined && !afterCoordinate(page.nextCursor, after)) ||
            references.some((ref) => afterCoordinate(ref, page.nextCursor!))))
      )
        throw Error('Review scan page did not advance within its bound.');
      return { receipt: { ...receipt }, sheet, ...(after === undefined ? {} : { after }), ...page };
    },
    async read(
      receipt: ReviewOwnerReceipt,
      identity: string,
      sheet: string,
      cell: string,
      signal: AbortSignal
    ): Promise<ReviewCellRead> {
      coordinate(cell);
      const frame = capture(receipt, identity, sheet, signal),
        raw = await frame.read(sheet, cell);
      capture(receipt, identity, sheet, signal);
      if (!raw || typeof raw !== 'object' || Array.isArray(raw))
        throw Error('The Review cell read returned no valid result.');
      const cells = (raw as { cells?: unknown }).cells;
      if (!cells || typeof cells !== 'object' || Array.isArray(cells) || Object.keys(cells).some((ref) => ref !== cell))
        throw Error('The Review cell read returned unrelated cells.');
      const value = (cells as Record<string, WorkbookCell>)[cell];
      if (value !== undefined && (!value || typeof value !== 'object' || Array.isArray(value)))
        throw Error('The Review cell read returned an invalid cell.');
      return { receipt: { ...receipt }, sheet, cell, value };
    },
  };
}
