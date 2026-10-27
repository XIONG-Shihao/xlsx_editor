/** Dictionary callbacks retain the original host owner, including cancellation. */
import type { WorkbookSpellingRequest } from '../adapters/workbook-spelling-jobs.mjs';
import type {
  OfflineSpellingAvailability,
  OfflineSpellingReply,
  OfflineSpellingRequest,
} from '../adapters/offline-spelling.mjs';
import type { ReviewOwnerReceipt } from './reviewScanAdapter';

export interface WorkbookReviewControls {
  receipt: ReviewOwnerReceipt;
  availability: () => Promise<OfflineSpellingAvailability>;
  check: (request: OfflineSpellingRequest) => Promise<OfflineSpellingReply>;
  cancel: (requestId: string) => Promise<void>;
}
const plain = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);
const sameReceipt = (raw: unknown, receipt: ReviewOwnerReceipt): boolean =>
  plain(raw) &&
  Object.keys(raw).sort().join(',') === 'documentId,ownerId,revision,state' &&
  Object.entries(receipt).every(([key, value]) => raw[key] === value);
function envelope(raw: unknown, key: string): Record<string, unknown> {
  if (!plain(raw) || !plain(raw[key]) || raw[key].schema !== 1)
    throw Error('The spelling host returned no valid receipt.');
  return raw[key];
}
export function createWorkbookReviewControls(
  receipt: ReviewOwnerReceipt,
  invoke: (request: WorkbookSpellingRequest) => Promise<unknown>
): WorkbookReviewControls {
  if (
    !sameReceipt(receipt, receipt) ||
    Object.values(receipt).some((value) => typeof value !== 'string' || !value.length || value.length > 256)
  )
    throw Error('Review needs an accepted host owner receipt.');
  const captured = Object.freeze({ ...receipt });
  return Object.freeze({
    receipt: captured,
    async availability() {
      const requestId = crypto.randomUUID(),
        reply = envelope(
          await invoke({ operation: 'spellingAvailability', ...captured, requestId }),
          'spellingAvailability'
        );
      if (reply.requestId !== requestId || !sameReceipt(reply.receipt, captured) || !plain(reply.availability))
        throw Error('Spelling availability belongs to another workbook or request.');
      return reply.availability as unknown as OfflineSpellingAvailability;
    },
    async check(request: OfflineSpellingRequest) {
      const reply = envelope(
        await invoke({
          operation: 'spellingCheck',
          ...captured,
          requestId: request.requestId,
          locale: request.locale,
          words: [...request.words],
        }),
        'spellingCheck'
      );
      if (
        !sameReceipt(reply.receipt, captured) ||
        !plain(reply.result) ||
        reply.result.requestId !== request.requestId ||
        reply.result.locale !== request.locale
      )
        throw Error('Spelling result belongs to another workbook or request.');
      return reply.result as unknown as OfflineSpellingReply;
    },
    async cancel(requestId: string) {
      const reply = envelope(
        await invoke({ operation: 'spellingCancel', ...captured, requestId }),
        'spellingCancellation'
      );
      if (reply.requestId !== requestId || typeof reply.cancelled !== 'boolean')
        throw Error('Spelling cancellation belongs to another request.');
    },
  });
}
