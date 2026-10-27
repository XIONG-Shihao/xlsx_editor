import type {
  WorkbookUsedRangeInner,
  WorkbookUsedRangeReceipt,
  WorkbookUsedRangeResult,
  WorkbookUsedRangeNativeIdentity,
} from './contracts.js';
export function usedRangeSheet(sheet: unknown): string;
export function usedRangeFields<T>(request: T, options?: { transport?: boolean }): T;
export function usedRangeReceipt(receipt: unknown): WorkbookUsedRangeReceipt;
export function usedRangeInner(
  reply: unknown,
  expected?: { sheet?: string; native?: WorkbookUsedRangeNativeIdentity }
): WorkbookUsedRangeInner;
export function normalizeNativeUsedRange(
  reply: unknown,
  expected: { sheet: string; native: WorkbookUsedRangeNativeIdentity }
): WorkbookUsedRangeInner;
export function bindUsedRange(
  reply: unknown,
  receipt: WorkbookUsedRangeReceipt,
  expected?: { sheet?: string; native?: WorkbookUsedRangeNativeIdentity }
): WorkbookUsedRangeResult;
export function validateUsedRange(
  reply: unknown,
  expected?: { sheet?: string; native?: WorkbookUsedRangeNativeIdentity; receipt?: WorkbookUsedRangeReceipt }
): WorkbookUsedRangeResult;
