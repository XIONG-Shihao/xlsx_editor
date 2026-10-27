import type {
  WorkbookNoteMembershipInner,
  WorkbookNoteMembershipReceipt,
  WorkbookNoteMembershipResult,
  WorkbookNoteMembershipSource,
} from './contracts.js';
export const NOTE_MEMBERSHIP_MAX_ANNOTATIONS: number;
export const NOTE_MEMBERSHIP_MAX_REPLY_BYTES: number;
export function noteMembershipSheet(sheet: unknown): string;
export function noteMembershipFields<T>(request: T, options?: { transport?: boolean }): T;
export function noteMembershipReceipt(receipt: unknown): WorkbookNoteMembershipReceipt;
export function noteMembershipSource(source: unknown): WorkbookNoteMembershipSource;
export function noteMembershipCell(ref: unknown): number;
export function noteMembershipInner(
  reply: unknown,
  expected?: { sheet?: string; source?: WorkbookNoteMembershipSource }
): WorkbookNoteMembershipInner;
export function bindNoteMembership(
  reply: unknown,
  receipt: WorkbookNoteMembershipReceipt,
  expected?: { sheet?: string; source?: WorkbookNoteMembershipSource }
): WorkbookNoteMembershipResult;
export function validateNoteMembership(
  reply: unknown,
  expected?: { sheet?: string; source?: WorkbookNoteMembershipSource; receipt?: WorkbookNoteMembershipReceipt }
): WorkbookNoteMembershipResult;
