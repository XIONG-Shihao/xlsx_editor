import type { WorkbookCell, WorkbookCommand } from './contracts';
export type ReviewKind = 'blank' | 'text' | 'number' | 'boolean' | 'error' | 'unavailable';
export type ReviewAuthority =
  | 'absent'
  | 'literal'
  | 'current'
  | 'stored'
  | 'pending'
  | 'unresolved'
  | 'circular'
  | 'unqualified';
export type ReviewFormulaState =
  | 'calculated'
  | 'stored'
  | 'pending'
  | 'unresolved'
  | 'error'
  | 'circular'
  | 'unqualified';
export interface ReviewContext {
  revision: string;
  special?: boolean;
  protected?: boolean;
}
export interface ReviewCell extends WorkbookCell {
  protected?: boolean;
  readOnly?: boolean;
  protection?: unknown;
  scenario?: unknown;
}
export interface ReviewClassification {
  presence: 'absent' | 'authored';
  authored: 'blank' | 'literal' | 'formula' | null;
  kind: ReviewKind;
  authority: ReviewAuthority;
  formulaState: ReviewFormulaState | null;
  special: boolean;
  protected: boolean;
  unsafeNumber: boolean;
  spellingEligible: boolean;
  reason: string | null;
}
export interface ReviewStatistics {
  schema: 1;
  revision: string;
  positions: number;
  authoredCells: number;
  absentPositions: number;
  authoredBlankCells: number;
  literalCells: number;
  formulaCells: number;
  specialCells: number;
  protectedCells: number;
  unsafeNumericCells: number;
  spellingEligibleCells: number;
  literalKinds: Record<ReviewKind, number>;
  formulaStates: Record<ReviewFormulaState, number>;
  results: Record<ReviewAuthority, Record<ReviewKind, number>>;
}
/** Offsets are exact original UTF-16 indices. Pending tokens may span pages. */
export interface ReviewTokenCursor {
  offset: number;
  start: number | null;
  points: number;
  hasNumber: boolean;
  oversized: boolean;
  ready: boolean;
}
export interface ReviewToken {
  text: string;
  start: number;
  end: number;
}
export interface ReviewExcludedToken {
  start: number;
  end: number;
  reason: 'token-too-long' | 'numeric-or-identifier' | 'cell-text-limit';
}
export type ReviewPageReason =
  | 'token-limit'
  | 'exclusion-limit'
  | 'byte-limit'
  | 'work-limit'
  | 'cell-text-limit'
  | null;
export interface ReviewTokenPage {
  schema: 1;
  tokens: ReviewToken[];
  excluded: ReviewExcludedToken[];
  nextCursor: ReviewTokenCursor | null;
  coverage: { complete: boolean; reason: ReviewPageReason; offset: number; workUnits: number };
}
/** IDs are the verified, pinned provider/dictionary content identities. */
export interface ReviewReceipt {
  documentId: string;
  ownerId: string;
  state: string;
  revision: string;
  locale: 'en-US' | 'en-GB';
  providerId: string;
  dictionaryId: string;
}
export interface ReviewTarget {
  sheet: string;
  cell: string;
}
export interface SpellingFinding {
  schema: 1;
  receipt: ReviewReceipt;
  target: ReviewTarget;
  cellIdentity: string;
  token: ReviewToken;
}
export const WORKBOOK_REVIEW_LIMITS: Readonly<{
  statisticsCells: 256;
  statisticsBytes: 1048576;
  textBytes: 1048576;
  wordCodePoints: 256;
  tokens: 100;
  workUnits: 32768;
  pageBytes: 32768;
  commandBytes: 1048576;
}>;
export function classifyReviewCell(cell: ReviewCell | undefined, context: ReviewContext): ReviewClassification;
export function reviewStatistics(cells: Array<ReviewCell | undefined>, context: ReviewContext): ReviewStatistics;
export function addReviewStatistics(
  summary: ReviewStatistics,
  cells: Array<ReviewCell | undefined>,
  context: ReviewContext
): ReviewStatistics;
export function reviewTokens(
  text: string,
  cursor?: ReviewTokenCursor,
  limits?: { maxTokens?: number; maxWork?: number; maxBytes?: number }
): ReviewTokenPage;
export function validateReviewReceipt(receipt: ReviewReceipt, expected?: ReviewReceipt): ReviewReceipt;
export function captureSpellingFinding(input: {
  receipt: ReviewReceipt;
  target: ReviewTarget;
  cell: ReviewCell;
  token: ReviewToken;
}): SpellingFinding;
export function planSpellingReplacement(
  finding: SpellingFinding,
  current: {
    receipt: ReviewReceipt;
    cell: ReviewCell;
    replacement: string;
    writable: boolean;
    special?: boolean;
    protected?: boolean;
  }
): WorkbookCommand;
