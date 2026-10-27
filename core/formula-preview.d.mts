import type {
  WorkbookFormulaPreview,
  WorkbookFormulaPreviewReceipt,
  WorkbookFormulaPreviewCancellation,
} from './contracts.js';
export type FormulaPreviewInner = Omit<WorkbookFormulaPreview, 'receipt'>;
export function formulaPreviewRequestId(value: unknown): string;
export function formulaPreviewTarget(
  sheet: unknown,
  cell: unknown,
  formula: unknown
): { sheet: string; cell: string; formula: string };
export function formulaPreviewFields<T>(request: T, options?: { transport?: boolean; cancel?: boolean }): T;
export function formulaPreviewReceipt(receipt: unknown): WorkbookFormulaPreviewReceipt;
export function formulaPreviewInner(
  reply: unknown,
  expected?: Pick<FormulaPreviewInner, 'requestId' | 'sheet' | 'cell' | 'formula'>
): FormulaPreviewInner;
export function bindFormulaPreview(
  reply: unknown,
  receipt: WorkbookFormulaPreviewReceipt,
  expected?: Pick<FormulaPreviewInner, 'requestId' | 'sheet' | 'cell' | 'formula'>
): WorkbookFormulaPreview;
export function validateFormulaPreview(
  reply: unknown,
  expected?: Pick<WorkbookFormulaPreview, 'requestId' | 'sheet' | 'cell' | 'formula'> & {
    receipt?: WorkbookFormulaPreviewReceipt;
  }
): WorkbookFormulaPreview;
export function formulaPreviewCancellation(reply: unknown, requestId: string): WorkbookFormulaPreviewCancellation;
