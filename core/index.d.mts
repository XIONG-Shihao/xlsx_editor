import type { WorkbookEdit, WorkbookView } from './contracts.js';
export const MAX_WORKBOOK_BYTES: number;
export function readWorkbook(bytes: Uint8Array): WorkbookView;
export function editWorkbook(bytes: Uint8Array, edits: WorkbookEdit[], expectedRevision?: string): Uint8Array;
export function createWorkbook(): Uint8Array;
export function columnName(index: number): string;
export function cellPosition(ref: string): { row: number; col: number };

export function editWorkbookResult(
  bytes: Uint8Array,
  edits: WorkbookEdit[],
  expectedRevision?: string
): { bytes: Uint8Array; view: WorkbookView };
export function editWorkbookResult(
  bytes: Uint8Array,
  edits: WorkbookEdit[],
  expectedRevision: string | undefined,
  includeView: false,
  structuralView?: WorkbookView,
  retainContext?: boolean,
  options?: { sourceView?: WorkbookView }
): { bytes: Uint8Array; view?: WorkbookView };
export function readCalculatedWorkbook(bytes: Uint8Array): Promise<WorkbookView>;
export function editCalculatedWorkbookResult(
  bytes: Uint8Array,
  edits: WorkbookEdit[],
  expectedRevision?: string
): Promise<{ bytes: Uint8Array; view: WorkbookView }>;
