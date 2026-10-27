import type { WorkbookView, WorkbookCell } from './contracts.js';
import type { CsvCommand, CsvEdit } from './csv-workbook.mjs';
import type { SelectionSummary } from './indexed-session.mjs';
export interface CsvFileSessionOptions {
  expectedRevision?: string;
  delimiter?: ',' | ';' | '\t' | 'auto';
  storageDirectory?: string;
  maxSessionBytes?: number;
  maxIndexBytes?: number;
  maxHistoryBytes?: number;
  maxProjectionBytes?: number;
  maxReplyBytes?: number;
  maxStates?: number;
}
export type CsvFileCommand = CsvCommand | { type: 'clearRange'; sheet: string; range: string };
export class CsvFileSession {
  readonly id: string;
  readonly cacheHit: false;
  open(source: string, options?: CsvFileSessionOptions): Promise<WorkbookView>;
  view(state: string): WorkbookView;
  apply(state: string, commands: CsvFileCommand[]): Promise<{ view: WorkbookView; edits: CsvEdit[] }>;
  read(
    state: string,
    sheet: string,
    range: string
  ): Promise<{ sheet: string; range: string; cells: Record<string, WorkbookCell> }>;
  scan(
    state: string,
    sheet: string,
    after?: string
  ): Promise<{ cells: Record<string, WorkbookCell>; nextCursor: string | null }>;
  query(
    state: string,
    sheet: string,
    query: string,
    mode?: 'display' | 'raw' | 'formula',
    after?: string,
    options?: { range?: string; matchLimit?: number }
  ): Promise<{ cells: Record<string, WorkbookCell>; matches: number; scanned: number; nextCursor: string | null }>;
  selectionSummary(state: string, sheet: string, range: string): Promise<SelectionSummary>;
  export(state: string): Promise<{ path: string; revision: string }>;
  compact(retain: string[]): Promise<{ compacted: true }>;
  diagnosticStats(): {
    retainedStates: number;
    storageRoots: number;
    indexBytes: number;
    historyBytes: number;
    exportBytes: number;
    decodedRowCacheBytes: number;
    closed: boolean;
  };
  close(): Promise<void>;
}
