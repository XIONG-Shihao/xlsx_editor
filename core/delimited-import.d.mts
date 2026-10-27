import { FORMAT } from './format.mjs';
import type { PastePlan, PastePage, PasteDescriptor } from './paste-protocol.mjs';
export interface DelimitedSource {
  readonly id: string;
  readonly name: string;
  readonly size: number;
  readonly sha256?: string;
  chunks(): AsyncIterable<Uint8Array>;
  check?(): void | Promise<void>;
  dispose?(): Promise<void>;
}
export interface DelimitedInspection extends PastePlan {
  schema: typeof FORMAT.SCHEMA_DELIMITED_FILE_IMPORT_V1;
  source: { id: string; name: string; bytes: number; sha256: string };
  delimiter: ',' | ';' | '\t';
  encoding: 'utf-8' | 'utf-16le' | 'utf-16be';
  bom: boolean;
  policy: 'text';
}
export interface DelimitedProgress {
  phase: string;
  completed: number;
  total: number;
}
export const DELIMITED_READ_BYTES: number, DELIMITED_IMPORT_SCHEMA: string;
export function delimitedFields(
  source: DelimitedSource,
  delimiter: ',' | ';' | '\t',
  options?: { signal?: AbortSignal; proof?: Record<string, unknown>; progress?: (done: number, total: number) => void }
): AsyncGenerator<{ row: number; col: number; text: string }>;
export function inspectDelimitedSource(
  source: DelimitedSource,
  options: {
    sheet: string;
    start: string;
    delimiter?: 'auto' | ',' | ';' | '\t';
    signal?: AbortSignal;
    onProgress?: (event: DelimitedProgress) => void;
  }
): Promise<DelimitedInspection>;
export function delimitedImportPages(
  source: DelimitedSource,
  inspection: DelimitedInspection,
  options?: { signal?: AbortSignal; onProgress?: (event: DelimitedProgress) => void }
): AsyncGenerator<PastePage>;
export function stageDelimitedImport(
  source: DelimitedSource,
  inspection: DelimitedInspection,
  options: {
    state: string;
    signal?: AbortSignal;
    checkCurrent?: () => void | Promise<void>;
    sink: {
      begin: (input: any) => Promise<any>;
      upload: (input: any) => Promise<any>;
      seal: (input: any) => Promise<any>;
      cancel: (input: any) => Promise<any>;
      prune: (input: any) => Promise<any>;
    };
    onProgress?: (event: DelimitedProgress) => void;
  }
): Promise<{
  descriptor: PasteDescriptor;
  source: DelimitedInspection['source'];
  delimiter: DelimitedInspection['delimiter'];
  encoding: DelimitedInspection['encoding'];
  bom: boolean;
}>;
export function createBlobDelimitedSource(
  blob: Blob,
  options?: { name?: string; chunkBytes?: number; signal?: AbortSignal }
): DelimitedSource;
