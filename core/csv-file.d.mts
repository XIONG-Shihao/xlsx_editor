import type { CsvTransferPayload } from './csv-transfer.js';
import type { CsvDocumentView } from './contracts.js';
export function readCsvFileTransfer(
  file: string,
  options?: { delimiter?: string; maxIndexBytes?: number },
  expectedRevision?: string
): Omit<CsvDocumentView, 'text' | 'sourceSnapshot'> & CsvTransferPayload;

import type { CsvIndex } from './csv-workbook.mjs';
export function readCsvFileIndex(
  file: string,
  options?: { delimiter?: string; maxIndexBytes?: number },
  expectedRevision?: string
): CsvIndex & { revision: string };
export function checkCsvFileIndexSource(index: CsvIndex, source?: string, descriptor?: number): boolean;
