import type { CsvIndex, CsvCommand, CsvEdit, CsvAffected } from './csv-workbook.mjs';
export function readCsvWorkbook(bytes: Uint8Array, options?: { delimiter?: string }): CsvIndex & { revision: string };
export function editCsvWorkbook(
  bytes: Uint8Array,
  commands: CsvCommand[],
  expectedRevision?: string,
  options?: { delimiter?: string }
): { bytes: Uint8Array; index: CsvIndex & { revision: string }; edits: CsvEdit[]; affected: CsvAffected[] };
export function createCsvWorkbook(rows: Array<Array<string | number | boolean | null>>): Uint8Array;

export function editCsvBytes(
  bytes: Uint8Array,
  commands: CsvCommand[],
  expectedRevision?: string,
  options?: { delimiter?: string }
): { bytes: Uint8Array; revision: string };
