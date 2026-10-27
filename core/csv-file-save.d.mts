import type { CsvIndex, CsvCommand } from './csv-workbook.mjs';
export function editCsvFile(
  source: string,
  commands: CsvCommand[],
  expectedRevision: string | undefined,
  options: { delimiter?: string } | undefined,
  destination: string
): { revision: string };
export function writeCsvIndex(
  index: CsvIndex & { revision?: string },
  source: string,
  destination: string,
  expectedRevision?: string,
  options?: { maxOutputBytes?: number }
): string;
