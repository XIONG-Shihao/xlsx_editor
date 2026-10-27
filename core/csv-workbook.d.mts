export type CsvScalar = string | number | boolean | null;
export interface CsvEdit {
  sheet: string;
  cell: string;
  value: CsvScalar;
  expectedValue?: CsvScalar;
  expectedRawValue?: string | null;
  expectedFormula?: null;
}
export type CsvCommand =
  | CsvEdit
  | {
      type: 'setCells';
      sheet: string;
      cells: Omit<CsvEdit, 'sheet'>[];
    }
  | {
      type: 'updateRows';
      sheet: string;
      range: string;
      keyColumn: string;
      updates: {
        key: CsvScalar;
        cells: ({
          column: string;
        } & Omit<CsvEdit, 'sheet' | 'cell'>)[];
      }[];
    };
export interface CsvAffected {
  sheet: string;
  cell: string;
  value: string;
  rawValue: string;
  dataType: 'text';
}
export interface CsvIndex {
  kind: 'csv';
  sheet: 'CSV';
  text: string | { readonly length: number; slice(start?: number, end?: number): string };
  /** Optional owner budget for one decoded row and its field records. */
  maxRowBytes?: number;
  rowOffsets: Float64Array;
  rowWidths: Uint32Array;
  rows: number;
  columns: number;
  nonemptyCells: number;
  storedCells: number;
  delimiter: ',' | ';' | '\t';
  encoding: 'utf-8' | 'utf-16le' | 'utf-16be';
  bom: boolean;
  newline: string;
  terminalNewline: boolean;
  changes: Record<string, string>;
  widthOverrides: Record<string, number>;
}
export const CSV_SHEET: 'CSV';
export function indexCsv(
  text: string,
  options?: {
    delimiter?: ',' | ';' | '\t' | 'auto';
    encoding?: CsvIndex['encoding'];
    bom?: boolean;
  }
): CsvIndex;
export function csvCell(index: CsvIndex, row: number, column: number): string | null;
export function csvCellInfo(
  index: CsvIndex,
  row: number,
  column: number
): {
  value: string;
  rawValue: string;
  dataType: 'text';
} | null;
export function csvReadRange(
  index: CsvIndex,
  range: string,
  options?: {
    offset?: number;
    limit?: number;
  }
): {
  items: {
    sheet: string;
    cell: string;
    value: string | null;
    rawValue: string | null;
    dataType: string;
  }[];
  total: number;
  offset: number;
  nextOffset: number | null;
};
export function applyCsvCommands(
  index: CsvIndex,
  commands: unknown[]
): {
  index: CsvIndex;
  edits: CsvEdit[];
  affected: CsvAffected[];
};
export function serializeCsv(index: CsvIndex, edits?: unknown[]): string;
export function csvTextEdits(index: CsvIndex): Array<{ start: number; end: number; text: string }>;
export function detectCsvDelimiter(text: string): CsvIndex['delimiter'];
export function csvCacheStatistics(index: CsvIndex): {
  rows: number;
  estimatedBytes: number;
  maxRows: number;
  maxEstimatedBytes: number;
};

export function csvPosition(address: string): { row: number; column: number };
export function csvAddress(row: number, column: number): string;

export function indexCsvChunks(
  chunks: Iterable<string>,
  options?: { delimiter?: ',' | ';' | '\t'; encoding?: CsvIndex['encoding']; bom?: boolean; maxIndexBytes?: number }
): Omit<CsvIndex, 'text'>;
