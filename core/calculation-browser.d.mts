export interface CellCalculation {
  status: 'calculated' | 'stored' | 'unresolved' | 'error' | 'circular';
  value?: string | number | boolean;
  error?: string;
  revision?: string;
}
export interface CalculationCell {
  value?: string | number | boolean | null;
  formula?: string;
  dataType?: string;
  rawValue?: string | null;
  cachedValue?: string | number | boolean | null;
  calculation?: CellCalculation;
}
export interface CalculationView {
  revision?: string;
  sheets: { name: string; cells: Record<string, CalculationCell>; rows?: number; columns?: number }[];
  namedRanges?: { name: string; formula: string; sheet?: string }[];
  tables?: {
    name: string;
    sheet: string;
    ref: string;
    columns?: ({ name: string } | string)[];
    headerRowCount?: number;
    totalsRowCount?: number;
  }[];
}
export function calculateWorkbook<T extends CalculationView>(
  view: T,
  options?: { timeoutMs?: number }
): Promise<T & { calculationEngine: string }>;
export function shiftFormula(formula: string, columnOffset: number, rowOffset: number): string;

export function transformWorkbook<T extends CalculationView>(
  view: T,
  edits: { sheet: string; operation: 'insertRow'; index: number; count?: number }[]
): Promise<T>;

export function shiftFormulas(requests: { formula: string; columnOffset: number; rowOffset: number }[]): string[];
export function formulaReferences(formula: string): {
  references: {
    unitId: string;
    sheetName: string;
    range: {
      startRow: number;
      endRow: number;
      startColumn: number;
      endColumn: number;
      startAbsoluteRefType?: number;
      endAbsoluteRefType?: number;
    };
  }[];
  unsupported: boolean;
};

export function createCalculationSession<T extends CalculationView>(
  view: T,
  options?: { timeoutMs?: number; takeOwnership?: boolean }
): Promise<{
  view: T;
  update(
    next: T,
    changes: { sheet: string; cell: string }[]
  ): Promise<{ sheet: string; cell: string; value: CalculationCell }[]>;
  dispose(): void;
}>;
