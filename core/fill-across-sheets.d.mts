import type { WorkbookCommand, WorkbookFillAcrossSheetsCommand, WorkbookView } from './contracts.js';
export const FILL_ACROSS_SHEETS_LIMITS: Readonly<{ destinationPositions: 10000; actionBytes: number }>;
export function fillAcrossSheetsCommand(value: unknown): WorkbookFillAcrossSheetsCommand;
export function normalizeFillAcrossSheetsRuns(
  commands: WorkbookCommand[],
  view?: Pick<WorkbookView, 'sheets'>
): WorkbookCommand[];
