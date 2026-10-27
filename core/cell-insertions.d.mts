import type { WorkbookBooleanCellCommand, WorkbookCell, WorkbookEdit, WorkbookCommand } from './contracts';
export const symbolSamples: ReadonlyArray<Readonly<{ id: string; label: string; value: string }>>;
export function booleanCellCommand(value: unknown): WorkbookBooleanCellCommand;
export function normalizeBooleanCellCommands(commands: WorkbookCommand[]): WorkbookCommand[];
export function symbolCellEdit(sheet: string, cell: string, source: WorkbookCell | null, id: string): WorkbookEdit;
export function booleanCellValue(cell: Readonly<WorkbookCell> | null): boolean | null;
