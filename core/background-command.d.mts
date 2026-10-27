import type { WorkbookCommand, WorkbookBackgroundCommand } from './contracts';
export function backgroundCommand(value: unknown): WorkbookBackgroundCommand;
export function normalizeBackgroundCommands(commands: WorkbookCommand[]): WorkbookCommand[];
