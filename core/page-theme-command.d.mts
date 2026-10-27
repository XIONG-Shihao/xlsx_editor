import type { WorkbookCommand, WorkbookPageThemeCommand } from './contracts';
export const ownedPageThemePart: string;
export function pageThemeCommand(value: unknown): WorkbookPageThemeCommand;
export function normalizePageThemeCommands(commands: WorkbookCommand[]): WorkbookCommand[];
