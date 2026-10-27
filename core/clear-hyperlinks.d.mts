import type { WorkbookCommand, WorkbookClearHyperlinksCommand, WorkbookView } from './contracts.js';
export const CLEAR_HYPERLINKS_LIMITS: Readonly<{ ranges: 10000; actionBytes: number }>;
export function clearHyperlinksCommand(value: unknown): WorkbookClearHyperlinksCommand;
export function normalizeClearHyperlinksCommands(
  commands: WorkbookCommand[],
  view?: Pick<WorkbookView, 'sheets'>
): WorkbookCommand[];
