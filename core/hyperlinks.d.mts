import type {
  WorkbookCommand,
  WorkbookView,
  WorkbookHyperlinkCommand,
  WorkbookHyperlinkDefinition,
  WorkbookHyperlinkInner,
  WorkbookHyperlinkDestination,
  WorkbookHyperlinkReceipt,
} from './contracts.js';
export function hyperlinkCell(value: unknown): string;
export function hyperlinkDestination(value: unknown): WorkbookHyperlinkDestination;
export function hyperlinkCommand(value: unknown): WorkbookHyperlinkCommand;
export function normalizeHyperlinkCommands(commands: WorkbookCommand[], view?: WorkbookView): WorkbookCommand[];
export function hyperlinkFields(request: unknown, options?: { transport?: boolean }): unknown;
export function bindHyperlink(
  value: WorkbookHyperlinkInner,
  binding: WorkbookHyperlinkReceipt,
  expected: { sheet: string; cell: string }
): WorkbookHyperlinkDefinition;
export function validateHyperlink(
  value: unknown,
  binding: WorkbookHyperlinkReceipt,
  expected: { sheet: string; cell: string }
): WorkbookHyperlinkDefinition;

export function hyperlinkDisplayText(value: unknown): string;
