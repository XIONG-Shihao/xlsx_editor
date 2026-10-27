import type { WorkbookDuplicatePreview, WorkbookRemoveDuplicatesCommand } from './contracts.js';
export interface DuplicateConfiguration {
  sheet: string;
  range: string;
  hasHeaders: boolean;
  columns: string[];
}
export type DuplicateReceipt = Pick<WorkbookDuplicatePreview, 'documentId' | 'ownerId' | 'state' | 'revision'>;
export type DuplicatePage = Pick<
  WorkbookDuplicatePreview,
  'fingerprint' | 'totalRows' | 'keptRows' | 'removedRows' | 'comparison' | 'samples' | 'offset' | 'nextOffset'
>;
export function duplicateConfiguration(
  sheet: string,
  range: string,
  hasHeaders: boolean,
  columns: string[]
): DuplicateConfiguration & { first: { row: number; col: number }; last: { row: number; col: number }; body: number };
export function duplicateCommand(command: unknown): WorkbookRemoveDuplicatesCommand;
export function duplicatePage(reply: unknown, config: DuplicateConfiguration, offset?: number): DuplicatePage;
export function bindDuplicatePreview(
  reply: DuplicatePage,
  receipt: DuplicateReceipt,
  config: DuplicateConfiguration,
  offset?: number
): WorkbookDuplicatePreview;
export function validateDuplicatePreview(
  reply: unknown,
  receipt: DuplicateReceipt,
  config: DuplicateConfiguration,
  offset?: number
): WorkbookDuplicatePreview;
