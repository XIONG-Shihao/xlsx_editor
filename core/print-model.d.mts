import type { WorkbookPrintSettings } from './contracts.js';
export const PRINT_LIMITS: Readonly<{
  sourcePositions: number;
  pages: number;
  renderedPositions: number;
  snapshotBytes: number;
  layoutBytes: number;
  htmlBytes: number;
  objects: number;
}>;
export function defaultPrintSettings(): WorkbookPrintSettings;
export function printPaperSize(settings: WorkbookPrintSettings): [number, number];
export function validatePrintSettings(
  input: WorkbookPrintSettings,
  options?: { allowAbsentArea?: boolean }
): WorkbookPrintSettings;
