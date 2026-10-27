import type { WorkbookPrintHeaderFooter, WorkbookPrintTextSections } from './contracts.js';
export const PRINT_HEADER_FOOTER_LIMIT: 255;
export function emptyPrintHeaderFooter(): WorkbookPrintHeaderFooter;
export function encodePrintSections(sections: WorkbookPrintTextSections): string;
export function decodePrintSections(encoded: string): WorkbookPrintTextSections;
export function validatePrintHeaderFooter(value: WorkbookPrintHeaderFooter): WorkbookPrintHeaderFooter;
