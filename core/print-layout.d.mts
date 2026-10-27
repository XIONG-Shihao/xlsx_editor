import type { WorkbookPrintSnapshot, WorkbookPrintLayout } from './contracts.js';
export function composePrint(snapshot: WorkbookPrintSnapshot): WorkbookPrintLayout;
export function printPageHtml(layout: WorkbookPrintLayout, pageNumber: number): string;
export function printDocumentHtml(layout: WorkbookPrintLayout, options?: { pages?: number[] }): string;
export function printOverflowInspectionSource(): string;
