import { FORMAT } from './format.mjs';
import type { NoteReceipt } from './notes.mjs';
export interface AccessibilityFinding {
  rule: 'missing-alt-text' | 'blank-table-header' | 'duplicate-table-header';
  message: string;
  sheet: string;
  cell: string;
  objectId: string | null;
  tableName: string | null;
  relatedCells: string[];
}
export interface AccessibilityReport {
  schema: typeof FORMAT.SCHEMA_ACCESSIBILITY_V1;
  scope: 'workbook';
  status: 'complete' | 'partial';
  rules: string[];
  findings: AccessibilityFinding[];
  coverage: {
    objects: { total: number | null; checked: number; skipped: number };
    tables: { total: number | null; checked: number; skipped: number };
    headerCellsRead: number;
    skipped: Array<{ kind: 'objects' | 'table'; sheet: string | null; name: string | null; reason: string }>;
  };
}
export const ACCESSIBILITY_LIMITS: Readonly<{
  tables: number;
  columns: number;
  headerCells: number;
  findings: number;
  skips: number;
}>;
export const ACCESSIBILITY_RULES: readonly string[];
export function accessibilityFields(v: unknown, options?: { transport?: boolean }): unknown;
export function accessibilityReport(v: unknown): AccessibilityReport;
export function bindAccessibility(v: unknown, binding: NoteReceipt): AccessibilityReport & NoteReceipt;
export function validateAccessibility(v: unknown, binding: NoteReceipt): AccessibilityReport & NoteReceipt;
