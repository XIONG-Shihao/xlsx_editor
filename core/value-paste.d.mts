import type { WorkbookCell, WorkbookSheet } from './contracts';
export interface ValuePasteSnapshot {
  readonly value: string | number | boolean | null;
  readonly valueType?: 'error';
}
export const VALUE_PASTE_ERRORS: readonly string[];
export function validateValueType<T extends { value?: unknown; valueType?: unknown; formula?: unknown }>(edit: T): T;
/** Requires an accepted owner read. Derived source fields alone do not prove native admission. */
export interface ValuePasteSourceContext {
  readonly sheet: WorkbookSheet;
  readonly cell: string;
}
/** Complete captured accepted metadata is required; known missing formula inventory refuses. */
export function createValuePasteSourceContext(sheet: WorkbookSheet): Readonly<{ sheet: WorkbookSheet }>;
export function normalizeValuePasteCell(
  cell: WorkbookCell | null | undefined,
  context?: ValuePasteSourceContext
): ValuePasteSnapshot;
