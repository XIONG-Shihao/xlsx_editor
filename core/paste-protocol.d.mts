import { FORMAT } from './format.mjs';
export interface PasteShape {
  sheet: string;
  start: string;
  rows: number;
  columns: number;
}
import type { TextConversionPolicy } from './text-conversion.mjs';
export interface TypedPasteConversion {
  schema: 1;
  policies: TextConversionPolicy[];
  date1904: boolean;
  keepHeader: boolean;
  outputRows: number;
  outputColumns: number;
}
export interface LegacyPastePlan {
  shape: PasteShape;
  count: number;
  policy: 'grid' | 'text';
  conversion?: never;
}
export interface TypedPastePlan {
  shape: PasteShape;
  count: number;
  policy: 'typed-text-v1';
  conversion: TypedPasteConversion;
}
export type PastePlan = LegacyPastePlan | TypedPastePlan;
export type PasteDescriptor = (
  | (LegacyPastePlan & { schema?: never })
  | (TypedPastePlan & { schema: typeof FORMAT.SCHEMA_STAGED_PASTE_V2 })
) & { type: 'importCells'; id: string; sha256: string; bytes: number };
export interface PastePage {
  sequence: number;
  cells: [number, string][];
}
export const PASTE_POSITION_LIMIT: number,
  PASTE_BYTE_LIMIT: number,
  PASTE_PAGE_CELLS: number,
  PASTE_PAGE_BYTES: number,
  PASTE_SCHEMA: string,
  TYPED_PASTE_SCHEMA: string,
  TYPED_PASTE_POLICY: 'typed-text-v1';
export function pasteBytes(text: string): number;
export function pasteId(id: unknown): boolean;
export function validatePasteState(state: string): string;
export function validatePasteShape(shape: PasteShape): PasteShape;
export function validatePastePlan(plan: PastePlan): PastePlan;
export function validatePasteDescriptor(descriptor: PasteDescriptor): PasteDescriptor;
export function validatePasteCells(
  cells: [number, string][],
  shape: PasteShape,
  previous?: number
): { lastOffset: number; bytes: number };

export function validateTypedPasteConversion(value: unknown, shape: PasteShape): TypedPasteConversion;
export function pasteHeader(plan: PastePlan, documentId: string, state: string): Record<string, unknown>;
export function pasteDescriptorIdentity(descriptor: PasteDescriptor): string;
