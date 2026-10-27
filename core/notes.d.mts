import { FORMAT } from './format.mjs';
import type { WorkbookView } from './contracts';
export interface LegacyNote {
  cell: string;
  author: string;
  text: string;
}
export type NoteCommand =
  | {
      type: 'setNote';
      sheet: string;
      cell: string;
      expectedNoteRevision: string | null;
      note: { author: string; text: string };
    }
  | { type: 'removeNote'; sheet: string; cell: string; expectedNoteRevision: string | null };
export interface NoteInventory {
  schema: typeof FORMAT.SCHEMA_LEGACY_NOTES_V1;
  sheet: string;
  noteRevision: string | null;
  notes: LegacyNote[] | null;
  canEdit: boolean;
  protectedReason: string | null;
}
export interface NoteReceipt {
  documentId: string;
  ownerId: string;
  state: string;
  revision: string;
}
export const NOTE_LIMITS: Readonly<{ notes: number; author: number; text: number; batchBytes: number }>;
export function isNoteCommand(v: unknown): v is NoteCommand;
export function noteCommand(v: unknown): NoteCommand;
export function captureNoteCommands(v: unknown, view?: Pick<WorkbookView, 'sheets'>): unknown[];
export function notesFields(v: unknown, options?: { transport?: boolean }): unknown;
export function noteInventory(v: unknown, expected: { sheet: string }): NoteInventory;
export function bindNotes(v: unknown, binding: NoteReceipt, expected: { sheet: string }): NoteInventory & NoteReceipt;
export function validateNotes(
  v: unknown,
  binding: NoteReceipt,
  expected: { sheet: string }
): NoteInventory & NoteReceipt;
