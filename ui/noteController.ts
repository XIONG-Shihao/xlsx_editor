import type { WorkbookCommand } from '../core/contracts';
import type { NoteInventory, NoteReceipt } from '../core/notes.mjs';
import { noteCommand, validateNotes } from '../core/notes.mjs';
export interface NoteFrame {
  key: string;
  sheet: string;
  cell: string;
  receipt: NoteReceipt;
  active: boolean;
  writable: boolean;
  reason?: string;
  read: () => Promise<NoteInventory & NoteReceipt>;
  apply: (commands: WorkbookCommand[]) => Promise<void>;
  navigate: (sheet: string, cell: string) => void;
}
export interface NoteSnapshot {
  key: string;
  inventory: (NoteInventory & NoteReceipt) | null;
  loading: boolean;
  pending: boolean;
  error: string;
}
/** A retired read joins before its successor; the accepted write lane survives frame changes. */
export function createNoteController() {
  let frame: NoteFrame | null = null,
    wanted: NoteFrame | null = null,
    reading = false,
    working = false,
    disposed = false;
  let snapshot: NoteSnapshot = { key: '', inventory: null, loading: false, pending: false, error: '' };
  const listeners = new Set<() => void>();
  const emit = (next: NoteSnapshot) => {
    snapshot = next;
    for (const fn of listeners) fn();
  };
  const same = (captured: NoteFrame) => !disposed && frame?.active && frame.key === captured.key;
  const capture = (next: NoteFrame) => ({ ...next, receipt: { ...next.receipt } });
  const message = (error: unknown) => String(error instanceof Error ? error.message : error);
  async function drain() {
    if (reading || disposed) return;
    reading = true;
    try {
      while (wanted && !disposed) {
        const captured = wanted;
        wanted = null;
        if (!same(captured)) continue;
        try {
          const value = await captured.read();
          validateNotes(value, captured.receipt, { sheet: captured.sheet });
          if (same(captured)) emit({ ...snapshot, inventory: structuredClone(value), loading: false, error: '' });
        } catch (error) {
          if (same(captured)) emit({ ...snapshot, inventory: null, loading: false, error: message(error) });
        }
      }
    } finally {
      reading = false;
    }
  }
  function target(captured: NoteFrame) {
    if (!same(captured) || snapshot.key !== captured.key) throw Error('The original note cell or workbook changed.');
    if (working || snapshot.loading) throw Error('Wait for the note operation to finish.');
    if (!snapshot.inventory) throw Error('Load the current notes first.');
    return snapshot.inventory;
  }
  return {
    subscribe: (fn: () => void) => {
      listeners.add(fn);
      return () => {
        listeners.delete(fn);
      };
    },
    getSnapshot: () => snapshot,
    setFrame: (next: NoteFrame | null) => {
      const changed = frame?.key !== next?.key || frame?.active !== next?.active;
      frame = next;
      if (!changed) return;
      wanted = next?.active ? capture(next) : null;
      emit({ key: next?.key ?? '', inventory: null, loading: !!wanted, pending: working, error: '' });
      void drain();
    },
    refresh: () => {
      if (frame?.active) {
        wanted = capture(frame);
        emit({ ...snapshot, inventory: null, loading: true, error: '' });
        void drain();
      }
    },
    write: async (captured: NoteFrame, note: { author: string; text: string } | null) => {
      const inventory = target(captured);
      if (!frame!.writable || !captured.writable || !inventory.canEdit)
        throw Error(frame!.reason ?? inventory.protectedReason ?? 'This note is read-only.');
      const command = noteCommand({
        type: note === null ? 'removeNote' : 'setNote',
        sheet: captured.sheet,
        cell: captured.cell,
        expectedNoteRevision: inventory.noteRevision,
        ...(note ? { note } : {}),
      });
      working = true;
      emit({ ...snapshot, pending: true, error: '' });
      try {
        await captured.apply([command]);
        if (same(captured)) {
          wanted = capture(frame!);
          emit({ ...snapshot, inventory: null, loading: true });
        }
        return { accepted: true, current: same(captured) };
      } catch (error) {
        if (same(captured)) emit({ ...snapshot, error: message(error) });
        throw error;
      } finally {
        working = false;
        emit({ ...snapshot, pending: false });
        if (wanted) void drain();
      }
    },
    navigate: (captured: NoteFrame, cell: string) => {
      const inventory = target(captured);
      if (!inventory.notes?.some((n) => n.cell === cell)) throw Error('Choose a note from the current worksheet list.');
      captured.navigate(captured.sheet, cell);
    },
    dispose: () => {
      disposed = true;
      frame = null;
      wanted = null;
      listeners.clear();
    },
  };
}
