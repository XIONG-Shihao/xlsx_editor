import type { WorkbookCell, WorkbookCommand, WorkbookBooleanCellCommand } from '../core/contracts';
import { booleanCellCommand, booleanCellValue, symbolCellEdit } from '../core/cell-insertions.mjs';
export interface CellInsertionFrame {
  key: string;
  family: string;
  sheet: string;
  cell: string;
  active: boolean;
  writable: boolean;
  reason?: string;
  read: () => Promise<WorkbookCell | null>;
  apply: (commands: WorkbookCommand[]) => Promise<void>;
}
export interface CellInsertionSnapshot {
  key: string;
  cell: Readonly<WorkbookCell> | null;
  loaded: boolean;
  loading: boolean;
  pending: boolean;
  error: string;
  notice: string;
}
/** Shared Checkbox/Symbols inspection lane. Each action freezes the displayed cell. */
export function createCellInsertionController() {
  let frame: CellInsertionFrame | null = null,
    wanted: CellInsertionFrame | null = null,
    reading = false,
    working = false,
    disposed = false;
  let snapshot: CellInsertionSnapshot = {
    key: '',
    cell: null,
    loaded: false,
    loading: false,
    pending: false,
    error: '',
    notice: '',
  };
  const listeners = new Set<() => void>();
  const emit = (next: CellInsertionSnapshot) => {
      snapshot = next;
      for (const listener of listeners) listener();
    },
    same = (f: CellInsertionFrame) => !disposed && frame?.key === f.key && frame.active;
  const text = (e: unknown) => String(e instanceof Error ? e.message : e);
  async function drain() {
    if (reading || disposed) return;
    reading = true;
    try {
      while (wanted && !disposed) {
        const captured = wanted;
        wanted = null;
        if (!same(captured)) continue;
        emit({ ...snapshot, key: captured.key, cell: null, loaded: false, loading: true, error: '' });
        try {
          const result = await captured.read();
          if (result !== null && (typeof result !== 'object' || Array.isArray(result)))
            throw Error('Invalid selected cell reply.');
          if (same(captured))
            emit({
              ...snapshot,
              cell: result ? Object.freeze(structuredClone(result)) : null,
              loaded: true,
              loading: false,
              error: '',
            });
        } catch (e) {
          if (same(captured)) emit({ ...snapshot, cell: null, loaded: false, loading: false, error: text(e) });
        }
      }
    } finally {
      reading = false;
    }
  }
  function target(captured: CellInsertionFrame) {
    if (!same(captured) || snapshot.key !== captured.key)
      throw Error('The original cell or workbook changed. Select it again.');
    if (working || snapshot.loading) throw Error('Wait for the cell insertion operation to finish.');
    if (!snapshot.loaded) throw Error('Load the selected cell first.');
    if (!frame!.writable || !captured.writable) throw Error(frame!.reason ?? 'This cell is read-only.');
    return { ...frame! };
  }
  async function write(captured: CellInsertionFrame, commands: WorkbookCommand[], notice: string) {
    working = true;
    emit({ ...snapshot, pending: true, error: '', notice: '' });
    try {
      await captured.apply(commands);
      if (frame?.active && frame.family === captured.family) emit({ ...snapshot, notice });
      return { accepted: true, current: same(captured) };
    } catch (e) {
      if (same(captured)) emit({ ...snapshot, error: text(e) });
      throw e;
    } finally {
      working = false;
      emit({ ...snapshot, pending: false });
      if (wanted) void drain();
    }
  }
  return {
    subscribe: (listener: () => void) => {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    getSnapshot: () => snapshot,
    setFrame: (next: CellInsertionFrame | null) => {
      const changed = frame?.key !== next?.key || frame?.active !== next?.active;
      frame = next;
      if (!changed) return;
      wanted = next?.active ? { ...next } : null;
      emit({
        key: next?.key ?? '',
        cell: null,
        loaded: false,
        loading: !!wanted,
        pending: working,
        error: '',
        notice: '',
      });
      void drain();
    },
    refresh: () => {
      if (frame?.active) {
        wanted = { ...frame };
        void drain();
      }
    },
    boolean: async (captured: CellInsertionFrame, action: WorkbookBooleanCellCommand['action']) => {
      const current = target(captured),
        value = booleanCellValue(snapshot.cell);
      if (action === 'remove' && value === null) throw Error('There is no Boolean-cell checkbox to remove.');
      const command = booleanCellCommand({
        type: 'booleanCell',
        sheet: current.sheet,
        cell: current.cell,
        action,
        expectedValue: value,
      });
      return write(
        current,
        [command],
        action === 'remove' ? 'Checkbox removed.' : action === 'insert' ? 'Checkbox inserted.' : 'Checkbox toggled.'
      );
    },
    symbol: async (captured: CellInsertionFrame, id: string) => {
      const current = target(captured),
        command = symbolCellEdit(current.sheet, current.cell, snapshot.cell, id);
      return write(current, [command], 'Symbol inserted.');
    },
    dispose: () => {
      disposed = true;
      frame = null;
      wanted = null;
      listeners.clear();
    },
  };
}
