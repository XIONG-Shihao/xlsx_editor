import {
  localExecutionPreview,
  type LocalExecutionAction,
  type LocalExecutionPreview,
} from '../core/local-action-execution.mjs';
import type { WorkbookCommand } from '../core/contracts';
export interface LocalExecutionFrame {
  key: string;
  ownerKey: string;
  sheet: string;
  selection: string;
  active: boolean;
  reason?: string;
  apply: (commands: WorkbookCommand[]) => Promise<unknown>;
}
export interface LocalExecutionSnapshot {
  key: string;
  ownerKey: string;
  phase: 'idle' | 'preview' | 'running' | 'completed' | 'failed';
  preview: LocalExecutionPreview | null;
  message: string;
  pending: boolean;
}
/** Prepared targets are private; UI copies cannot change the submitted command. */
export function createLocalExecutionController() {
  let frame: LocalExecutionFrame | null = null,
    prepared: { frame: LocalExecutionFrame; preview: LocalExecutionPreview } | null = null,
    flight = false,
    disposed = false;
  let snapshot: LocalExecutionSnapshot = {
    key: '',
    ownerKey: '',
    phase: 'idle',
    preview: null,
    message: '',
    pending: false,
  };
  const listeners = new Set<() => void>();
  const emit = (phase: LocalExecutionSnapshot['phase'], preview: LocalExecutionPreview | null = null, message = '') => {
    snapshot = {
      key: frame?.key ?? '',
      ownerKey: frame?.ownerKey ?? '',
      phase,
      preview: preview ? structuredClone(preview) : null,
      message,
      pending: flight,
    };
    for (const l of listeners) l();
  };
  const sameOwner = (f: LocalExecutionFrame) => !disposed && !!frame && frame.active && frame.ownerKey === f.ownerKey;
  const sameTarget = (f: LocalExecutionFrame) => sameOwner(f) && frame?.key === f.key && !frame.reason;
  return {
    subscribe: (f: () => void) => {
      listeners.add(f);
      return () => {
        listeners.delete(f);
      };
    },
    getSnapshot: () => snapshot,
    setFrame(next: LocalExecutionFrame | null) {
      if (disposed) return;
      const before = frame;
      frame = next;
      if (prepared && !sameTarget(prepared.frame)) {
        prepared = null;
        emit(
          flight ? 'running' : 'idle',
          null,
          next?.active ? 'The workbook or selection changed. Preview the action again.' : ''
        );
        return;
      }
      if (before?.ownerKey !== next?.ownerKey || before?.active !== next?.active) emit(flight ? 'running' : 'idle');
    },
    preview(action: LocalExecutionAction) {
      if (flight || !frame || !frame.active || frame.reason) {
        if (frame?.active) emit('failed', null, frame.reason ?? 'Wait for the current action to finish.');
        return false;
      }
      try {
        const captured = { ...frame },
          preview = localExecutionPreview(action, { sheet: captured.sheet, selection: captured.selection });
        prepared = { frame: captured, preview: structuredClone(preview) };
        emit('preview', preview);
        return true;
      } catch (e) {
        prepared = null;
        emit('failed', null, e instanceof Error ? e.message : String(e));
        return false;
      }
    },
    cancel() {
      if (flight) return false;
      prepared = null;
      emit('idle', null, 'Action preview cancelled.');
      return true;
    },
    async run() {
      const target = prepared;
      if (flight || !target || !sameTarget(target.frame)) return { accepted: false, current: false };
      const commands = structuredClone(target.preview.commands);
      prepared = null;
      flight = true;
      emit('running', target.preview, 'Running ' + target.preview.label + '…');
      try {
        await target.frame.apply(commands);
        const current = sameOwner(target.frame);
        if (current)
          emit(
            'completed',
            null,
            target.preview.label + ' completed for ' + target.preview.sheet + '. You can Undo this action.'
          );
        return { accepted: true, current };
      } catch (e) {
        const current = sameOwner(target.frame) && frame?.key === target.frame.key;
        if (current) emit('failed', null, e instanceof Error ? e.message : String(e));
        return { accepted: false, current };
      } finally {
        flight = false;
        if (sameOwner(target.frame))
          emit(snapshot.phase === 'running' ? 'idle' : snapshot.phase, snapshot.preview, snapshot.message);
        else emit('idle');
      }
    },
    dispose() {
      disposed = true;
      frame = null;
      prepared = null;
      emit(flight ? 'running' : 'idle');
      listeners.clear();
    },
  };
}
