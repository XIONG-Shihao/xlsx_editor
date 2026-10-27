import { validateAccessibility, type AccessibilityReport, type AccessibilityFinding } from '../core/accessibility.mjs';
import type { NoteReceipt } from '../core/notes.mjs';

export interface AccessibilityFrame {
  key: string;
  receipt: NoteReceipt;
  active: boolean;
  visibleSheets: string[];
  read: () => Promise<AccessibilityReport & NoteReceipt>;
  navigate: (sheet: string, cell: string) => void;
}
interface Snapshot {
  key: string;
  report: (AccessibilityReport & NoteReceipt) | null;
  loading: boolean;
  error: string;
}
/** One physical read joins before the latest frame; findings never grant a write. */
export function createAccessibilityController() {
  let frame: AccessibilityFrame | null = null,
    wanted: AccessibilityFrame | null = null,
    reading = false,
    disposed = false;
  let snapshot: Snapshot = { key: '', report: null, loading: false, error: '' };
  const listeners = new Set<() => void>();
  const emit = (next: Snapshot) => {
    snapshot = next;
    for (const fn of listeners) fn();
  };
  const same = (captured: AccessibilityFrame) => !disposed && frame?.active && frame.key === captured.key;
  const capture = (next: AccessibilityFrame) => ({
    ...next,
    receipt: { ...next.receipt },
    visibleSheets: [...next.visibleSheets],
  });
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
          validateAccessibility(value, captured.receipt);
          if (same(captured)) emit({ ...snapshot, report: structuredClone(value), loading: false, error: '' });
        } catch (error) {
          if (same(captured))
            emit({
              ...snapshot,
              report: null,
              loading: false,
              error: String(error instanceof Error ? error.message : error),
            });
        }
      }
    } finally {
      reading = false;
    }
  }
  return {
    subscribe: (fn: () => void) => {
      listeners.add(fn);
      return () => {
        listeners.delete(fn);
      };
    },
    getSnapshot: () => snapshot,
    setFrame: (next: AccessibilityFrame | null) => {
      const changed = frame?.key !== next?.key || frame?.active !== next?.active;
      frame = next;
      if (!changed) return;
      wanted = next?.active ? capture(next) : null;
      emit({ key: next?.key ?? '', report: null, loading: !!wanted, error: '' });
      void drain();
    },
    refresh: (captured: AccessibilityFrame) => {
      if (!same(captured) || snapshot.key !== captured.key || snapshot.loading)
        throw Error('Wait for the current accessibility check.');
      wanted = capture(frame!);
      emit({ ...snapshot, report: null, loading: true, error: '' });
      void drain();
    },
    navigate: (captured: AccessibilityFrame, index: number, cell?: string) => {
      if (!same(captured) || snapshot.key !== captured.key || snapshot.loading || !snapshot.report)
        throw Error('The original accessibility report changed.');
      const finding: AccessibilityFinding | undefined = snapshot.report.findings[index],
        target = cell ?? finding?.cell;
      if (!finding || !target || (target !== finding.cell && !finding.relatedCells.includes(target)))
        throw Error('Choose a current finding location.');
      if (!frame!.visibleSheets.includes(finding.sheet) || !captured.visibleSheets.includes(finding.sheet))
        throw Error('This finding is on a hidden worksheet. Unhide it separately to navigate.');
      captured.navigate(finding.sheet, target);
    },
    dispose: () => {
      disposed = true;
      frame = null;
      wanted = null;
      listeners.clear();
    },
  };
}
