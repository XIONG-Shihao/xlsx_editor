import { INK_LIMITS, type InkPenId } from '../core/ink-stroke.mjs';
import type { WorkbookObjectCommand } from '../core/contracts';
import { anchoredInkStroke, inkWorksheetPoint, type InkViewport } from './inkGeometry';
import type { ObjectAxes } from './worksheetObjectGeometry';
export interface InkGestureFrame {
  key: string;
  sheet: string;
  active: boolean;
  reason?: string;
  pen: InkPenId;
  axes: ObjectAxes;
  view: InkViewport;
  accept: (commands: WorkbookObjectCommand[]) => Promise<unknown>;
}
export interface InkGestureSnapshot {
  key: string;
  points: readonly (readonly [number, number])[];
  pending: boolean;
  error: string | null;
}
interface Draft {
  frame: InkGestureFrame;
  pointer: number;
  points: [number, number][];
}
/** One draft and one physical write lane. A frame carries the original owner/head and viewport. */
export function createInkGestureController() {
  let frame: InkGestureFrame | null = null,
    draft: Draft | null = null,
    flight = false,
    disposed = false;
  let snapshot: InkGestureSnapshot = { key: '', points: [], pending: false, error: null };
  const listeners = new Set<() => void>();
  const publish = (error: string | null = null) => {
    snapshot = {
      key: frame?.key ?? '',
      points: draft ? draft.points.map((p) => [...p] as [number, number]) : [],
      pending: flight,
      error,
    };
    for (const f of listeners) f();
  };
  const current = (captured: InkGestureFrame) =>
    !disposed && !!frame && frame.active && !frame.reason && frame.key === captured.key;
  const cancel = () => {
    if (draft) {
      draft = null;
      publish();
    }
  };
  const add = (d: Draft, x: number, y: number) => {
    const p = inkWorksheetPoint(d.frame.view, x, y);
    if (!p) throw Error('Keep the stroke inside the worksheet.');
    const last = d.points.at(-1);
    if (last && last[0] === p[0] && last[1] === p[1]) return;
    if (d.points.length >= INK_LIMITS.points) throw Error('A stroke can contain at most 512 points.');
    d.points.push(p);
  };
  return {
    subscribe: (f: () => void) => {
      listeners.add(f);
      return () => {
        listeners.delete(f);
      };
    },
    getSnapshot: () => snapshot,
    setFrame(next: InkGestureFrame | null) {
      if (disposed) return;
      const changed = frame?.key !== next?.key || frame?.active !== next?.active || frame?.reason !== next?.reason;
      frame = next;
      if (draft && !current(draft.frame)) draft = null;
      if (changed) publish();
    },
    start(pointer: number, x: number, y: number) {
      if (!frame || !current(frame) || flight || draft) return false;
      const captured = { ...frame, view: { ...frame.view } };
      try {
        const d: Draft = { frame: captured, pointer, points: [] };
        add(d, x, y);
        draft = d;
        publish();
        return true;
      } catch (e) {
        publish(e instanceof Error ? e.message : String(e));
        return false;
      }
    },
    move(pointer: number, x: number, y: number) {
      const d = draft;
      if (!d || d.pointer !== pointer) return false;
      if (!current(d.frame)) {
        cancel();
        return false;
      }
      try {
        add(d, x, y);
        publish();
        return true;
      } catch (e) {
        draft = null;
        publish(e instanceof Error ? e.message : String(e));
        return false;
      }
    },
    async finish(pointer: number, x: number, y: number) {
      const d = draft;
      if (!d || d.pointer !== pointer || flight) return { accepted: false, current: false };
      if (!current(d.frame)) {
        cancel();
        return { accepted: false, current: false };
      }
      try {
        add(d, x, y);
        const value = anchoredInkStroke(d.points, d.frame.pen, d.frame.axes);
        const command: WorkbookObjectCommand = {
          operation: 'createInkStroke',
          sheet: d.frame.sheet,
          objectId: crypto.randomUUID(),
          expectedObjectRevision: null,
          stroke: value.stroke,
          anchor: value.anchor,
          altText: '',
        };
        draft = null;
        flight = true;
        publish();
        await d.frame.accept([command]);
        return { accepted: true, current: current(d.frame) };
      } catch (e) {
        if (current(d.frame)) publish(e instanceof Error ? e.message : String(e));
        return { accepted: false, current: current(d.frame) };
      } finally {
        flight = false;
        draft = null;
        publish(snapshot.error);
      }
    },
    cancel,
    dispose() {
      disposed = true;
      frame = null;
      draft = null;
      publish();
      listeners.clear();
    },
  };
}
