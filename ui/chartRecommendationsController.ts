import { createChartRecommendations } from '../core/chart-authoring-model.mjs';
import type { WorkbookChartModel } from '../core/contracts';
export interface ChartRecommendationFrame {
  key: string;
  active: boolean;
  reason?: string;
  sheet: string;
  range: string;
  view: { sheets: Array<{ name: string; hiddenRows?: number[] }> };
  read: (sheet: string, range: string) => Promise<Record<string, unknown>>;
  apply: (command: import('../core/contracts').WorkbookObjectCommand) => Promise<unknown>;
}
export type ChartRecommendations = Awaited<ReturnType<typeof createChartRecommendations>>;
export interface ChartRecommendationSnapshot {
  key: string;
  loading: boolean;
  error: string;
  result: ChartRecommendations | null;
}
/** One source read joins retirement; no old reply may populate a later range/frame. */
export function createChartRecommendationsController() {
  let frame: ChartRecommendationFrame | null = null,
    wanted: ChartRecommendationFrame | null = null,
    reading = false,
    disposed = false,
    snapshot: ChartRecommendationSnapshot = { key: '', loading: false, error: '', result: null };
  const listeners = new Set<() => void>(),
    emit = (next: ChartRecommendationSnapshot) => {
      snapshot = next;
      for (const fn of listeners) fn();
    },
    same = (f: ChartRecommendationFrame) => !disposed && frame?.active && frame.key === f.key;
  async function drain() {
    if (reading || disposed) return;
    reading = true;
    try {
      while (wanted && !disposed) {
        const captured = wanted;
        wanted = null;
        if (!same(captured)) continue;
        try {
          const result = await createChartRecommendations({
            sheet: captured.sheet,
            range: captured.range,
            view: captured.view,
            readRange: captured.read,
          });
          if (same(captured)) emit({ key: captured.key, loading: false, error: '', result: structuredClone(result) });
        } catch (error) {
          if (same(captured))
            emit({
              key: captured.key,
              loading: false,
              error: String(error instanceof Error ? error.message : error),
              result: null,
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
    setFrame: (next: ChartRecommendationFrame | null) => {
      const changed = frame?.key !== next?.key || frame?.active !== next?.active;
      frame = next;
      if (!changed) return;
      wanted = next?.active ? { ...next, view: structuredClone(next.view) } : null;
      emit({ key: next?.key ?? '', loading: !!wanted, error: '', result: null });
      void drain();
    },
    refresh: () => {
      if (frame?.active) {
        wanted = { ...frame, view: structuredClone(frame.view) };
        emit({ key: frame.key, loading: true, error: '', result: null });
        void drain();
      }
    },
    capture: (captured: ChartRecommendationFrame, id: string, style: WorkbookChartModel['style']) => {
      if (!same(captured) || snapshot.key !== captured.key || snapshot.loading || !snapshot.result)
        throw Error('The chart preview changed. Preview the current selection again.');
      const p = snapshot.result.proposals.find((p) => p.id === id);
      if (!p) throw Error('Choose a current chart proposal.');
      return structuredClone({ ...p.chart, style: style ?? p.chart.style });
    },
    dispose: () => {
      disposed = true;
      frame = null;
      wanted = null;
      listeners.clear();
    },
  };
}
