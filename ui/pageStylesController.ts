import type { WorkbookCommand, WorkbookView, WorkbookSheet } from '../core/contracts';
import { pageThemeCommand } from '../core/page-theme-command.mjs';
import { backgroundCommand } from '../core/background-command.mjs';
import { pageThemeSample, worksheetBackgroundSample } from '../core/page-style-samples.mjs';
export interface PageStylesFrame {
  key: string;
  family: string;
  sheet: string;
  active: boolean;
  writable: boolean;
  reason?: string;
  theme: NonNullable<WorkbookView['pageTheme']>;
  background: WorkbookSheet['background'];
  apply: (commands: WorkbookCommand[]) => Promise<void>;
}
export interface PageStylesSnapshot {
  key: string;
  theme: string | null;
  background: string | null;
  pending: boolean;
  error: string;
  notice: string;
}
/** Local sample previews never write. A physical write keeps its original owner,
 * captured revisions and one lane, even when the retained pane changes scope. */
export function createPageStylesController() {
  let frame: PageStylesFrame | null = null,
    working = false,
    disposed = false;
  let snapshot: PageStylesSnapshot = { key: '', theme: null, background: null, pending: false, error: '', notice: '' };
  const listeners = new Set<() => void>();
  const emit = (next: PageStylesSnapshot) => {
      snapshot = next;
      for (const listener of listeners) listener();
    },
    same = (f: PageStylesFrame) => !disposed && frame?.key === f.key && frame.active;
  function target(captured: PageStylesFrame, write = false) {
    if (!same(captured) || snapshot.key !== captured.key)
      throw Error('The original page-style preview changed. Choose a sample again.');
    if (working) throw Error('Wait for the page-style operation to finish.');
    if (write && (!frame!.writable || !captured.writable)) throw Error(frame!.reason ?? 'This workbook is read-only.');
    return {
      ...frame!,
      theme: structuredClone(frame!.theme),
      background: frame!.background ? structuredClone(frame!.background) : undefined,
    };
  }
  async function write(current: PageStylesFrame, command: WorkbookCommand, notice: string) {
    working = true;
    emit({ ...snapshot, pending: true, error: '', notice: '' });
    try {
      await current.apply([command]);
      if (frame?.active && frame.family === current.family) emit({ ...snapshot, notice });
      return { accepted: true, current: same(current) };
    } catch (e) {
      if (same(current)) emit({ ...snapshot, error: String(e instanceof Error ? e.message : e) });
      throw e;
    } finally {
      working = false;
      if (!disposed) emit({ ...snapshot, pending: false });
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
    setFrame: (next: PageStylesFrame | null) => {
      const changed = frame?.key !== next?.key || frame?.active !== next?.active;
      frame = next;
      if (changed)
        emit({
          key: next?.key ?? '',
          theme: next?.theme.sample ?? null,
          background: next?.background?.sample ?? null,
          pending: working,
          error: '',
          notice: '',
        });
    },
    choose: (captured: PageStylesFrame, mode: 'themes' | 'backgrounds', id: string) => {
      target(captured);
      if (mode === 'themes') pageThemeSample(id);
      else worksheetBackgroundSample(id);
      emit({ ...snapshot, [mode === 'themes' ? 'theme' : 'background']: id, error: '', notice: '' });
    },
    reset: (captured: PageStylesFrame) => {
      const current = target(captured);
      emit({
        ...snapshot,
        theme: current.theme.sample,
        background: current.background?.sample ?? null,
        error: '',
        notice: '',
      });
    },
    theme: async (captured: PageStylesFrame) => {
      const current = target(captured, true);
      if (!current.theme.canApply) throw Error(current.theme.protectedReason ?? 'This imported theme is protected.');
      if (!snapshot.theme) throw Error('Choose a theme sample first.');
      const command = pageThemeCommand({
        type: 'applyPageTheme',
        sample: snapshot.theme,
        expectedThemeRevision: current.theme.revision,
      }) as WorkbookCommand;
      return write(current, command, 'Theme applied.');
    },
    background: async (captured: PageStylesFrame, remove = false) => {
      const current = target(captured, true);
      if (current.background && !current.background.canEdit)
        throw Error(current.background.protectedReason ?? 'This imported background is protected.');
      if (remove && !current.background?.sample) throw Error('There is no owned background to remove.');
      if (!remove && !snapshot.background) throw Error('Choose a background sample first.');
      const command = backgroundCommand({
        type: remove ? 'removeSheetBackground' : 'setSheetBackground',
        sheet: current.sheet,
        expectedBackgroundRevision: current.background?.revision ?? null,
        ...(!remove ? { sample: snapshot.background } : {}),
      }) as WorkbookCommand;
      return write(current, command, remove ? 'Background removed.' : 'Background applied.');
    },
    dispose: () => {
      disposed = true;
      frame = null;
      listeners.clear();
    },
  };
}
