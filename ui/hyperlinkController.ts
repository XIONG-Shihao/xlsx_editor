import type {
  WorkbookHyperlinkDefinition,
  WorkbookHyperlinkDestination,
  WorkbookHyperlinkCommand,
  WorkbookHyperlinkReceipt,
  WorkbookCell,
  WorkbookCommand,
} from '../core/contracts';
import { hyperlinkCommand, validateHyperlink, hyperlinkDisplayText } from '../core/hyperlinks.mjs';
export type HyperlinkFollowResult = { kind: 'url'; opened: true } | { kind: 'internal'; sheet: string; cell: string };
export interface HyperlinkFrame {
  key: string;
  family: string;
  sheet: string;
  cell: string;
  receipt: WorkbookHyperlinkReceipt;
  active: boolean;
  writable: boolean;
  reason?: string;
  read: () => Promise<WorkbookHyperlinkDefinition>;
  apply: (commands: WorkbookCommand[]) => Promise<void>;
  readCell?: () => Promise<WorkbookCell | null>;
  follow?: (revision: string) => Promise<HyperlinkFollowResult>;
  navigate?: (sheet: string, cell: string) => void;
}
export interface HyperlinkSnapshot {
  key: string;
  definition: WorkbookHyperlinkDefinition | null;
  cell: WorkbookCell | null;
  labelAvailable: boolean;
  loading: boolean;
  pending: boolean;
  error: string;
  notice: string;
}
/** One physical inspection and one accepted action; retired replies stay retired. */
export function createHyperlinkController() {
  let frame: HyperlinkFrame | null = null,
    wanted: HyperlinkFrame | null = null,
    reading = false,
    working = false,
    disposed = false;
  let snapshot: HyperlinkSnapshot = {
    key: '',
    definition: null,
    cell: null,
    labelAvailable: false,
    loading: false,
    pending: false,
    error: '',
    notice: '',
  };
  const listeners = new Set<() => void>();
  const emit = (value: HyperlinkSnapshot) => {
      snapshot = value;
      for (const listener of listeners) listener();
    },
    same = (captured: HyperlinkFrame) => !disposed && frame?.key === captured.key && frame.active;
  const errorText = (error: unknown) => String(error instanceof Error ? error.message : error);
  async function drain() {
    if (reading || disposed) return;
    reading = true;
    try {
      while (wanted && !disposed) {
        const captured = wanted;
        wanted = null;
        if (!same(captured)) continue;
        emit({
          ...snapshot,
          key: captured.key,
          definition: null,
          cell: null,
          labelAvailable: false,
          loading: true,
          error: '',
        });
        try {
          const result = await captured.read();
          const cell = captured.readCell && same(captured) ? structuredClone(await captured.readCell()) : null;
          if (same(captured)) {
            validateHyperlink(result, captured.receipt, { sheet: captured.sheet, cell: captured.cell });
            emit({
              ...snapshot,
              definition: result,
              cell,
              labelAvailable: !!captured.readCell,
              loading: false,
              error: '',
            });
          }
        } catch (error) {
          if (same(captured))
            emit({
              ...snapshot,
              definition: null,
              cell: null,
              labelAvailable: false,
              loading: false,
              error: errorText(error),
            });
        }
      }
    } finally {
      reading = false;
    }
  }
  function target(captured: HyperlinkFrame) {
    if (!same(captured) || snapshot.key !== captured.key)
      throw Error('The original cell or workbook changed. Select it again.');
    if (working || snapshot.loading) throw Error('Wait for the link operation to finish.');
    if (!snapshot.definition) throw Error('Load the selected cell link first.');
    return snapshot.definition;
  }
  return {
    subscribe: (listener: () => void) => {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
    getSnapshot: () => snapshot,
    setFrame: (next: HyperlinkFrame | null) => {
      const changed = frame?.key !== next?.key || frame?.active !== next?.active;
      frame = next;
      if (!changed) return;
      wanted = next?.active ? { ...next, receipt: { ...next.receipt } } : null;
      emit({
        key: next?.key ?? '',
        definition: null,
        cell: null,
        labelAvailable: false,
        loading: !!wanted,
        pending: working,
        error: '',
        notice: '',
      });
      void drain();
    },
    refresh: () => {
      if (frame?.active) {
        wanted = { ...frame, receipt: { ...frame.receipt } };
        void drain();
      }
    },
    write: async (
      captured: HyperlinkFrame,
      link: WorkbookHyperlinkDestination | null,
      tooltip = '',
      label?: string
    ) => {
      const definition = target(captured);
      if (!frame!.writable || !captured.writable || !definition.canEdit)
        throw Error(frame!.reason ?? definition.protectedReason ?? 'This link is read-only.');
      const command = hyperlinkCommand(
        link === null
          ? {
              type: 'removeHyperlink',
              sheet: captured.sheet,
              cell: captured.cell,
              expectedHyperlinkRevision: definition.hyperlinkRevision,
            }
          : {
              type: 'setHyperlink',
              sheet: captured.sheet,
              cell: captured.cell,
              expectedHyperlinkRevision: definition.hyperlinkRevision,
              link,
              tooltip,
            }
      );
      const commands: WorkbookCommand[] = [command];
      if (link !== null && label !== undefined && snapshot.labelAvailable) {
        const cell = snapshot.cell;
        const plain =
          !cell ||
          (!cell.formula &&
            !cell.formulaType &&
            !cell.spill &&
            (cell.dataType === 'text' || cell.dataType === 'blank' || cell.dataType === undefined) &&
            (cell.value == null || typeof cell.value === 'string'));
        if (!plain) throw Error('Use the worksheet to edit this cell value.');
        hyperlinkDisplayText(label);
        const original = cell?.value ?? null;
        const value =
          !definition.link && (original == null || original === '') && label === ''
            ? link.kind === 'url'
              ? link.target
              : link.sheet + '!' + link.cell
            : label;
        if (value !== (original ?? ''))
          commands.push({
            sheet: captured.sheet,
            cell: captured.cell,
            value,
            expectedValue: original,
            expectedFormula: null,
          });
      }
      working = true;
      emit({ ...snapshot, pending: true, error: '', notice: '' });
      try {
        await captured.apply(commands);
        if (frame?.active && frame.family === captured.family)
          emit({ ...snapshot, notice: link === null ? 'Link removed.' : 'Link updated.' });
        return { accepted: true, current: same(captured), key: captured.key };
      } catch (error) {
        if (same(captured)) emit({ ...snapshot, error: errorText(error) });
        throw error;
      } finally {
        working = false;
        emit({ ...snapshot, pending: false });
        if (wanted) void drain();
      }
    },
    follow: async (captured: HyperlinkFrame) => {
      const definition = target(captured);
      if (!definition.canEdit || !definition.link || !definition.hyperlinkRevision || !captured.follow)
        throw Error('This link cannot be opened in this host.');
      working = true;
      emit({ ...snapshot, pending: true, error: '', notice: '' });
      try {
        const result = await captured.follow(definition.hyperlinkRevision);
        if (same(captured)) {
          if (result.kind === 'internal') captured.navigate?.(result.sheet, result.cell);
          else emit({ ...snapshot, notice: 'Opened link in the browser.' });
        }
        return { completed: true, current: same(captured) };
      } catch (error) {
        if (same(captured)) emit({ ...snapshot, error: errorText(error) });
        throw error;
      } finally {
        working = false;
        emit({ ...snapshot, pending: false });
        if (wanted) void drain();
      }
    },
    dispose: () => {
      disposed = true;
      wanted = null;
      frame = null;
      listeners.clear();
    },
  };
}
export type HyperlinkController = ReturnType<typeof createHyperlinkController>;
