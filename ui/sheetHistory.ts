import type { CsvIndex } from '../core/csv-workbook.mjs';
import type { WorkbookEdit, WorkbookHistoryEdit, WorkbookSheet, WorkbookView } from '../core/contracts';
import { applyGridEdit } from './sheetModel';
import { clearHyperlinksCommand } from '../core/clear-hyperlinks.mjs';

type CalculationHistoryEdit = Extract<
  WorkbookHistoryEdit,
  { type: 'recalculate' | 'calculateSheet' | 'setCalculationOptions' }
>;
function calculationEdit(edit: WorkbookHistoryEdit): edit is CalculationHistoryEdit {
  return edit.type === 'recalculate' || edit.type === 'calculateSheet' || edit.type === 'setCalculationOptions';
}
function retainedEdit(edit: WorkbookHistoryEdit): WorkbookHistoryEdit {
  if (edit.type === 'clearHyperlinks') {
    const command = clearHyperlinksCommand(edit);
    return Object.freeze({ ...command, ranges: Object.freeze(command.ranges) as unknown as string[] });
  }
  if ('operation' in edit && edit.operation === 'removeDuplicates')
    return Object.freeze({ ...edit, columns: Object.freeze([...edit.columns]) as unknown as string[] });
  if ('operation' in edit && edit.operation === 'removeFilter') return Object.freeze({ ...edit });
  if (!calculationEdit(edit)) return edit;
  const allowed =
    edit.type === 'recalculate'
      ? ['type']
      : edit.type === 'calculateSheet'
        ? ['type', 'sheet']
        : ['type', 'mode', 'recalculateBeforeSave'];
  if (
    Object.keys(edit).some((key) => !allowed.includes(key)) ||
    (edit.type === 'calculateSheet' && (typeof edit.sheet !== 'string' || !edit.sheet)) ||
    (edit.type === 'setCalculationOptions' &&
      (!['auto', 'manual'].includes(edit.mode) ||
        (edit.recalculateBeforeSave !== undefined && typeof edit.recalculateBeforeSave !== 'boolean')))
  )
    throw Error('Invalid calculation history command');
  return Object.freeze({ ...edit });
}
function acceptedCalculationView(edits: WorkbookHistoryEdit[], view?: WorkbookView): void {
  if (
    edits.some((edit) => edit.type === 'clearHyperlinks') &&
    (!view ||
      !Array.isArray(view.sheets) ||
      typeof view.revision !== 'string' ||
      !view.revision ||
      view.storage?.kind !== 'rust-indexed-v1' ||
      typeof view.storage.sessionId !== 'string' ||
      !view.storage.sessionId ||
      typeof view.storage.state !== 'string' ||
      !view.storage.state)
  )
    throw Error(
      'Clear Hyperlinks history requires its accepted native XLSX owner view; it cannot be projected in browser history'
    );
  if (
    edits.some(
      (edit) => 'operation' in edit && (edit.operation === 'removeDuplicates' || edit.operation === 'removeFilter')
    ) &&
    (!view || !Array.isArray(view.sheets) || typeof view.revision !== 'string' || !view.revision)
  )
    throw Error(
      'Duplicate/filter history requires its accepted workbook view; it cannot be projected in browser history'
    );
  if (
    edits.some(calculationEdit) &&
    (!view || !Array.isArray(view.sheets) || typeof view.revision !== 'string' || !view.revision)
  )
    throw Error(
      'Calculation history requires its accepted workbook view; it cannot be calculated in the browser history'
    );
}

interface JournalNode {
  edit?: WorkbookHistoryEdit;
  checkpoint?: WorkbookHistoryEdit[];
  previous?: JournalNode;
}

export interface SheetFrame {
  sheets: WorkbookSheet[];
  view?: WorkbookView;
  csvIndex?: CsvIndex;
  readonly edits: WorkbookHistoryEdit[];
  id: number;
  focus?: { sheet: string; cell: string };
}

/** Session-local undo frames share immutable cells and a persistent edit journal. */
export class SheetHistory {
  private past: SheetFrame[] = [];
  private future: SheetFrame[] = [];
  private sequence = 0;
  private saved = 0;
  private editsSinceCompaction = 0;
  private journal?: JournalNode;
  private materializedJournal?: JournalNode;
  private materializedEdits: WorkbookHistoryEdit[] = [];
  private journals = new WeakMap<SheetFrame, JournalNode | undefined>();
  current: SheetFrame;
  constructor(sheets: WorkbookSheet[], edits: WorkbookHistoryEdit[] = [], dirty = false, view?: WorkbookView) {
    acceptedCalculationView(edits, view);
    for (const edit of edits) this.journal = { edit: retainedEdit(edit), previous: this.journal };
    this.current = this.frame(view?.sheets ?? sheets, this.journal, 0);
    this.current.view = view;
    this.saved = dirty ? -1 : 0;
    this.compactJournal();
  }
  retainedNativeStates(sessionId: string): string[] {
    return [
      ...new Set(
        [...this.past, this.current, ...this.future].flatMap((frame) =>
          frame.view?.storage?.sessionId === sessionId ? [frame.view.storage.state] : []
        )
      ),
    ];
  }
  get canUndo(): boolean {
    return this.past.length > 0;
  }
  get canRedo(): boolean {
    return this.future.length > 0;
  }
  get dirty(): boolean {
    return this.current.id !== this.saved;
  }
  markSaved(): void {
    this.saved = this.current.id;
    this.compactJournal();
  }
  edit(edit: WorkbookEdit): SheetFrame {
    return this.batch([edit]);
  }
  batch(edits: WorkbookHistoryEdit[], sheets?: WorkbookSheet[], view?: WorkbookView): SheetFrame {
    acceptedCalculationView(edits, view);
    const retained = edits.map(retainedEdit);
    let projected = view?.sheets ?? sheets ?? this.current.sheets;
    if (!sheets && !view)
      for (const edit of retained) {
        if ('operation' in edit && (edit.operation === 'removeDuplicates' || edit.operation === 'removeFilter'))
          throw Error('Duplicate/filter history requires its accepted workbook view');
        if (calculationEdit(edit)) throw Error('Calculation history requires its accepted workbook view');
        if (edit.type === 'importCells') throw Error('Staged paste history requires its accepted workbook view');
        projected = projected.map((s) => (s.name === edit.sheet ? applyGridEdit(s, edit) : s));
      }
    this.past.push(this.current);
    if (this.past.length > 100) this.past.shift();
    this.future = [];
    for (const edit of retained) this.journal = { edit, previous: this.journal };
    const last = retained.at(-1);
    const focus =
      last && calculationEdit(last)
        ? this.current.focus
        : last?.type === 'importCells'
          ? { sheet: last.shape.sheet, cell: last.shape.start }
          : last?.operation === 'removeDuplicates'
            ? { sheet: last.sheet, cell: last.range.split(':')[0] }
            : last?.type === 'clearHyperlinks'
              ? { sheet: last.sheet, cell: last.ranges?.[0]?.split(':')[0] ?? 'A1' }
              : last
                ? { sheet: last.sheet, cell: last.cell ?? last.range?.split(':')[0] ?? 'A1' }
                : undefined;
    this.current = this.frame(projected, this.journal, ++this.sequence, focus);
    this.current.view = view;
    if (++this.editsSinceCompaction >= 100) this.compactJournal();
    return this.current;
  }
  undo(): SheetFrame {
    const focus = this.current.focus;
    const next = this.past.pop();
    if (next) {
      this.future.push(this.current);
      this.current = next;
      this.journal = this.frameJournal(next);
    }
    return { ...this.current, focus };
  }
  redo(): SheetFrame {
    const next = this.future.pop();
    if (next) {
      this.past.push(this.current);
      this.current = next;
      this.journal = this.frameJournal(next);
    }
    return this.current;
  }

  /** Replace the prefix older than the undo window with one compact checkpoint.
   * Structural operations remain ordered barriers; never flatten them into cell writes.
   * All retained frames descend from this node, so undo/redo keep their original meaning.
   */
  private compactJournal(): void {
    const oldest = this.past[0] ?? this.current;
    const node = this.frameJournal(oldest);
    if (!node) return;
    const checkpoint = this.materialize(node);
    node.checkpoint = checkpoint;
    delete node.edit;
    delete node.previous;
    this.materializedJournal = undefined;
    this.materializedEdits = [];
    this.editsSinceCompaction = 0;
  }

  private frame(
    sheets: WorkbookSheet[],
    journal: JournalNode | undefined,
    id: number,
    focus?: { sheet: string; cell: string }
  ): SheetFrame {
    const history = this;
    const result: SheetFrame = {
      sheets,
      id,
      focus,
      get edits(): WorkbookHistoryEdit[] {
        return history.materialize(journal);
      },
    };
    this.journals.set(result, journal);
    return result;
  }

  private frameJournal(frame: SheetFrame): JournalNode | undefined {
    return this.journals.get(frame);
  }

  private materialize(journal: JournalNode | undefined): WorkbookHistoryEdit[] {
    if (journal === this.materializedJournal) return this.materializedEdits;
    const raw: WorkbookHistoryEdit[] = [];
    let prefix: WorkbookHistoryEdit[] = [];
    for (let node = journal; node; node = node.previous) {
      if (node.checkpoint) {
        prefix = node.checkpoint;
        break;
      }
      if (node.edit) raw.push(node.edit);
    }
    raw.reverse();

    const compacted: WorkbookHistoryEdit[] = [];
    const cells = new Map<string, number>();
    for (const edit of [...prefix, ...raw]) {
      if (
        calculationEdit(edit) ||
        edit.type === 'importCells' ||
        edit.operation ||
        !edit.cell ||
        edit.expectedValue !== undefined ||
        edit.expectedFormula !== undefined ||
        edit.expectedRawValue !== undefined
      ) {
        compacted.push(edit);
        cells.clear();
        continue;
      }
      const key = `${edit.sheet}\0${edit.cell}`;
      const index = cells.get(key);
      if (index === undefined) {
        cells.set(key, compacted.length);
        compacted.push(edit);
        continue;
      }
      const merged: WorkbookEdit = { ...(compacted[index] as WorkbookEdit), ...edit };
      if (edit.value !== undefined) delete merged.formula;
      if (edit.formula !== undefined) delete merged.value;
      compacted[index] = merged;
    }
    this.materializedJournal = journal;
    this.materializedEdits = compacted;
    return compacted;
  }
}
