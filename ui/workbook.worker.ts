import * as engine from '../core/calculation-browser.mjs';
import { projectCellCommands, projectWorkbookCommands } from '../core/workbook-projection.mjs';
import type { WorkbookCell, WorkbookCommand, WorkbookEdit, WorkbookView } from '../core/contracts';
const scope = globalThis as unknown as {
  onmessage: ((event: MessageEvent) => void) | null;
  postMessage: (value: unknown) => void;
};
type Session = Awaited<ReturnType<typeof engine.createCalculationSession<WorkbookView>>>;
let current: WorkbookView | undefined, session: Session | undefined;
let formulaRefs: Array<{
  sheet: string;
  cell: string;
}> = [];
function indexFormulas(): Record<string, number> {
  const counts: Record<string, number> = {};
  formulaRefs = [];
  for (const sheet of current!.sheets) {
    counts[sheet.name] = 0;
    for (const cell in sheet.cells)
      if (sheet.cells[cell].formula || sheet.cells[cell].formulaType) {
        counts[sheet.name]++;
        formulaRefs.push({ sheet: sheet.name, cell });
      }
  }
  return counts;
}
function assign(
  changes: Array<{
    sheet: string;
    cell: string;
    value: WorkbookCell;
  }>
): void {
  for (const change of changes) {
    const sheet = current!.sheets.find((item) => item.name === change.sheet)!;
    sheet.cells[change.cell] = change.value;
    const match = /^([A-Z]+)([1-9]\d*)$/.exec(change.cell)!;
    let column = 0;
    for (const char of match[1]) column = column * 26 + char.charCodeAt(0) - 64;
    sheet.rows = Math.max(sheet.rows, Number(match[2]));
    sheet.columns = Math.max(sheet.columns, column);
  }
}
let queue = Promise.resolve();
scope.onmessage = ({
  data,
}: MessageEvent<{
  id: number;
  view?: WorkbookView;
  baseRevision: string;
  commands: WorkbookCommand[];
  revision: string;
  prepare?: boolean;
}>) => {
  queue = queue.then(async () => {
    try {
      if (data.view) {
        session?.dispose();
        session = undefined;
        current = data.view;
        indexFormulas();
      }
      if (!current || current.revision !== data.baseRevision) throw Error('Workbook calculation base revision changed');
      if (data.prepare) {
        if (!formulaRefs.length) {
          scope.postMessage({
            id: data.id,
            prepared: true,
            formulaCounts: Object.fromEntries(current.sheets.map((sheet) => [sheet.name, 0])),
            edits: [],
          });
          return;
        }
        try {
          session = await engine.createCalculationSession(current, { takeOwnership: true });
          current = session.view;
        } catch (error) {
          current = await engine.calculateWorkbook(current);
          if (!current) throw error;
        }
        scope.postMessage({ id: data.id, view: current, edits: [] });
        return;
      }
      const bounded = projectCellCommands(current, data.commands);
      const incremental =
        (session || !formulaRefs.length) &&
        bounded &&
        bounded.changes.every((change) => {
          const previous = current!.sheets.find((sheet) => sheet.name === change.sheet)?.cells[change.cell];
          return !previous?.formula && !previous?.formulaType && !change.value.formula && !change.value.formulaType;
        });
      let edits: WorkbookEdit[];
      if (incremental && bounded) {
        assign(bounded.changes);
        current.revision = data.revision;
        const results = session ? await session.update(current, bounded.changes) : [];
        const patches = new Map<
          string,
          {
            sheet: string;
            cell: string;
            value: WorkbookCell;
          }
        >();
        for (const change of [...bounded.changes, ...results]) {
          const sheet = current.sheets.find((sheet) => sheet.name === change.sheet)!;
          sheet.cells[change.cell] = change.value as WorkbookCell;
          patches.set(
            `${change.sheet}!${change.cell}`,
            change as {
              sheet: string;
              cell: string;
              value: WorkbookCell;
            }
          );
        }
        for (const sheet of current.sheets) sheet.calculationRevision = data.revision;
        scope.postMessage({
          id: data.id,
          revision: data.revision,
          patches: [...patches.values()],
          edits: bounded.edits,
        });
        return;
      }
      session?.dispose();
      session = undefined;
      let projected: {
        view: WorkbookView;
        edits: WorkbookEdit[];
      };
      if (bounded) {
        assign(bounded.changes);
        projected = { view: current, edits: bounded.edits };
      } else
        projected = (await projectWorkbookCommands(current, data.commands, engine)) as {
          view: WorkbookView;
          edits: WorkbookEdit[];
        };
      projected.view.revision = data.revision;
      edits = projected.edits;
      current = projected.view;
      indexFormulas();
      if (formulaRefs.length) {
        try {
          session = await engine.createCalculationSession(current, { takeOwnership: true });
          current = session.view;
        } catch (error) {
          current = await engine.calculateWorkbook(current);
          if (!current) throw error;
        }
      }
      if (bounded) {
        const patches = new Map<
          string,
          {
            sheet: string;
            cell: string;
            value: WorkbookCell;
          }
        >();
        for (const ref of [...bounded.changes, ...formulaRefs]) {
          const value = current.sheets.find((sheet) => sheet.name === ref.sheet)!.cells[ref.cell];
          patches.set(`${ref.sheet}!${ref.cell}`, { sheet: ref.sheet, cell: ref.cell, value });
        }
        for (const sheet of current.sheets) sheet.calculationRevision = data.revision;
        scope.postMessage({ id: data.id, revision: data.revision, patches: [...patches.values()], edits });
        return;
      }
      scope.postMessage({ id: data.id, view: current, edits });
    } catch (error) {
      session?.dispose();
      session = undefined;
      current = undefined;
      scope.postMessage({ id: data.id, error: String(error) });
    }
  });
};
