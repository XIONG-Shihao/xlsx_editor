import type { WorkbookCell, WorkbookCommand, WorkbookHistoryEdit, WorkbookView } from '../core/contracts';
import { projectCellCommands } from '../core/workbook-projection.mjs';
import {
  applyCalculatedPatches,
  formulaCount,
  plainSheets,
  prepareCellRecord,
  rememberFormulaCount,
} from './sheetModel';
import { registerWorkbookDiagnostics } from './workbookDiagnostics';
export type WorkbookMutationOwner = (
  view: WorkbookView,
  commands: WorkbookCommand[]
) => Promise<{ view: WorkbookView; edits: WorkbookHistoryEdit[] }>;
export class WorkbookWorker {
  private owner?: WorkbookMutationOwner;
  constructor(owner?: WorkbookMutationOwner) {
    this.owner = owner;
    this.observeDiagnostics();
  }
  private releaseDiagnostics?: () => void;
  private observeDiagnostics(): void {
    this.releaseDiagnostics ??= registerWorkbookDiagnostics('renderer-worker', () => ({
      liveBrowserWorkers: this.worker ? 1 : 0,
      pendingOwnerRequests: this.ownerPending.size,
      pendingWorkerRequests: this.pending.size,
      acceptedRevision: this.acceptedRevision ? 1 : 0,
    }));
  }
  private ownerGeneration = 0;
  private ownerPending = new Set<(error: Error) => void>();
  private worker?: Worker;
  private idleTimer?: ReturnType<typeof setTimeout>;
  private armIdle(): void {
    clearTimeout(this.idleTimer);
    if (!this.worker || this.pending.size || this.ownerPending.size) return;
    const worker = this.worker;
    this.idleTimer = setTimeout(() => {
      if (this.worker !== worker || this.pending.size || this.ownerPending.size) return;
      worker.terminate();
      this.worker = undefined;
      this.acceptedRevision = undefined;
      // The renderer owns the accepted view and dirty buffer. The next request
      // rehydrates from that view; retiring this cache never publishes a file.
    }, 15000);
  }
  private sequence = 0;
  private session = crypto.randomUUID();
  private acceptedRevision?: string;
  private pending = new Map<
    number,
    {
      base: WorkbookView;
      resolve: (value: { view: WorkbookView; edits: WorkbookHistoryEdit[] }) => void;
      reject: (error: Error) => void;
    }
  >();
  private ensureWorker(): Worker {
    clearTimeout(this.idleTimer);
    if (this.worker) return this.worker;
    const worker = new Worker(new URL('./workbook.worker.ts', import.meta.url), { type: 'module' });
    this.worker = worker;
    worker.onmessage = ({ data }) => {
      if (this.worker !== worker) return;
      const request = this.pending.get(data.id);
      if (!request) return;
      this.pending.delete(data.id);
      if (data.error) {
        this.acceptedRevision = undefined;
        this.armIdle();
        request.reject(Error(data.error));
        return;
      }
      let view: WorkbookView = data.view;
      if (data.prepared) {
        view = request.base;
        for (const sheet of view.sheets) rememberFormulaCount(sheet, data.formulaCounts[sheet.name] ?? 0);
      } else if (!view) {
        const patches = new Map<string, Record<string, WorkbookCell>>();
        for (const patch of data.patches) {
          if (!patches.has(patch.sheet)) patches.set(patch.sheet, {});
          patches.get(patch.sheet)![patch.cell] = patch.value;
        }
        view = {
          ...request.base,
          revision: data.revision,
          sheets: request.base.sheets.map((sheet) =>
            applyCalculatedPatches(sheet, patches.get(sheet.name) ?? {}, data.revision)
          ),
        };
      } else
        for (const sheet of view.sheets) {
          prepareCellRecord(sheet.cells);
          formulaCount(sheet);
        }
      view = { ...view, sourceData: request.base.sourceData, sourceSnapshot: request.base.sourceSnapshot };
      this.acceptedRevision = view.revision;
      this.armIdle();
      request.resolve({ view, edits: data.edits });
    };
    worker.onerror = (event) => {
      if (this.worker === worker) this.dispose(Error(event.message || 'Workbook worker failed'));
    };
    return worker;
  }
  async prepare(view: WorkbookView): Promise<WorkbookView> {
    this.observeDiagnostics();
    if (view.storage) return view;
    // A literal-only workbook needs no calculation session. Keep its cell map
    // in the renderer instead of cloning every cell into a second worker heap.
    if (view.sheets.every((sheet) => formulaCount(sheet) === 0)) {
      if (this.worker) this.dispose(Error('Workbook calculation base changed'));
      this.acceptedRevision = undefined;
      return view;
    }
    return (await this.request(view, [], true)).view;
  }
  async apply(
    view: WorkbookView,
    commands: WorkbookCommand[]
  ): Promise<{
    view: WorkbookView;
    edits: WorkbookHistoryEdit[];
  }> {
    this.observeDiagnostics();
    if (view.storage) {
      if (!this.owner) throw Error('The indexed workbook owner is unavailable');
      const generation = this.ownerGeneration;
      return new Promise((resolve, reject) => {
        this.ownerPending.add(reject);
        this.owner!(view, commands).then(
          (result) => {
            this.ownerPending.delete(reject);
            if (generation === this.ownerGeneration) resolve(result);
          },
          (error) => {
            this.ownerPending.delete(reject);
            reject(error);
          }
        );
      });
    }
    if (view.sheets.every((sheet) => formulaCount(sheet) === 0)) {
      const bounded = projectCellCommands(view, commands);
      if (bounded && bounded.changes.every((change) => !change.value.formula && !change.value.formulaType)) {
        const revision = `${view.revision.split(':draft:')[0]}:draft:${this.session}:${++this.sequence}`;
        const patches = new Map<string, Record<string, WorkbookCell>>();
        for (const change of bounded.changes) {
          if (!patches.has(change.sheet)) patches.set(change.sheet, {});
          patches.get(change.sheet)![change.cell] = change.value;
        }
        this.acceptedRevision = undefined;
        return {
          view: {
            ...view,
            revision,
            sheets: view.sheets.map((sheet) => applyCalculatedPatches(sheet, patches.get(sheet.name) ?? {}, revision)),
          },
          edits: bounded.edits,
        };
      }
    }
    return this.request(view, commands, false);
  }
  private request(
    view: WorkbookView,
    commands: WorkbookCommand[],
    prepare: boolean
  ): Promise<{
    view: WorkbookView;
    edits: WorkbookHistoryEdit[];
  }> {
    const id = ++this.sequence;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { base: view, resolve, reject });
      this.ensureWorker().postMessage({
        id,
        ...(view.revision === this.acceptedRevision && !prepare
          ? {}
          : { view: { ...view, sheets: plainSheets(view.sheets), sourceData: undefined } }),
        baseRevision: view.revision,
        commands,
        prepare,
        revision: prepare ? view.revision : `${view.revision.split(':draft:')[0]}:draft:${this.session}:${id}`,
      });
    });
  }
  dispose(error = Error('Workbook editor closed')): void {
    clearTimeout(this.idleTimer);
    this.ownerGeneration++;
    for (const reject of this.ownerPending) reject(error);
    this.ownerPending.clear();
    this.worker?.terminate();
    this.worker = undefined;
    this.acceptedRevision = undefined;
    for (const request of this.pending.values()) {
      request.reject(error);
    }
    this.pending.clear();
    this.releaseDiagnostics?.();
    this.releaseDiagnostics = undefined;
  }
}
