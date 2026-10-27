import { validateHyperlink } from '../core/hyperlinks.mjs';
import type { HyperlinkFollowResult } from '../ui/hyperlinkController';
import { validateDuplicatePreview } from '../core/duplicate-model.mjs';
import { useEffect, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { SpreadsheetGrid, createOwnerReadControls, createWorkbookReviewControls } from '../ui';
import { acceptRecoveredCommand, OwnerCommandRetired, OwnerCommandCancelled } from '../ui/commandAcceptance';
import { configureXlsxHost } from '../ui/host';
import type { WorkbookCommand } from '../core/contracts';
import { decodeOwnerResult } from '../core/transport';
import { validateFormulaPreview, formulaPreviewCancellation } from '../core/formula-preview.mjs';
import type { SpreadsheetStatus, SpreadsheetRequest, SpreadsheetRecoveryEntry } from '../adapters/document-session.mjs';
import type { WorkbookSpellingRequest } from '../adapters/workbook-spelling-jobs.mjs';
import type { LocalActionList } from '../core/local-actions.mjs';
import type { LocalActionRequest } from '../adapters/local-actions-host.mjs';
import { Icon } from './Icon';
import { pastePages, type PasteInspection } from '../ui/largePaste';
import { pickWorkbookPng } from '../ui/pngPicker';
import type { LocalQueryDraft, LocalQueryList } from '../core/local-query.mjs';
import type {
  CalculationAction,
  CalculationGroupReceipt,
  CalculationProgress,
} from '../adapters/workbook-calculation-group.mjs';
import { WorkbookVba } from '../ui/WorkbookVba';
import '../ui/styles/01.css';
import '../ui/styles/02.css';
import './theme.css';

type OpenDocument = SpreadsheetStatus & {
  id: string;
  tabId?: string;
  epoch?: number;
  reuse?: boolean;
  reload?: boolean;
  conflict?: boolean;
  suspended?: boolean;
  operationCancelled?: boolean;
};
interface Operation {
  operationId: string;
  documentId: string | null;
  label: string;
  phase: string;
  cancellable: boolean;
}
interface CancelResult {
  outcome: 'cancelled' | 'completed' | 'finishing';
  recovered?: boolean;
  document: OpenDocument | null;
}
interface RecentEntry {
  id: string;
  name: string;
  displayPath: string;
  lastOpened: number;
  available: boolean;
  reason?: string;
}
interface RecentFiles {
  entries: RecentEntry[];
  warning?: string;
}
interface RecoveryFiles {
  entries: SpreadsheetRecoveryEntry[];
  warnings: string[];
}
interface Documents {
  activeId: string | null;
  documents: OpenDocument[];
  limit: number;
  applicationCalculation?: { mode: 'auto' | 'manual'; recalculateBeforeSave: boolean; revision: number } | null;
}
interface Desktop {
  followHyperlink(
    id: string,
    request: { documentId: string; state: string; sheet: string; cell: string; expectedHyperlinkRevision: string }
  ): Promise<HyperlinkFollowResult>;
  localActions(id: string, request: LocalActionRequest): Promise<LocalActionList>;
  spelling(id: string, request: WorkbookSpellingRequest): Promise<unknown>;
  calculation(
    id: string,
    request: { documentId: string; state: string; requestId: string; action: CalculationAction }
  ): Promise<CalculationGroupReceipt>;
  cancelCalculation(id: string, documentId: string, requestId: string): Promise<CancelResult>;
  onCalculationProgress(callback: (value: CalculationProgress) => void): () => void;
  dataOnlyExport(
    id: string,
    request: { documentId: string; state: string; format: 'xls' | 'xlsb' | 'xlsx' }
  ): Promise<any>;
  queryConnection(
    id: string,
    request: {
      operation: string;
      documentId: string;
      state: string;
      connectionId?: string;
      expectedRevision?: number;
      draft?: LocalQueryDraft;
      requestId?: string;
      offset?: number;
      replacementOffset?: number;
      expectedConversionRevision?: string;
      expectedReplacementRevision?: string;
    }
  ): Promise<any>;
  refreshAllQueries(
    id: string,
    request: {
      operation: 'preview' | 'refresh';
      documentId: string;
      state: string;
      requestId: string;
      previewId?: string;
    }
  ): Promise<any>;
  cancelQuery(id: string, documentId: string, requestId: string): Promise<CancelResult>;
  importFile(
    id: string,
    request: { documentId: string; state: string; sheet: string; start: string; requestId: string }
  ): Promise<{
    imported?: boolean;
    cancelled?: boolean;
    document: OpenDocument;
    source?: { id: string; name: string; bytes: number; sha256: string };
    count?: number;
  }>;
  cancelPaste(id: string, documentId: string, requestId: string): Promise<CancelResult>;
  documents(): Promise<Documents>;
  activateTab(id: string): Promise<OpenDocument>;
  closeTab(id: string): Promise<{ closed: boolean; cancelled: boolean }>;
  closeAll(): Promise<{ closed: boolean; cancelled: boolean }>;
  onDocuments(callback: (value: Documents) => void): () => void;
  current(): Promise<OpenDocument | null>;
  open(): Promise<OpenDocument | null>;
  activateOpen(id: string, discardPrevious?: boolean): Promise<OpenDocument | null>;
  cancelOpen(id: string): Promise<void>;
  request(id: string, request: SpreadsheetRequest): Promise<any>;
  copy(id: string): Promise<boolean>;
  confirmLeave(id: string, allowKeepRecovery?: boolean): Promise<'save' | 'discard' | 'cancel' | 'keep'>;
  publishPrint(
    id: string,
    request: { documentId: string; state: string; jobId: string; kind: 'pdf' | 'printer'; requestId: string }
  ): Promise<{ exported?: boolean; submitted?: boolean; cancelled?: boolean }>;
  cancelPrint(id: string, documentId: string, requestId: string): Promise<CancelResult>;
  copyText(text: string): Promise<void>;
  readClipboardText(): Promise<string>;
  closeResult(accepted: boolean): Promise<void>;
  onClose(callback: () => void): () => void;
  onAction(callback: (action: string) => void): () => void;
  recentFiles(): Promise<RecentFiles>;
  openRecent(id: string): Promise<OpenDocument>;
  removeRecent(id: string): Promise<RecentFiles>;
  newWorkbook(): Promise<OpenDocument>;
  save(id: string): Promise<OpenDocument | null>;
  recoveryFiles(): Promise<RecoveryFiles>;
  openRecovery(id: string): Promise<OpenDocument>;
  operation(): Promise<Operation | null>;
  onProgress(callback: (operation: Operation | null) => void): () => void;
  cancelOperation(id: string, documentId: string | null): Promise<CancelResult>;
  reload(id: string): Promise<OpenDocument>;
  reopenRecovery(id: string): Promise<OpenDocument>;
  copyRecovery(id: string): Promise<boolean>;
  discardRecovery(id: string): Promise<boolean>;
}
declare global {
  interface Window {
    xlsxDesktop: Desktop;
  }
}
const desktop = window.xlsxDesktop;
configureXlsxHost({ api: desktop, Icon, useI18n: () => ({ locale: 'en' }) });
function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
function App(): JSX.Element {
  const [documents, setDocuments] = useState<OpenDocument[]>([]),
    [activeId, setActiveId] = useState<string | null>(null),
    [busy, setBusy] = useState(false),
    [error, setError] = useState(''),
    [notice, setNotice] = useState('');
  const [recents, setRecents] = useState<RecentFiles>({ entries: [] }),
    [recentOpen, setRecentOpen] = useState(false);
  const [recoveries, setRecoveries] = useState<RecoveryFiles>({ entries: [], warnings: [] });
  const [operation, setOperation] = useState<Operation | null>(null),
    [cancelling, setCancelling] = useState(false);
  const [applicationCalculation, setApplicationCalculation] = useState<Documents['applicationCalculation']>(null);
  const info = documents.find((value) => value.id === activeId) ?? null;
  const latest = useRef<OpenDocument | null>(null),
    all = useRef<OpenDocument[]>([]),
    activeDocument = useRef<string | null>(null),
    queue = useRef<Promise<unknown>>(Promise.resolve()),
    pending = useRef(0),
    ownerQueues = useRef(new Map<string, { tail: Promise<unknown>; count: number }>()),
    [ownerBusy, setOwnerBusy] = useState<Record<string, number>>({});
  function accept(value: Documents): void {
    const live = new Set(value.documents.map((record) => record.tabId ?? record.id));
    for (const [key, owner] of ownerQueues.current)
      if (!live.has(key) && owner.count === 0) ownerQueues.current.delete(key);
    setOwnerBusy((previous) => Object.fromEntries(Object.entries(previous).filter(([key]) => live.has(key))));
    all.current = value.documents;
    activeDocument.current = value.activeId;
    latest.current = value.documents.find((record) => record.id === value.activeId) ?? null;
    setDocuments(value.documents);
    setActiveId(value.activeId);
    if (Object.hasOwn(value, 'applicationCalculation')) setApplicationCalculation(value.applicationCalculation);
  }
  function update(value: OpenDocument): void {
    const index = all.current.findIndex((record) => (record.tabId ?? record.id) === (value.tabId ?? value.id));
    if (index < 0) return;
    const previous = all.current[index];
    if ((value.epoch ?? 1) < (previous.epoch ?? 1)) return;
    const next = [...all.current];
    next[index] = value;
    accept({
      activeId: activeDocument.current === previous.id ? value.id : activeDocument.current,
      documents: next,
      limit: 4,
    });
  }
  useEffect(() => {
    const off = desktop.onDocuments(accept);
    void desktop
      .current()
      .then(() => desktop.documents())
      .then(accept)
      .catch((error) => setError(errorMessage(error)));
    return off;
  }, []);
  function refreshRecents(): Promise<void> {
    return desktop
      .recentFiles()
      .then(setRecents)
      .catch((error) => setRecents((previous) => ({ ...previous, warning: errorMessage(error) })));
  }
  function refreshRecovery(): Promise<void> {
    return desktop
      .recoveryFiles()
      .then(setRecoveries)
      .catch((error) => setRecoveries((previous) => ({ ...previous, warnings: [errorMessage(error)] })));
  }
  async function refreshFiles(): Promise<void> {
    await Promise.all([refreshRecents(), refreshRecovery()]);
  }
  useEffect(() => {
    void refreshFiles();
  }, []);
  function enqueue<T>(work: () => Promise<T>, id?: string): Promise<T> {
    const record = id ? all.current.find((value) => value.id === id) : null,
      key = record?.tabId ?? record?.id;
    if (id && !record) return Promise.reject(Error('Document changed before the action could run.'));
    const showResult = (): boolean =>
      !id ||
      (activeDocument.current === id &&
        all.current.some(
          (value) => value.id === id && value.documentId === record?.documentId && value.epoch === record?.epoch
        ));
    const owner = key ? (ownerQueues.current.get(key) ?? { tail: Promise.resolve(), count: 0 }) : null;
    if (owner && key) {
      ownerQueues.current.set(key, owner);
      owner.count++;
      setOwnerBusy((value) => ({ ...value, [key]: owner.count }));
    } else {
      pending.current++;
      setBusy(true);
    }
    const task = (owner?.tail ?? queue.current).then(async () => {
      if (showResult()) {
        setError('');
        setNotice('');
      }
      if (id && !all.current.some((value) => value.id === id))
        throw Error('Document changed before the action could run.');
      return work();
    });
    const settled = task
      .then(
        () => undefined,
        async (error) => {
          if (showResult()) setError(errorMessage(error));
          await desktop
            .documents()
            .then(accept)
            .catch(() => {});
        }
      )
      .finally(() => {
        if (owner && key) {
          owner.count--;
          if (owner.count === 0 && !all.current.some((record) => (record.tabId ?? record.id) === key)) {
            ownerQueues.current.delete(key);
            setOwnerBusy((value) => Object.fromEntries(Object.entries(value).filter(([entry]) => entry !== key)));
          } else setOwnerBusy((value) => ({ ...value, [key]: owner.count }));
        } else if (--pending.current === 0) setBusy(false);
      });
    if (owner) owner.tail = settled;
    else queue.current = settled;
    return task;
  }
  async function switchTab(id: string): Promise<void> {
    const departing = latest.current;
    (document.activeElement as HTMLElement | null)?.blur();
    await ownerQueues.current.get(departing?.tabId ?? departing?.id ?? '')?.tail;
    if (all.current.some((record) => record.id === id)) {
      await desktop.activateTab(id);
      await desktop.documents().then(accept);
    }
  }
  async function leave(current: OpenDocument, deferDiscard = false): Promise<boolean | 'discard'> {
    if (!current?.dirty || current.suspended) return true;
    const choice = await desktop.confirmLeave(current.id, !deferDiscard);
    if (choice === 'cancel') return false;
    if (choice === 'keep') return true;
    if (choice === 'discard' && deferDiscard) return 'discard';
    const next =
      choice === 'save' ? await desktop.save(current.id) : await desktop.request(current.id, { operation: 'discard' });
    if (!next || next.operationCancelled) return false;
    update(next);
    return true;
  }
  function action(kind: string, recentId?: string): void {
    // The native Edit menu targets workbook history. A live IME candidate must
    // finish or cancel before a menu action can commit/undo a cell draft.
    if (
      (kind === 'undo' || kind === 'redo') &&
      ((document.activeElement as HTMLElement | null)?.dataset.composing === 'true' || pending.current > 0)
    )
      return; // Blur commits the active grid editor before the shell operation is queued.
    (document.activeElement as HTMLElement | null)?.blur();
    const captured = latest.current;
    enqueue(
      async () => {
        if (['open', 'recent', 'new', 'recover', 'reload', 'close-tab'].includes(kind))
          await ownerQueues.current.get(captured?.tabId ?? captured?.id ?? '')?.tail;
        if (kind === 'close') await Promise.all([...ownerQueues.current.values()].map((owner) => owner.tail));
        const current = captured ? (all.current.find((record) => record.id === captured.id) ?? null) : null;
        if (['open', 'recent', 'new', 'recover', 'reload'].includes(kind)) {
          try {
            const next =
              kind === 'new'
                ? await desktop.newWorkbook()
                : kind === 'recover'
                  ? await desktop.openRecovery(recentId!)
                  : kind === 'reload' && current
                    ? await desktop.reload(current.id)
                    : kind === 'recent'
                      ? await desktop.openRecent(recentId!)
                      : await desktop.open();
            if (!next) return;
            try {
              const decision = kind === 'reload' && current ? await leave(current, true) : true;
              if (decision) {
                const activated = await desktop.activateOpen(next.id, decision === 'discard');
                if (activated) {
                  await desktop.documents().then(accept);
                  setRecentOpen(false);
                }
              } else await desktop.cancelOpen(next.id);
            } catch (error) {
              await desktop.cancelOpen(next.id).catch((e) => console.error('xlsx.cancel-open', e));
              throw error;
            }
          } finally {
            await refreshFiles();
          }
        } else if (kind === 'save' && current) {
          const next = await desktop.save(current.id);
          if (next) {
            update(next);
            if (!next.operationCancelled) setNotice('Saved ' + next.name);
            await refreshFiles();
          }
        } else if (kind === 'copy' && current) {
          if (await desktop.copy(current.id)) setNotice('Copy saved. The current file remains open.');
        } else if ((kind === 'undo' || kind === 'redo') && current)
          update(await desktop.request(current.id, { operation: kind }));
        else if (kind === 'resume' && current) update(await desktop.reopenRecovery(current.id));
        else if (kind === 'recovery-copy') {
          if (await desktop.copyRecovery(recentId!)) setNotice('Recovery copy saved.');
          await refreshRecovery();
        } else if (kind === 'recovery-discard') {
          await desktop.discardRecovery(recentId!);
          await refreshRecovery();
        } else if (kind === 'close-tab' && current) {
          await desktop.closeTab(current.id);
          await desktop.documents().then(accept);
          await refreshFiles();
        } else if (kind === 'close') {
          await desktop.closeAll();
          await desktop
            .documents()
            .then(accept)
            .catch(() => {});
        }
      },
      captured && ['save', 'copy', 'undo', 'redo', 'resume'].includes(kind) ? captured.id : undefined
    );
  }
  useEffect(() => {
    const offClose = desktop.onClose(() => action('close')),
      offAction = desktop.onAction(action);
    return () => {
      offClose();
      offAction();
    };
  }, []);
  useEffect(() => {
    const off = desktop.onProgress(setOperation);
    void desktop.operation().then(setOperation);
    return off;
  }, []);
  async function cancel(): Promise<void> {
    if (!operation || cancelling) return;
    setCancelling(true);
    try {
      const result = await desktop.cancelOperation(operation.operationId, operation.documentId);
      if (result.document) update(result.document);
      setNotice(
        result.outcome === 'cancelled'
          ? result.recovered
            ? 'Operation cancelled. Accepted edits and undo history were recovered.'
            : result.document
              ? 'Operation cancelled. The current workbook and its accepted edits were kept.'
              : 'Operation cancelled.'
          : result.outcome === 'completed'
            ? 'The operation completed before cancellation.'
            : 'The operation is finishing a durable write.'
      );
      await refreshFiles();
    } catch (error) {
      setError(errorMessage(error));
      const current = await desktop.current().catch(() => null);
      if (current) update(current);
    } finally {
      setCancelling(false);
    }
  }
  function edit(id: string, commands: WorkbookCommand[]): Promise<void> {
    return enqueue(async () => {
      const current = all.current.find((record) => record.id === id);
      if (!current) throw Error('Document changed before this edit could be applied.');
      update(await desktop.request(id, { operation: 'apply', state: current.state, commands }));
    }, id);
  }
  /** Registered actions keep the exact initiating accepted head. Legacy cell
   * edits retain their existing FIFO/blur behavior through edit(). */
  function editCommand(bound: OpenDocument, commands: WorkbookCommand[]): Promise<void> {
    const captured = { id: bound.id, documentId: bound.documentId, state: bound.state, epoch: bound.epoch };
    const acceptedCommands = structuredClone(commands);
    const check = (): void => {
      const current = all.current.find((record) => record.id === captured.id);
      if (
        !current ||
        current.documentId !== captured.documentId ||
        current.state !== captured.state ||
        current.epoch !== captured.epoch ||
        activeDocument.current !== captured.id
      )
        throw new OwnerCommandRetired();
    };
    try {
      check();
    } catch (error) {
      return Promise.reject(error);
    }
    return enqueue(async () => {
      check();
      acceptRecoveredCommand(
        await desktop.request(captured.id, {
          operation: 'apply',
          documentId: captured.documentId,
          state: captured.state,
          commands: acceptedCommands,
        }),
        update
      );
    }, captured.id);
  }
  function historyCommand(bound: OpenDocument, operation: 'undo' | 'redo'): Promise<void> {
    const captured = { id: bound.id, documentId: bound.documentId, state: bound.state, epoch: bound.epoch };
    const check = (): void => {
      const current = all.current.find((record) => record.id === captured.id);
      if (
        !current ||
        current.documentId !== captured.documentId ||
        current.state !== captured.state ||
        current.epoch !== captured.epoch ||
        activeDocument.current !== captured.id
      )
        throw new OwnerCommandRetired();
      if (current.view.readOnly || current.suspended) throw Error('Workbook is read-only or unavailable.');
      if (operation === 'undo' ? !current.canUndo : !current.canRedo)
        throw Error('No accepted ' + operation + ' is available.');
    };
    try {
      check();
    } catch (error) {
      return Promise.reject(error);
    }
    return enqueue(async () => {
      check();
      acceptRecoveredCommand(
        await desktop.request(captured.id, { operation, documentId: captured.documentId, state: captured.state }),
        update
      );
    }, captured.id);
  }
  function importFile(bound: OpenDocument, sheet: string, start: string): Promise<void> {
    const initiating = all.current.find((record) => record.id === bound.id);
    if (
      !initiating ||
      initiating.epoch !== bound.epoch ||
      initiating.state !== bound.state ||
      initiating.documentId !== bound.documentId ||
      activeDocument.current !== bound.id
    )
      return Promise.reject(Error('The initiating workbook changed before importing.'));
    // Clicking the import control blurs an editor first. Flush those preceding
    // owner actions, then capture the accepted revision before opening the chooser.
    return enqueue(async () => {
      const current = all.current.find((record) => record.id === bound.id);
      if (
        !current ||
        current.epoch !== bound.epoch ||
        current.documentId !== bound.documentId ||
        activeDocument.current !== bound.id
      )
        throw Error('The initiating workbook changed before importing.');
      const result = await desktop.importFile(bound.id, {
        documentId: current.documentId,
        state: current.state,
        sheet,
        start,
        requestId: crypto.randomUUID(),
      });
      if (result.document) update(result.document);
      if (result.imported)
        setNotice(
          'Imported ' + result.count + ' literal cells from ' + result.source?.name + ' at ' + sheet + '!' + start
        );
      else if (result.cancelled) setNotice('File import cancelled. Accepted workbook edits were kept.');
    }, bound.id);
  }
  function runCalculation(
    bound: OpenDocument,
    action: CalculationAction,
    requestId: string,
    onProgress: (value: CalculationProgress) => void
  ): Promise<CalculationGroupReceipt> {
    const captured = { id: bound.id, epoch: bound.epoch, documentId: bound.documentId, state: bound.state };
    const check = () => {
      const live = all.current.find((r) => r.id === captured.id);
      if (
        !live ||
        live.epoch !== captured.epoch ||
        live.documentId !== captured.documentId ||
        live.state !== captured.state ||
        activeDocument.current !== captured.id
      )
        throw new OwnerCommandRetired();
    };
    try {
      check();
    } catch (error) {
      return Promise.reject(error);
    }
    return enqueue(async () => {
      check();
      const off = desktop.onCalculationProgress((value) => {
        if (value.requestId === requestId) onProgress(value);
      });
      try {
        return await desktop.calculation(captured.id, {
          documentId: captured.documentId,
          state: captured.state,
          requestId,
          action,
        });
      } finally {
        try {
          off();
        } catch (error) {
          console.warn('xlsx.calculation-progress-cleanup', error);
        }
        await desktop
          .documents()
          .then(accept)
          .catch((error) => console.warn('xlsx.calculation-refresh', error));
      }
    });
  }
  function queryCall(
    bound: OpenDocument,
    operation: string,
    fields: {
      connectionId?: string;
      expectedRevision?: number;
      draft?: LocalQueryDraft;
      requestId?: string;
      offset?: number;
      replacementOffset?: number;
      expectedConversionRevision?: string;
      expectedReplacementRevision?: string;
    } = {}
  ): Promise<any> {
    bound = { ...bound };
    fields = structuredClone(fields);
    const initiating = all.current.find((value) => value.id === bound.id);
    if (
      !initiating ||
      initiating.epoch !== bound.epoch ||
      initiating.documentId !== bound.documentId ||
      initiating.state !== bound.state ||
      activeDocument.current !== bound.id
    )
      return Promise.reject(Error('Activate the initiating workbook at its current revision before using a query.'));
    const run = async () => {
      const current = all.current.find((value) => value.id === bound.id);
      if (
        !current ||
        current.epoch !== bound.epoch ||
        current.documentId !== bound.documentId ||
        current.state !== bound.state ||
        activeDocument.current !== bound.id
      )
        throw Error('The query workbook changed before its action.');
      const result = await desktop.queryConnection(bound.id, {
        operation,
        documentId: bound.documentId,
        state: bound.state,
        ...fields,
      });
      if (result.document) update(result.document);
      return result;
    };
    // Wait for preceding owner actions without rebinding this query to their head.
    return operation === 'list' ? run() : enqueue(run, bound.id);
  }
  function refreshAllCall(
    bound: OpenDocument,
    operation: 'preview' | 'refresh',
    requestId: string,
    previewId?: string
  ): Promise<any> {
    bound = { ...bound };
    const original = all.current.find((value) => value.id === bound.id);
    if (
      !original ||
      original.epoch !== bound.epoch ||
      original.documentId !== bound.documentId ||
      original.state !== bound.state ||
      activeDocument.current !== bound.id
    )
      return Promise.reject(Error('Activate the original workbook at its captured revision before Refresh All.'));
    return enqueue(async () => {
      const current = all.current.find((value) => value.id === bound.id);
      if (
        !current ||
        current.epoch !== bound.epoch ||
        current.documentId !== bound.documentId ||
        current.state !== bound.state ||
        activeDocument.current !== bound.id
      )
        throw Error('The Refresh All workbook changed before its request.');
      const result = await desktop.refreshAllQueries(bound.id, {
        operation,
        documentId: bound.documentId,
        state: bound.state,
        requestId,
        ...(previewId ? { previewId } : {}),
      });
      if (result.document) update(result.document);
      return result;
    }, bound.id);
  }
  function importPaste(
    bound: OpenDocument,
    text: string,
    inspection: PasteInspection,
    options: { signal: AbortSignal; onProgress: (completed: number, total: number) => void }
  ): Promise<void> {
    const initiating = (): OpenDocument | undefined =>
      all.current.find(
        (record) => record.id === bound.id && record.documentId === bound.documentId && record.epoch === bound.epoch
      );
    const currentHead = (): boolean => activeDocument.current === bound.id && initiating()?.state === bound.state;
    if (!currentHead()) return Promise.reject(new OwnerCommandRetired());
    return enqueue(async () => {
      const request = { documentId: bound.documentId, state: bound.state },
        check = (): void => {
          if (options.signal.aborted) throw Error('Paste cancelled. Accepted workbook edits were kept.');
          if (!currentHead()) throw new OwnerCommandRetired();
        };
      const requestId = crypto.randomUUID();
      let uploadId: string | undefined,
        sealed = false,
        accepted = false,
        applying = false,
        finished = false,
        cancellation: Promise<CancelResult> | undefined;
      const abort = (): void => {
        if (!applying || finished || cancellation) return;
        cancellation = desktop.cancelPaste(bound.id, bound.documentId, requestId).then((result) => {
          if (result.document?.documentId === bound.documentId) update(result.document);
          return result;
        });
        void cancellation.catch((error) => {
          if (currentHead()) setError(errorMessage(error));
        });
      };
      options.signal.addEventListener('abort', abort);
      try {
        check();
        const upload = await desktop.request(bound.id, {
          operation: 'importBegin',
          ...request,
          shape: inspection.shape,
          count: inspection.count,
          policy: inspection.policy,
        });
        uploadId = upload.id;
        let completed = 0;
        for (const page of pastePages(text, inspection)) {
          check();
          const result = await desktop.request(bound.id, {
            operation: 'importUpload',
            ...request,
            uploadId,
            sequence: page.sequence,
            cells: page.cells,
          });
          completed += page.cells.length;
          if (result.sequence !== page.sequence + 1 || result.count !== completed)
            throw Error('Paste upload acknowledgement changed.');
          options.onProgress(completed, inspection.count);
        }
        check();
        const descriptor = await desktop.request(bound.id, { operation: 'importSeal', ...request, uploadId });
        sealed = true;
        check();
        applying = true;
        const next = await desktop.request(bound.id, {
          operation: 'apply',
          ...request,
          requestId,
          commands: [descriptor],
        });
        acceptRecoveredCommand(next, update);
        if (cancellation && (await cancellation).outcome === 'cancelled') throw new OwnerCommandCancelled();
        accepted = true;
      } finally {
        finished = true;
        applying = false;
        options.signal.removeEventListener('abort', abort);
        if (cancellation) await cancellation.catch(() => {});
        if (uploadId && !sealed)
          await desktop
            .request(bound.id, { operation: 'importCancel', ...request, uploadId })
            .catch((error) => console.error('xlsx.cancel-paste-upload', error));
        if (sealed && !accepted) {
          const restored = (await desktop.documents().catch(() => null))?.documents.find(
            (record) => record.tabId === bound.tabId
          );
          if (restored?.documentId === bound.documentId)
            await desktop
              .request(restored.id, {
                operation: 'importPrune',
                documentId: restored.documentId,
                state: restored.state,
              })
              .catch((error) => console.error('xlsx.prune-unaccepted-paste', error));
        }
      }
    }, bound.id);
  }
  function removeRecent(id: string): void {
    enqueue(async () => {
      setRecents(await desktop.removeRecent(id));
    });
  }
  // The grid's read-only owner controls, the same as in every host (ui/ownerReadControls.ts).
  function ownerReads(bound: OpenDocument) {
    if (bound.kind !== 'xlsx' || bound.view.storage?.kind !== 'rust-indexed-v1') return undefined;
    const epoch = bound.epoch ?? 1,
      sourceView = bound.view,
      privateStorageState = sourceView.storage?.state;
    const receipt = {
      documentId: bound.documentId,
      ownerId: bound.view.storage.sessionId,
      state: bound.state,
      revision: bound.view.revision,
    };
    return createOwnerReadControls({
      receipt,
      current: (operation) => {
        const live = all.current.find((record) => record.id === bound.id && record.tabId === bound.tabId);
        if (
          activeDocument.current !== bound.id ||
          !live ||
          live.suspended ||
          (live.epoch ?? 1) !== epoch ||
          live.documentId !== receipt.documentId ||
          live.kind !== 'xlsx' ||
          live.view.storage?.kind !== 'rust-indexed-v1' ||
          live.view.storage.sessionId !== receipt.ownerId ||
          live.state !== receipt.state ||
          live.view.revision !== receipt.revision ||
          (operation === 'noteMembership' &&
            (live.view.passiveNotice || live.view !== sourceView || live.view.storage.state !== privateStorageState))
        )
          throw new OwnerCommandRetired();
      },
      request: (operation, fields) => desktop.request(bound.id, { operation, ...fields }),
    });
  }
  return (
    <main className="xlsx-desktop">
      <header className="desktop-header">
        <strong>XLSX Editor</strong>
        <button disabled={busy} onClick={() => action('new')}>
          New workbook
        </button>
        <button disabled={busy} onClick={() => action('open')}>
          Open…
        </button>
        <button
          disabled={busy}
          aria-expanded={recentOpen || !info}
          aria-controls="desktop-recents"
          onClick={() => {
            setRecentOpen((value) => !value);
            void refreshRecents();
          }}
        >
          Recent files
        </button>
        <button disabled={busy || !info || (!info.dirty && !info.unsaved)} onClick={() => action('save')}>
          Save
        </button>
        <button disabled={busy || !info} onClick={() => action('copy')}>
          Save a copy…
        </button>
        <span className="desktop-title">
          {info?.name ?? 'No spreadsheet open'}
          {info?.dirty ? ' •' : ''}
        </span>
      </header>
      {!!documents.length && (
        <nav className="desktop-tabs" aria-label="Workbook tabs">
          <div role="tablist" aria-label="Open workbooks">
            {documents.map((record) => (
              <div className="desktop-tab" key={record.tabId ?? record.id}>
                <button
                  role="tab"
                  aria-selected={record.id === activeId}
                  aria-controls={'workbook-' + (record.tabId ?? record.id)}
                  disabled={busy}
                  onClick={() => void switchTab(record.id).catch((error) => setError(errorMessage(error)))}
                >
                  {record.name}
                  {record.dirty ? ' •' : ''}
                </button>
                <button
                  disabled={busy}
                  aria-label={'Close workbook ' + record.name}
                  onClick={() => {
                    const departing = latest.current;
                    (document.activeElement as HTMLElement | null)?.blur();
                    enqueue(async () => {
                      await ownerQueues.current.get(departing?.tabId ?? departing?.id ?? '')?.tail;
                      await ownerQueues.current.get(record.tabId ?? record.id)?.tail;
                      await desktop.closeTab(record.id);
                      await desktop.documents().then(accept);
                      await refreshFiles();
                    });
                  }}
                >
                  ×
                </button>
              </div>
            ))}
          </div>
          <span>{documents.length} / 4</span>
        </nav>
      )}
      {error && (
        <div className="desktop-error" role="alert">
          {error}
        </div>
      )}
      {notice && (
        <div className="desktop-notice" role="status">
          {notice}
        </div>
      )}
      {operation && (
        <div className="desktop-progress" role="status" aria-live="polite">
          <span className="desktop-spinner" aria-hidden="true" />
          <strong>{operation.label}</strong>
          <span>
            {/^(?:[^ ]*-file|checking-query-output|uploading-query) /.test(operation.phase)
              ? (
                  {
                    'checking-file': 'Checking imported file',
                    'copying-file': 'Capturing imported file',
                    'inspecting-file': 'Inspecting CSV/TSV',
                    'uploading-file': 'Staging imported cells',
                    'reading-file': 'Reading imported cells',
                    'querying-file': 'Selecting and filtering source rows',
                    'uploading-query': 'Staging query result',
                    'checking-query-output': 'Checking the previous query result',
                  } as Record<string, string>
                )[operation.phase.split(' ')[0]] +
                ' ' +
                operation.phase.split(' ')[1]
              : operation.phase.startsWith('applying-paste ')
                ? 'Applying paste ' + operation.phase.slice('applying-paste '.length)
                : ((
                    {
                      starting: 'Starting…',
                      'checking-source': 'Checking the source file…',
                      'creating-workbook': 'Creating a blank workbook…',
                      'copying-source': 'Preparing a private recovery copy…',
                      'checking-recovery': 'Checking accepted recovery edits…',
                      indexing: 'Indexing the workbook…',
                      calculating: 'Calculating formulas…',
                      applying: 'Applying edits…',
                      exporting: 'Preparing workbook output…',
                      'saving-recovery': 'Saving accepted recovery edits…',
                      'writing-output': 'Publishing the spreadsheet file…',
                      'cleaning-up': 'Cleaning temporary indexes…',
                      stopping: 'Stopping the operation…',
                      finishing: 'Finishing a durable write…',
                      completed: 'Completed',
                      ready: 'Ready',
                      'evaluating-pure-function': 'Evaluating the selected pure numeric function…',
                      'exporting-stored-values': 'Preparing stored cell values…',
                      'choosing-printer': 'Preparing the native print dialog…',
                      printing: 'Native printing is finishing…',
                    } as Record<string, string>
                  )[operation.phase] ?? 'Working…')}
          </span>
          <button disabled={!operation.cancellable || cancelling} onClick={() => void cancel()}>
            {cancelling ? 'Cancelling…' : 'Cancel operation'}
          </button>
        </div>
      )}
      {recents.warning && (
        <div className="desktop-recents-warning" role="status">
          {recents.warning}
        </div>
      )}
      {recoveries.warnings.map((warning, index) => (
        <div className="desktop-recents-warning" role="status" key={index}>
          {warning}
        </div>
      ))}
      {!!recoveries.entries.length && (
        <section className="desktop-recents desktop-drafts" aria-label="Recoverable workbooks">
          <strong>Recoverable workbooks</strong>
          <p>Continue, save a copy, or explicitly discard accepted recovery edits.</p>
          <ul>
            {recoveries.entries.map((entry) => (
              <li key={entry.id}>
                <button
                  className="desktop-recent-open"
                  disabled={busy}
                  aria-label={'Recover ' + entry.name}
                  onClick={() => action('recover', entry.id)}
                >
                  <strong>{entry.name}</strong>
                  <span>{entry.displayPath ?? 'Not saved to a file yet'}</span>
                </button>
                <button
                  disabled={busy}
                  aria-label={'Save recovery copy ' + entry.name}
                  onClick={() => action('recovery-copy', entry.id)}
                >
                  Save copy…
                </button>
                <button
                  disabled={busy}
                  aria-label={'Discard recovery ' + entry.name}
                  onClick={() => action('recovery-discard', entry.id)}
                >
                  Discard…
                </button>
              </li>
            ))}
          </ul>
        </section>
      )}
      {(recentOpen || !info) && (
        <section id="desktop-recents" className="desktop-recents" aria-label="Recent files">
          <strong>Recent files</strong>
          {!recents.entries.length ? (
            <p>
              {recents.warning ? 'Recent history is unavailable.' : 'No recent files yet.'} Use Open… to choose a
              spreadsheet.
            </p>
          ) : (
            <ul>
              {recents.entries.map((entry) => (
                <li key={entry.id}>
                  <button
                    className="desktop-recent-open"
                    title={entry.displayPath}
                    aria-label={'Open recent ' + entry.name}
                    disabled={busy || !entry.available}
                    onClick={() => action('recent', entry.id)}
                  >
                    <strong>{entry.name}</strong>
                    <span>{entry.displayPath}</span>
                    {entry.reason && <small>{entry.reason}</small>}
                  </button>
                  <button
                    disabled={busy}
                    aria-label={'Remove recent ' + entry.name}
                    title="Remove from this list; the file and recovery data are kept."
                    onClick={() => removeRecent(entry.id)}
                  >
                    Remove
                  </button>
                </li>
              ))}
            </ul>
          )}
        </section>
      )}
      {info?.recovered && <div className="desktop-recovery">Recovered accepted edits from the previous session.</div>}
      {info?.view.calculationPolicy?.preserveSource && !info.view.readOnly && (
        <div className="desktop-recents-warning" role="status" aria-label="Preserved formula results">
          Unsupported formulas are preserved. Ordinary cells can be edited; stored results need recalculation in
          Microsoft Excel.
        </div>
      )}
      {info?.warning && (
        <div className="desktop-error" role="status">
          {info.warning}
        </div>
      )}
      {info && (['xls', 'xlsb'].includes(info.kind) || info.view.macroProject) && (
        <WorkbookVba
          key={info.id + ':' + info.state}
          scope={info.id + ':' + info.state}
          busy={busy}
          request={(request) =>
            desktop.request(info.id, { ...request, documentId: info.documentId, state: info.state })
          }
          exportValues={
            ['xls', 'xlsb'].includes(info.kind)
              ? (format) =>
                  enqueue(
                    () => desktop.dataOnlyExport(info.id, { documentId: info.documentId, state: info.state, format }),
                    info.id
                  )
              : undefined
          }
        />
      )}
      {info?.conflict && (
        <div className="desktop-conflict" role="alert">
          <span>The original file changed on disk. Your accepted edits are preserved.</span>
          <button disabled={busy} onClick={() => action('copy')}>
            Save conflict copy…
          </button>
          <button disabled={busy} onClick={() => action('reload')}>
            Reload file…
          </button>
          <button disabled={busy} onClick={() => action('resume')}>
            Reopen accepted recovery
          </button>
        </div>
      )}
      {info?.suspended && (
        <div className="desktop-conflict" role="alert">
          <span>The worker stopped. Accepted recovery data is retained.</span>
          <button disabled={busy} onClick={() => action('resume')}>
            Reopen accepted recovery
          </button>
        </div>
      )}
      {documents.map((info) => {
        const tabKey = info.tabId ?? info.id,
          tabBusy = Boolean(ownerBusy[tabKey]) || operation?.documentId === info.id,
          reads = ownerReads(info),
          visible = info.id === activeId;
        return (
          <section
            key={tabKey}
            id={'workbook-' + tabKey}
            data-workbook-id={info.id}
            data-tab-id={tabKey}
            className="desktop-workbook"
            role="tabpanel"
            aria-label={info.name}
            hidden={!visible}
            ref={(node) => {
              if (node) node.inert = !visible;
            }}
            onKeyDownCapture={(event) => {
              if (!visible) event.stopPropagation();
            }}
            onPasteCapture={(event) => {
              if (!visible) event.stopPropagation();
            }}
          >
            <SpreadsheetGrid
              key={tabKey}
              active={visible}
              ownerIdentity={info.id}
              documentId={info.documentId}
              viewPreferenceId={info.viewPreferenceId}
              sheets={info.view.sheets}
              storageKey={info.id + ':' + info.state}
              namedRanges={info.view.namedRanges}
              tables={info.view.tables}
              pivots={info.view.pivots}
              passiveNotice={info.view.passiveNotice}
              nameCommands={info.view.storage?.kind === 'rust-indexed-v1'}
              pageTheme={info.view.storage?.kind === 'rust-indexed-v1' ? info.view.pageTheme : undefined}
              literal={info.kind === 'csv'}
              readOnly={info.view.readOnly || Boolean(info.suspended) || !visible}
              structuralReason={info.view.structuralReason}
              busy={tabBusy}
              canUndo={info.canUndo}
              canRedo={info.canRedo}
              onUndo={() => {
                if (visible) action('undo');
              }}
              onRedo={() => {
                if (visible) action('redo');
              }}
              onCommandUndo={() => (visible ? historyCommand(info, 'undo') : Promise.reject(new OwnerCommandRetired()))}
              onCommandRedo={() => (visible ? historyCommand(info, 'redo') : Promise.reject(new OwnerCommandRetired()))}
              status={
                tabBusy
                  ? 'Working…'
                  : info.dirty
                    ? 'Unsaved changes · recovery journal saved'
                    : info.unsaved
                      ? 'Not saved to a file yet'
                      : 'Saved'
              }
              readRange={async (sheet, range) =>
                (await desktop.request(info.id, { operation: 'read', state: info.state, sheet, range })).cells
              }
              scan={(sheet, after, limit) =>
                desktop.request(info.id, {
                  operation: 'scan',
                  documentId: info.documentId,
                  state: info.state,
                  sheet,
                  after,
                  ...(limit === undefined ? {} : { limit }),
                })
              }
              actionsControls={{
                list: () =>
                  desktop.localActions(info.id, { operation: 'list', documentId: info.documentId, state: info.state }),
                create: (draft, expectedRevision) =>
                  desktop.localActions(info.id, {
                    operation: 'create',
                    documentId: info.documentId,
                    state: info.state,
                    draft: structuredClone(draft),
                    expectedRevision,
                  }),
                update: (actionId, draft, expectedRevision) =>
                  desktop.localActions(info.id, {
                    operation: 'update',
                    documentId: info.documentId,
                    state: info.state,
                    actionId,
                    draft: structuredClone(draft),
                    expectedRevision,
                  }),
                remove: (actionId, expectedRevision) =>
                  desktop.localActions(info.id, {
                    operation: 'remove',
                    documentId: info.documentId,
                    state: info.state,
                    actionId,
                    expectedRevision,
                  }),
              }}
              reviewControls={
                info.kind === 'xlsx' && info.view.storage?.kind === 'rust-indexed-v1' && !info.view.passiveNotice
                  ? createWorkbookReviewControls(
                      {
                        documentId: info.documentId,
                        ownerId: info.view.storage.sessionId,
                        state: info.state,
                        revision: info.view.revision,
                      },
                      (request) => desktop.spelling(info.id, request)
                    )
                  : undefined
              }
              formulaPreviewControls={
                info.kind === 'xlsx' && info.view.storage?.kind === 'rust-indexed-v1'
                  ? {
                      receipt: {
                        documentId: info.documentId,
                        ownerId: info.view.storage.sessionId,
                        state: info.state,
                        revision: info.view.revision,
                      },
                      preview: async (sheet, cell, formula, requestId) => {
                        const result = decodeOwnerResult(
                          await desktop.request(info.id, {
                            operation: 'previewFormula',
                            documentId: info.documentId,
                            state: info.state,
                            requestId,
                            sheet,
                            cell,
                            formula,
                          })
                        );
                        return validateFormulaPreview(result.formulaPreview, {
                          requestId,
                          sheet,
                          cell,
                          formula,
                          receipt: {
                            documentId: info.documentId,
                            ownerId: info.view.storage!.sessionId,
                            state: info.state,
                            revision: info.view.revision,
                          },
                        });
                      },
                      cancel: async (requestId) => {
                        const result = await desktop.request(info.id, {
                          operation: 'cancelFormulaPreview',
                          documentId: info.documentId,
                          state: info.state,
                          requestId,
                        });
                        formulaPreviewCancellation(result.formulaPreviewCancellation, requestId);
                      },
                    }
                  : undefined
              }
              traceControls={
                info.kind === 'xlsx' && info.view.storage?.kind === 'rust-indexed-v1'
                  ? {
                      receipt: {
                        documentId: info.documentId,
                        ownerId: info.view.storage.sessionId,
                        state: info.state,
                        revision: info.view.revision,
                      },
                      trace: (sheet, cell, direction, after, limit) =>
                        desktop.request(info.id, {
                          operation: 'trace',
                          documentId: info.documentId,
                          state: info.state,
                          sheet,
                          cell,
                          direction,
                          ...(after !== undefined ? { after } : {}),
                          ...(limit !== undefined ? { limit } : {}),
                        }),
                    }
                  : undefined
              }
              usedRangeControls={info.view.passiveNotice ? undefined : reads?.usedRangeControls}
              noteMembershipControls={info.view.passiveNotice ? undefined : reads?.noteMembershipControls}
              calculationControls={
                info.kind === 'xlsx' && info.view.calculationPolicy
                  ? {
                      receipt: {
                        documentId: info.documentId,
                        ownerId: info.id,
                        state: info.state,
                        revision: info.view.revision,
                      },
                      policy: info.view.calculationPolicy,
                      optionsScope: 'application',
                      applicationMode: applicationCalculation
                        ? { mode: applicationCalculation.mode, revision: applicationCalculation.revision }
                        : undefined,
                      run: (action, requestId, onProgress) => runCalculation(info, action, requestId, onProgress),
                      cancel: (requestId) =>
                        desktop.cancelCalculation(info.id, info.documentId, requestId).then((result) => {
                          if (result.document) update(result.document);
                          return { outcome: result.outcome };
                        }),
                    }
                  : undefined
              }
              query={async (sheet, query, mode, after, options) => {
                const result = await desktop.request(info.id, {
                  operation: 'query',
                  documentId: info.documentId,
                  state: info.state,
                  sheet,
                  query,
                  mode,
                  after,
                  ...options,
                });
                return { cells: result.cells, matches: result.matches, nextCursor: result.nextCursor };
              }}
              summarize={(sheet, range) =>
                desktop.request(info.id, { operation: 'selection', state: info.state, sheet, range })
              }
              duplicatePreviewReceipt={
                info.kind === 'xlsx' && info.view.storage?.kind === 'rust-indexed-v1'
                  ? {
                      documentId: info.documentId,
                      ownerId: info.view.storage.sessionId,
                      state: info.state,
                      revision: info.view.revision,
                    }
                  : undefined
              }
              duplicatePreview={
                info.kind === 'xlsx' && info.view.storage?.kind === 'rust-indexed-v1'
                  ? async (sheet, range, hasHeaders, columns, offset) => {
                      const receipt = {
                          documentId: info.documentId,
                          ownerId: info.view.storage!.sessionId,
                          state: info.state,
                          revision: info.view.revision,
                        },
                        config = { sheet, range, hasHeaders, columns };
                      const result = await desktop.request(info.id, {
                        operation: 'duplicatePreview',
                        documentId: receipt.documentId,
                        state: receipt.state,
                        ...config,
                        offset,
                      });
                      return validateDuplicatePreview(result.duplicatePreview, receipt, config, offset);
                    }
                  : undefined
              }
              filterValues={
                info.kind === 'xlsx'
                  ? (sheet, range, column, query, offset) =>
                      desktop.request(info.id, {
                        operation: 'filterValues',
                        state: info.state,
                        sheet,
                        range,
                        column,
                        query,
                        offset,
                        limit: 100,
                      })
                  : undefined
              }
              cellInsertionControls={info.kind === 'xlsx' && info.view.storage?.kind === 'rust-indexed-v1'}
              accessibilityControls={reads?.accessibilityControls}
              notesControls={reads?.notesControls}
              hyperlinkControls={
                info.kind === 'xlsx' && info.view.storage?.kind === 'rust-indexed-v1'
                  ? {
                      receipt: {
                        documentId: info.documentId,
                        ownerId: info.view.storage.sessionId,
                        state: info.state,
                        revision: info.view.revision,
                      },
                      read: async (sheet, cell) => {
                        const receipt = {
                          documentId: info.documentId,
                          ownerId: info.view.storage!.sessionId,
                          state: info.state,
                          revision: info.view.revision,
                        };
                        const result = await desktop.request(info.id, {
                          operation: 'hyperlink',
                          documentId: receipt.documentId,
                          state: receipt.state,
                          sheet,
                          cell,
                        });
                        return validateHyperlink(result.hyperlink, receipt, { sheet, cell });
                      },
                      follow: (sheet, cell, revision) =>
                        desktop.followHyperlink(info.id, {
                          documentId: info.documentId,
                          state: info.state,
                          sheet,
                          cell,
                          expectedHyperlinkRevision: revision,
                        }),
                    }
                  : undefined
              }
              objectControls={
                info.kind === 'xlsx'
                  ? {
                      documentId: info.id,
                      revision: info.state,
                      objects: (sheet, options) =>
                        desktop.request(info.id, {
                          operation: 'objects',
                          documentId: info.documentId,
                          state: info.state,
                          sheet,
                          ...options,
                        }),
                      object: (sheet, objectId) =>
                        desktop.request(info.id, {
                          operation: 'object',
                          documentId: info.documentId,
                          state: info.state,
                          sheet,
                          objectId,
                        }),
                      onBatch: (commands) => {
                        if (!visible) return Promise.reject(Error('Activate this workbook before editing.'));
                        return edit(info.id, commands);
                      },
                      pickPng: pickWorkbookPng,
                    }
                  : undefined
              }
              printControls={
                info.kind === 'xlsx'
                  ? {
                      documentId: info.id,
                      revision: info.state,
                      canPrint: true,
                      definition: (sheet) =>
                        desktop.request(info.id, {
                          operation: 'printDefinition',
                          documentId: info.documentId,
                          state: info.state,
                          sheet,
                        }),
                      onBatch: (commands) => {
                        if (!visible) return Promise.reject(Error('Activate this workbook before editing.'));
                        return edit(info.id, commands);
                      },
                      preview: (sheet, settings, requestId, geometryProfile) =>
                        enqueue(
                          () =>
                            desktop.request(info.id, {
                              operation: 'printPrepare',
                              documentId: info.documentId,
                              state: info.state,
                              sheet,
                              settings,
                              requestId,
                              ...(geometryProfile === undefined ? {} : { geometryProfile }),
                            }),
                          info.id
                        ),
                      page: (jobId, pageNumber) =>
                        desktop.request(info.id, {
                          operation: 'printPage',
                          documentId: info.documentId,
                          state: info.state,
                          jobId,
                          pageNumber,
                        }),
                      publish: (jobId, kind, requestId) =>
                        enqueue(
                          () =>
                            desktop.publishPrint(info.id, {
                              documentId: info.documentId,
                              state: info.state,
                              jobId,
                              kind,
                              requestId,
                            }),
                          info.id
                        ),
                      release: (jobId) =>
                        desktop.request(info.id, {
                          operation: 'printRelease',
                          documentId: info.documentId,
                          state: info.state,
                          jobId,
                        }),
                      cancel: (requestId) =>
                        desktop.cancelPrint(info.id, info.documentId, requestId).then((result) => {
                          if (result.document) update(result.document);
                          return result;
                        }),
                    }
                  : undefined
              }
              conditionalControls={
                info.kind === 'xlsx'
                  ? {
                      documentId: info.id,
                      revision: info.state,
                      inventory: (sheet) =>
                        desktop.request(info.id, {
                          operation: 'conditionalRules',
                          documentId: info.documentId,
                          state: info.state,
                          sheet,
                        }),
                      onBatch: (commands) => {
                        if (!visible) return Promise.reject(Error('Activate this workbook before editing.'));
                        return editCommand(info, commands);
                      },
                    }
                  : undefined
              }
              validationControls={
                info.kind === 'xlsx'
                  ? {
                      documentId: info.id,
                      revision: info.state,
                      inventory: (sheet) =>
                        desktop.request(info.id, {
                          operation: 'validationRules',
                          documentId: info.documentId,
                          state: info.state,
                          sheet,
                        }),
                      onBatch: (commands) => {
                        if (!visible) return Promise.reject(Error('Activate this workbook before editing.'));
                        return edit(info.id, commands);
                      },
                    }
                  : undefined
              }
              queryControls={
                info.kind === 'xlsx' && !info.view.readOnly
                  ? {
                      documentId: info.id,
                      revision: info.state,
                      ownerId: info.view.storage?.sessionId,
                      refreshAll: {
                        name: info.name,
                        scope: {
                          documentId: info.documentId,
                          ownerId: info.view.storage?.sessionId,
                          state: info.state,
                        },
                        preview: (requestId) => refreshAllCall(info, 'preview', requestId),
                        refresh: (previewId, requestId) => refreshAllCall(info, 'refresh', requestId, previewId),
                        cancel: (requestId) =>
                          desktop.cancelQuery(info.id, info.documentId, requestId).then((result) => {
                            if (result.document) update(result.document);
                            return result;
                          }),
                      },
                      list: () => queryCall(info, 'list') as Promise<LocalQueryList>,
                      create: (draft) => queryCall(info, 'create', { draft }),
                      update: (connectionId, expectedRevision, draft) =>
                        queryCall(info, 'update', { connectionId, expectedRevision, draft }),
                      remove: (connectionId, expectedRevision) =>
                        queryCall(info, 'remove', { connectionId, expectedRevision }),
                      preview: (connectionId, expectedRevision, requestId, options) =>
                        queryCall(info, 'preview', {
                          connectionId,
                          expectedRevision,
                          requestId,
                          ...(options?.offset !== undefined ? { offset: options.offset } : {}),
                          ...(options?.replacementOffset !== undefined
                            ? { replacementOffset: options.replacementOffset }
                            : {}),
                          ...(options?.expectedConversionRevision !== undefined
                            ? { expectedConversionRevision: options.expectedConversionRevision }
                            : {}),
                        }),
                      refresh: (connectionId, expectedRevision, requestId, options) =>
                        queryCall(info, 'refresh', {
                          connectionId,
                          expectedRevision,
                          requestId,
                          ...(options?.expectedConversionRevision !== undefined
                            ? { expectedConversionRevision: options.expectedConversionRevision }
                            : {}),
                          ...(options?.expectedReplacementRevision !== undefined
                            ? { expectedReplacementRevision: options.expectedReplacementRevision }
                            : {}),
                        }),
                      cancel: (requestId) =>
                        desktop.cancelQuery(info.id, info.documentId, requestId).then((result) => {
                          if (result.document) update(result.document);
                          return result;
                        }),
                    }
                  : undefined
              }
              importFile={
                info.kind === 'xlsx' && visible ? (sheet, start) => importFile(info, sheet, start) : undefined
              }
              importPaste={
                info.kind === 'xlsx'
                  ? (text, inspection, options) => importPaste(info, text, inspection, options)
                  : undefined
              }
              onEdit={(command) =>
                visible ? edit(info.id, [command]) : Promise.reject(Error('Activate this workbook before editing.'))
              }
              onBatch={(commands) =>
                visible ? edit(info.id, commands) : Promise.reject(Error('Activate this workbook before editing.'))
              }
              onCommandBatch={(commands) =>
                visible ? editCommand(info, commands) : Promise.reject(Error('Activate this workbook before editing.'))
              }
            />
          </section>
        );
      })}
      {!documents.length && (
        <section className="desktop-empty">
          <h1>Create or open a spreadsheet</h1>
          <p>XLSX workbooks and literal CSV/TSV files use the same grid as the embedded editor.</p>
          <button disabled={busy} onClick={() => action('new')}>
            New workbook
          </button>{' '}
          <button disabled={busy} onClick={() => action('open')}>
            Open spreadsheet…
          </button>
        </section>
      )}
    </main>
  );
}
createRoot(document.getElementById('root')!).render(<App />);
