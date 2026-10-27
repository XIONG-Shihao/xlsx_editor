import { accessibilityFields, validateAccessibility } from '../core/accessibility.mjs';
import { notesFields, validateNotes } from '../core/notes.mjs';
import { duplicateConfiguration, validateDuplicatePreview } from '../core/duplicate-model.mjs';
import { standalonePlatform } from './platform.mjs';
import path from 'node:path';
import { realpath, stat, mkdir, mkdtemp, rm } from 'node:fs/promises';
import { randomUUID, createHash } from 'node:crypto';
import { SpreadsheetWorker } from '../adapters/worker-client.mjs';
import {
  revision,
  listSpreadsheetRecovery,
  discardSpreadsheetRecovery,
  writePrivateSnapshot,
  publishCopy,
  validId,
  readJson,
  validateRecoveryManifest,
  RECOVERY_MANIFEST_LIMIT,
} from '../adapters/storage.mjs';
import { RecentFiles } from './recent-files.mjs';
import { KeptRecoveryBindings } from '../adapters/kept-recovery.mjs';
import { WorkbookTabRegistry } from '../adapters/workbook-tab-registry.mjs';
import { runCalculationGroup, validateCalculationAction } from '../adapters/workbook-calculation-group.mjs';
import { prepareDelimitedFileImport } from '../adapters/delimited-file-import.mjs';
import { LocalQueryConnections } from '../adapters/local-query-connections.mjs';
import { LocalActions } from '../adapters/local-actions.mjs';
import { localActionRequest, invokeLocalActionRequest } from '../adapters/local-actions-host.mjs';
import { LocalQueryHost } from '../adapters/local-query-host.mjs';
import { LocalQueryRefreshAll } from '../adapters/local-query-refresh-all.mjs';
import { queryRefreshAllRequest } from '../adapters/query-refresh-all-request.mjs';
import { releaseQueryInspections } from '../adapters/local-query-inspection.mjs';
import { WorkbookSpellingJobs, spellingRequest } from '../adapters/workbook-spelling-jobs.mjs';
import { cellPosition, rangeBounds } from '../core/workbook-commands.mjs';
import { formulaPreviewFields, validateFormulaPreview, formulaPreviewCancellation } from '../core/formula-preview.mjs';
import { usedRangeFields, validateUsedRange } from '../core/used-range.mjs';
import { noteMembershipFields, noteMembershipReceipt, validateNoteMembership } from '../core/note-membership.mjs';
import { decodeOwnerResult } from '../dist/core/transport.js';

// The shell chose the desktop before importing this module (platform.mjs).
const platform = standalonePlatform();
const { app, BrowserWindow, clipboard, dialog, ipcMain, Menu, session, shell: electronShell } = platform;
let routeLaunch;
const launchQueue = [];
function incomingFile(filename) {
  if (typeof filename === 'string' && /\.(xlsx|xlsm|xls|xlsb|csv|tsv)$/i.test(filename)) {
    if (routeLaunch) void routeLaunch(filename);
    else launchQueue.push(filename);
  }
}
app.setName('XLSX Editor');
// The native shell owns profile selection and its lease.
async function boot() {
  if (app.isPackaged)
    process.env.XLSX_EDITOR_NATIVE_BIN = path.join(
      process.resourcesPath,
      'native',
      process.platform === 'win32' ? 'workbook-engine.exe' : 'workbook-engine'
    );
  const recoveryDirectory = path.join(app.getPath('userData'), 'recovery');
  const queryConnections = await LocalQueryConnections.open(path.join(app.getPath('userData'), 'local-queries.json'));
  const localActions = await LocalActions.open(path.join(app.getPath('userData'), 'local-actions.json'));
  const recentFiles = await RecentFiles.open(path.join(app.getPath('userData'), 'recent-files.json'));
  const keptRecovery = await KeptRecoveryBindings.open(path.join(app.getPath('userData'), 'kept-recovery.json'));
  const renderPdf = platform.createWorkbookPdfRenderer({ BrowserWindow, session });
  const printNative = platform.createWorkbookNativePrinter({ BrowserWindow, session });
  let shellQueue = Promise.resolve(),
    untitled = 0;
  const recoveryGrants = new Map();
  function shell(work) {
    const next = shellQueue.then(work);
    shellQueue = next.catch(() => {});
    return next;
  }
  let window,
    current = null,
    pending = null,
    permitClose = false,
    rendererReady = false,
    closing = false,
    closeTransaction = false,
    startupOpen = Promise.resolve();
  const registry = new WorkbookTabRegistry();
  let applicationCalculation = null,
    calculationRevision = 0;
  const spellingJobs = new WeakMap(),
    refreshAllJobs = new WeakMap();
  let active = null;
  const completed = new Map();
  const snapshot = () => ({
    activeId: registry.active?.id ?? null,
    documents: registry.records().map(status),
    limit: registry.limit,
    applicationCalculation,
  });
  function publishDocuments() {
    current = registry.active;
    if (window && !window.isDestroyed()) window.webContents.send('xlsx:documents', snapshot());
  }
  function activateRecord(record) {
    registry.activate(record.id);
    publishDocuments();
    return status(record);
  }
  const page = 'workbook-standalone://editor';
  const cleanupErrors = [];
  function publishOperation() {
    if (window && !window.isDestroyed())
      window.webContents.send(
        'xlsx:progress',
        active
          ? {
              operationId: active.id,
              documentId: active.documentId,
              label: active.label,
              phase: active.phase,
              cancellable: active.cancellable,
            }
          : null
      );
  }
  function observe(operation, owner) {
    operation.owner = owner;
    return {
      onStart: (requestId) => {
        operation.requestId = requestId;
      },
      onProgress: (event) => {
        if (active !== operation || operation.cancelled) return;
        operation.phase = event.phase;
        operation.cancellable = event.cancellable;
        if (event.recoveryId) operation.recoveryId = event.recoveryId;
        publishOperation();
      },
    };
  }
  async function perform(label, record, work, cancelValue) {
    if (active) throw Error('A spreadsheet operation is already running.');
    const originalOwner = record?.owner,
      originalEpoch = record?.epoch,
      originalDocument = record?.status.documentId;
    let finish;
    const operation = {
      id: randomUUID(),
      documentId: record?.id ?? null,
      label: record ? label + ' · ' + record.status.name : label,
      phase: 'starting',
      cancellable: false,
      record,
      finished: new Promise((resolve) => {
        finish = resolve;
      }),
    };
    active = operation;
    if (record) record.operation = operation;
    publishOperation();
    try {
      return await work(operation);
    } catch (error) {
      if (operation.cancelled) {
        await operation.cancelAttempt;
        return cancelValue?.() ?? null;
      }
      if (record && registry.contains(record) && record.owner === originalOwner && record.epoch === originalEpoch) {
        if (record.owner.closed) {
          record.suspended = true;
          record.status = {
            ...record.status,
            warning: 'The workbook worker stopped. Accepted recovery data was retained.',
          };
        } else
          try {
            const accepted = await originalOwner.request({ operation: 'status' });
            if (
              registry.contains(record) &&
              record.owner === originalOwner &&
              record.epoch === originalEpoch &&
              accepted.documentId === originalDocument
            ) {
              record.status = accepted;
              publishDocuments();
            }
          } catch (statusError) {
            console.warn('xlsx.accepted-status-after-failure', statusError);
          }
      }
      throw error;
    } finally {
      // The worker grant and stop/recovery must settle before the serialized
      // shell can start a replacement or another mutation.
      await operation.cancelAttempt?.catch(() => {});
      completed.set(operation.id, { documentId: operation.documentId, tabId: record?.tabId });
      if (completed.size > 8) completed.delete(completed.keys().next().value);
      if (active === operation) {
        active = null;
        publishOperation();
      }
      if (record?.operation === operation) delete record.operation;
      publishDocuments();
      finish();
    }
  }
  async function resumeRecord(record) {
    const options =
      record.status.documentId.length === 36 ? { recoveryId: record.status.documentId } : { resume: true };
    try {
      const next = await prepare(options.recoveryId ? undefined : (record.filename ?? undefined), options);
      registry.replace(record, next);
      publishDocuments();
      return next;
    } catch (error) {
      record.suspended = true;
      record.status = {
        ...record.status,
        warning: 'Accepted recovery data was retained, but could not be reopened. ' + String(error.message ?? error),
      };
      throw error;
    }
  }
  async function cancelOperation(id, documentId) {
    const operation = active;
    if (!operation || operation.id !== id || operation.documentId !== documentId) {
      const receipt = completed.get(id),
        record = registry.records().find((record) => record.tabId === receipt?.tabId);
      if (receipt && receipt.documentId === documentId && (documentId === null || record?.id === documentId))
        return { outcome: 'completed', document: status(record) };
      throw Error('Operation access is not granted.');
    }
    if (operation.cancelAttempt) return operation.cancelAttempt;
    if (operation.abortPrint) {
      if (!operation.cancellable)
        return {
          outcome: 'finishing',
          document: status(
            operation.record && registry.records().find((record) => record.tabId === operation.record.tabId)
          ),
        };
      operation.cancelled = true;
      operation.cancellable = false;
      operation.phase = 'stopping';
      publishOperation();
      operation.abortPrint();
      operation.cancelAttempt = Promise.resolve({
        outcome: 'cancelled',
        document: status(
          operation.record && registry.records().find((record) => record.tabId === operation.record.tabId)
        ),
      });
      await operation.finished;
      return operation.cancelAttempt;
    }
    if (!operation.owner || !operation.requestId) {
      if (operation.calculationAbort && operation.cancellable) {
        operation.calculationAbort();
        operation.cancellable = false;
        operation.phase = 'stopping';
        publishOperation();
        return { outcome: 'cancelled', document: status(registry.active) };
      }
      return {
        outcome: 'finishing',
        document: status(
          operation.record && registry.records().find((record) => record.tabId === operation.record.tabId)
        ),
      };
    }
    operation.cancelAttempt = (async () => {
      const grant = await operation.owner.authorizeCancel(operation.requestId);
      if (!grant.granted) {
        operation.calculationAbort?.();
        operation.cancellable = false;
        operation.phase = grant.reason === 'completed' ? 'completed' : 'finishing';
        publishOperation();
        return {
          outcome: grant.reason,
          document: status(
            operation.record && registry.records().find((record) => record.tabId === operation.record.tabId)
          ),
        };
      }
      operation.cancelled = true;
      operation.calculationAbort?.();
      operation.cancellable = false;
      operation.phase = 'stopping';
      publishOperation();
      await operation.owner.cancel();
      const recovered = Boolean(operation.record && registry.contains(operation.record));
      if (recovered) await resumeRecord(operation.record);
      else if (operation.createdCandidate && operation.recoveryId)
        await discardSpreadsheetRecovery(recoveryDirectory, operation.recoveryId, { preparedOnly: true }).catch(
          (error) => {
            if (error.code !== 'ENOENT') console.warn('xlsx.cancel-candidate', error);
          }
        );
      return {
        outcome: 'cancelled',
        recovered,
        document: status(
          operation.record && registry.records().find((record) => record.tabId === operation.record.tabId)
        ),
      };
    })();
    const result = await operation.cancelAttempt;
    if (result.outcome === 'cancelled') await operation.finished;
    return result;
  }
  async function prepare(filename, options = {}, operation) {
    let owner;
    const callbacks = operation
      ? {
          onStart: (id) => {
            operation.requestId = id;
          },
          onProgress: (event) => {
            if (active !== operation || operation.cancelled) return;
            operation.phase = event.phase;
            operation.cancellable = event.cancellable;
            if (event.recoveryId) operation.recoveryId = event.recoveryId;
            publishOperation();
          },
        }
      : {};
    owner = new SpreadsheetWorker({
      filename,
      ...(applicationCalculation
        ? {
            initialCalculationPolicy: {
              mode: applicationCalculation.mode,
              recalculateBeforeSave: applicationCalculation.recalculateBeforeSave,
            },
          }
        : {}),
      ...options,
      recoveryDirectory,
      ...callbacks,
    });
    if (operation) {
      operation.owner = owner;
      operation.createdCandidate = Boolean(options.newWorkbook || options.fresh);
    }
    try {
      const status = await owner.open(),
        binding = await owner.request({ operation: 'binding' });
      const policy = status.view?.calculationPolicy;
      if (!applicationCalculation && status.kind === 'xlsx' && policy && !policy.protectedSource)
        applicationCalculation = {
          mode: policy.mode,
          recalculateBeforeSave: policy.recalculateBeforeSave,
          revision: ++calculationRevision,
        };
      return { id: randomUUID(), filename: binding.filename, owner, status };
    } catch (error) {
      await owner.close().catch((e) => console.error('xlsx.open-cleanup', e));
      throw error;
    }
  }
  const status = (record) =>
    record
      ? {
          id: record.id,
          ...(record.tabId ? { tabId: record.tabId, epoch: record.epoch } : {}),
          ...record.status,
          ...(record.reuseId ? { reuse: true } : {}),
          ...(record.reload ? { reload: true } : {}),
          ...(record.conflict ? { conflict: true } : {}),
          ...(record.suspended ? { suspended: true } : {}),
          ...(record.operationCancelled ? { operationCancelled: true } : {}),
        }
      : null;
  async function retire(record) {
    if (!record?.owner) return;
    refreshAllJobs.get(record.owner)?.coordinator.dispose();
    refreshAllJobs.delete(record.owner);
    const documentId = record.status?.documentId,
      ownerId = record.status?.view?.storage?.sessionId;
    try {
      await spellingJobs.get(record.owner)?.close();
    } catch (error) {
      cleanupErrors.push(String(error));
      console.error('xlsx.spelling-close', error);
    }
    try {
      await record.owner.close();
    } catch (error) {
      cleanupErrors.push(String(error));
      console.error('xlsx.owner-close', error);
    }
    if (documentId && ownerId)
      await releaseQueryInspections('standalone:' + documentId, path.join(app.getPath('userData'), 'query-staging'), {
        ownerId,
      }).catch((error) => console.warn('xlsx.query-inspection-cleanup', error));
  }
  async function remember(record) {
    if (record.filename)
      try {
        await recentFiles.remember(record.filename);
      } catch (error) {
        console.warn('xlsx.recent-files', String(error));
      }
  }
  async function prepareReplacement(filename, operation) {
    const canonical = await realpath(filename),
      existing = registry.byFilename(canonical);
    if (existing) return { id: randomUUID(), reuseId: existing.id, status: existing.status };
    registry.assertRoom();
    return prepare(canonical, { prepared: true }, operation);
  }
  async function destinationFor(record, filename) {
    const target = await realpath(filename).catch(async (error) => {
      if (error.code !== 'ENOENT') throw error;
      return path.join(await realpath(path.dirname(filename)), path.basename(filename));
    });
    registry.assertDestination(record, target);
    return target;
  }
  async function recoveryFiles() {
    recoveryGrants.clear();
    try {
      const result = await listSpreadsheetRecovery(recoveryDirectory);
      // Legacy journals have no filename. Explicit Keep pins their trusted
      // owner binding independently of removable/evictable recent history.
      const candidates = [...result.entries];
      const pins = await keptRecovery.list();
      if (pins.warning) result.warnings.push(pins.warning);
      const known = new Map(
        pins.entries.map((entry) => [
          entry.filename,
          {
            displayPath: entry.filename,
            name: path.basename(entry.filename),
            lastOpened: entry.createdAt,
            pinned: true,
          },
        ])
      );
      for (const recent of (await recentFiles.list()).entries)
        if (!known.has(recent.displayPath)) known.set(recent.displayPath, recent);
      for (const recent of known.values()) {
        const filename = recent.displayPath,
          id = createHash('sha256').update(filename).digest('hex');
        if (!recent.pinned && result.entries.some((entry) => entry.displayPath === filename)) continue;
        try {
          const manifest = validateRecoveryManifest(
            await readJson(path.join(recoveryDirectory, id, 'current.json'), RECOVERY_MANIFEST_LIMIT)
          );
          if (manifest.schema !== 1) continue;
          const tip = manifest.history?.frames.at(-1)?.head ?? manifest.head;
          if (manifest.head === manifest.savedHead && tip === manifest.savedHead && !manifest.pendingPublication)
            continue;
          const entry = {
            id,
            name: recent.name,
            displayPath: filename,
            unsaved: false,
            dirty: manifest.head !== manifest.savedHead,
            createdAt: recent.lastOpened,
          };
          candidates.push(entry);
          recoveryGrants.set(id, { ...entry, legacyFilename: filename });
        } catch (error) {
          if (error.code !== 'ENOENT')
            result.warnings.push(
              'A saved-file recovery record could not be read. Its data was preserved. ' +
                String(error.message ?? error)
            );
        }
      }
      const entries = candidates.filter(
        (entry) => !registry.byDocument(entry.id) && entry.id !== pending?.status.documentId
      );
      const duplicatePaths = new Map();
      for (const entry of entries)
        if (entry.displayPath) duplicatePaths.set(entry.displayPath, (duplicatePaths.get(entry.displayPath) ?? 0) + 1);
      if ([...duplicatePaths.values()].some((count) => count > 1))
        result.warnings.push(
          'Multiple recovery records refer to the same file. Recover a specific record; their data was preserved.'
        );
      for (const entry of entries) if (!recoveryGrants.has(entry.id)) recoveryGrants.set(entry.id, entry);
      for (const id of recoveryGrants.keys()) if (!entries.some((entry) => entry.id === id)) recoveryGrants.delete(id);
      return { ...result, entries };
    } catch (error) {
      return { entries: [], warnings: [String(error.message ?? error)] };
    }
  }
  function prepareGrantedRecovery(entry, operation) {
    return entry.legacyFilename
      ? prepare(entry.legacyFilename, { resume: true }, operation)
      : prepare(undefined, { recoveryId: entry.id }, operation);
  }
  async function pinLegacyBinding(filename, id, expectedState) {
    if (!filename || !/^[a-f0-9]{64}$/.test(id ?? '')) return;
    let manifest;
    try {
      manifest = validateRecoveryManifest(
        await readJson(path.join(recoveryDirectory, id, 'current.json'), RECOVERY_MANIFEST_LIMIT)
      );
    } catch (error) {
      if (error.code === 'ENOENT' && expectedState === undefined) return;
      throw error;
    }
    if (manifest.schema !== 1 || (expectedState !== undefined && (manifest.head ?? 'base') !== expectedState))
      throw Error('The accepted recovery changed before it could be kept.');
    const tip = manifest.history?.frames.at(-1)?.head ?? manifest.head;
    if (manifest.head !== manifest.savedHead || tip !== manifest.savedHead || manifest.pendingPublication)
      await keptRecovery.pin(filename, id);
  }
  async function pinRecordRecovery(record) {
    if (record && !record.discardedRecovery)
      await pinLegacyBinding(record.filename, record.status.documentId, record.status.state);
  }
  async function saveRequest(record) {
    let request = { operation: 'save' };
    if (record.status.unsaved) {
      const result = await dialog.showSaveDialog(window, {
        defaultPath: record.status.name + '.xlsx',
        filters: [{ name: 'Workbook', extensions: ['xlsx'] }],
      });
      if (result.canceled || !result.filePath) return null;
      const existing = await stat(result.filePath).catch((error) => {
        if (error.code === 'ENOENT') return null;
        throw error;
      });
      request = {
        operation: 'saveAs',
        destination: await destinationFor(record, result.filePath),
        expectedRevision: existing ? await revision(result.filePath) : null,
      };
    }
    return request;
  }
  async function saveSelected(id, preparedRequest) {
    const record = selected(id),
      request = preparedRequest ?? (await saveRequest(record));
    if (!request) return null;
    if (request.destination) await destinationFor(record, request.destination);
    return perform(
      'Saving workbook',
      record,
      async (operation) => {
        try {
          record.status = await record.owner.request(request, observe(operation, record.owner));
          record.conflict = false;
          record.operationCancelled = false;
        } catch (error) {
          if (request.operation === 'save' && error.code === 'REVISION_CONFLICT') record.conflict = true;
          throw error;
        }
        record.filename = (await record.owner.request({ operation: 'binding' })).filename;
        if (
          !record.operationCancelled &&
          !record.status.dirty &&
          !record.status.canRedo &&
          /^[a-f0-9]{64}$/.test(record.status.documentId)
        )
          await keptRecovery
            .remove(record.status.documentId)
            .catch((error) => console.warn('xlsx.kept-recovery-cleanup', String(error)));
        if (request.operation === 'saveAs') await remember(record);
        return status(record);
      },
      () => {
        const restored = registry.records().find((item) => item.tabId === record.tabId);
        if (restored) restored.operationCancelled = true;
        return status(restored);
      }
    );
  }
  function selected(id, allowSuspended = false) {
    const record = registry.selected(id);
    if (record.suspended && !allowSuspended) throw Error('Reopen the accepted recovery before using this workbook.');
    return record;
  }
  function handle(name, callback) {
    ipcMain.handle('xlsx:' + name, (event, ...args) => {
      if (
        !window ||
        event.sender !== window.webContents ||
        event.senderFrame !== window.webContents.mainFrame ||
        event.senderFrame.url !== page
      )
        throw Error('Untrusted spreadsheet request.');
      return callback(...args);
    });
  }
  handle('current', () => {
    rendererReady = true;
    return status(registry.active);
  });
  handle('documents', () => snapshot());
  handle('local-actions', (id, body) => {
    const request = localActionRequest(body),
      record = selected(id),
      owner = record.owner,
      epoch = record.epoch;
    const current = () =>
      !closing &&
      !closeTransaction &&
      registry.contains(record) &&
      registry.active === record &&
      record.owner === owner &&
      !owner.closed &&
      record.epoch === epoch &&
      record.status.documentId === request.documentId &&
      record.status.state === request.state;
    if (!current()) throw Error('The original Local Actions workbook changed.');
    // Library definitions are private-profile metadata, never workbook commands.
    // Capture the initiating owner before queuing; durable library acceptance
    // remains accepted even if its UI retires while the disk write completes.
    return shell(async () => {
      if (!current()) throw Error('The original Local Actions workbook changed.');
      return invokeLocalActionRequest(localActions, request);
    });
  });
  handle('spelling', (id, body) => {
    const request = spellingRequest(body),
      record = selected(id),
      owner = record.owner,
      epoch = record.epoch,
      filename = record.filename ?? null;
    const cancelling = request.operation === 'spellingCancel',
      bindings = new WeakMap();
    const present = () =>
      !closing &&
      !closeTransaction &&
      registry.contains(record) &&
      record.owner === owner &&
      !owner.closed &&
      record.epoch === epoch &&
      record.status.documentId === request.documentId &&
      record.status.view.storage?.sessionId === request.ownerId;
    const current = () =>
      present() &&
      registry.active === record &&
      record.status.state === request.state &&
      record.status.view.revision === request.revision &&
      (record.filename ?? null) === filename;
    const requireCurrent = () => {
      if (!(cancelling ? present() : current()))
        throw Error('The original spelling workbook or accepted revision changed.');
    };
    requireCurrent();
    let jobs = spellingJobs.get(owner);
    if (!jobs) {
      jobs = new WorkbookSpellingJobs();
      spellingJobs.set(owner, jobs);
    }
    const access = {
      capture: (r) => {
        requireCurrent();
        return { documentId: r.documentId, ownerId: r.ownerId, state: r.state, revision: r.revision };
      },
      authorize: async () => {
        requireCurrent();
      },
      verify: async (r) => {
        requireCurrent();
        if (cancelling) return;
        const binding = await owner.request({ operation: 'binding' });
        requireCurrent();
        const captured = bindings.get(r);
        if (captured && (captured.filename !== binding.filename || captured.sourceRevision !== binding.sourceRevision))
          throw Error('The original spelling file binding changed.');
        if (!captured) bindings.set(r, { filename: binding.filename, sourceRevision: binding.sourceRevision });
        if (binding.filename !== null && (await revision(binding.filename)) !== binding.sourceRevision)
          throw Error('The original spelling source changed.');
        requireCurrent();
      },
    };
    return cancelling ? jobs.cancel(request, access) : jobs.run(request, access);
  });
  handle('calculation', (id, request) => {
    const record = selected(id),
      epoch = record.epoch;
    if (
      !request ||
      Object.keys(request).some((k) => !['documentId', 'state', 'requestId', 'action'].includes(k)) ||
      request.documentId !== record.status.documentId ||
      request.state !== record.status.state ||
      !validId(request.requestId)
    )
      throw Error('Calculation access is not granted.');
    const action = validateCalculationAction(request.action);
    const target = (r) => ({
      record: r,
      epoch: r.epoch,
      documentId: r.status.documentId,
      ownerId: r.id,
      state: r.status.state,
      revision: r.status.view.revision,
      name: r.status.name,
      profile:
        r.status.kind === 'xlsx'
          ? 'xlsx'
          : r.status.kind === 'csv' || r.status.kind === 'tsv'
            ? 'delimited'
            : 'passive',
      readOnly: r.suspended || r.status.view.readOnly,
      canRecalculate: r.status.view.calculationPolicy?.canRecalculate,
      canChangeOptions: r.status.view.calculationPolicy?.canChangeOptions,
      canCalculateSheet: r.status.view.calculationPolicy?.canCalculateSheet,
      reason: r.status.view.calculationPolicy?.reason,
    });
    // Capture each grant/head now, before waiting behind prior shell work.
    const targets = registry.records().map(target),
      initiator = target(record);
    return shell(async () => {
      if (
        closing ||
        closeTransaction ||
        registry.active !== record ||
        !registry.contains(record) ||
        record.epoch !== epoch ||
        record.status.state !== request.state
      )
        throw Error('The initiating calculation workbook changed.');
      const controller = new AbortController();
      let receipt;
      return perform(
        action.type === 'setOptions' ? 'Updating calculation options' : 'Calculating workbooks',
        record,
        async (operation) => {
          operation.calculationRequestId = request.requestId;
          operation.calculationInitiator = id;
          operation.calculationAbort = () => controller.abort();
          receipt = await runCalculationGroup({
            requestId: request.requestId,
            initiator,
            action,
            targets,
            maxParticipants: registry.limit,
            signal: controller.signal,
            current: (t, state) =>
              registry.contains(t.record) &&
              t.record.epoch === t.epoch &&
              t.record.id === t.ownerId &&
              t.record.status.documentId === t.documentId &&
              t.record.status.state === state &&
              !t.record.owner.closed,
            authorize: async () => {},
            apply: async (t, commands) => {
              operation.record = t.record;
              try {
                const next = await t.record.owner.request(
                  { operation: 'apply', documentId: t.documentId, state: t.state, commands },
                  observe(operation, t.record.owner)
                );
                if (registry.contains(t.record) && t.record.epoch === t.epoch && next.documentId === t.documentId) {
                  t.record.status = next;
                  if (action.type === 'setOptions') {
                    const p = next.view.calculationPolicy;
                    applicationCalculation = {
                      mode: p.mode,
                      recalculateBeforeSave: p.recalculateBeforeSave,
                      revision: ++calculationRevision,
                    };
                  }
                  publishDocuments();
                }
                return {
                  documentId: next.documentId,
                  ownerId: t.record.id,
                  state: next.state,
                  revision: next.view.revision,
                };
              } catch (error) {
                const cancellation = await operation.cancelAttempt?.catch(() => undefined);
                if (cancellation?.outcome === 'cancelled')
                  throw Object.assign(Error('Calculation cancellation was granted.', { cause: error }), {
                    code: 'CALCULATION_CANCEL_GRANTED',
                  });
                throw error;
              } finally {
                await operation.cancelAttempt?.catch(() => {});
                if (!controller.signal.aborted) operation.cancelAttempt = undefined;
                operation.owner = null;
                operation.requestId = null;
                operation.record = undefined;
                operation.phase = controller.signal.aborted ? 'stopping' : 'between workbooks';
                operation.cancellable = !controller.signal.aborted;
                publishOperation();
              }
            },
            progress: (value) => {
              if (window && !window.isDestroyed()) window.webContents.send('xlsx:calculation-progress', value);
            },
          });
          return {
            ...receipt,
            ...(applicationCalculation
              ? { applicationMode: { mode: applicationCalculation.mode, revision: applicationCalculation.revision } }
              : {}),
          };
        },
        () =>
          receipt ?? {
            schema: 1,
            requestId: request.requestId,
            initiator: Object.fromEntries(['documentId', 'ownerId', 'state', 'revision'].map((k) => [k, initiator[k]])),
            action,
            consistency: 'independent-workbook-commits',
            outcome: 'cancelled',
            results: [
              {
                documentId: initiator.documentId,
                ownerId: initiator.ownerId,
                name: initiator.name,
                beforeState: initiator.state,
                status: 'cancelled',
                reason: 'Calculation cancelled before its first acceptance.',
              },
            ],
          }
      );
    });
  });
  handle('cancel-calculation', (id, documentId, requestId) => {
    const record = selected(id, true);
    if (
      record.status.documentId !== documentId ||
      !validId(requestId) ||
      active?.calculationRequestId !== requestId ||
      active.calculationInitiator !== id
    )
      throw Error('Calculation cancellation access is not granted.');
    if (!active.owner || !active.requestId) {
      active.calculationAbort();
      return { outcome: 'cancelled', document: status(registry.active) };
    }
    return cancelOperation(active.id, active.documentId);
  });
  handle('activate-tab', (id) => {
    if (closing) throw Error('Finish the close decision first.');
    return activateRecord(selected(id, true));
  });
  handle('operation', () =>
    active
      ? {
          operationId: active.id,
          documentId: active.documentId,
          label: active.label,
          phase: active.phase,
          cancellable: active.cancellable,
        }
      : null
  );
  handle('cancel-operation', (id, documentId) => cancelOperation(id, documentId));
  handle('cancel-paste', (id, documentId, requestId) => {
    const record = selected(id, true);
    if (record.status.documentId !== documentId || !validId(requestId))
      throw Error('Paste cancellation access is not granted.');
    if (!active || active.record !== record || active.pasteRequestId !== requestId)
      return { outcome: 'completed', document: status(record) };
    return cancelOperation(active.id, active.documentId);
  });
  handle('cancel-print', (id, documentId, requestId) => {
    const record = selected(id);
    if (
      record.status.documentId !== documentId ||
      !validId(requestId) ||
      active?.printRequestId !== requestId ||
      active.record !== record
    )
      throw Error('Print cancellation access is not granted.');
    return cancelOperation(active.id, id);
  });
  handle('recent-files', () => recentFiles.list());
  handle('recovery-files', () => recoveryFiles());
  handle('new', () =>
    shell(() =>
      perform('Creating workbook', null, async (operation) => {
        registry.assertRoom();
        await retire(pending);
        pending = null;
        pending = await prepare(
          undefined,
          { newWorkbook: { name: 'Untitled ' + ++untitled }, prepared: true },
          operation
        );
        return status(pending);
      })
    )
  );
  handle('open-recovery', (id) =>
    shell(async () => {
      const entry = grantedRecovery(id);
      registry.assertRoom();
      if (entry.displayPath && registry.byFilename(entry.displayPath))
        throw Error('That file is already open in another workbook tab.');
      return perform('Recovering workbook', null, async (operation) => {
        await retire(pending);
        pending = null;
        pending = await prepareGrantedRecovery(entry, operation);
        return status(pending);
      });
    })
  );
  handle('remove-recent', async (id) => {
    const entry = (await recentFiles.list()).entries.find((entry) => entry.id === id);
    if (entry) await pinLegacyBinding(entry.displayPath, createHash('sha256').update(entry.displayPath).digest('hex'));
    await recentFiles.remove(id);
    return recentFiles.list();
  });
  handle('open-recent', (id) =>
    shell(async () => {
      const filename = await recentFiles.resolve(id);
      await retire(pending);
      pending = null;
      try {
        return await perform('Opening workbook', null, async (operation) => {
          pending = await prepareReplacement(filename, operation);
          return status(pending);
        });
      } catch (error) {
        throw Error(
          'This recent spreadsheet could not be opened. The current spreadsheet is still open. ' +
            String(error.message ?? error)
        );
      }
    })
  );
  handle('open', () =>
    shell(async () => {
      const result = await dialog.showOpenDialog(window, {
        properties: ['openFile'],
        filters: [{ name: 'Spreadsheets', extensions: ['xlsx', 'xlsm', 'xls', 'xlsb', 'csv', 'tsv'] }],
      });
      if (result.canceled) return null;
      if (!Array.isArray(result.filePaths) || result.filePaths.length !== 1)
        throw Error('Choose one spreadsheet file.');
      await retire(pending);
      pending = null;
      return perform('Opening workbook', null, async (operation) => {
        pending = await prepareReplacement(result.filePaths[0], operation);
        return status(pending);
      });
    })
  );
  handle('reload', (id) =>
    shell(async () => {
      const record = selected(id);
      if (!record.filename) throw Error('Save this workbook before reloading a file.');
      await retire(pending);
      pending = {
        id: randomUUID(),
        filename: record.filename,
        reload: true,
        replaceId: record.id,
        status: record.status,
      };
      return status(pending);
    })
  );
  handle('activate-open', (id, discardPrevious = false) =>
    shell(async () => {
      if (!pending || pending.id !== id) throw Error('No prepared document is waiting to open.');
      if (typeof discardPrevious !== 'boolean') throw Error('Invalid replacement decision.');
      if (pending.reuseId) {
        const reused = selected(pending.reuseId, true);
        pending = null;
        await remember(reused);
        return activateRecord(reused);
      }
      const previous = pending.replaceId ? selected(pending.replaceId, true) : null;
      if (previous?.status.dirty && !previous.suspended && !discardPrevious)
        throw Error('Save or discard that spreadsheet before reloading it.');
      if (!previous) registry.assertRoom();
      return perform(
        'Opening workbook',
        null,
        async (operation) => {
          if (previous && !discardPrevious) await pinRecordRecovery(previous);
          const candidate = pending.reload
            ? await prepare(pending.filename, { fresh: true, prepared: true }, operation)
            : pending;
          try {
            candidate.status = await candidate.owner.request(
              { operation: 'activate' },
              observe(operation, candidate.owner)
            );
          } catch (error) {
            if (candidate !== pending) await retire(candidate);
            throw error;
          }
          try {
            if (previous) registry.replace(previous, candidate);
            else registry.add(candidate);
          } catch (error) {
            await retire(candidate);
            pending = null;
            throw error;
          }
          pending = null;
          activateRecord(candidate);
          if (discardPrevious && !previous?.suspended)
            try {
              previous.status = await previous.owner.request({ operation: 'discard' });
              previous.discardedRecovery = true;
              if (/^[a-f0-9]{64}$/.test(previous.status.documentId))
                await keptRecovery
                  .remove(previous.status.documentId)
                  .catch((error) => console.warn('xlsx.kept-recovery-cleanup', String(error)));
            } catch (error) {
              candidate.status = {
                ...candidate.status,
                warning:
                  'The reload opened, but prior recovery cleanup failed. Its data was retained where possible. ' +
                  String(error.message ?? error),
              };
            }
          await retire(previous);
          await remember(candidate);
          publishDocuments();
          return status(candidate);
        },
        () => {
          pending = null;
          return null;
        }
      );
    })
  );
  handle('cancel-open', (id) =>
    shell(async () => {
      if (!pending || pending.id !== id) throw Error('No prepared document is waiting to open.');
      await retire(pending);
      pending = null;
    })
  );
  handle('save', (id) => shell(() => saveSelected(id)));
  handle('import-delimited', (id, request) =>
    shell(async () => {
      const record = selected(id),
        epoch = record.epoch,
        state = record.status.state;
      if (
        closeTransaction ||
        !request ||
        Object.keys(request).some((key) => !['documentId', 'state', 'sheet', 'start', 'requestId'].includes(key)) ||
        record.status.kind !== 'xlsx' ||
        record.status.view.readOnly ||
        request.documentId !== record.status.documentId ||
        request.state !== state ||
        !validId(request.requestId) ||
        !record.status.view.sheets.some((sheet) => sheet.name === request.sheet)
      )
        throw Error('Delimited import access is not granted.');
      cellPosition(request.start);
      const checkCurrent = () => {
        if (
          !registry.contains(record) ||
          record.owner.closed ||
          record.epoch !== epoch ||
          record.status.state !== state ||
          registry.active !== record
        )
          throw Error('The initiating workbook changed before file import acceptance.');
      };
      checkCurrent();
      const chosen = await dialog.showOpenDialog(window, {
        properties: ['openFile'],
        filters: [{ name: 'CSV or TSV data', extensions: ['csv', 'tsv'] }],
      });
      checkCurrent();
      if (chosen.canceled) return { cancelled: true, document: status(record) };
      if (!Array.isArray(chosen.filePaths) || chosen.filePaths.length !== 1 || typeof chosen.filePaths[0] !== 'string')
        throw Error('Choose exactly one CSV or TSV source file.');
      return perform(
        'Importing CSV/TSV at ' + request.sheet + '!' + request.start,
        record,
        async (operation) => {
          const controller = new AbortController();
          operation.abortPrint = () => controller.abort();
          operation.cancellable = true;
          operation.phase = 'checking-file';
          publishOperation();
          let prepared,
            accepted = false,
            lastProgress = '';
          const cleanupRecord = () => {
            const latest = registry.records().find((item) => item.tabId === record.tabId);
            if (!latest || latest.status.documentId !== request.documentId || latest.owner.closed)
              throw Error('The import cleanup owner changed.');
            return latest;
          };
          const prune = async () => {
            const latest = cleanupRecord();
            return latest.owner.request({
              operation: 'importPrune',
              documentId: latest.status.documentId,
              state: latest.status.state,
            });
          };
          const bound = { documentId: request.documentId, state };
          const sink = {
            begin: (input) =>
              record.owner.request({
                operation: 'importBegin',
                ...bound,
                shape: input.shape,
                count: input.count,
                policy: input.policy,
              }),
            upload: (input) =>
              record.owner.request({
                operation: 'importUpload',
                ...bound,
                uploadId: input.id,
                sequence: input.sequence,
                cells: input.cells,
              }),
            seal: (input) => record.owner.request({ operation: 'importSeal', ...bound, uploadId: input.id }),
            cancel: (input) => record.owner.request({ operation: 'importCancel', ...bound, uploadId: input.id }),
            prune,
          };
          try {
            prepared = await prepareDelimitedFileImport({
              source: chosen.filePaths[0],
              directory: path.join(app.getPath('userData'), 'file-import-staging'),
              sheet: request.sheet,
              start: request.start,
              state,
              sink,
              signal: controller.signal,
              checkCurrent,
              onProgress: (event) => {
                if (operation.cancelled) return;
                const percent = event.total > 0 ? Math.min(100, Math.floor((event.completed * 100) / event.total)) : 0,
                  key = event.phase + ':' + percent;
                if (key === lastProgress) return;
                lastProgress = key;
                operation.phase = event.phase + ' ' + percent + '%';
                operation.cancellable = true;
                publishOperation();
              },
            });
            await prepared.check();
            checkCurrent();
            controller.signal.throwIfAborted();
            // Staging aborts do not kill an owner. Applying uses its authoritative
            // worker grant, including the journal/publication critical boundary.
            operation.abortPrint = null;
            operation.cancellable = false;
            operation.phase = 'applying';
            publishOperation();
            record.status = await record.owner.request(
              { operation: 'apply', state, commands: [prepared.descriptor] },
              observe(operation, record.owner)
            );
            record.discardedRecovery = false;
            record.operationCancelled = false;
            accepted = true;
            return {
              imported: true,
              document: status(record),
              source: prepared.source,
              shape: prepared.descriptor.shape,
              count: prepared.descriptor.count,
              delimiter: prepared.delimiter,
              encoding: prepared.encoding,
            };
          } finally {
            if (!accepted && prepared) {
              await operation.cancelAttempt?.catch(() => {});
              await prune().catch((error) => console.warn('xlsx.import-file-prune', String(error)));
            }
            try {
              await prepared?.dispose();
            } catch (error) {
              console.warn('xlsx.import-file-cleanup', String(error));
              if (!accepted && !operation.cancelled) throw error;
            }
          }
        },
        () => {
          const restored = registry.records().find((item) => item.tabId === record.tabId);
          if (restored) restored.operationCancelled = true;
          return { cancelled: true, document: status(restored) };
        }
      );
    })
  );
  handle('query-connection', (id, request) =>
    shell(async () => {
      try {
        const record = selected(id),
          epoch = record.epoch,
          state = record.status.state,
          documentId = record.status.documentId;
        if (
          closeTransaction ||
          !request ||
          Object.keys(request).some(
            (key) =>
              ![
                'operation',
                'documentId',
                'state',
                'connectionId',
                'expectedRevision',
                'draft',
                'requestId',
                'offset',
                'replacementOffset',
                'expectedConversionRevision',
                'expectedReplacementRevision',
              ].includes(key)
          ) ||
          !['list', 'create', 'update', 'remove', 'preview', 'refresh'].includes(request.operation) ||
          record.status.kind !== 'xlsx' ||
          record.status.view.readOnly ||
          request.documentId !== documentId ||
          request.state !== state
        )
          throw Error('Local query access is not granted.');
        if (
          ['update', 'remove', 'preview', 'refresh'].includes(request.operation) &&
          (!validId(request.connectionId) ||
            !Number.isSafeInteger(request.expectedRevision) ||
            request.expectedRevision < 1)
        )
          throw Error('Invalid query connection guard.');
        if (['preview', 'refresh'].includes(request.operation) && !validId(request.requestId))
          throw Error('Invalid query request identity.');
        if (
          request.offset !== undefined &&
          (request.operation !== 'preview' ||
            !Number.isSafeInteger(request.offset) ||
            request.offset < 0 ||
            request.offset > 1000000)
        )
          throw Error('Invalid query conversion page.');
        if (
          request.replacementOffset !== undefined &&
          (request.operation !== 'preview' ||
            !Number.isSafeInteger(request.replacementOffset) ||
            request.replacementOffset < 0 ||
            request.replacementOffset > 1000000 ||
            request.offset !== undefined)
        )
          throw Error('Invalid query replacement page.');
        if (
          request.expectedConversionRevision !== undefined &&
          (!['preview', 'refresh'].includes(request.operation) || !validId(request.expectedConversionRevision))
        )
          throw Error('Invalid query conversion review.');
        if (
          request.expectedReplacementRevision !== undefined &&
          (request.operation !== 'refresh' || !validId(request.expectedReplacementRevision))
        )
          throw Error('Invalid query replacement review.');
        if (['create', 'update'].includes(request.operation)) {
          if (!request.draft || !record.status.view.sheets.some((sheet) => sheet.name === request.draft.target?.sheet))
            throw Error('Query target worksheet is unavailable.');
          cellPosition(request.draft.target.start);
        }
        const latest = () => {
          const value = registry.records().find((r) => r.tabId === record.tabId);
          if (!value || value.status.documentId !== documentId || value.owner.closed)
            throw Error('The query document owner is no longer available.');
          return value;
        };
        const check = () => {
          if (
            latest() !== record ||
            record.epoch !== epoch ||
            record.status.state !== state ||
            registry.active !== record
          )
            throw Error('The initiating workbook changed before query acceptance.');
        };
        check();
        const service = (operation) =>
          new LocalQueryHost({
            store: queryConnections,
            document: 'standalone:' + documentId,
            directory: path.join(app.getPath('userData'), 'query-staging'),
            context: {
              current: (fenced) => {
                if (fenced !== false) check();
                const value = latest();
                return { state: value.status.state, documentId, ownerId: value.status.view.storage?.sessionId };
              },
              calendar: (sheet) => {
                check();
                const target = record.status.view.sheets.find((value) => value.name === sheet);
                if (!target) throw Error('Query target worksheet is unavailable.');
                return typeof target.date1904 === 'boolean' ? target.date1904 : undefined;
              },
              request: (name, input, fenced = true) => {
                if (fenced) check();
                return latest().owner.request({ operation: name, documentId, state: latest().status.state, ...input });
              },
              apply: async (parent, descriptor) => {
                check();
                record.status = await record.owner.request(
                  { operation: 'apply', state: parent, commands: [descriptor] },
                  observe(operation, record.owner)
                );
                record.discardedRecovery = false;
                record.operationCancelled = false;
                return status(record);
              },
            },
          });
        if (request.operation === 'list') return service().list();
        if (request.operation === 'create')
          return service().create(request.draft, async () => {
            const chosen = await dialog.showOpenDialog(window, {
              properties: ['openFile'],
              filters: [{ name: 'Local query source', extensions: ['csv', 'tsv'] }],
            });
            check();
            if (chosen.canceled) return null;
            if (chosen.filePaths.length !== 1) throw Error('Choose one local CSV/TSV source.');
            return chosen.filePaths[0];
          });
        if (request.operation === 'update')
          return service().update(request.connectionId, request.expectedRevision, request.draft);
        if (request.operation === 'remove') return service().remove(request.connectionId, request.expectedRevision);
        return perform(
          request.operation === 'refresh' ? 'Refreshing local query' : 'Previewing local query',
          record,
          async (operation) => {
            const controller = new AbortController();
            operation.queryRequestId = request.requestId;
            operation.abortPrint = () => controller.abort();
            operation.cancellable = true;
            publishOperation();
            let lastProgress = '';
            const options = {
              signal: controller.signal,
              ...(request.offset !== undefined ? { offset: request.offset } : {}),
              ...(request.replacementOffset !== undefined ? { replacementOffset: request.replacementOffset } : {}),
              ...(request.expectedConversionRevision !== undefined
                ? { expectedConversionRevision: request.expectedConversionRevision }
                : {}),
              ...(request.expectedReplacementRevision !== undefined
                ? { expectedReplacementRevision: request.expectedReplacementRevision }
                : {}),
              onProgress: (event) => {
                if (operation.cancelled) return;
                const percent = event.total ? Math.min(100, Math.floor((event.completed * 100) / event.total)) : 100,
                  key = event.phase + ':' + percent;
                if (key === lastProgress) return;
                lastProgress = key;
                operation.phase = event.phase + ' ' + percent + '%';
                publishOperation();
              },
              beforeApply: () => {
                check();
                controller.signal.throwIfAborted();
                operation.abortPrint = null;
                operation.cancellable = false;
                operation.phase = 'applying';
                publishOperation();
              },
            };
            try {
              const result =
                request.operation === 'refresh'
                  ? await service(operation).refresh(request.connectionId, request.expectedRevision, options)
                  : await service(operation).preview(request.connectionId, request.expectedRevision, options);
              const { acceptance, ...publicResult } = result;
              return { ...publicResult, document: status(latest()) };
            } catch (error) {
              if (error?.code) throw Error('Local query could not read its granted source (' + error.code + ').');
              throw error;
            }
          },
          () => ({ cancelled: true, document: status(latest()) })
        );
      } catch (error) {
        if (
          ['ENOENT', 'ENOTDIR', 'EACCES', 'EPERM', 'EIO', 'EISDIR', 'ELOOP', 'EMFILE', 'ENFILE'].includes(error?.code)
        )
          throw Error('Local query could not read its granted source or private registry (' + error.code + ').');
        throw error;
      }
    })
  );
  handle('query-refresh-all', (id, body) => {
    const request = queryRefreshAllRequest(body);
    return shell(async () => {
      const record = selected(id),
        owner = record.owner,
        epoch = record.epoch,
        documentId = record.status.documentId;
      const checkOwner = () => {
        if (
          closing ||
          closeTransaction ||
          !registry.contains(record) ||
          registry.active !== record ||
          record.owner !== owner ||
          owner.closed ||
          record.epoch !== epoch ||
          record.status.documentId !== documentId ||
          record.status.kind !== 'xlsx' ||
          record.status.view.readOnly ||
          record.status.view.passiveNotice
        )
          throw Error('The original Refresh All workbook owner retired.');
      };
      checkOwner();
      if (request.documentId !== documentId || request.state !== record.status.state)
        throw Error('Refresh All original workbook state changed. Preview again.');
      let entry = refreshAllJobs.get(owner);
      if (entry && entry.epoch !== epoch) {
        entry.coordinator.dispose();
        refreshAllJobs.delete(owner);
        entry = null;
      }
      if (!entry) {
        if (request.operation !== 'preview')
          throw Error('Preview Refresh All on the current workbook before running it.');
        entry = { epoch, operation: null, coordinator: null };
        const owned = entry;
        const context = {
          current: () => {
            checkOwner();
            return { documentId, ownerId: record.status.view.storage?.sessionId, state: record.status.state };
          },
          calendar: (sheet) => {
            checkOwner();
            const target = record.status.view.sheets.find((value) => value.name === sheet);
            if (!target) throw Error('Query target worksheet is unavailable.');
            return target.date1904;
          },
          request: async (name, input) => {
            checkOwner();
            const result = await owner.request({ operation: name, documentId, state: record.status.state, ...input });
            checkOwner();
            return result;
          },
          apply: async (parent, descriptor) => {
            checkOwner();
            if (parent !== record.status.state) throw Error('Refresh All original parent changed.');
            const callbacks = observe(owned.operation, owner),
              progress = callbacks.onProgress;
            callbacks.onProgress = (event) => progress({ ...event, cancellable: true });
            // Cancel requests stop subsequent units; they never kill this
            // currently dispatched native acceptance/recovery transaction.
            record.status = await owner.request(
              { operation: 'apply', documentId, state: parent, commands: [descriptor] },
              callbacks
            );
            record.discardedRecovery = false;
            record.operationCancelled = false;
            checkOwner();
            return status(record);
          },
        };
        owned.coordinator = new LocalQueryRefreshAll({
          store: queryConnections,
          document: 'standalone:' + documentId,
          directory: path.join(app.getPath('userData'), 'query-staging'),
          context,
        });
        refreshAllJobs.set(owner, owned);
      }
      return perform(
        request.operation === 'preview' ? 'Previewing all local queries' : 'Refreshing all local queries',
        record,
        async (operation) => {
          const controller = new AbortController();
          operation.queryRequestId = request.requestId;
          operation.abortPrint = () => controller.abort();
          operation.cancellable = true;
          entry.operation = operation;
          publishOperation();
          const options = {
            signal: controller.signal,
            onProgress: (event) => {
              if (operation.cancelled) return;
              operation.phase =
                event.phase +
                (event.total ? ' ' + Math.min(100, Math.floor((event.completed * 100) / event.total)) + '%' : '');
              publishOperation();
            },
            beforeApply: () => {
              checkOwner();
              controller.signal.throwIfAborted();
              operation.phase = 'Applying current connection; Cancel stops remaining connections';
              operation.cancellable = true;
              publishOperation();
            },
            afterApply: () => {
              if (!operation.cancelled) {
                operation.phase = 'Connection accepted; checking remaining connections';
                operation.cancellable = true;
                publishOperation();
              }
            },
          };
          try {
            const result =
              request.operation === 'preview'
                ? await entry.coordinator.preview(options)
                : await entry.coordinator.refresh(request.previewId, options);
            return { ...result, document: status(record) };
          } finally {
            entry.operation = null;
          }
        },
        () => ({ cancelled: true, document: status(record) })
      );
    });
  });
  handle('cancel-query', (id, documentId, requestId) => {
    const record = selected(id, true);
    if (record.status.documentId !== documentId || !validId(requestId))
      throw Error('Query cancellation access is not granted.');
    if (!active || active.record?.tabId !== record.tabId || active.queryRequestId !== requestId)
      return { outcome: 'completed' };
    return cancelOperation(active.id, active.documentId);
  });
  handle('request', async (id, request) => {
    const record = selected(id);
    if (request?.operation === 'accessibility') {
      const captured = Object.freeze(structuredClone(request));
      accessibilityFields(captured);
      const owner = record.owner,
        epoch = record.epoch,
        sourceView = record.status.view,
        receipt = Object.freeze(
          noteMembershipReceipt({
            documentId: record.status.documentId,
            ownerId: record.status.view.storage?.sessionId,
            state: record.status.state,
            revision: record.status.view.revision,
          })
        );
      const current = () => {
        if (
          closing ||
          closeTransaction ||
          !registry.contains(record) ||
          selected(id) !== record ||
          registry.active !== record ||
          record.owner !== owner ||
          owner.closed ||
          record.epoch !== epoch ||
          record.status.kind !== 'xlsx' ||
          record.status.view !== sourceView ||
          record.status.view.passiveNotice ||
          record.status.view.storage?.kind !== 'rust-indexed-v1' ||
          record.status.view.storage.sessionId !== receipt.ownerId ||
          record.status.documentId !== receipt.documentId ||
          captured.documentId !== receipt.documentId ||
          captured.state !== receipt.state ||
          record.status.state !== receipt.state ||
          record.status.view.revision !== receipt.revision
        )
          throw Error('The accessibility owner or accepted revision changed.');
      };
      current();
      const result = await owner.request(captured);
      current();
      return { accessibility: validateAccessibility(result.accessibility, receipt) };
    }
    if (request?.operation === 'notes') {
      const captured = Object.freeze(structuredClone(request));
      notesFields(captured);
      const owner = record.owner,
        epoch = record.epoch,
        sourceView = record.status.view,
        receipt = Object.freeze(
          noteMembershipReceipt({
            documentId: record.status.documentId,
            ownerId: record.status.view.storage?.sessionId,
            state: record.status.state,
            revision: record.status.view.revision,
          })
        );
      const current = () => {
        if (
          closing ||
          closeTransaction ||
          !registry.contains(record) ||
          selected(id) !== record ||
          registry.active !== record ||
          record.owner !== owner ||
          owner.closed ||
          record.epoch !== epoch ||
          record.status.kind !== 'xlsx' ||
          record.status.view !== sourceView ||
          record.status.view.passiveNotice ||
          record.status.view.storage?.kind !== 'rust-indexed-v1' ||
          record.status.view.storage.sessionId !== receipt.ownerId ||
          record.status.documentId !== receipt.documentId ||
          captured.documentId !== receipt.documentId ||
          captured.state !== receipt.state ||
          record.status.state !== receipt.state ||
          record.status.view.revision !== receipt.revision
        )
          throw Error('The note owner or accepted revision changed.');
      };
      current();
      const result = await owner.request(captured);
      current();
      return { notes: validateNotes(result.notes, receipt, { sheet: captured.sheet }) };
    }
    if (request?.operation === 'noteMembership') {
      const captured = Object.freeze({ ...request });
      noteMembershipFields(captured);
      const owner = record.owner,
        epoch = record.epoch,
        sourceView = record.status.view,
        privateStorageState = sourceView.storage?.state,
        receipt = Object.freeze(
          noteMembershipReceipt({
            documentId: record.status.documentId,
            ownerId: record.status.view.storage?.sessionId,
            state: record.status.state,
            revision: record.status.view.revision,
          })
        );
      const current = () => {
        if (
          closing ||
          closeTransaction ||
          !registry.contains(record) ||
          selected(id) !== record ||
          record.owner !== owner ||
          owner.closed ||
          record.epoch !== epoch ||
          registry.active !== record ||
          record.status.kind !== 'xlsx' ||
          record.status.view !== sourceView ||
          record.status.view.passiveNotice ||
          record.status.view.storage?.kind !== 'rust-indexed-v1' ||
          record.status.view.storage.sessionId !== receipt.ownerId ||
          record.status.view.storage.state !== privateStorageState ||
          record.status.documentId !== receipt.documentId ||
          captured.documentId !== receipt.documentId ||
          captured.state !== receipt.state ||
          record.status.state !== receipt.state ||
          record.status.view.revision !== receipt.revision
        )
          throw Error('The note membership owner or accepted revision changed.');
      };
      const binding = async () => {
        current();
        const value = await owner.request({ operation: 'binding' });
        current();
        if (
          !value ||
          Object.keys(value).length !== 2 ||
          !Object.hasOwn(value, 'filename') ||
          !Object.hasOwn(value, 'sourceRevision') ||
          (value.filename !== null && typeof value.filename !== 'string') ||
          (value.sourceRevision !== null &&
            (typeof value.sourceRevision !== 'string' || !/^[a-f0-9]{64}$/.test(value.sourceRevision)))
        )
          throw Error('Invalid note membership source binding.');
        return Object.freeze({ ...value });
      };
      current();
      const original = await binding();
      current();
      const result = await owner.request(captured);
      current();
      const completed = await binding();
      current();
      if (completed.filename !== original.filename || completed.sourceRevision !== original.sourceRevision)
        throw Error('The note membership source revision changed.');
      const noteMembership = validateNoteMembership(result.noteMembership, { sheet: captured.sheet, receipt });
      return { noteMembership };
    }
    if (request?.operation === 'usedRange') {
      const captured = Object.freeze({ ...request });
      usedRangeFields(captured);
      const owner = record.owner,
        epoch = record.epoch,
        receipt = Object.freeze({
          documentId: record.status.documentId,
          ownerId: record.status.view.storage?.sessionId,
          state: record.status.state,
          revision: record.status.view.revision,
        });
      const current = () => {
        if (
          closing ||
          closeTransaction ||
          !registry.contains(record) ||
          selected(id) !== record ||
          record.owner !== owner ||
          owner.closed ||
          record.epoch !== epoch ||
          registry.active !== record ||
          record.status.kind !== 'xlsx' ||
          record.status.view.storage?.kind !== 'rust-indexed-v1' ||
          record.status.view.storage.sessionId !== receipt.ownerId ||
          record.status.documentId !== receipt.documentId ||
          captured.documentId !== receipt.documentId ||
          captured.state !== receipt.state ||
          record.status.state !== receipt.state ||
          record.status.view.revision !== receipt.revision
        )
          throw Error('The UsedRange owner or accepted revision changed.');
      };
      current();
      const result = await owner.request(captured);
      current();
      validateUsedRange(result.usedRange, { sheet: captured.sheet, receipt });
      return result;
    }
    if (request?.operation === 'duplicatePreview') {
      const captured = { ...request, columns: Array.isArray(request.columns) ? [...request.columns] : request.columns };
      if (
        Object.keys(captured).some(
          (k) => !['operation', 'documentId', 'state', 'sheet', 'range', 'hasHeaders', 'columns', 'offset'].includes(k)
        )
      )
        throw Error('Invalid duplicate preview request fields');
      duplicateConfiguration(captured.sheet, captured.range, captured.hasHeaders, captured.columns);
      const epoch = record.epoch,
        receipt = {
          documentId: record.status.documentId,
          ownerId: record.status.view.storage?.sessionId,
          state: record.status.state,
          revision: record.status.view.revision,
        };
      const current = () => {
        if (
          !registry.contains(record) ||
          selected(id) !== record ||
          record.owner.closed ||
          record.epoch !== epoch ||
          record.status.kind !== 'xlsx' ||
          captured.documentId !== receipt.documentId ||
          captured.state !== receipt.state ||
          record.status.state !== receipt.state ||
          record.status.view.revision !== receipt.revision ||
          registry.active !== record
        )
          throw Error('The duplicate preview owner or accepted revision changed.');
      };
      current();
      const result = await record.owner.request(captured);
      current();
      validateDuplicatePreview(result.duplicatePreview, receipt, captured, captured.offset ?? 0);
      return result;
    }
    if (request?.operation === 'previewFormula' || request?.operation === 'cancelFormulaPreview') {
      const captured = Object.freeze({ ...request }),
        cancel = captured.operation === 'cancelFormulaPreview';
      formulaPreviewFields(captured, { cancel });
      const epoch = record.epoch,
        receipt = {
          documentId: record.status.documentId,
          ownerId: record.status.view.storage?.sessionId,
          state: record.status.state,
          revision: record.status.view.revision,
        };
      const current = () => {
        if (
          !registry.contains(record) ||
          selected(id) !== record ||
          record.owner.closed ||
          record.epoch !== epoch ||
          record.status.kind !== 'xlsx' ||
          captured.documentId !== receipt.documentId ||
          captured.state !== receipt.state ||
          record.status.state !== receipt.state ||
          record.status.view.revision !== receipt.revision ||
          (!cancel && registry.active !== record)
        )
          throw Error('The formula preview owner or accepted revision changed.');
      };
      current();
      const result = await record.owner.request(captured);
      current();
      const decoded = decodeOwnerResult(result);
      if (cancel) formulaPreviewCancellation(decoded.formulaPreviewCancellation, captured.requestId);
      else validateFormulaPreview(decoded.formulaPreview, { ...captured, receipt });
      // Readonly preview has its own cancellation, never the global busy operation.
      return result;
    }
    if (
      closeTransaction &&
      ['apply', 'undo', 'redo', 'discard', 'importBegin', 'importUpload', 'importSeal'].includes(request?.operation)
    )
      throw Error('Finish the close decision before editing.');
    const allowed = [
      'status',
      'apply',
      'undo',
      'redo',
      'discard',
      'read',
      'scan',
      'trace',
      'query',
      'selection',
      'filterValues',
      'hyperlink',
      'objects',
      'object',
      'vbaModules',
      'vbaModule',
      'vbaFunctions',
      'evaluateVba',
      'conditionalRules',
      'validationRules',
      'importBegin',
      'importUpload',
      'importSeal',
      'importCancel',
      'importPrune',
      'printDefinition',
      'printPrepare',
      'printPage',
      'printRelease',
    ];
    if (
      !request ||
      !allowed.includes(request.operation) ||
      Object.hasOwn(request, 'destination') ||
      Object.hasOwn(request, 'expectedRevision')
    )
      throw Error('Unsupported spreadsheet request.');
    // Registered actions carry the original grant/head. Legacy point edits
    // omit documentId and keep their existing serialized FIFO behavior.
    const strictCommand = ['apply', 'undo', 'redo'].includes(request.operation) && request.documentId !== undefined,
      epoch = record.epoch;
    const checkStrictCommand = () => {
      if (
        strictCommand &&
        (!registry.contains(record) ||
          selected(id) !== record ||
          record.owner.closed ||
          record.epoch !== epoch ||
          registry.active !== record ||
          request.documentId !== record.status.documentId ||
          request.state !== record.status.state)
      )
        throw Error('The initiating workbook or accepted revision changed before this command.');
    };
    checkStrictCommand();
    const stagedPaste =
      request.operation === 'apply' &&
      request.commands?.length === 1 &&
      request.commands[0]?.type === 'importCells' &&
      request.requestId !== undefined;
    if (stagedPaste && (!strictCommand || !validId(request.requestId)))
      throw Error('Invalid staged paste request identity.');
    if (
      [
        'trace',
        'hyperlink',
        'objects',
        'object',
        'vbaModules',
        'vbaModule',
        'vbaFunctions',
        'evaluateVba',
        'conditionalRules',
        'validationRules',
        'importBegin',
        'importUpload',
        'importSeal',
        'importCancel',
        'importPrune',
        'printDefinition',
        'printPrepare',
        'printPage',
        'printRelease',
      ].includes(request.operation) &&
      request.documentId !== record.status.documentId
    )
      throw Error('Spreadsheet request belongs to another document.');
    if (request.operation === 'printPrepare' && !validId(request.requestId))
      throw Error('Invalid print preparation request.');
    if (request.operation === 'filterValues') {
      if (record.status.kind !== 'xlsx' || !record.status.view.sheets.some((sheet) => sheet.name === request.sheet))
        throw Error('Filter value pages require a granted XLSX worksheet.');
      const bounds = rangeBounds(request.range),
        column = request.column,
        query = request.query ?? '',
        offset = request.offset ?? 0,
        limit = request.limit ?? 100;
      if (
        typeof column !== 'string' ||
        !/^([A-Z]{1,3})$/.test(column) ||
        typeof query !== 'string' ||
        Buffer.byteLength(query) > 32768 ||
        !Number.isSafeInteger(offset) ||
        offset < 0 ||
        !Number.isInteger(limit) ||
        limit < 1 ||
        limit > 100
      )
        throw Error('Invalid filter value page request.');
      const col = cellPosition(column + '1').col;
      if (col < bounds.first.col || col > bounds.last.col) throw Error('Filter column must be inside its range.');
    }
    const mutating = ['apply', 'undo', 'redo', 'discard'].includes(request.operation);
    const observed = mutating || ['printPrepare', 'evaluateVba'].includes(request.operation);
    const serialized =
      observed ||
      [
        'importBegin',
        'importUpload',
        'importSeal',
        'importCancel',
        'importPrune',
        'printDefinition',
        'printPage',
        'printRelease',
        'conditionalRules',
        'validationRules',
      ].includes(request.operation);
    const run = async () => {
      checkStrictCommand();
      const selectedRecord = selected(id);
      const work = async (operation) => {
        if (operation && request.operation === 'printPrepare') operation.printRequestId = request.requestId;
        if (operation && stagedPaste) operation.pasteRequestId = request.requestId;
        const result = await selectedRecord.owner.request(
          request,
          operation ? observe(operation, selectedRecord.owner) : undefined
        );
        if (!registry.contains(selectedRecord))
          throw Error('The workbook owner changed before its response completed.');
        if (['status', 'apply', 'undo', 'redo', 'discard'].includes(request.operation)) {
          selectedRecord.status = result;
          selectedRecord.operationCancelled = false;
          if (request.operation === 'discard') {
            selectedRecord.discardedRecovery = true;
            if (/^[a-f0-9]{64}$/.test(selectedRecord.status.documentId))
              await keptRecovery
                .remove(selectedRecord.status.documentId)
                .catch((error) => console.warn('xlsx.kept-recovery-cleanup', String(error)));
          } else if (['apply', 'undo', 'redo'].includes(request.operation)) selectedRecord.discardedRecovery = false;
          return status(selectedRecord);
        }
        return result;
      };
      return observed
        ? perform(
            request.operation === 'printPrepare'
              ? 'Preparing print pages'
              : request.operation === 'apply'
                ? 'Applying edits'
                : 'Updating workbook',
            selectedRecord,
            work,
            () => {
              if (request.operation === 'printPrepare') throw Error('Print preparation cancelled.');
              const restored = registry.records().find((item) => item.tabId === record.tabId);
              if (restored) restored.operationCancelled = true;
              return status(restored);
            }
          )
        : work();
    };
    return serialized ? shell(run) : run();
  });
  handle('publish-print', (id, request) =>
    shell(async () => {
      const record = selected(id);
      if (
        !request ||
        Object.keys(request).some((key) => !['documentId', 'state', 'jobId', 'kind', 'requestId'].includes(key)) ||
        request.documentId !== record.status.documentId ||
        request.state !== record.status.state ||
        !validId(request.jobId) ||
        !validId(request.requestId) ||
        !['pdf', 'printer'].includes(request.kind)
      )
        throw Error('Print output access is not granted.');
      const checkCurrent = () => {
        if (!registry.contains(record) || record.owner.closed || record.status.state !== request.state)
          throw Error('The workbook revision changed. Prepare a new print preview.');
      };
      checkCurrent();
      // Validate the owner-resident preview before asking for an output path.
      await record.owner.request({
        operation: 'printPage',
        documentId: request.documentId,
        state: request.state,
        jobId: request.jobId,
        pageNumber: 1,
      });
      if (request.kind === 'printer')
        return perform(
          'Printing workbook',
          record,
          async (operation) => {
            operation.printRequestId = request.requestId;
            const document = await record.owner.request(
              {
                operation: 'printDocument',
                documentId: request.documentId,
                state: request.state,
                jobId: request.jobId,
              },
              observe(operation, record.owner)
            );
            checkCurrent();
            const controller = new AbortController();
            operation.abortPrint = () => controller.abort();
            operation.cancellable = true;
            publishOperation();
            return printNative(document, {
              signal: controller.signal,
              checkCurrent,
              parentWindow: window,
              onProgress: (phase) => {
                operation.phase = phase;
                publishOperation();
              },
              beforeDispatch: () => {
                checkCurrent();
                controller.signal.throwIfAborted();
                operation.cancellable = false;
                operation.phase = 'printing';
                publishOperation();
              },
            });
          },
          () => ({ cancelled: true })
        );
      const destination = await dialog.showSaveDialog(window, {
        defaultPath: path.basename(record.status.name, path.extname(record.status.name)) + '.pdf',
        filters: [{ name: 'PDF', extensions: ['pdf'] }],
      });
      checkCurrent();
      if (destination.canceled || !destination.filePath) return { cancelled: true };
      if (path.extname(destination.filePath).toLowerCase() !== '.pdf') throw Error('Choose a PDF output filename.');
      const existing = await stat(destination.filePath).catch((error) => {
          if (error.code === 'ENOENT') return null;
          throw error;
        }),
        expected = existing ? await revision(destination.filePath) : null;
      return perform(
        'Exporting PDF',
        record,
        async (operation) => {
          operation.printRequestId = request.requestId;
          const document = await record.owner.request(
            { operation: 'printDocument', documentId: request.documentId, state: request.state, jobId: request.jobId },
            observe(operation, record.owner)
          );
          checkCurrent();
          const controller = new AbortController();
          operation.abortPrint = () => controller.abort();
          operation.cancellable = true;
          publishOperation();
          const result = await renderPdf(document, {
            signal: controller.signal,
            checkCurrent,
            onProgress: (phase) => {
              operation.phase = phase;
              publishOperation();
            },
          });
          checkCurrent();
          controller.signal.throwIfAborted();
          // Output publication has one critical boundary. Cancellation before this
          // boundary preserves an existing PDF; publication uses its observed hash.
          operation.cancellable = false;
          operation.phase = 'writing-output';
          publishOperation();
          const staging = path.join(app.getPath('userData'), 'print-staging');
          await mkdir(staging, { recursive: true, mode: 0o700 });
          const owned = await mkdtemp(path.join(staging, 'export-'));
          try {
            const source = path.join(owned, 'pages.pdf');
            await writePrivateSnapshot(source, result.bytes);
            checkCurrent();
            await publishCopy(source, destination.filePath, expected);
            return { exported: true, pageCount: result.pageCount };
          } finally {
            await rm(owned, { recursive: true, force: true });
          }
        },
        () => ({ cancelled: true })
      );
    })
  );
  handle('data-only-export', (id, request) =>
    shell(async () => {
      const record = selected(id),
        state = record.status.state,
        epoch = record.epoch;
      if (
        !request ||
        Object.keys(request).some((key) => !['documentId', 'state', 'format'].includes(key)) ||
        request.documentId !== record.status.documentId ||
        request.state !== state ||
        !['xls', 'xlsb'].includes(record.status.kind) ||
        !['xls', 'xlsb', 'xlsx'].includes(request.format)
      )
        throw Error('Passive data-only export access is not granted.');
      const check = () => {
        if (
          !registry.contains(record) ||
          record.owner.closed ||
          record.epoch !== epoch ||
          record.status.state !== state ||
          registry.active !== record
        )
          throw Error('The passive workbook changed before export.');
      };
      check();
      const chosen = await dialog.showSaveDialog(window, {
        defaultPath:
          path.basename(record.status.name, path.extname(record.status.name)) + ' data-only.' + request.format,
        filters: [{ name: 'Data-only ' + request.format.toUpperCase(), extensions: [request.format] }],
      });
      check();
      if (chosen.canceled || !chosen.filePath) return { cancelled: true };
      const destination = path.join(await realpath(path.dirname(chosen.filePath)), path.basename(chosen.filePath));
      if (
        path.extname(destination).toLowerCase() !== '.' + request.format ||
        destination === record.filename ||
        registry.records().some((r) => r.filename === destination)
      )
        throw Error('Choose a new filename with the requested data-only format.');
      return perform(
        'Exporting stored values',
        record,
        (operation) =>
          record.owner.request(
            {
              operation: 'dataOnlyExport',
              documentId: record.status.documentId,
              state,
              format: request.format,
              destination,
              expectedRevision: null,
            },
            observe(operation, record.owner)
          ),
        () => ({ cancelled: true })
      );
    })
  );
  handle('copy', (id) =>
    shell(async () => {
      const record = selected(id),
        extension = record.filename ? path.extname(record.filename).slice(1).toLowerCase() : 'xlsx';
      const result = await dialog.showSaveDialog(window, {
        defaultPath:
          (record.filename ? path.basename(record.filename, path.extname(record.filename)) : record.status.name) +
          ' copy.' +
          extension,
        filters: [{ name: 'Spreadsheet', extensions: [extension] }],
      });
      if (result.canceled || !result.filePath) return false;
      const destination = await destinationFor(record, result.filePath),
        existing = await stat(destination).catch((error) => {
          if (error.code === 'ENOENT') return null;
          throw error;
        });
      return perform(
        'Saving a copy',
        record,
        async (operation) => {
          await record.owner.request(
            { operation: 'copy', destination, expectedRevision: existing ? await revision(destination) : null },
            observe(operation, record.owner)
          );
          return true;
        },
        () => false
      );
    })
  );
  handle('reopen-recovery', (id) =>
    shell(async () => {
      const record = selected(id, true);
      if (pending) throw Error('Finish the prepared replacement before reopening recovery.');
      if (!record.owner.closed) await record.owner.cancel();
      return status(await resumeRecord(record));
    })
  );
  function grantedRecovery(id) {
    if (!recoveryGrants.has(id) || registry.byDocument(id) || id === pending?.status.documentId)
      throw Error('Workbook recovery access is not granted. Refresh the recovery list.');
    return recoveryGrants.get(id);
  }
  handle('copy-recovery', (id) =>
    shell(async () => {
      if (pending) throw Error('Finish the prepared workbook before saving a recovery copy.');
      const entry = grantedRecovery(id),
        extension = entry.displayPath ? path.extname(entry.displayPath).slice(1) : 'xlsx';
      const result = await dialog.showSaveDialog(window, {
        defaultPath: entry.name + ' recovered copy.' + extension,
        filters: [{ name: 'Spreadsheet', extensions: [extension] }],
      });
      if (result.canceled || !result.filePath) return false;
      const destination = await destinationFor(null, result.filePath),
        existing = await stat(destination).catch((error) => {
          if (error.code === 'ENOENT') return null;
          throw error;
        });
      return perform(
        'Saving recovery copy',
        null,
        async (operation) => {
          const record = await prepareGrantedRecovery(entry, operation);
          try {
            await record.owner.request(
              { operation: 'copy', destination, expectedRevision: existing ? await revision(destination) : null },
              observe(operation, record.owner)
            );
            return true;
          } finally {
            await record.owner.cancel();
          }
        },
        () => false
      );
    })
  );
  handle('discard-recovery', (id) =>
    shell(async () => {
      const entry = grantedRecovery(id),
        result = await dialog.showMessageBox(window, {
          type: 'warning',
          buttons: ['Keep recovery', 'Discard recovery'],
          defaultId: 0,
          cancelId: 0,
          message: 'Discard recovery for ' + entry.name + '?',
          detail: 'This removes its accepted recovery edits. The original spreadsheet file is kept.',
        });
      if (result.response !== 1) return false;
      // Validate the exact UUID manifest again. Corrupt data is preserved, never
      // interpreted as permission to delete an arbitrary directory or user file.
      grantedRecovery(id);
      if (entry.legacyFilename) {
        const record = await prepareGrantedRecovery(entry);
        try {
          await record.owner.request({ operation: 'discard' });
        } finally {
          await retire(record);
        }
      } else await discardSpreadsheetRecovery(recoveryDirectory, id);
      if (entry.legacyFilename)
        await keptRecovery.remove(id).catch((error) => console.warn('xlsx.kept-recovery-cleanup', String(error)));
      recoveryGrants.delete(id);
      return true;
    })
  );
  async function leaveChoice(record, allowKeepRecovery = true) {
    const state = record.status.state,
      choices = ['save', 'cancel', 'discard', ...(allowKeepRecovery ? ['keep'] : [])];
    const { response } = await dialog.showMessageBox(window, {
      type: 'question',
      buttons: ['Save', 'Cancel', 'Discard changes', ...(allowKeepRecovery ? ['Keep recovery and close'] : [])],
      defaultId: 0,
      cancelId: 1,
      message: 'Save changes before leaving ' + record.status.name + '?',
      detail: 'Accepted edits already have a recovery journal. Discard removes that journal.',
    });
    if (!registry.contains(record) || record.status.state !== state)
      throw Error('The workbook changed during its close decision.');
    return choices[response] ?? 'cancel';
  }
  handle('confirm-leave', (id, allowKeepRecovery = false) =>
    shell(() => {
      if (typeof allowKeepRecovery !== 'boolean') throw Error('Invalid close decision.');
      return leaveChoice(selected(id, true), allowKeepRecovery);
    })
  );
  async function closeRecords(records, quit) {
    if (closeTransaction) throw Error('A workbook close is already running.');
    closeTransaction = true;
    closing = true;
    try {
      const decisions = [];
      // Collect choices and Save filenames before changing any workbook. Cancel
      // in a later dialog leaves every accepted cursor, file and tab untouched.
      for (const record of records) {
        const choice = record.status.dirty && !record.suspended ? await leaveChoice(record) : 'keep';
        if (choice === 'cancel') return { closed: false, cancelled: true };
        decisions.push({ record, choice });
      }
      const destinations = new Set();
      for (const decision of decisions)
        if (decision.choice === 'save') {
          decision.request = await saveRequest(decision.record);
          if (!decision.request) return { closed: false, cancelled: true };
          if (decision.request.destination) {
            if (destinations.has(decision.request.destination))
              throw Error('Two workbooks selected the same Save destination. Choose separate filenames.');
            destinations.add(decision.request.destination);
          }
        }
      for (const decision of decisions)
        if (decision.choice === 'save') {
          const saved = await saveSelected(decision.record.id, decision.request);
          if (!saved || saved.operationCancelled) return { closed: false, cancelled: true };
        }
      for (const decision of decisions) if (decision.choice !== 'discard') await pinRecordRecovery(decision.record);
      for (const { record, choice } of decisions)
        if (choice === 'discard' && !record.suspended) {
          record.status = await record.owner.request({ operation: 'discard' });
          record.discardedRecovery = true;
          if (/^[a-f0-9]{64}$/.test(record.status.documentId)) await keptRecovery.remove(record.status.documentId);
        }
      for (const { record } of decisions) {
        await retire(record);
        registry.remove(record);
      }
      publishDocuments();
      if (quit) {
        await retire(pending);
        pending = null;
        permitClose = true;
        window.close();
      }
      return { closed: true, cancelled: false, ...snapshot() };
    } finally {
      closeTransaction = false;
      closing = false;
      publishDocuments();
    }
  }
  handle('close-tab', (id) => shell(() => closeRecords([selected(id, true)], false)));
  handle('close-all', () => shell(() => closeRecords(registry.records(), true)));
  handle('follow-hyperlink', (id, request) =>
    shell(async () => {
      if (
        !request ||
        typeof request !== 'object' ||
        Array.isArray(request) ||
        Object.keys(request).sort().join(',') !== 'cell,documentId,expectedHyperlinkRevision,sheet,state' ||
        typeof request.expectedHyperlinkRevision !== 'string' ||
        !/^[a-f0-9]{64}$/.test(request.expectedHyperlinkRevision)
      )
        throw Error('Invalid link handoff request.');
      const record = selected(id),
        epoch = record.epoch;
      const current = () => {
        if (
          !registry.contains(record) ||
          registry.active !== record ||
          selected(id) !== record ||
          record.epoch !== epoch ||
          request.documentId !== record.status.documentId ||
          request.state !== record.status.state
        )
          throw Error('The initiating link or workbook changed.');
      };
      current();
      const reply = await record.owner.request({
        operation: 'hyperlink',
        documentId: request.documentId,
        state: request.state,
        sheet: request.sheet,
        cell: request.cell,
      });
      current();
      const definition = reply.hyperlink;
      if (
        !definition?.canEdit ||
        !definition.link ||
        definition.hyperlinkRevision !== request.expectedHyperlinkRevision
      )
        throw Error('The link changed or is outside the supported profile.');
      if (definition.link.kind === 'internal')
        return { kind: 'internal', sheet: definition.link.sheet, cell: definition.link.cell };
      const address = new URL(definition.link.target);
      if (!['https:', 'http:'].includes(address.protocol) || address.username || address.password)
        throw Error('Only ordinary web addresses can be opened.');
      await electronShell.openExternal(definition.link.target);
      return { kind: 'url', opened: true };
    })
  );
  handle('clipboard-read', () => clipboard.readText());
  handle('clipboard-write', (text) => {
    if (typeof text !== 'string' || text.length > 16 * 1024 * 1024) throw Error('Clipboard text exceeds its limit.');
    return clipboard.writeText(text);
  });
  handle('close-result', (accepted) =>
    shell(async () => {
      if (accepted !== true) {
        closing = false;
        return;
      }
      return closeRecords(registry.records(), true);
    })
  );
  session.defaultSession.setPermissionCheckHandler(() => false);
  session.defaultSession.setPermissionRequestHandler((_contents, _permission, callback) => callback(false));
  session.defaultSession.setDevicePermissionHandler(() => false);
  const launchFile = process.argv
    .slice(1)
    .find((value) => !value.startsWith('-') && /\.(xlsx|xlsm|xls|xlsb|csv|tsv)$/i.test(value));
  window = new BrowserWindow({
    width: 1400,
    height: 920,
    minWidth: 720,
    minHeight: 500,
    title: 'XLSX Editor',
  });
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  window.webContents.on('will-navigate', (event) => event.preventDefault());
  // A native shell reports a lost host connection; close without asking.
  app.on('host-disconnected', () => {
    rendererReady = false;
    window.close();
  });
  window.on('close', (event) => {
    if (permitClose) return;
    event.preventDefault();
    if (!rendererReady) {
      if (!closing) {
        closing = true;
        // Launch-file preparation may still own a worker before current is
        // assigned. Keep its accepted binding before retiring either owner.
        void shell(async () => {
          await startupOpen;
          for (const record of registry.records()) await pinRecordRecovery(record);
          for (const record of registry.records()) await retire(record);
          await retire(pending);
          current = null;
          pending = null;
          permitClose = true;
          window.close();
        })
          .catch(async (error) => {
            closing = false;
            // A failed binding write must leave the accepted owner available.
            await dialog.showMessageBox(window, {
              type: 'error',
              buttons: ['Reopen editor'],
              defaultId: 0,
              cancelId: 0,
              message: 'Recovery could not be kept before closing.',
              detail: 'Accepted edits were retained. ' + String(error.message ?? error),
            });
            if (window.isDestroyed()) return;
            for (const record of registry.records())
              if (!record.owner.closed) record.status = await record.owner.request({ operation: 'status' });
            await window.loadURL(page);
          })
          .catch((error) => console.error('xlsx.exceptional-close', error));
      }
      return;
    }
    if (!closing) {
      closing = true;
      window.webContents.send('xlsx:close-request');
    }
  });
  window.webContents.on('render-process-gone', async () => {
    rendererReady = false;
    closing = false;
    const { response } = await dialog.showMessageBox(window, {
      type: 'error',
      buttons: ['Reopen editor', 'Close'],
      defaultId: 0,
      cancelId: 1,
      message: 'The spreadsheet view stopped unexpectedly.',
      detail: 'Accepted edits remain in the worker and recovery journal.',
    });
    if (response === 0) {
      for (const record of registry.records())
        if (!record.owner.closed) record.status = await record.owner.request({ operation: 'status' });
      window.reload();
    } else window.close();
  });
  Menu.setApplicationMenu(
    Menu.buildFromTemplate([
      ...(process.platform === 'darwin' ? [{ role: 'appMenu' }] : []),
      {
        label: 'File',
        submenu: [
          {
            label: 'New workbook',
            accelerator: 'CmdOrCtrl+N',
            click: () => window.webContents.send('xlsx:action', 'new'),
          },
          { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: () => window.webContents.send('xlsx:action', 'open') },
          { label: 'Save', accelerator: 'CmdOrCtrl+S', click: () => window.webContents.send('xlsx:action', 'save') },
          {
            label: 'Save a copy…',
            accelerator: 'CmdOrCtrl+Shift+S',
            click: () => window.webContents.send('xlsx:action', 'copy'),
          },
          { label: 'Reload file…', click: () => window.webContents.send('xlsx:action', 'reload') },
          {
            label: 'Close workbook',
            accelerator: 'CmdOrCtrl+W',
            click: () => window.webContents.send('xlsx:action', 'close-tab'),
          },
        ],
      },
      {
        label: 'Edit',
        submenu: [
          { label: 'Undo', accelerator: 'CmdOrCtrl+Z', click: () => window.webContents.send('xlsx:action', 'undo') },
          {
            label: 'Redo',
            accelerator: 'CmdOrCtrl+Shift+Z',
            click: () => window.webContents.send('xlsx:action', 'redo'),
          },
          { type: 'separator' },
          { role: 'cut' },
          { role: 'copy' },
          { role: 'paste' },
          { role: 'selectAll' },
        ],
      },
      { role: 'viewMenu' },
      { role: 'windowMenu' },
    ])
  );
  routeLaunch = (filename) =>
    shell(async () => {
      try {
        if (closing || pending) throw Error('Finish the close or open decision before opening another workbook.');
        const canonical = await realpath(path.resolve(filename)),
          existing = registry.byFilename(canonical);
        if (existing) {
          activateRecord(existing);
          window.show();
          window.focus();
          return;
        }
        registry.assertRoom();
        const candidate = await prepare(canonical, { prepared: true });
        try {
          candidate.status = await candidate.owner.request({ operation: 'activate' });
          registry.add(candidate);
          activateRecord(candidate);
          await remember(candidate);
        } catch (error) {
          await retire(candidate);
          throw error;
        }
      } catch (error) {
        await dialog.showMessageBox(window, {
          type: 'error',
          message: 'Spreadsheet could not be opened.',
          detail: String(error.message ?? error),
        });
      } finally {
        publishDocuments();
      }
    });
  app.on('window-all-closed', () => app.quit());
  if (launchFile) {
    startupOpen = (async () => {
      try {
        current = registry.add(await prepare(path.resolve(launchFile)));
        await remember(current);
      } catch (error) {
        await dialog.showMessageBox(window, {
          type: 'error',
          message: 'Spreadsheet could not be opened.',
          detail: String(error),
        });
      }
    })();
    await startupOpen;
  }
  if (closing || permitClose) return;
  try {
    await window.loadURL(page);
  } catch (error) {
    if (!permitClose) throw error;
  }
  for (const filename of launchQueue.splice(0)) await routeLaunch(filename);
}
const singleInstance = app.requestSingleInstanceLock();
if (!singleInstance) app.quit();
else {
  app.on('second-instance', (_event, args, workingDirectory) => {
    const window = BrowserWindow.getAllWindows()[0];
    if (window) {
      if (window.isMinimized()) window.restore();
      window.show();
      window.focus();
    }
    const file = args
      .slice(1)
      .find((value) => !value.startsWith('-') && /\.(xlsx|xlsm|xls|xlsb|csv|tsv)$/i.test(value));
    if (file) incomingFile(path.resolve(workingDirectory, file));
  });
  app.on('open-file', (event, filename) => {
    event.preventDefault();
    incomingFile(filename);
  });
  void app
    .whenReady()
    .then(boot)
    .catch((error) => {
      console.error(error);
      app.exit(1);
    });
}
