import { accessibilityFields, bindAccessibility } from '../core/accessibility.mjs';
import { notesFields, bindNotes } from '../core/notes.mjs';
import { hyperlinkFields, bindHyperlink } from '../core/hyperlinks.mjs';
import { randomUUID } from 'node:crypto';
import { mkdir, readFile, readdir, realpath, rename, rm, stat, lstat, chmod } from 'node:fs/promises';
import path from 'node:path';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { PassiveWorkbook } from './passive-workbook.mjs';
import { PASSIVE_FORMAT_LIMITS } from '../core/passive-container.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync, nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { indexedImplementationIdentity } from '../core/indexed-cache.mjs';
import { createWorkbook } from '../core/workbook-package.mjs';
import { cellPosition, rangeBounds } from '../core/workbook-commands.mjs';
import { ownerCommands, decodedCommands, encodeOwnerResult } from '../dist/core/transport.js';
import { validateDuplicatePreview } from '../core/duplicate-model.mjs';
import { formulaPreviewFields, bindFormulaPreview, formulaPreviewCancellation } from '../core/formula-preview.mjs';
import { usedRangeFields, bindUsedRange } from '../core/used-range.mjs';
import { noteMembershipFields, bindNoteMembership } from '../core/note-membership.mjs';
import { openCsv, applyCsv, csvView, readCsv, scanCsv, summarizeCsv, exportCsv } from './csv-session.mjs';
import { createPasteStager } from './large-paste.mjs';
import { normalizeOwnerRequest } from './owner-request.mjs';
import { validatePasteDescriptor, validatePastePlan, validatePasteState } from '../core/paste-protocol.mjs';
import { validateImportReceiptRequest, findImportReceipt } from './import-receipt.mjs';
import { composePrint, printDocumentHtml } from '../core/print-layout.mjs';
import {
  sha256,
  validHash,
  validId,
  revision,
  atomicJson,
  readJson,
  copySnapshot,
  writePrivateSnapshot,
  stageCopy,
  assertRevision,
  publishCopy,
  syncDirectory,
  directoryBytes,
  validateRecoveryManifest,
  validateSavedSource,
  destinationIdentity,
  listDraftRecords,
  isOoxmlExtension,
  RECOVERY_DRAFT_LIMIT,
  RECOVERY_MANIFEST_LIMIT,
} from './storage.mjs';
import { refusal } from '../core/refusal-codes.mjs';

const JOURNAL_LIMIT = 64 * 1024 * 1024,
  COMMAND_LIMIT = 2 * 1024 * 1024,
  NODE_LIMIT = 2 * COMMAND_LIMIT + 4096;
const key = (head) => head ?? 'base';
const filePath = (value) =>
  typeof value === 'string' && value.length <= 4096 && !value.includes('\0') && path.isAbsolute(value);
const objectHandle = (value) =>
  typeof value === 'string' && /^(?:[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}|opaque-[a-f0-9]{24})$/.test(value);
export { listSpreadsheetRecovery } from './storage.mjs';

/** Trusted Node owner, normally isolated in a worker. UI sees metadata/ranges only. */
export class SpreadsheetDocument {
  #directory;
  #filename;
  #manifest;
  #book;
  // Reads the macros of an editable macro workbook (see #macroReader).
  #macroBook;
  #frames = [];
  #cursor = 0;
  #queue = Promise.resolve();
  #closed = false;
  #recoveryRoot;
  #budget;
  #journalBytes = 0;
  #recovered = false;
  #native;
  #csvBaseline;
  #maintenanceWarning;
  #indexDirectory;
  #cacheIdentity;
  #automaticCache;
  #cacheBudget;
  #publicationWarning;
  #discarded = false;
  #ready = false;
  #created = false;
  #verifiedSnapshot;
  #checkedJournalKey;
  #checkedJournal;
  #onProgress;
  #beforeCritical;
  #leaveCritical;
  #checkCancelled;
  #lastPhase = '';
  #phaseCount = 0;
  #creatingRecoveryId;
  #paste;
  #pasteIdentity;
  #importParent;
  #applyingImport = false;
  #importPercent = -1;
  #cleanupNative = false;
  #printJob;
  #initialCalculationPolicy;
  #saveRecalculateRevision;
  #formulaPreviews = new Map();
  #formulaPreviewIds = new Set();
  constructor({
    recoveryDirectory,
    maxRecoveryBytes = 512 * 1024 * 1024,
    maxIndexCacheBytes = 2 * 1024 ** 3,
    native = nativeWorkbookOperationAsync,
    indexDirectory,
    cacheIdentity,
    initialCalculationPolicy,
    onProgress,
    beforeCritical,
    leaveCritical,
    checkCancelled,
  }) {
    if (!Number.isSafeInteger(maxRecoveryBytes) || maxRecoveryBytes < 1024 * 1024)
      throw refusal('internal-invalid-field', 'Invalid recovery budget.');
    if (!Number.isSafeInteger(maxIndexCacheBytes) || maxIndexCacheBytes < 4096 || maxIndexCacheBytes > 2 * 1024 ** 3)
      throw refusal('session-malformed', 'Invalid workbook index cache budget.');
    if (
      initialCalculationPolicy !== undefined &&
      (!initialCalculationPolicy ||
        typeof initialCalculationPolicy !== 'object' ||
        Array.isArray(initialCalculationPolicy) ||
        Object.keys(initialCalculationPolicy).sort().join(',') !== 'mode,recalculateBeforeSave' ||
        !['auto', 'manual'].includes(initialCalculationPolicy.mode) ||
        typeof initialCalculationPolicy.recalculateBeforeSave !== 'boolean')
    )
      throw refusal('internal-invalid-field', 'Invalid initial calculation policy');
    this.#initialCalculationPolicy =
      initialCalculationPolicy === undefined ? undefined : structuredClone(initialCalculationPolicy);
    this.#recoveryRoot = path.resolve(recoveryDirectory);
    this.#budget = maxRecoveryBytes;
    this.#indexDirectory = indexDirectory;
    this.#onProgress = onProgress;
    this.#beforeCritical = beforeCritical;
    this.#leaveCritical = leaveCritical;
    this.#checkCancelled = checkCancelled;
    this.#native = async (input) => {
      if (!(this.#cleanupNative && input.operation === 'compact')) this.#checkCancelled?.();
      const phase = {
        index: 'indexing',
        calculate: 'calculating',
        export: 'exporting',
        'copy-source': 'copying-source',
        patch: 'applying',
        bulk: 'applying',
        format: 'applying',
        transform: 'applying',
        'package-edit': 'applying',
        compact: 'cleaning-up',
      }[input.operation];
      if (phase && !(this.#applyingImport && input.operation === 'patch')) this.#emit(phase, phase !== 'cleaning-up');
      return native(input);
    };
    this.#cacheIdentity = cacheIdentity;
    this.#automaticCache = cacheIdentity === undefined && native === nativeWorkbookOperationAsync;
    this.#cacheBudget = maxIndexCacheBytes;
  }
  #emit(phase, cancellable = true) {
    const recoveryId = this.#manifest?.documentId ?? this.#creatingRecoveryId,
      key = phase + ':' + cancellable + ':' + (recoveryId ?? '');
    if (key === this.#lastPhase) return;
    this.#lastPhase = key;
    this.#phaseCount++;
    try {
      this.#onProgress?.({ phase, cancellable, ...(recoveryId ? { recoveryId } : {}) });
    } catch {}
  }
  #critical(phase) {
    this.#beforeCritical?.(phase);
    this.#emit(phase, false);
  }
  #releaseCritical() {
    this.#leaveCritical?.();
  }
  #history(frames = this.#frames, cursor = this.#cursor) {
    return { cursor, frames: frames.map(({ head, hash }) => ({ head, hash })) };
  }
  #tip(manifest) {
    const tip = manifest.history?.frames.at(-1);
    return tip ? { ...manifest, head: tip.head, headHash: tip.hash } : manifest;
  }
  #documentId(manifest = this.#manifest) {
    return manifest.documentId ?? sha256(this.#filename);
  }
  async #pasteOwner(manifest = this.#manifest) {
    const documentId = this.#documentId(manifest);
    if (this.#paste) {
      if (this.#pasteIdentity !== documentId)
        throw refusal('session-stale', 'Paste storage document identity changed.');
      return this.#paste;
    }
    const directoryInfo = await lstat(this.#directory),
      recoveryRoot = await realpath(this.#recoveryRoot),
      ownerRoot = await realpath(this.#directory);
    if (!directoryInfo.isDirectory() || directoryInfo.isSymbolicLink() || path.dirname(ownerRoot) !== recoveryRoot)
      throw refusal('session-bad-input', 'Paste recovery owner directory is not confined.');
    const directory = path.join(ownerRoot, 'paste-v1');
    this.#paste = createPasteStager({
      directory,
      documentId,
      checkCancelled: () => this.#checkCancelled?.(),
      accountBytes: async ({ delta }) => {
        if (
          delta > 0 &&
          (this.#journalBytes + (this.#paste?.stats().attachmentBytes ?? 0) + delta + RECOVERY_MANIFEST_LIMIT >
            JOURNAL_LIMIT ||
            (await directoryBytes(this.#directory)) + delta + RECOVERY_MANIFEST_LIMIT > this.#budget)
        )
          throw refusal(
            'session-limit',
            'Recovery attachment disk budget exceeded. Save and reopen before continuing.'
          );
      },
    });
    this.#pasteIdentity = documentId;
    return this.#paste;
  }
  #boundImport(request, current = true) {
    if (!this.#book || !isOoxmlExtension(this.#manifest.extension))
      throw refusal('session-bad-input', 'Staged imports require an XLSX workbook.');
    if (request.documentId !== this.#documentId())
      throw refusal('session-stale', 'Paste request belongs to another document.');
    validatePasteState(request.state);
    if (current && request.state !== key(this.#current().head))
      throw refusal('session-stale', 'Spreadsheet import is stale.');
  }
  #printFrame(request) {
    if (!this.#book || !isOoxmlExtension(this.#manifest.extension) || request.documentId !== this.#documentId())
      throw refusal('session-stale', 'Print request belongs to another XLSX document.');
    if (request.state !== key(this.#current().head))
      throw refusal('session-unsupported', 'Print preview belongs to an earlier workbook revision. Prepare it again.');
    return this.#current();
  }
  #print(request) {
    const frame = this.#printFrame(request),
      job = this.#printJob;
    if (!job || job.id !== request.jobId || job.state !== request.state || job.revision !== frame.view.revision)
      throw refusal('session-stale', 'Print preview has expired. Prepare it again.');
    return job;
  }
  #importDescriptors(commands) {
    const imports = commands.filter((command) => command?.type === 'importCells');
    if (imports.length && (imports.length !== 1 || commands.length !== 1))
      throw refusal('session-bad-input', 'A staged import must be one complete owner command.');
    return imports.map(validatePasteDescriptor);
  }
  async #verifyImports(commands, parent, manifest = this.#manifest) {
    const imports = this.#importDescriptors(commands);
    if (!imports.length) return [];
    if (!isOoxmlExtension(manifest.extension))
      throw refusal('session-bad-input', 'Staged imports require an XLSX workbook.');
    const owner = await this.#pasteOwner(manifest);
    for (const descriptor of imports)
      for await (const _page of owner.pages(descriptor, { state: key(parent) })) this.#checkCancelled?.();
    return imports;
  }
  #importProgress({ completed, total }) {
    if (
      !this.#applyingImport ||
      !Number.isInteger(completed) ||
      !Number.isInteger(total) ||
      total < 1 ||
      completed < 0 ||
      completed > total
    )
      return;
    const percent = Math.floor((completed * 100) / total);
    if (percent === this.#importPercent) return;
    this.#importPercent = percent;
    this.#emit('applying-paste ' + percent + '%');
  }
  async open(filename, { fresh = false, prepared = false } = {}) {
    this.#emit('checking-source');
    if (this.#directory) throw refusal('session-refused', 'This document owner has already been opened.');
    const extension = path.extname(filename).toLowerCase();
    if (!['.xlsx', '.xlsm', '.xls', '.xlsb', '.csv', '.tsv'].includes(extension))
      throw refusal('session-bad-input', 'Select an XLSX, XLSM, XLS, XLSB, CSV or TSV file.');
    const canonical = await realpath(filename),
      info = await stat(canonical);
    if (!info.isFile() || info.size > this.#budget)
      throw refusal('session-limit', 'Source exceeds the recovery disk budget.');
    if (['.xls', '.xlsb'].includes(extension) && info.size > PASSIVE_FORMAT_LIMITS.sourceBytes)
      throw refusal('session-limit', 'Passive workbook source exceeds its 64 MiB byte budget.');
    if (fresh) {
      // A same-file reload stages an independent snapshot, never a second writer
      // over the live owner's recovery manifest. Only activation makes it mutable.
      return this.#newDraft(path.basename(canonical).slice(0, 80), canonical, prepared);
    }
    const { records } = await listDraftRecords(this.#recoveryRoot),
      matches = [];
    for (const record of records) {
      const m = record.manifest;
      if (m.prepared) continue;
      if (
        m.filename === canonical ||
        (m.pendingPublication?.kind === 'bind' &&
          m.pendingPublication.destination === canonical &&
          (await revision(canonical)) === m.pendingPublication.revision)
      )
        matches.push(record);
    }
    if (matches.length > 1)
      throw refusal(
        'session-unsupported',
        'Conflicting workbook recovery identities exist for this file. Existing data was preserved; recover a specific draft first.'
      );
    this.#filename = canonical;
    this.#directory = matches[0]?.directory ?? path.join(this.#recoveryRoot, sha256(canonical));
    await mkdir(this.#directory, { recursive: true, mode: 0o700 });
    let saved;
    try {
      saved = validateRecoveryManifest(await readJson(this.#manifestPath()), matches[0]?.id);
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
    if (saved && saved.extension !== extension)
      throw refusal('internal-invalid-field', 'Invalid spreadsheet recovery format.');
    if (
      saved &&
      (saved.pendingPublication || saved.head !== saved.savedHead || this.#tip(saved).head !== saved.savedHead)
    ) {
      await this.#verifySnapshot(saved);
      await this.#journalChain(this.#tip(saved));
      await this.#reconcile(saved);
    }
    if (saved && (saved.head !== saved.savedHead || this.#tip(saved).head !== saved.savedHead)) {
      this.#manifest = saved;
      this.#recovered = true;
      this.#filename = saved.schema === 2 ? saved.filename : canonical;
    } else {
      const snapshot = randomUUID(),
        baseRevision = await copySnapshot(canonical, path.join(this.#directory, snapshot + extension));
      this.#manifest = {
        ...(saved?.schema === 2
          ? {
              schema: 2,
              documentId: saved.documentId,
              name: saved.name,
              filename: canonical,
              createdAt: saved.createdAt,
              prepared: false,
            }
          : { schema: 1 }),
        extension,
        snapshot,
        baseRevision,
        sourceRevision: baseRevision,
        head: null,
        headHash: null,
        savedHead: null,
        ...(this.#initialCalculationPolicy ? { initialCalculationPolicy: this.#initialCalculationPolicy } : {}),
      };
      this.#critical('saving-recovery');
      try {
        await atomicJson(this.#manifestPath(), this.#manifest);
      } finally {
        this.#releaseCritical();
      }
    }
    return this.#hydrate();
  }
  async create(name = 'Untitled workbook', { prepared = false } = {}) {
    this.#emit('creating-workbook');
    if (this.#directory) throw refusal('session-refused', 'This document owner has already been opened.');
    if (typeof name !== 'string' || !name.trim() || name.length > 80)
      throw refusal('internal-invalid-field', 'Invalid untitled workbook name.');
    return this.#newDraft(name, null, prepared);
  }
  async #newDraft(name, filename, prepared) {
    const { count } = await listDraftRecords(this.#recoveryRoot);
    if (count >= RECOVERY_DRAFT_LIMIT)
      throw refusal(
        'session-limit',
        'Workbook recovery has reached its 128-document limit. Save or discard existing drafts first.'
      );
    const documentId = randomUUID(),
      snapshot = randomUUID(),
      extension = filename ? path.extname(filename).toLowerCase() : '.xlsx';
    this.#filename = filename;
    this.#directory = path.join(this.#recoveryRoot, 'draft-' + documentId);
    this.#creatingRecoveryId = documentId;
    this.#emit('copying-source');
    await mkdir(this.#directory, { recursive: false, mode: 0o700 }).catch(async (error) => {
      if (error.code !== 'ENOENT') throw error;
      await mkdir(this.#recoveryRoot, { recursive: true, mode: 0o700 });
      await mkdir(this.#directory, { mode: 0o700 });
    });
    this.#created = true;
    this.#emit('copying-source');
    let baseRevision;
    if (filename) baseRevision = await copySnapshot(filename, path.join(this.#directory, snapshot + extension));
    else {
      const bytes = createWorkbook();
      await writePrivateSnapshot(path.join(this.#directory, snapshot + extension), bytes);
      baseRevision = sha256(bytes);
    }
    this.#manifest = {
      schema: 2,
      documentId,
      name,
      filename,
      createdAt: Date.now(),
      prepared,
      extension,
      snapshot,
      baseRevision,
      sourceRevision: filename ? baseRevision : null,
      head: null,
      headHash: null,
      savedHead: null,
      ...(this.#initialCalculationPolicy ? { initialCalculationPolicy: this.#initialCalculationPolicy } : {}),
    };
    this.#critical('saving-recovery');
    try {
      await atomicJson(this.#manifestPath(), this.#manifest);
    } finally {
      this.#releaseCritical();
    }
    return this.#hydrate();
  }
  async recover(documentId) {
    this.#emit('checking-recovery');
    if (this.#directory || !validId(documentId))
      throw refusal('internal-invalid-field', 'Invalid workbook recovery handle.');
    this.#directory = path.join(this.#recoveryRoot, 'draft-' + documentId);
    if (!(await lstat(this.#directory)).isDirectory())
      throw refusal('internal-invalid-field', 'Invalid workbook recovery directory.');
    this.#manifest = validateRecoveryManifest(
      await readJson(this.#manifestPath(), RECOVERY_MANIFEST_LIMIT),
      documentId
    );
    if (this.#manifest.schema !== 2 || this.#manifest.prepared)
      throw refusal('session-refused', 'This prepared workbook was never activated.');
    await this.#verifySnapshot(this.#manifest);
    await this.#journalChain(this.#tip(this.#manifest));
    await this.#reconcile(this.#manifest);
    this.#filename = this.#manifest.filename;
    this.#recovered = true;
    return this.#hydrate();
  }
  async resume(filename) {
    if (this.#directory || !filePath(filename))
      throw refusal('internal-invalid-field', 'Invalid saved workbook recovery binding.');
    this.#filename = filename;
    this.#directory = path.join(this.#recoveryRoot, sha256(filename));
    this.#manifest = validateRecoveryManifest(await readJson(this.#manifestPath(), RECOVERY_MANIFEST_LIMIT));
    if (this.#manifest.schema !== 1) throw refusal('session-malformed', 'Unexpected saved workbook recovery identity.');
    await this.#verifySnapshot(this.#manifest);
    await this.#journalChain(this.#tip(this.#manifest));
    await this.#reconcile(this.#manifest);
    this.#recovered = true;
    return this.#hydrate();
  }
  async #reconcile(saved) {
    if (!saved.pendingPublication) return;
    const pending = saved.pendingPublication,
      target = pending.kind === 'bind' ? pending.destination : (this.#filename ?? saved.filename);
    let actual;
    try {
      actual = await revision(target);
    } catch (error) {
      if (!['ENOENT', 'ENOTDIR'].includes(error.code)) throw error;
    }
    // Historical intents predate staged-file identity. New intents require it
    // so byte-identical prior content cannot masquerade as a completed Save.
    const confirmed =
      actual === pending.revision &&
      (!pending.destinationIdentity || (await this.#destinationMatches(target, pending.destinationIdentity)));
    if (confirmed) {
      saved.sourceRevision = pending.revision;
      saved.savedHead = pending.head;
      if (pending.kind === 'bind') saved.filename = target;
      if (pending.savedSource) {
        await this.#verifySavedSource(pending.savedSource);
        const chain = await this.#journalChain(this.#tip(saved)),
          ordered = [
            { head: null, hash: null },
            ...chain
              .slice()
              .reverse()
              .map((item) => ({ head: item.head, hash: item.hash })),
          ];
        saved.savedSources = this.#selectSavedSources(
          [...(saved.savedSources ?? []).filter((source) => source.head !== pending.head), pending.savedSource],
          saved.history?.frames ?? ordered.slice(-101),
          ordered
        );
      }
    }
    delete saved.pendingPublication;
    this.#critical('saving-recovery');
    try {
      await atomicJson(this.#manifestPath(), saved);
    } finally {
      this.#releaseCritical();
    }
  }
  async #verifySnapshot(manifest) {
    const identity = manifest.snapshot + ':' + manifest.baseRevision;
    if (this.#verifiedSnapshot === identity) return;
    const filename = path.join(this.#directory, manifest.snapshot + manifest.extension),
      info = await lstat(filename);
    if (!info.isFile() || info.size > this.#budget)
      throw refusal('internal-invalid-field', 'Invalid recovery source snapshot.');
    if (['.xls', '.xlsb'].includes(manifest.extension) && info.size > PASSIVE_FORMAT_LIMITS.sourceBytes)
      throw refusal('session-limit', 'Passive recovery snapshot exceeds its 64 MiB byte budget.');
    if ((await revision(filename)) !== manifest.baseRevision)
      throw refusal('session-malformed', 'Recovery source snapshot is damaged. The original file is unchanged.');
    this.#verifiedSnapshot = identity;
  }
  #savedSourcePath(source) {
    return path.join(this.#directory, source.snapshot + '.xlsx');
  }
  #savedCalculationPath(calculation) {
    return path.join(this.#directory, calculation.snapshot + '.calculation.sqlite');
  }
  async #verifySavedCalculation(source) {
    validateSavedSource(source);
    if (!source.calculation) return undefined;
    const item = source.calculation,
      file = this.#savedCalculationPath(item),
      info = await lstat(file);
    if (
      !info.isFile() ||
      info.isSymbolicLink() ||
      info.nlink !== 1 ||
      info.size !== item.bytes ||
      info.size > this.#budget ||
      (await revision(file)) !== item.revision
    )
      throw refusal(
        'session-malformed',
        'Saved-source calculation checkpoint is damaged. Accepted recovery was preserved.'
      );
    return { path: file, revision: item.revision, bytes: item.bytes };
  }
  async #destinationMatches(target, identity) {
    const actual = await destinationIdentity(target).catch(() => null);
    return identity.ino !== '0' && !!actual && actual.dev === identity.dev && actual.ino === identity.ino;
  }
  async #verifySavedSource(source) {
    validateSavedSource(source);
    const file = this.#savedSourcePath(source),
      info = await lstat(file);
    if (
      !info.isFile() ||
      info.isSymbolicLink() ||
      info.nlink !== 1 ||
      info.size !== source.bytes ||
      info.size > this.#budget ||
      (await revision(file)) !== source.revision
    )
      throw refusal(
        'session-malformed',
        'Saved-source recovery checkpoint is damaged. Accepted recovery was preserved.'
      );
    await this.#verifySavedCalculation(source);
    return file;
  }
  #selectSavedSources(sources, frames, ordered) {
    const positions = new Map(ordered.map((frame, index) => [key(frame.head), { ...frame, index }])),
      retained = new Set(frames.map((frame) => key(frame.head))),
      first = positions.get(key(frames[0].head));
    if (!first) throw refusal('session-unsupported', 'Saved-source recovery history is unavailable.');
    const valid = sources.filter((source) => {
      const original = positions.get(key(source.head));
      return original && original.hash === source.headHash;
    });
    const preceding = valid
      .filter((source) => positions.get(key(source.head)).index < first.index)
      .sort((a, b) => positions.get(key(b.head)).index - positions.get(key(a.head)).index)[0];
    return valid.filter((source) => retained.has(key(source.head)) || source === preceding);
  }
  async #prepareSavedSource(output, savedRevision) {
    if (!(this.#book instanceof IndexedWorkbook) || !isOoxmlExtension(this.#manifest.extension)) return null;
    // A new branch can retire redo checkpoints before the next maintenance
    // pass. Reclaim verified orphan payloads before charging a new Save's quota.
    this.#checkCancelled?.();
    await this.#prune();
    const frame = this.#current(),
      book = this.#book,
      info = await lstat(output);
    if (!info.isFile() || info.isSymbolicLink() || info.size < 1 || info.size > this.#budget)
      throw refusal('session-limit', 'Saved-source checkpoint exceeds its recovery budget.');
    let snapshot,
      created = false,
      calculationCreated = false;
    const existing = (this.#manifest.savedSources ?? []).find(
      (source) => source.revision === savedRevision && source.bytes === info.size
    );
    if (existing) {
      await this.#verifySavedSource(existing);
      snapshot = existing.snapshot;
    } else if (
      savedRevision === this.#manifest.baseRevision &&
      (await lstat(this.#snapshotPath())).size === info.size
    ) {
      await this.#verifySnapshot(this.#manifest);
      snapshot = this.#manifest.snapshot;
    } else {
      if ((await directoryBytes(this.#directory)) + info.size + RECOVERY_MANIFEST_LIMIT > this.#budget)
        throw refusal('session-limit', 'Saved-source checkpoint exceeds the recovery disk budget.');
      snapshot = randomUUID();
      const file = path.join(this.#directory, snapshot + '.xlsx');
      try {
        if ((await copySnapshot(output, file)) !== savedRevision)
          throw refusal('session-stale', 'Saved-source checkpoint revision changed.');
        created = true;
      } catch (error) {
        await rm(file, { force: true });
        throw error;
      }
    }
    const checkpoint = {
      schema: 1,
      head: frame.head,
      headHash: frame.hash,
      snapshot,
      revision: savedRevision,
      bytes: info.size,
    };
    let token;
    try {
      token = await book.preparePublicationRebase(frame.state, this.#savedSourcePath(checkpoint), savedRevision);
      if (this.#current() !== frame || this.#book !== book)
        throw refusal('session-stale', 'Saved-source owner/head changed during preparation.');
      const pin = book.publicationCheckpoint(token);
      const prior = (this.#manifest.savedSources ?? []).find(
        (source) => source.calculation?.revision === pin.revision && source.calculation.bytes === pin.bytes
      );
      if (prior) {
        await this.#verifySavedCalculation(prior);
        checkpoint.calculation = { ...prior.calculation };
      } else {
        if ((await directoryBytes(this.#directory)) + pin.bytes + RECOVERY_MANIFEST_LIMIT > this.#budget)
          throw refusal('session-limit', 'Saved-source calculation checkpoint exceeds the recovery disk budget.');
        checkpoint.calculation = { schema: 1, snapshot: randomUUID(), revision: pin.revision, bytes: pin.bytes };
        const file = this.#savedCalculationPath(checkpoint.calculation);
        try {
          if ((await copySnapshot(pin.path, file)) !== pin.revision)
            throw refusal('session-stale', 'Saved-source calculation checkpoint revision changed.');
          await chmod(file, 0o600);
          calculationCreated = true;
        } catch (error) {
          await rm(file, { force: true });
          throw error;
        }
      }
      await this.#verifySavedCalculation(checkpoint);
      const chain = await this.#journalChain(this.#tip(this.#manifest)),
        ordered = [
          { head: null, hash: null },
          ...chain
            .slice()
            .reverse()
            .map((item) => ({ head: item.head, hash: item.hash })),
        ];
      const savedSources = this.#selectSavedSources(
        [...(this.#manifest.savedSources ?? []).filter((source) => source.head !== frame.head), checkpoint],
        this.#frames,
        ordered
      );
      if (savedSources.length > 102)
        throw refusal('session-limit', 'Saved-source checkpoint inventory exceeds its history bound.');
      return {
        frame,
        book,
        token,
        checkpoint,
        savedSources,
        created,
        calculationCreated,
        intent: false,
        adopted: false,
      };
    } catch (error) {
      if (token) await book.discardPublicationRebase(token);
      if (calculationCreated) await rm(this.#savedCalculationPath(checkpoint.calculation), { force: true });
      if (created) await rm(this.#savedSourcePath(checkpoint), { force: true });
      throw error;
    }
  }
  async #finishSavedSource(prepared, next, publicationWarning) {
    if (prepared) {
      if (
        this.#current() !== prepared.frame ||
        this.#book !== prepared.book ||
        next.head !== prepared.frame.head ||
        next.headHash !== prepared.frame.hash ||
        next.savedHead !== prepared.frame.head ||
        next.sourceRevision !== prepared.checkpoint.revision
      )
        throw refusal('session-stale', 'Published saved-source owner/head changed.');
      const result = prepared.book.adoptPublicationRebase(prepared.token);
      this.#frames[this.#cursor] = { ...prepared.frame, state: result.view.storage.state, view: result.view };
      prepared.adopted = true;
      next = { ...next, savedSources: prepared.savedSources };
      this.#printJob = undefined;
    }
    this.#manifest = next;
    try {
      await atomicJson(this.#manifestPath(), next);
      this.#publicationWarning = publicationWarning;
      await this.#maintain({ prune: true });
    } catch (error) {
      this.#publicationWarning =
        'The workbook was saved, but its recovery receipt could not be finalized. Reopen to reconcile it. ' +
        String(error.message ?? error);
    }
    return this.status();
  }
  async #discardSavedSource(prepared) {
    if (!prepared || prepared.adopted) return;
    await prepared.book.discardPublicationRebase(prepared.token);
    // A durable intent may describe a completed publication after a lost reply.
    if (prepared.created && !prepared.intent) await rm(this.#savedSourcePath(prepared.checkpoint), { force: true });
    if (prepared.calculationCreated && !prepared.intent)
      await rm(this.#savedCalculationPath(prepared.checkpoint.calculation), { force: true });
  }
  #assertSavedSource(prepared) {
    if (!prepared) return;
    if (this.#current() !== prepared.frame || this.#book !== prepared.book)
      throw refusal('session-stale', 'Saved-source owner/head changed before publication.');
    this.#checkCancelled?.();
    prepared.book.assertPublicationRebase(prepared.token);
  }
  async #restoreSavedSource(frame, checkpoint) {
    if (checkpoint.head !== frame.head || checkpoint.headHash !== frame.hash)
      throw refusal('session-stale', 'Saved-source checkpoint belongs to another history frame.');
    const source = await this.#verifySavedSource(checkpoint),
      calculationCheckpoint = await this.#verifySavedCalculation(checkpoint),
      book = this.#book;
    if (!calculationCheckpoint) {
      const exported = await book.export(frame.state, { publication: true });
      if (exported.revision !== checkpoint.revision)
        throw refusal(
          'session-unsupported',
          'Legacy saved-source calculation state cannot be reproduced from accepted recovery.'
        );
    }
    const token = await book.preparePublicationRebase(frame.state, source, checkpoint.revision, {
      ...(calculationCheckpoint ? { calculationCheckpoint } : {}),
    });
    try {
      const result = book.adoptPublicationRebase(token);
      return { ...frame, state: result.view.storage.state, view: result.view };
    } finally {
      await book.discardPublicationRebase(token);
    }
  }
  async #journalChain(manifest) {
    const identity = key(manifest.head) + ':' + manifest.headHash;
    if (this.#checkedJournalKey === identity) return this.#checkedJournal;
    this.#journalBytes = 0;
    const chain = [],
      attachments = new Map();
    let head = manifest.head,
      expectedHash = manifest.headHash,
      attachmentBytes = 0;
    const visited = new Set();
    while (head !== null) {
      if (!validId(head) || visited.has(head))
        throw refusal('internal-invalid-field', 'Invalid recovery command chain.');
      visited.add(head);
      const filename = this.#nodePath(head);
      const info = await lstat(filename);
      if (!info.isFile() || info.size > NODE_LIMIT)
        throw refusal('session-limit', 'Recovery command exceeds its budget or is not a private regular file.');
      const bytes = await readFile(filename);
      this.#journalBytes += bytes.length;
      if (
        bytes.length > NODE_LIMIT ||
        this.#journalBytes + attachmentBytes + RECOVERY_MANIFEST_LIMIT > JOURNAL_LIMIT ||
        sha256(bytes) !== expectedHash
      )
        throw refusal('session-malformed', 'Recovery commands are damaged or exceed their budget.');
      const node = JSON.parse(bytes);
      if (
        node.schema !== 1 ||
        typeof node.commands !== 'string' ||
        (node.parent !== null && (!validId(node.parent) || !validHash(node.parentHash)))
      )
        throw refusal('internal-invalid-field', 'Invalid recovery command record.');
      const commands = decodedCommands({ encodedCommands: node.commands });
      if (!Array.isArray(commands) || !commands.length || commands.length > 10000)
        throw refusal('internal-invalid-field', 'Invalid recovery command batch.');
      for (const descriptor of await this.#verifyImports(commands, node.parent, manifest)) {
        const prior = attachments.get(descriptor.id);
        if (prior && JSON.stringify(prior) !== JSON.stringify(descriptor))
          throw refusal('session-stale', 'Recovery import attachment identity changed.');
        if (!prior) {
          attachments.set(descriptor.id, descriptor);
          attachmentBytes += descriptor.bytes;
        }
        if (this.#journalBytes + attachmentBytes + RECOVERY_MANIFEST_LIMIT > JOURNAL_LIMIT)
          throw refusal('session-limit', 'Recovery command attachments exceed their journal budget.');
      }
      chain.push({ head, hash: expectedHash, node });
      head = node.parent;
      expectedHash = node.parentHash;
    }
    this.#checkedJournalKey = identity;
    this.#checkedJournal = chain;
    return chain;
  }
  async #hydrate() {
    this.#emit('checking-recovery');
    const extension = this.#manifest.extension;
    await this.#verifySnapshot(this.#manifest);
    const chain = await this.#journalChain(this.#tip(this.#manifest));
    const ordered = [
        { head: null, hash: null },
        ...chain
          .slice()
          .reverse()
          .map((item) => ({ head: item.head, hash: item.hash })),
      ],
      positions = new Map(ordered.map((frame, index) => [key(frame.head), { ...frame, index }]));
    const expectedFrames = this.#manifest.history?.frames ?? ordered.slice(-101),
      first = positions.get(key(expectedFrames[0].head));
    const savedSources = new Map();
    for (const checkpoint of this.#manifest.savedSources ?? []) {
      const original = positions.get(key(checkpoint.head));
      if (!original || original.hash !== checkpoint.headHash)
        throw refusal('session-unsupported', 'Saved-source checkpoint is outside its verified recovery chain.');
      savedSources.set(key(checkpoint.head), checkpoint);
    }
    const savedStart = [...savedSources.values()]
      .filter((source) => positions.get(key(source.head)).index <= first.index)
      .sort((a, b) => positions.get(key(b.head)).index - positions.get(key(a.head)).index)[0];
    let startIndex = 0;
    if (isOoxmlExtension(extension)) {
      if (this.#automaticCache) {
        try {
          this.#cacheIdentity = await indexedImplementationIdentity(nativeWorkbookBinary());
        } catch {
          this.#cacheIdentity = undefined;
        }
      }
      this.#book = new IndexedWorkbook(
        this.#native,
        async () => {
          throw refusal('session-unsupported', 'This transformation has no qualified native implementation.');
        },
        (input) => calculateIndexedWorkbook(input, this.#native),
        {
          importPages: (descriptor) => {
            if (!this.#paste || this.#importParent === undefined)
              throw refusal('session-bad-input', 'Staged import has no trusted Document parent.');
            return this.#paste.pages(descriptor, { state: this.#importParent });
          },
          importProgress: (event) => this.#importProgress(event),
        }
      );
      const source = savedStart ? await this.#verifySavedSource(savedStart) : this.#snapshotPath();
      const savedCalculationCheckpoint = savedStart ? await this.#verifySavedCalculation(savedStart) : undefined;
      const view = await this.#book.open(source, {
        storageDirectory: this.#indexDirectory,
        cacheIdentity: this.#cacheIdentity,
        ...(savedCalculationCheckpoint ? { savedCalculationCheckpoint } : {}),
        cacheOptions: { directory: path.join(this.#recoveryRoot, '.native-index-cache'), maxBytes: this.#cacheBudget },
        ...(!savedStart && this.#manifest.initialCalculationPolicy
          ? { initialCalculationPolicy: this.#manifest.initialCalculationPolicy }
          : {}),
      });
      this.#frames.push({
        head: savedStart?.head ?? null,
        hash: savedStart?.headHash ?? null,
        state: view.storage.state,
        view,
      });
      if (savedStart) startIndex = positions.get(key(savedStart.head)).index;
    } else if (['.xls', '.xlsb'].includes(extension)) {
      if (chain.length) throw refusal('session-unsupported', 'Passive workbook recovery cannot contain edit commands.');
      this.#book = new PassiveWorkbook();
      this.#emit('reading-stored-workbook');
      const view = await this.#book.open(this.#snapshotPath(), {
        format: extension.slice(1),
        storageDirectory: this.#indexDirectory,
      });
      this.#checkCancelled?.();
      this.#frames.push({ head: null, hash: null, state: '0', view });
    } else {
      const index = await openCsv(
        this.#snapshotPath(),
        extension === '.tsv' ? '\t' : 'auto',
        this.#manifest.baseRevision
      );
      this.#csvBaseline = { rows: index.rows, storedCells: index.storedCells };
      this.#frames.push({ head: null, hash: null, index });
    }
    for (const record of chain.slice().reverse().slice(startIndex)) {
      const commands = decodedCommands({ encodedCommands: record.node.commands });
      const frame = await this.#applyFrame(this.#current(), commands);
      delete frame.commands;
      let accepted = { ...frame, head: record.head, hash: record.hash };
      const checkpoint = savedSources.get(key(record.head));
      if (checkpoint) accepted = await this.#restoreSavedSource(accepted, checkpoint);
      this.#frames.push(accepted);
      if (this.#frames.length > 101) this.#frames.shift();
      this.#cursor = this.#frames.length - 1;
    }
    if (this.#manifest.history) {
      const history = this.#manifest.history;
      if (JSON.stringify(this.#history(this.#frames, history.cursor)) !== JSON.stringify(history))
        throw refusal('session-bad-input', 'Recovery undo frames do not match the verified command chain.');
      this.#cursor = history.cursor;
    }
    this.#critical('cleaning-up');
    try {
      await this.#prune();
      await this.#maintain();
    } finally {
      this.#releaseCritical();
    }
    this.#ready = true;
    this.#checkedJournal = undefined;
    this.#checkedJournalKey = undefined;
    this.#emit('ready', false);
    return this.status();
  }
  /** Macros are carried untouched and can be read, never run. The passive reader
   * that serves old-format workbooks also reads the macros of an editable macro
   * workbook, from its stored file, when the first request for them arrives. */
  async #macroReader() {
    if (this.#book instanceof PassiveWorkbook) return this.#book;
    if (!this.#current().view?.macroProject) return undefined;
    if (!this.#macroBook) {
      const reader = new PassiveWorkbook();
      try {
        await reader.open(this.#snapshotPath(), { format: 'xlsm', storageDirectory: this.#indexDirectory });
      } catch (error) {
        await reader.close().catch(() => {});
        throw error;
      }
      this.#macroBook = reader;
    }
    return this.#macroBook;
  }
  #manifestPath() {
    return path.join(this.#directory, 'current.json');
  }
  #snapshotPath() {
    return path.join(this.#directory, this.#manifest.snapshot + this.#manifest.extension);
  }
  #nodePath(head) {
    if (!validId(head)) throw refusal('session-malformed', 'Invalid recovery record identity.');
    return path.join(this.#directory, head + '.json');
  }
  #current() {
    if (this.#closed) throw refusal('session-closed', 'Document owner is closed.');
    return this.#frames[this.#cursor];
  }
  #frame(state) {
    const found = this.#frames.find((frame) => key(frame.head) === state);
    if (!found) throw refusal('session-stale', 'Spreadsheet view has expired.');
    return found;
  }
  status() {
    const frame = this.#current();
    return {
      state: key(frame.head),
      name: this.#filename ? path.basename(this.#filename) : this.#manifest.name,
      kind: ['.csv', '.tsv'].includes(this.#manifest.extension)
        ? 'csv'
        : isOoxmlExtension(this.#manifest.extension)
          ? 'xlsx'
          : this.#manifest.extension.slice(1),
      documentId: this.#manifest.documentId ?? sha256(this.#filename),
      // Private display preferences outlive temporary recovery/grant identities.
      // A clean UUID draft/reload can close and reopen through its canonical file.
      viewPreferenceId: this.#filename ? sha256(this.#filename) : this.#manifest.documentId,
      unsaved: !this.#filename,
      view: this.#book ? frame.view : csvView(frame.index, this.#manifest.baseRevision + ':' + key(frame.head)),
      dirty: frame.head !== this.#manifest.savedHead,
      canUndo: this.#cursor > 0,
      canRedo: this.#cursor < this.#frames.length - 1,
      recovered: this.#recovered,
      warning: this.#publicationWarning ?? this.#maintenanceWarning,
    };
  }
  #serial(operation, options = {}) {
    const work = this.#queue.then(async () => {
      this.#current();
      const previous = [this.#onProgress, this.#beforeCritical, this.#leaveCritical, this.#checkCancelled];
      this.#onProgress = options.onProgress;
      this.#beforeCritical = options.beforeCritical;
      this.#leaveCritical = options.leaveCritical;
      this.#checkCancelled = options.checkCancelled;
      this.#lastPhase = '';
      this.#phaseCount = 0;
      try {
        this.#checkCancelled?.();
        return await operation();
      } finally {
        [this.#onProgress, this.#beforeCritical, this.#leaveCritical, this.#checkCancelled] = previous;
      }
    });
    this.#queue = work.catch(() => {});
    return work;
  }
  async request(request, options = {}) {
    request = normalizeOwnerRequest(request);
    if (request.operation === 'previewFormula') return this.#previewFormula(request, options);
    if (request.operation === 'cancelFormulaPreview') return this.#cancelFormulaPreview(request);
    return this.#serial(async () => {
      if (
        this.#manifest.prepared &&
        ![
          'status',
          'binding',
          'read',
          'scan',
          'trace',
          'usedRange',
          'noteMembership',
          'duplicatePreview',
          'query',
          'selection',
          'filterValues',
          'objects',
          'object',
          'activate',
        ].includes(request.operation)
      )
        throw refusal('session-bad-input', 'Activate this prepared workbook before changing it.');
      switch (request.operation) {
        case 'vbaModules':
        case 'vbaModule':
        case 'vbaFunctions':
        case 'evaluateVba':
        case 'dataOnlyExport': {
          const reader = request.operation === 'dataOnlyExport' ? this.#book : await this.#macroReader();
          if (
            !(reader instanceof PassiveWorkbook) ||
            request.documentId !== this.#documentId() ||
            request.state !== key(this.#current().head)
          )
            throw refusal('session-stale', 'VBA source request is stale or belongs to another workbook.');
          const fields = {
            vbaModules: ['offset', 'limit'],
            vbaModule: ['moduleId', 'offset', 'limit'],
            vbaFunctions: ['moduleId'],
            evaluateVba: ['moduleId', 'functionName', 'args'],
            dataOnlyExport: ['format', 'destination', 'expectedRevision'],
          }[request.operation];
          if (Object.keys(request).some((field) => !['operation', 'documentId', 'state', ...fields].includes(field)))
            throw refusal('internal-invalid-field', 'Invalid passive workbook request fields.');
          // The reader of an editable macro workbook only ever has its first state.
          const state = reader === this.#book ? this.#current().state : '0';
          if (request.operation === 'vbaModules')
            return reader.vbaModules(state, { offset: request.offset, limit: request.limit });
          if (request.operation === 'vbaModule')
            return reader.vbaModule(state, request.moduleId, { offset: request.offset, limit: request.limit });
          if (request.operation === 'vbaFunctions') return reader.vbaFunctions(state, request.moduleId);
          if (request.operation === 'evaluateVba') {
            this.#emit('evaluating-pure-function');
            return reader.evaluateVba(state, request.moduleId, request.functionName, request.args);
          }
          return this.#dataOnlyCopy(request);
        }
        case 'status':
          return this.status();
        case 'binding':
          return { filename: this.#filename ?? null, sourceRevision: this.#manifest.sourceRevision ?? null };
        case 'apply':
          if (request.documentId !== undefined && request.documentId !== this.#documentId())
            throw refusal('session-stale', 'Spreadsheet edit belongs to another document.');
          return this.#apply(request.commands, request.state);
        case 'importBegin': {
          this.#boundImport(request);
          const plan = validatePastePlan({
            shape: request.shape,
            count: request.count,
            policy: request.policy,
            ...(request.conversion !== undefined ? { conversion: request.conversion } : {}),
          });
          if (
            !this.#current().view.sheets.some((sheet) => sheet.name === plan.shape.sheet) ||
            this.#current().view.readOnly
          )
            throw refusal('session-protected', 'Paste target worksheet is unavailable or read-only.');
          if (
            plan.policy === 'typed-text-v1' &&
            this.#current().view.sheets.find((sheet) => sheet.name === plan.shape.sheet).date1904 !==
              plan.conversion.date1904
          )
            throw refusal('session-stale', 'Typed query target calendar changed.');
          this.#emit('uploading-paste');
          return (await this.#pasteOwner()).begin({ ...plan, state: request.state });
        }
        case 'importUpload':
          this.#boundImport(request);
          this.#emit('uploading-paste');
          return (await this.#pasteOwner()).upload({
            id: request.uploadId,
            state: request.state,
            sequence: request.sequence,
            cells: request.cells,
          });
        case 'importSeal':
          this.#boundImport(request);
          this.#emit('checking-paste');
          return (await this.#pasteOwner()).seal({ id: request.uploadId, state: request.state });
        case 'importCancel':
          this.#boundImport(request, false);
          return (await this.#pasteOwner()).cancel({ id: request.uploadId, state: request.state });
        case 'importPrune':
          this.#boundImport(request);
          this.#emit('cleaning-paste', false);
          await this.#prune();
          return { pruned: true };
        case 'importReceipt': {
          this.#boundImport(request);
          validateImportReceiptRequest(request);
          // Re-read the complete hash-verified tip, rather than treating a
          // prior cached verification as proof after a possible interrupted job.
          this.#checkedJournalKey = undefined;
          this.#checkedJournal = undefined;
          const tip = this.#tip(this.#manifest),
            chain = await this.#journalChain(tip);
          this.#current();
          return findImportReceipt(
            chain.map(({ head, node }) => ({
              head,
              parent: node.parent,
              commands: decodedCommands({ encodedCommands: node.commands }),
            })),
            request,
            { baseState: 'base', currentState: key(this.#current().head), tipState: key(tip.head) }
          );
        }
        case 'conditionalRules':
        case 'validationRules': {
          if (
            !this.#book ||
            !isOoxmlExtension(this.#manifest.extension) ||
            request.documentId !== this.#documentId() ||
            request.state !== key(this.#current().head)
          )
            throw refusal('session-stale', 'Workbook rules request is stale or belongs to another document.');
          return this.#book[request.operation](this.#current().state, request.sheet);
        }
        case 'printDefinition': {
          const frame = this.#printFrame(request);
          return this.#book.printDefinition(frame.state, request.sheet);
        }
        case 'printPrepare': {
          const frame = this.#printFrame(request);
          this.#emit('preparing-print');
          const snapshot = await this.#book.printSnapshot(
            frame.state,
            request.sheet,
            request.settings,
            request.geometryProfile
          );
          this.#checkCancelled?.();
          const layout = composePrint(snapshot);
          this.#checkCancelled?.();
          // Prove the bounded document before adopting a preview. It remains
          // owner-resident; no source cell graph enters the desktop host.
          printDocumentHtml(layout);
          const job = {
            ...layout,
            id: randomUUID(),
            state: request.state,
            revision: frame.view.revision,
            pageCount: layout.pages.length,
          };
          this.#printJob = job;
          return { ...job, pages: [job.pages[0]] };
        }
        case 'printPage': {
          const job = this.#print(request);
          if (!Number.isInteger(request.pageNumber) || request.pageNumber < 1 || request.pageNumber > job.pageCount)
            throw refusal('internal-unknown-reference', 'Unknown print page.');
          return job.pages[request.pageNumber - 1];
        }
        case 'printDocument': {
          const job = this.#print(request);
          this.#emit('preparing-print');
          return {
            html: printDocumentHtml(job),
            pageWidthPt: job.pageSizePt[0],
            pageHeightPt: job.pageSizePt[1],
            pageCount: job.pageCount,
          };
        }
        case 'printRelease': {
          // Release is harmless after an accepted edit invalidates the job; it
          // still needs the document identity and the exact opaque job handle.
          if (request.documentId !== this.#documentId() || !validId(request.jobId))
            throw refusal('internal-invalid-field', 'Invalid print preview release.');
          if (this.#printJob?.id === request.jobId) this.#printJob = undefined;
          return { released: true };
        }
        case 'undo':
          return this.#move(-1);
        case 'redo':
          return this.#move(1);
        case 'save':
          return this.#save();
        case 'saveAs':
          return this.#saveAs(request.destination, request.expectedRevision);
        case 'activate':
          if (this.#manifest.prepared) {
            const next = { ...this.#manifest, prepared: false };
            this.#critical('saving-recovery');
            await atomicJson(this.#manifestPath(), next);
            this.#manifest = next;
          }
          return this.status();
        case 'copy':
          return this.#copy(request.destination, request.expectedRevision);
        case 'discard':
          this.#critical('saving-recovery');
          await rm(this.#manifestPath(), { force: true });
          await syncDirectory(this.#directory);
          this.#manifest.savedHead = this.#current().head;
          this.#discarded = true;
          return this.status();
        case 'read': {
          const frame = this.#frame(request.state);
          this.#sheet(request.sheet);
          return this.#book
            ? this.#book.read(frame.state, request.sheet, request.range)
            : readCsv(frame.index, request.range);
        }
        case 'scan':
        case 'query': {
          const frame = this.#frame(request.state);
          this.#sheet(request.sheet);
          if (request.operation === 'query' && (typeof request.query !== 'string' || request.query.length > 10000))
            throw refusal('internal-invalid-field', 'Invalid search text.');
          if (
            request.operation === 'scan' &&
            request.limit !== undefined &&
            (!Number.isInteger(request.limit) ||
              request.limit < 1 ||
              request.limit > 2048 ||
              !(this.#book instanceof IndexedWorkbook))
          )
            throw refusal('session-limit', 'Limited scanning requires a native XLSX page limit from 1 to 2048');
          return this.#book
            ? request.operation === 'scan'
              ? this.#book.scan(
                  frame.state,
                  request.sheet,
                  request.after,
                  request.limit === undefined ? undefined : { limit: request.limit }
                )
              : this.#book.query(frame.state, request.sheet, request.query, request.mode, request.after, {
                  range: request.range,
                  matchLimit: request.matchLimit,
                })
            : scanCsv(
                frame.index,
                request.after,
                request.operation === 'query' ? request.query : undefined,
                request.mode,
                { range: request.range, matchLimit: request.matchLimit }
              );
        }
        case 'usedRange': {
          usedRangeFields(request);
          if (
            !(this.#book instanceof IndexedWorkbook) ||
            !isOoxmlExtension(this.#manifest.extension) ||
            request.documentId !== this.#documentId() ||
            request.state !== key(this.#current().head)
          )
            throw refusal('session-stale', 'UsedRange is stale or belongs to another native XLSX owner');
          const frame = this.#current(),
            book = this.#book,
            receipt = {
              documentId: this.#documentId(),
              ownerId: frame.view.storage.sessionId,
              state: key(frame.head),
              revision: frame.view.revision,
            };
          this.#checkCancelled?.();
          const raw = await book.usedRange(frame.state, request.sheet);
          this.#checkCancelled?.();
          if (
            this.#current() !== frame ||
            this.#book !== book ||
            request.documentId !== this.#documentId() ||
            request.state !== key(frame.head) ||
            frame.view.revision !== receipt.revision
          )
            throw refusal('session-stale', 'UsedRange owner/head retired before its reply');
          return { usedRange: bindUsedRange(raw, receipt, { sheet: request.sheet }) };
        }
        case 'noteMembership': {
          noteMembershipFields(request);
          if (
            !(this.#book instanceof IndexedWorkbook) ||
            !isOoxmlExtension(this.#manifest.extension) ||
            request.documentId !== this.#documentId() ||
            request.state !== key(this.#current().head)
          )
            throw refusal('session-stale', 'Notes membership is stale or belongs to another native XLSX owner');
          const frame = this.#current(),
            book = this.#book,
            receipt = {
              documentId: this.#documentId(),
              ownerId: frame.view.storage.sessionId,
              state: key(frame.head),
              revision: frame.view.revision,
            };
          this.#checkCancelled?.();
          const raw = await book.noteMembership(frame.state, request.sheet);
          this.#checkCancelled?.();
          if (
            this.#current() !== frame ||
            this.#book !== book ||
            request.documentId !== this.#documentId() ||
            request.state !== key(frame.head) ||
            frame.view.revision !== receipt.revision
          )
            throw refusal('session-stale', 'Notes membership owner/head retired before its reply');
          return { noteMembership: bindNoteMembership(raw, receipt, { sheet: request.sheet }) };
        }
        case 'duplicatePreview': {
          if (
            !(this.#book instanceof IndexedWorkbook) ||
            !isOoxmlExtension(this.#manifest.extension) ||
            request.documentId !== this.#documentId() ||
            request.state !== key(this.#current().head)
          )
            throw refusal('session-duplicate', 'Duplicate preview is stale or belongs to another native XLSX owner');
          if (
            Object.keys(request).some(
              (k) =>
                !['operation', 'documentId', 'state', 'sheet', 'range', 'hasHeaders', 'columns', 'offset'].includes(k)
            )
          )
            throw refusal('internal-invalid-field', 'Invalid duplicate preview request fields');
          const frame = this.#current(),
            book = this.#book,
            receipt = {
              documentId: this.#documentId(),
              ownerId: frame.view.storage.sessionId,
              state: key(frame.head),
              revision: frame.view.revision,
            };
          this.#checkCancelled?.();
          const raw = await book.duplicatePreview(
            frame.state,
            request.sheet,
            request.range,
            request.hasHeaders,
            request.columns,
            request.offset ?? 0
          );
          this.#checkCancelled?.();
          if (this.#current() !== frame || this.#book !== book || request.documentId !== this.#documentId())
            throw refusal('session-duplicate', 'Duplicate preview owner/head retired before its reply');
          const result = { ...raw, ...receipt };
          validateDuplicatePreview(result, receipt, request, request.offset ?? 0);
          return { duplicatePreview: result };
        }
        case 'trace': {
          if (
            !(this.#book instanceof IndexedWorkbook) ||
            !isOoxmlExtension(this.#manifest.extension) ||
            request.documentId !== this.#documentId() ||
            request.state !== key(this.#current().head)
          )
            throw refusal('session-stale', 'Formula trace is stale or belongs to another native XLSX owner.');
          if (
            Object.keys(request).some(
              (k) => !['operation', 'documentId', 'state', 'sheet', 'cell', 'direction', 'after', 'limit'].includes(k)
            )
          )
            throw refusal('internal-invalid-field', 'Invalid formula trace request fields.');
          const frame = this.#current();
          this.#checkCancelled?.();
          const result = await this.#book.trace(frame.state, request.sheet, request.cell, {
            direction: request.direction,
            ...(request.after !== undefined ? { after: request.after } : {}),
            ...(request.limit !== undefined ? { limit: request.limit } : {}),
          });
          this.#checkCancelled?.();
          if (this.#current() !== frame || request.documentId !== this.#documentId())
            throw refusal('session-stale', 'Formula trace owner/head retired before its reply');
          return {
            ...result,
            documentId: this.#documentId(),
            ownerId: frame.view.storage.sessionId,
            state: key(frame.head),
            revision: frame.view.revision,
          };
        }
        case 'selection': {
          const frame = this.#frame(request.state);
          this.#sheet(request.sheet);
          return this.#book
            ? this.#book.selectionSummary(frame.state, request.sheet, request.range)
            : summarizeCsv(frame.index, request.range);
        }
        case 'filterValues': {
          if (!this.#book) throw refusal('session-bad-input', 'Filter value pages require an XLSX workbook.');
          const frame = this.#frame(request.state);
          this.#sheet(request.sheet);
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
            throw refusal('internal-invalid-field', 'Invalid filter value page request.');
          const col = cellPosition(column + '1').col;
          if (col < bounds.first.col || col > bounds.last.col)
            throw refusal('session-bad-input', 'Filter column must be inside its range.');
          return this.#book.filterValues(frame.state, request.sheet, request.range, column, query, offset, limit);
        }
        case 'accessibility': {
          accessibilityFields(request);
          if (
            !(this.#book instanceof IndexedWorkbook) ||
            !isOoxmlExtension(this.#manifest.extension) ||
            request.documentId !== this.#documentId() ||
            request.state !== key(this.#current().head)
          )
            throw refusal('session-stale', 'Accessibility belongs to another or retired native XLSX owner.');
          const frame = this.#current(),
            book = this.#book,
            receipt = {
              documentId: this.#documentId(),
              ownerId: frame.view.storage.sessionId,
              state: key(frame.head),
              revision: frame.view.revision,
            };
          this.#checkCancelled?.();
          const raw = await book.accessibility(frame.state);
          this.#checkCancelled?.();
          if (
            this.#current() !== frame ||
            this.#book !== book ||
            request.documentId !== this.#documentId() ||
            request.state !== key(frame.head) ||
            frame.view.revision !== receipt.revision
          )
            throw refusal('session-stale', 'Accessibility owner/head retired before its reply.');
          return { accessibility: bindAccessibility(raw, receipt) };
        }
        case 'notes': {
          notesFields(request);
          if (
            !(this.#book instanceof IndexedWorkbook) ||
            !isOoxmlExtension(this.#manifest.extension) ||
            request.documentId !== this.#documentId() ||
            request.state !== key(this.#current().head)
          )
            throw refusal('session-stale', 'Notes belong to another or retired native XLSX owner.');
          const frame = this.#current(),
            book = this.#book,
            receipt = {
              documentId: this.#documentId(),
              ownerId: frame.view.storage.sessionId,
              state: key(frame.head),
              revision: frame.view.revision,
            };
          this.#checkCancelled?.();
          const raw = await book.notes(frame.state, request.sheet);
          this.#checkCancelled?.();
          if (
            this.#current() !== frame ||
            this.#book !== book ||
            request.documentId !== this.#documentId() ||
            request.state !== key(frame.head) ||
            frame.view.revision !== receipt.revision
          )
            throw refusal('session-stale', 'Note owner/head retired before its reply.');
          return { notes: bindNotes(raw, receipt, { sheet: request.sheet }) };
        }
        case 'hyperlink': {
          hyperlinkFields(request);
          if (
            !(this.#book instanceof IndexedWorkbook) ||
            !isOoxmlExtension(this.#manifest.extension) ||
            request.documentId !== this.#documentId() ||
            request.state !== key(this.#current().head)
          )
            throw refusal('session-stale', 'Hyperlink inspection is stale or belongs to another native XLSX owner.');
          const frame = this.#current(),
            book = this.#book,
            receipt = {
              documentId: this.#documentId(),
              ownerId: frame.view.storage.sessionId,
              state: key(frame.head),
              revision: frame.view.revision,
            };
          this.#checkCancelled?.();
          const raw = await book.hyperlink(frame.state, request.sheet, request.cell);
          this.#checkCancelled?.();
          if (
            this.#current() !== frame ||
            this.#book !== book ||
            request.documentId !== this.#documentId() ||
            request.state !== key(frame.head) ||
            frame.view.revision !== receipt.revision
          )
            throw refusal('session-stale', 'Hyperlink owner/head retired before its reply.');
          return { hyperlink: bindHyperlink(raw, receipt, { sheet: request.sheet, cell: request.cell }) };
        }
        case 'objects':
        case 'object': {
          if (!this.#book) throw refusal('session-bad-input', 'Workbook objects require an XLSX workbook.');
          if (request.documentId !== this.#documentId() || request.state !== key(this.#current().head))
            throw refusal('session-stale', 'Workbook object request is stale or belongs to another document.');
          if (
            typeof request.sheet !== 'string' ||
            request.sheet.length < 1 ||
            request.sheet.length > 31 ||
            !this.#current().view.sheets.some((sheet) => sheet.name === request.sheet)
          )
            throw refusal('internal-unknown-reference', 'Unknown object worksheet.');
          const frame = this.#current();
          if (request.operation === 'objects') {
            const limit = request.limit ?? 50;
            if (
              !Number.isInteger(limit) ||
              limit < 1 ||
              limit > 50 ||
              (request.after !== undefined && !objectHandle(request.after)) ||
              request.objectId !== undefined
            )
              throw refusal('internal-invalid-field', 'Invalid object page request.');
            return this.#book.objects(frame.state, request.sheet, { after: request.after, limit });
          }
          if (!objectHandle(request.objectId) || request.after !== undefined || request.limit !== undefined)
            throw refusal('internal-invalid-field', 'Invalid object preview request.');
          return this.#book.object(frame.state, request.sheet, request.objectId);
        }
        default:
          throw refusal('session-unsupported', 'Unsupported spreadsheet operation.');
      }
    }, options);
  }
  #sheet(name) {
    if (typeof name !== 'string' || (!this.#book && name !== 'CSV'))
      throw refusal('internal-unknown-reference', 'Unknown worksheet.');
  }
  async #applyFrame(previous, commands) {
    if (this.#book) {
      const priorParent = this.#importParent,
        priorImport = this.#applyingImport;
      this.#importParent = key(previous.head);
      this.#applyingImport = commands.some((command) => command?.type === 'importCells');
      this.#importPercent = -1;
      try {
        const result = await this.#book.apply(previous.state, commands);
        return {
          view: result.view,
          state: result.view.storage.state,
          ...(result.unchanged === true ? { unchanged: true } : {}),
          ...(result.commands ? { commands: result.commands } : {}),
        };
      } finally {
        this.#importParent = priorParent;
        this.#applyingImport = priorImport;
      }
    }
    if (
      commands.some(
        (command) =>
          Object.hasOwn(command, 'valueType') ||
          (command.type === 'setCells' &&
            Array.isArray(command.cells) &&
            command.cells.some((cell) => Object.hasOwn(cell, 'valueType')))
      )
    )
      throw refusal('session-unsupported', 'CSV supports literal text only; declared error edits are unsupported.');
    const index = applyCsv(previous.index, commands);
    // Existing source text stays in the immutable snapshot. Large blank expansions would
    // otherwise create an unbounded patch array during CSV serialization.
    const expansion =
      Math.max(0, index.storedCells - this.#csvBaseline.storedCells) * 48 +
      Math.max(0, index.rows - this.#csvBaseline.rows) * 32;
    if (!Number.isSafeInteger(expansion) || expansion > JOURNAL_LIMIT)
      throw refusal('session-limit', 'CSV expansion exceeds the 64 MiB edit projection budget.');
    return { index };
  }
  async #apply(commands, expectedState) {
    this.#emit('applying');
    if (this.#manifest.prepared) throw refusal('session-bad-input', 'Activate this prepared workbook before editing.');
    const previous = this.#current();
    if (expectedState !== key(previous.head)) throw refusal('session-stale', 'Spreadsheet edit is stale.');
    if (!Array.isArray(commands) || !commands.length || commands.length > 10000)
      throw refusal('session-bad-input', 'Submit between one and 10,000 commands.');
    const head = randomUUID();
    let frame, node, bytes, hash;
    try {
      const preliminary = ownerCommands(commands),
        incoming = preliminary.encodedCommands ?? JSON.stringify(preliminary.commands);
      if (Buffer.byteLength(incoming) > COMMAND_LIMIT)
        throw refusal('session-limit', 'Recovery command budget exceeded.');
      await this.#verifyImports(commands, previous.head);
      if (
        this.#journalBytes +
          (this.#paste?.stats().attachmentBytes ?? 0) +
          Buffer.byteLength(incoming) +
          RECOVERY_MANIFEST_LIMIT >
          JOURNAL_LIMIT ||
        (await directoryBytes(this.#directory)) + Buffer.byteLength(incoming) + RECOVERY_MANIFEST_LIMIT > this.#budget
      )
        throw refusal('session-limit', 'Recovery disk budget exceeded. Save and reopen before continuing.');
      frame = await this.#applyFrame(previous, commands);
      if (frame.unchanged === true) {
        if (
          !commands.every((command) =>
            [
              'setSheetBackground',
              'removeSheetBackground',
              'applyPageTheme',
              'booleanCell',
              'clearHyperlinks',
              'setHyperlink',
              'removeHyperlink',
              'setNote',
              'removeNote',
            ].includes(command.type)
          ) ||
          frame.state !== previous.state ||
          this.#current() !== previous
        )
          throw refusal('internal-invalid-field', 'Invalid no-change hyperlink metadata owner receipt.');
        this.#checkCancelled?.();
        return this.status();
      }
      const durableCommands = frame.commands ?? commands;
      if (!Array.isArray(durableCommands) || !durableCommands.length || durableCommands.length > 10000)
        throw refusal('internal-invalid-field', 'Invalid normalized recovery commands.');
      await this.#verifyImports(durableCommands, previous.head);
      const wire = ownerCommands(durableCommands),
        serialized = wire.encodedCommands ?? JSON.stringify(wire.commands);
      node = { schema: 1, parent: previous.head, parentHash: previous.hash, commands: serialized };
      const nodeText = JSON.stringify(node);
      bytes = Buffer.byteLength(nodeText);
      hash = sha256(nodeText);
      if (
        Buffer.byteLength(serialized) > COMMAND_LIMIT ||
        bytes > NODE_LIMIT ||
        this.#journalBytes + (this.#paste?.stats().attachmentBytes ?? 0) + bytes + RECOVERY_MANIFEST_LIMIT >
          JOURNAL_LIMIT ||
        (await directoryBytes(this.#directory)) + bytes + RECOVERY_MANIFEST_LIMIT > this.#budget
      )
        throw refusal('session-limit', 'Normalized recovery disk budget exceeded. Save and reopen before continuing.');
      const accepted = { ...frame };
      delete accepted.commands;
      const frames = [...this.#frames.slice(0, this.#cursor + 1), { ...accepted, head, hash }].slice(-101),
        next = { ...this.#manifest, head, headHash: hash, history: this.#history(frames, frames.length - 1) };
      if (this.#manifest.savedSources) {
        const parent = await this.#journalChain({ ...this.#manifest, head: previous.head, headHash: previous.hash }),
          ordered = [
            { head: null, hash: null },
            ...parent
              .slice()
              .reverse()
              .map((item) => ({ head: item.head, hash: item.hash })),
            { head, hash },
          ];
        next.savedSources = this.#selectSavedSources(this.#manifest.savedSources, frames, ordered);
      }
      this.#checkCancelled?.();
      this.#critical('saving-recovery');
      this.#current();
      if (this.#saveRecalculateRevision !== undefined)
        await assertRevision(this.#filename, this.#saveRecalculateRevision);
      await atomicJson(this.#nodePath(head), node);
      await atomicJson(this.#manifestPath(), next);
      this.#manifest = next;
      this.#journalBytes += bytes;
      this.#discarded = false;
      this.#frames = frames;
      this.#cursor = this.#frames.length - 1;
      this.#checkedJournal = undefined;
      this.#checkedJournalKey = undefined;
      this.#printJob = undefined;
    } catch (error) {
      await this.#maintain({ prune: !!this.#paste });
      throw error;
    }
    if (this.#frames.length % 25 === 0 || this.#paste) await this.#maintain({ prune: !!this.#paste });
    return this.status();
  }
  async #move(direction) {
    const cursor = this.#cursor + direction;
    if (cursor < 0 || cursor >= this.#frames.length) return this.status();
    const frame = this.#frames[cursor],
      next = {
        ...this.#manifest,
        head: frame.head,
        headHash: frame.hash,
        history: this.#history(this.#frames, cursor),
      };
    this.#critical('saving-recovery');
    await atomicJson(this.#manifestPath(), next);
    this.#manifest = next;
    this.#cursor = cursor;
    this.#printJob = undefined;
    return this.status();
  }
  async #export(publication = false) {
    this.#emit('exporting');
    if (this.#book)
      return (
        await this.#book.export(
          this.#current().state,
          ...(this.#book instanceof IndexedWorkbook ? [{ publication }] : [])
        )
      ).path;
    const output = path.join(this.#directory, 'export-' + randomUUID() + this.#manifest.extension);
    await exportCsv(this.#current().index, this.#snapshotPath(), output);
    return output;
  }
  async #recalculateForPublication(checkSource = false) {
    if (!this.#book) return;
    const frame = this.#current(),
      policy = frame.view.calculationPolicy;
    if (frame.view.readOnly || policy?.protectedSource || policy?.mode !== 'manual' || !policy.recalculateBeforeSave)
      return;
    this.#saveRecalculateRevision = checkSource ? this.#manifest.sourceRevision : undefined;
    try {
      await this.#apply([{ type: 'recalculate' }], key(frame.head));
    } finally {
      this.#saveRecalculateRevision = undefined;
    }
    if (checkSource) await assertRevision(this.#filename, this.#manifest.sourceRevision);
  }
  async #save() {
    if (this.#book instanceof PassiveWorkbook)
      throw refusal('session-protected', 'Passive workbooks are read-only; use same-format Save Copy.');
    this.#emit('checking-source');
    if (!this.#filename)
      throw refusal('session-bad-input', 'Choose a filename through Save before saving this untitled workbook.');
    await assertRevision(this.#filename, this.#manifest.sourceRevision);
    await this.#recalculateForPublication(true);
    const output = await this.#export(true),
      savedRevision = await revision(output),
      prepared = await this.#prepareSavedSource(output, savedRevision),
      previous = this.#manifest;
    let temporary,
      stagedIdentity,
      intent = false,
      published = false;
    try {
      this.#critical('writing-output');
      this.#assertSavedSource(prepared);
      temporary = await stageCopy(output, this.#filename);
      await chmod(temporary, (await stat(this.#filename)).mode & 0o777);
      stagedIdentity = await destinationIdentity(temporary);
      if ((await revision(temporary)) !== savedRevision)
        throw refusal('session-stale', 'Saved workbook bytes changed before publication.');
      this.#assertSavedSource(prepared);
      const pending = {
        ...previous,
        pendingPublication: {
          head: this.#current().head,
          revision: savedRevision,
          destinationIdentity: stagedIdentity,
          ...(prepared ? { savedSource: prepared.checkpoint } : {}),
        },
      };
      await atomicJson(this.#manifestPath(), pending);
      intent = true;
      if (prepared) prepared.intent = true;
      await assertRevision(this.#filename, previous.sourceRevision);
      this.#assertSavedSource(prepared);
      await rename(temporary, this.#filename);
      published = true;
      await syncDirectory(path.dirname(this.#filename));
      const next = {
        ...previous,
        sourceRevision: savedRevision,
        savedHead: prepared ? prepared.frame.head : this.#current().head,
      };
      // A crash between rename and this manifest update is resolved by the saved hash.
      return await this.#finishSavedSource(prepared, next);
    } catch (error) {
      if (
        (published || (intent && stagedIdentity && (await this.#destinationMatches(this.#filename, stagedIdentity)))) &&
        (await revision(this.#filename).catch(() => null)) === savedRevision
      )
        return await this.#finishSavedSource(
          prepared,
          {
            ...previous,
            sourceRevision: savedRevision,
            savedHead: prepared ? prepared.frame.head : this.#current().head,
          },
          'The workbook was saved; publication reported an error after the saved bytes were confirmed. ' +
            String(error.message ?? error)
        );
      this.#manifest = previous;
      if (intent)
        try {
          await atomicJson(this.#manifestPath(), previous);
          if (prepared) prepared.intent = false;
        } catch {
          this.#publicationWarning = 'Save failed; its recovery receipt will be checked on restart.';
        }
      throw error;
    } finally {
      if (temporary) await rm(temporary, { force: true });
      await this.#discardSavedSource(prepared);
      if (!this.#book) await rm(output, { force: true });
    }
  }
  async #saveAs(destination, expectedRevision) {
    this.#emit('checking-recovery');
    if (this.#filename)
      throw refusal('session-duplicate', 'This workbook already has a filename. Use Save or Save Copy.');
    if (this.#manifest.prepared) throw refusal('session-bad-input', 'Activate this prepared workbook before saving.');
    if (
      !filePath(destination) ||
      path.extname(destination).toLowerCase() !== '.xlsx' ||
      (expectedRevision !== null && !validHash(expectedRevision))
    )
      throw refusal('session-bad-input', 'Choose an XLSX destination with an expected revision.');
    const target = path.join(await realpath(path.dirname(destination)), path.basename(destination)),
      relative = path.relative(this.#recoveryRoot, target);
    if (!relative.startsWith('..' + path.sep) && relative !== '..' && !path.isAbsolute(relative))
      throw refusal('session-unsupported', 'Choose a destination outside private workbook recovery storage.');
    const { records } = await listDraftRecords(this.#recoveryRoot);
    if (
      records.some(
        ({ manifest: m }) =>
          m.documentId !== this.#manifest.documentId &&
          m.filename === target &&
          (m.head !== m.savedHead || this.#tip(m).head !== m.savedHead || m.pendingPublication)
      )
    )
      throw refusal(
        'session-refused',
        'The destination has pending workbook recovery. Recover or save that document before replacing it.'
      );
    const legacy = path.join(this.#recoveryRoot, sha256(target), 'current.json');
    try {
      const saved = validateRecoveryManifest(await readJson(legacy));
      if (saved.head !== saved.savedHead || this.#tip(saved).head !== saved.savedHead)
        throw refusal(
          'session-refused',
          'The destination has pending workbook recovery. Recover it before replacing it.'
        );
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
    await this.#recalculateForPublication();
    const output = await this.#export(true),
      candidateRevision = await revision(output),
      prepared = await this.#prepareSavedSource(output, candidateRevision),
      previous = this.#manifest;
    let stagedIdentity,
      intent = false,
      published = false;
    try {
      this.#critical('writing-output');
      this.#assertSavedSource(prepared);
      await publishCopy(output, target, expectedRevision, async (identity, stagedRevision) => {
        if (stagedRevision !== candidateRevision)
          throw refusal('session-stale', 'Saved workbook bytes changed before publication.');
        stagedIdentity = identity;
        this.#assertSavedSource(prepared);
        const pending = {
          ...previous,
          pendingPublication: {
            kind: 'bind',
            destination: target,
            expectedRevision,
            head: this.#current().head,
            revision: candidateRevision,
            destinationIdentity: identity,
            ...(prepared ? { savedSource: prepared.checkpoint } : {}),
          },
        };
        await atomicJson(this.#manifestPath(), pending);
        this.#manifest = pending;
        intent = true;
        if (prepared) prepared.intent = true;
        this.#assertSavedSource(prepared);
      });
      published = true;
      this.#filename = target;
      return await this.#finishSavedSource(prepared, {
        ...previous,
        filename: target,
        sourceRevision: candidateRevision,
        savedHead: prepared ? prepared.frame.head : this.#current().head,
      });
    } catch (error) {
      if (
        error.code !== 'EEXIST' &&
        (published || (intent && stagedIdentity && (await this.#destinationMatches(target, stagedIdentity)))) &&
        (await revision(target).catch(() => null)) === candidateRevision
      ) {
        this.#filename = target;
        return await this.#finishSavedSource(
          prepared,
          {
            ...previous,
            filename: target,
            sourceRevision: candidateRevision,
            savedHead: prepared ? prepared.frame.head : this.#current().head,
          },
          'The workbook was saved; publication reported an error after the saved bytes were confirmed. ' +
            String(error.message ?? error)
        );
      }
      this.#manifest = previous;
      if (intent)
        try {
          await atomicJson(this.#manifestPath(), previous);
          if (prepared) prepared.intent = false;
        } catch {
          this.#publicationWarning = 'The first Save failed; its recovery receipt will be checked again on restart.';
        }
      throw error;
    } finally {
      await this.#discardSavedSource(prepared);
    }
  }
  async #copy(destination, expectedRevision) {
    if (typeof destination !== 'string' || path.extname(destination).toLowerCase() !== this.#manifest.extension)
      throw refusal('session-bad-input', 'Use the same file format for Save Copy.');
    const target = path.join(await realpath(path.dirname(destination)), path.basename(destination));
    const relative = path.relative(this.#recoveryRoot, target);
    if (!relative.startsWith('..' + path.sep) && relative !== '..' && !path.isAbsolute(relative))
      throw refusal('session-unsupported', 'Choose a copy destination outside private workbook recovery storage.');
    if (target === this.#filename) throw refusal('session-bad-input', 'Choose another filename, or use Save.');
    await this.#recalculateForPublication();
    const output = await this.#export();
    this.#critical('writing-output');
    try {
      await publishCopy(output, target, expectedRevision);
      return { copied: true };
    } finally {
      if (!this.#book) await rm(output, { force: true });
    }
  }
  async #dataOnlyCopy(request) {
    const { format, destination, expectedRevision } = request;
    if (
      !['xls', 'xlsb', 'xlsx'].includes(format) ||
      !filePath(destination) ||
      path.extname(destination).toLowerCase() !== '.' + format ||
      expectedRevision !== null
    )
      throw refusal(
        'session-unsupported',
        'Choose a new XLS, XLSB or XLSX data-only destination. Existing files cannot be replaced.'
      );
    const target = path.join(await realpath(path.dirname(destination)), path.basename(destination));
    for (const root of [this.#recoveryRoot, ...this.#book.storageRoots()]) {
      const relative = path.relative(await realpath(root), target);
      if (!relative.startsWith('..' + path.sep) && relative !== '..' && !path.isAbsolute(relative))
        throw refusal('session-unsupported', 'Choose a new data-only destination outside private workbook storage.');
    }
    if (target === this.#filename)
      throw refusal('session-unsupported', 'The original file cannot be a data-only destination.');
    try {
      await lstat(target);
      throw refusal('session-duplicate', 'Choose a new data-only filename; the destination already exists.');
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
    this.#emit('exporting-stored-values');
    const output = await this.#book.dataOnlyExport(this.#current().state, { format });
    this.#checkCancelled?.();
    this.#current();
    this.#critical('writing-output');
    try {
      await publishCopy(output.path, target, null);
      const { path: privatePath, ...metadata } = output;
      return { copied: true, ...metadata };
    } finally {
      this.#releaseCritical();
    }
  }
  async #compact() {
    if (!this.#book) return;
    // Cancellation forbids further publication. Compact is private reclamation
    // of unaccepted versions and explicitly keeps every accepted history state.
    this.#cleanupNative = true;
    try {
      await this.#book.compact(
        this.#frames.map((frame) => frame.state),
        { retainBase: false }
      );
    } finally {
      this.#cleanupNative = false;
    }
  }
  async #maintain({ prune = false } = {}) {
    try {
      await this.#compact();
      if (prune) await this.#prune();
      this.#maintenanceWarning = undefined;
    } catch (error) {
      this.#maintenanceWarning =
        'Accepted edits are preserved, but temporary index cleanup failed: ' + String(error.message ?? error);
      console.warn(this.#maintenanceWarning);
    }
  }
  async #prune() {
    const keep = new Set(['current.json', path.basename(this.#snapshotPath())]);
    const attachments = new Set(),
      chain = this.#discarded ? [] : await this.#journalChain(this.#tip(this.#manifest));
    if (this.#manifest.savedSources) {
      const ordered = [
          { head: null, hash: null },
          ...chain
            .slice()
            .reverse()
            .map((item) => ({ head: item.head, hash: item.hash })),
        ],
        sources = this.#discarded ? [] : this.#selectSavedSources(this.#manifest.savedSources, this.#frames, ordered);
      if (JSON.stringify(sources) !== JSON.stringify(this.#manifest.savedSources)) {
        const next = { ...this.#manifest, savedSources: sources };
        await atomicJson(this.#manifestPath(), next);
        this.#manifest = next;
      }
      for (const source of sources) {
        keep.add(source.snapshot + '.xlsx');
        if (source.calculation) keep.add(source.calculation.snapshot + '.calculation.sqlite');
      }
    }
    if (this.#manifest.pendingPublication?.savedSource) {
      const source = this.#manifest.pendingPublication.savedSource;
      keep.add(source.snapshot + '.xlsx');
      if (source.calculation) keep.add(source.calculation.snapshot + '.calculation.sqlite');
    }
    for (const { head, node } of chain) {
      keep.add(head + '.json');
      for (const descriptor of this.#importDescriptors(decodedCommands({ encodedCommands: node.commands })))
        attachments.add(descriptor.id);
    }
    if (!this.#paste) {
      try {
        await lstat(path.join(this.#directory, 'paste-v1'));
        await this.#pasteOwner();
      } catch (error) {
        if (error.code !== 'ENOENT') throw error;
      }
    }
    if (this.#paste) await this.#paste.prune([...attachments]);
    for (const name of await readdir(this.#directory))
      if (
        !keep.has(name) &&
        (/^[a-f0-9-]{36}\.(json|xlsx|xlsm|xls|xlsb|csv|tsv|calculation\.sqlite)$/.test(name) ||
          /^export-[a-f0-9-]{36}\.(csv|tsv)$/.test(name))
      )
        await rm(path.join(this.#directory, name), { force: true });
  }
  #previewBinding(request) {
    const frame = this.#current();
    if (
      !(this.#book instanceof IndexedWorkbook) ||
      !isOoxmlExtension(this.#manifest.extension) ||
      request.documentId !== this.#documentId() ||
      request.state !== key(frame.head)
    )
      throw refusal('session-stale', 'Formula preview belongs to a retired or different XLSX owner.');
    return {
      frame,
      book: this.#book,
      receipt: {
        documentId: this.#documentId(),
        ownerId: frame.view.storage.sessionId,
        state: key(frame.head),
        revision: frame.view.revision,
      },
    };
  }
  #previewFormula(raw, options) {
    const request = Object.freeze({ ...raw });
    formulaPreviewFields(request);
    const binding = this.#previewBinding(request);
    if (this.#formulaPreviews.size >= 8 || this.#formulaPreviewIds.has(request.requestId))
      throw refusal('session-unsupported', 'Formula preview request IDs cannot be reused.');
    if (this.#formulaPreviewIds.size >= 4096)
      throw refusal('session-limit', 'Formula preview session limit reached. Save and reopen this workbook.');
    this.#formulaPreviewIds.add(request.requestId);
    const job = { ...binding, cancelled: false, started: false };
    this.#formulaPreviews.set(request.requestId, job);
    return this.#serial(async () => {
      const current = this.#previewBinding(request);
      if (current.frame !== job.frame || current.book !== job.book)
        throw refusal('session-stale', 'Formula preview owner/head retired before reading.');
      const cancelled = () => ({
        schema: 1,
        requestId: request.requestId,
        sheet: request.sheet,
        cell: request.cell,
        formula: request.formula,
        authority: 'draft-preview',
        result: { status: 'cancelled', reason: 'The formula draft was retired.' },
      });
      let reply;
      if (job.cancelled) reply = cancelled();
      else {
        job.started = true;
        reply = await job.book.previewFormula(job.frame.state, request.sheet, request.cell, request.formula, {
          requestId: request.requestId,
        });
      }
      const final = this.#previewBinding(request);
      if (final.frame !== job.frame || final.book !== job.book)
        throw refusal('session-stale', 'Formula preview owner/head retired before its reply.');
      // This private job can be cancelled without invoking the owner-wide gate.
      if (job.cancelled) reply = cancelled();
      return encodeOwnerResult({ formulaPreview: bindFormulaPreview(reply, job.receipt, request) });
    }, options).finally(() => {
      if (this.#formulaPreviews.get(request.requestId) === job) this.#formulaPreviews.delete(request.requestId);
    });
  }
  async #cancelFormulaPreview(raw) {
    const request = Object.freeze({ ...raw });
    formulaPreviewFields(request, { cancel: true });
    this.#previewBinding(request);
    const job = this.#formulaPreviews.get(request.requestId);
    if (!job || job.frame !== this.#current() || job.book !== this.#book)
      return { formulaPreviewCancellation: { schema: 1, requestId: request.requestId, cancelled: false } };
    job.cancelled = true;
    if (job.started)
      formulaPreviewCancellation(
        await job.book.cancelFormulaPreview(job.frame.state, request.requestId),
        request.requestId
      );
    return { formulaPreviewCancellation: { schema: 1, requestId: request.requestId, cancelled: true } };
  }
  async close() {
    if (this.#closed) return;
    this.#closed = true;
    await this.#book?.close();
    await this.#macroBook?.close();
    await this.#queue;
    if (this.#ready && this.#paste) {
      try {
        await this.#prune();
      } catch {
        /* Retain recovery attachments when validation or cancellation prevents safe pruning. */
      }
    }
    await this.#paste?.close();
    this.#paste = undefined;
    this.#frames = [];
    this.#book = undefined;
    this.#printJob = undefined;
    if (this.#created && !this.#manifest) {
      await rm(this.#directory, { recursive: true, force: true });
      return;
    }
    // Untouched/cancelled candidates and explicitly discarded drafts have no
    // accepted recovery to retain. Clean UUID bindings can reopen from their file.
    if (
      (this.#ready || this.#created) &&
      this.#manifest?.schema === 2 &&
      !this.#publicationWarning &&
      (this.#manifest.prepared ||
        this.#discarded ||
        (this.#manifest.head === this.#manifest.savedHead &&
          this.#tip(this.#manifest).head === this.#manifest.savedHead))
    )
      await rm(this.#directory, { recursive: true, force: true });
  }
}
