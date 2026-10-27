import { FORMAT, formatTemplate, matchesFormat, stripFormatPrefix } from './format.mjs';
import { isRefusal, refusal, reworded } from './refusal-codes.mjs';
import { preservationPolicy, assertPreservationCommands } from './preservation-editing.mjs';
import { captureInkStrokeCommands, isInkStrokeCommand } from './ink-stroke.mjs';
import { indexedAccessibility } from './indexed-accessibility.mjs';
import { captureNoteCommands, isNoteCommand, noteInventory } from './notes.mjs';
import { captureModel3DCommands, isModel3DCommand } from './model3d-authoring.mjs';
import { normalizeBackgroundCommands } from './background-command.mjs';
import { normalizePageThemeCommands } from './page-theme-command.mjs';
import { normalizeBooleanCellCommands } from './cell-insertions.mjs';
import { normalizeHyperlinkCommands, hyperlinkCell } from './hyperlinks.mjs';
import { normalizeFillAcrossSheetsRuns } from './fill-across-sheets.mjs';
import { duplicateConfiguration, duplicatePage, bindDuplicatePreview } from './duplicate-model.mjs';
import { formulaPreviewRequestId, formulaPreviewTarget, formulaPreviewInner } from './formula-preview.mjs';
import { usedRangeSheet, normalizeNativeUsedRange } from './used-range.mjs';
import { noteMembershipSheet, noteMembershipInner } from './note-membership.mjs';
import { indexedNoteMembership } from './indexed-notes.mjs';
import { traceCursorKey, traceRequest, readTraceCursor, traceReceipt } from './indexed-trace.mjs';
import { isIndexedTableCommand } from './indexed-tables.mjs';
import { captureSmartArtCommands, isSmartArtCommand } from './smartart-model.mjs';
import { captureShapeCommands, isShapeCommand } from './shape-model.mjs';
import {
  isIndexedPivotCommand,
  indexedPivotPlan,
  capturePivotCommands,
  indexedPivotInventory,
} from './indexed-pivots.mjs';
import {
  indexedObjectList,
  indexedObjectPreview,
  indexedObjectCachePlan,
  indexedObjectHasManagedCharts,
} from './indexed-object-authoring.mjs';
import { indexedPrintDefinition } from './indexed-print.mjs';
import { createPrintSnapshot } from './print-snapshot.mjs';
import { indexedRulesDefinition, indexedRuleDuplicates } from './indexed-rules.mjs';
import { indexedSpillMetadata, withSpillDimensions } from './indexed-spills.mjs';
/** Host-neutral session coordinator. Rust owns cells and package writes.
 * Small metadata/reference projections preserve the existing JS semantics;
 * unqualified mutations reject before publication, never load a workbook writer. */
import fs from 'node:fs/promises';
import { strFromU8, unzipSync } from 'fflate';
import path from 'node:path';
import { indexedCache, fileDigest } from './indexed-cache.mjs';
import {
  indexedCalculationReusable,
  publishIndexedCheckpoint,
  restoreIndexedCheckpoint,
  resetIndexedSnapshotExtents,
} from './indexed-checkpoint.mjs';
import { indexedViewGuards } from './indexed-metadata.mjs';
import { createIndexedStorage } from './indexed-storage.mjs';
import { randomUUID } from 'node:crypto';
import { viewFromEngine, indexedCalculationMetadata, indexedStylePublication } from './workbook-package.mjs';
import { normalizeClearHyperlinksCommands } from './clear-hyperlinks.mjs';
import { cellPosition, columnName, rangeBounds } from './workbook-commands.mjs';
import { boundedObjectMedia, empty, nativeValue, own, FACTS_V2, uuidV4 } from './commands/shared.mjs';
import { admitAgainstMetadata, directModules, guardedModules, routeCommands } from './commands/index.mjs';

// The code is the signal. The wording is still checked because a few layers
// re-throw this error as text; remove that once they keep the code.
const cacheQuotaFailure = (error) =>
  ['CACHE_SESSION_BUDGET', 'session-disk-budget'].includes(error?.code) ||
  /Workbook session disk budget/.test(String(error));
const sharedSourceQualificationFailure = (error) =>
  /Shared formula expansion (?:minimum byte budget|byte budget|would orphan a follower)|Invalid shared formula master|Inconsistent shared formula follower|Extended shared formula attributes|Duplicate or excessive shared groups|Overlapping shared formula master ranges/.test(
    String(error)
  );
const protectSharedSource = (view, error) => {
  const reason =
    'Stored formula results; recalculate in Microsoft Excel after editing. ' + String(error.message ?? error);
  view.structuralReason = reason;
  view.calculationPolicy = preservationPolicy(view.calculationPolicy, reason);
};
const PUBLICATION_CHECKPOINT_LIMIT = 256 * 1024 ** 2;
function publicationCheckpointFields(checkpoint) {
  if (
    !checkpoint ||
    typeof checkpoint !== 'object' ||
    Array.isArray(checkpoint) ||
    Object.keys(checkpoint).sort().join(',') !== 'bytes,path,revision' ||
    typeof checkpoint.path !== 'string' ||
    !path.isAbsolute(checkpoint.path) ||
    checkpoint.path.includes('\0') ||
    typeof checkpoint.revision !== 'string' ||
    !/^[a-f0-9]{64}$/.test(checkpoint.revision) ||
    !Number.isSafeInteger(checkpoint.bytes) ||
    checkpoint.bytes < 1 ||
    checkpoint.bytes > PUBLICATION_CHECKPOINT_LIMIT
  )
    throw refusal('internal-invalid-field', 'Invalid publication calculation checkpoint');
  return checkpoint;
}
function decodeCell(cell, revision) {
  if (cell.calculation && revision) {
    cell.calculation = { ...cell.calculation, revision };
    delete cell.calculation.nativePartial;
  }
  for (const key of ['value', 'cachedValue'])
    if (cell[key] && typeof cell[key] === 'object' && own(cell[key], '$workbookNumber'))
      cell[key] = Number(cell[key].$workbookNumber);
  if (
    cell.calculation?.value &&
    typeof cell.calculation.value === 'object' &&
    own(cell.calculation.value, '$workbookNumber')
  )
    cell.calculation.value = Number(cell.calculation.value.$workbookNumber);
  if (
    cell.cachedResult?.value &&
    typeof cell.cachedResult.value === 'object' &&
    own(cell.cachedResult.value, '$workbookNumber')
  )
    cell.cachedResult.value = Number(cell.cachedResult.value.$workbookNumber);
  return cell;
}
export class IndexedWorkbook {
  id = randomUUID();
  #cacheIdentity;
  #cacheOptions;
  #cacheMode = 'all';
  #native;
  #legacy;
  #calculate;
  #context;
  #root;
  #states = new Map();
  #sequence = 0;
  #closed = false;
  #work = Promise.resolve();
  #roots = new Set();
  #abort = new AbortController();
  #diagnostics = {
    calculationCalls: 0,
    lastProjectedBytes: 0,
    peakProjectedBytes: 0,
    lastInputCells: 0,
    peakInputCells: 0,
    lastEvaluatedFormulaCells: 0,
    lastExportStyleCount: 0,
    lastExportMetadataBytes: 0,
  };
  #traceKey = traceCursorKey();
  #previews = new Map();
  #previewIds = new Set();
  #publicationRebases = new Map();
  #publicationEpoch = 0;
  #initialCalculationPolicy;
  constructor(native, legacy, calculate, context = {}) {
    this.#context = context;
    this.#native = native;
    this.#legacy = (source, commands, output) => legacy(source, commands, output, this.#abort.signal);
    this.#calculate = async (input) => {
      if (input.view.calculationPolicy?.preserveSource) {
        if (input.command || input.calculationIntent)
          throw refusal(
            'session-protected',
            'Calculation is protected; preserved formulas must be recalculated in Microsoft Excel'
          );
        const deferred = await this.#native({
          action: 'workbook',
          operation: 'calculation-policy',
          method: 'preserve',
          root: input.root,
          version: input.version,
          view: input.view,
        });
        return { version: deferred.version, calculationPolicy: preservationPolicy(deferred.policy), skipped: true };
      }
      if (!input.command && input.view.calculationPolicy?.mode === 'manual' && !input.calculationIntent) {
        const deferred = await this.#native({
          action: 'workbook',
          operation: 'calculation-policy',
          method: 'defer',
          root: input.root,
          version: input.version,
          view: input.view,
          changedRanges: input.changedRanges,
        });
        return { version: deferred.version, calculationPolicy: deferred.policy, skipped: true, mode: 'manual-pending' };
      }
      this.#diagnostics.calculationCalls++;
      const result = await calculate(input, this.#abort.signal),
        count = (value) => (Number.isFinite(value) && value >= 0 ? value : 0);
      this.#diagnostics.lastProjectedBytes = count(result.projectedBytes);
      this.#diagnostics.peakProjectedBytes = Math.max(
        this.#diagnostics.peakProjectedBytes,
        this.#diagnostics.lastProjectedBytes
      );
      this.#diagnostics.lastInputCells = count(result.inputCells);
      this.#diagnostics.peakInputCells = Math.max(this.#diagnostics.peakInputCells, this.#diagnostics.lastInputCells);
      this.#diagnostics.lastEvaluatedFormulaCells = count(result.evaluatedFormulaCells);
      if (!input.command) {
        const settled = await this.#native({
          action: 'workbook',
          operation: 'calculation-policy',
          method: 'settle',
          root: input.root,
          version: result.version,
          view: input.view,
          ...(input.calculationSheet ? { sheet: input.calculationSheet, plan: input.calculationPlan } : {}),
        });
        result.calculationPolicy = settled.policy;
      }
      return result;
    };
  }
  /** Scalar ownership counters; projected bytes are serialized accounting, not retained heap. */
  diagnosticStats() {
    return {
      ...this.#diagnostics,
      retainedStates: this.#states.size,
      storageRoots: this.#roots.size,
      closed: Number(this.#closed),
    };
  }
  open(source, options = {}) {
    const initial = options.initialCalculationPolicy;
    if (
      initial !== undefined &&
      (!initial ||
        typeof initial !== 'object' ||
        Array.isArray(initial) ||
        Object.keys(initial).sort().join(',') !== 'mode,recalculateBeforeSave' ||
        !['auto', 'manual'].includes(initial.mode) ||
        typeof initial.recalculateBeforeSave !== 'boolean')
    )
      return Promise.reject(refusal('internal-invalid-field', 'Invalid initial calculation policy'));
    this.#initialCalculationPolicy = options.initialCalculationPolicy;
    try {
      if (options.savedCalculationCheckpoint !== undefined)
        publicationCheckpointFields(options.savedCalculationCheckpoint);
    } catch (error) {
      return Promise.reject(error);
    }
    if (this.#closed) return Promise.reject(refusal('session-closed', 'Workbook session is closed'));
    const work = this.#open(source, options);
    this.#work = work.then(
      () => {},
      () => {}
    );
    return work;
  }
  async #open(source, options) {
    this.#cacheMode = options.cacheMode ?? 'all';
    if (!['all', 'saved-only'].includes(this.#cacheMode))
      throw refusal('session-malformed', 'Invalid workbook index cache mode');
    this.#cacheIdentity = options.cacheIdentity
      ? `xlsx-b16-preserved-calculation-editing:${options.cacheIdentity}:${JSON.stringify(this.#initialCalculationPolicy ?? null)}`
      : null;
    this.#cacheOptions = { ...options.cacheOptions, signal: this.#abort.signal };
    this.#root = await createIndexedStorage(this.id, options.maxSessionBytes, options.storageDirectory);
    try {
      const root = path.join(this.#root, 'base');
      await fs.mkdir(root);
      this.#roots.add(root);
      let cache = null,
        cached = null,
        initial = null;
      if (this.#cacheIdentity && !options.savedCalculationCheckpoint) {
        try {
          cache = await indexedCache(source, this.#cacheIdentity, this.#cacheOptions);
          cached = await cache.restore(root, options.maxSessionBytes ?? 4 * 1024 ** 3);
        } catch (error) {
          if (this.#abort.signal.aborted || error.code === 'CACHE_SESSION_BUDGET') throw error;
          cache = null;
        }
      }
      if (cached) {
        try {
          const index = JSON.parse(await fs.readFile(path.join(root, 'index.json'), 'utf8'));
          if (matchesFormat(index.sourceFactSchema, 'SCHEMA_USED_RANGE_FACTS_V2'))
            await resetIndexedSnapshotExtents(root, cached.version, index.revision, this.#native);
          if (cached.checkpoint) initial = await restoreIndexedCheckpoint(root, cached);
        } catch (error) {
          this.#abort.signal.throwIfAborted();
          if (cacheQuotaFailure(error)) throw error;
          await cache.invalidate();
          await fs.rm(root, { recursive: true, force: true });
          await fs.mkdir(root);
          cached = null;
        }
      }
      if (!cached) {
        this.#roots.add(root);
        await this.#native({ action: 'workbook', operation: 'copy-source', root, source });
      }
      initial ??= await this.#index(root, [], cached, undefined, { publication: !!options.savedCalculationCheckpoint });
      if (options.savedCalculationCheckpoint)
        initial = await this.#publicationCalculation(initial, undefined, options.savedCalculationCheckpoint);
      this.#roots.add(root);
      this.cacheHit = !!cached;
      if (cache && !cached && this.#cacheMode === 'all') {
        if (initial.index.revision !== cache.revision)
          throw refusal('session-stale', 'Workbook changed while preparing cache');
        try {
          const proof = await indexedCalculationReusable(initial, this.#native);
          const snapshot = matchesFormat(initial.index.sourceFactSchema, 'SCHEMA_USED_RANGE_FACTS_V2')
            ? {
                workingBytes: (await fs.stat(path.join(root, 'cells.sqlite'))).size * 2,
                prepareSnapshot: async (clone) => {
                  const cancel = () => {
                    void fs.writeFile(path.join(clone, 'cancelled'), '1').catch(() => {});
                  };
                  this.#abort.signal.throwIfAborted();
                  this.#abort.signal.addEventListener('abort', cancel, { once: true });
                  try {
                    await resetIndexedSnapshotExtents(clone, initial.version, initial.index.revision, this.#native);
                    this.#abort.signal.throwIfAborted();
                  } finally {
                    this.#abort.signal.removeEventListener('abort', cancel);
                  }
                },
              }
            : {};
          await cache.publish(root, initial.version, { reuseCalculation: proof.qualified, ...snapshot });
        } catch (error) {
          this.#abort.signal.throwIfAborted();
          if (cacheQuotaFailure(error)) throw error;
          this.cachePublication = { published: false, reason: 'Optional index cache publication failed' };
        }
      }
      this.#states.set('0', initial);
      this.#initialCalculationPolicy = undefined;
      return this.view('0');
    } catch (error) {
      this.#closed = true;
      if (this.#root && !this.#abort.signal.aborted) {
        await fs.rm(this.#root, { recursive: true, force: true });
        this.#roots.clear();
      }
      throw error;
    }
  }
  async #index(root, _journal, cached, inheritedPolicy, options = {}) {
    this.#roots.add(root);
    if (this.#closed) throw refusal('session-closed', 'Workbook session is closed');
    const index = cached
      ? JSON.parse(await fs.readFile(path.join(root, 'index.json'), 'utf8'))
      : await this.#native({ action: 'workbook', operation: 'index', root });
    const metadataBytes = await fs.readFile(path.join(root, 'metadata.xlsx'));
    // The engine answers what the workbook contains; this layer decides what to allow.
    const answer = await this.#native({ action: 'workbook', operation: 'workbook-view', root });
    const { view, styles } = viewFromEngine(answer, index.revision);
    await this.#native({
      action: 'workbook',
      operation: 'artifact-write',
      root,
      name: 'display-formats.json',
      data: JSON.stringify(styles.map((s) => s.format ?? null)),
    });
    for (const sheet of view.sheets) {
      const metadata = index.sheets.find((item) => item.path === sheet.path);
      if (!metadata) throw refusal('session-bad-input', 'The native index did not contain a workbook sheet');
      if (metadata.background) {
        const value = metadata.background;
        if (
          !matchesFormat(value.schema, 'SCHEMA_SHEET_BACKGROUND_V1') ||
          value.sheet !== sheet.path ||
          typeof value.canEdit !== 'boolean' ||
          (value.revision !== null && !/^[a-f0-9]{64}$/.test(value.revision)) ||
          (value.sample !== null && !['graph-paper', 'soft-dots'].includes(value.sample)) ||
          (value.canEdit && (value.sample === null || value.revision === null))
        )
          throw refusal('internal-invalid-field', 'Invalid native background metadata.');
        sheet.background = structuredClone(value);
      }
      Object.assign(sheet, {
        rows: metadata.rows,
        columns: metadata.columns,
        hiddenRows: metadata.hiddenRows,
        visibleRows: metadata.visibleRows ?? [],
        rowHeights: metadata.rowHeights,
        storedCells: metadata.storedCells,
        storedFormulaCount: metadata.formulaCount,
      });
    }
    indexedViewGuards(view, answer, index.sheets);
    view.pivots = indexedPivotInventory(metadataBytes, view).map(({ model, parts, targetPath, ...definition }) => ({
      ...definition,
      ...(model ? { fieldNames: model.fields.map((field) => field.name) } : {}),
    }));
    try {
      const policy = await this.#native({
        action: 'workbook',
        operation: 'calculation-policy',
        method: 'initialize',
        root,
        version: cached?.reuseCalculation ? cached.version : 0,
        view,
        ...((inheritedPolicy ?? this.#initialCalculationPolicy)
          ? {
              initialPolicy: {
                mode: (inheritedPolicy ?? this.#initialCalculationPolicy).mode,
                recalculateBeforeSave: (inheritedPolicy ?? this.#initialCalculationPolicy).recalculateBeforeSave,
              },
            }
          : {}),
      });
      view.calculationPolicy = policy.policy ?? policy;
      // Normal Automatic import initializes its graph during evaluation. A saved
      // checkpoint must create that graph without running the evaluator.
      if (
        options.publication &&
        !inheritedPolicy &&
        !this.#initialCalculationPolicy &&
        view.calculationPolicy.mode === 'auto' &&
        !view.calculationPolicy.protectedSource
      ) {
        const initialized = await this.#native({
          action: 'workbook',
          operation: 'calculation-policy',
          method: 'initialize',
          root,
          version: 0,
          view,
          initialPolicy: {
            mode: view.calculationPolicy.mode,
            recalculateBeforeSave: view.calculationPolicy.recalculateBeforeSave,
          },
        });
        view.calculationPolicy = initialized.policy ?? initialized;
      }
    } catch (error) {
      this.#abort.signal.throwIfAborted();
      if (!sharedSourceQualificationFailure(error)) throw error;
      // An inherited application mode initializes dependencies before the
      // evaluator. That transaction has rolled back; status reads the original
      // policy without requesting an unqualified shared dependency graph.
      view.calculationPolicy = await this.#native({
        action: 'workbook',
        operation: 'calculation-policy',
        method: 'status',
        root,
        version: cached?.reuseCalculation ? cached.version : 0,
        view,
      });
      this.#abort.signal.throwIfAborted();
      protectSharedSource(view, error);
    }
    if (index.sheets.some((sheet) => sheet.protectedFormulaCaches > 0) || view.calculationPolicy.protectedSource)
      view.calculationPolicy = preservationPolicy(view.calculationPolicy);
    let version = cached?.reuseCalculation ? cached.version : 0,
      spillOwnership = false;
    if (!options.publication && view.calculationPolicy.mode === 'manual' && !view.calculationPolicy.protectedSource) {
      const adopted = await this.#native({
        action: 'workbook',
        operation: 'spill-stage',
        root,
        version,
        revision: index.revision,
        view: empty(view),
        storedOnly: true,
      });
      version = adopted.version;
      spillOwnership = !!adopted.groups;
    }
    if (
      !options.publication &&
      (!cached || !cached.reuseCalculation) &&
      view.sheets.some((sheet) => sheet.storedFormulaCount > 0) &&
      view.calculationPolicy.mode !== 'manual' &&
      !view.calculationPolicy.protectedSource
    ) {
      try {
        const result = await this.#calculate({ root, version, view: empty(view), revision: index.revision });
        version = result.version;
        view.sheets = withSpillDimensions(view, result).sheets;
        view.calculationPolicy = result.calculationPolicy;
        spillOwnership = !!result.spillDimensions;
      } catch (error) {
        // Unsupported/malformed shared source geometry cannot own a calculated rewrite.
        // Keep the original package accessible; edits/export refresh remain fenced.
        this.#abort.signal.throwIfAborted();
        if (!sharedSourceQualificationFailure(error)) throw error;
        protectSharedSource(view, error);
      }
    }
    if (view.calculationPolicy.preserveSource && !view.readOnly) {
      const kept = await this.#native({
        action: 'workbook',
        operation: 'calculation-policy',
        method: 'preserve',
        root,
        version,
        view,
      });
      view.calculationPolicy = preservationPolicy(kept.policy, view.calculationPolicy.reason);
      for (const sheet of view.sheets)
        sheet.capabilities = {
          ...sheet.capabilities,
          insertRows: false,
          deleteRows: false,
          referenceColumns: false,
          insertRowsReason: view.calculationPolicy.reason,
        };
    }
    return {
      root,
      index,
      view: empty(view),
      styles,
      version,
      journal: null,
      spillOwnership,
      managedCharts: indexedObjectHasManagedCharts(metadataBytes),
    };
  }
  #rewriteCheck(candidate) {
    this.#abort.signal.throwIfAborted();
    if (this.#state(candidate.state) !== candidate.anchor || !this.#roots.has(candidate.previous.root))
      throw refusal('session-stale', 'Private rewrite accepted source changed');
  }
  /** Allocate the exact target before native export seals its rewrite input.
   * Force controls export behavior; only v2 internal purpose grants receipts. */
  async #rewriteSource(state, force = true, working, privatePrefix) {
    const anchor = this.#state(state),
      previous = working ?? anchor;
    const next = privatePrefix ? null : String(++this.#sequence),
      root = path.join(this.#root, privatePrefix ? privatePrefix + '-' + randomUUID() : `version-${next}`);
    const candidate = {
      state: String(state),
      anchor,
      previous,
      next,
      root,
      enabled: matchesFormat(previous.index.sourceFactSchema, 'SCHEMA_USED_RANGE_FACTS_V2'),
    };
    await fs.mkdir(root, { mode: 0o700 });
    this.#roots.add(root);
    try {
      this.#rewriteCheck(candidate);
      const source = await this.#export(state, force, previous, false, candidate.enabled ? root : undefined);
      this.#rewriteCheck(candidate);
      if (candidate.enabled && !uuidV4.test(source.rewriteInputId))
        throw refusal('internal-missing-field', 'Missing native private rewrite input');
      return { ...source, candidate };
    } catch (error) {
      await fs.rm(root, { recursive: true, force: true });
      this.#roots.delete(root);
      throw error;
    }
  }
  /**
   * Excel stores a formula that was filled down a column once and lets the
   * other cells point at it. Inserting or deleting rows changes each of those
   * cells differently, so the groups are first written out in full, as Excel
   * itself does. Array and data-table formulas cannot be treated that way.
   */
  async #expandShared(current, refusalMessage, sheetName) {
    const sheets = (current.logicalSheets ?? current.index.sheets).filter((s) => s.specialFormulaCount || s.extended);
    if (!sheets.length) return current;
    // A sheet with extension records (form controls, embedded objects, newer
    // rule kinds) is left alone unless it is the edited sheet or those records
    // mention the edited sheet, which this writer could not then keep correct.
    const extended = sheets.filter((s) => s.extended),
      target = current.view.sheets.find((s) => s.name === sheetName);
    if (extended.length) {
      if (!target || extended.some((s) => s.path === target.path)) throw refusal('session-refused', refusalMessage);
      const parts = unzipSync(await fs.readFile(path.join(current.root, 'metadata.xlsx')), {
          filter: (file) => extended.some((s) => s.path === file.name),
        }),
        escaped = sheetName.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
      for (const sheet of extended) {
        const text = strFromU8(parts[sheet.path] ?? new Uint8Array()),
          // An empty <extLst/> (WPS writes one inside its filters) holds nothing.
          blocks = text.match(
            /<(?:\w+:)?extLst\s*\/>|<((?:\w+:)?extLst|(?:\w+:)?AlternateContent)[\s>][\s\S]*?<\/\1>/g
          );
        if (!blocks || blocks.some((block) => block.includes(sheetName) || block.includes(escaped)))
          throw refusal('session-refused', refusalMessage);
      }
    }
    if (sheets.some((s) => s.specialFormulaCount)) {
      const kinds = await this.#native({ action: 'workbook', operation: 'special-kinds', root: current.root });
      if (kinds?.notShared !== 0) throw refusal('session-refused', refusalMessage);
    } else return current;
    const normalized = await this.#native({
      action: 'workbook',
      operation: 'normalize-shared',
      root: current.root,
      version: current.version,
      ranges: sheets.filter((s) => s.specialFormulaCount).map((s) => ({ sheet: s.path, range: 'A1:XFD1048576' })),
    });
    return normalized.version === current.version ? current : { ...current, version: normalized.version };
  }
  #rewriteFields(source) {
    return source.candidate.enabled ? { rewriteInputId: source.rewriteInputId } : {};
  }
  /** Index truthful rewritten bytes, then fork only their extent authority.
   * Selected-pivot certification runs before the extent fork; later patches and
   * calculation descend from the returned actual native version. */
  async #reindexRewrite(source, result, policy, afterIndex, preserveCalculation = false) {
    const candidate = source.candidate,
      { root, previous } = candidate;
    try {
      this.#rewriteCheck(candidate);
      if (
        !result ||
        typeof result.file !== 'string' ||
        path.basename(result.file) !== result.file ||
        !/^[a-zA-Z0-9_-][a-zA-Z0-9._-]*\.xlsx$/.test(result.file)
      )
        throw refusal('internal-invalid-field', 'Invalid native private rewrite output');
      if (candidate.enabled && !uuidV4.test(result.rewriteId))
        throw refusal('internal-missing-field', 'Missing native private rewrite receipt');
      await this.#native({
        action: 'workbook',
        operation: 'copy-source',
        root,
        source: path.join(previous.root, result.file),
      });
      this.#rewriteCheck(candidate);
      let indexed = await this.#index(root, [], undefined, policy, preserveCalculation ? { publication: true } : {});
      this.#rewriteCheck(candidate);
      if (afterIndex) {
        indexed = await afterIndex(indexed);
        this.#rewriteCheck(candidate);
      }
      if (candidate.enabled) {
        const version = indexed.version,
          transferred = await this.#native({
            action: 'workbook',
            operation: 'used-range-transfer',
            root,
            version,
            sourceRevision: indexed.index.revision,
            previousRoot: previous.root,
            previousVersion: previous.version,
            previousSourceRevision: previous.index.revision,
            rewriteId: result.rewriteId,
          });
        this.#rewriteCheck(candidate);
        if (
          !transferred ||
          typeof transferred !== 'object' ||
          Array.isArray(transferred) ||
          Object.keys(transferred).sort().join(',') !==
            'incompleteSheets,previousSourceRevision,previousVersion,rewriteId,schema,sourceRevision,transferredSheets,version' ||
          !matchesFormat(transferred.schema, 'SCHEMA_USED_RANGE_TRANSFER_V1') ||
          transferred.sourceRevision !== indexed.index.revision ||
          transferred.previousSourceRevision !== previous.index.revision ||
          transferred.previousVersion !== previous.version ||
          transferred.rewriteId !== result.rewriteId ||
          !Number.isSafeInteger(transferred.version) ||
          transferred.version < 0 ||
          transferred.version === version ||
          ['transferredSheets', 'incompleteSheets'].some(
            (key) =>
              !Number.isSafeInteger(transferred[key]) ||
              transferred[key] < 0 ||
              transferred[key] > indexed.view.sheets.length
          )
        )
          throw refusal('internal-invalid-field', 'Invalid native private extent transfer');
        indexed = { ...indexed, version: transferred.version };
      }
      this.#rewriteCheck(candidate);
      return indexed;
    } catch (error) {
      await fs.rm(root, { recursive: true, force: true });
      this.#roots.delete(root);
      throw error;
    }
  }
  /** Only a native clear-link writer receipt can carry unchanged authority
   * across this metadata rewrite; source XML caches grant no new authority. */
  async #clearHyperlinksCalculation(indexed, previous, rewriteId, purpose = 'clear-hyperlinks') {
    const result = await this.#native({
      action: 'workbook',
      operation: 'calculation-transfer',
      purpose,
      root: indexed.root,
      version: indexed.version,
      sourceRevision: indexed.index.revision,
      previousRoot: previous.root,
      previousVersion: previous.version,
      rewriteId,
      view: indexed.view,
    });
    const fields =
      'policy,schema,transferredDerivedCells,transferredFormulaCells,transferredSpillGroups,version,witness';
    if (
      !result ||
      typeof result !== 'object' ||
      Array.isArray(result) ||
      Object.keys(result).sort().join(',') !== fields ||
      result.schema !==
        (purpose === 'note-authoring'
          ? FORMAT.SCHEMA_NOTE_CALCULATION_TRANSFER_V1
          : purpose === 'sheet-background'
            ? FORMAT.SCHEMA_BACKGROUND_CALCULATION_TRANSFER_V1
            : purpose === 'page-theme'
              ? FORMAT.SCHEMA_PAGE_THEME_CALCULATION_TRANSFER_V1
              : purpose === 'hyperlink-authoring'
                ? FORMAT.SCHEMA_HYPERLINK_CALCULATION_TRANSFER_V1
                : FORMAT.SCHEMA_CLEAR_HYPERLINKS_CALCULATION_TRANSFER_V1) ||
      !Number.isSafeInteger(result.version) ||
      result.version < 0 ||
      result.version === indexed.version ||
      !result.policy ||
      typeof result.policy !== 'object' ||
      Array.isArray(result.policy) ||
      ['transferredFormulaCells', 'transferredDerivedCells', 'transferredSpillGroups'].some(
        (key) => !Number.isSafeInteger(result[key]) || result[key] < 0
      ) ||
      ['mode', 'recalculateBeforeSave', 'pendingFormulaCells', 'pendingCacheObjects'].some(
        (key) => (result.policy[key] ?? 0) !== (previous.view.calculationPolicy[key] ?? 0)
      )
    )
      throw refusal('internal-invalid-field', 'Invalid Clear Hyperlinks calculation-preservation receipt.');
    const witness = result.witness;
    if (
      !witness ||
      typeof witness !== 'object' ||
      Array.isArray(witness) ||
      Object.keys(witness).sort().join(',') !== 'bytes,file,revision,schema' ||
      witness.schema !== 1 ||
      witness.file !== 'publication-calculation.sqlite'
    )
      throw refusal('internal-invalid-field', 'Invalid Clear Hyperlinks calculation checkpoint.');
    const checkpoint = Object.freeze({
      path: path.join(indexed.root, witness.file),
      revision: witness.revision,
      bytes: witness.bytes,
    });
    await this.#verifyPublicationCheckpoint(checkpoint);
    return {
      ...indexed,
      version: result.version,
      view: { ...indexed.view, calculationPolicy: result.policy },
      spillOwnership: result.transferredSpillGroups > 0,
      publicationCheckpoint: checkpoint,
    };
  }
  #state(state) {
    if (this.#closed) throw refusal('session-closed', 'Workbook session is closed');
    const value = this.#states.get(String(state));
    if (!value) throw refusal('internal-unknown-reference', 'Unknown workbook session revision');
    return value;
  }
  view(state) {
    const current = this.#state(state);
    const revision = current.publicRevision ?? current.index.revision;
    return {
      ...empty(current.view),
      revision: state === '0' ? revision : `${revision}:draft:${this.id}:${state}`,
      storage: { kind: 'rust-indexed-v1', sessionId: this.id, state: String(state) },
    };
  }
  async read(state, sheetName, range) {
    return this.#read(this.#state(state), state, sheetName, range);
  }
  async #read(current, state, sheetName, range, calculationInput = false) {
    const sheet = current.view.sheets.find((sheet) => sheet.name === sheetName);
    if (!sheet) throw refusal('internal-unknown-reference', 'Unknown worksheet');
    const bounds = rangeBounds(range);
    if (bounds.rows * bounds.columns > 10000)
      throw refusal('session-limit', 'Read ranges in pages of at most 10,000 positions');
    const cells = {},
      revision = this.view(state).revision;
    let offset = 0;
    do {
      const result = await this.#native({
        action: 'workbook',
        operation: 'range',
        root: current.root,
        version: current.version,
        sheet: sheet.path,
        range,
        offset,
        limit: 2048,
        ...(calculationInput ? { calculationInput: true } : {}),
      });
      this.#state(state);
      for (const [ref, record] of result.cells) {
        const cell = decodeCell(record, revision);
        cells[ref] = { ...(current.styles[cell.styleId ?? 0] ?? { wrapText: false }), ...cell };
      }
      offset = result.nextOffset;
    } while (offset !== null);
    return { sheet: sheetName, range, cells };
  }
  async scan(state, sheetName, after, options) {
    if (
      options !== undefined &&
      (!options ||
        typeof options !== 'object' ||
        Array.isArray(options) ||
        Object.keys(options).some((key) => key !== 'limit'))
    )
      throw refusal('internal-invalid-field', 'Invalid worksheet scan options');
    const limit = options?.limit === undefined ? 2048 : options.limit;
    if (!Number.isInteger(limit) || limit < 1 || limit > 2048)
      throw refusal('internal-invalid-field', 'Invalid worksheet scan page limit');
    const current = this.#state(state),
      sheet = current.view.sheets.find((sheet) => sheet.name === sheetName);
    if (!sheet) throw refusal('internal-unknown-reference', 'Unknown worksheet');
    let pageLimit = limit;
    for (;;) {
      let result;
      try {
        result = await this.#native({
          action: 'workbook',
          operation: 'range',
          root: current.root,
          version: current.version,
          sheet: sheet.path,
          range: 'A1:XFD1048576',
          after,
          limit: pageLimit,
        });
      } catch (error) {
        // Only an explicitly bounded readonly page may retry this exact
        // recoverable transport refusal; mutations and other failures propagate.
        if (
          options === undefined ||
          pageLimit === 1 ||
          error?.message !== 'Native workbook reply exceeds transport bound'
        )
          throw error;
        this.#state(state);
        pageLimit = Math.max(1, Math.floor(pageLimit / 2));
        continue;
      }
      this.#state(state);
      const cells = {},
        revision = this.view(state).revision;
      for (const [ref, record] of result.cells) {
        const cell = decodeCell(record, revision);
        cells[ref] = { ...(current.styles[cell.styleId ?? 0] ?? { wrapText: false }), ...cell };
      }
      const page = { cells, nextCursor: result.nextCursor };
      if (options !== undefined) {
        // Account for the real exceptional-number owner encoding and its JSON
        // framing. A count limit alone cannot bound long literal cell records.
        let exceptional = false;
        const encodedCells = JSON.stringify(cells, (_key, item) => {
          if (typeof item === 'number' && (!Number.isFinite(item) || Object.is(item, -0))) {
            exceptional = true;
            return { $workbookNumber: Object.is(item, -0) ? '-0' : String(item) };
          }
          return item;
        });
        const bytes = Buffer.byteLength(
          exceptional ? JSON.stringify({ encodedCells, nextCursor: page.nextCursor }) : JSON.stringify(page)
        );
        if (bytes > 1024 * 1024) {
          if (pageLimit === 1)
            throw refusal('session-limit', 'A single worksheet cell exceeds the bounded scan reply size');
          pageLimit = Math.max(1, Math.floor(pageLimit / 2));
          continue;
        }
      }
      return page;
    }
  }
  async trace(state, sheetName, cell, options) {
    const current = this.#state(state),
      request = traceRequest(sheetName, cell, options);
    if (!current.view.sheets.some((s) => s.name === sheetName))
      throw refusal('internal-unknown-reference', 'Unknown trace worksheet');
    const binding = {
        ownerId: this.id,
        state: String(state),
        revision: this.view(state).revision,
        version: current.version,
        sheet: sheetName,
        cell,
        direction: request.direction,
        sheets: current.view.sheets.map((s) => s.name),
      },
      cursor = readTraceCursor(this.#traceKey, options.after, binding);
    this.#abort.signal.throwIfAborted();
    const result = await this.#native({
      action: 'workbook',
      operation: 'trace',
      root: current.root,
      version: current.version,
      view: current.view,
      sheet: sheetName,
      cell,
      direction: request.direction,
      cursor: cursor.cursor,
      limit: request.limit,
    });
    this.#abort.signal.throwIfAborted();
    if (this.#state(state) !== current)
      throw refusal('session-stale', 'Formula trace owner changed its accepted state');
    return traceReceipt(this.#traceKey, binding, cursor, result, request.limit);
  }
  async usedRange(state, sheetName) {
    usedRangeSheet(sheetName);
    const current = this.#state(state),
      sheet = current.view.sheets.find((item) => item.name === sheetName);
    if (!sheet) throw refusal('internal-unknown-reference', 'Unknown UsedRange worksheet');
    const expected = {
      sheet: sheetName,
      native: { version: current.version, sourceRevision: current.index.revision, sheetPath: sheet.path },
    };
    this.#abort.signal.throwIfAborted();
    const raw = await this.#native({
      action: 'workbook',
      operation: 'used-range',
      root: current.root,
      version: current.version,
      sourceRevision: current.index.revision,
      sheet: sheet.path,
    });
    this.#abort.signal.throwIfAborted();
    if (this.#state(state) !== current) throw refusal('session-stale', 'UsedRange owner changed its accepted state');
    return normalizeNativeUsedRange(raw, expected);
  }
  async noteMembership(state, sheetName) {
    noteMembershipSheet(sheetName);
    const current = this.#state(state),
      sheet = current.view.sheets.find((item) => item.name === sheetName);
    if (!sheet) throw refusal('internal-unknown-reference', 'Unknown Notes membership worksheet');
    const source = { version: current.version, sourceRevision: current.index.revision, sheetPath: sheet.path },
      expected = { sheet: sheetName, source };
    const check = () => {
      this.#abort.signal.throwIfAborted();
      if (this.#state(state) !== current)
        throw refusal('session-stale', 'Notes membership owner changed its accepted state');
    };
    check();
    const raw = await indexedNoteMembership({
      metadataFile: path.join(current.root, 'metadata.xlsx'),
      view: current.view,
      sheet: sheetName,
      source,
      signal: this.#abort.signal,
      check,
    });
    check();
    return noteMembershipInner(raw, expected);
  }
  async duplicatePreview(state, sheet, range, hasHeaders, columns, offset = 0) {
    const current = this.#state(state),
      config = duplicateConfiguration(sheet, range, hasHeaders, columns),
      target = current.view.sheets.find((s) => s.name === sheet);
    if (!target) throw refusal('internal-unknown-reference', 'Unknown duplicate worksheet');
    if (!Number.isSafeInteger(offset) || offset < 0 || offset > 1000000)
      throw refusal('internal-invalid-field', 'Invalid duplicate preview offset');
    const raw = await this.#native({
      action: 'workbook',
      operation: 'dedup-preview',
      root: current.root,
      version: current.version,
      sheet: target.path,
      range,
      hasHeaders,
      columns: config.columns,
      offset,
      limit: 100,
    });
    this.#state(state);
    duplicatePage(raw, config, offset);
    return bindDuplicatePreview(
      raw,
      { documentId: this.id, ownerId: this.id, state, revision: this.view(state).revision },
      config,
      offset
    );
  }
  previewFormula(state, sheet, cell, formula, options) {
    const current = this.#state(state);
    formulaPreviewTarget(sheet, cell, formula);
    if (
      !options ||
      typeof options !== 'object' ||
      Array.isArray(options) ||
      Object.keys(options).join(',') !== 'requestId'
    )
      return Promise.reject(refusal('internal-invalid-field', 'Invalid formula preview options'));
    const requestId = formulaPreviewRequestId(options.requestId),
      expected = { requestId, sheet, cell, formula };
    if (!current.view.sheets.some((item) => item.name === sheet))
      return Promise.reject(refusal('internal-unknown-reference', 'Unknown preview worksheet'));
    if (this.#previewIds.has(requestId))
      return Promise.reject(refusal('session-refused', 'Formula preview request ID has already been used'));
    if (this.#previewIds.size >= 4096)
      return Promise.reject(refusal('session-limit', 'Formula preview request lifetime bound; reopen this workbook'));
    if (this.#previews.size)
      return Promise.reject(refusal('session-busy', 'A formula preview request is still stopping'));
    this.#previewIds.add(requestId);
    const request = { current, state: String(state), cancelled: false, started: false };
    this.#previews.set(requestId, request);
    const cancelled = () => ({
      ...expected,
      schema: 1,
      authority: 'draft-preview',
      result: { status: 'cancelled', reason: 'Draft preview request cancelled' },
    });
    const work = this.#work
      .then(async () => {
        this.#abort.signal.throwIfAborted();
        if (this.#state(state) !== current)
          throw refusal('session-stale', 'Formula preview owner changed its accepted state');
        if (request.cancelled) return cancelled();
        request.started = true;
        let reply = await this.#native({
          action: 'workbook',
          operation: 'preview-formula',
          root: current.root,
          version: current.version,
          view: current.view,
          ...expected,
        });
        this.#abort.signal.throwIfAborted();
        if (this.#state(state) !== current)
          throw refusal('session-stale', 'Formula preview owner changed its accepted state');
        if (request.cancelled) return cancelled();
        if (
          reply?.result?.type === 'number' &&
          reply.result.value &&
          typeof reply.result.value === 'object' &&
          Object.keys(reply.result.value).join(',') === '$workbookNumber' &&
          reply.result.value.$workbookNumber === '-0'
        )
          reply = { ...reply, result: { ...reply.result, value: -0 } };
        return formulaPreviewInner(reply, expected);
      })
      .finally(async () => {
        await request.cancelWork;
        await fs.rm(path.join(current.root, 'preview-cancel-' + requestId), { force: true });
        this.#previews.delete(requestId);
      });
    this.#work = work.then(
      () => {},
      () => {}
    );
    return work;
  }
  async cancelFormulaPreview(state, requestId) {
    const current = this.#state(state);
    formulaPreviewRequestId(requestId);
    const request = this.#previews.get(requestId);
    if (!request || request.current !== current || request.state !== String(state))
      return { schema: 1, requestId, cancelled: false };
    request.cancelled = true;
    if (request.started) {
      request.cancelWork ??= fs
        .writeFile(path.join(current.root, 'preview-cancel-' + requestId), '1', { flag: 'wx' })
        .catch((error) => {
          if (error.code !== 'EEXIST' && error.code !== 'ENOENT') throw error;
        });
      await request.cancelWork;
    }
    return { schema: 1, requestId, cancelled: true };
  }
  async accessibility(state) {
    const current = this.#state(state),
      check = () => {
        this.#abort.signal.throwIfAborted();
        if (this.#state(state) !== current) throw refusal('session-stale', 'Accessibility state retired.');
      };
    check();
    const bytes = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
    check();
    return indexedAccessibility(bytes, current.view, {
      readRange: (sheet, range) => this.#read(current, state, sheet, range),
      check,
    });
  }
  async objects(state, sheetName, options = {}) {
    const current = this.#state(state);
    try {
      const bytes = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
      this.#state(state);
      const result = indexedObjectList(bytes, current.view, sheetName, options);
      this.#state(state);
      return result;
    } catch (error) {
      if (error?.code && !isRefusal(error))
        throw refusal('session-unsupported', 'Object data is unavailable; reload this workbook.');
      throw error;
    }
  }
  async object(state, sheetName, objectId) {
    const current = this.#state(state);
    try {
      const metadataBytes = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
      this.#state(state);
      const result = await indexedObjectPreview(
        {
          metadataBytes,
          view: current.view,
          readRange: (sheet, range) => this.#read(current, state, sheet, range),
          readMedia: (part, { maxBytes }) =>
            boundedObjectMedia(path.join(current.root, 'source.xlsx'), part, maxBytes, this.#abort.signal),
        },
        sheetName,
        objectId
      );
      this.#state(state);
      return result;
    } catch (error) {
      if (error?.code && !isRefusal(error))
        throw refusal('session-unsupported', 'Object data is unavailable; reload this workbook.');
      throw error;
    }
  }
  async printDefinition(state, sheetName) {
    const current = this.#state(state),
      metadataBytes = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
    this.#state(state);
    return indexedPrintDefinition(metadataBytes, current.view, sheetName);
  }
  async #ruleDefinitions(current) {
    if (!current.ruleDefinitions) {
      const bytes = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
      this.#abort.signal.throwIfAborted();
      const conditional = {},
        validation = {};
      for (const sheet of current.view.sheets) {
        conditional[sheet.name] = indexedRulesDefinition(bytes, current.view, sheet.name, 'conditional');
        validation[sheet.name] = indexedRulesDefinition(bytes, current.view, sheet.name, 'validation');
      }
      current.ruleDefinitions = { conditional, validation };
    }
    return current.ruleDefinitions;
  }
  async conditionalRules(state, sheetName) {
    const current = this.#state(state),
      definitions = await this.#ruleDefinitions(current);
    this.#state(state);
    const found = definitions.conditional[sheetName];
    if (!found) throw refusal('internal-unknown-reference', 'Unknown worksheet');
    const result = structuredClone(found);
    try {
      const duplicates = await indexedRuleDuplicates(
        result.rules,
        sheetName,
        (sheet, range) => this.#read(current, state, sheet, range),
        () => {
          this.#abort.signal.throwIfAborted();
          this.#state(state);
        }
      );
      this.#state(state);
      return { ...result, duplicates };
    } catch (error) {
      this.#abort.signal.throwIfAborted();
      this.#state(state);
      return { ...result, duplicates: {}, evaluationReasons: [String(error.message ?? error)] };
    }
  }
  async validationRules(state, sheetName) {
    const current = this.#state(state),
      definitions = await this.#ruleDefinitions(current);
    this.#state(state);
    const found = definitions.validation[sheetName];
    if (!found) throw refusal('internal-unknown-reference', 'Unknown worksheet');
    return structuredClone(found);
  }
  async printSnapshot(state, sheetName, settings, geometryProfile) {
    const current = this.#state(state),
      sheet = current.view.sheets.find((item) => item.name === sheetName);
    if (!sheet) throw refusal('internal-unknown-reference', 'Unknown worksheet');
    const check = () => {
      this.#abort.signal.throwIfAborted();
      this.#state(state);
    };
    // Native format admission proves the entire current worksheet's special
    // masters before exposing any bounded page, including off-page producers.
    const metadataBytes = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
    check();
    return createPrintSnapshot({
      sessionId: this.id,
      state,
      revision: this.view(state).revision,
      metadataBytes,
      view: current.view,
      sheet: sheetName,
      settings,
      geometryProfile,
      check,
      readRange: (sheet, range) => this.#read(current, state, sheet, range),
      objects: (sheet, options) => this.objects(state, sheet, options),
      object: (sheet, id) => this.object(state, sheet, id),
      qualifyCells: async (range) => {
        let proof;
        try {
          proof = await this.#native({
            action: 'workbook',
            operation: 'print-check',
            root: current.root,
            version: current.version,
            sheet: sheet.path,
            range,
            view: current.view,
          });
        } catch (error) {
          const code = /\b(PRINT_[A-Z_]+)\b/.exec(String(error.message ?? error))?.[1];
          if (code) error.code = code;
          throw error;
        }
        check();
        if (proof?.version !== current.version || proof?.sheet !== sheet.path || proof?.range !== range)
          throw refusal('session-stale', 'Print ownership proof changed its accepted version or page');
        return proof;
      },
    });
  }
  async query(state, sheetName, query, mode, after, options = {}) {
    const current = this.#state(state),
      sheet = current.view.sheets.find((s) => s.name === sheetName);
    if (!sheet) throw refusal('internal-unknown-reference', 'Unknown worksheet');
    const range = options.range ?? 'A1:XFD1048576',
      matchLimit = options.matchLimit;
    rangeBounds(range);
    if (matchLimit !== undefined && (!Number.isInteger(matchLimit) || matchLimit < 1 || matchLimit > 1000))
      throw refusal('internal-invalid-field', 'Invalid search match limit');
    const result = await this.#native({
      action: 'workbook',
      operation: 'query',
      root: current.root,
      version: current.version,
      sheet: sheet.path,
      query,
      mode,
      after,
      range,
      matchLimit,
      date1904: !!sheet.date1904,
    });
    this.#state(state);
    const cells = {},
      revision = this.view(state).revision;
    for (const [ref, record] of result.cells) {
      const cell = decodeCell(record, revision);
      cells[ref] = { ...(current.styles[cell.styleId ?? 0] ?? { wrapText: false }), ...cell };
    }
    return { cells, matches: result.matches, scanned: result.scanned, nextCursor: result.nextCursor };
  }
  async notes(state, sheetName) {
    const current = this.#state(state);
    if (!current.view.sheets.some((s) => s.name === sheetName))
      throw refusal('internal-unknown-reference', 'Unknown note worksheet.');
    const raw = await this.#native({
      action: 'workbook',
      operation: 'notes',
      root: current.root,
      version: current.version,
      sourceRevision: current.index.revision,
      sheet: sheetName,
    });
    this.#state(state);
    return noteInventory(raw, { sheet: sheetName });
  }
  async hyperlink(state, sheetName, cell) {
    const current = this.#state(state);
    hyperlinkCell(cell);
    if (!current.view.sheets.some((s) => s.name === sheetName))
      throw refusal('internal-unknown-reference', 'Unknown hyperlink worksheet.');
    const result = await this.#native({
      action: 'workbook',
      operation: 'hyperlink',
      root: current.root,
      version: current.version,
      sourceRevision: current.index.revision,
      sheet: sheetName,
      cell,
    });
    this.#state(state);
    return result;
  }
  async selectionSummary(state, sheetName, range) {
    const current = this.#state(state),
      sheet = current.view.sheets.find((s) => s.name === sheetName);
    if (!sheet) throw refusal('internal-unknown-reference', 'Unknown worksheet');
    const result = await this.#native({
      action: 'workbook',
      operation: 'selection',
      root: current.root,
      version: current.version,
      sheet: sheet.path,
      range,
    });
    this.#state(state);
    return result;
  }
  /** Read-only, first-seen typed choices. Native TEMP storage owns grouping;
   * every page rescans the exact accepted state, including offscreen values. */
  async filterValues(state, sheetName, range, column, query = '', offset = 0, limit = 100) {
    const current = this.#state(state),
      sheet = current.view.sheets.find((s) => s.name === sheetName);
    if (!sheet) throw refusal('internal-unknown-reference', 'Unknown worksheet');
    const bounds = rangeBounds(range);
    if (typeof column !== 'string' || !/^[A-Z]{1,3}$/.test(column))
      throw refusal('internal-invalid-field', 'Invalid filter column');
    const col = cellPosition(column + '1').col;
    if (col < bounds.first.col || col > bounds.last.col)
      throw refusal('session-unsupported', 'Filter column is outside the requested range');
    if (typeof query !== 'string' || Buffer.byteLength(query) > 32768)
      throw refusal('internal-invalid-field', 'Invalid filter search text');
    if (!Number.isSafeInteger(offset) || offset < 0) throw refusal('internal-invalid-field', 'Invalid filter offset');
    if (!Number.isInteger(limit) || limit < 1 || limit > 100)
      throw refusal('internal-invalid-field', 'Invalid filter page limit');
    const result = await this.#native({
      action: 'workbook',
      operation: 'filter-values',
      root: current.root,
      version: current.version,
      sheet: sheet.path,
      range,
      column,
      query,
      offset,
      limit,
    });
    this.#state(state);
    for (const item of result.values)
      if (item.value && typeof item.value === 'object' && own(item.value, '$workbookNumber'))
        item.value = Number(item.value.$workbookNumber);
    return result;
  }
  async keys(state, sheetName, range, column, values, options = {}) {
    return this.#keys(this.#state(state), sheetName, range, column, values, options);
  }
  async #keys(current, sheetName, range, column, values, options) {
    const sheet = current.view.sheets.find((s) => s.name === sheetName);
    if (!sheet) throw refusal('internal-unknown-reference', 'Unknown worksheet');
    return this.#native({
      action: 'workbook',
      operation: 'keys',
      root: current.root,
      version: current.version,
      sheet: sheet.path,
      range,
      column,
      values: nativeValue(values),
      ...options,
    });
  }
  async #keyedCommands(current, command) {
    if (!Array.isArray(command.updates) || !command.updates.length)
      throw refusal('session-bad-input', 'Keyed updates must not be empty');
    const keys = command.updates.map((u) => u.key);
    if (new Set(keys).size !== keys.length) throw refusal('session-duplicate', 'Duplicate requested key');
    const matched = await this.#keys(current, command.sheet, command.range, command.keyColumn, keys, {
      unique: true,
      limit: 0,
    });
    const bounds = rangeBounds(command.range),
      commands = [];
    for (const [index, update] of command.updates.entries()) {
      if (!matched.matches[index]) throw refusal('internal-missing-field', `Missing key: ${String(update.key)}`);
      if (!Array.isArray(update.cells) || !update.cells.length)
        throw refusal('session-bad-input', 'Keyed update cells must not be empty');
      for (const { column, ...cell } of update.cells) {
        if (!own(cell, 'expectedValue') && !own(cell, 'expectedFormula'))
          throw refusal('session-bad-input', 'Keyed updates require expected old values or formulas');
        const ref = `${column}${matched.firstRows[index] + 1}`,
          position = cellPosition(ref);
        if (position.col < bounds.first.col || position.col > bounds.last.col)
          throw refusal('session-unsupported', `Update is outside the keyed range: ${ref}`);
        commands.push({ ...cell, sheet: command.sheet, cell: ref });
      }
    }
    return commands;
  }
  async calculationSummary(state) {
    const current = this.#state(state),
      counts = { calculated: 0, unresolved: 0, error: 0, circular: 0 };
    for (const sheet of current.view.sheets)
      if (sheet.storedFormulaCount) {
        const summary = await this.#native({
          action: 'workbook',
          operation: 'summary',
          root: current.root,
          version: current.version,
          sheet: sheet.path,
        });
        for (const key of Object.keys(counts)) counts[key] += summary[key];
      }
    this.#state(state);
    return { revision: this.view(state).revision, engine: 'rust-sqlite + univer-oss/0.25.2', ...counts };
  }
  async affected(before, after) {
    const original = this.#state(before),
      current = this.#state(after);
    if (original.root !== current.root) {
      let total = 0;
      const affected = [];
      const names = new Set([...original.view.sheets, ...current.view.sheets].map((s) => s.name));
      for (const name of names) {
        const a = original.view.sheets.find((s) => s.name === name),
          b = current.view.sheets.find((s) => s.name === name);
        let cursor;
        do {
          const page = await this.#native({
            action: 'workbook',
            operation: 'diff-bases',
            root: current.root,
            beforeBase: path.basename(original.root),
            before: original.version,
            version: current.version,
            beforeSheet: a?.path,
            sheet: b?.path,
            beforeStyles: original.styles,
            styles: current.styles,
            after: cursor,
          });
          total += page.count;
          for (const change of page.changes) if (affected.length < 100) affected.push({ ...change, sheet: name });
          cursor = page.nextCursor;
        } while (cursor);
      }
      this.#state(after);
      return { affected, total };
    }
    let offset = 0,
      total = 0;
    const affected = [];
    do {
      const page = await this.#native({
        action: 'workbook',
        operation: 'diff',
        root: current.root,
        before: original.version,
        version: current.version,
        styles: current.styles,
        offset,
      });
      total += page.count;
      for (const change of page.changes)
        if (affected.length < 100)
          affected.push({ ...change, sheet: current.view.sheets.find((s) => s.path === change.sheet).name });
      offset = page.nextOffset;
    } while (offset !== null);
    this.#state(after);
    return { affected, total };
  }
  async #inputs(state, commands, current = this.#state(state)) {
    const view = empty(current.view),
      windows = new Map();
    const include = (sheet, range) => windows.set(`${sheet}!${range}`, { sheet, range });
    const includeCell = (sheet, cell) => {
      const p = cellPosition(cell),
        row = Math.floor(p.row / 32) * 32,
        col = Math.floor(p.col / 32) * 32;
      include(
        sheet,
        `${columnName(col)}${row + 1}:${columnName(Math.min(16383, col + 31))}${Math.min(1048576, row + 32)}`
      );
    };
    for (const command of commands) {
      if (command.cell) includeCell(command.sheet, command.cell);
      for (const cell of command.cells ?? []) includeCell(command.sheet, cell.cell);
      for (const range of [command.range, command.source])
        if (range) {
          const bounds = rangeBounds(range);
          if (bounds.rows * bounds.columns > 10000) return null;
          include(command.sheet, range);
        }
      if (command.operation || ['fillRange', 'pasteFormulas'].includes(command.type)) return null;
    }
    for (const window of windows.values()) {
      const result = await this.#read(current, state, window.sheet, window.range, true);
      Object.assign(view.sheets.find((sheet) => sheet.name === window.sheet).cells, result.cells);
    }
    return view;
  }
  apply(state, commands) {
    if (
      Array.isArray(commands) &&
      commands.some((c) => ['createChart', 'updateChart', 'deleteObject'].includes(c?.operation))
    ) {
      try {
        commands = structuredClone(commands);
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (Array.isArray(commands) && commands.some(isNoteCommand)) {
      try {
        commands = captureNoteCommands(commands);
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (Array.isArray(commands) && commands.some(isModel3DCommand)) {
      try {
        commands = captureModel3DCommands(commands);
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (Array.isArray(commands) && commands.some(isSmartArtCommand)) {
      try {
        commands = captureSmartArtCommands(commands);
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (Array.isArray(commands) && commands.some(isInkStrokeCommand)) {
      try {
        commands = captureInkStrokeCommands(commands);
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (Array.isArray(commands) && commands.some((c) => isShapeCommand(c) || c?.operation === 'deleteObject')) {
      try {
        commands = captureShapeCommands(commands);
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (Array.isArray(commands) && commands.some(isIndexedPivotCommand)) {
      try {
        commands = capturePivotCommands(commands);
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (Array.isArray(commands) && commands.some(isIndexedTableCommand)) {
      try {
        commands = structuredClone(commands);
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (
      Array.isArray(commands) &&
      commands.some((c) => ['setSheetBackground', 'removeSheetBackground'].includes(c?.type))
    ) {
      try {
        commands = normalizeBackgroundCommands(commands);
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (Array.isArray(commands) && commands.some((c) => c?.type === 'applyPageTheme')) {
      try {
        commands = normalizePageThemeCommands(commands);
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (Array.isArray(commands) && commands.some((c) => c?.type === 'booleanCell')) {
      try {
        commands = normalizeBooleanCellCommands(commands);
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (Array.isArray(commands) && commands.some((c) => ['setHyperlink', 'removeHyperlink'].includes(c?.type))) {
      try {
        commands = normalizeHyperlinkCommands(commands);
      } catch (error) {
        return Promise.reject(error);
      }
    }
    // Capture the closed source/target action before this owner queues work.
    if (Array.isArray(commands) && commands.some((command) => command?.type === 'fillAcrossSheets')) {
      try {
        commands = normalizeFillAcrossSheetsRuns(structuredClone(commands));
      } catch (error) {
        return Promise.reject(error);
      }
    }
    if (Array.isArray(commands) && commands.some((command) => command?.type === 'clearHyperlinks')) {
      try {
        commands = normalizeClearHyperlinksCommands(structuredClone(commands));
      } catch (error) {
        return Promise.reject(error);
      }
    }
    const work = this.#work.then(async () => {
      this.#publicationEpoch++;
      await this.#retirePublicationRebases();
      return this.#mutate(state, commands);
    });
    const guarded = work.catch(async (error) => {
      const accepted = new Map();
      for (const state of this.#states.values()) {
        if (!accepted.has(state.root)) accepted.set(state.root, new Set());
        accepted.get(state.root).add(state.version);
      }
      for (const root of this.#roots)
        if (!this.#closed && !accepted.has(root)) {
          await fs.rm(root, { recursive: true, force: true });
          this.#roots.delete(root);
        }
      // Any late failure can leave private same-root SQL versions, including a
      // later name/object guard after an earlier cell patch in an ordered batch.
      if (!this.#closed)
        for (const [root, versions] of accepted) {
          try {
            await this.#native({ action: 'workbook', operation: 'compact', root, keep: [...versions] });
          } catch {
            /* Preserve the original failure; accepted states and recovery remain intact. */
          }
        }
      throw error;
    });
    this.#work = guarded.catch(() => {});
    return guarded;
  }
  /** What a command module may use: the native engine and this session's private state, under plain names. */
  #commandContext(state, current) {
    return {
      state,
      current,
      id: this.id,
      signal: this.#abort.signal,
      context: this.#context,
      native: (request) => this.#native(request),
      calculate: (input) => this.#calculate(input),
      // Throws if this revision is no longer the accepted one (or the session closed).
      recheck: () => this.#state(state),
      view: (next) => this.view(next),
      object: (sheet, objectId) => this.object(state, sheet, objectId),
      mutate: (commands) => this.#mutate(state, commands),
      rewriteSource: (force, working, privatePrefix) => this.#rewriteSource(state, force, working, privatePrefix),
      rewriteFields: (source) => this.#rewriteFields(source),
      rewriteCheck: (candidate) => this.#rewriteCheck(candidate),
      reindexRewrite: (...args) => this.#reindexRewrite(...args),
      clearHyperlinksCalculation: (...args) => this.#clearHyperlinksCalculation(...args),
      expandShared: (...args) => this.#expandShared(...args),
      read: (from, sheet, range, ...rest) => this.#read(from, state, sheet, range, ...rest),
      keyedCommands: (from, command) => this.#keyedCommands(from, command),
      inputs: (commands, working) => this.#inputs(state, commands, working),
      loadRuleDefinitions: () => this.#ruleDefinitions(current),
      finishDerived: (working) => this.#finishDerived(working, state),
      finishSelectedPivots: (working, sheet, objects) => this.#finishSelectedPivots(working, state, sheet, objects),
      finishCharts: (working, sheet) => this.#finishCharts(working, state, sheet),
      publish: (next, value) => this.#states.set(next, value),
      nextState: () => String(++this.#sequence),
      peekNextState: () => this.#sequence + 1,
      dropRoot: (root) => this.#roots.delete(root),
      isClosed: () => this.#closed,
      acceptedStates: () => this.#states.values(),
    };
  }
  async #mutate(state, commands) {
    if (
      Array.isArray(commands) &&
      commands.some(
        (command) =>
          command?.type === 'recalculate' &&
          (Array.isArray(command) || Object.keys(command).length !== 1 || !Object.hasOwn(command, 'type'))
      )
    )
      throw refusal(
        'session-unsupported',
        'Recalculate requires exactly {type: "recalculate"}; scoped calculation and mode options are unavailable'
      );
    if (Array.isArray(commands) && commands.length !== 1 && commands.some((command) => command.type === 'importCells'))
      throw refusal('session-bad-input', 'Staged paste must be one complete descriptor command');
    if (
      Array.isArray(commands) &&
      commands.length > 1 &&
      commands.some(
        (c) =>
          c.operation ||
          [
            'setSheetBackground',
            'removeSheetBackground',
            'applyPageTheme',
            'booleanCell',
            'setHyperlink',
            'removeHyperlink',
            'clearHyperlinks',
            'recalculate',
            'calculateSheet',
            'setCalculationOptions',
          ].includes(c.type)
      )
    ) {
      return this.#mutateBatch(state, commands);
    }
    const current = this.#state(state);
    if (Array.isArray(commands) && commands.some((command) => command?.type === 'fillAcrossSheets'))
      commands = normalizeFillAcrossSheetsRuns(commands, current.view);
    if (current.index.sheets.some((s) => s.dataStart === null))
      throw refusal('internal-missing-field', 'Worksheet is missing required sheetData; editing is not qualified');
    if (current.view.readOnly) throw refusal('session-protected', 'Workbook is read-only');
    // Copying a sheet in Excel adds a code module for it to the macro project. The
    // project stays byte for byte as it was, so the copy would share the code name
    // of the sheet it came from.
    if (current.view.macroProject && Array.isArray(commands) && commands.some((c) => c?.operation === 'duplicateSheet'))
      throw refusal(
        'macro-protected',
        'This workbook has macros. Copying a sheet would change them, so the editor leaves it to Excel'
      );
    // A protected sheet refuses changes, as it does in Excel; other sheets do
    // not. Renaming is allowed: sheet protection does not cover the name.
    const locked = current.view.protectedSheets ?? [];
    if (locked.length && Array.isArray(commands))
      for (const command of commands) {
        const touched = [command?.sheet, command?.sourceSheet, command?.targetSheet, ...(command?.sheets ?? [])];
        const name = command?.operation === 'renameSheet' ? undefined : touched.find((sheet) => locked.includes(sheet));
        if (name) throw refusal('sheet-protected', `Sheet “${name}” is protected; unprotect it in Excel to change it`);
      }
    // Changing how a sheet is laid out touches no formula and no stored result,
    // so it is allowed on a workbook whose formulas are kept as stored.
    const layoutOnly =
      Array.isArray(commands) &&
      commands.length > 0 &&
      commands.every((command) =>
        ['resizeColumn', 'resizeRow', 'freeze', 'mergeCells', 'unmergeCells'].includes(command?.operation)
      );
    if (
      current.view.calculationPolicy?.preserveSource &&
      !(commands.length === 1 && commands[0].operation === 'addSheet') &&
      !layoutOnly
    ) {
      for (const footprint of assertPreservationCommands(commands)) {
        const sheet = current.view.sheets.find((s) => s.name === footprint.sheet);
        if (!sheet) throw refusal('internal-unknown-reference', 'Unknown worksheet');
        await this.#native({
          action: 'workbook',
          operation: 'preservation-point-check',
          root: current.root,
          version: current.version,
          sheet: sheet.path,
          range: footprint.range,
          content: footprint.content,
        });
      }
    } else if (current.view.calculationPolicy?.protectedSource && !current.view.calculationPolicy?.preserveSource)
      throw refusal(
        'session-protected',
        'Workbook calculation source settings are protected; its stored source is preserved'
      );
    if (!Array.isArray(commands) || !commands.length)
      throw refusal('internal-missing-field', 'Missing workbook commands');
    const ctx = this.#commandContext(state, current);
    const direct = routeCommands(directModules, commands);
    if (direct) return direct.run(ctx, commands);
    await admitAgainstMetadata(ctx, commands);
    const owner = routeCommands(guardedModules, commands);
    if (owner) return owner.run(ctx, commands);
    throw refusal(
      'session-unsupported',
      'This command has no qualified indexed implementation; accepted workbook is unchanged'
    );
  }
  /** A batch with a special command runs one command at a time on a private candidate, then publishes the whole batch. */
  async #mutateBatch(state, commands) {
    // A private candidate owner keeps intermediate revisions unreachable by
    // reads/undo. Publish only the complete batch, preserving ordered replay.
    this.#state(state);
    const candidate = new IndexedWorkbook(this.#native, this.#legacy, this.#calculate, this.#context);
    candidate.id = this.id;
    candidate.#root = this.#root;
    candidate.#states = new Map(this.#states);
    candidate.#sequence = this.#sequence;
    candidate.#roots = this.#roots;
    candidate.#abort = this.#abort;
    candidate.#cacheIdentity = this.#cacheIdentity;
    candidate.#cacheOptions = this.#cacheOptions;
    candidate.#cacheMode = this.#cacheMode;
    const oldRoots = new Set(this.#roots),
      edits = [],
      durableCommands = [];
    let next = state,
      normalized = false;
    try {
      for (const command of commands) {
        this.#state(state);
        const result = await candidate.#mutate(next, [command]);
        next = result.view.storage.state;
        edits.push(...result.edits);
        durableCommands.push(...(result.commands ?? [structuredClone(command)]));
        normalized ||= !!result.commands;
      }
      const accepted = candidate.#state(next);
      this.#state(state);
      for (const root of this.#roots)
        if (!oldRoots.has(root) && root !== accepted.root) {
          await fs.rm(root, { recursive: true, force: true });
          this.#roots.delete(root);
        }
      this.#state(state);
      this.#states.set(next, accepted);
      return {
        view: this.view(next),
        edits,
        ...(next === state &&
        commands.every((command) =>
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
        )
          ? { unchanged: true }
          : {}),
        ...(normalized ? { commands: durableCommands } : {}),
      };
    } finally {
      this.#sequence = candidate.#sequence;
    }
  }
  export(state, options = {}) {
    const work = this.#work.then(() => {
      if (
        !options ||
        typeof options !== 'object' ||
        Array.isArray(options) ||
        Object.keys(options).some((key) => key !== 'publication') ||
        (options.publication !== undefined && typeof options.publication !== 'boolean')
      )
        throw refusal('internal-invalid-field', 'Invalid workbook export options');
      return this.#export(state, false, undefined, options.publication === true);
    });
    this.#work = work.then(
      () => {},
      () => {}
    );
    return work;
  }
  async #verifyPublicationCheckpoint(checkpoint) {
    publicationCheckpointFields(checkpoint);
    const info = await fs.lstat(checkpoint.path);
    if (
      !info.isFile() ||
      info.isSymbolicLink() ||
      info.nlink !== 1 ||
      info.size !== checkpoint.bytes ||
      (await fileDigest(checkpoint.path, { signal: this.#abort.signal })) !== checkpoint.revision
    )
      throw refusal('session-malformed', 'Publication calculation checkpoint is damaged');
    return checkpoint;
  }
  /** Publication imports accepted SQL authority, never evaluates source caches. */
  async #publicationCalculation(indexed, current, checkpoint) {
    if (checkpoint) await this.#verifyPublicationCheckpoint(checkpoint);
    const restoring = !!checkpoint;
    const transferred = await this.#native({
      action: 'workbook',
      operation: 'calculation-transfer',
      purpose: restoring ? 'publication-restore' : 'publication',
      root: indexed.root,
      version: indexed.version,
      sourceRevision: indexed.index.revision,
      view: indexed.view,
      ...(restoring
        ? { witness: checkpoint.path, witnessRevision: checkpoint.revision }
        : { previousRoot: current.root, previousVersion: current.version }),
    });
    const fields =
      'policy,schema,transferredDerivedCells,transferredFormulaCells,transferredSpillGroups,version' +
      (restoring ? '' : ',witness');
    if (
      !transferred ||
      typeof transferred !== 'object' ||
      Array.isArray(transferred) ||
      Object.keys(transferred).sort().join(',') !== fields ||
      stripFormatPrefix(transferred.schema, 'SCHEMA_PUBLICATION_CALCULATION_') !==
        (restoring ? 'restore' : 'transfer') + FORMAT.V1 ||
      !Number.isSafeInteger(transferred.version) ||
      transferred.version < 0 ||
      !transferred.policy ||
      typeof transferred.policy !== 'object' ||
      Array.isArray(transferred.policy) ||
      ['transferredFormulaCells', 'transferredDerivedCells', 'transferredSpillGroups'].some(
        (key) => !Number.isSafeInteger(transferred[key]) || transferred[key] < 0
      )
    )
      throw refusal('internal-invalid-field', 'Invalid native publication calculation transfer');
    let pin = checkpoint;
    if (!restoring) {
      const witness = transferred.witness;
      if (
        !witness ||
        typeof witness !== 'object' ||
        Array.isArray(witness) ||
        Object.keys(witness).sort().join(',') !== 'bytes,file,revision,schema' ||
        witness.schema !== 1 ||
        witness.file !== 'publication-calculation.sqlite'
      )
        throw refusal('internal-invalid-field', 'Invalid native publication calculation witness');
      pin = { path: path.join(indexed.root, witness.file), revision: witness.revision, bytes: witness.bytes };
    }
    await this.#verifyPublicationCheckpoint(pin);
    return {
      ...indexed,
      version: transferred.version,
      view: {
        ...indexed.view,
        calculationPolicy: indexed.view.calculationPolicy?.preserveSource
          ? preservationPolicy(transferred.policy)
          : transferred.policy,
      },
      spillOwnership: transferred.transferredSpillGroups > 0,
      publicationCheckpoint: Object.freeze({ ...pin }),
    };
  }
  /** Prepare the actual published bytes without changing any accepted state. */
  preparePublicationRebase(state, source, expectedRevision, options = {}) {
    const work = this.#work.then(async () => {
      const current = this.#state(state),
        epoch = this.#publicationEpoch;
      if (
        !options ||
        typeof options !== 'object' ||
        Array.isArray(options) ||
        Object.keys(options).some((key) => key !== 'calculationCheckpoint')
      )
        throw refusal('internal-invalid-field', 'Invalid publication rebase options');
      if (options.calculationCheckpoint !== undefined)
        await this.#verifyPublicationCheckpoint(options.calculationCheckpoint);
      if (
        typeof source !== 'string' ||
        !path.isAbsolute(source) ||
        typeof expectedRevision !== 'string' ||
        !/^[a-f0-9]{64}$/.test(expectedRevision)
      )
        throw refusal('internal-invalid-field', 'Invalid publication rebase source');
      if (this.#publicationRebases.size) throw refusal('session-refused', 'A publication rebase is already prepared');
      const info = await fs.lstat(source);
      if (!info.isFile() || info.isSymbolicLink())
        throw refusal('session-bad-input', 'Publication rebase requires an immutable regular source');
      if ((await fileDigest(source, { signal: this.#abort.signal })) !== expectedRevision)
        throw refusal('internal-mismatch', 'Publication rebase source revision mismatch');
      const id = randomUUID(),
        root = path.join(this.#root, 'publication-' + id);
      await fs.mkdir(root);
      this.#roots.add(root);
      try {
        await this.#native({ action: 'workbook', operation: 'copy-source', root, source });
        const indexed = await this.#publicationCalculation(
          await this.#index(root, [], undefined, current.view.calculationPolicy, { publication: true }),
          current,
          options.calculationCheckpoint
        );
        this.#abort.signal.throwIfAborted();
        if (this.#state(state) !== current || this.#publicationEpoch !== epoch)
          throw refusal('session-stale', 'Publication rebase accepted source changed');
        if (
          indexed.index.revision !== expectedRevision ||
          (await fileDigest(source, { signal: this.#abort.signal })) !== expectedRevision
        )
          throw refusal('internal-mismatch', 'Publication rebase source revision mismatch');
        const token = Object.freeze({ id, sourceRevision: expectedRevision });
        this.#publicationRebases.set(id, { state: String(state), current, epoch, indexed, token });
        return token;
      } catch (error) {
        await fs.rm(root, { recursive: true, force: true });
        this.#roots.delete(root);
        throw error;
      }
    });
    this.#work = work.then(
      () => {},
      () => {}
    );
    return work;
  }
  /** Only a trusted owner that confirmed publication may adopt this token. */
  #publicationRebase(token) {
    this.#abort.signal.throwIfAborted();
    if (
      !token ||
      typeof token !== 'object' ||
      Array.isArray(token) ||
      Object.keys(token).sort().join(',') !== 'id,sourceRevision'
    )
      throw refusal('internal-invalid-field', 'Invalid publication rebase token');
    const prepared = this.#publicationRebases.get(token.id);
    if (
      !prepared ||
      token.sourceRevision !== prepared.token.sourceRevision ||
      this.#state(prepared.state) !== prepared.current ||
      this.#publicationEpoch !== prepared.epoch
    )
      throw refusal('session-stale', 'Publication rebase is unavailable or retired');
    return prepared;
  }
  assertPublicationRebase(token) {
    this.#publicationRebase(token);
  }
  publicationCheckpoint(token) {
    return { ...this.#publicationRebase(token).indexed.publicationCheckpoint };
  }
  adoptPublicationRebase(token) {
    const prepared = this.#publicationRebase(token);
    const state = String(++this.#sequence);
    this.#states.set(state, prepared.indexed);
    this.#publicationRebases.delete(token.id);
    return { view: this.view(state) };
  }
  discardPublicationRebase(token) {
    const work = this.#work.then(async () => {
      if (!token || typeof token.id !== 'string')
        throw refusal('internal-invalid-field', 'Invalid publication rebase token');
      const prepared = this.#publicationRebases.get(token.id);
      if (!prepared) return;
      if (token.sourceRevision !== prepared.token.sourceRevision)
        throw refusal('internal-invalid-field', 'Invalid publication rebase revision');
      this.#publicationRebases.delete(token.id);
      await fs.rm(prepared.indexed.root, { recursive: true, force: true });
      this.#roots.delete(prepared.indexed.root);
    });
    this.#work = work.then(
      () => {},
      () => {}
    );
    return work;
  }
  async #retirePublicationRebases() {
    const prepared = [...this.#publicationRebases.values()];
    this.#publicationRebases.clear();
    for (const candidate of prepared) {
      await fs.rm(candidate.indexed.root, { recursive: true, force: true });
      this.#roots.delete(candidate.indexed.root);
    }
  }
  async #finishSelectedPivots(working, state, sheet, objects) {
    for (const object of objects) {
      if (!matchesFormat(object, 'XL_PIVOTTABLES_PIVOT1_XML'))
        throw refusal('session-unsupported', 'Selected pivot source profile is not qualified');
      const pivot = working.view.pivots?.find((p) => p.canEdit && p.revision);
      if (!pivot) throw refusal('internal-missing-field', 'Required selected pivot is protected or missing');
      const source = await this.#rewriteSource(state, true, working, 'selected-pivot');
      this.#state(state);
      const prepared = await indexedPivotPlan({
        metadataBytes: await fs.readFile(path.join(working.root, 'metadata.xlsx')),
        view: working.view,
        command: {
          operation: 'refreshPivot',
          sheet: pivot.sheet,
          name: pivot.name,
          expectedPivotRevision: pivot.revision,
        },
        readRange: (name, range) => this.#read(working, state, name, range),
        readWorksheet: (part, { maxBytes }) => boundedObjectMedia(source.path, part, maxBytes, this.#abort.signal),
      });
      await this.#native({
        action: 'workbook',
        operation: 'artifact-write',
        root: working.root,
        name: 'package-metadata.json',
        data: JSON.stringify(prepared.plan),
      });
      const refreshed = await this.#native({
        action: 'workbook',
        operation: 'package-edit',
        root: working.root,
        version: working.version,
        source: path.basename(source.path),
        ...this.#rewriteFields(source),
      });
      const { root } = source.candidate;
      // An internal reindex starts with stored caches. Only the authoritative
      // transfer below can preserve certification from the original SQL owner.
      working = await this.#reindexRewrite(
        source,
        refreshed,
        { mode: 'manual', recalculateBeforeSave: working.view.calculationPolicy.recalculateBeforeSave },
        async (indexed) => {
          const transferred = await this.#native({
            action: 'workbook',
            operation: 'calculation-transfer',
            root,
            version: indexed.version,
            previousRoot: working.root,
            previousVersion: working.version,
            sheet,
            view: indexed.view,
          });
          return {
            ...indexed,
            version: transferred.version,
            view: { ...indexed.view, calculationPolicy: transferred.policy },
          };
        }
      );
    }
    const scope = await this.#native({
      action: 'workbook',
      operation: 'calculation-policy',
      method: 'scope',
      root: working.root,
      version: working.version,
      view: working.view,
      sheet,
    });
    const calculated = await this.#calculate({
      root: working.root,
      version: working.version,
      view: working.view,
      calculationIntent: true,
      calculationSheet: sheet,
      calculationPlan: scope.plan,
      revision: working.index.revision + ':selected-pivot',
    });
    working = { ...working, version: calculated.version, view: withSpillDimensions(working.view, calculated) };
    const proofs = [];
    for (const pivot of working.view.pivots.filter((p) => p.canEdit)) {
      const prepared = await indexedPivotPlan({
        metadataBytes: await fs.readFile(path.join(working.root, 'metadata.xlsx')),
        view: working.view,
        command: {
          operation: 'refreshPivot',
          sheet: pivot.sheet,
          name: pivot.name,
          expectedPivotRevision: pivot.revision,
        },
        readRange: (name, range) => this.#read(working, state, name, range),
        readWorksheet: (part, { maxBytes }) =>
          boundedObjectMedia(path.join(working.root, 'source.xlsx'), part, maxBytes, this.#abort.signal),
      });
      proofs.push(prepared.plan);
    }
    const settled = await this.#native({
      action: 'workbook',
      operation: 'calculation-policy',
      method: 'settle-derived',
      kind: 'pivot',
      root: working.root,
      version: working.version,
      view: working.view,
      sheet,
      proofs,
    });
    return { ...working, view: { ...working.view, calculationPolicy: settled.policy } };
  }
  async #finishDerived(current, state) {
    let working = current;
    const pivots = working.view.pivots ?? [];
    if (!pivots.length && working.managedCharts === false && !working.view.calculationPolicy?.pendingCacheObjects)
      return working;
    if (pivots.some((pivot) => !pivot.canEdit))
      throw refusal('session-unsupported', 'Full recalculation cannot refresh an unqualified foreign pivot cache');
    // Package/cache refreshes are private steps of the single recalculate
    // command. They do not publish another history frame or trigger Save.
    for (const pivot of pivots) {
      const source = await this.#rewriteSource(state, true, working, 'calculated');
      this.#state(state);
      const prepared = await indexedPivotPlan({
        metadataBytes: await fs.readFile(path.join(working.root, 'metadata.xlsx')),
        view: working.view,
        command: {
          operation: 'refreshPivot',
          sheet: pivot.sheet,
          name: pivot.name,
          expectedPivotRevision: pivot.revision,
        },
        readRange: (sheet, range) => this.#read(working, state, sheet, range),
        readWorksheet: (part, { maxBytes }) => boundedObjectMedia(source.path, part, maxBytes, this.#abort.signal),
      });
      await this.#native({
        action: 'workbook',
        operation: 'artifact-write',
        root: working.root,
        name: 'package-metadata.json',
        data: JSON.stringify(prepared.plan),
      });
      const result = await this.#native({
        action: 'workbook',
        operation: 'package-edit',
        root: working.root,
        version: working.version,
        source: path.basename(source.path),
        ...this.#rewriteFields(source),
      });
      working = await this.#reindexRewrite(source, result, working.view.calculationPolicy);
    }
    if (pivots.length) {
      const calculated = await this.#calculate({
        root: working.root,
        version: working.version,
        view: working.view,
        calculationIntent: true,
        revision: working.index.revision + ':derived',
      });
      working = {
        ...working,
        version: calculated.version,
        view: withSpillDimensions(working.view, calculated),
        spillOwnership: working.spillOwnership || !!calculated.spillDimensions,
      };
      const proofs = [];
      for (const pivot of working.view.pivots) {
        const prepared = await indexedPivotPlan({
          metadataBytes: await fs.readFile(path.join(working.root, 'metadata.xlsx')),
          view: working.view,
          command: {
            operation: 'refreshPivot',
            sheet: pivot.sheet,
            name: pivot.name,
            expectedPivotRevision: pivot.revision,
          },
          readRange: (sheet, range) => this.#read(working, state, sheet, range),
          readWorksheet: (part, { maxBytes }) =>
            boundedObjectMedia(path.join(working.root, 'source.xlsx'), part, maxBytes, this.#abort.signal),
        });
        proofs.push(prepared.plan);
      }
      const accepted = await this.#native({
        action: 'workbook',
        operation: 'calculation-policy',
        method: 'settle-derived',
        kind: 'pivot',
        root: working.root,
        version: working.version,
        view: working.view,
        proofs,
      });
      working.view = { ...working.view, calculationPolicy: accepted.policy };
    }
    return this.#finishCharts(working, state);
  }
  async #finishCharts(working, state, sheet) {
    const metadataBytes = await fs.readFile(path.join(working.root, 'metadata.xlsx'));
    const proof = await indexedObjectCachePlan({
      metadataBytes,
      view: working.view,
      sheet,
      readRange: (sheet, range) => this.#read(working, state, sheet, range),
    });
    const accepted = await this.#native({
      action: 'workbook',
      operation: 'calculation-policy',
      method: 'settle-derived',
      kind: 'chart',
      root: working.root,
      version: working.version,
      view: working.view,
      ...(sheet === undefined ? {} : { sheet }),
      proof: { managedChartCaches: proof.managedChartCaches ?? {} },
    });
    return { ...working, view: { ...working.view, calculationPolicy: accepted.policy } };
  }
  async #export(state, force = false, working, publication = false, rewriteTargetRoot) {
    const current = working ?? this.#state(state);
    const internalRewrite = rewriteTargetRoot !== undefined;
    if (
      internalRewrite &&
      (publication ||
        !matchesFormat(current.index.sourceFactSchema, 'SCHEMA_USED_RANGE_FACTS_V2') ||
        current.view.readOnly ||
        (current.view.calculationPolicy?.protectedSource && !current.view.calculationPolicy?.preserveSource))
    )
      throw refusal('session-unsupported', 'Private rewrite source is not qualified');
    // Passive read-only viewing may calculate display values privately. Public
    // Copy/export must preserve the original package, including signatures and
    // macro/opaque graphs, before any chart or spill cache refresh is considered.
    if (!internalRewrite && !force && current.view.readOnly)
      return { path: path.join(current.root, 'source.xlsx'), revision: current.index.revision };
    if (
      !internalRewrite &&
      !force &&
      !publication &&
      current.view.calculationPolicy?.preserveSource &&
      !current.journal
    )
      return { path: path.join(current.root, 'source.xlsx'), revision: current.index.revision };
    if (current.view.calculationPolicy?.protectedSource && !current.view.calculationPolicy?.preserveSource) {
      if (current.journal) throw refusal('session-protected', 'Protected calculation settings cannot be rewritten');
      return { path: path.join(current.root, 'source.xlsx'), revision: current.index.revision };
    }
    // Public caches are computed before source/saved-export fast paths. A loaded
    // formula chart can already differ from its stored cache without cell edits.
    // Private mutation export skips this gate, so delete/fix remains available.
    let objectMetadata, objectCaches;
    if (
      !current.view.calculationPolicy?.preserveSource &&
      !force &&
      current.view.calculationPolicy?.mode === 'manual' &&
      current.view.calculationPolicy.pendingCacheObjects
    ) {
      const certified = await this.#native({
        action: 'workbook',
        operation: 'calculation-policy',
        method: 'chart-status',
        root: current.root,
        version: current.version,
        view: current.view,
      });
      if (certified.parts.length) {
        objectMetadata = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
        objectCaches = await indexedObjectCachePlan({
          metadataBytes: objectMetadata,
          view: current.view,
          parts: certified.parts,
          readRange: (sheet, range) => this.#read(current, state, sheet, range),
        });
      }
      this.#state(state);
    }
    if (
      !current.view.calculationPolicy?.preserveSource &&
      !force &&
      (current.view.calculationPolicy?.mode !== 'manual' ||
        (!current.view.calculationPolicy.pendingFormulaCells && !current.view.calculationPolicy.pendingCacheObjects)) &&
      current.managedCharts !== false
    ) {
      objectMetadata = await fs.readFile(path.join(current.root, 'metadata.xlsx'));
      current.managedCharts ??= indexedObjectHasManagedCharts(objectMetadata);
      if (current.managedCharts)
        objectCaches = await indexedObjectCachePlan({
          metadataBytes: objectMetadata,
          view: current.view,
          readRange: (sheet, range) => this.#read(current, state, sheet, range),
        });
      this.#state(state);
    }
    const manual = current.view.calculationPolicy?.mode === 'manual';
    if (
      !internalRewrite &&
      !publication &&
      !manual &&
      !objectCaches &&
      current.savedExport?.version === current.version
    )
      return { path: current.savedExport.path, revision: current.savedExport.revision };
    if (
      !internalRewrite &&
      !publication &&
      !manual &&
      !objectCaches &&
      !current.spillOwnership &&
      !current.journal &&
      (!force || current.version === 0)
    ) {
      let required = false;
      if (matchesFormat(current.index.sourceFactSchema, 'SCHEMA_USED_RANGE_FACTS_V2')) {
        const normalization = await this.#native({
          action: 'workbook',
          operation: 'used-range-export-needed',
          root: current.root,
          version: current.version,
          sourceRevision: current.index.revision,
        });
        if (
          !normalization ||
          typeof normalization !== 'object' ||
          Array.isArray(normalization) ||
          Object.keys(normalization).sort().join(',') !== 'required,schema,sourceRevision,version' ||
          !matchesFormat(normalization.schema, 'SCHEMA_USED_RANGE_EXPORT_NEEDED_V1') ||
          normalization.sourceRevision !== current.index.revision ||
          normalization.version !== current.version ||
          typeof normalization.required !== 'boolean'
        )
          throw refusal('internal-invalid-field', 'Invalid native source normalization plan');
        this.#state(state);
        required = normalization.required;
      }
      if (!required) return { path: path.join(current.root, 'source.xlsx'), revision: current.index.revision };
    }
    const plan = await this.#native({
      action: 'workbook',
      operation: 'export-plan',
      root: current.root,
      version: current.version,
      ...(current.spillOwnership ? { view: current.view } : {}),
    });
    this.#diagnostics.lastExportStyleCount = plan.styleCount;
    if (!plan.special) {
      const metadataBytes = objectMetadata ?? (await fs.readFile(path.join(current.root, 'metadata.xlsx')));
      this.#diagnostics.lastExportMetadataBytes = metadataBytes.byteLength;
      const hasFormulas =
        current.index.sheets.some((sheet) => sheet.formulaCount > 0) ||
        current.view.sheets.some((sheet) => sheet.storedFormulaCount > 0);
      const metadata = indexedCalculationMetadata(
        metadataBytes,
        plan.incomplete ||
          !!current.view.calculationPolicy?.pendingFormulaCells ||
          !!current.view.calculationPolicy?.pendingCacheObjects,
        current.view.calculationPolicy
      );
      if (plan.spills) metadata.spillMetadata = indexedSpillMetadata(metadataBytes, metadata, plan.spills);
      if (objectCaches) metadata.managedChartCaches = objectCaches.managedChartCaches;
      // One synthetic cell per unique style combination, independent of the
      // number of formatted cells. The qualified style-table writer stays shared.
      const styles = [];
      let after = 0;
      while (styles.length < plan.styleCount) {
        const page = await this.#native({
          action: 'workbook',
          operation: 'export-styles',
          root: current.root,
          version: current.version,
          after,
        });
        if (!page.styles.length) throw refusal('session-unsupported', 'Incomplete native export styles');
        styles.push(...page.styles);
        after = page.next;
      }
      const styled = indexedStylePublication(metadataBytes, styles);
      Object.assign(metadata.replacements, styled.replacements);
      metadata.styleIds = styled.styleIds;
      await this.#native({
        action: 'workbook',
        operation: 'artifact-write',
        root: current.root,
        name: 'export-metadata.json',
        data: JSON.stringify(metadata),
      });
      try {
        const result = await this.#native({
          action: 'workbook',
          operation: 'export',
          root: current.root,
          revision: current.index.revision,
          version: current.version,
          calculationMetadata: true,
          ...(current.view.calculationPolicy?.preserveSource ? { preservationExport: true } : {}),
          ...(publication ? { publication: true } : {}),
          ...(internalRewrite ? { internalRewrite: true, rewriteTargetRoot } : {}),
        });
        this.#state(state);
        if (
          internalRewrite &&
          (!result ||
            typeof result !== 'object' ||
            Array.isArray(result) ||
            Object.keys(result).sort().join(',') !== 'file,revision,rewriteInputId' ||
            result.file !== 'export.xlsx' ||
            typeof result.revision !== 'string' ||
            !/^[a-f0-9]{64}$/.test(result.revision) ||
            !uuidV4.test(result.rewriteInputId))
        )
          throw refusal('internal-invalid-field', 'Invalid native private rewrite export');
        const output = {
          path: path.join(current.root, result.file),
          revision: result.revision,
          ...(internalRewrite ? { rewriteInputId: result.rewriteInputId } : {}),
        };
        if (!internalRewrite && !force && this.#cacheIdentity) {
          try {
            this.cachePublication = await publishIndexedCheckpoint({
              source: output.path,
              identity: this.#cacheIdentity,
              current,
              native: this.#native,
              options: this.#cacheOptions,
            });
          } catch (error) {
            this.#abort.signal.throwIfAborted();
            if (cacheQuotaFailure(error)) throw error;
            this.cachePublication = { published: false, reason: 'Optional saved-index cache publication failed' };
          }
        }
        return output;
      } catch (error) {
        if (error?.code !== 'needs-structural-writer') throw error;
      }
    }
    throw refusal(
      'session-unsupported',
      'This export shape is not qualified for native preservation; accepted workbook is unchanged'
    );
  }
  compact(retain, options = {}) {
    const work = this.#work.then(async () => {
      if (this.#closed) throw refusal('session-closed', 'Workbook session is closed');
      this.#publicationEpoch++;
      await this.#retirePublicationRebases();
      if (
        !options ||
        typeof options !== 'object' ||
        Array.isArray(options) ||
        Object.keys(options).some((key) => key !== 'retainBase') ||
        (options.retainBase !== undefined && typeof options.retainBase !== 'boolean')
      )
        throw refusal('internal-invalid-field', 'Invalid workbook compaction options');
      const keep = new Set([...(options.retainBase === false ? [] : ['0']), ...retain.map(String)]);
      if (!keep.size) throw refusal('session-bad-input', 'Keep at least one accepted workbook state');
      for (const state of keep) this.#state(state);
      const roots = new Map();
      for (const state of keep) {
        const current = this.#state(state);
        if (!roots.has(current.root)) roots.set(current.root, new Set());
        roots.get(current.root).add(current.version);
      }
      const reports = [];
      for (const [root, versions] of roots)
        reports.push(await this.#native({ action: 'workbook', operation: 'compact', root, keep: [...versions] }));
      for (const state of this.#states.keys()) if (!keep.has(state)) this.#states.delete(state);
      for (const root of this.#roots)
        if (!roots.has(root)) {
          await fs.rm(root, { recursive: true, force: true });
          this.#roots.delete(root);
        }
      return { compacted: true, reports };
    });
    this.#work = work.catch(() => {});
    return work;
  }
  async close() {
    this.#closed = true;
    this.#abort.abort(refusal('operation-cancelled', 'Workbook operation cancelled'));
    await Promise.all(
      [...this.#roots].map((root) =>
        fs.writeFile(path.join(root, 'cancelled'), '1').catch((error) => {
          if (error.code !== 'ENOENT') throw error;
        })
      )
    );
    await this.#work;
    this.#states.clear();
    this.#publicationRebases.clear();
    if (this.#root) {
      // A terminated JS worker can leave its Rust child finishing the current
      // frame. Keep cancellation markers until that native lease is released.
      const lease = path.join(this.#root, 'native-active.json');
      for (let attempt = 0; ; attempt++) {
        let active;
        try {
          active = JSON.parse(await fs.readFile(lease, 'utf8'));
        } catch (error) {
          if (error.code === 'ENOENT') break;
          if (error instanceof SyntaxError && attempt < 500) {
            await new Promise((resolve) => setTimeout(resolve, 10));
            continue;
          }
          throw error;
        }
        try {
          process.kill(active.pid, 0);
        } catch (error) {
          if (error.code === 'ESRCH') break;
          throw error;
        }
        if (attempt === 500)
          throw refusal(
            'operation-cancelled',
            'Native workbook work is still stopping; its cancellation markers and private files were retained'
          );
        await new Promise((resolve) => setTimeout(resolve, 10));
      }
      await fs.rm(this.#root, { recursive: true, force: true });
      this.#roots.clear();
    }
  }
}
