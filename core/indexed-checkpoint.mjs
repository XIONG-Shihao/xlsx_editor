import { FORMAT, matchesFormat } from './format.mjs';
/** A checkpoint may reuse native spans only for the exact serialized source.
 * Derived calculation authority cannot rebind old authored value/raw/cache fields
 * to a different saved ZIP merely by changing the public revision. */
import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { indexedCache } from './indexed-cache.mjs';
import { viewFromEngine } from './workbook-package.mjs';
import { indexedViewGuards, indexedCalendarGuards } from './indexed-metadata.mjs';
import { refusal } from './refusal-codes.mjs';

const CHECKPOINT_SCHEMA = FORMAT.SCHEMA_EXPORTED_CHECKPOINT_V1;
const MAX_METADATA = 16 * 1024 * 1024;
const hash = /^[a-f0-9]{64}$/;
const FACTS_V2 = FORMAT.SCHEMA_USED_RANGE_FACTS_V2;
/** Fresh source caches retain calculation certification, never editor extent
 * ancestry. Native ownership checks restrict this operation to a private clone
 * or the single restored-source boundary, before that root becomes live. */
export async function resetIndexedSnapshotExtents(root, version, sourceRevision, native) {
  if (
    typeof root !== 'string' ||
    !path.isAbsolute(root) ||
    root.includes('\0') ||
    !Number.isSafeInteger(version) ||
    version < 0 ||
    typeof sourceRevision !== 'string' ||
    !hash.test(sourceRevision)
  )
    throw refusal('session-malformed', 'Invalid source snapshot identity');
  const result = await native({
    action: 'workbook',
    operation: 'used-range-source-snapshot',
    root,
    version,
    sourceRevision,
  });
  if (
    !result ||
    typeof result !== 'object' ||
    Array.isArray(result) ||
    Object.keys(result).sort().join(',') !== 'schema,sourceRevision,version' ||
    !matchesFormat(result.schema, 'SCHEMA_USED_RANGE_SOURCE_SNAPSHOT_V1') ||
    result.version !== version ||
    result.sourceRevision !== sourceRevision
  )
    throw refusal('internal-invalid-field', 'Invalid native source snapshot receipt');
  return result;
}
const graphMetadataIdentity = (view) => {
  indexedCalendarGuards(view);
  return createHash('sha256')
    .update(
      JSON.stringify({
        sheets: view.sheets.map((s) => ({ name: s.name, path: s.path, date1904: !!s.date1904 })),
        names: view.namedRanges,
        tables: view.tables,
      })
    )
    .digest('hex');
};
const metadataIdentity = (view) =>
  createHash('sha256')
    .update(JSON.stringify([graphMetadataIdentity(view), view.calculationPolicy ?? null]))
    .digest('hex');
function validate(checkpoint, manifest) {
  if (
    !matchesFormat(checkpoint.schema, 'SCHEMA_EXPORTED_CHECKPOINT_V1') ||
    !hash.test(checkpoint.publicRevision) ||
    !hash.test(checkpoint.baseRevision) ||
    !Number.isSafeInteger(checkpoint.version) ||
    checkpoint.version < 0 ||
    !Array.isArray(checkpoint.view?.sheets) ||
    !Array.isArray(checkpoint.styles) ||
    !Array.isArray(checkpoint.logicalSheets)
  )
    throw refusal('internal-invalid-field', 'Invalid exported workbook checkpoint');
  if (
    checkpoint.publicRevision !== manifest.revision ||
    checkpoint.version !== manifest.version ||
    checkpoint.policy !== manifest.identity ||
    checkpoint.metadata !== metadataIdentity(checkpoint.view)
  )
    throw refusal('internal-mismatch', 'Exported workbook checkpoint identity mismatch');
  if (checkpoint.view.sheets.some((sheet) => !sheet.cells || Object.keys(sheet.cells).length))
    throw refusal('session-unsupported', 'Workbook checkpoint cannot contain cell projections');
  if (
    checkpoint.logicalSheets.length !== checkpoint.view.sheets.length ||
    checkpoint.view.sheets.some(
      (sheet) =>
        !checkpoint.logicalSheets.some(
          (item) =>
            item.path === sheet.path && Number.isSafeInteger(item.specialFormulaCount) && item.specialFormulaCount >= 0
        )
    )
  )
    throw refusal('internal-invalid-field', 'Invalid checkpoint logical worksheet guards');
  if (checkpoint.baseRevision !== manifest.files.find((file) => file.name === 'source.xlsx')?.sha256)
    throw refusal('internal-mismatch', 'Exported workbook checkpoint base mismatch');
  if (checkpoint.baseRevision !== checkpoint.publicRevision)
    throw refusal(
      'session-bad-input',
      'Exported checkpoint spans and cached source fields do not describe the saved workbook'
    );
}
export async function indexedCalculationReusable(current, native, { exported = false } = {}) {
  if (
    current.view.calculationPolicy?.preserveSource ||
    current.view.calculationPolicy?.mode === 'manual' ||
    current.view.calculationPolicy?.pendingFormulaCells
  )
    return { qualified: false, reason: 'Stored or pending workbook results are not reusable calculation checkpoints' };
  return native({
    action: 'workbook',
    operation: 'dependencies',
    method: 'cache-status',
    root: current.root,
    version: current.version,
    view: current.view,
    metadata: graphMetadataIdentity(current.view),
    exported,
  });
}
export async function publishIndexedCheckpoint({ source, identity, current, native, options = {} }) {
  if (!identity) return { published: false, reason: 'Index cache is disabled' };
  const cache = await indexedCache(source, identity, options);
  // The old export optimization kept an earlier package/index plus value
  // overlays. Those overlays prove calculation results, not the saved package's
  // authored/raw/cache metadata. Reimport rewritten packages before reuse.
  if (cache.revision !== current.index.revision)
    return {
      published: false,
      reason: 'Saved package differs from the indexed source; exact saved cache metadata requires a fresh import',
    };
  const qualification = await indexedCalculationReusable(current, native, { exported: true });
  if (!qualification.qualified) return { published: false, reason: qualification.reason };
  const view = {
    ...current.view,
    sheets: current.view.sheets.map((sheet) => {
      const { storedCells, storedFormulaCount, rows, columns } = qualification.sheets.find(
        (item) => item.path === sheet.path
      );
      return { ...sheet, storedCells, storedFormulaCount, rows, columns };
    }),
  };
  const answer = await native({ action: 'workbook', operation: 'workbook-view', root: current.root });
  indexedViewGuards(view, answer, qualification.sheets, viewFromEngine(answer, current.index.revision).view);
  const checkpoint = {
    schema: CHECKPOINT_SCHEMA,
    publicRevision: cache.revision,
    baseRevision: current.index.revision,
    version: current.version,
    policy: identity,
    metadata: metadataIdentity(view),
    view,
    styles: current.styles,
    logicalSheets: qualification.sheets,
  };
  const bytes = Buffer.byteLength(JSON.stringify(checkpoint));
  if (bytes > MAX_METADATA) return { published: false, reason: 'Checkpoint metadata exceeds its bounded projection' };
  const databaseBytes = (await fs.stat(path.join(current.root, 'cells.sqlite'))).size;
  let compaction;
  const result = await cache.publish(current.root, current.version, {
    checkpoint,
    // Compaction runs only on the private cache clone. SQLite may temporarily
    // need journal and VACUUM copies; reserve them inside the cache's global cap.
    workingBytes: databaseBytes * 2,
    prepareSnapshot: async (root) => {
      const cancel = () => {
        void fs.writeFile(path.join(root, 'cancelled'), '1').catch(() => {});
      };
      options.signal?.throwIfAborted();
      options.signal?.addEventListener('abort', cancel, { once: true });
      try {
        compaction = await native({
          action: 'workbook',
          operation: 'compact',
          root,
          keep: [current.version],
          vacuum: false,
        });
        if (matchesFormat(current.index.sourceFactSchema, 'SCHEMA_USED_RANGE_FACTS_V2'))
          await resetIndexedSnapshotExtents(root, current.version, current.index.revision, native);
        return compaction;
      } finally {
        options.signal?.removeEventListener('abort', cancel);
      }
    },
  });
  return result
    ? { ...result, ...(compaction ? { compaction } : {}) }
    : { published: false, reason: 'Checkpoint exceeds cache staging quota' };
}
export async function restoreIndexedCheckpoint(root, manifest) {
  if (!manifest?.checkpoint) return null;
  const file = path.join(root, 'checkpoint.json'),
    stat = await fs.lstat(file);
  if (!stat.isFile() || stat.isSymbolicLink() || stat.size > MAX_METADATA)
    throw refusal('internal-invalid-field', 'Invalid exported workbook checkpoint');
  const checkpoint = JSON.parse(await fs.readFile(file, 'utf8'));
  validate(checkpoint, manifest);
  const index = JSON.parse(await fs.readFile(path.join(root, 'index.json'), 'utf8'));
  if (index.revision !== checkpoint.baseRevision)
    throw refusal('internal-mismatch', 'Exported workbook checkpoint index mismatch');
  return {
    root,
    index,
    view: checkpoint.view,
    styles: checkpoint.styles,
    logicalSheets: checkpoint.logicalSheets,
    version: checkpoint.version,
    journal: { previous: null, edits: [] },
    publicRevision: checkpoint.publicRevision,
    savedExport: {
      path: path.join(root, 'saved.xlsx'),
      revision: checkpoint.publicRevision,
      version: checkpoint.version,
    },
  };
}
