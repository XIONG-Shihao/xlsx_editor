/** Shared trusted host service. All file grants/receipts stay in the main process. */
import { realpath } from 'node:fs/promises';
import { randomUUID } from 'node:crypto';
import { snapshotDelimitedFile } from './delimited-file-import.mjs';
import {
  inspectLocalQuery,
  stageLocalQuery,
  queryCellFingerprint,
  queryCellIdentity,
  typedQueryCellIdentity,
  queryHasConversions,
  QUERY_TYPED_FINGERPRINT,
  queryReplacementIdentity,
  queryReplacementCurrent,
  queryConversionDisplay,
  localQueryFields,
} from '../core/local-query.mjs';
import { validatePasteShape, pasteId } from '../core/paste-protocol.mjs';
import {
  createQueryInspectionSpool,
  readQueryInspection,
  querySourceStamp,
  releaseQueryInspections,
} from './local-query-inspection.mjs';
import { cellPosition, columnName } from '../core/workbook-commands.mjs';
import { refusal } from '../core/refusal-codes.mjs';
export async function destinationFingerprint(
  shape,
  { read, check = () => {}, signal, onProgress = () => {}, fingerprintProfile, conversion } = {}
) {
  if (fingerprintProfile !== undefined && (fingerprintProfile !== QUERY_TYPED_FINGERPRINT || !conversion))
    throw refusal('internal-invalid-field', 'Invalid managed query fingerprint profile');
  shape = validatePasteShape(shape);
  const start = cellPosition(shape.start),
    hash = queryCellFingerprint();
  let completed = 0;
  try {
    for (let row = 0; row < shape.rows; row++)
      for (let col = 0; col < shape.columns; col += 16) {
        signal?.throwIfAborted();
        await check();
        const end = Math.min(shape.columns, col + 16),
          ref = (c) => columnName(start.col + c) + (start.row + row + 1),
          range = ref(col) + ':' + ref(end - 1),
          result = await read(shape.sheet, range);
        await check();
        signal?.throwIfAborted();
        for (let c = col; c < end; c++) {
          if (!result?.cells || typeof result.cells !== 'object' || Array.isArray(result.cells))
            throw refusal('internal-invalid-field', 'Invalid managed output read receipt');
          const index = row * shape.columns + c,
            [kind, value] = fingerprintProfile
              ? typedQueryCellIdentity(result.cells[ref(c)], { conversion, shape, index })
              : queryCellIdentity(result.cells[ref(c)]);
          hash.add(row * shape.columns + c, kind, value);
          completed++;
        }
        onProgress({ phase: 'checking-query-output', completed, total: shape.rows * shape.columns });
      }
    return hash.finish();
  } finally {
    hash.close();
  }
}
/** Captured sparse range reads; cache one <=1 MiB rectangle and adapt on
 * oversized replies. The caller hashes only positions it actually claims. */
function destinationReader(shape, { read, check, signal }) {
  // Older helpers retired on replies over16MiB; qualified helpers return the
  // recoverable transport error. Start at16 positions, never grow from samples.
  const start = cellPosition(shape.start);
  let rowStep = Math.max(1, Math.floor(16 / Math.min(16, shape.columns))),
    colStep = Math.min(16, shape.columns),
    cache = null;
  return async (row, col) => {
    for (;;) {
      signal?.throwIfAborted();
      if (cache && row >= cache.row && row <= cache.lastRow && col >= cache.col && col <= cache.lastCol)
        return cache.cells[columnName(start.col + col) + (start.row + row + 1)];
      await check();
      const firstRow = Math.floor(row / rowStep) * rowStep,
        lastRow = Math.min(shape.rows - 1, firstRow + rowStep - 1),
        firstCol = Math.floor(col / colStep) * colStep,
        lastCol = Math.min(shape.columns - 1, firstCol + colStep - 1),
        a = columnName(start.col + firstCol) + (start.row + firstRow + 1),
        b = columnName(start.col + lastCol) + (start.row + lastRow + 1);
      let reply;
      try {
        reply = await read(shape.sheet, a + ':' + b);
      } catch (error) {
        await check();
        if (error?.message !== 'Native workbook reply exceeds transport bound') throw error;
        const rows = lastRow - firstRow + 1,
          columns = lastCol - firstCol + 1;
        if (rows === 1 && columns === 1)
          throw refusal('query-limit', 'A destination cell exceeds the native review transport limit');
        if (rows > 1) rowStep = Math.max(1, Math.floor(rows / 2));
        else colStep = Math.max(1, Math.floor(columns / 2));
        cache = null;
        continue;
      }
      await check();
      signal?.throwIfAborted();
      if (!reply?.cells || typeof reply.cells !== 'object' || Array.isArray(reply.cells))
        throw refusal('internal-invalid-field', 'Invalid query destination read receipt');
      if (Buffer.byteLength(JSON.stringify(reply.cells)) > 1024 * 1024) {
        cache = null;
        if (lastRow > firstRow) {
          rowStep = Math.max(1, Math.floor((lastRow - firstRow + 1) / 2));
          continue;
        }
        if (lastCol > firstCol) {
          colStep = Math.max(1, Math.floor((lastCol - firstCol + 1) / 2));
          continue;
        }
        throw refusal('query-limit', 'A destination cell exceeds the 1 MiB review limit');
      }
      for (const [ref, cell] of Object.entries(reply.cells)) {
        const p = cellPosition(ref);
        if (
          p.row < start.row + firstRow ||
          p.row > start.row + lastRow ||
          p.col < start.col + firstCol ||
          p.col > start.col + lastCol ||
          !cell ||
          typeof cell !== 'object' ||
          Array.isArray(cell)
        )
          throw refusal('query-unsupported', 'Destination read returned a foreign or invalid cell');
      }
      cache = { row: firstRow, lastRow, col: firstCol, lastCol, cells: reply.cells };
    }
  };
}
const newPosition = (row, col, previous) => !previous || row >= previous.rows || col >= previous.columns;
async function destinationReviewFingerprint(shape, previous, options) {
  const get = destinationReader(shape, options),
    hash = queryCellFingerprint();
  let occupied = 0,
    blocked = 0,
    totalNewPositions = 0;
  try {
    for (let row = 0; row < shape.rows; row++)
      for (let col = 0; col < shape.columns; col++)
        if (newPosition(row, col, previous)) {
          const cell = await get(row, col),
            [kind, value] = queryReplacementIdentity(cell),
            status = queryReplacementCurrent(cell);
          hash.add(row * shape.columns + col, kind, value);
          totalNewPositions++;
          if (status.occupied) occupied++;
          if (status.blocked) blocked++;
        }
    return { fingerprint: hash.finish(), occupied, blocked, totalNewPositions };
  } finally {
    hash.close();
  }
}
function destinationReviewCollector(record, inspection, { read, check, signal, spool }) {
  const start = cellPosition(record.target.start),
    hash = queryCellFingerprint();
  const layout = { width: inspection.shape.columns, outputColumns: inspection.outputColumns },
    get = destinationReader(inspection.shape, { read, check, signal });
  let occupied = 0,
    blocked = 0,
    totalNewPositions = 0;
  const visit = async (row, col, raw, sourceRef, result) => {
    if (!newPosition(row, col, record.last?.shape)) return;
    const cell = await get(row, col),
      [kind, value] = queryReplacementIdentity(cell),
      status = queryReplacementCurrent(cell);
    hash.add(row * layout.width + col, kind, value);
    totalNewPositions++;
    if (status.blocked) blocked++;
    if (status.occupied) {
      occupied++;
      await spool.addReplacement({
        index: row * layout.width + col,
        sourceRef,
        destinationRef: columnName(start.col + col) + (start.row + row + 1),
        sourceText: raw,
        outputKind: result.status === 'converted' ? result.kind : 'rejected',
        outputDisplay: queryConversionDisplay(result) ?? result.reason,
        current: status.current,
      });
    }
  };
  const empty = { status: 'converted', kind: 'text', value: '', notices: [], empty: true, padded: true };
  return {
    async onCell(cell) {
      await visit(
        cell.outputRow - 1,
        cell.outputColumn - 1,
        cell.raw,
        columnName(cell.sourceColumn - 1) + cell.sourceRow,
        cell.result
      );
      if (cell.outputColumn === layout.outputColumns)
        for (let col = layout.outputColumns; col < layout.width; col++)
          await visit(cell.outputRow - 1, col, '', null, empty);
    },
    async finish(inspection) {
      for (let row = inspection.outputRows; row < inspection.shape.rows; row++)
        for (let col = 0; col < layout.width; col++) await visit(row, col, '', null, empty);
      return {
        fingerprint: hash.finish(),
        occupied,
        blocked,
        totalNewPositions,
        shape: inspection.shape,
        outputRows: inspection.outputRows,
        outputColumns: inspection.outputColumns,
      };
    },
    close() {
      hash.close();
    },
  };
}
/** Context owns accepted state/actor lifetime; apply crosses its authoritative
 * worker or journal fence. Service cancellation never invents acceptance. */
export class LocalQueryHost {
  #store;
  #document;
  #directory;
  #context;
  constructor({ store, document, directory, context }) {
    this.#store = store;
    this.#document = document;
    this.#directory = directory;
    this.#context = context;
  }
  async #receipt(record) {
    const scope = this.#context.current(false);
    return this.#context.request(
      'importReceipt',
      {
        state: scope.state,
        documentId: scope.documentId,
        descriptor: record.pending.descriptor,
        parentState: record.pending.parentState,
      },
      false
    );
  }
  async list() {
    const found = await this.#store.list(this.#document),
      warnings = [];
    for (const visible of found.entries)
      if (visible.pending) {
        const r = await this.#store.granted(this.#document, visible.id);
        try {
          const receipt = await this.#receipt(r);
          if (receipt.found) await this.#store.finalize(this.#document, r.id, r.pending.id, receipt);
          else
            warnings.push(
              r.name + ': uncertain refresh is retained; no accepted receipt was found in current retained history.'
            );
        } catch (e) {
          warnings.push(r.name + ': pending receipt could not be verified. ' + String(e.message ?? e));
        }
      }
    const latest = await this.#store.list(this.#document);
    return {
      ...latest,
      ...(warnings.length ? { warning: [latest.warning, ...warnings].filter(Boolean).join(' ') } : {}),
    };
  }
  async create(draft, choose) {
    const scope = this.#context.current(),
      filename = await choose();
    this.#context.current();
    if (!filename) return { cancelled: true };
    return { connection: await this.#store.create(this.#document, draft, filename) };
  }
  async update(id, revision, draft) {
    const ownerId = this.#context.current().ownerId ?? null;
    const connection = await this.#store.update(this.#document, id, revision, draft);
    await releaseQueryInspections(this.#document, this.#directory, { ownerId, connectionId: id }).catch((e) =>
      console.warn('Query inspection cleanup deferred', e)
    );
    return { connection };
  }
  async remove(id, revision) {
    const ownerId = this.#context.current().ownerId ?? null;
    const result = await this.#store.remove(this.#document, id, revision);
    await releaseQueryInspections(this.#document, this.#directory, { ownerId, connectionId: id }).catch((e) =>
      console.warn('Query inspection cleanup deferred', e)
    );
    return result;
  }
  async #snapshot(record, options) {
    this.#context.current();
    if ((await realpath(record.source.filename)) !== record.source.filename)
      throw refusal(
        'query-stale',
        'Saved query source changed its canonical identity. Create a new explicit source grant.'
      );
    return snapshotDelimitedFile({
      source: record.source.filename,
      directory: this.#directory,
      signal: options.signal,
      checkCurrent: () => this.#context.current(),
      onProgress: options.onProgress,
    });
  }
  async #assertScope(scope) {
    const live = this.#context.current();
    if (['documentId', 'state', 'ownerId'].some((key) => (scope[key] ?? null) !== (live[key] ?? null)))
      throw refusal('query-stale', 'Query workbook scope retired');
  }
  async #reviewBinding(record, scope) {
    const typed = queryHasConversions(record.query),
      date1904 = this.#context.calendar?.(record.target.sheet);
    if (typed && (typeof scope.ownerId !== 'string' || !scope.ownerId || typeof date1904 !== 'boolean'))
      throw refusal('query-bad-input', 'Typed queries require the captured workbook owner and date calendar');
    return {
      document: this.#document,
      connectionId: record.id,
      definitionRevision: record.revision,
      query: record.query,
      target: record.target,
      previousShape: record.last?.shape ?? null,
      scope: { documentId: scope.documentId, state: scope.state, ownerId: scope.ownerId ?? null },
      date1904: date1904 ?? null,
      sourceGrant: {
        id: record.source.id,
        filename: record.source.filename,
        stamp: await querySourceStamp(record.source.filename),
      },
    };
  }
  async disposeInspections(ownerId) {
    await releaseQueryInspections(this.#document, this.#directory, { ownerId });
  }
  async preview(id, revision, options = {}) {
    const { offset = 0, replacementOffset = 0, expectedConversionRevision } = options;
    if (
      !Number.isInteger(offset) ||
      offset < 0 ||
      !Number.isInteger(replacementOffset) ||
      replacementOffset < 0 ||
      (expectedConversionRevision !== undefined && !pasteId(expectedConversionRevision))
    )
      throw refusal('internal-invalid-field', 'Invalid query conversion page request');
    const record = await this.#store.granted(this.#document, id, revision);
    if (record.pending) throw refusal('query-refused', 'Reconcile the pending refresh before previewing it');
    const scope = this.#context.current(),
      binding = await this.#reviewBinding(record, scope),
      check = async () => {
        options.signal?.throwIfAborted();
        await this.#assertScope(scope);
        if ((this.#context.calendar?.(record.target.sheet) ?? null) !== binding.date1904)
          throw refusal('query-stale', 'Query target date calendar changed');
        if ((await querySourceStamp(record.source.filename)) !== binding.sourceGrant.stamp)
          throw refusal('query-stale', 'Query source changed; Preview again');
      };
    let page;
    if (expectedConversionRevision)
      page = await readQueryInspection(expectedConversionRevision, binding, offset, {
        signal: options.signal,
        check,
        replacementOffset,
        replacementFirst: Object.hasOwn(options, 'replacementOffset'),
      });
    else {
      if (offset !== 0 || replacementOffset !== 0)
        throw refusal('query-bad-input', 'Preview the query before requesting another conversion page');
      const source = await this.#snapshot(record, options);
      let spool,
        collector,
        sealed = false;
      try {
        spool = await createQueryInspectionSpool({
          directory: this.#directory,
          binding,
          signal: options.signal,
          check,
        });
        const inspection = await inspectLocalQuery(source, {
          ...record.target,
          query: record.query,
          previousShape: record.last?.shape,
          date1904: typeof binding.date1904 === 'boolean' ? binding.date1904 : undefined,
          ...options,
          onCell: (cell) => spool.add(cell),
        });
        await source.check();
        await check();
        if ((await querySourceStamp(record.source.filename)) !== binding.sourceGrant.stamp)
          throw refusal('query-stale', 'Query source changed during inspection');
        collector = destinationReviewCollector(record, inspection, {
          read: (sheet, range) => this.#context.request('read', { state: scope.state, sheet, range }),
          check,
          signal: options.signal,
          spool,
        });
        let complete = false;
        for await (const cell of localQueryFields(source, inspection, {
          signal: options.signal,
          onProgress: options.onProgress,
          onComplete: (proof) => {
            if (
              proof.source.sha256 !== inspection.source.sha256 ||
              proof.source.bytes !== inspection.source.bytes ||
              proof.outputRows !== inspection.outputRows ||
              proof.outputColumns !== inspection.outputColumns ||
              JSON.stringify(proof.shape) !== JSON.stringify(inspection.shape) ||
              proof.fingerprint !== inspection.conversionSummary.fingerprint
            )
              throw refusal('query-stale', 'Query source or conversion changed between review passes');
            complete = true;
          },
        }))
          await collector.onCell(cell);
        if (!complete) throw refusal('query-bad-input', 'Query replacement source proof did not complete');
        const replacementSummary = await collector.finish(inspection);
        await source.check();
        await check();
        await spool.seal(inspection, replacementSummary);
        sealed = true;
        page = await readQueryInspection(spool.id, binding, 0, { signal: options.signal, check });
      } finally {
        collector?.close();
        if (spool && !sealed) await spool.dispose().catch((e) => console.warn('Query review cleanup deferred', e));
        await source.dispose().catch((e) => console.warn('Query snapshot cleanup deferred', e));
      }
    }
    await check();
    const { entry, issues } = page,
      inspection = entry.inspection;
    const result = {
      queryPreview: {
        connectionId: id,
        revision,
        source: inspection.source,
        outputRows: inspection.outputRows,
        outputColumns: inspection.outputColumns,
        rows: inspection.preview,
        truncated: inspection.previewTruncated,
        target: record.target,
        conversionInspection: {
          schema: 1,
          revision: entry.id,
          date1904: typeof binding.date1904 === 'boolean' ? binding.date1904 : null,
          policies: Array.from({ length: inspection.outputColumns }, (_, i) => record.query.policies?.[i] ?? 'text'),
          rejected: inspection.conversionSummary.rejected,
          changed: inspection.conversionSummary.changed,
          totalIssues: entry.count,
          offset: page.offset,
          nextOffset: page.nextOffset,
          issues,
          expiresAt: entry.expires,
        },
        replacementInspection: {
          schema: 1,
          revision: entry.id,
          occupied: entry.replacementSummary.occupied,
          blocked: entry.replacementSummary.blocked,
          totalNewPositions: entry.replacementSummary.totalNewPositions,
          offset: page.replacementOffset,
          nextOffset: page.nextReplacementOffset,
          rows: page.replacements,
        },
      },
    };
    if (Buffer.byteLength(JSON.stringify(result)) > 1024 * 1024)
      throw refusal('query-limit', 'Query preview exceeds its 1 MiB response limit');
    return result;
  }
  /** Internal trusted preparation for a complete Refresh All plan. Not an IPC DTO.
   * Reuses paged inspection, but captures small immutable proofs, not 32 spools. */
  async prepareRefreshAllUnit(id, revision, options = {}) {
    const scope = this.#context.current(),
      record = await this.#store.granted(this.#document, id, revision);
    const { queryPreview } = await this.preview(id, revision, options),
      binding = await this.#reviewBinding(record, scope);
    const { entry } = await readQueryInspection(queryPreview.conversionInspection.revision, binding, 0, {
      signal: options.signal,
      check: () => this.#assertScope(scope),
    });
    if (record.last) {
      const actual = await destinationFingerprint(record.last.shape, {
        fingerprintProfile: record.last.fingerprintProfile,
        conversion: record.last.conversion,
        read: (sheet, range) => this.#context.request('read', { state: scope.state, sheet, range }),
        check: () => this.#assertScope(scope),
        signal: options.signal,
        onProgress: options.onProgress,
      });
      if (actual !== record.last.fingerprint)
        throw refusal('query-stale', 'Query output changed since its last accepted refresh. Your edits were kept.');
    }
    await this.#assertScope(scope);
    const { inspection, replacementSummary } = entry;
    return {
      preview: queryPreview,
      shape: structuredClone(inspection.shape),
      proof: {
        definition: {
          id: record.id,
          revision: record.revision,
          name: record.name,
          target: record.target,
          query: record.query,
        },
        sourceGrant: binding.sourceGrant,
        source: { name: inspection.source.name, bytes: inspection.source.bytes, sha256: inspection.source.sha256 },
        conversionFingerprint: inspection.conversionSummary.fingerprint,
        date1904: binding.date1904,
        previous: record.last ?? null,
        shape: inspection.shape,
        replacementFingerprint: replacementSummary.fingerprint,
      },
    };
  }
  async refresh(
    id,
    revision,
    {
      signal,
      onProgress = () => {},
      authorizeApply = async () => {},
      beforeApply = () => {},
      expectedConversionRevision,
      expectedReplacementRevision,
    } = {}
  ) {
    if (expectedReplacementRevision !== undefined && !pasteId(expectedReplacementRevision))
      throw refusal('internal-invalid-field', 'Invalid destination replacement review');
    if (
      expectedConversionRevision &&
      expectedReplacementRevision &&
      expectedConversionRevision !== expectedReplacementRevision
    )
      throw refusal('internal-mismatch', 'Destination and conversion review receipts differ');
    if (expectedConversionRevision !== undefined && !pasteId(expectedConversionRevision))
      throw refusal('internal-invalid-field', 'Invalid reviewed query receipt');
    await this.list();
    const record = await this.#store.granted(this.#document, id, revision);
    if (record.pending)
      throw refusal(
        'session-busy',
        'An uncertain refresh receipt is retained. Accepted attachments were kept; recover its original history before retrying.'
      );
    const scope = this.#context.current(),
      binding = await this.#reviewBinding(record, scope);
    if (queryHasConversions(record.query) && !expectedConversionRevision)
      throw refusal('query-bad-input', 'Preview and review typed conversions before Refresh');
    const source = await this.#snapshot(record, { signal, onProgress });
    let staged,
      intent,
      sent = false,
      accepted = false,
      acceptance;
    const check = async () => {
      signal?.throwIfAborted();
      await this.#assertScope(scope);
      if ((this.#context.calendar?.(record.target.sheet) ?? null) !== binding.date1904)
        throw refusal('query-stale', 'Query target date calendar changed');
    };
    const prune = () => this.#context.request('importPrune', { state: this.#context.current(false).state }, false);
    const call = (operation, input = {}) => this.#context.request(operation, { state: scope.state, ...input });
    try {
      if (record.last) {
        const actual = await destinationFingerprint(record.last.shape, {
          fingerprintProfile: record.last.fingerprintProfile,
          conversion: record.last.conversion,
          read: (sheet, range) => call('read', { sheet, range }),
          check,
          signal,
          onProgress,
        });
        if (actual !== record.last.fingerprint)
          throw refusal(
            'query-stale',
            'Query output changed since its last accepted refresh. Undo/Redo to that result or choose a new target; your edits were kept.'
          );
      }
      const inspection = await inspectLocalQuery(source, {
        ...record.target,
        query: record.query,
        previousShape: record.last?.shape,
        date1904: typeof binding.date1904 === 'boolean' ? binding.date1904 : undefined,
        signal,
        onProgress,
      });
      await check();
      const reviewId = expectedConversionRevision ?? expectedReplacementRevision;
      let reviewed;
      if (reviewId) {
        reviewed = (await readQueryInspection(reviewId, binding, 0, { signal, check })).entry;
        if (
          reviewed.inspection.source.sha256 !== inspection.source.sha256 ||
          reviewed.inspection.source.bytes !== inspection.source.bytes ||
          reviewed.inspection.conversionSummary.fingerprint !== inspection.conversionSummary.fingerprint ||
          JSON.stringify(reviewed.inspection.shape) !== JSON.stringify(inspection.shape)
        )
          throw refusal('query-stale', 'Query review changed; Preview again before Refresh');
      }
      if (inspection.conversionSummary.rejected)
        throw refusal(
          'query-refused',
          'Query has ' +
            inspection.conversionSummary.rejected +
            ' rejected conversions. Review the issues or choose Text before Refresh'
        );
      const currentReplacement = await destinationReviewFingerprint(inspection.shape, record.last?.shape, {
        read: (sheet, range) => call('read', { sheet, range }),
        check,
        signal,
      });
      if (currentReplacement.blocked)
        throw refusal('query-protected', 'Query destination contains protected special cells that cannot be replaced.');
      if (reviewed && currentReplacement.fingerprint !== reviewed.replacementSummary.fingerprint)
        throw refusal('query-stale', 'Query destination changed after review. Preview the replacements again.');
      if (currentReplacement.occupied && (!expectedReplacementRevision || !reviewed))
        throw refusal(
          'query-conflict',
          'Preview and explicitly acknowledge occupied destination replacements before Refresh.'
        );
      const sink = {
        begin: (input) => call('importBegin', input),
        upload: ({ id, ...input }) => call('importUpload', { ...input, uploadId: id }),
        seal: ({ id, ...input }) => call('importSeal', { ...input, uploadId: id }),
        cancel: ({ id, ...input }) => call('importCancel', { ...input, uploadId: id }),
        prune,
      };
      staged = await stageLocalQuery(source, inspection, {
        state: scope.state,
        signal,
        checkCurrent: check,
        sink,
        onProgress,
      });
      await source.check();
      await check();
      intent = {
        id: randomUUID(),
        descriptor: staged.descriptor,
        parentState: scope.state,
        output: {
          shape: inspection.shape,
          fingerprint: staged.outputFingerprint,
          ...(inspection.conversion
            ? { fingerprintProfile: QUERY_TYPED_FINGERPRINT, conversion: inspection.conversion }
            : {}),
          sourceSha256: inspection.source.sha256,
          sourceBytes: inspection.source.bytes,
          outputRows: inspection.outputRows,
          outputColumns: inspection.outputColumns,
          acceptedHead: scope.state,
          updatedAt: Date.now(),
        },
      };
      await this.#store.begin(this.#document, id, revision, intent);
      await authorizeApply();
      await source.check();
      await check();
      const finalReplacement = await destinationReviewFingerprint(inspection.shape, record.last?.shape, {
        read: (sheet, range) => call('read', { sheet, range }),
        check,
        signal,
      });
      if (finalReplacement.fingerprint !== currentReplacement.fingerprint || finalReplacement.blocked)
        throw refusal('query-stale', 'Query destination changed before acceptance; Preview again.');
      if (record.last) {
        const actual = await destinationFingerprint(record.last.shape, {
          fingerprintProfile: record.last.fingerprintProfile,
          conversion: record.last.conversion,
          read: (sheet, range) => call('read', { sheet, range }),
          check,
          signal,
        });
        if (actual !== record.last.fingerprint)
          throw refusal('query-stale', 'Query output changed before acceptance; your edits were kept.');
      }
      await source.check();
      await check();
      beforeApply();
      sent = true;
      acceptance = await this.#context.apply(scope.state, staged.descriptor);
      accepted = true;
      // Receipt/metadata failure after accepted apply is a warning, never a false
      // data rejection. Pending intent remains available for crash reconciliation.
      let connection, warning;
      try {
        const receipt = await this.#receipt({ ...record, pending: intent });
        if (!receipt.found) throw refusal('query-bad-input', 'Accepted receipt is not in retained history');
        connection = await this.#store.finalize(this.#document, id, intent.id, receipt);
      } catch (e) {
        warning = 'Refresh data was accepted; saved connection receipt update is pending. ' + String(e.message ?? e);
      }
      return {
        refreshed: true,
        acceptance,
        ...(connection ? { connection } : {}),
        ...(warning ? { warning } : {}),
        querySource: inspection.source,
        outputRows: inspection.outputRows,
        outputColumns: inspection.outputColumns,
        count: inspection.count,
      };
    } finally {
      if (intent && !sent)
        await this.#store
          .abandonBeforeApply(this.#document, id, intent.id)
          .catch((e) => console.warn('Pending unsubmitted query intent retained', e));
      // This prune consults actual trusted retained history, including Redo; it
      // cannot interpret a missing receipt as permission to discard accepted data.
      if (staged && !accepted)
        await Promise.resolve()
          .then(prune)
          .catch((e) => console.warn('Query attachments retained after cleanup refusal', e));
      await source.dispose().catch((e) => console.warn('Query private source retained after cleanup refusal', e));
    }
  }
}
