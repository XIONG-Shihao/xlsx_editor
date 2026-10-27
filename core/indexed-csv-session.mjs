/** Host-neutral CSV owner. One verified immutable source and row index serve
 * every revision; hosts retain recovery snapshots and publish exported files. */
import fs from 'node:fs/promises';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { createIndexedStorage } from './indexed-storage.mjs';
import { readCsvFileIndex, checkCsvFileIndexSource } from './csv-file.mjs';
import { writeCsvIndex } from './csv-file-save.mjs';
import { applyCsvCommands, csvPosition, csvAddress, csvCacheStatistics } from './csv-workbook.mjs';
import { encodeCsvText } from './csv-codec.mjs';
import { csvView, readCsv, scanCsv, summarizeCsv } from './csv-session-operations.mjs';
import { refusal } from './refusal-codes.mjs';

const COMMAND_BYTES = 2 * 1024 ** 2,
  MAX_COMMANDS = 10000,
  PROJECTION_BYTES = 64 * 1024 ** 2;
const own = (value, key) => Object.hasOwn(value, key);
const stamp = (value) => [value.dev, value.ino, value.size, value.mtimeNs, value.ctimeNs].join(':');
function budget(value, fallback, min, max, label) {
  value ??= fallback;
  if (!Number.isSafeInteger(value) || value < min || value > max)
    throw refusal('internal-invalid-field', `Invalid CSV ${label} budget`);
  return value;
}
function keys(value, allowed) {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.keys(value).some((key) => value[key] !== undefined && !allowed.includes(key))
  )
    throw refusal('csv-bad-input', 'CSV accepts literal cell edits only');
}
function validateCommands(commands, encoding) {
  if (!Array.isArray(commands) || !commands.length || commands.length > MAX_COMMANDS)
    throw refusal('csv-bad-input', 'Submit between one and 10,000 CSV commands');
  if (Buffer.byteLength(JSON.stringify(commands)) > COMMAND_BYTES)
    throw refusal('csv-limit', 'CSV commands exceed the 2 MiB transport budget');
  let cells = 0;
  const literal = (cell) => {
    if (++cells > MAX_COMMANDS) throw refusal('csv-limit', 'Edit at most 10,000 CSV cells at a time');
    if (!own(cell, 'value')) throw refusal('csv-bad-input', 'CSV cell edit requires a literal value');
    if (
      (cell.value !== null && !['string', 'number', 'boolean'].includes(typeof cell.value)) ||
      (typeof cell.value === 'number' && !Number.isFinite(cell.value))
    )
      throw refusal('internal-invalid-field', 'Invalid CSV literal value');
    // Reject malformed Unicode before accepting a revision, not during save.
    encodeCsvText(cell.value === null ? '' : String(cell.value), encoding, false);
  };
  const fields = ['cell', 'value', 'expectedValue', 'expectedRawValue', 'expectedFormula'];
  for (const command of commands) {
    if (command?.type === 'setCells') {
      keys(command, ['type', 'sheet', 'cells']);
      if (!Array.isArray(command.cells) || !command.cells.length)
        throw refusal('csv-bad-input', 'CSV setCells requires cells');
      for (const cell of command.cells) {
        keys(cell, fields);
        literal(cell);
      }
    } else if (command?.type === 'updateRows') {
      keys(command, ['type', 'sheet', 'range', 'keyColumn', 'updates']);
      const [first, last = first] = String(command.range).split(':'),
        a = csvPosition(first),
        b = csvPosition(last);
      if (b.row < a.row || b.column < a.column || (b.row - a.row + 1) * (b.column - a.column + 1) > 1000000)
        throw refusal('csv-limit', 'CSV keyed updates support at most one million positions');
      if (!Array.isArray(command.updates) || !command.updates.length)
        throw refusal('csv-bad-input', 'CSV keyed updates require updates');
      for (const update of command.updates) {
        keys(update, ['key', 'cells']);
        if (!Array.isArray(update.cells) || !update.cells.length)
          throw refusal('csv-bad-input', 'CSV keyed update requires cells');
        for (const cell of update.cells) {
          keys(cell, ['column', ...fields.filter((key) => key !== 'cell')]);
          literal(cell);
        }
      }
    } else {
      keys(command, ['sheet', ...fields]);
      literal(command);
    }
  }
}
function literalCommands(commands) {
  if (!Array.isArray(commands) || !commands.length || commands.length > MAX_COMMANDS)
    throw refusal('csv-bad-input', 'Submit between one and 10,000 CSV commands');
  if (Buffer.byteLength(JSON.stringify(commands)) > COMMAND_BYTES)
    throw refusal('csv-limit', 'CSV commands exceed the 2 MiB transport budget');
  const expanded = [];
  for (const command of commands) {
    if (command?.type !== 'clearRange') {
      expanded.push(command);
      continue;
    }
    keys(command, ['type', 'sheet', 'range']);
    const parts = String(command.range).split(':');
    if (parts.length > 2) throw refusal('internal-invalid-field', 'Invalid CSV clear range');
    const first = csvPosition(parts[0]),
      last = csvPosition(parts[1] ?? parts[0]),
      rows = last.row - first.row + 1,
      columns = last.column - first.column + 1;
    if (
      rows < 1 ||
      columns < 1 ||
      !Number.isSafeInteger(rows * columns) ||
      rows * columns > MAX_COMMANDS ||
      expanded.length + rows * columns > MAX_COMMANDS
    )
      throw refusal('csv-limit', 'Clear at most 10,000 CSV positions at a time');
    for (let row = first.row; row <= last.row; row++)
      for (let column = first.column; column <= last.column; column++)
        expanded.push({ sheet: command.sheet, cell: csvAddress(row, column), value: null });
  }
  return expanded;
}
function overlayBytes(index) {
  let bytes = 256;
  for (const [key, value] of Object.entries(index.changes)) bytes += 80 + key.length * 2 + value.length * 2;
  return bytes + Object.keys(index.widthOverrides).length * 48;
}

export class CsvFileSession {
  id = randomUUID();
  cacheHit = false;
  #source;
  #base;
  #root;
  #states = new Map();
  #sequence = 0;
  #work = Promise.resolve();
  #opening = false;
  #closed = false;
  #exports = new Map();
  #exportBytes = 0;
  #leaseBytes = 0;
  #historyBytes = 0;
  #indexBytes = 0;
  #maxStates = 102;
  #maxHistoryBytes = PROJECTION_BYTES;
  #maxProjectionBytes = PROJECTION_BYTES;
  #maxReplyBytes = 8 * 1024 ** 2;
  #maxSessionBytes = 4 * 1024 ** 3;
  #serial(operation) {
    const result = this.#work.then(() => {
      if (this.#closed) throw refusal('session-closed', 'CSV session is closed');
      return operation();
    });
    this.#work = result.then(
      () => {},
      () => {}
    );
    return result;
  }
  open(source, options = {}) {
    if (this.#opening || this.#closed)
      return Promise.reject(refusal('session-closed', 'CSV session has already been opened or closed'));
    this.#opening = true;
    return this.#serial(async () => {
      try {
        this.#maxStates = budget(options.maxStates, 102, 2, 10001, 'retained-state');
        this.#maxHistoryBytes = budget(options.maxHistoryBytes, PROJECTION_BYTES, 256, PROJECTION_BYTES, 'history');
        this.#maxProjectionBytes = budget(
          options.maxProjectionBytes,
          PROJECTION_BYTES,
          1024,
          PROJECTION_BYTES,
          'projection'
        );
        this.#maxReplyBytes = budget(options.maxReplyBytes, 8 * 1024 ** 2, 1024, PROJECTION_BYTES, 'reply');
        const maxIndexBytes = budget(options.maxIndexBytes, PROJECTION_BYTES, 20, PROJECTION_BYTES, 'row-index');
        this.#maxSessionBytes = budget(
          options.maxSessionBytes,
          4 * 1024 ** 3,
          128 * 1024,
          4 * 1024 ** 3,
          'session disk'
        );
        if (
          options.expectedRevision !== undefined &&
          (typeof options.expectedRevision !== 'string' || !/^(?:sha256:)?[a-f0-9]{64}$/.test(options.expectedRevision))
        )
          throw refusal('internal-invalid-field', 'Invalid CSV source revision');
        this.#root = await createIndexedStorage(this.id, this.#maxSessionBytes, options.storageDirectory);
        this.#leaseBytes = (await fs.stat(path.join(this.#root, 'owner.json'))).size;
        if (this.#closed) throw refusal('session-closed', 'CSV session is closed');
        this.#source = path.resolve(source);
        this.#base = readCsvFileIndex(
          this.#source,
          { delimiter: options.delimiter ?? 'auto', maxIndexBytes },
          options.expectedRevision
        );
        this.#base.maxRowBytes = this.#maxProjectionBytes;
        this.#indexBytes = this.#base.rowOffsets.byteLength + this.#base.rowWidths.byteLength;
        const bytes = overlayBytes(this.#base);
        this.#historyBytes = bytes;
        this.#states.set('0', { index: this.#base, bytes });
        return this.view('0');
      } catch (error) {
        this.#closed = true;
        this.#states.clear();
        this.#base = undefined;
        this.#historyBytes = 0;
        this.#indexBytes = 0;
        if (this.#root) {
          await fs.rm(this.#root, { recursive: true, force: true });
          this.#root = undefined;
        }
        throw error;
      }
    });
  }
  #state(state) {
    if (this.#closed) throw refusal('session-closed', 'CSV session is closed');
    const current = this.#states.get(String(state));
    if (!current) throw refusal('internal-unknown-reference', 'Unknown CSV session revision');
    if (!checkCsvFileIndexSource(current.index, this.#source))
      throw refusal('csv-unsupported', 'CSV source index is unavailable');
    return current;
  }
  #sheet(sheet) {
    if (sheet !== 'CSV') throw refusal('internal-unknown-reference', 'Unknown CSV worksheet');
  }
  #reply(value) {
    if (Buffer.byteLength(JSON.stringify(value)) > this.#maxReplyBytes)
      throw refusal('csv-limit', 'CSV reply exceeds its byte budget; request a smaller range');
    return value;
  }
  view(state) {
    state = String(state);
    const current = this.#state(state),
      revision = this.#base.revision;
    return {
      ...csvView(current.index, state === '0' ? revision : `${revision}:draft:${this.id}:${state}`),
      storage: { kind: 'csv-indexed-v1', sessionId: this.id, state },
    };
  }
  apply(state, commands) {
    return this.#serial(() => {
      const current = this.#state(state);
      commands = literalCommands(commands);
      validateCommands(commands, current.index.encoding);
      if (this.#states.size >= this.#maxStates)
        throw refusal('csv-limit', 'CSV retained-state budget exceeded; compact unused revisions before editing');
      const result = applyCsvCommands(current.index, commands),
        index = result.index;
      const expansion =
        Math.max(0, index.storedCells - this.#base.storedCells) * 48 + Math.max(0, index.rows - this.#base.rows) * 32;
      const bytes = overlayBytes(index);
      if (!Number.isSafeInteger(expansion) || expansion + bytes > this.#maxProjectionBytes)
        throw refusal('csv-limit', 'CSV edit projection budget exceeded');
      if (this.#historyBytes + bytes > this.#maxHistoryBytes)
        throw refusal('csv-limit', 'CSV retained-history budget exceeded; compact unused revisions before editing');
      this.#state(state);
      const next = String(this.#sequence + 1),
        view = {
          ...csvView(index, `${this.#base.revision}:draft:${this.id}:${next}`),
          storage: { kind: 'csv-indexed-v1', sessionId: this.id, state: next },
        };
      const reply = this.#reply({ view, edits: result.edits });
      this.#sequence++;
      this.#states.set(next, { index, bytes });
      this.#historyBytes += bytes;
      return reply;
    });
  }
  read(state, sheet, range) {
    return this.#serial(() => {
      this.#sheet(sheet);
      const result = readCsv(this.#state(state).index, range);
      this.#state(state);
      return this.#reply({ sheet, range, ...result });
    });
  }
  scan(state, sheet, after) {
    return this.#serial(() => {
      this.#sheet(sheet);
      const { cells, nextCursor } = scanCsv(this.#state(state).index, after);
      this.#state(state);
      return this.#reply({ cells, nextCursor });
    });
  }
  query(state, sheet, query, mode = 'display', after, options = {}) {
    return this.#serial(() => {
      this.#sheet(sheet);
      if (typeof query !== 'string' || query.length > 10000)
        throw refusal('internal-invalid-field', 'Invalid CSV search text');
      const result = scanCsv(this.#state(state).index, after, query, mode, options);
      this.#state(state);
      return this.#reply(result);
    });
  }
  selectionSummary(state, sheet, range) {
    return this.#serial(() => {
      this.#sheet(sheet);
      const result = summarizeCsv(this.#state(state).index, range);
      this.#state(state);
      return result;
    });
  }
  export(state) {
    return this.#serial(async () => {
      state = String(state);
      const current = this.#state(state),
        cached = this.#exports.get(state);
      if (cached) {
        if (stamp(await fs.stat(cached.path, { bigint: true })) !== cached.stamp)
          throw refusal('csv-stale', 'CSV export candidate changed');
        this.#state(state);
        return { path: cached.path, revision: cached.revision };
      }
      const output = path.join(this.#root, `export-${randomUUID()}.csv`);
      let created = false;
      try {
        const revision = writeCsvIndex(current.index, this.#source, output, this.#base.revision, {
          maxOutputBytes: this.#maxSessionBytes - this.#leaseBytes - this.#exportBytes,
        });
        created = true;
        const info = await fs.stat(output, { bigint: true });
        this.#state(state);
        const bytes = Number(info.size);
        this.#exports.set(state, { path: output, revision, bytes, stamp: stamp(info) });
        this.#exportBytes += bytes;
        return { path: output, revision };
      } catch (error) {
        if (created) await fs.rm(output, { force: true });
        throw error;
      }
    });
  }
  compact(retain) {
    return this.#serial(async () => {
      if (!Array.isArray(retain) || retain.length > this.#maxStates)
        throw refusal('internal-invalid-field', 'Invalid CSV retained revisions');
      // The immutable source/index is owned separately by #base. A discarded
      // base view must not consume a hidden history slot after undo expires.
      const keep = new Set(retain.map(String));
      for (const state of keep) this.#state(state);
      for (const [state, candidate] of this.#exports)
        if (!keep.has(state)) {
          await fs.rm(candidate.path, { force: true });
          this.#exports.delete(state);
          this.#exportBytes -= candidate.bytes;
        }
      for (const [state, current] of this.#states)
        if (!keep.has(state)) {
          this.#states.delete(state);
          this.#historyBytes -= current.bytes;
        }
      return { compacted: true };
    });
  }
  diagnosticStats() {
    return {
      retainedStates: this.#states.size,
      storageRoots: this.#root ? 1 : 0,
      indexBytes: this.#indexBytes,
      historyBytes: this.#historyBytes,
      exportBytes: this.#exportBytes,
      decodedRowCacheBytes: this.#base ? csvCacheStatistics(this.#base).estimatedBytes : 0,
      closed: this.#closed,
    };
  }
  async close() {
    this.#closed = true;
    await this.#work;
    this.#states.clear();
    this.#exports.clear();
    this.#base = undefined;
    this.#source = undefined;
    this.#indexBytes = 0;
    this.#historyBytes = 0;
    this.#exportBytes = 0;
    if (this.#root) {
      await fs.rm(this.#root, { recursive: true, force: true });
      this.#root = undefined;
    }
  }
}
