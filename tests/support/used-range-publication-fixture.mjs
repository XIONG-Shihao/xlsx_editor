/** Bounded native-envelope substitute for JS/FS publication checks.
 * No SQLite, native executable or Excel calculation is exercised here. */
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { SpreadsheetDocument } from '../../adapters/document-session.mjs';
import { createWorkbook } from '../../core/workbook-package.mjs';
import { cellPosition, columnName, rangeBounds } from '../../core/workbook-commands.mjs';
import { engineQuestion } from './engine-questions.mjs';
export const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const union = (a, b) =>
  !a
    ? b
    : !b
      ? a
      : {
          minRow: Math.min(a.minRow, b.minRow),
          minColumn: Math.min(a.minColumn, b.minColumn),
          maxRow: Math.max(a.maxRow, b.maxRow),
          maxColumn: Math.max(a.maxColumn, b.maxColumn),
        };
const single = (ref) => {
  const p = cellPosition(ref);
  return { minRow: p.row, minColumn: p.col, maxRow: p.row, maxColumn: p.col };
};
const bounds = (cells) => Object.keys(cells).reduce((result, ref) => union(result, single(ref)), null);
const ref = (row, col) => columnName(col) + (row + 1),
  range = (b) =>
    !b
      ? 'A1'
      : ref(b.minRow, b.minColumn) === ref(b.maxRow, b.maxColumn)
        ? ref(b.minRow, b.minColumn)
        : ref(b.minRow, b.minColumn) + ':' + ref(b.maxRow, b.maxColumn);
const cell = (value) => ({ value, dataType: 'number', rawValue: String(value), cachedValue: null, styleId: 0 });
function pack(template, cells) {
  const parts = unzipSync(template),
    rows = new Map();
  for (const address of Object.keys(cells).sort(
    (a, b) => cellPosition(a).row - cellPosition(b).row || cellPosition(a).col - cellPosition(b).col
  )) {
    const row = cellPosition(address).row + 1;
    if (!rows.has(row)) rows.set(row, []);
    rows.get(row).push('<c r="' + address + '"><v>' + cells[address].value + '</v></c>');
  }
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="' +
      range(bounds(cells)) +
      '"/><sheetData>' +
      [...rows].map(([row, items]) => '<row r="' + row + '">' + items.join('') + '</row>').join('') +
      '</sheetData></worksheet>'
  );
  return zipSync(parts, { level: 0, mtime: new Date('2020-01-01T00:00:00Z') });
}
function parse(bytes) {
  const source = strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']),
    cells = {};
  for (const match of source.matchAll(/<c r="([A-Z]+\d+)"><v>([^<]+)<\/v><\/c>/g))
    cells[match[1]] = cell(Number(match[2]));
  return cells;
}
export async function publicationFixture(t, { untitled = false, maxRecoveryBytes, witnessBytes } = {}) {
  const directory = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-used-range-publication-'))),
    source = path.join(directory, 'source.xlsx'),
    recovery = path.join(directory, 'recovery'),
    indices = path.join(directory, 'indices');
  const original = pack(createWorkbook(), { A1: cell(7), C6: cell(3) });
  await fs.writeFile(source, original);
  const roots = new Map(),
    calls = [],
    owners = [],
    policy = {
      schema: 1,
      mode: 'auto',
      recalculateBeforeSave: true,
      pendingFormulaCells: 0,
      canCalculateSheet: true,
      canChangeOptions: true,
    };
  let onCall,
    publicationReply = (reply) => reply;
  const native = async (input) => {
    const answer = engineQuestion(input, ['range', 'used-range']);
    if (answer) return answer;
    calls.push(structuredClone(input));
    if (onCall) await onCall(input);
    if (input.operation === 'copy-source') {
      await fs.copyFile(input.source, path.join(input.root, 'source.xlsx'));
      return {};
    }
    if (input.operation === 'index') {
      const bytes = await fs.readFile(path.join(input.root, 'source.xlsx')),
        cells = parse(bytes),
        extent = bounds(cells),
        revision = hash(bytes),
        root = {
          bytes,
          revision,
          sourceCells: extent,
          states: new Map([[0, { cells, history: null }]]),
          exports: new Map(),
          next: 0,
        };
      roots.set(input.root, root);
      const index = {
        revision,
        sheets: [
          {
            name: 'Sheet1',
            path: 'xl/worksheets/sheet1.xml',
            rows: Math.max(20, (extent?.maxRow ?? 0) + 1),
            columns: Math.max(8, (extent?.maxColumn ?? 0) + 1),
            hiddenRows: [],
            visibleRows: [],
            rowHeights: {},
            storedCells: Object.keys(cells).length,
            formulaCount: 0,
          },
        ],
      };
      await fs.writeFile(path.join(input.root, 'metadata.xlsx'), bytes);
      await fs.writeFile(path.join(input.root, 'index.json'), JSON.stringify(index));
      await fs.writeFile(path.join(input.root, 'cells.sqlite'), 'explicit mock boundary, never executed');
      return index;
    }
    if (input.operation === 'artifact-write') {
      await fs.writeFile(path.join(input.root, input.name), input.data);
      return {};
    }
    if (input.operation === 'calculation-policy' && input.method === 'initialize')
      return { policy: structuredClone(policy) };
    if (input.operation === 'compact') return { retainedVersions: input.keep.length };
    const root = roots.get(input.root),
      state = root?.states.get(input.version);
    assert(root && state, 'Known mock root/version required');
    if (input.operation === 'calculation-transfer') {
      assert.equal(input.sourceRevision, root.revision, 'Publication is bound to actual copied bytes');
      let witness;
      if (input.purpose === 'publication') {
        assert.deepEqual(Object.keys(input).sort(), [
          'action',
          'operation',
          'previousRoot',
          'previousVersion',
          'purpose',
          'root',
          'sourceRevision',
          'version',
          'view',
        ]);
        const previous = roots.get(input.previousRoot);
        assert(previous?.states.has(input.previousVersion), 'Captured prior mock selected version exists');
        assert.equal(
          previous.exports.get(input.previousVersion),
          root.revision,
          'Warm transfer requires an actual publication export receipt'
        );
        const body = Buffer.from(
          JSON.stringify({
            schema: 'workbook.xlsx.mock-calculation-witness.v1',
            sourceRevision: root.revision,
            policy,
            contentHash: hash(Buffer.from(JSON.stringify(state.cells))),
          })
        );
        assert(
          witnessBytes === undefined || (Number.isSafeInteger(witnessBytes) && witnessBytes >= body.length),
          'Mock witness byte size is bounded by explicit test setup'
        );
        const bytes = Buffer.alloc(witnessBytes ?? body.length, 0x20);
        body.copy(bytes);
        const file = 'publication-calculation.sqlite';
        await fs.writeFile(path.join(input.root, file), bytes, { flag: 'wx' });
        witness = { schema: 1, file, revision: hash(bytes), bytes: bytes.length };
      } else {
        assert.equal(input.purpose, 'publication-restore');
        assert.deepEqual(Object.keys(input).sort(), [
          'action',
          'operation',
          'purpose',
          'root',
          'sourceRevision',
          'version',
          'view',
          'witness',
          'witnessRevision',
        ]);
        const bytes = await fs.readFile(input.witness);
        assert.equal(hash(bytes), input.witnessRevision, 'Restore uses hash-verified real filesystem bytes');
        const saved = JSON.parse(bytes.toString());
        assert.equal(saved.schema, 'workbook.xlsx.mock-calculation-witness.v1');
        assert.equal(saved.sourceRevision, root.revision);
        assert.deepEqual(saved.policy, policy);
        assert.equal(
          saved.contentHash,
          hash(Buffer.from(JSON.stringify(state.cells))),
          'Mock witness cannot attach to another source content'
        );
      }
      const version = ++root.next;
      root.states.set(version, structuredClone(state));
      return publicationReply(
        {
          schema: 'workbook.xlsx.publication-calculation-' + (witness ? 'transfer' : 'restore') + '.v1',
          version,
          policy: structuredClone(policy),
          transferredFormulaCells: 0,
          transferredDerivedCells: 0,
          transferredSpillGroups: 0,
          ...(witness ? { witness } : {}),
        },
        input
      );
    }
    if (input.operation === 'range') {
      const b = rangeBounds(input.range);
      return {
        cells: Object.entries(state.cells).filter(([address]) => {
          const p = cellPosition(address);
          return p.row >= b.first.row && p.row <= b.last.row && p.col >= b.first.col && p.col <= b.last.col;
        }),
        nextOffset: null,
      };
    }
    if (input.operation === 'patch') {
      const next = structuredClone(state);
      for (const change of input.changes) {
        if (change.value.value === null) delete next.cells[change.cell];
        else next.cells[change.cell] = structuredClone(change.value);
        next.history = union(next.history, single(change.cell));
      }
      const version = ++root.next;
      root.states.set(version, next);
      return { version };
    }
    if (input.operation === 'bulk') {
      const next = structuredClone(state);
      for (const command of input.commands) {
        assert.equal(command.type, 'clearRange');
        const b = rangeBounds(command.range);
        for (const address of Object.keys(next.cells)) {
          const p = cellPosition(address);
          if (p.row >= b.first.row && p.row <= b.last.row && p.col >= b.first.col && p.col <= b.last.col)
            delete next.cells[address];
        }
      }
      const version = ++root.next;
      root.states.set(version, next);
      return { version, formulaTouched: false, formulaDelta: {} };
    }
    if (input.operation === 'used-range') {
      assert.deepEqual(Object.keys(input).sort(), [
        'action',
        'operation',
        'root',
        'sheet',
        'sourceRevision',
        'version',
      ]);
      assert.equal(input.sourceRevision, root.revision, 'Actual owner must send its captured index source hash');
      const effectiveCells = bounds(state.cells),
        extent = union(union(root.sourceCells, effectiveCells), state.history),
        scope = extent ?? { minRow: 0, minColumn: 0, maxRow: 0, maxColumn: 0 };
      return {
        schema: 'workbook.xlsx.used-range.v1',
        policy: 'xlsx-used-range-v1',
        status: 'complete',
        version: input.version,
        sourceRevision: root.revision,
        sheet: 'Sheet1',
        sheetPath: input.sheet,
        scopeMinMax: scope,
        range: range(scope),
        lastCell: ref(scope.maxRow, scope.maxColumn),
        empty: !extent,
        provenance: {
          factsSchema: 'workbook.xlsx.used-range-facts.v1',
          declaration: {
            status: 'valid',
            ref: range(root.sourceCells),
            bounds: root.sourceCells ?? { minRow: 0, minColumn: 0, maxRow: 0, maxColumn: 0 },
          },
          sourceCells: root.sourceCells,
          sourceMerges: null,
          effectiveCells,
          sourceAxisFormatting: null,
          acceptedHistory: state.history,
          sourceDefaultFormatting: false,
        },
        reasons: [],
      };
    }
    if (input.operation === 'export-plan') return { styleCount: 0, special: false, incomplete: false };
    if (input.operation === 'export') {
      const bytes = pack(root.bytes, state.cells),
        revision = hash(bytes);
      await fs.writeFile(path.join(input.root, 'saved.xlsx'), bytes);
      if (input.publication === true) root.exports.set(input.version, revision);
      return { file: 'saved.xlsx', revision };
    }
    assert.fail('Unexpected native envelope in publication fixture: ' + input.operation);
  };
  const make = () => {
    const owner = new SpreadsheetDocument({
      native,
      recoveryDirectory: recovery,
      indexDirectory: indices,
      cacheIdentity: null,
      ...(maxRecoveryBytes ? { maxRecoveryBytes } : {}),
    });
    owners.push(owner);
    return owner;
  };
  const owner = make();
  if (untitled) await owner.create('publication fixture');
  else await owner.open(source);
  t.after(async () => {
    for (const book of owners) await book.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  return {
    directory,
    source,
    recovery,
    indices,
    original,
    owner,
    make,
    native,
    calls,
    setCall: (callback) => (onCall = callback),
    setPublicationReply: (callback) => (publicationReply = callback ?? ((reply) => reply)),
    manifestPath: () =>
      path.join(recovery, untitled ? 'draft-' + owner.status().documentId : hash(source), 'current.json'),
  };
}
export const used = (owner) =>
  owner
    .request({
      operation: 'usedRange',
      documentId: owner.status().documentId,
      state: owner.status().state,
      sheet: 'Sheet1',
    })
    .then((result) => result.usedRange);
export const edit = (owner, cell, value) =>
  owner.request({ operation: 'apply', state: owner.status().state, commands: [{ sheet: 'Sheet1', cell, value }] });
export async function farClear(owner) {
  await edit(owner, 'A100', 9);
  await edit(owner, 'Z1', 11);
  const far = owner.status();
  await owner.request({
    operation: 'apply',
    state: far.state,
    commands: [
      { type: 'clearRange', sheet: 'Sheet1', range: 'A100' },
      { type: 'clearRange', sheet: 'Sheet1', range: 'Z1' },
    ],
  });
  return far;
}
