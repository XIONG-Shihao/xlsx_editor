/** Actual private/current helper direct owner qualification; no Excel runtime claim. */
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { createWorkbook } from '../core/workbook-package.mjs';
import { calculationTraceFixture, traceExpected } from './support/calculation-trace-fixture.mjs';
import { record } from './support/file-record.mjs';
const output = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-formula-trace-'));
const helper = nativeWorkbookBinary();
const client = createWorkbookNativeClient(helper),
  owners = [];
let root;
const native = async (request) => {
  if (request.operation === 'trace') root = request.root;
  return client.call(request);
};
const make = () => {
  const book = new IndexedWorkbook(
    native,
    () => assert.fail('No legacy writer'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  owners.push(book);
  return book;
};
async function footprint(book, source) {
  return {
    view: book.view('0'),
    stats: book.diagnosticStats(),
    source: (await record(source)).sha256,
    sqlite: root ? (await record(path.join(root, 'cells.sqlite'))).sha256 : null,
  };
}
async function full(book, state, sheet, cell, direction, limit = 100) {
  let after;
  const edges = [],
    markers = [];
  let response;
  do {
    response = await book.trace(state, sheet, cell, { direction, limit, ...(after ? { after } : {}) });
    assert(response.edges.length + response.markers.length <= limit);
    edges.push(...response.edges);
    markers.push(...response.markers);
    after = response.nextCursor;
  } while (after);
  return { edges, markers, coverage: response.coverage };
}
try {
  const source = path.join(output, 'source.xlsx'),
    bytes = calculationTraceFixture();
  await fs.writeFile(source, bytes);
  const book = make();
  await book.open(source);
  await book.trace('0', 'Inputs', 'B1', { direction: 'precedents' });
  const before = await footprint(book, source);
  for (const expected of traceExpected.filter((e) => !e.complete && e.cell[0] !== 'C')) {
    const result = await full(book, '0', expected.sheet, expected.cell, expected.direction);
    assert.deepEqual(
      result.edges.map(({ kind, ...e }) => e),
      expected.edges
    );
    assert(result.edges.every((e) => e.kind === 'reference'));
    assert.equal(result.coverage.complete, true);
  }
  const after = await footprint(book, source);
  assert.deepEqual(after, before);
  const first = await book.trace('0', 'Inputs', 'A1', { direction: 'dependents', limit: 1 });
  assert(first.nextCursor);
  const paged = await full(book, '0', 'Inputs', 'A1', 'dependents', 1);
  assert.deepEqual(
    paged.edges.map(({ kind, ...e }) => e),
    traceExpected.find((e) => e.direction === 'dependents').edges
  );
  await assert.rejects(
    () => book.trace('0', 'Inputs', 'A2', { direction: 'dependents', after: first.nextCursor, limit: 1 }),
    /stale|target/
  );
  await assert.rejects(
    () => book.trace('0', 'Inputs', 'A1', { direction: 'precedents', after: first.nextCursor, limit: 1 }),
    /stale|target/
  );
  await assert.rejects(() =>
    book.trace('0', 'Inputs', 'A1', { direction: 'dependents', after: first.nextCursor.slice(0, -1) + 'x', limit: 1 })
  );
  const changed = await book.apply('0', [{ sheet: 'Inputs', cell: 'B1', formula: 'A2+7' }]),
    state = changed.view.storage.state;
  assert.equal((await book.trace(state, 'Inputs', 'B1', { direction: 'precedents' })).edges[0].from.range, 'A2');
  assert.equal((await book.trace('0', 'Inputs', 'B1', { direction: 'precedents' })).edges[0].from.range, 'A1');
  await assert.rejects(
    () => book.trace(state, 'Inputs', 'A1', { direction: 'dependents', after: first.nextCursor }),
    /stale|target/
  );
  const second = make();
  await second.open(source);
  await assert.rejects(
    () => second.trace('0', 'Inputs', 'A1', { direction: 'dependents', after: first.nextCursor }),
    /another owner/
  );
  const markedSource = path.join(output, 'marked.xlsx');
  await fs.writeFile(markedSource, calculationTraceFixture({ markers: true }));
  const marked = make();
  await marked.open(markedSource);
  for (const expected of traceExpected.filter((e) => e.cell[0] === 'C')) {
    const r = await full(marked, '0', expected.sheet, expected.cell, expected.direction);
    assert.deepEqual(
      r.edges.map(({ kind, ...e }) => e),
      expected.edges
    );
    if (expected.requiredMarker) {
      assert(r.markers.some((m) => m.kind === expected.requiredMarker));
      assert.equal(r.coverage.complete, false);
    } else assert.equal(r.coverage.complete, true);
  }
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>SEQUENCE(3,2)</f></c><c r="D1"><f>SUM(A1#)</f></c><c r="F1"><f>LEN("Z999")</f></c></row></sheetData></worksheet>'
  );
  const spillSource = path.join(output, 'spill.xlsx');
  await fs.writeFile(spillSource, zipSync(parts));
  const spill = make();
  await spill.open(spillSource);
  const follower = await spill.trace('0', 'Sheet1', 'B2', { direction: 'precedents' });
  assert.deepEqual(follower.edges, [
    {
      from: { sheet: 'Sheet1', range: 'A1' },
      to: { sheet: 'Sheet1', range: 'B2' },
      kind: 'spill-origin',
      via: 'formula',
    },
  ]);
  const consumer = await spill.trace('0', 'Sheet1', 'D1', { direction: 'precedents' });
  assert.equal(consumer.edges[0].kind, 'spill-range');
  assert.equal(consumer.edges[0].from.range, 'A1:B3');
  const outputs = await full(spill, '0', 'Sheet1', 'A1', 'dependents', 1);
  assert(outputs.edges.some((e) => e.kind === 'spill-range' && e.to.range === 'A1:B3'));
  assert.equal((await spill.trace('0', 'Sheet1', 'F1', { direction: 'precedents' })).edges.length, 0);
  const saved = await book.export(state),
    sourceHash = (await record(source)).sha256,
    saveHash = (await record(saved.path)).sha256;
  await full(book, state, 'Inputs', 'A2', 'dependents');
  assert.equal((await record(saved.path)).sha256, saveHash);
  assert.equal((await record(source)).sha256, sourceHash);
  for (const options of [
    { direction: 'bad' },
    { direction: 'precedents', limit: 101 },
    { direction: 'precedents', limit: 0 },
    { direction: 'precedents', after: 'x.y' },
    { direction: 'precedents', view: {} },
  ])
    await assert.rejects(() => book.trace('0', 'Inputs', 'B1', options));
  await assert.rejects(() => book.trace('0', 'Inputs', 'A0', { direction: 'precedents' }));
  const version = Number(book.view('0').storage.state);
  await assert.rejects(
    () =>
      native({
        action: 'workbook',
        operation: 'trace',
        root,
        version: 999999,
        view: book.view('0'),
        sheet: 'Inputs',
        cell: 'A1',
        direction: 'precedents',
        limit: 100,
        cursor: null,
      }),
    /Unknown trace version/
  );
} finally {
  const closing = await Promise.allSettled(owners.map((owner) => owner.close()));
  client.close();
  await fs.rm(output, { recursive: true, force: true });
  for (const result of closing) if (result.status === 'rejected') throw result.reason;
}
