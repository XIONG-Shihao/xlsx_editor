/** Formula preview on a real native indexed workbook: results, isolation, cancellation and refusals. */
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { randomUUID, createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { formulaPreviewCoreFixture, formulaPreviewCoreCases } from './support/formula-preview-core-fixture.mjs';
const output = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-formula-preview-core-')),
  record = async (file) => createHash('sha256').update(await fs.readFile(file)).digest('hex');
const client = createWorkbookNativeClient(nativeWorkbookBinary()),
  owners = [],
  calls = [];
let root, hold;
const native = async (request) => {
  calls.push({ operation: request.operation, version: request.version, requestId: request.requestId });
  if (request.operation === 'preview-formula') root = request.root;
  const result = await client.call(request);
  if (hold && request.operation === 'preview-formula') {
    hold.reply = result;
    hold.started();
    await hold.promise;
  }
  return result;
};
const make = () => {
  const owner = new IndexedWorkbook(
    native,
    () => assert.fail('No legacy writer'),
    (input) => calculateIndexedWorkbook(input, native)
  );
  owners.push(owner);
  return owner;
};
async function footprint(owner, source, state = '0') {
  const names = await fs.readdir(root);
  return {
    view: owner.view(state),
    stats: owner.diagnosticStats(),
    source: await record(source),
    files: await Promise.all(names.sort().map((n) => record(path.join(root, n)))),
    database: JSON.parse(
      execFileSync(
        process.env.XLSX_DESKTOP_PYTHON ?? 'python3',
        [
          '-c',
          `import sqlite3,json,sys\nd=sqlite3.connect('file:'+sys.argv[1]+'?mode=ro',uri=True)\nprint(json.dumps({n:d.execute('SELECT count(*) FROM '+n).fetchone()[0] for (n,) in d.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")}))`,
          path.join(root, 'cells.sqlite'),
        ],
        { encoding: 'utf8' }
      )
    ),
  };
}
const request = async (owner, formula, target = 'E1', state = '0') =>
  owner.previewFormula(state, 'Sheet1', target, formula, { requestId: randomUUID() });
try {
  const source = path.join(output, 'source.xlsx');
  await fs.writeFile(source, formulaPreviewCoreFixture());
  const owner = make();
  await owner.open(source);
  await request(owner, '1');
  const before = await footprint(owner, source);
  for (const [formula, expected] of formulaPreviewCoreCases) {
    const result = await request(owner, formula);
    assert.deepEqual(result.result, expected, formula);
  }
  assert.deepEqual(await footprint(owner, source), before);
  assert.equal((await owner.read('0', 'Sheet1', 'B1:B2')).cells.B2.calculation.pending, true);
  for (const formula of ['F1', 'E1+1']) assert.equal((await request(owner, formula)).result.status, 'invalid', formula);
  const replaced = await request(owner, 'A1+20', 'B1');
  assert.equal(replaced.result.value, 22);
  assert.equal((await request(owner, 'B1+1', 'B1')).result.status, 'invalid');
  for (const formula of [
    'BadScope',
    'Dup',
    'IF(FALSE,UNKNOWN(A1),12)',
    'IFERROR(UNKNOWN(A1),99)',
    'IF(TRUE,12,INDIRECT("A1"))',
    'G1+1',
    '9007199254740993',
    '1e999',
    'SUM(A1:A10000)',
  ])
    assert.equal((await request(owner, formula)).result.status, 'unavailable', formula);
  assert.equal((await request(owner, '1', 'M1')).result.status, 'unavailable');
  assert.deepEqual(await footprint(owner, source), before);
  const requestId = randomUUID();
  let start;
  const started = new Promise((r) => (start = r));
  let release;
  hold = { started: start, promise: new Promise((r) => (release = r)) };
  const pending = owner.previewFormula('0', 'Sheet1', 'E1', 'B2+1', { requestId });
  await started;
  const ack = await owner.cancelFormulaPreview('0', requestId);
  assert.deepEqual(ack, { schema: 1, requestId, cancelled: true });
  assert.equal(hold.reply.result.value, 11);
  release();
  const cancelled = await pending;
  assert.equal(cancelled.result.status, 'cancelled');
  hold = null;
  assert.equal((await owner.cancelFormulaPreview('0', requestId)).cancelled, false);
  await assert.rejects(() => owner.previewFormula('0', 'Sheet1', 'E1', '1', { requestId }), /already been used/);
  assert.deepEqual(await footprint(owner, source), before);
  assert.equal((await request(owner, '2')).result.value, 2);
  const current = await owner.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 5 }]),
    state = current.view.storage.state;
  const edited = await request(owner, 'B2+1', 'E1', state),
    old = await request(owner, 'B2+1');
  assert.equal(edited.result.value, 17);
  assert.equal(old.result.value, 11);
  const sourceView = owner.view(state);
  const raw = await native({
    action: 'workbook',
    operation: 'preview-formula',
    root,
    version: 999999,
    view: sourceView,
    sheet: 'Sheet1',
    cell: 'E1',
    formula: '1',
    requestId: randomUUID(),
  });
  assert.equal(raw.result.status, 'unavailable');
  const forged = structuredClone(sourceView);
  forged.namedRanges[0].formula = "'Other'!$A$1";
  const refusal = await native({
    action: 'workbook',
    operation: 'preview-formula',
    root,
    version: 0,
    view: forged,
    sheet: 'Sheet1',
    cell: 'E1',
    formula: 'Amount',
    requestId: randomUUID(),
  });
  assert.match(refusal.result.reason, /metadata differs/);
  const specialSource = path.join(output, 'special.xlsx');
  await fs.writeFile(specialSource, formulaPreviewCoreFixture({ special: true }));
  const special = make();
  await special.open(specialSource);
  for (const formula of ['K1', 'L2', 'SUM(K1:L2)'])
    assert.equal((await request(special, formula)).result.status, 'unavailable');
  assert.equal((await request(special, '1', 'L1')).result.status, 'unavailable');
  const cancelId = randomUUID(),
    marker = path.join(root, 'preview-cancel-' + cancelId);
  await fs.writeFile(marker, '1');
  const nativeCancelled = await native({
    action: 'workbook',
    operation: 'preview-formula',
    root,
    version: 0,
    view: special.view('0'),
    sheet: 'Sheet1',
    cell: 'E1',
    formula: 'SUM(A1:A3)',
    requestId: cancelId,
  });
  await fs.rm(marker);
  assert.equal(nativeCancelled.result.status, 'cancelled');
  assert.equal((await request(special, '1')).result.value, 1);
  for (const extra of [{ cells: {} }, { range: 'A1:A2' }, { overrides: { A1: 99 } }, { cancelFile: '/tmp/x' }])
    await assert.rejects(
      () =>
        native({
          action: 'workbook',
          operation: 'preview-formula',
          root,
          version: 0,
          view: special.view('0'),
          sheet: 'Sheet1',
          cell: 'E1',
          formula: '1',
          requestId: randomUUID(),
          ...extra,
        }),
      /Unexpected formula preview field/
    );
  const calendar = structuredClone(special.view('0'));
  calendar.sheets[0].date1904 = true;
  const mismatch = await native({
    action: 'workbook',
    operation: 'preview-formula',
    root,
    version: 0,
    view: calendar,
    sheet: 'Sheet1',
    cell: 'E1',
    formula: 'DATE(2000,1,1)',
    requestId: randomUUID(),
  });
  assert.match(mismatch.result.reason, /calendar differs/);

} finally {
  const closing = await Promise.allSettled(owners.map((owner) => owner.close()));
  client.close();
  await fs.rm(output, { recursive: true, force: true });
  for (const result of closing) if (result.status === 'rejected') throw result.reason;
}
