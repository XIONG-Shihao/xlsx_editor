import assert from 'node:assert/strict';
import { test, after } from 'node:test';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { createWorkbook } from '../core/workbook-package.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';

const roots = [],
  clients = [];
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
const text = (ref, value) => `<c r="${ref}" t="inlineStr"><is><t>${value}</t></is></c>`;
const number = (ref, value) => `<c r="${ref}"><v>${value}</v></c>`;
function bytes(body) {
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet xmlns="${ns}"><sheetData>${body}</sheetData></worksheet>`);
  parts['custom/dedup-opaque.bin'] = new Uint8Array([7, 0, 255]);
  return zipSync(parts);
}
async function fixture(body) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-dedup-preview-'));
  roots.push(root);
  const source = path.join(root, 'source.xlsx'),
    original = Buffer.from(bytes(body));
  await fs.writeFile(source, original);
  const client = createWorkbookNativeClient(nativeWorkbookBinary());
  clients.push(client);
  const call = (request) => client.call({ action: 'workbook', root, ...request });
  const index = await call({ operation: 'index' }),
    sheet = index.sheets[0].path;
  const preview = (options) =>
    call({
      operation: 'dedup-preview',
      sheet,
      version: 0,
      range: 'A1:B30',
      hasHeaders: true,
      columns: ['A'],
      ...options,
    });
  const read = (version) => call({ operation: 'range', version, sheet, range: 'A1:B30' });
  return { root, source, original, call, sheet, preview, read };
}
function check(name, fn) {
  test(name, async () => {
    const result = await fn();
  });
}
after(async () => {
  for (const client of clients) client.close();
  for (const root of roots) await fs.rm(root, { recursive: true, force: true });
});

check('native exact typed duplicates retain source decimal precision and stable first occurrences', async () => {
  const cells = [
    number('A2', '1'),
    text('A3', '1'),
    '<c r="A4" t="b"><v>1</v></c>',
    text('A5', ''),
    '',
    text('A7', '#N/A'),
    '<c r="A8" t="e"><v>#N/A</v></c>',
    number('A9', '1.00'),
    number('A10', '9007199254740992'),
    number('A11', '9007199254740993'),
    number('A12', '9007199254740992'),
    text('A13', '1'),
    '<c r="A14" t="b"><v>1</v></c>',
    '',
    text('A16', ''),
    '<c r="A17" t="e"><v>#N/A</v></c>',
    text('A18', '#N/A'),
    number('A19', '-0'),
    number('A20', '0'),
    number('A21', '0.1234567890123456789'),
    number('A22', '0.1234567890123456788'),
  ];
  const body =
    `<row r="1">${text('A1', 'Key')}${text('B1', 'Record')}</row>` +
    cells.map((cell, i) => `<row r="${i + 2}">${cell}${text('B' + (i + 2), 'row-' + (i + 2))}</row>`).join('');
  const f = await fixture(body),
    before = await f.read(0),
    db = await fs.readFile(path.join(f.root, 'cells.sqlite'));
  const result = await f.preview({ range: 'A1:B22' });
  assert.equal(result.totalRows, 21);
  assert.equal(result.keptRows, 12);
  assert.equal(result.removedRows, 9);
  assert.deepEqual(result.samples, [
    { row: 9, keptRow: 2 },
    { row: 12, keptRow: 10 },
    { row: 13, keptRow: 3 },
    { row: 14, keptRow: 4 },
    { row: 15, keptRow: 6 },
    { row: 16, keptRow: 5 },
    { row: 17, keptRow: 8 },
    { row: 18, keptRow: 7 },
    { row: 20, keptRow: 19 },
  ]);
  assert.equal(result.comparison, 'exact-typed-values-v1');
  assert.match(result.fingerprint, /^[a-f0-9]{64}$/);
  assert.deepEqual(await f.read(0), before);
  assert.deepEqual(await fs.readFile(f.source), f.original);
  assert.deepEqual(await fs.readFile(path.join(f.root, 'cells.sqlite')), db);
  return { ...result, sourceAndIndexUnchanged: true, unsafeDecimalIdentitiesRetained: true };
});
check('multicolumn keys and explicit header choice determine complete records', async () => {
  const f = await fixture(
    `<row r="1">${text('A1', 'x')}${number('B1', '1')}</row><row r="2">${text('A2', 'x')}${number('B2', '1.0')}</row><row r="3">${text('A3', 'x')}${number('B3', '2')}</row><row r="4">${text('A4', 'x')}${number('B4', '1')}</row>`
  );
  const both = await f.preview({ range: 'A1:B4', hasHeaders: false, columns: ['A', 'B'] });
  assert.equal(both.keptRows, 2);
  assert.deepEqual(both.samples, [
    { row: 2, keptRow: 1 },
    { row: 4, keptRow: 1 },
  ]);
  const header = await f.preview({ range: 'A1:B4', columns: ['A', 'B'] });
  assert.equal(header.totalRows, 3);
  assert.deepEqual(header.samples, [{ row: 4, keptRow: 2 }]);
  const one = await f.preview({ range: 'A1:B4', hasHeaders: false });
  assert.equal(one.keptRows, 1);
  assert.deepEqual(one.samples, [
    { row: 2, keptRow: 1 },
    { row: 3, keptRow: 1 },
    { row: 4, keptRow: 1 },
  ]);
  assert.notEqual(both.fingerprint, header.fingerprint);
  assert.notEqual(both.fingerprint, one.fingerprint);
  return { allColumns: both, headerExcluded: header, oneColumn: one };
});
check('preview pages and sparse blank rows share one deterministic receipt', async () => {
  const f = await fixture(`<row r="1">${text('A1', 'Key')}</row><row r="2">${number('A2', '1')}</row>`);
  const first = await f.preview({ range: 'A1:B220', limit: 100 });
  assert.equal(first.totalRows, 219);
  assert.equal(first.keptRows, 2);
  assert.equal(first.removedRows, 217);
  assert.equal(first.samples.length, 100);
  assert.deepEqual(first.samples[0], { row: 4, keptRow: 3 });
  assert.equal(first.nextOffset, 100);
  const next = await f.preview({ range: 'A1:B220', offset: 100, limit: 100 }),
    tail = await f.preview({ range: 'A1:B220', offset: 200, limit: 100 });
  assert.equal(next.fingerprint, first.fingerprint);
  assert.equal(tail.fingerprint, first.fingerprint);
  assert.equal(next.samples[0].row, 104);
  assert.equal(tail.samples.length, 17);
  assert.equal(tail.nextOffset, null);
  return {
    rows: first.totalRows,
    kept: first.keptRows,
    removed: first.removedRows,
    pages: [first.samples.length, next.samples.length, tail.samples.length],
    sameReceipt: true,
  };
});
check('accepted branch edits change the receipt without changing the original version', async () => {
  const f = await fixture(
    `<row r="1">${text('A1', 'Key')}</row><row r="2">${number('A2', '1')}</row><row r="3">${number('A3', '1')}</row>`
  );
  const original = await f.preview({ range: 'A1:B3' }),
    patch = await f.call({
      operation: 'patch',
      version: 0,
      changes: [
        {
          sheet: f.sheet,
          cell: 'A3',
          value: { value: 2, dataType: 'number', rawValue: '2', cachedValue: null, styleId: 0 },
        },
      ],
    });
  const branch = await f.preview({ range: 'A1:B3', version: patch.version });
  assert.equal(branch.removedRows, 0);
  assert.notEqual(branch.fingerprint, original.fingerprint);
  assert.deepEqual(await f.preview({ range: 'A1:B3' }), original);
  assert.deepEqual(await fs.readFile(f.source), f.original);
  return { original, branch, originalVersionAndSourceRetained: true };
});
check('invalid scope, source changes and pending formula keys refuse without publication', async () => {
  const f = await fixture(`<row r="1">${text('A1', 'Key')}</row><row r="2">${number('A2', '1')}</row>`),
    before = await f.read(0),
    indexHash = digest(await fs.readFile(path.join(f.root, 'cells.sqlite')));
  const invalid = [
    { columns: [] },
    { columns: ['A', 'A'] },
    { columns: ['C'] },
    { columns: ['a'] },
    { columns: ['$A'] },
    { hasHeaders: 'true' },
    { range: 'A1', hasHeaders: true },
    { range: 'A1:XFD1048576' },
    { range: 'B3:A1' },
    { version: 999 },
    { sheet: '../unknown.xml' },
    { limit: 101 },
    { limit: 0 },
    { offset: -1 },
  ];
  for (const patch of invalid) await assert.rejects(f.preview(patch));
  assert.deepEqual(await f.read(0), before);
  assert.equal(digest(await fs.readFile(path.join(f.root, 'cells.sqlite'))), indexHash);
  assert.deepEqual(await fs.readFile(f.source), f.original);
  await fs.appendFile(f.source, Buffer.from([0]));
  await assert.rejects(f.preview(), /source changed/i);
  await fs.writeFile(f.source, f.original);
  const pending = await fixture(
    `<row r="1">${text('A1', 'Key')}</row><row r="2"><c r="A2"><f>1+1</f><v>99</v></c></row>`
  );
  await assert.rejects(pending.preview({ range: 'A1:B2' }), /pending|unresolved|Calculate/i);
  return {
    invalidScopes: invalid.length,
    changedSourceRefused: true,
    pendingSourceCacheNeverUsed: true,
    sourceAndIndexUnchanged: true,
  };
});
