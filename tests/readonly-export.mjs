import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const id = '11111111-1111-4111-8111-111111111111',
  hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
async function fixture(t) {
  const directory = await fs.realpath(await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-readonly-export-'))),
    source = path.join(directory, 'source.xlsx'),
    owners = [];
  const make = (bridge = native) => {
    const book = new IndexedWorkbook(
      bridge,
      () => assert.fail('No compatibility writer'),
      (input) => calculateIndexedWorkbook(input, bridge)
    );
    owners.push(book);
    return book;
  };
  t.after(async () => {
    try {
      for (const owner of owners) await owner.close();
    } finally {
      await fs.rm(directory, { recursive: true, force: true });
    }
  });
  await fs.copyFile(path.resolve('tests/support/w7-reference-data/source.xlsx'), source);
  const author = make();
  await author.open(source);
  const created = await author.apply('0', [
      {
        operation: 'createChart',
        sheet: 'Sales',
        objectId: id,
        expectedObjectRevision: null,
        chart: { kind: 'column', range: 'A1:C4', title: 'Passive cache test' },
      },
    ]),
    saved = await author.export(created.view.storage.state),
    parts = unzipSync(await fs.readFile(saved.path));
  const sheet = strFromU8(parts['xl/worksheets/sheet1.xml']);
  assert.match(sheet, /<c\b[^>]*r="B2"/);
  const changed = sheet.replace(/(<c\b[^>]*r="B2"[^>]*>[\s\S]*?<v(?:\s[^>]*)?>)[^<]*(<\/v>)/, '$1' + '99' + '$2');
  assert.notEqual(changed, sheet);
  parts['xl/worksheets/sheet1.xml'] = strToU8(changed);
  parts['custom/passive.bin'] = new Uint8Array([255, 0, 19, 31]);
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace(
      '</Types>',
      // An ActiveX part keeps a workbook read-only (a macro project no longer does).
      '<Override PartName="/custom/passive.bin" ContentType="application/vnd.ms-office.activeX+xml"/></Types>'
    )
  );
  const bytes = Buffer.from(zipSync(parts));
  await fs.writeFile(source, bytes);
  await author.close();
  return { directory, source, bytes, parts, make, owners };
}
test('read-only managed chart displays current accepted values but export preserves the exact original stale cache/package bytes', async (t) => {
  const f = await fixture(t),
    calls = [],
    bridge = async (input) => {
      calls.push(input.operation);
      return native(input);
    },
    book = f.make(bridge),
    view = await book.open(f.source);
  assert(view.readOnly);
  const preview = await book.object(view.storage.state, 'Sales', id);
  assert.equal(preview.data.series[0].values[0], 99);
  const before = calls.length,
    output = await book.export(view.storage.state);
  assert.equal(output.revision, hash(f.bytes));
  assert.deepEqual(await fs.readFile(output.path), f.bytes);
  assert(!calls.slice(before).some((operation) => ['export-plan', 'artifact-write', 'export'].includes(operation)));
  await assert.rejects(book.apply(view.storage.state, [{ sheet: 'Sales', cell: 'B2', value: 23 }]), /read-only/);
  assert.equal(hash(await fs.readFile(f.source)), hash(f.bytes));
});
test('read-only Document Save Copy is byte-identical and cannot publish a fresh managed chart cache', async (t) => {
  const f = await fixture(t),
    owner = new SpreadsheetDocument({ recoveryDirectory: path.join(f.directory, 'recovery') });
  f.owners.push(owner);
  const current = await owner.open(f.source),
    destination = path.join(f.directory, 'copy.xlsx');
  assert(current.view.readOnly);
  await owner.request({ operation: 'copy', destination, expectedRevision: null });
  assert.deepEqual(await fs.readFile(destination), f.bytes);
  assert.equal(hash(await fs.readFile(f.source)), hash(f.bytes));
  const original = unzipSync(f.bytes),
    copied = unzipSync(await fs.readFile(destination));
  assert.deepEqual(Object.keys(copied).sort(), Object.keys(original).sort());
  for (const part of Object.keys(original)) assert.deepEqual(copied[part], original[part], part);
});

test('inherited Automatic and Manual policy admission separates unqualified shared calculations from ordinary editing in Document open/activate/copy', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-shared-policy-')),
    source = path.join(directory, 'source.xlsx');
  const { createWorkbook } = await import('../core/workbook-package.mjs');
  const parts = unzipSync(createWorkbook());
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f t="shared" si="0" ref="A1:A2" ca="1">B1+1</f><v>2</v></c><c r="B1"><v>1</v></c></row><row r="2"><c r="A2"><f t="shared" si="1" ref="A2:A3">B2+2</f><v>3</v></c><c r="B2"><v>1</v></c></row></sheetData></worksheet>'
  );
  const bytes = Buffer.from(zipSync(parts));
  await fs.writeFile(source, bytes);
  const owners = [];
  t.after(async () => {
    for (const owner of owners) await owner.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  for (const mode of ['auto', 'manual']) {
    const owner = new SpreadsheetDocument({
      recoveryDirectory: path.join(directory, mode),
      cacheIdentity: null,
      initialCalculationPolicy: { mode, recalculateBeforeSave: true },
    });
    owners.push(owner);
    const opened = await owner.open(source, { prepared: true });
    assert.equal(opened.view.readOnly, false);
    assert.equal(opened.view.calculationPolicy.preserveSource, true);
    assert.match(opened.view.calculationPolicy.reason, /stored|Stored/);
    assert.equal(opened.view.calculationPolicy.canRecalculate, false);
    const active = await owner.request({ operation: 'activate' });
    assert.equal(active.view.readOnly, false);
    const before = structuredClone(owner.status());
    await assert.rejects(
      owner.request({
        operation: 'apply',
        state: active.state,
        commands: [{ sheet: 'Sheet1', cell: 'A1', value: 99 }],
      }),
      /protect/
    );
    assert.deepEqual(owner.status(), before);
    const destination = path.join(directory, mode + '-copy.xlsx');
    await owner.request({ operation: 'copy', destination, expectedRevision: null });
    assert.deepEqual(await fs.readFile(destination), bytes);
    assert.deepEqual(await fs.readFile(source), bytes);
  }
});
