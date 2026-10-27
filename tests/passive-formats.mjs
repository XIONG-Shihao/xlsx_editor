import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strFromU8, strToU8 } from 'fflate';
import { PassiveWorkbook } from '../adapters/passive-workbook.mjs';
import { admitPassiveContainer, validatePassiveCfb } from '../core/passive-container.mjs';
import { passiveCell } from '../core/passive-model.mjs';
const fixtures = path.resolve('fixtures/remaining-formats'),
  sha = (bytes) => createHash('sha256').update(bytes).digest('hex'),
  formats = [
    ['xlsm', 'vba-user-function.xlsm', '20ae63044d0f5edbbdd522cc710a4ef0333e9391cc33aafed875287b91452d4c'],
    ['xls', 'universal-content.xls', '3f84bf513e191642835acbd56a40361bdcf1e4ef1a094b5a444b80f240fd889c'],
    ['xlsb', 'universal-content.xlsb', '3d425c433c4f0688d4d2f2c84bbcfc25494641d0d38cbe96d6a954d820536170'],
  ];
for (const [format, filename, expectedSha] of formats)
  test('authentic ' + format + ' keeps typed stored data and exact same-format original Copy', async (t) => {
    const owner = new PassiveWorkbook(),
      original = await fs.readFile(path.join(fixtures, filename)),
      root = await fs.mkdtemp(path.join(os.tmpdir(), 'passive-format-test-'));
    t.after(async () => {
      await owner.close();
      await fs.rm(root, { recursive: true, force: true });
    });
    assert.equal(sha(original), expectedSha);
    const opened = owner.open(path.join(fixtures, filename), { format, storageDirectory: root });
    await assert.rejects(owner.open(path.join(fixtures, filename), { format }), /already/);
    const view = await opened;
    assert(view.readOnly);
    assert.equal(view.format, format);
    assert.equal(view.originalFormat, format);
    assert.equal(view.storage.state, '0');
    assert(view.passiveNotice.includes('Read-only'));
    assert(view.sheets.every((sheet) => !Object.keys(sheet.cells).length));
    assert(!JSON.stringify(view).includes(root));
    const read = await owner.read('0', 'Sheet1', 'A1:C4'),
      query = await owner.query('0', 'Sheet1', format === 'xlsm' ? 'myfunc' : '2*3', 'formula');
    assert.equal(query.cells.A1?.formula ?? query.cells.C1.formula, format === 'xlsm' ? 'myfunc()' : '2*3');
    assert.equal(query.matches, 1);
    assert(query.scanned <= 8192);
    const scan = await owner.scan('0', 'Sheet1');
    assert(Object.keys(scan.cells).length <= 1000);
    assert.equal(scan.nextCursor, null);
    if (format === 'xlsm') {
      assert.equal(read.cells.A1.value, 42);
      assert.equal(read.cells.A1.calculation.status, 'stored');
      const modules = await owner.vbaModules('0', { offset: 0, limit: 2 });
      assert.equal(modules.total, 5);
      assert.equal(modules.modules.length, 2);
      assert.equal(modules.nextOffset, 2);
      const all = await owner.vbaModules('0'),
        module = all.modules.find((item) => item.name === 'Module1');
      let code = '',
        offset = 0;
      for (;;) {
        const page = await owner.vbaModule('0', module.id, { offset, limit: 7 });
        code += page.code;
        if (page.nextOffset === null) break;
        offset = page.nextOffset;
      }
      assert(code.includes('Function MyFunc()\r\n    MyFunc = 42\r\nEnd Function'));
      assert.equal(sha(Buffer.from(code)), module.sha256);
      assert.equal(module.sha256, '20954c4747f4f596157a4fde8c25c55db505190d77e518ef381b3561aa738a38');
    } else {
      assert.equal(read.cells.A1.value, 1);
      assert.equal(read.cells.A2.value, 2);
      assert.equal(read.cells.B1.value, 'String1');
      assert.equal(read.cells.B2.value, 'String2');
      for (const [ref, value] of [
        ['C1', 6],
        ['C2', 5],
        ['C3', -1],
        ['C4', 11],
      ]) {
        assert.equal(read.cells[ref].calculation.status, 'stored');
        assert.equal(read.cells[ref].calculation.value, value);
      }
      assert(view.sheets[0].merges.includes('E2:F3'));
      const summary = await owner.selectionSummary('0', 'Sheet1', 'A1:C4');
      assert.equal(summary.count, 8);
      assert.equal(summary.numericCount, 6);
      assert.equal(summary.sum, 24);
      assert.equal(summary.unresolved, 0);
      assert.equal((await owner.vbaModules('0')).total, 0);
    }
    await assert.rejects(owner.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 9 }]), /read-only/);
    await assert.rejects(owner.read('old', 'Sheet1', 'A1'), /stale/);
    await assert.rejects(owner.read('0', 'Sheet1', 'A1:B1000'), /budget/);
    await assert.rejects(owner.query('0', 'Sheet1', 'x', 'unknown'), /search/);
    await assert.rejects(owner.compact(['guessed']), /retain/);
    assert((await owner.compact()).unchanged);
    assert((await owner.objects('0', 'Sheet1')).protectedReason);
    const exported = await owner.export('0'),
      copied = path.join(root, 'copy.' + format);
    assert.equal(exported.revision, expectedSha);
    assert(exported.path.endsWith('.' + format));
    await fs.copyFile(exported.path, copied);
    assert.equal(sha(await fs.readFile(copied)), expectedSha);
    assert.equal(sha(await fs.readFile(path.join(fixtures, filename))), expectedSha);
    const privateRoot = owner.storageRoots()[0];
    await owner.close();
    await assert.rejects(owner.read('0', 'Sheet1', 'A1'), /closed/);
    await assert.rejects(fs.stat(privateRoot), /ENOENT/);
    assert.equal(owner.diagnosticStats().storageRoots, 0);
  });
test('module identities and stored cache types cannot cross owners or become fresh calculations', async (t) => {
  const owners = [new PassiveWorkbook(), new PassiveWorkbook()];
  t.after(() => Promise.all(owners.map((owner) => owner.close())));
  for (const owner of owners) await owner.open(path.join(fixtures, 'vba-user-function.xlsm'), { format: 'xlsm' });
  const id = (await owners[0].vbaModules('0')).modules[0].id;
  await assert.rejects(owners[1].vbaModule('0', id), /unknown module/);
  await assert.rejects(owners[0].vbaModule('old', id), /stale/);
  await assert.rejects(owners[0].vbaModule('0', id, { path: 'private' }), /fields/);
  await assert.rejects(owners[0].vbaModules('0', { limit: 101 }), /limit/);
  for (const source of [
    { t: 'b', v: 1, f: 'TRUE()' },
    { t: 'n', v: true, f: '1' },
    { t: 's', v: 7, f: '"7"' },
    { t: 'n', v: Infinity, f: '1/0' },
    { t: 'n', f: '1' },
  ]) {
    const cell = passiveCell(source, 'rev');
    assert.equal(cell.calculation.status, 'unresolved');
    assert.equal(cell.value, null);
  }
  assert.equal(passiveCell({ t: 's', v: '', f: '""' }, 'rev').calculation.value, '');
  assert.equal(passiveCell({ t: 'b', v: false, f: 'FALSE()' }, 'rev').calculation.value, false);
  assert(Object.is(passiveCell({ t: 'n', v: -0, f: '0' }, 'rev').calculation.value, -0));
});
test('authentic MIME/CRC/binary headers fence renamed, truncated and corrupt sources', async () => {
  const macro = await fs.readFile(path.join(fixtures, 'vba-user-function.xlsm')),
    binary = await fs.readFile(path.join(fixtures, 'universal-content.xlsb')),
    xls = await fs.readFile(path.join(fixtures, 'universal-content.xls'));
  assert.throws(() => admitPassiveContainer(macro, 'xlsb'), /MIME/);
  assert.throws(() => admitPassiveContainer(binary, 'xlsm'), /MIME/);
  assert.throws(() => admitPassiveContainer(macro.subarray(0, -1), 'xlsm'), /ZIP/);
  const badOle = Buffer.from(xls);
  badOle[28] = 0;
  assert.throws(() => admitPassiveContainer(badOle, 'xls'), /header/);
  const parts = unzipSync(macro);
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace(
      'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
    )
  );
  assert.throws(() => admitPassiveContainer(zipSync(parts), 'xlsm'), /MIME/);
  const corrupt = Buffer.from(macro);
  const signature = Buffer.from([0x50, 0x4b, 0x01, 0x02]),
    at = corrupt.indexOf(signature);
  assert(at > 0);
  corrupt[at + 16] ^= 1;
  assert.throws(() => admitPassiveContainer(corrupt, 'xlsm'), /checksum|disagree/);
});
test('close during preparation removes its owned snapshot and unexpected files retain private storage', async (t) => {
  const source = path.join(fixtures, 'vba-user-function.xlsm'),
    owner = new PassiveWorkbook(),
    opening = owner.open(source, { format: 'xlsm' });
  opening.catch(() => {});
  await owner.close();
  await assert.rejects(opening, /closed/);
  assert.equal(owner.storageRoots().length, 0);
  const other = new PassiveWorkbook();
  await other.open(source, { format: 'xlsm' });
  const root = other.storageRoots()[0];
  t.after(async () => {
    await other.close().catch(() => {});
    await fs.rm(root, { recursive: true, force: true });
  });
  await fs.writeFile(path.join(root, 'unexpected.bin'), 'preserve');
  await assert.rejects(other.close(), /unexpected/);
  assert.equal(await fs.readFile(path.join(root, 'unexpected.bin'), 'utf8'), 'preserve');
});
test('compound directory entry kinds, declared stream sizes and free-slot contents stay closed', () => {
  const valid = {
    FullPaths: ['Root Entry/', 'Root Entry/stream'],
    FileIndex: [{ type: 5 }, { type: 2, size: 3, content: Uint8Array.from([1, 2, 3]) }],
  };
  assert.equal(validatePassiveCfb(valid), valid);
  for (const change of [
    (entry) => (entry.type = 3),
    (entry) => (entry.size = 4),
    (entry) => {
      entry.type = 0;
    },
    (entry) => (entry.type = 5),
  ]) {
    const input = structuredClone(valid);
    change(input.FileIndex[1]);
    assert.throws(() => validatePassiveCfb(input), /compound/);
  }
});
