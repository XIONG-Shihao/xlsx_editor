/** Month-date lifecycle from actual readonly Excel numeric reads; no saved-Excel cache claim. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { editWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
// openpyxl emits an empty workbookProtection marker. The production owner
// conservatively protects imported markers; remove ONLY that empty node from a
// disposable test clone, preserving the actual captured source bytes/hash.
function ordinaryClone(bytes) {
  const parts = unzipSync(bytes),
    part = 'xl/workbook.xml',
    xml = strFromU8(parts[part]);
  assert.equal((xml.match(/<workbookProtection\/>/g) ?? []).length, 1);
  parts[part] = strToU8(xml.replace('<workbookProtection/>', ''));
  return Buffer.from(zipSync(parts));
}
async function owner(t, bytes) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-months-')),
    source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, bytes);
  const make = () =>
    new IndexedWorkbook(
      native,
      () => assert.fail('No legacy writer'),
      (input) => calculateIndexedWorkbook(input, native)
    );
  const book = make();
  await book.open(source);
  t.after(async () => {
    await book.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  return { book, make, source, bytes, root };
}
for (const system of [1900, 1904])
  test(`indexed ${system} monthly numeric/error results retain exact references through edit, old state, replay, export and reopen`, async (t) => {
    const oracle = JSON.parse(
        await fs.readFile(new URL(`../fixtures/excel-month-${system}-reference.json`, import.meta.url), 'utf8')
      ),
      source = await fs.readFile(new URL(`../fixtures/${oracle.source}`, import.meta.url));
    assert.equal(createHash('sha256').update(source).digest('hex'), oracle.sourceSha256);
    const start = system === 1904 ? 43844 : 45306,
      expected = system === 1904 ? [43875, 43889] : [45337, 45351];
    const bytes = editWorkbook(ordinaryClone(source), [
      { sheet: 'Cases', cell: 'C1', value: start },
      { sheet: 'Cases', cell: 'J1', value: 1 },
      { sheet: 'Cases', cell: 'D1', formula: 'EDATE(C1,J1)' },
      { sheet: 'Cases', cell: 'E1', formula: 'EOMONTH(C1,J1)' },
      { sheet: 'Cases', cell: 'F1', formula: 'D1+E1' },
      { sheet: 'Cases', cell: 'G1', formula: 'IFERROR(B22,99)' },
      { sheet: 'Cases', cell: 'H1', formula: 'DEC2HEX(D1,6)' },
    ]);
    const f = await owner(t, bytes);
    async function verify(book, state, pair = expected) {
      const cells = (await book.read(state, 'Cases', 'B1:J29')).cells;
      for (const [ref, item] of Object.entries(oracle.cells)) {
        const result = cells[ref].calculation;
        if (item.dataType === 'number') {
          assert.equal(result.status, 'calculated', ref);
          assert.equal(result.value, item.expected, ref);
        } else {
          assert.equal(result.status, 'error', ref);
          assert.equal(result.error, item.observedError, ref);
        }
      }
      assert.equal(cells.D1.calculation.value, pair[0]);
      assert.equal(cells.E1.calculation.value, pair[1]);
      assert.equal(cells.F1.calculation.value, pair[0] + pair[1]);
      assert.equal(cells.G1.calculation.value, 99);
      assert.equal(cells.H1.calculation.value, pair[0].toString(16).toUpperCase().padStart(6, '0'));
      return cells;
    }
    await verify(f.book, '0');
    const edit = [{ sheet: 'Cases', cell: 'J1', value: 2 }],
      changed = await f.book.apply('0', edit),
      state = changed.view.storage.state,
      pair = system === 1904 ? [43904, 43920] : [45366, 45382];
    await verify(f.book, state, pair);
    await verify(f.book, '0');
    const output = await f.book.export(state),
      saved = await fs.readFile(output.path);
    const disk = readWorkbook(saved);
    assert.equal(disk.sheets[0].date1904, system === 1904);
    const cells = await verify(f.book, state, pair);
    for (const [ref, item] of Object.entries(oracle.cells))
      if (item.dataType === 'number') assert.equal(disk.sheets[0].cells[ref].cachedValue, item.expected, ref);
    const replay = f.make();
    t.after(() => replay.close());
    await replay.open(f.source);
    await verify(replay, (await replay.apply('0', edit)).view.storage.state, pair);
    const reopened = f.make();
    t.after(() => reopened.close());
    await reopened.open(output.path);
    await verify(reopened, '0', pair);
    assert.deepEqual(new Uint8Array(await fs.readFile(f.source)), new Uint8Array(bytes));
    assert.deepEqual(
      new Uint8Array(await fs.readFile(new URL(`../fixtures/${oracle.source}`, import.meta.url))),
      new Uint8Array(source)
    );
    const independent = spawnSync(
      'python3',
      [
        '-c',
        `import sys,zipfile,xml.etree.ElementTree as E
n='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
with zipfile.ZipFile(sys.argv[1]) as z:
 c={x.get('r'):x for x in E.fromstring(z.read('xl/worksheets/sheet1.xml')).iter(n+'c')}
 for ref in ['D1','E1','F1']:assert float(c[ref].find(n+'v').text)==float(sys.argv[['D1','E1','F1'].index(ref)+2])
 assert c['D1'].find(n+'f').text=='EDATE(C1,J1)' and c['E1'].find(n+'f').text=='EOMONTH(C1,J1)'
 assert c['B22'].get('t')=='e' and c['B22'].find(n+'v').text=='#NUM!'
`,
        output.path,
        String(pair[0]),
        String(pair[1]),
        String(pair[0] + pair[1]),
      ],
      { encoding: 'utf8' }
    );
    assert.equal(independent.status, 0, independent.stderr);
  });
test('monthly shared formulas keep compact source membership and branch caches without stale-input reuse', async (t) => {
  const base = await fs.readFile(new URL('../fixtures/excel-month-1900-source.xlsx', import.meta.url)),
    bytes = editWorkbook(ordinaryClone(base), [
      { sheet: 'Cases', cell: 'C1', value: 45306 },
      { sheet: 'Cases', cell: 'C2', value: 45306 },
      { sheet: 'Cases', cell: 'C3', value: 45306 },
      { sheet: 'Cases', cell: 'J1', value: 1 },
      { sheet: 'Cases', cell: 'J2', value: 2 },
      { sheet: 'Cases', cell: 'J3', value: 3 },
    ]);
  const parts = unzipSync(bytes),
    xml = strFromU8(parts['xl/worksheets/sheet1.xml']);
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    xml.replace(
      /(<row\b[^>]*r="([123])"[^>]*>)([\s\S]*?)(<\/row>)/g,
      (_all, head, row, body, tail) =>
        head +
        body +
        `<c r="I${row}"><f t="shared" si="0"${row === '1' ? ' ref="I1:I3"' : ''}>${row === '1' ? 'EDATE(C1,J1)' : ''}</f><v>999</v></c>` +
        tail
    )
  );
  const f = await owner(t, Buffer.from(zipSync(parts))),
    initial = (await f.book.read('0', 'Cases', 'I1:I3')).cells;
  assert.deepEqual(
    [initial.I1.calculation.value, initial.I2.calculation.value, initial.I3.calculation.value],
    [45337, 45366, 45397]
  );
  const changed = await f.book.apply('0', [{ sheet: 'Cases', cell: 'J2', value: 1 }]),
    state = changed.view.storage.state;
  assert.equal((await f.book.read(state, 'Cases', 'I2')).cells.I2.calculation.value, 45337);
  assert.equal((await f.book.read('0', 'Cases', 'I2')).cells.I2.calculation.value, 45366);
  const output = await f.book.export(state),
    saved = await fs.readFile(output.path),
    sheet = strFromU8(unzipSync(saved)['xl/worksheets/sheet1.xml']);
  assert.equal((sheet.match(/t="shared"/g) ?? []).length, 3);
  assert.deepEqual(new Uint8Array(await fs.readFile(f.source)), new Uint8Array(f.bytes));
  const again = f.make();
  t.after(() => again.close());
  await again.open(output.path);
  assert.equal((await again.read('0', 'Cases', 'I2')).cells.I2.calculation.value, 45337);
});
