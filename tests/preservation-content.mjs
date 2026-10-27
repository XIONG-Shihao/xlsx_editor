/** Actual native and compatibility writers: imported text metadata must survive or refuse. */
import { test, after } from 'node:test';
import { spawnSync } from 'node:child_process';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, editWorkbookResult, readWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { createWorkbookNativeClient } from '../core/native-client.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  rich = '<r><rPr><b/></rPr><t>Bold </t></r><r><t>plain</t></r>',
  annotated = '<t>Tokyo</t><rPh sb="0" eb="5"><t>phonetic</t></rPh>',
  opaque = '<t xmlns:u="urn:roundtrip:test" u:tag="keep">Plain</t>',
  variants = {
    rich,
    annotated,
    opaque,
    comment: '<t>Plain</t><!--retain-->',
    empty: '<t xmlns:u="urn:roundtrip:test" u:tag="keep"/>',
  };
const helper =
  process.env.XLSX_EDITOR_PRESERVATION_HELPER ??
  new URL('../native/bin/darwin-arm64/workbook-engine', import.meta.url).pathname;
const client = createWorkbookNativeClient(helper),
  cases = [];
const prefix = new URL('../.build/runs/compatibility/', import.meta.url).pathname;
await fs.mkdir(prefix, { recursive: true });
const evidence = await fs.mkdtemp(path.join(prefix, 'content-'));
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
function fixture(kind, storage = 'inline') {
  const parts = unzipSync(createWorkbook()),
    a =
      storage === 'shared' ? '<c r="A1" t="s"><v>0</v></c>' : `<c r="A1" t="inlineStr"><is>${variants[kind]}</is></c>`;
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${NS}" xmlns:u="urn:roundtrip:test"><dimension ref="A1:B1"/><sheetData><row r="1" u:marker="row">${a}<c r="B1" u:marker="cell"><v>10</v><extLst><ext uri="urn:roundtrip:test"><u:payload>Keep me</u:payload></ext></extLst></c></row></sheetData><extLst><ext uri="urn:roundtrip:test"><u:payload>Outside cells</u:payload></ext></extLst></worksheet>`
  );
  parts['customXml/opaque.bin'] = new Uint8Array([0, 255, 11, 99]);
  if (storage === 'shared') {
    parts['xl/sharedStrings.xml'] = strToU8(
      `<sst xmlns="${NS}" count="1" uniqueCount="1"><si>${variants[kind]}</si></sst>`
    );
    parts['[Content_Types].xml'] = strToU8(
      strFromU8(parts['[Content_Types].xml']).replace(
        '</Types>',
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'
      )
    );
    parts['xl/_rels/workbook.xml.rels'] = strToU8(
      strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
        '</Relationships>',
        '<Relationship Id="sst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'
      )
    );
  }
  return Buffer.from(zipSync(parts));
}
async function checked(name, run) {
  try {
    await run();
    cases.push({ name, passed: true });
  } catch (error) {
    cases.push({ name, passed: false, error: String(error) });
    throw error;
  }
}
for (const storage of ['inline', 'shared'])
  for (const kind of Object.keys(variants)) {
    test(`compatibility writer refuses ${storage} ${kind} content replacement atomically`, () =>
      checked(`compatibility-${storage}-${kind}`, async () => {
        const source = fixture(kind, storage),
          before = digest(source);
        assert.throws(
          () =>
            editWorkbookResult(source, [
              { sheet: 'Sheet1', cell: 'B1', value: 14 },
              { sheet: 'Sheet1', cell: 'A1', value: 'changed' },
            ]),
          /rich|metadata|preservation/i
        );
        assert.equal(digest(source), before);
      }));
    test(`native writer refuses ${storage} ${kind} and retains unrelated edits/formatting`, (t) =>
      checked(`native-${storage}-${kind}`, async () => {
        const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-preserve-')),
          file = path.join(dir, 'source.xlsx'),
          source = fixture(kind, storage);
        await fs.writeFile(file, source);
        const book = new IndexedWorkbook(
          client.call,
          () => assert.fail('No compatibility fallback'),
          () => assert.fail('No formula calculation')
        );
        try {
          await book.open(file);
          const before = book.view('0');
          await assert.rejects(
            book.apply('0', [
              { sheet: 'Sheet1', cell: 'B1', value: 14 },
              { sheet: 'Sheet1', cell: 'A1', value: 'changed' },
            ]),
            /rich|metadata|preservation/i
          );
          assert.equal(book.view('0').storage.state, before.storage.state);
          assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.value, 10);
          const plain = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', value: '00123' }]);
          const formatted = await book.apply(plain.view.storage.state, [{ sheet: 'Sheet1', cell: 'A1', bold: true }]);
          const output = await book.export(formatted.view.storage.state),
            saved = await fs.readFile(output.path),
            parts = unzipSync(saved),
            original = unzipSync(source);
          assert.equal(readWorkbook(saved).sheets[0].cells.B1.value, '00123');
          const valuePart = storage === 'shared' ? 'xl/sharedStrings.xml' : 'xl/worksheets/sheet1.xml';
          assert(strFromU8(parts[valuePart]).includes(variants[kind]));
          assert(strFromU8(parts['xl/worksheets/sheet1.xml']).includes('u:marker="cell"'));
          assert(strFromU8(parts['xl/worksheets/sheet1.xml']).includes('<u:payload>Keep me</u:payload>'));
          assert.deepEqual(parts['customXml/opaque.bin'], original['customXml/opaque.bin']);
          assert.deepEqual(await fs.readFile(file), source);
          await fs.writeFile(path.join(evidence, `${storage}-${kind}-source.xlsx`), source);
          await fs.writeFile(path.join(evidence, `${storage}-${kind}-saved.xlsx`), saved);
          const audit = spawnSync(
            'python3',
            [new URL('../scripts/test/xlsx-package-check.py', import.meta.url).pathname],
            {
              input: JSON.stringify({
                source: path.join(evidence, `${storage}-${kind}-source.xlsx`),
                saved: path.join(evidence, `${storage}-${kind}-saved.xlsx`),
                targets: { Sheet1: ['A1', 'B1'] },
                styleTargets: { Sheet1: ['A1'] },
                appendStyles: true,
                strictPointEdits: true,
                expected: { 'Sheet1!B1': { value: '00123' } },
              }),
              encoding: 'utf8',
            }
          );
          assert.equal(audit.status, 0, audit.stdout || audit.stderr);
          await fs.writeFile(path.join(evidence, `${storage}-${kind}-oracle.json`), audit.stdout);
        } finally {
          await book.close();
          await fs.rm(dir, { recursive: true, force: true });
        }
      }));
  }

for (const payload of [
  '<v xmlns:u="urn:roundtrip:test" u:tag="keep">10</v>',
  '<f xmlns:u="urn:roundtrip:test" u:tag="keep">2+8</f><v>10</v>',
])
  test('native and compatibility refuse opaque value/formula content fields ' + payload.slice(1, 2), () =>
    checked('opaque-' + payload.slice(1, 2), async () => {
      const parts = unzipSync(fixture('rich'));
      parts['xl/worksheets/sheet1.xml'] = strToU8(
        `<worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1">${payload}</c><c r="B1"><v>10</v></c></row></sheetData></worksheet>`
      );
      const source = Buffer.from(zipSync(parts)),
        file = path.join(evidence, 'opaque-' + payload.slice(1, 2) + '.xlsx');
      await fs.writeFile(file, source);
      assert.throws(
        () => editWorkbookResult(source, [{ sheet: 'Sheet1', cell: 'A1', value: 11 }]),
        /metadata|preservation/i
      );
      const book = new IndexedWorkbook(
        client.call,
        () => assert.fail('fallback'),
        (input) => calculateIndexedWorkbook(input, client.call)
      );
      try {
        await book.open(file);
        await assert.rejects(
          book.apply('0', [
            { sheet: 'Sheet1', cell: 'B1', value: 15 },
            { sheet: 'Sheet1', cell: 'A1', value: 11 },
          ]),
          /metadata|preservation/i
        );
        assert.equal((await book.read('0', 'Sheet1', 'B1')).cells.B1.value, 10);
      } finally {
        await book.close();
      }
    })
  );
test('opaque formula cache preserves results, permits unrelated literals and protects formula content', () =>
  checked('opaque-cache-readonly', async () => {
    const parts = unzipSync(fixture('rich'));
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      `<worksheet xmlns="${NS}"><sheetData><row r="1"><c r="A1"><f>2+8</f><v xmlns:u="urn:roundtrip:test" u:tag="keep">10</v></c><c r="B1"><v>4</v></c></row></sheetData></worksheet>`
    );
    const source = Buffer.from(zipSync(parts)),
      file = path.join(evidence, 'opaque-cache.xlsx');
    await fs.writeFile(file, source);
    const book = new IndexedWorkbook(
      client.call,
      () => assert.fail('fallback'),
      () => assert.fail('calculation must not replace opaque cache')
    );
    try {
      await book.open(file);
      assert.equal(book.view('0').readOnly, false);
      assert.equal(book.view('0').calculationPolicy.preserveSource, true);
      await assert.rejects(book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 11 }]), /protect/);
      assert.deepEqual(await fs.readFile((await book.export('0')).path), source);
      const next = await book.apply('0', [{ sheet: 'Sheet1', cell: 'B1', value: 8 }]);
      assert(
        strFromU8(
          unzipSync(await fs.readFile((await book.export(next.view.storage.state)).path))['xl/worksheets/sheet1.xml']
        ).includes('<v xmlns:u="urn:roundtrip:test" u:tag="keep">10</v>')
      );
    } finally {
      await book.close();
    }
  }));
test('rich copy refuses but explicit Paste Values extracts plain text without changing source', () =>
  checked('paste-values-boundary', async () => {
    const file = path.join(evidence, 'bulk-source.xlsx');
    await fs.writeFile(file, fixture('rich'));
    const book = new IndexedWorkbook(
      client.call,
      () => assert.fail('fallback'),
      (input) => calculateIndexedWorkbook(input, client.call)
    );
    try {
      await book.open(file);
      await assert.rejects(
        book.apply('0', [{ type: 'fillRange', sheet: 'Sheet1', source: 'A1', range: 'B1' }]),
        /metadata|preservation/i
      );
      const pasted = await book.apply('0', [{ type: 'pasteValues', sheet: 'Sheet1', source: 'A1', range: 'B1' }]);
      const bytes = await fs.readFile((await book.export(pasted.view.storage.state)).path);
      assert.equal(readWorkbook(bytes).sheets[0].cells.B1.value, 'Bold plain');
      assert(strFromU8(unzipSync(bytes)['xl/worksheets/sheet1.xml']).includes(rich));
    } finally {
      await book.close();
    }
  }));
test('strict independent oracle rejects target metadata loss and authority shortcuts', () =>
  checked('oracle-negative-controls', async () => {
    const original = fixture('rich'),
      source = path.join(evidence, 'oracle-source.xlsx');
    await fs.writeFile(source, original);
    for (const [name, mutate, extra] of [
      [
        'envelope',
        (p) => {
          p['xl/worksheets/sheet1.xml'] = strToU8(
            strFromU8(p['xl/worksheets/sheet1.xml']).replace('u:marker="cell"', '')
          );
        },
        {},
      ],
      [
        'rich',
        (p) => {
          p['xl/worksheets/sheet1.xml'] = strToU8(
            strFromU8(p['xl/worksheets/sheet1.xml']).replace(rich, '<t>Bold plain</t>')
          );
        },
        {},
      ],
      [
        'opaque',
        (p) => {
          p['customXml/opaque.bin'] = new Uint8Array([1]);
        },
        {},
      ],
      [
        'allow-bypass',
        (p) => {
          p['customXml/opaque.bin'] = new Uint8Array([1]);
        },
        { allowedChanged: ['customXml/opaque.bin'] },
      ],
      [
        'structural-bypass',
        (p) => {
          p['xl/worksheets/sheet1.xml'] = strToU8(
            strFromU8(p['xl/worksheets/sheet1.xml']).replace('u:marker="cell"', '')
          );
        },
        { structural: { Sheet1: { index: 0, count: 1 } } },
      ],
    ]) {
      const parts = unzipSync(original);
      mutate(parts);
      const saved = path.join(evidence, 'oracle-' + name + '.xlsx');
      await fs.writeFile(saved, zipSync(parts));
      const audit = spawnSync('python3', [new URL('../scripts/test/xlsx-package-check.py', import.meta.url).pathname], {
        input: JSON.stringify({ source, saved, strictPointEdits: true, targets: { Sheet1: ['A1', 'B1'] }, ...extra }),
        encoding: 'utf8',
      });
      assert.equal(audit.status, 1, audit.stdout || audit.stderr);
      assert.equal(JSON.parse(audit.stdout).passed, false);
    }
  }));

after(async () => {
  client.close();
});
