import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook, readWorkbook } from '../core/workbook-package.mjs';
import { WorkbookSession } from '../core/workbook-service.mjs';

function independentTexts(bytes) {
  // XML normalization and SpreadsheetML escape interpretation are independent
  // of the JavaScript parser/writer. The inspector decodes escapes once, preserving
  // escaped literal _x0041_, _x0001_ and _x005F_x0041_ text.
  const parsed = spawnSync(
    'python3',
    [
      '-c',
      `import sys,json,zipfile,io,re,xml.etree.ElementTree as E
z=zipfile.ZipFile(io.BytesIO(sys.stdin.buffer.read()))
r=E.fromstring(z.read('xl/worksheets/sheet1.xml'))
def decode(t):return re.sub(r'_x([0-9a-fA-F]{4})_',lambda m:chr(int(m.group(1),16)),t)
values={}
for c in r.findall('.//{*}c'):
 if c.find('{*}is') is not None:values[c.attrib['r']]=decode(''.join(c.find('{*}is').itertext()))
 elif c.attrib.get('t')=='str':values[c.attrib['r']]=decode(c.find('{*}v').text or '')
print(json.dumps(values))`,
    ],
    { input: bytes, maxBuffer: 1024 * 1024 }
  );
  assert.equal(parsed.status, 0, parsed.stderr.toString());
  return JSON.parse(parsed.stdout);
}
test('invalid XML characters and unpaired surrogates refuse an entire workbook batch', async () => {
  for (const value of ['A\u0000B', 'A\u0001B', 'A\u000bB', 'A\ufffeB', 'A\ud800B', 'A\udc00B']) {
    for (const edit of [{ value }, { formula: `"${value}"` }]) {
      const source = createWorkbook(),
        session = new WorkbookSession(source),
        before = await session.read();
      await assert.rejects(
        session.apply({
          operationId: 'invalid-text',
          expectedRevision: before.revision,
          commands: [
            { sheet: 'Sheet1', cell: 'A1', value: 'Would change first' },
            { sheet: 'Sheet1', cell: 'B1', ...edit },
          ],
        }),
        /character|Unicode|surrogate|XML|text/i
      );
      assert.strictEqual(await session.read(), before);
      assert.deepEqual(session.bytes, source);
    }
  }
});
test('saved legal whitespace, Unicode and literal SpreadsheetML escape patterns retain exact text', async () => {
  const texts = ['English / 中文 / 😀', 'A\tB\nC', 'A\rB', 'A\r\nB', '_x0041_', '_x0001_', '_x005F_x0041_'];
  const session = new WorkbookSession(createWorkbook()),
    before = await session.read();
  const result = await session.apply({
    operationId: 'text-roundtrip',
    expectedRevision: before.revision,
    commands: texts.map((value, index) => ({ sheet: 'Sheet1', cell: `A${index + 1}`, value })),
  });
  const independent = independentTexts(result.bytes),
    reopened = readWorkbook(result.bytes);
  for (const [index, value] of texts.entries()) {
    const cell = `A${index + 1}`;
    assert.equal(independent[cell], value, `Independent saved ${cell}`);
    assert.equal(result.view.sheets[0].cells[cell].value, value, `Accepted ${cell}`);
    assert.equal(reopened.sheets[0].cells[cell].value, value, `Reopened ${cell}`);
  }
});
test('reader decodes inline/shared string escapes once and preserves literal escape prefixes', () => {
  const parts = unzipSync(createWorkbook());
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>_x0041_</t></is></c><c r="B1" t="inlineStr"><is><t>_x005F_x0041_</t></is></c><c r="C1" t="s"><v>0</v></c><c r="D1" t="s"><v>1</v></c></row></sheetData></worksheet>`
  );
  parts['xl/sharedStrings.xml'] = strToU8(
    `<sst xmlns="${ns}" count="2" uniqueCount="2"><si><t>_x0041_</t></si><si><t>_x005F_x0041_</t></si></sst>`
  );
  parts['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
      '</Relationships>',
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'
    )
  );
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace(
      '</Types>',
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'
    )
  );
  const view = readWorkbook(zipSync(parts));
  for (const [ref, expected] of Object.entries({ A1: 'A', B1: '_x0041_', C1: 'A', D1: '_x0041_' }))
    assert.equal(view.sheets[0].cells[ref].value, expected, ref);
});
test('formula expressions keep literal escapes while saved text caches use SpreadsheetML encoding', async () => {
  const texts = ['_x0041_', '_x0001_', '_x005F_x0041_'];
  const session = new WorkbookSession(createWorkbook()),
    before = await session.read();
  const result = await session.apply({
    operationId: 'formula-text',
    expectedRevision: before.revision,
    commands: texts.map((value, index) => ({ sheet: 'Sheet1', cell: `A${index + 1}`, formula: `"${value}"` })),
  });
  const independent = independentTexts(result.bytes),
    reopened = readWorkbook(result.bytes);
  for (const [index, value] of texts.entries()) {
    const ref = `A${index + 1}`,
      accepted = result.view.sheets[0].cells[ref],
      saved = reopened.sheets[0].cells[ref];
    assert.equal(accepted.formula, `"${value}"`);
    assert.equal(saved.formula, accepted.formula);
    assert.equal(accepted.calculation.value, value);
    assert.equal(saved.value, value);
    assert.equal(independent[ref], value);
    assert.equal(accepted.rawValue, saved.rawValue);
  }
});
