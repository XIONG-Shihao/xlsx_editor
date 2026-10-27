import assert from 'node:assert/strict';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { createWorkbook, readWorkbook, editWorkbook, MAX_WORKBOOK_BYTES } from '@office-editors/xlsx-editor/engine';
let parts = unzipSync(createWorkbook());
const sheet = strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
  '<sheetData/>',
  '<cols><col min="1" max="1" width="24" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>00123</t></is></c><c r="B1"><v>45000</v></c><c r="C1"><f>B1+1</f><v>45001</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A3:C3"/></mergeCells>'
);
parts['xl/worksheets/sheet1.xml'] = strToU8(sheet);
parts['xl/worksheets/sheet2.xml'] = strToU8(sheet);
parts['xl/workbook.xml'] = strToU8(
  strFromU8(parts['xl/workbook.xml']).replace(
    '</sheets>',
    '<sheet name="Hidden data" sheetId="2" r:id="rId3" state="hidden"/></sheets>'
  )
);
parts['xl/_rels/workbook.xml.rels'] = strToU8(
  strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
    '</Relationships>',
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>'
  )
);
parts['[Content_Types].xml'] = strToU8(
  strFromU8(parts['[Content_Types].xml']).replace(
    '</Types>',
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
  )
);
parts['xl/media/picture.png'] = Uint8Array.of(137, 80, 78, 71, 23, 42);
parts['customXml/item1.xml'] = strToU8('<custom>Keep this exactly</custom>');
let bytes = zipSync(parts);
const original = readWorkbook(bytes);
assert.equal(original.sheets.length, 2);
assert.equal(original.sheets[0].cells.A1.value, '00123');
assert.equal(original.sheets[0].cells.C1.value, 45001);
bytes = editWorkbook(
  bytes,
  [
    { sheet: 'Sheet1', cell: 'B1', value: 46000, numberFormat: 'date', bold: true },
    { sheet: 'Sheet1', cell: 'C1', formula: 'B1+2' },
  ],
  original.revision
);
const after = readWorkbook(bytes),
  updated = unzipSync(bytes);
assert.equal(after.sheets[0].cells.B1.value, 46000);
assert.equal(after.sheets[0].cells.B1.format, 14);
assert.equal(after.sheets[0].cells.B1.bold, true);
assert.equal(after.sheets[0].cells.C1.formula, 'B1+2');
assert.equal(after.sheets[0].cells.A1.value, '00123');
assert.equal(after.sheets[1].state, 'hidden');
assert.deepEqual(after.sheets[0].merges, ['A3:C3']);
assert.equal(after.sheets[0].widths[0].width, 24);
for (const path of ['xl/media/picture.png', 'customXml/item1.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels'])
  assert.deepEqual(updated[path], parts[path], path);
assert.throws(() => editWorkbook(bytes, [], original.revision), /Conflict/);
assert.throws(
  () => editWorkbook(bytes, [{ sheet: 'Sheet1', operation: 'insertRow', index: 1 }]),
  /Row\/column changes/
);
// A macro project does not make a workbook read-only: it is carried byte for
// byte, as Excel does when it edits a macro workbook.
{
  const macros = strToU8('macro project bytes'),
    withMacros = zipSync({ ...parts, 'xl/vbaProject.bin': macros });
  assert.equal(readWorkbook(withMacros).readOnly, false);
  const edited = editWorkbook(withMacros, [{ sheet: 'Sheet1', cell: 'A1', value: 'changed' }]);
  assert.equal(readWorkbook(edited).sheets[0].cells.A1.value, 'changed');
  assert.deepEqual(unzipSync(edited)['xl/vbaProject.bin'], macros, 'the macro part is unchanged');
}
// ActiveX controls still make a workbook read-only: an edit could move them.
for (const path of ['xl/activeX/activeX1.bin']) {
  const unsupported = zipSync({ ...parts, [path]: strToU8('unsupported') });
  assert.equal(readWorkbook(unsupported).readOnly, true);
  assert.throws(() => editWorkbook(unsupported, [{ sheet: 'Sheet1', cell: 'A1', value: 'changed' }]), /Read-only/);
}
for (const path of [
  'xl/externalLinks/externalLink1.xml',
  'xl/pivotTables/pivotTable1.xml',
  'xl/embeddings/opaque.bin',
]) {
  const originalParts = { ...parts, [path]: strToU8('opaque bytes retained without interpretation') };
  const supported = zipSync(originalParts);
  assert.equal(readWorkbook(supported).readOnly, false);
  const changed = editWorkbook(supported, [{ sheet: 'Sheet1', cell: 'A1', expectedValue: '00123', value: 'changed' }]);
  assert.equal(readWorkbook(changed).sheets[0].cells.A1.value, 'changed');
  assert.deepEqual(unzipSync(changed)[path], originalParts[path], `${path} literal edit must preserve opaque content`);
  assert.throws(
    () => editWorkbook(supported, [{ sheet: 'Sheet1', operation: 'insertRow', index: 1 }]),
    /Row\/column|preserve|qualified/
  );
}
let simple = editWorkbook(createWorkbook(), [
  { sheet: 'Sheet1', cell: 'A1', value: 'first' },
  { sheet: 'Sheet1', cell: 'A2', value: 'second' },
]);
simple = editWorkbook(simple, [{ sheet: 'Sheet1', operation: 'insertRow', index: 0 }]);
assert.equal(readWorkbook(simple).sheets[0].cells.A2.value, 'first');
simple = editWorkbook(simple, [{ sheet: 'Sheet1', operation: 'deleteRow', index: 1 }]);
assert.equal(readWorkbook(simple).sheets[0].cells.A2.value, 'second');
console.log(
  'PASS shared XLSX adapter: typed cells, formulas, cached-result invalidation, dates, styles, hidden sheets, merges, widths, unrelated OOXML/media, conflicts, structural edits, unsupported fallback'
);

const chained = { ...parts };
chained['xl/calcChain.xml'] = strToU8(
  '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="C1" i="1"/></calcChain>'
);
chained['xl/_rels/workbook.xml.rels'] = strToU8(
  strFromU8(chained['xl/_rels/workbook.xml.rels']).replace(
    '</Relationships>',
    '<Relationship Id="calc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>'
  )
);
chained['[Content_Types].xml'] = strToU8(
  strFromU8(chained['[Content_Types].xml']).replace(
    '</Types>',
    '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>'
  )
);
const recalculated = unzipSync(editWorkbook(zipSync(chained), [{ sheet: 'Sheet1', cell: 'C1', formula: 'B1+3' }]));
assert(!recalculated['xl/calcChain.xml']);
assert(!strFromU8(recalculated['xl/_rels/workbook.xml.rels']).includes('calcChain'));
assert(!strFromU8(recalculated['[Content_Types].xml']).includes('calcChain'));
assert(readWorkbook(zipSync(parts)).structuralReason.includes('Row/column'));
console.log('PASS calculation-chain invalidation and upfront structural-edit limitation');

const largeParts = unzipSync(createWorkbook());
largeParts['xl/media/large.bin'] = new Uint8Array(11 * 1024 * 1024);
const largeWorkbook = zipSync(largeParts, { level: 0 });
assert(largeWorkbook.length > 10 * 1024 * 1024);
assert.equal(readWorkbook(largeWorkbook).sheets.length, 1);
assert.equal(MAX_WORKBOOK_BYTES, Infinity);
console.log('PASS XLSX packages above 10 MiB are accepted without a fixed file-size ceiling');

const phonetic = unzipSync(createWorkbook());
phonetic['xl/sharedStrings.xml'] = strToU8(
  '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><t>Unit </t></r><r><t>cost</t></r><rPh sb="0" eb="4"><t>pronunciation-only</t></rPh></si></sst>'
);
phonetic['xl/worksheets/sheet1.xml'] = strToU8(
  strFromU8(phonetic['xl/worksheets/sheet1.xml']).replace(
    '<sheetData/>',
    '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>東京</t><rPh sb="0" eb="2"><t>とうきょう</t></rPh></is></c></row></sheetData>'
  )
);
const displayed = readWorkbook(zipSync(phonetic)).sheets[0].cells;
assert.equal(displayed.A1.value, 'Unit cost');
assert.equal(displayed.B1.value, '東京');
console.log('PASS shared and inline string display excludes phonetic annotations while preserving rich text');
