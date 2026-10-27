import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { encodeCsvText } from '@office-editors/xlsx-editor/core/csv-codec.mjs';
import { readCsvWorkbook, editCsvBytes, editCsvWorkbook } from '@office-editors/xlsx-editor/core/csv-document.mjs';
const hash = (b) => createHash('sha256').update(b).digest('hex');
let count = 0;
for (const [encoding, bom] of [
  ['utf-8', false],
  ['utf-8', true],
  ['utf-16le', true],
  ['utf-16be', true],
])
  for (const terminal of ['', '\r\n']) {
    const text = 'id;value\r\n1;"🙂中文\r\nquote""x"\n2;keep\r3;tail' + terminal;
    const bytes = encodeCsvText(text, encoding, bom),
      original = bytes.slice(),
      revision = hash(bytes);
    const commands = [
      { sheet: 'CSV', cell: 'B2', value: 'new;🙂' },
      { sheet: 'CSV', cell: 'B4', value: '終' },
      { sheet: 'CSV', cell: 'C6', value: 'last' },
    ];
    const expected =
      'id;value\r\n1;"new;🙂"\n2;keep\r3;終' + terminal + (terminal ? '' : '\r\n') + '\r\n;;last' + terminal;
    const actual = editCsvBytes(bytes, commands, revision, { delimiter: ';' });
    assert.deepEqual(actual.bytes, encodeCsvText(expected, encoding, bom));
    assert.equal(actual.revision, hash(actual.bytes));
    assert.deepEqual(bytes, original);
    assert.deepEqual(actual.bytes, editCsvWorkbook(bytes, commands, revision, { delimiter: ';' }).bytes);
    const noChange = editCsvBytes(bytes, [{ sheet: 'CSV', cell: 'B3', value: 'keep' }], revision, { delimiter: ';' });
    assert.deepEqual(noChange.bytes, bytes);
    assert.equal(noChange.revision, revision);
    assert.throws(() => editCsvBytes(bytes, commands, 'stale', { delimiter: ';' }), /stale/);
    assert.throws(
      () =>
        editCsvBytes(bytes, [...commands, { sheet: 'CSV', cell: 'B3', value: '\ud800' }], revision, { delimiter: ';' }),
      { code: 'csv-malformed' }
    );
    assert.deepEqual(bytes, original);
    count++;
  }
const empty = new Uint8Array();
assert.equal(
  new TextDecoder().decode(editCsvBytes(empty, [{ sheet: 'CSV', cell: 'B2', value: 'x' }], hash(empty)).bytes),
  '\n,x'
);
console.log(
  `PASS ${count} sparse byte-save cases: UTF-8/BOM/UTF-16LE/BE, mixed newlines, multiline quotes, non-BMP characters, row/column extension, exact expected bytes, no-op/stale/invalid-input atomicity, empty source`
);
