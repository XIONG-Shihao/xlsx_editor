import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import { createWorkbook, readWorkbook, readWorkbookMetadata } from '../core/workbook-package.mjs';
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  PACKAGE = 'http://schemas.openxmlformats.org/package/2006/relationships';
const PART = 'xl/nonstandard/identity-sheet.xml',
  TARGET = 'nonstandard/identity-sheet.xml';
const esc = (s) => String(s).replaceAll('&', '&amp;').replaceAll('"', '&quot;');
function fixture({
  prefix = false,
  absolute = false,
  bookNamespace = MAIN,
  packageNamespace = PACKAGE,
  idNamespace = OFFICE,
  sheetXml,
  edges,
  relationshipBody,
  worksheetPrefix = false,
} = {}) {
  const parts = unzipSync(createWorkbook());
  delete parts['xl/worksheets/sheet1.xml'];
  const tag = (name) => (prefix ? 'm:' + name : name),
    rel = (name) => (prefix ? 'p:' + name : name);
  const sheet = sheetXml ?? `<${tag('sheet')} name="Data Ω" sheetId="1" ${prefix ? 'q' : 'r'}:id="sheet"/>`;
  parts['xl/workbook.xml'] = strToU8(
    `<${tag('workbook')} ${prefix ? 'xmlns:m' : 'xmlns'}="${bookNamespace}" xmlns:${prefix ? 'q' : 'r'}="${idNamespace}"><${tag('sheets')}>${sheet}</${tag('sheets')}><${tag('definedNames')}><${tag('definedName')} name="Input">'Data Ω'!$A$1</${tag('definedName')}></${tag('definedNames')}></${tag('workbook')}>`
  );
  const records = edges ?? [
    { Id: 'sheet', Type: OFFICE + '/worksheet', Target: absolute ? '/' + PART : TARGET },
    { Id: 'styles', Type: OFFICE + '/styles', Target: 'styles.xml' },
  ];
  const body =
    relationshipBody ??
    records
      .map(
        (record) =>
          `<${rel('Relationship')}${Object.entries(record)
            .map(([key, value]) => ` ${key}="${esc(value)}"`)
            .join('')}/>`
      )
      .join('');
  parts['xl/_rels/workbook.xml.rels'] = strToU8(
    `<${rel('Relationships')} ${prefix ? 'xmlns:p' : 'xmlns'}="${packageNamespace}">${body}</${rel('Relationships')}>`
  );
  const w = (name) => (worksheetPrefix ? 'w:' + name : name);
  parts[PART] = strToU8(
    `<${w('worksheet')} ${worksheetPrefix ? 'xmlns:w' : 'xmlns'}="${MAIN}"><${w('sheetData')}><${w('row')} r="1"><${w('c')} r="A1" t="inlineStr"><${w('is')}><${w('t')}>雪 😀</${w('t')}></${w('is')}></${w('c')}></${w('row')}></${w('sheetData')}></${w('worksheet')}>`
  );
  return Buffer.from(zipSync(parts, { level: 0, mtime: new Date('2020-01-01T00:00:00Z') }));
}
const readers = [
  ['cells', readWorkbook],
  ['metadata', (bytes) => readWorkbookMetadata(bytes).view],
];
for (const [name, read] of readers) {
  test(`${name}: exact namespaces admit default/prefixed aliases and nonstandard worksheet parts`, () => {
    for (const options of [
      {},
      { prefix: true },
      { prefix: true, absolute: true },
      { prefix: true, worksheetPrefix: true },
    ]) {
      const bytes = fixture(options),
        before = createHash('sha256').update(bytes).digest('hex'),
        view = read(bytes);
      assert.equal(view.sheets.length, 1);
      assert.equal(view.sheets[0].name, 'Data Ω');
      assert.equal(view.sheets[0].path, PART);
      assert.equal(view.sheets[0].cells.A1.value, '雪 😀');
      assert.deepEqual(view.namedRanges, [{ name: 'Input', formula: "'Data Ω'!$A$1" }]);
      assert.equal(createHash('sha256').update(bytes).digest('hex'), before);
    }
  });
  test(`${name}: foreign and strict namespaces cannot masquerade as supported identities`, () => {
    for (const options of [
      { prefix: true, bookNamespace: 'urn:foreign' },
      { prefix: true, packageNamespace: 'urn:foreign' },
      { prefix: true, idNamespace: 'urn:foreign' },
      { prefix: true, bookNamespace: 'http://purl.oclc.org/ooxml/spreadsheetml/main' },
      { prefix: true, sheetXml: '<other:sheet xmlns:other="urn:foreign" name="Data Ω" sheetId="1" q:id="sheet"/>' },
      {
        prefix: true,
        relationshipBody: `<other:Relationship xmlns:other="urn:foreign" Id="sheet" Type="${OFFICE}/worksheet" Target="${TARGET}"/>`,
      },
    ])
      assert.throws(() => read(fixture(options)), /namespace|foreign|missing|sheet inventory/i);
  });
  test(`${name}: missing, external and wrong-type worksheet edges fail closed`, () => {
    for (const edge of [
      { Id: 'other', Type: OFFICE + '/worksheet', Target: TARGET },
      { Id: 'sheet', Type: OFFICE + '/worksheet', Target: 'missing.xml' },
      { Id: 'sheet', Type: OFFICE + '/worksheet', Target: TARGET, TargetMode: 'External' },
      { Id: 'sheet', Type: OFFICE + '/worksheet', Target: TARGET, TargetMode: 'Unknown' },
      { Id: 'sheet', Type: OFFICE + '/styles', Target: TARGET },
      { Type: OFFICE + '/worksheet', Target: TARGET },
    ])
      assert.throws(() => read(fixture({ prefix: true, edges: [edge] })), /missing|external|malformed/i);
  });
  test(`${name}: duplicate edges, expanded-name aliases and worksheet bindings are rejected`, () => {
    const edge = { Id: 'sheet', Type: OFFICE + '/worksheet', Target: TARGET };
    assert.throws(() => read(fixture({ prefix: true, edges: [edge, edge] })), /duplicate/i);
    assert.throws(
      () =>
        read(
          fixture({
            prefix: true,
            sheetXml: `<m:sheet xmlns:r="${OFFICE}" name="Data Ω" sheetId="1" q:id="sheet" r:id="sheet"/>`,
          })
        ),
      /missing|external|duplicate/i
    );
    assert.throws(
      () =>
        read(
          fixture({
            prefix: true,
            sheetXml: '<m:sheet name="One" sheetId="1" q:id="sheet"/><m:sheet name="Two" sheetId="2" q:id="sheet"/>',
          })
        ),
      /duplicate/i
    );
    assert.throws(
      () =>
        read(
          fixture({
            prefix: true,
            relationshipBody: `<p:Relationship Id="sheet" Type="${OFFICE}/worksheet" Target="${TARGET}"><p:Relationship Id="nested" Type="${OFFICE}/styles" Target="styles.xml"/></p:Relationship>`,
          })
        ),
      /nested|malformed/i
    );
  });
}
