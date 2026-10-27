import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { buildSpillMetadata } from '../core/spill-metadata.mjs';
import { createWorkbook } from '../core/workbook-package.mjs';
const base = { schema: 1, sheet: 'xl/worksheets/sheet1.xml' };
const positive = (master, range) => ({ ...base, master, range, status: 'calculated' });
const blocked = (master, intendedRange) => ({
  ...base,
  master,
  range: master,
  intendedRange,
  status: 'error',
  error: '#SPILL!',
});
async function checked(t, source, result) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-spill-metadata-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const before = path.join(root, 'source.xlsx'),
    after = path.join(root, 'saved.xlsx'),
    mapping = path.join(root, 'mapping.json'),
    parts = unzipSync(source);
  for (const [part, xml] of Object.entries(result.replacements)) parts[part] = strToU8(xml);
  await fs.writeFile(before, source);
  await fs.writeFile(after, zipSync(parts));
  await fs.writeFile(mapping, JSON.stringify(result.masters));
  const probe = spawnSync(
    'python3',
    [fileURLToPath(new URL('./support/check-spill-metadata.py', import.meta.url)), before, after, mapping],
    { encoding: 'utf8' }
  );
  assert.equal(probe.status, 0, probe.stderr);
  return JSON.parse(probe.stdout);
}
test('retained actual metadata reuses exact XLDAPR and rich SPILL records without rewriting any part', async (t) => {
  const source = await fs.readFile(new URL('../fixtures/excel-dynamic-reference.xlsx', import.meta.url)),
    result = buildSpillMetadata(source, [
      positive('A1', 'A1:B3'),
      positive('A6', 'A6:B7'),
      positive('A56', 'A56:B57'),
      blocked('A41', 'A41:B42'),
      blocked('A60', 'A60:B61'),
      blocked('A69', 'A69:B70'),
    ]);
  assert.deepEqual(result.replacements, {});
  for (const ref of ['A41', 'A60', 'A69']) assert.equal(result.masters[`${base.sheet}\0${ref}`].vm, 2);
  const proof = await checked(t, source, result);
  assert.deepEqual(proof.masters[`${base.sheet}\0A41`], {
    dynamic: true,
    colOffset: 1,
    errorType: 8,
    rwOffset: 1,
    subType: 1,
  });
});
test('new dynamic and blocked metadata preserve exact existing records and have independently resolved indexes', async (t) => {
  for (const source of [
    createWorkbook(),
    await fs.readFile(new URL('../fixtures/excel-dynamic-source.xlsx', import.meta.url)),
  ]) {
    const result = buildSpillMetadata(source, [positive('A1', 'A1:B3'), blocked('A8', 'A8:C11')]),
      proof = await checked(t, source, result);
    assert.deepEqual(proof.masters[`${base.sheet}\0A8`], {
      dynamic: true,
      colOffset: 2,
      errorType: 8,
      rwOffset: 3,
      subType: 1,
    });
  }
});
test('appending a new rich SPILL shape preserves original unknown modern error and opaque package records', async (t) => {
  const source = await fs.readFile(new URL('../fixtures/excel-dynamic-reference.xlsx', import.meta.url)),
    result = buildSpillMetadata(source, [blocked('A90', 'A90:D94')]),
    proof = await checked(t, source, result);
  assert.equal(result.masters[`${base.sheet}\0A90`].vm, 4);
  assert.deepEqual(proof.masters[`${base.sheet}\0A90`], {
    dynamic: true,
    colOffset: 3,
    errorType: 8,
    rwOffset: 4,
    subType: 1,
  });
  assert(!Object.hasOwn(result.replacements, 'xl/richData/rdrichvaluestructure.xml'));
  assert(!Object.hasOwn(result.replacements, 'xl/richData/rdRichValueTypes.xml'));
});
test('malformed dynamic metadata, unknown modern errors and overlapping or excessive output shapes refuse', () => {
  const valid = createWorkbook();
  for (const groups of [
    [{ ...blocked('A1', 'A1:B2'), error: '#VALUE!' }],
    [positive('A1', 'A1:A1001')],
    [positive('A1', 'B1:B2')],
    [positive('A1', 'A1:B2'), positive('B1', 'B1:B3')],
    [{ ...positive('A1', 'A1'), schema: 2 }],
  ])
    assert.throws(() => buildSpillMetadata(valid, groups));
  const parts = unzipSync(valid);
  parts['xl/metadata.xml'] = strToU8(
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><metadataTypes count="2"/></metadata>'
  );
  assert.throws(() => buildSpillMetadata(zipSync(parts), [positive('A1', 'A1')]));
  assert.deepEqual(buildSpillMetadata(valid, []), { replacements: {}, masters: {} });
});

test('independent checker rejects exact graph, index, namespace, count and modern-error mutants', async (t) => {
  const source = createWorkbook(),
    result = buildSpillMetadata(source, [blocked('A1', 'A1:B2')]),
    root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-spill-mutants-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const input = path.join(root, 'source.xlsx'),
    output = path.join(root, 'saved.xlsx'),
    mapping = path.join(root, 'mapping.json');
  await fs.writeFile(input, source);
  const original = {
      ...unzipSync(source),
      ...Object.fromEntries(Object.entries(result.replacements).map(([part, xml]) => [part, strToU8(xml)])),
    },
    key = `${base.sheet}\0A1`;
  const replace = (part, from, to) => (parts) => {
    const text = strFromU8(parts[part]);
    assert(text.includes(from), from);
    parts[part] = strToU8(text.replace(from, to));
  };
  const mutations = [
    [
      'relationship type',
      replace('xl/_rels/workbook.xml.rels', 'relationships/sheetMetadata', 'relationships/worksheet'),
    ],
    ['content type', replace('[Content_Types].xml', 'application/vnd.ms-excel.rdrichvalue+xml', 'application/x-wrong')],
    [
      'duplicate relationship ID',
      (parts) => {
        const text = strFromU8(parts['xl/_rels/workbook.xml.rels']);
        parts['xl/_rels/workbook.xml.rels'] = strToU8(text.replace('Id="workbookSpill1"', 'Id="rId1"'));
      },
    ],
    [
      'cm zero',
      null,
      (masters) => {
        masters[key].cm = 0;
      },
    ],
    [
      'vm zero',
      null,
      (masters) => {
        masters[key].vm = 0;
      },
    ],
    ['rc type zero', replace('xl/metadata.xml', '<rc t="1" v="0"', '<rc t="0" v="0"')],
    ['future negative index', replace('xl/metadata.xml', '<rc t="1" v="0"', '<rc t="1" v="-1"')],
    ['rich negative index', replace('xl/metadata.xml', 'i="0"', 'i="-1"')],
    ['structure negative index', replace('xl/richData/rdrichvalue.xml', 's="0"', 's="-1"')],
    ['rc foreign namespace', replace('xl/metadata.xml', '<rc t="1"', '<rc xmlns="urn:foreign" t="1"')],
    ['rich key type', replace('xl/richData/rdrichvaluestructure.xml', 'n="errorType" t="i"', 'n="errorType" t="s"')],
    ['collapsed flag', replace('xl/metadata.xml', 'fCollapsed="0"', 'fCollapsed="1"')],
    ['new rich count', replace('xl/richData/rdrichvalue.xml', 'count="1"', 'count="2"')],
    ['rich global flags', replace('xl/richData/rdRichValueTypes.xml', 'value="1"', 'value="0"')],
    [
      'unknown new rich child',
      replace('xl/richData/rdRichValueTypes.xml', '</rvTypesInfo>', '<unknown/></rvTypesInfo>'),
    ],
  ];
  for (const [name, mutate, mutateMap] of mutations) {
    const parts = { ...original },
      masters = structuredClone(result.masters);
    mutate?.(parts);
    mutateMap?.(masters);
    await fs.writeFile(output, zipSync(parts));
    await fs.writeFile(mapping, JSON.stringify(masters));
    const probe = spawnSync(
      'python3',
      [fileURLToPath(new URL('./support/check-spill-metadata.py', import.meta.url)), input, output, mapping],
      { encoding: 'utf8' }
    );
    assert.notEqual(probe.status, 0, name);
  }
});

test('opaque metadata text/tails and namespace bindings cannot change during an append', async (t) => {
  const inputParts = unzipSync(await fs.readFile(new URL('../fixtures/excel-dynamic-reference.xlsx', import.meta.url))),
    part = 'xl/metadata.xml';
  inputParts[part] = strToU8(
    strFromU8(inputParts[part]).replace(
      '</metadata>',
      '<extLst xmlns:q="urn:original"><q:opaque> leading <q:inner/> trailing </q:opaque></extLst></metadata>'
    )
  );
  const source = zipSync(inputParts),
    result = buildSpillMetadata(source, [blocked('A90', 'A90:D94')]);
  await checked(t, source, result);
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-spill-opaque-mutants-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const input = path.join(root, 'source.xlsx'),
    output = path.join(root, 'saved.xlsx'),
    mapping = path.join(root, 'mapping.json');
  await fs.writeFile(input, source);
  await fs.writeFile(mapping, JSON.stringify(result.masters));
  const text = result.replacements[part];
  for (const changed of [
    text.replace(' leading ', 'leading'),
    text.replace(' trailing ', 'trailing'),
    text.replace('urn:original', 'urn:changed'),
  ]) {
    assert.notEqual(changed, text);
    const parts = {
      ...inputParts,
      ...Object.fromEntries(Object.entries(result.replacements).map(([p, xml]) => [p, strToU8(xml)])),
      [part]: strToU8(changed),
    };
    await fs.writeFile(output, zipSync(parts));
    const probe = spawnSync(
      'python3',
      [fileURLToPath(new URL('./support/check-spill-metadata.py', import.meta.url)), input, output, mapping],
      { encoding: 'utf8' }
    );
    assert.notEqual(probe.status, 0);
  }
});

test('writer refuses invalid value-metadata classification and duplicate or missing relationship identities', async () => {
  const source = await fs.readFile(new URL('../fixtures/excel-dynamic-reference.xlsx', import.meta.url)),
    parts = unzipSync(source),
    meta = strFromU8(parts['xl/metadata.xml']);
  const wrong = {
    ...parts,
    'xl/metadata.xml': strToU8(meta.replace('name="XLRICHVALUE"', 'name="XLRICHVALUE" cellMeta="1"')),
  };
  assert.throws(() => buildSpillMetadata(zipSync(wrong), [blocked('A90', 'A90:B91')]), /value metadata/);
  const rels = strFromU8(parts['xl/_rels/workbook.xml.rels']);
  for (const invalid of [rels.replace('Id="rId8"', 'Id="rId3"'), rels.replace('Id="rId8"', 'Id=""')])
    assert.throws(
      () =>
        buildSpillMetadata(zipSync({ ...parts, 'xl/_rels/workbook.xml.rels': strToU8(invalid) }), [
          positive('A1', 'A1:B3'),
        ]),
      { code: /^(spill-ambiguous|workbook-bad-input)$/ }
    );
});

test('existing metadata parts cannot be silently activated through newly invented package links', async () => {
  const source = await fs.readFile(new URL('../fixtures/excel-dynamic-reference.xlsx', import.meta.url)),
    parts = unzipSync(source);
  const relPart = 'xl/_rels/workbook.xml.rels',
    typePart = '[Content_Types].xml',
    rels = strFromU8(parts[relPart]),
    types = strFromU8(parts[typePart]);
  for (const [part, text, expression] of [
    [relPart, rels, /<Relationship\b[^>]*Type="[^"]*\/sheetMetadata"[^>]*\/>/],
    [typePart, types, /<Override\b[^>]*PartName="\/xl\/metadata.xml"[^>]*\/>/],
    [relPart, rels, /<Relationship\b[^>]*Type="[^"]*\/rdRichValue"[^>]*\/>/],
  ]) {
    const changed = text.replace(expression, '');
    assert.notEqual(changed, text);
    assert.throws(
      () => buildSpillMetadata(zipSync({ ...parts, [part]: strToU8(changed) }), [blocked('A90', 'A90:B91')]),
      /qualified relationship|qualified content type/
    );
  }
  const duplicate = types.replace(
    '</Types>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/x-other"/></Types>'
  );
  assert.throws(
    () => buildSpillMetadata(zipSync({ ...parts, [typePart]: strToU8(duplicate) }), [positive('A1', 'A1:B3')]),
    { code: /^(spill-ambiguous|workbook-bad-input)$/ }
  );
});

test('writer refuses mapped dynamic/rich leaves under a wrong extension or ancestry', async () => {
  const source = await fs.readFile(new URL('../fixtures/excel-dynamic-reference.xlsx', import.meta.url)),
    parts = unzipSync(source),
    part = 'xl/metadata.xml',
    text = strFromU8(parts[part]);
  for (const [from, to] of [
    ['{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}', '{wrong-dynamic-extension}'],
    ['{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}', '{wrong-rich-extension}'],
    [
      '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>',
      '<unknown><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></unknown>',
    ],
    ['<xlrd:rvb i="0"/>', '<unknown><xlrd:rvb i="0"/></unknown>'],
  ]) {
    assert(text.includes(from), from);
    assert.throws(
      () =>
        buildSpillMetadata(zipSync({ ...parts, [part]: strToU8(text.replace(from, to)) }), [blocked('A90', 'A90:B91')]),
      /qualification|extension URI/
    );
  }
});
