import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { buildSpillMetadata } from '../core/spill-metadata.mjs';
import { createWorkbook } from '../core/workbook-package.mjs';

const checker = fileURLToPath(new URL('./support/check-spill-metadata.py', import.meta.url));
const sheet = 'xl/worksheets/sheet1.xml';
const group = {
  schema: 1,
  sheet,
  master: 'A1',
  range: 'A1',
  intendedRange: 'A1:B2',
  status: 'error',
  error: '#SPILL!',
};
async function fixture(t, source = createWorkbook(), groups = [group]) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-spill-closed-oracle-'));
  t.after(() => fs.rm(root, { recursive: true, force: true }));
  const input = path.join(root, 'source.xlsx'),
    output = path.join(root, 'saved.xlsx'),
    mapping = path.join(root, 'mapping.json');
  const result = buildSpillMetadata(source, groups),
    original = {
      ...unzipSync(source),
      ...Object.fromEntries(Object.entries(result.replacements).map(([name, text]) => [name, strToU8(text)])),
    };
  await fs.writeFile(input, source);
  async function probe(mutate = () => {}, mutateMap = () => {}) {
    const parts = { ...original },
      masters = structuredClone(result.masters);
    mutate(parts);
    mutateMap(masters);
    await fs.writeFile(output, zipSync(parts));
    await fs.writeFile(mapping, JSON.stringify(masters));
    return spawnSync('python3', [checker, input, output, mapping], { encoding: 'utf8' });
  }
  const positive = await probe();
  assert.equal(positive.status, 0, positive.stderr);
  return { probe, result, original };
}
function replace(part, from, to) {
  return (parts) => {
    const text = strFromU8(parts[part]);
    assert(text.includes(from), from);
    parts[part] = strToU8(text.replace(from, to));
  };
}
function appendCounted(part, tag, child) {
  return (parts) => {
    const text = strFromU8(parts[part]),
      expression = new RegExp(`(<${tag}\\b[^>]*count=")(\\d+)(")`);
    assert(expression.test(text), tag);
    parts[part] = strToU8(
      text
        .replace(expression, (_, prefix, count, suffix) => prefix + (Number(count) + 1) + suffix)
        .replace(`</${tag}>`, child + `</${tag}>`)
    );
  };
}
function replacePattern(part, expression, replacement) {
  return (parts) => {
    const text = strFromU8(parts[part]);
    assert(expression.test(text), expression.toString());
    parts[part] = strToU8(text.replace(expression, replacement));
  };
}

test('independent spill checker rejects wrong extension ancestry, URI, sibling, attributes and text', async (t) => {
  const { probe } = await fixture(t),
    part = 'xl/metadata.xml';
  const mutations = [
    ['dynamic URI', replace(part, '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}', '{wrong-dynamic}')],
    ['rich URI', replace(part, '{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}', '{wrong-rich}')],
    ['dynamic ancestor', replace(part, '<xda:dynamicArrayProperties', '<foreign/><xda:dynamicArrayProperties')],
    ['rich ancestor', replacePattern(part, /(<xlrd:rvb\b[^>]*\/>)/, '<unknown>$1</unknown>')],
    ['dynamic leaf child', replace(part, 'fCollapsed="0"/>', 'fCollapsed="0"><extra/></xda:dynamicArrayProperties>')],
    ['rich leaf attribute', replacePattern(part, /<xlrd:rvb\b/, '<xlrd:rvb extra="1"')],
    ['extension text', replace(part, '<xda:dynamicArrayProperties', 'unrequested<xda:dynamicArrayProperties')],
    ['nested rich-type flag', replace('xl/richData/rdRichValueTypes.xml', 'value="1"/>', 'value="1"><extra/></flag>')],
    ['rich-type container text', replace('xl/richData/rdRichValueTypes.xml', '<global>', '<global>unrequested')],
  ];
  for (const [label, mutate] of mutations) {
    const reply = await probe(mutate);
    assert.notEqual(reply.status, 0, label);
  }
});

test('independent spill checker admits only new records and package closure reachable from requested mappings', async (t) => {
  const { probe } = await fixture(t),
    part = 'xl/metadata.xml';
  const mutations = [
    ['unused type', appendCounted(part, 'metadataTypes', '<metadataType name="OTHER"/>')],
    ['unused future', replace(part, '</metadata>', '<futureMetadata name="OTHER" count="0"/></metadata>')],
    ['unused root extension', replace(part, '</metadata>', '<extLst/></metadata>')],
    ['unused cell record', appendCounted(part, 'cellMetadata', '<bk><rc t="1" v="0"/></bk>')],
    ['unused value record', appendCounted(part, 'valueMetadata', '<bk><rc t="2" v="0"/></bk>')],
    [
      'unused rich value',
      appendCounted('xl/richData/rdrichvalue.xml', 'rvData', '<rv s="0"><v>1</v><v>8</v><v>1</v><v>1</v></rv>'),
    ],
    [
      'unused rich structure',
      appendCounted(
        'xl/richData/rdrichvaluestructure.xml',
        'rvStructures',
        '<s t="_error"><k n="colOffset" t="i"/><k n="errorType" t="i"/><k n="rwOffset" t="i"/><k n="subType" t="i"/></s>'
      ),
    ],
  ];
  for (const [label, mutate] of mutations) {
    const reply = await probe(mutate);
    assert.notEqual(reply.status, 0, label);
  }
  const unrequestedRich = await probe(
    () => {},
    (masters) => {
      delete masters[`${sheet}\0A1`].vm;
    }
  );
  assert.notEqual(unrequestedRich.status, 0, 'rich package closure without any requested vm');
});

test('independent spill geometry comes from explicit intended range rather than producer offsets', async (t) => {
  const { probe } = await fixture(t),
    key = `${sheet}\0A1`;
  for (const intendedRange of ['A1:C2', 'B1:C2', 'A1:A1001']) {
    const reply = await probe(
      () => {},
      (masters) => {
        masters[key].intendedRange = intendedRange;
      }
    );
    assert.notEqual(reply.status, 0, intendedRange);
  }
  const wrongPublished = await probe(
    () => {},
    (masters) => {
      masters[key].range = 'A1:B2';
    }
  );
  assert.notEqual(wrongPublished.status, 0);
});

test('independent spill checker permits exact reuse and required appends while preserving unrelated original records', async (t) => {
  const source = await fs.readFile(new URL('../fixtures/excel-dynamic-reference.xlsx', import.meta.url));
  const reused = await fixture(t, source, [{ ...group, master: 'A41', range: 'A41', intendedRange: 'A41:B42' }]);
  assert.deepEqual(reused.result.replacements, {});
  const appended = await fixture(t, source, [{ ...group, master: 'A90', range: 'A90', intendedRange: 'A90:D94' }]);
  assert(!Object.hasOwn(appended.result.replacements, 'xl/richData/rdrichvaluestructure.xml'));
  assert(!Object.hasOwn(appended.result.replacements, 'xl/richData/rdRichValueTypes.xml'));
  const parts = unzipSync(source),
    metadata = strFromU8(parts['xl/metadata.xml']);
  parts['xl/metadata.xml'] = strToU8(
    metadata.replace(
      '</metadata>',
      '<extLst xmlns:q="urn:retained"><q:opaque> text <q:item/> tail </q:opaque></extLst></metadata>'
    )
  );
  await fixture(t, zipSync(parts), [{ ...group, master: 'A90', range: 'A90', intendedRange: 'A90:D94' }]);
});
