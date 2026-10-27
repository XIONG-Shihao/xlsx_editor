import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';
import { build } from 'esbuild';
import { zipSync, strToU8, unzipSync, strFromU8 } from 'fflate';
import { openZipParts, writeZipParts, sameZipMember, decodeXmlUtf8 } from '../core/zip-parts.mjs';
import { createWorkbook, editWorkbook, readWorkbook, indexedCalculationMetadata } from '../core/workbook-package.mjs';
import { relationshipPart, resolvePackageTarget } from '../core/package-paths.mjs';
test('portable package roundtrip runs with no Buffer, process or require and preserves media', async () => {
  const result = await build({
    stdin: {
      contents: `import {createWorkbook,editWorkbook,readWorkbook} from ${JSON.stringify(fileURLToPath(new URL('../core/workbook-package.mjs', import.meta.url)))};globalThis.run=()=>{const original=createWorkbook(),next=editWorkbook(original,[{sheet:'Sheet1',cell:'A1',value:'中文 😀'}]);return {revision:readWorkbook(next).revision,value:readWorkbook(next).sheets[0].cells.A1.value,bytes:next};};`,
      resolveDir: process.cwd(),
    },
    bundle: true,
    platform: 'browser',
    format: 'iife',
    write: false,
  });
  const context = vm.createContext({ TextEncoder, TextDecoder, setTimeout, clearTimeout });
  vm.runInContext(result.outputFiles[0].text, context);
  assert.equal(
    vm.runInContext('typeof Buffer+","+typeof process+","+typeof require', context),
    'undefined,undefined,undefined'
  );
  const output = context.run();
  assert.equal(output.value, '中文 😀');
  assert.equal(output.revision, createHash('sha256').update(output.bytes).digest('hex'));
  const bytes = zipSync({ 'xl/data.xml': strToU8('<data/>'), 'media/opaque.bin': new Uint8Array([0, 1, 2, 255]) }),
    parts = openZipParts(bytes);
  parts['xl/data.xml'] = strToU8('<updated/>');
  const saved = writeZipParts(parts),
    next = openZipParts(saved);
  assert(sameZipMember(parts, next, 'media/opaque.bin'));
  assert.deepEqual(unzipSync(saved)['media/opaque.bin'], new Uint8Array([0, 1, 2, 255]));
});
test('ZIP parsing honors byte offsets, archive comments and data descriptors', () => {
  const raw = zipSync({ 'word.xml': strToU8('unchanged') }, { level: 0, mtime: new Date(2000, 0, 1) }),
    wrapped = new Uint8Array(raw.length + 19);
  wrapped.set(raw, 7);
  assert.equal(strFromU8(openZipParts(wrapped.subarray(7, 7 + raw.length))['word.xml']), 'unchanged');
  const descriptor = Uint8Array.from(raw),
    data = new DataView(descriptor.buffer),
    end = descriptor.length - 22,
    central = data.getUint32(end + 16, true);
  data.setUint16(6, data.getUint16(6, true) | 8, true);
  data.setUint16(central + 8, data.getUint16(central + 8, true) | 8, true);
  for (const at of [14, 18, 22]) data.setUint32(at, 0, true);
  const comment = strToU8('archive comment'),
    commented = new Uint8Array(descriptor.length + comment.length);
  commented.set(descriptor);
  commented.set(comment, descriptor.length);
  new DataView(commented.buffer).setUint16(end + 20, comment.length, true);
  const output = writeZipParts(openZipParts(commented));
  assert.equal(strFromU8(output.subarray(output.length - comment.length)), 'archive comment');
  assert.equal(strFromU8(openZipParts(output)['word.xml']), 'unchanged');
});
test('bad CRCs, mismatched headers and invalid expansion bounds are rejected', () => {
  const raw = zipSync({ 'word.xml': strToU8('123456789') }, { level: 0 }),
    end = raw.length - 22,
    central = new DataView(raw.buffer).getUint32(end + 16, true);
  assert.equal(new DataView(raw.buffer).getUint32(central + 16, true), 0xcbf43926);
  const corrupt = Uint8Array.from(raw),
    start = 30 + new DataView(corrupt.buffer).getUint16(26, true) + new DataView(corrupt.buffer).getUint16(28, true);
  corrupt[start] ^= 1;
  assert.throws(() => openZipParts(corrupt)['word.xml'], { code: 'internal-mismatch' });
  const mismatch = Uint8Array.from(raw);
  new DataView(mismatch.buffer).setUint16(8, 8, true);
  assert.throws(() => openZipParts(mismatch), { code: /^(internal-mismatch|layout-limit)$/ });
  assert.throws(() => openZipParts(raw, NaN), { code: 'package-malformed' });
  assert.throws(() => openZipParts(raw, 8), /exceeds/);
});
test('OPC targets resolve within the package without OS path semantics', () => {
  assert.equal(relationshipPart('xl/worksheets/sheet1.xml'), 'xl/worksheets/_rels/sheet1.xml.rels');
  assert.equal(relationshipPart('book.xml'), '_rels/book.xml.rels');
  assert.equal(resolvePackageTarget('xl/worksheets/sheet1.xml', '../tables/table1.xml'), 'xl/tables/table1.xml');
  assert.equal(resolvePackageTarget('xl/worksheets/sheet1.xml', '/xl/tables/table1.xml'), 'xl/tables/table1.xml');
  assert.throws(() => resolvePackageTarget('xl/book.xml', '../../outside.xml'), /escapes/);
});
test('XML graph UTF-8 decoding is fatal for malformed, overlong and surrogate bytes', () => {
  assert.equal(decodeXmlUtf8(strToU8('\uFEFF中文 😀')), '中文 😀');
  for (const invalid of [[0xc0, 0xaf], [0xed, 0xa0, 0x80], [0xf4, 0x90, 0x80, 0x80], [0xf0, 0x9f], [0x80]])
    assert.throws(() => decodeXmlUtf8(new Uint8Array(invalid)), { code: 'internal-invalid-field' });
});
test('complete default calculation policy retains exact workbook XML; manual and refresh settings remain explicit', () => {
  const bytes = createWorkbook(),
    source = strFromU8(unzipSync(bytes)['xl/workbook.xml']);
  const defaults = indexedCalculationMetadata(bytes, false, { mode: 'auto', recalculateBeforeSave: true });
  assert.equal(defaults.replacements['xl/workbook.xml'], source);
  assert.match(
    indexedCalculationMetadata(bytes, true, { mode: 'auto', recalculateBeforeSave: true }).replacements[
      'xl/workbook.xml'
    ],
    /fullCalcOnLoad="1"/
  );
  assert.match(
    indexedCalculationMetadata(bytes, false, { mode: 'manual', recalculateBeforeSave: false }).replacements[
      'xl/workbook.xml'
    ],
    /calcMode="manual"/
  );
  assert.equal(
    readWorkbook(editWorkbook(bytes, [{ sheet: 'Sheet1', cell: 'A1', value: 4 }])).sheets[0].cells.A1.value,
    4
  );
});
