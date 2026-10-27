import test from 'node:test';
import assert from 'node:assert/strict';
import { decompressVba } from '../core/vba-source.mjs';
const bytes = (values) => Uint8Array.from(values),
  text = (output) => new TextDecoder().decode(output);
test('MS-OVBA literal and overlapping-copy vectors preserve source bytes', () => {
  // One compressed literal sequence and an offset-one overlapping reference.
  assert.equal(text(decompressVba(bytes([1, 3, 0xb0, 0, 65, 66, 67]))), 'ABC');
  assert.equal(text(decompressVba(bytes([1, 3, 0xb0, 2, 65, 2, 0]))), 'AAAAAA');
  // Offset three copies ABC twice, including output written by this token.
  assert.equal(text(decompressVba(bytes([1, 5, 0xb0, 8, 65, 66, 67, 3, 0x20]))), 'ABCABCABC');
});
test('raw chunks reset the next compressed dictionary and respect a declared budget', () => {
  const raw = new Uint8Array(4099);
  raw.set([1, 0xff, 0x3f]);
  raw.fill(90, 3);
  assert.equal(decompressVba(raw).length, 4096);
  const two = new Uint8Array(raw.length + 6);
  two.set(raw);
  two.set([3, 0xb0, 0, 65, 66, 67], raw.length);
  const decoded = decompressVba(two);
  assert.equal(decoded.length, 4099);
  assert.equal(text(decoded.subarray(4096)), 'ABC');
  assert.throws(() => decompressVba(two, { maxBytes: 4096 }), /budget/);
  const cross = new Uint8Array(raw.length + 5);
  cross.set(raw);
  cross.set([2, 0xb0, 1, 0, 0], raw.length);
  assert.throws(() => decompressVba(cross), /copy token/);
});
test('malformed containers cannot yield truncated or overexpanded module code', () => {
  for (const input of [
    [0, 3, 0xb0, 0, 65, 66, 67],
    [1, 3, 0x80, 0, 65, 66, 67],
    [1, 3, 0xb0, 0, 65],
    [1, 0, 0xb0, 0],
    [1, 3, 0xb0, 1, 0, 0, 0],
    [1, 0, 0x30, 0],
  ])
    assert.throws(() => decompressVba(bytes(input)), { code: /^(session-stale|workbook-refused)$/ });
  const nonfinal = bytes([1, 3, 0xb0, 0, 65, 66, 67, 3, 0xb0, 0, 65, 66, 67]);
  assert.throws(() => decompressVba(nonfinal), /nonfinal/);
  assert.throws(() => decompressVba(bytes([1, 3, 0xb0, 0, 65, 66, 67]), { maxBytes: 0 }), /budget/);
});
