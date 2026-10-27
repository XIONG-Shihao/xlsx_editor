import test from 'node:test';
import assert from 'node:assert/strict';
import { encodeOwnerResult, decodeOwnerResult } from '../dist/core/transport.js';

test('object preview keeps signed zero through the real JSON transport codec', () => {
  const object = {
    id: '00000000-0000-4000-8000-000000000001',
    name: 'Chart',
    kind: 'chart',
    revision: 'a'.repeat(64),
    anchor: null,
    managed: true,
    protectedReason: null,
    canEdit: true,
    canDelete: true,
    chart: { kind: 'line', range: 'A1:B4' },
    data: {
      categories: ['negative zero', 'zero', 'large'],
      series: [{ name: 'Values', values: [-0, 0, 1e308] }],
      hidden: [false, false, false],
    },
  };
  const original = { object },
    wire = JSON.parse(JSON.stringify(encodeOwnerResult(original))),
    restored = decodeOwnerResult(wire);
  assert.equal(wire.object, undefined);
  assert.equal(typeof wire.encodedObject, 'string');
  assert(Object.is(restored.object.data.series[0].values[0], -0));
  assert(!Object.is(restored.object.data.series[0].values[1], -0));
  assert.equal(restored.object.data.series[0].values[2], 1e308);
  assert.deepEqual(restored.object, object);
  assert.equal(restored.encodedObject, undefined);
  assert.equal(original.object, object);
});

test('ordinary object list and preview preserve the existing unencoded path', () => {
  const list = { objects: { objects: [], total: 0, nextCursor: null } },
    preview = {
      object: {
        name: 'Plain',
        data: { categories: ['$workbookNumber'], series: [{ name: 'Values', values: [1, null] }], hidden: [false] },
      },
    };
  assert.equal(encodeOwnerResult(list), list);
  assert.equal(encodeOwnerResult(preview), preview);
  assert.deepEqual(decodeOwnerResult(JSON.parse(JSON.stringify(preview))), preview);
});
