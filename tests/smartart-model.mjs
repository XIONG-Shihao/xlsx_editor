import { test } from 'node:test';
import assert from 'node:assert/strict';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { SMARTART_IDS, smartartFixture } from './support/smartart-fixture.mjs';
import {
  smartartSampleModel,
  normalizeSmartArt,
  smartartGeometry,
  smartartPartNames,
} from '../core/smartart-model.mjs';
import { indexedObjectList, indexedObjectPreview } from '../core/indexed-object-authoring.mjs';
const view = { readOnly: false, sheets: [{ name: 'Sheet1', path: 'xl/worksheets/sheet1.xml' }] },
  list = (bytes) => indexedObjectList(bytes, view, 'Sheet1');
test('three actual linked DiagramML graphs expose real nodes/connections/layouts as editable readable samples', async () => {
  const bytes = smartartFixture(),
    entries = list(bytes).objects;
  assert.equal(entries.length, 3);
  for (const [i, id] of SMARTART_IDS.entries()) {
    const e = entries.find((e) => e.id === id),
      model = smartartSampleModel(['process', 'hierarchy', 'cycle'][i]);
    assert.deepEqual(e.smartart, model);
    assert.equal(e.kind, 'smartart');
    assert.equal(e.canEdit, true);
    assert.equal(e.canDelete, true);
    assert.equal(e.protectedReason, null);
    const p = await indexedObjectPreview({ metadataBytes: bytes, view }, 'Sheet1', id);
    assert.deepEqual(p.smartart, model);
    assert.equal(p.altText, 'Workbook ' + model.layout + ' sample');
  }
  assert(!Object.keys(unzipSync(bytes)).some((p) => p.startsWith('xl/media/')));
});
test('all supported two-to-six-node topologies are detached, explicit and finite', () => {
  for (const layout of ['process', 'hierarchy', 'cycle'])
    for (let n = 2; n <= 6; n++) {
      const source = {
          layout,
          nodes: Array.from({ length: n }, (_, i) => ({ id: 'n' + (i + 1), text: 'Node ' + (i + 1) })),
        },
        model = normalizeSmartArt(source),
        g = smartartGeometry(model);
      source.nodes[0].text = 'Caller mutation';
      assert.equal(model.nodes[0].text, 'Node 1');
      assert.equal(model.connections.length, n - (layout === 'cycle' ? 0 : 1));
      for (const b of g.boxes) assert([b.x, b.y, b.w, b.h].every(Number.isFinite));
      for (const e of g.edges) assert([e.x1, e.y1, e.x2, e.y2].every(Number.isFinite));
      const reverse = { ...model, connections: model.connections.map((c) => ({ to: c.to, from: c.from })) };
      assert.deepEqual(normalizeSmartArt(reverse), model);
    }
});
test('blank/oversized/unknown/dangling models refuse instead of producing an invented diagram', () => {
  for (const value of [
    { layout: 'unknown', nodes: [] },
    { layout: 'process', nodes: [{ id: 'n1', text: 'One' }] },
    {
      layout: 'cycle',
      nodes: [
        { id: 'n1', text: 'One' },
        { id: 'n2', text: ' ' },
      ],
    },
    {
      layout: 'process',
      nodes: [
        { id: 'n1', text: 'One' },
        { id: 'n2', text: 'x'.repeat(129) },
      ],
    },
    {
      layout: 'hierarchy',
      nodes: [
        { id: 'n1', text: 'One' },
        { id: 'n2', text: 'Two' },
      ],
      connections: [{ from: 'n2', to: 'missing' }],
    },
  ])
    assert.throws(() => normalizeSmartArt(value));
});
test('mutated data/layout/cache/colors/style/type/ownership become protected with no semantic preview', () => {
  const bytes = smartartFixture(),
    parts = unzipSync(bytes),
    names = smartartPartNames(SMARTART_IDS[0]),
    drawing = Object.keys(parts).find((p) => p.startsWith('xl/drawings/') && p.endsWith('.xml')),
    rel = drawing.replace(/([^/]+)$/, '_rels/$1.rels');
  const mutations = [
    (p) => (p[names.data] = strToU8(strFromU8(p[names.data]).replace('Plan', 'Changed'))),
    (p) => (p[names.layout] = strToU8(strFromU8(p[names.layout]).replace('type="lin"', 'type="cycle"'))),
    (p) => (p[names.drawing] = strToU8(strFromU8(p[names.drawing]).replace('sz="1400"', 'sz="1600"'))),
    (p) => (p[names.colors] = strToU8(strFromU8(p[names.colors]).replace('4472C4', 'FF0000'))),
    (p) => (p[names.style] = strToU8(strFromU8(p[names.style]).replace('idx="1"', 'idx="9"'))),
    (p) =>
      (p['[Content_Types].xml'] = strToU8(
        strFromU8(p['[Content_Types].xml']).replace('diagramColors+xml', 'wrong+xml')
      )),
    (p) =>
      (p[rel] = strToU8(
        strFromU8(p[rel]).replace(
          '</Relationships>',
          '<Relationship Id="shared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/' +
            names.data.split('/').at(-1) +
            '"/></Relationships>'
        )
      )),
  ];
  mutations.push(
    (p) =>
      (p[rel] = strToU8(
        strFromU8(p[rel]).replace(
          '</Relationships>',
          '<Relationship Id="absoluteShared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="/' +
            names.data +
            '" TargetMode="Internal"/></Relationships>'
        )
      ))
  );
  mutations.push(
    (p) =>
      (p[names.colors.replace(/([^/]+)$/, '_rels/$1.rels')] = strToU8(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="foreign" Type="unknown" Target="https://example.com/" TargetMode="External"/></Relationships>'
      ))
  );
  for (const mutate of mutations) {
    const p = structuredClone(parts);
    mutate(p);
    const e = list(zipSync(p)).objects.find((e) => e.id === SMARTART_IDS[0]);
    assert.equal(e.smartart, undefined);
    assert(e.protectedReason);
    assert.equal(e.canEdit, false);
  }
});
