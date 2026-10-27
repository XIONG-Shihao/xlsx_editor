import test from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { planWorksheetLayout } from '../core/layout-planning.mjs';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const parse = (xml) => new DOMParser().parseFromString(xml, 'application/xml');
const elements = (xml, name) => Array.from(parse(xml).getElementsByTagNameNS(NS, name));
test('column splitting keeps source properties and repeated widths use the final value', () => {
  const source = `<worksheet xmlns="${NS}"><cols><col min="1" max="4" width="12" hidden="1" style="3" outlineLevel="2" bestFit="1"/></cols><sheetData/></worksheet>`;
  const { fragments } = planWorksheetLayout(source, [
    { operation: 'resizeColumn', index: 1, width: 22 },
    { operation: 'resizeColumn', index: 1, width: 27 },
    { operation: 'resizeColumn', index: 5, width: 30 },
  ]);
  assert.deepEqual(
    fragments.map((f) => f.tag),
    ['cols']
  );
  const cols = elements(fragments[0].xml, 'col');
  assert.deepEqual(
    cols.map((c) => [c.getAttribute('min'), c.getAttribute('max'), c.getAttribute('width')]),
    [
      ['1', '1', '12'],
      ['2', '2', '27'],
      ['3', '4', '12'],
      ['6', '6', '30'],
    ]
  );
  for (const col of cols.slice(0, 3)) {
    for (const [key, value] of Object.entries({ hidden: '1', style: '3', outlineLevel: '2', bestFit: '1' }))
      assert.equal(col.getAttribute(key), value);
  }
  assert.equal(cols[3].hasAttribute('hidden'), false);
});
test('freeze and unfreeze preserve other view settings and additional views', () => {
  const source = `<worksheet xmlns="${NS}"><sheetViews><sheetView workbookViewId="0" zoomScale="140"><pane state="split" xSplit="300"/><selection pane="bottomRight" activeCell="C5" sqref="C5"/></sheetView><sheetView workbookViewId="1" zoomScale="90"><pane state="frozen" ySplit="3"/></sheetView></sheetViews><sheetData/></worksheet>`;
  const freeze = planWorksheetLayout(source, [{ operation: 'freeze', rows: 2, columns: 27 }]).fragments[0].xml;
  const views = elements(freeze, 'sheetView');
  assert.equal(views[0].getAttribute('zoomScale'), '140');
  const pane = views[0].getElementsByTagNameNS(NS, 'pane')[0];
  assert.equal(pane.getAttribute('topLeftCell'), 'AB3');
  assert.equal(pane.getAttribute('activePane'), 'bottomRight');
  assert.equal(views[0].getElementsByTagNameNS(NS, 'selection')[0].getAttribute('activeCell'), 'C5');
  assert.equal(views[1].getElementsByTagNameNS(NS, 'pane')[0].getAttribute('ySplit'), '3');
  const thaw = planWorksheetLayout(source, [{ operation: 'freeze', rows: 0, columns: 0 }]).fragments[0].xml;
  assert.equal(elements(thaw, 'sheetView')[0].getElementsByTagNameNS(NS, 'pane').length, 0);
});
test('missing and prefixed containers yield self-contained namespace-qualified fragments', () => {
  const source = `<s:worksheet xmlns:s="${NS}"><s:sheetData/></s:worksheet>`;
  const fragments = planWorksheetLayout(source, [
    { operation: 'freeze', rows: 1, columns: 0 },
    { operation: 'resizeColumn', index: 16383, width: 100 },
  ]).fragments;
  assert.deepEqual(
    fragments.map((f) => f.tag),
    ['sheetViews', 'cols']
  );
  for (const fragment of fragments) {
    const root = parse(fragment.xml).documentElement;
    assert.equal(root.namespaceURI, NS);
    assert.equal(root.localName, fragment.tag);
  }
  assert.equal(elements(fragments[1].xml, 'col')[0].getAttribute('min'), '16384');
});
test('ambiguous source geometry, entities and unqualified inputs are refused', () => {
  const book = (body) => `<worksheet xmlns="${NS}">${body}<sheetData/></worksheet>`;
  const resize = { operation: 'resizeColumn', index: 1, width: 22 };
  for (const input of [
    book('<cols><col min="1" max="3"/><col min="2" max="4"/></cols>'),
    book('<cols/><cols/>'),
    `<!DOCTYPE worksheet [<!ENTITY x "bad">]>${book('')}`,
  ])
    assert.throws(() => planWorksheetLayout(input, [resize]), /validation:/);
  for (const command of [
    { ...resize, width: 101 },
    { ...resize, index: 16384 },
    { ...resize, index: 0.5 },
    { ...resize, xml: 'injected' },
    { operation: 'freeze', rows: 101, columns: 0 },
    { operation: 'freeze', rows: 0, columns: 31 },
  ])
    assert.throws(() => planWorksheetLayout(book(''), [command]), /validation:/);
});
