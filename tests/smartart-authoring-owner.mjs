import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { spawnSync } from 'node:child_process';
import { unzipSync, zipSync, strToU8 } from 'fflate';
import { createWorkbook } from '../core/workbook-package.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { smartartSampleModel, normalizeSmartArt } from '../core/smartart-model.mjs';
after(() => closeNativeWorkbookWorker());
const id = '11111111-1111-4111-8111-ffffffffffff',
  other = '22222222-2222-4222-8222-222222222222';
const req = (o, operation, more = {}) =>
  o.request({ operation, documentId: o.status().documentId, state: o.status().state, ...more });
const preview = (o, objectId = id) => req(o, 'object', { sheet: 'Sheet1', objectId });
const create = (layout = 'process', more = {}) => ({
  operation: 'createSmartArt',
  sheet: 'Sheet1',
  objectId: id,
  expectedObjectRevision: null,
  smartart: smartartSampleModel(layout),
  altText: 'Semantic sample',
  anchor: { cell: 'E2' },
  ...more,
});
async function fixture(t) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'smartart-authoring-owner-')),
    source = path.join(root, 'source.xlsx'),
    parts = unzipSync(createWorkbook()),
    owners = [];
  parts['custom/smartart-opaque.bin'] = Uint8Array.from([0, 4, 255, 8]);
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Keep cells</t></is></c><c r="B1"><v>42</v></c></row><row r="2"><c r="A2"><f>B1+1</f><v>43</v></c></row></sheetData></worksheet>'
  );
  const bytes = Buffer.from(zipSync(parts));
  await fs.writeFile(source, bytes);
  let fault = null;
  const bridge = async (input) => {
    if (fault && input.operation === 'artifact-write' && input.name === 'package-metadata.json') {
      const p = JSON.parse(input.data);
      fault(p);
      input = { ...input, data: JSON.stringify(p) };
    }
    return native(input);
  };
  const open = async (file = source, recovery = path.join(root, 'recovery')) => {
    const o = new SpreadsheetDocument({ native: bridge, recoveryDirectory: recovery, cacheIdentity: null });
    owners.push(o);
    await o.open(file, { fresh: true });
    return o;
  };
  const o = await open();
  t.after(async () => {
    for (const owner of owners) await owner.close();
    await fs.rm(root, { recursive: true, force: true });
  });
  return { root, source, bytes, o, owners, open, fault: (f) => (fault = f) };
}
async function checkSaved(f, rows, name = 'saved') {
  const file = path.join(f.root, name + '.xlsx');
  await req(f.o, 'copy', { destination: file, expectedRevision: null });
  const expected = path.join(f.root, name + '.json');
  await fs.writeFile(expected, JSON.stringify(rows));
  const check = spawnSync(
    process.env.XLSX_EDITOR_PYTHON ?? 'python3',
    [path.resolve('tests/support/check-smartart-owner.py'), f.source, file, expected],
    { encoding: 'utf8' }
  );
  assert.equal(check.status, 0, check.stderr);
  console.log(check.stdout.trim());
  return file;
}
const dataFacts = (cells) =>
  Object.fromEntries(
    Object.entries(cells).map(([ref, cell]) => {
      const value = structuredClone(cell);
      if (value.calculation) delete value.calculation.revision;
      if (value.dataType === 'number' && value.rawValue !== null) value.rawValue = Number(value.rawValue);
      return [ref, value];
    })
  );
const expected = (p) => ({ id: p.id, model: p.smartart, anchor: p.anchor, alt: p.altText });
test('three SmartArt semantic layouts create/edit/resize and save/fresh reopen with independently checked graphs', async (t) => {
  for (const layout of ['process', 'hierarchy', 'cycle']) {
    const f = await fixture(t);
    const cells = (await req(f.o, 'read', { sheet: 'Sheet1', range: 'A1:B2' })).cells;
    await req(f.o, 'apply', { commands: [create(layout)] });
    let p = await preview(f.o);
    assert.deepEqual(p.smartart, smartartSampleModel(layout));
    assert.equal(p.canEdit, true);
    assert.equal(p.canDelete, true);
    const nodes = [
      { id: 'n1', text: 'Edited <&é' },
      { id: 'n2', text: 'Build "one"' },
      { id: 'n3', text: 'Review Ω' },
      { id: 'n4', text: 'Publish' },
    ];
    await req(f.o, 'apply', {
      commands: [
        {
          operation: 'updateSmartArt',
          sheet: 'Sheet1',
          objectId: id,
          expectedObjectRevision: p.revision,
          patch: { nodes },
          anchor: { cell: 'B4', dxEmu: 100, dyEmu: 200, widthEmu: 3810000, heightEmu: 2286000 },
          altText: 'Edited "diagram"',
        },
      ],
    });
    p = await preview(f.o);
    assert.deepEqual(p.smartart, normalizeSmartArt({ layout, nodes }));
    assert.equal(p.smartart.nodes.length, 4);
    assert.equal(p.anchor.cell, 'B4');
    assert.deepEqual(dataFacts((await req(f.o, 'read', { sheet: 'Sheet1', range: 'A1:B2' })).cells), dataFacts(cells));
    const file = await checkSaved(f, [expected(p)], 'edited-' + layout);
    await f.o.close();
    const cold = await f.open(file, path.join(f.root, 'cold'));
    assert.deepEqual((await preview(cold)).smartart, p.smartart);
    assert.deepEqual((await preview(cold)).anchor, p.anchor);
    assert.equal((await preview(cold)).altText, p.altText);
    assert.equal(cold.status().canUndo, false);
    assert.deepEqual(await fs.readFile(f.source), f.bytes);
  }
});
test('captured SmartArt topology edits have Undo/Redo, durable recovery and isolated sibling/final graph deletion', async (t) => {
  const f = await fixture(t),
    base = f.o.status().state,
    c = create(),
    pending = req(f.o, 'apply', { commands: [c] });
  c.smartart.nodes[0].text = 'Caller mutation';
  c.anchor.cell = 'A100';
  await pending;
  let p = await preview(f.o);
  assert.equal(p.smartart.nodes[0].text, 'Plan');
  assert.equal(p.anchor.cell, 'E2');
  const accepted = f.o.status().state;
  await req(f.o, 'undo');
  assert.equal(f.o.status().state, base);
  assert.equal((await req(f.o, 'objects', { sheet: 'Sheet1' })).total, 0);
  await req(f.o, 'redo');
  assert.equal(f.o.status().state, accepted);
  await req(f.o, 'apply', { commands: [create('hierarchy', { objectId: other })] });
  const sibling = await preview(f.o, other);
  await req(f.o, 'apply', {
    commands: [
      {
        operation: 'updateSmartArt',
        sheet: 'Sheet1',
        objectId: id,
        expectedObjectRevision: p.revision,
        patch: { layout: 'cycle' },
      },
    ],
  });
  p = await preview(f.o);
  assert.equal(p.smartart.connections.length, 3);
  assert.deepEqual(await preview(f.o, other), sibling);
  const documentId = f.o.status().documentId;
  await f.o.close();
  const recovered = new SpreadsheetDocument({ recoveryDirectory: path.join(f.root, 'recovery'), cacheIdentity: null });
  f.owners.push(recovered);
  await recovered.recover(documentId);
  f.o = recovered;
  p = await preview(f.o);
  assert.equal(p.smartart.layout, 'cycle');
  await req(f.o, 'apply', {
    commands: [{ operation: 'deleteSmartArt', sheet: 'Sheet1', objectId: id, expectedObjectRevision: p.revision }],
  });
  assert.equal((await req(f.o, 'objects', { sheet: 'Sheet1' })).total, 1);
  assert.deepEqual(await preview(f.o, other), sibling);
  await req(f.o, 'undo');
  assert.equal((await preview(f.o)).smartart.layout, 'cycle');
  await req(f.o, 'redo');
  const rest = await preview(f.o, other);
  await req(f.o, 'apply', {
    commands: [
      { operation: 'deleteSmartArt', sheet: 'Sheet1', objectId: other, expectedObjectRevision: rest.revision },
    ],
  });
  assert.equal((await req(f.o, 'objects', { sheet: 'Sheet1' })).total, 0);
  await checkSaved(f, [], 'empty');
});
test('all admitted two-to-six-node layouts produce finite native-validated cached geometry', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [create()] });
  for (const layout of ['process', 'hierarchy', 'cycle'])
    for (const count of [2, 3, 4, 5, 6]) {
      const p = await preview(f.o),
        nodes = Array.from({ length: count }, (_, i) => ({ id: 'n' + (i + 1), text: 'Node ' + (i + 1) }));
      await req(f.o, 'apply', {
        commands: [
          {
            operation: 'updateSmartArt',
            sheet: 'Sheet1',
            objectId: id,
            expectedObjectRevision: p.revision,
            patch: { layout, nodes },
            anchor: { widthEmu: 3333333, heightEmu: 2222222 },
          },
        ],
      });
      const q = await preview(f.o);
      assert.deepEqual(q.smartart, normalizeSmartArt({ layout, nodes }));
      await checkSaved(f, [expected(q)], layout + '-' + count);
    }
});
test('stale/invalid commands, forged semantic/cache/layout/type/edge plans and mixed failure cannot publish partial state', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [create()] });
  const p = await preview(f.o),
    before = structuredClone(f.o.status()),
    c = {
      operation: 'updateSmartArt',
      sheet: 'Sheet1',
      objectId: id,
      expectedObjectRevision: p.revision,
      patch: {
        nodes: [
          { id: 'n1', text: 'New plan' },
          { id: 'n2', text: 'Build' },
          { id: 'n3', text: 'Review' },
        ],
      },
    };
  for (const bad of [
    { ...c, expectedObjectRevision: '0'.repeat(64) },
    { ...c, patch: { nodes: [{ id: 'n1', text: 'One' }] } },
    {
      ...c,
      patch: {
        nodes: [
          { id: 'n1', text: 'x'.repeat(129) },
          { id: 'n2', text: 'Two' },
        ],
      },
    },
    { ...c, patch: { layout: 'foreign' } },
    { ...c, patch: { fontSize: 99 } },
    {
      ...c,
      patch: {
        nodes: [
          { id: 'n1', text: 'Line\nbreak' },
          { id: 'n2', text: 'Two' },
        ],
      },
    },
    { ...c, sheet: 'missing' },
  ]) {
    await assert.rejects(req(f.o, 'apply', { commands: [bad] }));
    assert.deepEqual(f.o.status(), before);
  }
  const part = (p, role) => Object.keys(p.replacements).find((n) => n.endsWith('-' + role + '.xml'));
  const faults = [
    (p) => (p.priorObjectRevision = '0'.repeat(64)),
    (p) =>
      (p.replacements[part(p, 'data')] = p.replacements[part(p, 'data')].replace(
        'type="parOf"',
        'type="unknownRelationship"'
      )),
    (p) =>
      (p.replacements[part(p, 'layout')] = p.replacements[part(p, 'layout')].replace('type="lin"', 'type="cycle"')),
    (p) => (p.replacements[part(p, 'drawing')] = p.replacements[part(p, 'drawing')].replace('sz="1400"', 'sz="9900"')),
    (p) => (p.replacements[part(p, 'drawing')] = p.replacements[part(p, 'drawing')].replace('New plan', 'Other label')),
    (p) => (p.replacements[part(p, 'colors')] = p.replacements[part(p, 'colors')].replace('4472C4', 'FF0000')),
    (p) => (p.replacements[part(p, 'style')] = p.replacements[part(p, 'style')].replace('idx="1"', 'idx="99"')),
    (p) =>
      (p.replacements[p.drawingPart] = p.replacements[p.drawingPart].replace(
        '<xdr:graphicFrame>',
        '<xdr:graphicFrame macro="run">'
      )),
    (p) => (p.replacements['custom/forged.xml'] = '<secret/>'),
    (p) => delete p.replacements[part(p, 'layout')],
    (p) => p.removed.push(part(p, 'data')),
  ];
  for (const fault of faults) {
    f.fault(fault);
    await assert.rejects(req(f.o, 'apply', { commands: [c] }));
    assert.deepEqual(f.o.status(), before);
  }
  f.fault(null);
  await assert.rejects(req(f.o, 'apply', { commands: [c, { sheet: 'missing', cell: 'A1', value: 99 }] }));
  assert.deepEqual(f.o.status(), before);
  assert.deepEqual(await fs.readFile(f.source), f.bytes);
});

test('native admission rejects forged drawing edges, graph content types and invalid XML text independently', async (t) => {
  const f = await fixture(t);
  await req(f.o, 'apply', { commands: [create()] });
  const source = path.join(f.root, 'before.xlsx');
  await req(f.o, 'copy', { destination: source, expectedRevision: null });
  const parts = unzipSync(await fs.readFile(source)),
    p = await preview(f.o),
    before = structuredClone(f.o.status()),
    c = {
      operation: 'updateSmartArt',
      sheet: 'Sheet1',
      objectId: id,
      expectedObjectRevision: p.revision,
      patch: {
        nodes: [
          { id: 'n1', text: 'New plan' },
          { id: 'n2', text: 'Build' },
          { id: 'n3', text: 'Review' },
        ],
      },
    };
  const rels = Object.keys(parts).find((n) => n.startsWith('xl/drawings/_rels/'));
  const faults = [
    (plan) => {
      plan.replacements[rels] = new TextDecoder()
        .decode(parts[rels])
        .replace(
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"',
          'TargetMode="Internal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"'
        );
    },
    (plan) => {
      plan.replacements[rels] = new TextDecoder()
        .decode(parts[rels])
        .replace(
          '</Relationships>',
          '<Relationship Id="forged-shared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/workbook-smartart-' +
            id +
            '-data.xml"/></Relationships>'
        );
    },
    (plan) => {
      plan.replacements['[Content_Types].xml'] = new TextDecoder()
        .decode(parts['[Content_Types].xml'])
        .replace('diagramData+xml', 'diagramLayout+xml');
    },
    (plan) => {
      for (const name of Object.keys(plan.replacements).filter(
        (n) => n.endsWith('-data.xml') || n.endsWith('-drawing.xml')
      ))
        plan.replacements[name] = plan.replacements[name].replace('New plan', 'New' + String.fromCharCode(1) + 'plan');
    },
  ];
  for (const fault of faults) {
    f.fault(fault);
    await assert.rejects(req(f.o, 'apply', { commands: [c] }));
    assert.deepEqual(f.o.status(), before);
  }
  f.fault(null);
});
