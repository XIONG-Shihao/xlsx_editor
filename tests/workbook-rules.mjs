import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';
import { createWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { indexedRulesPlan } from '../core/indexed-rules.mjs';
import { conditionalFill } from '../core/rules-model.mjs';
after(() => closeNativeWorkbookWorker());
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  decode = strFromU8,
  doc = (bytes) => new DOMParser().parseFromString(decode(bytes), 'application/xml'),
  nodes = (document, name) => Array.from(document.getElementsByTagNameNS(NS, name));
async function fixture(t, patch) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-rules-')),
    file = path.join(directory, 'source.xlsx'),
    parts = unzipSync(createWorkbook());
  parts['custom-preserved.bin'] = new Uint8Array([1, 3, 5, 7]);
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${NS}"><sheetData>${Array.from({ length: 20 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>${i % 2 ? 'Same' : 'Other'}</t></is></c><c r="B${i + 1}"><v>${i + 1}</v></c></row>`).join('')}</sheetData></worksheet>`
  );
  patch?.(parts);
  await fs.writeFile(file, zipSync(parts));
  const book = new IndexedWorkbook(
    native,
    async () => {
      throw Error('Legacy writer is forbidden');
    },
    calculateIndexedWorkbook
  );
  t.after(async () => {
    await book.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  const view = await book.open(file, { storageDirectory: directory });
  return { directory, file, parts, book, view };
}
const cf = (kind, extra = {}) => ({ range: 'B1:B20', kind, fillColor: '#FF0000', stopIfTrue: false, ...extra });
const dv = (kind, extra = {}) => ({
  range: 'C1:C10',
  kind,
  allowBlank: true,
  showDropdown: kind === 'list',
  promptTitle: 'Choose',
  prompt: 'Use a permitted value',
  errorTitle: 'Invalid value',
  error: 'Use the allowed values',
  ...extra,
});
const command = (operation, rules, expectedRulesRevision = null) => ({
  operation,
  sheet: 'Sheet1',
  expectedRulesRevision,
  rules,
});
function untouched(old, next, claimed) {
  for (const [name, bytes] of Object.entries(old))
    if (!claimed.includes(name)) assert.deepEqual(next[name], bytes, 'Unclaimed part ' + name);
}
test('six fill rule types save exact criteria/priority and preserve unrelated parts through updates/reorder/delete', async (t) => {
  const f = await fixture(t),
    rules = [
      cf('greaterThan', { values: [5] }),
      cf('lessThan', { values: [2], fillColor: '#00FF00', stopIfTrue: true }),
      cf('equal', { values: [3] }),
      cf('between', { values: [4, 8] }),
      cf('textContains', { range: 'A1:A20', text: 'Sam*?~"' }),
      cf('duplicateValues', { range: 'A1:A20' }),
    ],
    made = await f.book.apply('0', [command('setConditionalRules', rules)]),
    state = made.view.storage.state,
    definition = await f.book.conditionalRules(state, 'Sheet1');
  assert.equal(definition.canEdit, true);
  assert.equal(definition.rules.length, 6);
  assert.equal(definition.duplicates['A1:A20']['"Same"'], 10);
  assert.deepEqual(made.commands[0].rules, rules);
  assert.equal((await f.book.conditionalRules('0', 'Sheet1')).rules.length, 0);
  const saved = unzipSync(await fs.readFile((await f.book.export(state)).path)),
    sheet = doc(saved['xl/worksheets/sheet1.xml']),
    styles = doc(saved['xl/styles.xml']),
    records = nodes(sheet, 'cfRule');
  assert.deepEqual(
    records.map((node) => node.getAttribute('priority')),
    ['1', '2', '3', '4', '5', '6']
  );
  assert.deepEqual(
    records.map((node) => node.getAttribute('type')),
    ['cellIs', 'cellIs', 'cellIs', 'cellIs', 'containsText', 'duplicateValues']
  );
  assert.equal(records[1].getAttribute('stopIfTrue'), '1');
  assert.equal(
    records[4].getElementsByTagNameNS(NS, 'formula')[0].textContent,
    'NOT(ISERROR(SEARCH("Sam~*~?~~""",A1)))'
  );
  assert.equal(nodes(styles, 'dxf').length, 6);
  assert.equal(
    nodes(styles, 'fgColor')
      .filter((node) => node.parentNode.localName === 'patternFill')
      .at(-6)
      .getAttribute('rgb'),
    'FFFF0000'
  );
  untouched(f.parts, saved, ['xl/styles.xml', 'xl/worksheets/sheet1.xml']);
  await assert.rejects(f.book.apply(state, [command('setConditionalRules', [], null)]), /Conflict/);
  const reversed = await f.book.apply(state, [
      command('setConditionalRules', [...rules].reverse(), definition.revision),
    ]),
    next = await f.book.conditionalRules(reversed.view.storage.state, 'Sheet1');
  assert.equal(next.rules[0].kind, 'duplicateValues');
  const cleared = await f.book.apply(reversed.view.storage.state, [command('setConditionalRules', [], next.revision)]);
  assert.equal((await f.book.conditionalRules(cleared.view.storage.state, 'Sheet1')).rules.length, 0);
  assert.equal(
    nodes(
      doc(unzipSync(await fs.readFile((await f.book.export(cleared.view.storage.state)).path))['xl/styles.xml']),
      'dxf'
    ).length,
    12
  );
});
test('guarded validation saves list/numeric settings, accepts literals and atomically rejects invalid batches and bulk bypasses', async (t) => {
  const f = await fixture(t),
    rules = [
      dv('list', { values: ['Yes', 'No'] }),
      dv('wholeNumber', { range: 'D1:D10', minimum: 1, maximum: 5 }),
      dv('decimal', { range: 'E1:E10', minimum: 0.5, maximum: 2.5 }),
      dv('textLength', { range: 'F1:F10', minimum: 2, maximum: 5, allowBlank: false }),
    ],
    made = await f.book.apply('0', [command('setValidationRules', rules)]),
    state = made.view.storage.state;
  assert.deepEqual((await f.book.validationRules(state, 'Sheet1')).rules, rules);
  const good = await f.book.apply(state, [
      {
        type: 'setCells',
        sheet: 'Sheet1',
        cells: [
          { cell: 'C1', value: 'Yes' },
          { cell: 'D1', value: 3 },
          { cell: 'E1', value: 1.2 },
          { cell: 'F1', value: 'four' },
        ],
      },
    ]),
    accepted = good.view.storage.state;
  await assert.rejects(
    f.book.apply(accepted, [
      {
        type: 'setCells',
        sheet: 'Sheet1',
        cells: [
          { cell: 'A20', value: 'must not commit' },
          { cell: 'D1', value: 3.2 },
        ],
      },
    ]),
    { code: 'rules-refused' }
  );
  assert.equal((await f.book.read(accepted, 'Sheet1', 'A20:A20')).cells.A20.value, 'Same');
  await assert.rejects(f.book.apply(accepted, [{ sheet: 'Sheet1', cell: 'C1', formula: '"Yes"' }]), {
    code: 'rules-unsupported',
  });
  await assert.rejects(
    f.book.apply(accepted, [{ type: 'fillSeries', sheet: 'Sheet1', range: 'D1:D5' }]),
    /Bulk editing/
  );
  const renamed = await f.book.apply(accepted, [{ operation: 'renameSheet', sheet: 'Sheet1', newName: 'Other' }]);
  assert.equal((await f.book.validationRules(renamed.view.storage.state, 'Other')).rules[0].range, 'C1:C10');
  const saved = unzipSync(await fs.readFile((await f.book.export(accepted)).path)),
    xml = doc(saved['xl/worksheets/sheet1.xml']),
    defs = nodes(xml, 'dataValidation');
  assert.equal(defs.length, 4);
  assert.equal(defs[0].getAttribute('showDropDown'), '0');
  assert.equal(defs[0].getAttribute('errorStyle'), 'stop');
  assert.equal(defs[0].getElementsByTagNameNS(NS, 'formula1')[0].textContent, '"Yes,No"');
  assert.equal(defs[1].getAttribute('type'), 'whole');
  assert.equal(defs[2].getElementsByTagNameNS(NS, 'formula2')[0].textContent, '2.5');
  assert.equal(defs[3].getAttribute('allowBlank'), '0');
  const definition = await f.book.validationRules(accepted, 'Sheet1'),
    clear = await f.book.apply(accepted, [command('setValidationRules', [], definition.revision)]);
  assert.equal((await f.book.validationRules(clear.view.storage.state, 'Sheet1')).rules.length, 0);
  assert.equal((await f.book.validationRules(state, 'Sheet1')).rules.length, 4);
});
test('opaque imported rules are preserved/refused and style edits survive append-only differential fills', async (t) => {
  const f = await fixture(t, (parts) => {
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      decode(parts['xl/worksheets/sheet1.xml']).replace(
        '</worksheet>',
        '<conditionalFormatting sqref="A1:A3"><cfRule type="colorScale" priority="1"><colorScale/></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="custom" sqref="C1:C3"><formula1>LEN(C1)&gt;0</formula1></dataValidation></dataValidations></worksheet>'
      )
    );
  });
  assert.equal((await f.book.conditionalRules('0', 'Sheet1')).canEdit, false);
  assert.equal((await f.book.validationRules('0', 'Sheet1')).canEdit, false);
  await assert.rejects(f.book.apply('0', [command('setConditionalRules', [])]), /protected/);
  // A validation rule the editor cannot evaluate is kept and does not stop an edit.
  const edited = await f.book.apply('0', [{ sheet: 'Sheet1', cell: 'C1', value: 'edit' }]);
  assert.match(
    decode(
      unzipSync(await fs.readFile((await f.book.export(edited.view.storage.state)).path))['xl/worksheets/sheet1.xml']
    ),
    /<dataValidation type="custom" sqref="C1:C3"/
  );
  assert.deepEqual(await fs.readFile((await f.book.export('0')).path), await fs.readFile(f.file));
  const g = await fixture(t),
    styled = await g.book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', bold: true, fillColor: 'AABBCC' }]),
    made = await g.book.apply(styled.view.storage.state, [
      command('setConditionalRules', [cf('equal', { values: [1] })]),
    ]);
  const read = await g.book.read(made.view.storage.state, 'Sheet1', 'A1:A1');
  assert.equal(read.cells.A1.bold, true);
  assert.equal(read.cells.A1.fillColor, 'AABBCC');
});
test('native independently rejects arbitrary styles/fragments/stale definitions/foreign scope before package publication', async (t) => {
  const f = await fixture(t),
    root = f.view.storage?.root;
  void root;
  const storage = await fs.readdir(f.directory),
    session = storage.find((name) => name.startsWith('owner-'));
  const actualRoot = path.join(f.directory, session, 'base'),
    metadata = await fs.readFile(path.join(actualRoot, 'metadata.xlsx')),
    prepared = indexedRulesPlan(metadata, f.view, command('setConditionalRules', [cf('equal', { values: [1] })]));
  const failures = [
    (plan) => (plan.replacements['xl/styles.xml'] = plan.replacements['xl/styles.xml'].replace('<font>', '<font><b/>')),
    (plan) =>
      (plan.fragments[plan.sheetPart].conditionalFormatting = plan.fragments[
        plan.sheetPart
      ].conditionalFormatting.replace('equal', 'greaterThan')),
    (plan) => (plan.priorRulesRevision = 'a'.repeat(64)),
    (plan) => (plan.sheetName = 'foreign'),
    (plan) => (plan.scope = 2),
    (plan) => plan.removed.push('custom-preserved.bin'),
    (plan) => (plan.rules[0].extra = 'opaque'),
  ];
  for (const fail of failures) {
    const plan = structuredClone(prepared.plan);
    fail(plan);
    await native({
      action: 'workbook',
      operation: 'artifact-write',
      root: actualRoot,
      name: 'package-metadata.json',
      data: JSON.stringify(plan),
    });
    await assert.rejects(
      native({ action: 'workbook', operation: 'package-edit', root: actualRoot, version: 0, source: 'source.xlsx' }),
      /rule|Rule|Rules|Conflict|worksheet|style|scope/i
    );
    await assert.rejects(fs.stat(path.join(actualRoot, 'package.xlsx')), /ENOENT/);
  }
  assert.equal(f.book.diagnosticStats().retainedStates, 1);
});
test('allow-blank Clear works, required fields refuse, and opaque validation leaves unrelated cells editable', async (t) => {
  const f = await fixture(t),
    made = await f.book.apply('0', [
      command('setValidationRules', [
        dv('list', { values: ['Yes', 'No'] }),
        dv('wholeNumber', { range: 'D1:D5', minimum: 1, maximum: 5, allowBlank: false }),
      ]),
    ]),
    filled = await f.book.apply(made.view.storage.state, [{ sheet: 'Sheet1', cell: 'C1', value: 'Yes' }]),
    cleared = await f.book.apply(filled.view.storage.state, [{ type: 'clearRange', sheet: 'Sheet1', range: 'C1:C5' }]);
  assert.equal((await f.book.read(cleared.view.storage.state, 'Sheet1', 'C1')).cells.C1?.value ?? null, null);
  await assert.rejects(
    f.book.apply(cleared.view.storage.state, [{ type: 'clearRange', sheet: 'Sheet1', range: 'D1:D5' }]),
    /Bulk editing/
  );
  const g = await fixture(t, (parts) => {
      parts['xl/worksheets/sheet1.xml'] = strToU8(
        decode(parts['xl/worksheets/sheet1.xml']).replace(
          '</worksheet>',
          '<dataValidations count="1"><dataValidation type="custom" sqref="C1:C3" showErrorMessage="1"><formula1>LEN(C1)&gt;0</formula1></dataValidation></dataValidations></worksheet>'
        )
      );
    }),
    outside = await g.book.apply('0', [{ sheet: 'Sheet1', cell: 'A1', value: 'Still editable' }]);
  assert.equal((await g.book.read(outside.view.storage.state, 'Sheet1', 'A1')).cells.A1.value, 'Still editable');
  const inside = await g.book.apply(outside.view.storage.state, [{ sheet: 'Sheet1', cell: 'C1', value: 'accepted' }]);
  assert.equal((await g.book.read(inside.view.storage.state, 'Sheet1', 'C1')).cells.C1.value, 'accepted');
  const source = decode(
    unzipSync(await fs.readFile((await g.book.export(outside.view.storage.state)).path))['xl/worksheets/sheet1.xml']
  );
  assert.match(
    source,
    /<dataValidation type="custom" sqref="C1:C3" showErrorMessage="1"><formula1>LEN\(C1\)&gt;0<\/formula1>/
  );
});
test('duplicate metadata reads are bounded and unqualified evaluation keeps collection editing available', async (t) => {
  const f = await fixture(t),
    made = await f.book.apply('0', [command('setConditionalRules', [cf('duplicateValues', { range: 'A1:A42' })])]),
    state = made.view.storage.state,
    values = Array.from({ length: 42 }, (_, index) => ({
      cell: 'A' + (index + 1),
      value: String(index).padStart(3, '0') + 'x'.repeat(32000),
    })),
    changed = await f.book.apply(state, [{ type: 'setCells', sheet: 'Sheet1', cells: values }]),
    definition = await f.book.conditionalRules(changed.view.storage.state, 'Sheet1');
  assert.equal(definition.canEdit, true);
  assert.equal(definition.rules.length, 1);
  assert.deepEqual(definition.duplicates, {});
  assert.match(definition.evaluationReasons[0], /bounded metadata/);
  const deleted = await f.book.apply(changed.view.storage.state, [
    command('setConditionalRules', [], definition.revision),
  ]);
  assert.equal((await f.book.conditionalRules(deleted.view.storage.state, 'Sheet1')).rules.length, 0);
  await assert.rejects(
    f.book.apply('0', [
      command('setConditionalRules', [
        cf('duplicateValues', { range: 'A1:A10000' }),
        cf('duplicateValues', { range: 'B1:B10000' }),
      ]),
    ]),
    { code: /^(fill-limit|rules-duplicate)$/ }
  );
});
test('native CF admission independently checks actual styles graph and XML characters', async (t) => {
  const f = await fixture(t, (parts) => {
      parts['[Content_Types].xml'] = strToU8(
        decode(parts['[Content_Types].xml']).replace(
          'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
          'application/octet-stream'
        )
      );
    }),
    session = (await fs.readdir(f.directory)).find((name) => name.startsWith('owner-')),
    root = path.join(f.directory, session, 'base'),
    metadata = await fs.readFile(path.join(root, 'metadata.xlsx')),
    prepared = indexedRulesPlan(metadata, f.view, command('setConditionalRules', [cf('equal', { values: [1] })]));
  await native({
    action: 'workbook',
    operation: 'artifact-write',
    root,
    name: 'package-metadata.json',
    data: JSON.stringify(prepared.plan),
  });
  await assert.rejects(
    native({ action: 'workbook', operation: 'package-edit', root, version: 0, source: 'source.xlsx' }),
    { code: 'rules-unsupported' }
  );
  await assert.rejects(fs.stat(path.join(root, 'package.xlsx')), /ENOENT/);
  const g = await fixture(t),
    owner = (await fs.readdir(g.directory)).find((name) => name.startsWith('owner-')),
    groot = path.join(g.directory, owner, 'base'),
    bad = indexedRulesPlan(
      await fs.readFile(path.join(groot, 'metadata.xlsx')),
      g.view,
      command('setValidationRules', [dv('list', { values: ['Yes', 'No'] })])
    );
  bad.plan.rules[0].values = ['bad\u0000value'];
  await native({
    action: 'workbook',
    operation: 'artifact-write',
    root: groot,
    name: 'package-metadata.json',
    data: JSON.stringify(bad.plan),
  });
  await assert.rejects(
    native({ action: 'workbook', operation: 'package-edit', root: groot, version: 0, source: 'source.xlsx' }),
    { code: /^(package-malformed|package-unsupported|rules-malformed)$/ }
  );
  await assert.rejects(fs.stat(path.join(groot, 'package.xlsx')), /ENOENT/);
});
