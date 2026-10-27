/** Source/readback checks for accepted conditional fills in immutable print jobs. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';
import { createWorkbook } from '../core/workbook-package.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { defaultPrintSettings } from '../core/print-model.mjs';
import { composePrint, printDocumentHtml } from '../core/print-layout.mjs';
import { conditionalRulesXml, differentialRulesXml } from '../core/rules-xml.mjs';
after(() => closeNativeWorkbookWorker());
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const settings = { ...defaultPrintSettings(), area: 'A1:B6' };
const cf = (range, kind, fillColor, extra = {}) => ({ range, kind, fillColor, stopIfTrue: false, ...extra });
const rules = [
  cf('B2:B6', 'greaterThan', '#FF0000', { values: [3] }),
  cf('B2:B6', 'lessThan', '#00FF00', { values: [2] }),
  cf('B2:B6', 'equal', '#0000FF', { values: [2] }),
  cf('B2:B6', 'between', '#FFFF00', { values: [3, 4] }),
  cf('A3:A3', 'textContains', '#FFA500', { text: 'Amber' }),
  cf('A2:A1002', 'duplicateValues', '#FF00FF'),
];
async function fixture(t, patch) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-print-conditional-')),
    source = path.join(directory, 'source.xlsx'),
    parts = unzipSync(createWorkbook());
  const rows = Array.from({ length: 1002 }, (_, i) => {
    const row = i + 1,
      value = row === 2 || row === 1002 ? 'Same' : row === 3 ? 'Amber' : 'unique-' + row;
    return `<row r="${row}"><c r="A${row}" t="inlineStr"><is><t>${value}</t></is></c><c r="B${row}"><v>${i}</v></c></row>`;
  }).join('');
  parts['xl/worksheets/sheet1.xml'] = strToU8(`<worksheet xmlns="${NS}"><sheetData>${rows}</sheetData></worksheet>`);
  parts['xl/styles.xml'] = strToU8(
    `<styleSheet xmlns="${NS}"><fonts count="1"><font><name val="Arial"/><sz val="11"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`
  );
  patch?.(parts);
  await fs.writeFile(source, zipSync(parts));
  const original = await fs.readFile(source);
  let root;
  const bridge = async (request) => {
    if (request.operation === 'index') root = request.root;
    return native(request);
  };
  const book = new IndexedWorkbook(
    bridge,
    () => assert.fail('No compatibility writer'),
    (request) => calculateIndexedWorkbook(request, bridge)
  );
  t.after(async () => {
    await book.close();
    await fs.rm(directory, { recursive: true, force: true });
  });
  await book.open(source);
  return { book, source, original, directory, root: () => root };
}
const command = (values, revision = null) => ({
  operation: 'setConditionalRules',
  sheet: 'Sheet1',
  rules: values,
  expectedRulesRevision: revision,
});
function importedRules(parts, fragment, styles) {
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    strFromU8(parts['xl/worksheets/sheet1.xml']).replace('</worksheet>', fragment + '</worksheet>')
  );
  parts['xl/styles.xml'] = strToU8(
    strFromU8(parts['xl/styles.xml']).replace(
      '</styleSheet>',
      '<dxfs count="' + styles.length + '">' + differentialRulesXml(styles) + '</dxfs></styleSheet>'
    )
  );
}
test('all six accepted fill types print with global duplicates, priority, immutable revisions and saved OOXML', async (t) => {
  const f = await fixture(t),
    base = await f.book.printSnapshot('0', 'Sheet1', settings),
    made = await f.book.apply('0', [command(rules)]),
    state = made.view.storage.state,
    snapshot = await f.book.printSnapshot(state, 'Sheet1', settings),
    fills = Object.fromEntries(
      snapshot.cells.filter((cell) => cell.style.fillColor).map((cell) => [cell.ref, cell.style.fillColor])
    );
  assert.deepEqual(fills, {
    A2: '#FF00FF',
    B2: '#00FF00',
    A3: '#FFA500',
    B3: '#0000FF',
    B4: '#FFFF00',
    B5: '#FF0000',
    B6: '#FF0000',
  });
  assert.ok(!base.cells.some((cell) => cell.style.fillColor));
  assert.notEqual(snapshot.fingerprint, base.fingerprint);
  assert.ok(snapshot.conditionalRevision);
  const html = printDocumentHtml(composePrint(snapshot));
  assert.match(html, /background:#FF00FF/);
  assert.match(html, /background:#FFFF00/);
  const changed = await f.book.apply(state, [{ sheet: 'Sheet1', cell: 'B5', value: 0 }]),
    later = await f.book.printSnapshot(changed.view.storage.state, 'Sheet1', settings);
  assert.equal(later.cells.find((cell) => cell.ref === 'B5').style.fillColor, '#00FF00');
  assert.equal(snapshot.cells.find((cell) => cell.ref === 'B5').style.fillColor, '#FF0000');
  const output = unzipSync(await fs.readFile((await f.book.export(state)).path)),
    doc = new DOMParser().parseFromString(strFromU8(output['xl/worksheets/sheet1.xml']), 'application/xml'),
    nodes = (name) => Array.from(doc.getElementsByTagNameNS(NS, name));
  // Independent raw XML count includes the distant duplicate beyond a read page.
  assert.equal(nodes('t').filter((node) => node.textContent === 'Same').length, 2);
  assert.deepEqual(
    nodes('cfRule').map((node) => node.getAttribute('priority')),
    ['1', '2', '3', '4', '5', '6']
  );
  assert.equal(nodes('conditionalFormatting').at(-1).getAttribute('sqref'), 'A2:A1002');
  await fs.writeFile(path.join(f.directory, 'saved.xlsx'), zipSync(output));
  const reopened = new IndexedWorkbook(native, () => assert.fail(), calculateIndexedWorkbook);
  try {
    await reopened.open(path.join(f.directory, 'saved.xlsx'));
    const saved = await reopened.printSnapshot('0', 'Sheet1', settings);
    assert.deepEqual(saved.cells, snapshot.cells);
  } finally {
    await reopened.close();
  }
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
test('advanced imported definitions refuse both snapshot and independent native proof without changing source', async (t) => {
  const f = await fixture(t, (parts) => {
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
        '</worksheet>',
        '<conditionalFormatting sqref="A2:A3"><cfRule type="colorScale" priority="1"><colorScale/></cfRule></conditionalFormatting></worksheet>'
      )
    );
  });
  await assert.rejects(
    f.book.printSnapshot('0', 'Sheet1', settings),
    (error) => error.code === 'PRINT_UNQUALIFIED_CONDITIONAL_FORMATTING'
  );
  await assert.rejects(
    native({
      action: 'workbook',
      operation: 'print-check',
      root: f.root(),
      version: 0,
      sheet: 'xl/worksheets/sheet1.xml',
      range: 'A1:B6',
    }),
    /conditional|Conditional|rule|Rule/
  );
  assert.deepEqual(await fs.readFile(f.source), f.original);
  assert.equal(f.book.diagnosticStats().retainedStates, 1);
});
test('unresolved distant duplicate source and unsafe numeric criteria preserve accepted state and refuse printing', async (t) => {
  const f = await fixture(t),
    made = await f.book.apply('0', [command(rules)]),
    state = made.view.storage.state,
    changed = await f.book.apply(state, [{ sheet: 'Sheet1', cell: 'A1002', formula: 'UNKNOWN(1)' }]),
    unresolved = changed.view.storage.state;
  await assert.rejects(
    f.book.printSnapshot(unresolved, 'Sheet1', settings),
    (error) => error.code === 'PRINT_UNQUALIFIED_CONDITIONAL_FORMATTING'
  );
  assert.equal(
    (await f.book.printSnapshot(state, 'Sheet1', settings)).cells.find((cell) => cell.ref === 'A2').style.fillColor,
    '#FF00FF'
  );
  assert.deepEqual(await fs.readFile(f.source), f.original);
  const g = await fixture(t, (parts) => {
      parts['xl/worksheets/sheet1.xml'] = strToU8(
        strFromU8(parts['xl/worksheets/sheet1.xml']).replace('<v>1</v>', '<v>9007199254740993</v>')
      );
    }),
    unsafe = await g.book.apply('0', [command([cf('B2:B6', 'greaterThan', '#FF0000', { values: [0] })])]);
  await assert.rejects(
    g.book.printSnapshot(unsafe.view.storage.state, 'Sheet1', settings),
    (error) => error.code === 'PRINT_UNQUALIFIED_CONDITIONAL_FORMATTING'
  );
  assert.deepEqual(await fs.readFile(g.source), g.original);
});
test('PDF duplicate policy retains typed, case-sensitive keys and combines blank with empty text', async (t) => {
  // This locks the existing editor/grid policy, not broad Excel duplicate parity.
  const f = await fixture(t, (parts) => {
      const replacements = {
        A2: '<c r="A2" t="inlineStr"><is><t>Same</t></is></c>',
        A3: '<c r="A3" t="inlineStr"><is><t>same</t></is></c>',
        A4: '<c r="A4" t="inlineStr"><is><t>Same</t></is></c>',
        A5: '<c r="A5" t="inlineStr"><is><t>1</t></is></c>',
        A6: '<c r="A6"><v>1</v></c>',
        A7: '<c r="A7" t="inlineStr"><is><t>1</t></is></c>',
        A8: '<c r="A8"/>',
        A9: '<c r="A9" t="inlineStr"><is><t></t></is></c>',
      };
      let xml = strFromU8(parts['xl/worksheets/sheet1.xml']);
      for (const [ref, cell] of Object.entries(replacements))
        xml = xml.replace(new RegExp('<c r="' + ref + '"[^>]*>[\\s\\S]*?</c>'), cell);
      parts['xl/worksheets/sheet1.xml'] = strToU8(xml);
    }),
    made = await f.book.apply('0', [command([cf('A2:A9', 'duplicateValues', '#FF00FF')])]),
    state = made.view.storage.state;
  const snapshot = await f.book.printSnapshot(state, 'Sheet1', { ...settings, area: 'A2:A9' }),
    fills = Object.fromEntries(snapshot.cells.map((cell) => [cell.ref, cell.style.fillColor ?? null]));
  assert.deepEqual(fills, {
    A2: '#FF00FF',
    A3: null,
    A4: '#FF00FF',
    A5: '#FF00FF',
    A6: null,
    A7: '#FF00FF',
    A8: '#FF00FF',
    A9: '#FF00FF',
  });
  const grid = await f.book.conditionalRules(state, 'Sheet1');
  assert.deepEqual(grid.duplicates['A2:A9'], { '"Same"': 2, '"1"': 2, '""': 2 });
  assert.equal(grid.duplicates['A2:A9']['"same"'], undefined);
  assert.equal(grid.duplicates['A2:A9']['1'], undefined);
  assert.match(printDocumentHtml(composePrint(snapshot)), /background:#FF00FF/);
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
test('imported physical rule order does not override numeric priority or change an immutable snapshot', async (t) => {
  const values = [
    cf('B2:B6', 'greaterThan', '#FF0000', { values: [0] }),
    cf('B2:B6', 'greaterThan', '#00FF00', { values: [0], stopIfTrue: true }),
  ];
  const blocks = conditionalRulesXml(values).match(/<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g);
  const f = await fixture(t, (parts) => importedRules(parts, blocks.toReversed().join(''), values)),
    definition = await f.book.conditionalRules('0', 'Sheet1'),
    snapshot = await f.book.printSnapshot('0', 'Sheet1', settings);
  assert.deepEqual(
    definition.rules.map((rule) => rule.fillColor),
    ['#FF0000', '#00FF00']
  );
  assert.equal(snapshot.cells.find((cell) => cell.ref === 'B2').style.fillColor, '#FF0000');
  const xml = new DOMParser().parseFromString(
    strFromU8(unzipSync(f.original)['xl/worksheets/sheet1.xml']),
    'application/xml'
  );
  assert.deepEqual(
    Array.from(xml.getElementsByTagNameNS(NS, 'cfRule')).map((node) => node.getAttribute('priority')),
    ['2', '1']
  );
  const changed = await f.book.apply('0', [{ sheet: 'Sheet1', cell: 'B2', value: 0 }]);
  assert.equal(
    (await f.book.printSnapshot(changed.view.storage.state, 'Sheet1', settings)).cells.find((cell) => cell.ref === 'B2')
      .style.fillColor,
    undefined
  );
  assert.equal(snapshot.cells.find((cell) => cell.ref === 'B2').style.fillColor, '#FF0000');
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
test('independent native print admission rejects duplicate aggregate, reversed bounds and unsafe priority', async (t) => {
  const variants = [
    [
      'duplicate aggregate',
      [cf('A1:A10000', 'duplicateValues', '#FF0000'), cf('B1:B10000', 'duplicateValues', '#00FF00')],
      (xml) => xml,
      /10000 source positions/,
    ],
    [
      'reversed between',
      [cf('B2:B6', 'between', '#FF0000', { values: [5, 1] })],
      (xml) => xml,
      /minimum exceeds maximum/,
    ],
    [
      'unsafe priority',
      [cf('B2:B6', 'equal', '#FF0000', { values: [1] })],
      (xml) => xml.replace('priority="1"', 'priority="9007199254740992"'),
      /priority/,
    ],
  ];
  for (const [label, values, mutate, reason] of variants)
    await t.test(label, async (sub) => {
      const f = await fixture(sub, (parts) => importedRules(parts, mutate(conditionalRulesXml(values)), values));
      await assert.rejects(
        native({
          action: 'workbook',
          operation: 'print-check',
          root: f.root(),
          version: 0,
          sheet: 'xl/worksheets/sheet1.xml',
          range: 'A1:B6',
        }),
        reason
      );
      assert.equal(f.book.diagnosticStats().retainedStates, 1);
      assert.deepEqual(await fs.readFile(f.source), f.original);
    });
  const values = [cf('A1:A10000', 'duplicateValues', '#FF0000'), cf('A1:A10000', 'duplicateValues', '#00FF00')],
    f = await fixture(t, (parts) => importedRules(parts, conditionalRulesXml(values), values));
  assert.equal(
    (
      await native({
        action: 'workbook',
        operation: 'print-check',
        root: f.root(),
        version: 0,
        sheet: 'xl/worksheets/sheet1.xml',
        range: 'A1:B6',
      })
    ).qualified,
    true,
    'Repeated duplicate rectangle is counted once'
  );
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
