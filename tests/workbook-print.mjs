import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { calculateIndexedWorkbook } from '../core/indexed-calculation.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { deflateSync, crc32 } from 'node:zlib';
import { defaultPrintSettings } from '../core/print-model.mjs';
import { composePrint, printDocumentHtml } from '../core/print-layout.mjs';
import { indexedPrintPlan } from '../core/indexed-print.mjs';
import { printLayoutCase, checkPrintLayout } from '../scripts/test/w7-print-layout-oracle.mjs';
after(() => closeNativeWorkbookWorker());
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  base = unzipSync(await fs.readFile('tests/support/w7-reference-data/source.xlsx'));
const xml = (value) => String(value).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]);
// Independent protocol fixture: one red and one blue RGB pixel, ratio2:1.
function twoPixelPng() {
  const chunk = (name, data) => {
      const type = Buffer.from(name),
        length = Buffer.alloc(4),
        checksum = Buffer.alloc(4);
      length.writeUInt32BE(data.length);
      checksum.writeUInt32BE(crc32(Buffer.concat([type, data])));
      return Buffer.concat([length, type, data, checksum]);
    },
    header = Buffer.alloc(13);
  header.writeUInt32BE(2, 0);
  header.writeUInt32BE(1, 4);
  header[8] = 8;
  header[9] = 2;
  return Buffer.concat([
    Buffer.from('89504e470d0a1a0a', 'hex'),
    chunk('IHDR', header),
    chunk('IDAT', deflateSync(Buffer.from([0, 255, 0, 0, 0, 0, 255]))),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}
async function fixture(
  t,
  { rows = 50, columns = 3, widths = [72, 108, 90], rowHeight = 18, hidden = [], merges = [], patch } = {}
) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-print-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const parts = { ...base };
  parts['xl/workbook.xml'] = strToU8(strFromU8(base['xl/workbook.xml']).replace('name="Sales"', 'name="Print"'));
  const cols = widths
      .map((width, index) => `<col min="${index + 1}" max="${index + 1}" width="${width / 0.75 / 7}" customWidth="1"/>`)
      .join(''),
    data = Array.from(
      { length: rows },
      (_, r) =>
        `<row r="${r + 1}"${hidden.includes(r + 1) ? ' hidden="1"' : ''}>` +
        Array.from(
          { length: columns },
          (_, c) =>
            `<c r="${String.fromCharCode(65 + c)}${r + 1}" t="inlineStr"><is><t>${r === 0 && c === 0 ? 'Item' : 'row-' + (r + 1)}</t></is></c>`
        ).join('') +
        '</row>'
    ).join('');
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${ns}"><sheetFormatPr defaultRowHeight="${rowHeight}"/><cols>${cols}</cols><sheetData>${data}</sheetData>${merges.length ? '<mergeCells count="' + merges.length + '">' + merges.map((ref) => '<mergeCell ref="' + ref + '"/>').join('') + '</mergeCells>' : ''}</worksheet>`
  );
  if (patch) patch(parts);
  const source = path.join(directory, 'source.xlsx');
  await fs.writeFile(source, zipSync(parts));
  const original = await fs.readFile(source);
  let root;
  const bridge = async (input) => {
    if (input.operation === 'index') root = input.root;
    return native(input);
  };
  const book = new IndexedWorkbook(
    bridge,
    () => assert.fail('Print cannot invoke compatibility writer'),
    (input) => calculateIndexedWorkbook(input, bridge)
  );
  await book.open(source);
  t.after(() => book.close());
  const baseline = await book.export('0'),
    exportParts = unzipSync(await fs.readFile(baseline.path));
  return { book, source, directory, parts, exportParts, original, root: () => root };
}
const settings = (area = 'A1:C50', patch = {}) => ({ ...defaultPrintSettings(), area, ...patch });
const command = (value, expected = null) => ({
  operation: 'setPrintSettings',
  sheet: 'Print',
  expectedPrintRevision: expected,
  settings: value,
});

// Existing export normalization can reorder only the generated dimension's two
// attributes between an original and reindexed export. Keep all other bytes
// exact and retain the same reference/namespace; this is not broad XML equality.
function unchangedExportPart(actual, expected, part) {
  const canonical = (bytes) =>
    strFromU8(bytes).replace(
      /<dimension (?:ref="([^"]+)" xmlns="http:\/\/schemas.openxmlformats.org\/spreadsheetml\/2006\/main"|xmlns="http:\/\/schemas.openxmlformats.org\/spreadsheetml\/2006\/main" ref="([^"]+)")\/>/g,
      (_match, a, b) => '<dimension ref="' + (a ?? b) + '" xmlns="' + ns + '"/>'
    );
  if (/^xl\/worksheets\/sheet\d+\.xml$/.test(part))
    assert.equal(canonical(actual), canonical(expected), 'unrelated ' + part);
  else assert.deepEqual(actual, expected, 'unrelated ' + part);
}

test('actual print settings use scoped reserved names, save/reopen and one immutable owner state', async (t) => {
  const f = await fixture(t),
    initial = await f.book.printDefinition('0', 'Print');
  assert.equal(initial.revision, null);
  const made = await f.book.apply('0', [command(settings())]),
    state = made.view.storage.state,
    definition = await f.book.printDefinition(state, 'Print');
  assert.equal(definition.canEdit, true);
  assert.deepEqual(definition.settings, settings());
  assert.equal(f.book.diagnosticStats().retainedStates, 2);
  assert.deepEqual(made.commands, made.edits);
  assert.equal((await f.book.printDefinition('0', 'Print')).revision, null);
  const saved = await f.book.export(state),
    parts = unzipSync(await fs.readFile(saved.path)),
    workbook = strFromU8(parts['xl/workbook.xml']);
  await fs.copyFile(saved.path, path.join(f.directory, 'saved.xlsx'));
  assert.match(workbook, /_xlnm.Print_Area[^>]*localSheetId="0"[^>]*>'Print'!\$A\$1:\$C\$50/);
  assert.match(
    strFromU8(parts['xl/worksheets/sheet1.xml']),
    /<pageSetup[^>]*paperSize="9"[^>]*orientation="portrait"[^>]*scale="100"/
  );
  for (const [part, bytes] of Object.entries(f.exportParts))
    if (!['xl/workbook.xml', 'xl/worksheets/sheet1.xml'].includes(part)) unchangedExportPart(parts[part], bytes, part);
  await assert.rejects(f.book.apply(state, [command(settings('A1:C4'), null)]), {
    code: /^(print-stale|rules-stale)$/,
  });
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
test('accepted offscreen snapshot produces all two-page cells against the independent physical oracle', async (t) => {
  const f = await fixture(t),
    snapshot = await f.book.printSnapshot('0', 'Print', settings()),
    layout = composePrint(snapshot);
  assert.equal(snapshot.cells.length, 150);
  assert.equal(snapshot.objects.length, 0);
  assert.equal(checkPrintLayout(layout, 'a4-portrait-two-pages').passed, true);
  for (const page of layout.pages)
    for (const cell of page.cells) {
      const row = Number(/\d+$/.exec(cell.ref)[0]);
      assert.equal(cell.text, cell.ref === 'A1' ? 'Item' : 'row-' + row);
    }
  const frozen = structuredClone(snapshot),
    changed = await f.book.apply('0', [{ sheet: 'Print', cell: 'C50', value: 'later' }]);
  assert.equal(snapshot.cells.find((cell) => cell.ref === 'C50').text, 'row-50');
  assert.deepEqual(snapshot, frozen);
  assert.equal(
    (await f.book.printSnapshot(changed.view.storage.state, 'Print', settings())).cells.find(
      (cell) => cell.ref === 'C50'
    ).text,
    'later'
  );
  const html = printDocumentHtml(layout);
  assert.equal((html.match(/data-xlsx-print-page=/g) ?? []).length, 2);
  assert.ok(html.includes('row-50'));
  assert.ok(!html.includes('<script'));
  assert.equal(f.book.diagnosticStats().retainedStates, 2);
});
test('repeat prefixes, fit policy and hidden/merged geometry stay explicit', async (t) => {
  const f = await fixture(t),
    s = settings('A1:C50', {
      paper: 'Letter',
      scaling: { mode: 'fit', width: 1, height: 2 },
      repeatRows: { first: 1, last: 2 },
      repeatColumns: { first: 'A', last: 'A' },
      gridlines: true,
    }),
    made = await f.book.apply('0', [command(s)]),
    state = made.view.storage.state,
    definition = await f.book.printDefinition(state, 'Print');
  assert.deepEqual(definition.settings, s);
  const layout = composePrint(await f.book.printSnapshot(state, 'Print'));
  assert.ok(
    layout.pages.every(
      (page) => page.rowIndices.includes(0) && page.rowIndices.includes(1) && page.columnIndices.includes(0)
    )
  );
  const g = await fixture(t, {
      rows: 4,
      rowHeight: 24,
      hidden: [3],
      merges: ['B2:C2'],
      patch(parts) {
        parts['xl/worksheets/sheet1.xml'] = strToU8(
          strFromU8(parts['xl/worksheets/sheet1.xml'])
            .replace(
              '<c r="B2" t="inlineStr"><is><t>row-2</t></is></c>',
              '<c r="B2" t="inlineStr"><is><t>Merged heading</t></is></c>'
            )
            .replace('<c r="C4" t="inlineStr"><is><t>row-4</t></is></c>', '<c r="C4" s="1"><v>60</v></c>')
        );
        parts['xl/styles.xml'] = strToU8(
          `<styleSheet xmlns="${ns}"><fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`
        );
        parts['xl/_rels/workbook.xml.rels'] = strToU8(
          strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
            '</Relationships>',
            '<Relationship Id="rPrintStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
          )
        );
        parts['[Content_Types].xml'] = strToU8(
          strFromU8(parts['[Content_Types].xml']).replace(
            '</Types>',
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
          )
        );
      },
    }),
    hiddenLayout = composePrint(
      await g.book.printSnapshot('0', 'Print', settings('A1:C4', { paper: 'Letter', orientation: 'landscape' }))
    );
  assert.equal(checkPrintLayout(hiddenLayout, 'letter-landscape-hidden-and-merge').passed, true);
});
test('merge cuts, unresolved source, rich runs, unsupported styles and advanced imported print settings refuse', async (t) => {
  const cut = await fixture(t, { merges: ['B42:C43'] });
  await assert.rejects(
    async () => composePrint(await cut.book.printSnapshot('0', 'Print', settings())),
    (error) => error.code === 'PRINT_MERGE_PAGE_CUT'
  );
  const unresolved = await fixture(t, {
    rows: 4,
    patch(parts) {
      parts['xl/worksheets/sheet1.xml'] = strToU8(
        strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
          '<c r="B3" t="inlineStr"><is><t>row-3</t></is></c>',
          '<c r="B3"><f>UNKNOWN(1)</f><v>999</v></c>'
        )
      );
    },
  });
  await assert.rejects(
    unresolved.book.printSnapshot('0', 'Print', settings('A1:C4')),
    (error) => error.code === 'PRINT_UNRESOLVED_CALCULATION'
  );
  const rich = await fixture(t, {
    rows: 4,
    patch(parts) {
      parts['xl/worksheets/sheet1.xml'] = strToU8(
        strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
          '<is><t>Item</t></is>',
          '<is><r><rPr><b/></rPr><t>Item</t></r></is>'
        )
      );
    },
  });
  await assert.rejects(rich.book.printSnapshot('0', 'Print', settings('A1:C4')), /Rich|rich/);
  const advanced = await fixture(t, {
    rows: 4,
    patch(parts) {
      parts['xl/worksheets/sheet1.xml'] = strToU8(
        strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
          '</worksheet>',
          '<headerFooter><oddHeader>Protected</oddHeader></headerFooter></worksheet>'
        )
      );
    },
  });
  assert.equal((await advanced.book.printDefinition('0', 'Print')).canEdit, false);
  await assert.rejects(advanced.book.apply('0', [command(settings('A1:C4'))]), /protected/);
  assert.deepEqual(await fs.readFile(advanced.source), advanced.original);
});
test('native print plans validate exact prior definition and refuse unrelated workbook/sheet metadata before output', async (t) => {
  const f = await fixture(t, { rows: 4 }),
    metadata = await fs.readFile(path.join(f.root(), 'metadata.xlsx')),
    prepared = indexedPrintPlan(metadata, f.book.view('0'), command(settings('A1:C4'))),
    source = await f.book.export('0'),
    output = path.join(f.root(), 'package.xlsx'),
    sentinel = Buffer.from('previous output');
  await fs.writeFile(output, sentinel);
  const changes = [
    (plan) => (plan.priorPrintRevision = '0'.repeat(64)),
    (plan) => (plan.sheetName = 'Other'),
    (plan) => (plan.scope = 1),
    (plan) =>
      (plan.replacements['xl/workbook.xml'] = plan.replacements['xl/workbook.xml'].replace(
        'name="Other"',
        'name="Changed"'
      )),
    (plan) =>
      (plan.fragments[plan.sheetPart].sheetPr = plan.fragments[plan.sheetPart].sheetPr.replace(
        '</sheetPr>',
        '<tabColor rgb="FFFF0000"/></sheetPr>'
      )),
    (plan) =>
      (plan.fragments[plan.sheetPart].pageSetup = plan.fragments[plan.sheetPart].pageSetup.replace(
        'scale="100"',
        'scale="99"'
      )),
    (plan) => (plan.settings.extra = 'unqualified'),
  ];
  for (const change of changes) {
    const plan = structuredClone(prepared.plan);
    change(plan);
    await fs.writeFile(path.join(f.root(), 'package-metadata.json'), JSON.stringify(plan));
    await assert.rejects(
      native({
        action: 'workbook',
        operation: 'package-edit',
        root: f.root(),
        version: 0,
        source: path.basename(source.path),
      }),
      /Print|print|Conflict|workbook|sheetPr/
    );
    assert.deepEqual(await fs.readFile(output), sentinel);
    await assert.rejects(fs.stat(path.join(f.root(), 'package.partial')), (error) => error.code === 'ENOENT');
  }
});
test('real managed column/PNG objects share the exact print scale and area clipping, with all typed text retained', async (t) => {
  const values = [
      ['Item', 'Sales', 'Costs', 'Notes'],
      ['Jan', 2, 1, 'note-2'],
      ['Feb', 3, 2, 'note-3'],
      ['Mar', 5, 3, '#DIV/0!'],
    ],
    f = await fixture(t, {
      rows: 4,
      columns: 4,
      widths: [144, 144, 144, 144],
      rowHeight: 36,
      patch(parts) {
        const rows = values
          .map(
            (row, r) =>
              '<row r="' +
              (r + 1) +
              '">' +
              row
                .map((value, c) => {
                  const ref = String.fromCharCode(65 + c) + (r + 1);
                  return typeof value === 'number'
                    ? `<c r="${ref}"><v>${value}</v></c>`
                    : value === '#DIV/0!'
                      ? `<c r="${ref}" t="e"><v>${value}</v></c>`
                      : `<c r="${ref}" t="inlineStr"><is><t>${xml(value)}</t></is></c>`;
                })
                .join('') +
              '</row>'
          )
          .join('');
        parts['xl/worksheets/sheet1.xml'] = strToU8(
          strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
            /<sheetData>[\s\S]*?<\/sheetData>/,
            '<sheetData>' + rows + '</sheetData>'
          )
        );
      },
    }),
    chartId = '11111111-1111-4111-8111-111111111111',
    imageId = '22222222-2222-4222-8222-222222222222',
    png = twoPixelPng().toString('base64');
  const chart = await f.book.apply('0', [
      {
        operation: 'createChart',
        sheet: 'Print',
        objectId: chartId,
        expectedObjectRevision: null,
        chart: { kind: 'column', range: 'A1:C4', title: 'Sales trend' },
        anchor: { cell: 'A4', widthEmu: 240 * 12700, heightEmu: 120 * 12700 },
      },
    ]),
    both = await f.book.apply(chart.view.storage.state, [
      {
        operation: 'insertImage',
        sheet: 'Print',
        objectId: imageId,
        expectedObjectRevision: null,
        asset: { mime: 'image/png', base64: png },
        anchor: { cell: 'B2', dxEmu: 24 * 12700, dyEmu: 12 * 12700, widthEmu: 108 * 12700, heightEmu: 54 * 12700 },
        altText: 'Printed logo',
      },
    ]),
    snapshot = await f.book.printSnapshot(
      both.view.storage.state,
      'Print',
      settings('A1:D4', { paper: 'Letter', orientation: 'landscape', scaling: { mode: 'scale', percent: 50 } })
    ),
    layout = composePrint(snapshot);
  for (const cell of layout.pages[0].cells) {
    const c = cell.ref.charCodeAt(0) - 65,
      r = Number(cell.ref.slice(1)) - 1;
    assert.equal(cell.text, String(values[r][c]));
  }
  assert.equal(snapshot.objects.find((object) => object.id === imageId).image.base64, png);
  assert.deepEqual(
    snapshot.objects.find((object) => object.id === chartId).data.series.map((item) => item.values),
    [
      [2, 3, 5],
      [1, 2, 3],
    ]
  );
  const expected = JSON.parse(await fs.readFile('tests/support/w7-print-layout-expected.json', 'utf8')),
    entry = expected.cases.find((item) => item.id === 'letter-half-scale-objects');
  entry.expected.pages[0].objects[0].id = imageId;
  entry.expected.pages[0].objects[1].id = chartId;
  const expectation = path.join(f.directory, 'expected-layout.json');
  await fs.writeFile(expectation, JSON.stringify(expected));
  assert.equal(checkPrintLayout(layout, 'letter-half-scale-objects', expectation).passed, true);
  const html = printDocumentHtml(layout);
  assert.match(html, /<svg /);
  assert.ok(html.includes('data:image/png;base64,' + png));
  assert.ok(html.includes('data-print-object="' + chartId + '"'));
});
test('discrete fit respects requested page counts with wide repeated columns; horizontal merge cuts refuse', async (t) => {
  const f = await fixture(t, { rows: 5, columns: 6, widths: Array(6).fill(300) }),
    s = settings('A1:F5', {
      paper: 'Letter',
      orientation: 'landscape',
      scaling: { mode: 'fit', width: 2, height: 0 },
      repeatColumns: { first: 'A', last: 'A' },
    }),
    layout = composePrint(await f.book.printSnapshot('0', 'Print', s));
  assert.equal(layout.pages.length, 2);
  assert.ok(layout.pages.every((page) => page.columnIndices.includes(0)));
  assert.ok(layout.scale <= 0.601 && layout.scale >= 0.599);
  const cut = await fixture(t, { rows: 3, columns: 4, widths: Array(4).fill(200), merges: ['B1:C1'] });
  await assert.rejects(
    async () => composePrint(await cut.book.printSnapshot('0', 'Print', settings('A1:D3', { paper: 'Letter' }))),
    (error) => error.code === 'PRINT_MERGE_PAGE_CUT'
  );
});
test('plain shared strings qualify while rich shared strings and opaque whole-cell styles are protected', async (t) => {
  const shared = (rich) => (parts) => {
    parts['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
        '<c r="A1" t="inlineStr"><is><t>Item</t></is></c>',
        '<c r="A1" t="s"><v>0</v></c>'
      )
    );
    parts['xl/sharedStrings.xml'] = strToU8(
      `<sst xmlns="${ns}" count="1" uniqueCount="1"><si>${rich ? '<r><rPr><b/></rPr><t>Item</t></r>' : '<t>Item</t>'}</si></sst>`
    );
    parts['xl/_rels/workbook.xml.rels'] = strToU8(
      strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
        '</Relationships>',
        '<Relationship Id="rPrintStrings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'
      )
    );
    parts['[Content_Types].xml'] = strToU8(
      strFromU8(parts['[Content_Types].xml']).replace(
        '</Types>',
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'
      )
    );
  };
  const plain = await fixture(t, { rows: 4, patch: shared(false) });
  assert.equal((await plain.book.printSnapshot('0', 'Print', settings('A1:C4'))).cells[0].text, 'Item');
  const rich = await fixture(t, { rows: 4, patch: shared(true) });
  await assert.rejects(rich.book.printSnapshot('0', 'Print', settings('A1:C4')), { code: 'print-unsupported' });
  const styled = await fixture(t, {
    rows: 4,
    patch(parts) {
      parts['xl/styles.xml'] = strToU8(
        `<styleSheet xmlns="${ns}"><fonts count="1"><font><name val="Arial"/><sz val="11"/><scheme val="minor"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>`
      );
    },
  });
  await assert.rejects(styled.book.printSnapshot('0', 'Print', settings('A1:C4')), /PRINT_UNSUPPORTED_THEME/);
  assert.equal(styled.book.diagnosticStats().retainedStates, 1);
});
test('clearing print area/repeats keeps page settings and permits later reference edits; readiness flags survive settings edits', async (t) => {
  const f = await fixture(t, { rows: 4 }),
    made = await f.book.apply('0', [command(settings('A1:C4', { repeatRows: { first: 1, last: 1 } }))]),
    state = made.view.storage.state,
    definition = await f.book.printDefinition(state, 'Print');
  // A print area is a defined name: it follows a sheet rename like any other.
  const followed = await f.book.apply(state, [{ operation: 'renameSheet', sheet: 'Print', newName: 'Moved' }]);
  assert.equal((await f.book.printDefinition(followed.view.storage.state, 'Moved')).settings.area, 'A1:C4');
  const clear = await f.book.apply(state, [
      command({ ...definition.settings, area: null, repeatRows: null, repeatColumns: null }, definition.revision),
    ]),
    cleared = await f.book.printDefinition(clear.view.storage.state, 'Print');
  assert.equal(cleared.settings.area, null);
  assert.equal(cleared.settings.paper, 'A4');
  const renamed = await f.book.apply(clear.view.storage.state, [
    { operation: 'renameSheet', sheet: 'Print', newName: 'Renamed' },
  ]);
  assert.ok(renamed.view.sheets.some((sheet) => sheet.name === 'Renamed'));
  assert.equal((await f.book.printDefinition(state, 'Print')).settings.area, 'A1:C4');
  const g = await fixture(t, {
      rows: 4,
      patch(parts) {
        parts['xl/worksheets/sheet1.xml'] = strToU8(
          strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
            '<c r="B2" t="inlineStr"><is><t>row-2</t></is></c>',
            '<c r="B2"><f>UNKNOWN(1)</f><v>999</v></c>'
          )
        );
      },
    }),
    changed = await g.book.apply('0', [command(settings('A1:C4'))]),
    saved = await g.book.export(changed.view.storage.state);
  assert.match(strFromU8(unzipSync(await fs.readFile(saved.path))['xl/workbook.xml']), /fullCalcOnLoad="1"/);
});
test('product report keeps a readable chart, separate PNG and complete accepted data across physical pages', async (t) => {
  const values = Array.from({ length: 34 }, () => Array(8).fill(''));
  values[0] = ['Month', 'Revenue', 'Costs', '', 'Project', 'Sample', '', ''];
  values[1] = ['Jan', 120, 100, '', 'Exact reference', '9007199254740993', '', ''];
  values[2] = ['Feb', 170, 140, '', 'Date check', 60, '', ''];
  values[3] = ['Mar', 160, 130, '', 'Prepared', '2026-09-30', '', ''];
  values[5][0] = 'Quarterly sales report';
  values[12][6] = 'Company mark (PNG)';
  values[31][0] = 'Report notes';
  values[32][0] = 'The chart uses the accepted Revenue and Costs values shown above.';
  values[33][0] = "The date check preserves Excel's 1900-02-29 serial; exact reference digits remain unchanged.";
  const f = await fixture(t, {
      rows: 34,
      columns: 8,
      widths: Array(8).fill(90),
      rowHeight: 18,
      merges: ['F2:H2', 'F3:H3', 'F4:H4', 'A6:H6', 'G13:H13', 'A32:H32', 'A33:H33', 'A34:H34'],
      patch(parts) {
        const rows = values
          .map(
            (row, r) =>
              `<row r="${r + 1}"${[5, 31].includes(r) ? ' ht="30" customHeight="1"' : ''}>` +
              row
                .map((value, c) => {
                  const ref = String.fromCharCode(65 + c) + (r + 1);
                  if (!value) return '';
                  const style = r === 2 && c === 5 ? ' s="1"' : [5, 31].includes(r) && c === 0 ? ' s="2"' : '';
                  return typeof value === 'number' || (r === 1 && c === 5)
                    ? `<c r="${ref}"${style}><v>${value}</v></c>`
                    : `<c r="${ref}"${style} t="inlineStr"><is><t>${xml(value)}</t></is></c>`;
                })
                .join('') +
              '</row>'
          )
          .join('');
        parts['xl/worksheets/sheet1.xml'] = strToU8(
          strFromU8(parts['xl/worksheets/sheet1.xml']).replace(
            /<sheetData>[\s\S]*?<\/sheetData>/,
            '<sheetData>' + rows + '</sheetData>'
          )
        );
        parts['xl/styles.xml'] = strToU8(
          `<styleSheet xmlns="${ns}"><fonts count="2"><font><name val="Arial"/><sz val="11"/></font><font><name val="Arial"/><sz val="16"/><b/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0"/></cellXfs></styleSheet>`
        );
        parts['xl/_rels/workbook.xml.rels'] = strToU8(
          strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
            '</Relationships>',
            '<Relationship Id="rReportStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
          )
        );
        parts['[Content_Types].xml'] = strToU8(
          strFromU8(parts['[Content_Types].xml']).replace(
            '</Types>',
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
          )
        );
      },
    }),
    chartId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa',
    imageId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
  let result = await f.book.apply('0', [
    {
      operation: 'createChart',
      sheet: 'Print',
      objectId: chartId,
      expectedObjectRevision: null,
      chart: { kind: 'column', range: 'A1:C4', title: 'Quarterly sales' },
      anchor: { cell: 'A8', widthEmu: 480 * 12700, heightEmu: 250 * 12700 },
    },
  ]);
  result = await f.book.apply(result.view.storage.state, [
    {
      operation: 'insertImage',
      sheet: 'Print',
      objectId: imageId,
      expectedObjectRevision: null,
      asset: { mime: 'image/png', base64: twoPixelPng().toString('base64') },
      anchor: { cell: 'G8', widthEmu: 144 * 12700, heightEmu: 72 * 12700 },
      altText: 'Company mark',
    },
  ]);
  result = await f.book.apply(result.view.storage.state, [
    command(settings('A1:H34', { paper: 'Letter', orientation: 'landscape', repeatRows: { first: 1, last: 4 } })),
  ]);
  const state = result.view.storage.state,
    snapshot = await f.book.printSnapshot(state, 'Print'),
    layout = composePrint(snapshot),
    saved = await f.book.export(state);
  assert.equal(layout.pages.length, 2);
  assert.equal(layout.scale, 1);
  assert.equal(snapshot.cells.find((cell) => cell.ref === 'F2').text, '9007199254740993');
  assert.equal(snapshot.cells.find((cell) => cell.ref === 'F3').text, '1900-02-29');
  assert.equal(layout.pages[0].objects.length, 2);
  assert.ok(layout.pages[0].objects.every((object) => !object.clipped));
  const [a, b] = layout.pages[0].objects;
  assert.ok(a.bbox[0] + a.bbox[2] <= b.bbox[0] || b.bbox[0] + b.bbox[2] <= a.bbox[0]);
  const html = printDocumentHtml(layout);
});

test('actual native six-section plain headers persist across three presets, immutable prior states and cold reopen', async (t) => {
  const f = await fixture(t, { rows: 4 });
  let state = '0';
  const presets = [
    {
      header: { left: 'R&D', center: 'Print', right: '😀' },
      footer: { left: '', center: '', right: '' },
      scaleWithDoc: true,
    },
    {
      header: { left: '', center: 'Print', right: '' },
      footer: { left: '', center: 'Confidential', right: '' },
      scaleWithDoc: true,
    },
    {
      header: { left: '', center: 'Draft', right: '' },
      footer: { left: '', center: '', right: 'Print &P literal' },
      scaleWithDoc: false,
    },
  ];
  for (const hf of presets) {
    const before = await f.book.printDefinition(state, 'Print'),
      made = await f.book.apply(state, [command(settings('A1:C4', { headerFooter: hf }), before.revision)]);
    state = made.view.storage.state;
    const definition = await f.book.printDefinition(state, 'Print');
    assert.deepEqual(definition.settings.headerFooter, hf);
    assert.equal(definition.canEdit, true);
    const html = printDocumentHtml(composePrint(await f.book.printSnapshot(state, 'Print')));
    assert.ok(html.includes('data-print-furniture="header-center"'));
    const saved = await f.book.export(state),
      parts = unzipSync(await fs.readFile(saved.path)),
      sheet = strFromU8(parts['xl/worksheets/sheet1.xml']);
    assert.match(sheet, /<headerFooter /);
    assert.ok(sheet.indexOf('<headerFooter') > sheet.indexOf('<pageSetup'));
    for (const [part, bytes] of Object.entries(f.exportParts))
      if (!['xl/workbook.xml', 'xl/worksheets/sheet1.xml'].includes(part))
        unchangedExportPart(parts[part], bytes, part);
    const copy = path.join(f.directory, 'header-' + state + '.xlsx');
    await fs.copyFile(saved.path, copy);
    const cold = new IndexedWorkbook(
      native,
      () => assert.fail('No compatibility writer'),
      (input) => calculateIndexedWorkbook(input, native)
    );
    await cold.open(copy);
    try {
      assert.deepEqual((await cold.printDefinition('0', 'Print')).settings.headerFooter, hf);
    } finally {
      await cold.close();
    }
  }
  assert.equal((await f.book.printDefinition('0', 'Print')).settings.headerFooter, undefined);
  assert.deepEqual(await fs.readFile(f.source), f.original);
});
test('native guard independently rejects mismatched/header-code/extended fragment proposals before overwriting output', async (t) => {
  const f = await fixture(t, { rows: 4 }),
    metadata = await fs.readFile(path.join(f.root(), 'metadata.xlsx')),
    hf = {
      header: { left: '', center: 'Print', right: '' },
      footer: { left: '', center: 'Confidential', right: '' },
      scaleWithDoc: true,
    },
    prepared = indexedPrintPlan(metadata, f.book.view('0'), command(settings('A1:C4', { headerFooter: hf }))),
    source = await f.book.export('0'),
    output = path.join(f.root(), 'package.xlsx'),
    sentinel = Buffer.from('retain me');
  const changes = [
    (plan) => (plan.settings.headerFooter.header.center = 'mismatch'),
    (plan) =>
      (plan.fragments[plan.sheetPart].headerFooter = plan.fragments[plan.sheetPart].headerFooter.replace(
        '&amp;CPrint',
        '&amp;C&amp;P'
      )),
    (plan) =>
      (plan.fragments[plan.sheetPart].headerFooter = plan.fragments[plan.sheetPart].headerFooter.replace(
        'differentFirst="0"',
        'differentFirst="1"'
      )),
    (plan) =>
      (plan.fragments[plan.sheetPart].headerFooter = plan.fragments[plan.sheetPart].headerFooter.replace(
        'scaleWithDoc="1"',
        'scaleWithDoc="0"'
      )),
    (plan) => delete plan.settings.headerFooter,
  ];
  for (const change of changes) {
    const plan = structuredClone(prepared.plan);
    change(plan);
    await fs.writeFile(output, sentinel);
    await fs.writeFile(path.join(f.root(), 'package-metadata.json'), JSON.stringify(plan));
    await assert.rejects(
      native({
        action: 'workbook',
        operation: 'package-edit',
        root: f.root(),
        version: 0,
        source: path.basename(source.path),
      }),
      /header|Header|print|Print|Unqualified/
    );
    assert.deepEqual(await fs.readFile(output), sentinel);
  }
});
