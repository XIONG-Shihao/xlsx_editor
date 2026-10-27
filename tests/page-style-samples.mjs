import assert from 'node:assert/strict';
import { test, after } from 'node:test';
import { mkdtemp, writeFile, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8, strFromU8, unzlibSync } from 'fflate';
import * as XLSX from 'xlsx';
import { DOMParser } from '@xmldom/xmldom';
import { createWorkbook, readWorkbook, readWorkbookMetadata, editWorkbook } from '../core/workbook-package.mjs';
import { openZipParts, zipPartSize, writeZipParts } from '../core/zip-parts.mjs';
import { readWorkbookTheme, spreadsheetThemeColors } from '../core/workbook-theme.mjs';
import {
  pageThemeSamples,
  pageThemeSample,
  pageThemeXml,
  worksheetBackgroundSamples,
  worksheetBackgroundSample,
} from '../core/page-style-samples.mjs';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { closeNativeWorkbookWorker } from '../core/native-runtime.mjs';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const themePart = 'xl/theme/workbook-sample.xml';
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex');
const roots = [];
after(async () => {
  await closeNativeWorkbookWorker();
  await Promise.all(roots.map((root) => rm(root, { recursive: true, force: true })));
});
function fixture(id = 'forest', mutate = () => {}) {
  const parts = unzipSync(createWorkbook());
  parts['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
      '</Relationships>',
      `<Relationship Id="workbookTheme" Type="${REL}/theme" Target="theme/workbook-sample.xml"/></Relationships>`
    )
  );
  parts['[Content_Types].xml'] = strToU8(
    strFromU8(parts['[Content_Types].xml']).replace(
      '</Types>',
      `<Override PartName="/${themePart}" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`
    )
  );
  parts[themePart] = strToU8(pageThemeXml(id));
  parts['custom/unrelated.bin'] = new Uint8Array([0, 1, 29, 255]);
  parts['xl/styles.xml'] = strToU8(
    `<styleSheet xmlns="${NS}"><fonts count="3"><font><name val="Old body"/><sz val="11"/><color theme="1"/><scheme val="minor"/></font><font><name val="Old heading"/><sz val="16"/><color theme="4"/><scheme val="major"/></font><font><name val="Courier New"/><sz val="12"/><color rgb="FF11AA44"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="5"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`
  );
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${NS}"><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews><sheetFormatPr defaultRowHeight="19"/><cols><col min="2" max="2" width="24" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Body</t></is></c><c r="B1" s="1" t="inlineStr"><is><t>Heading</t></is></c><c r="C1" s="2" t="inlineStr"><is><t>Literal</t></is></c></row><row r="4" hidden="1"><c r="A4"><v>77</v></c></row></sheetData></worksheet>`
  );
  mutate(parts);
  return { bytes: zipSync(parts), parts };
}

test('three real themes resolve major/minor fonts, colors and fills; literal styles stay literal', () => {
  assert.equal(pageThemeSamples.length, 3);
  for (const sample of pageThemeSamples) {
    const { bytes } = fixture(sample.id),
      view = readWorkbook(bytes),
      sheet = view.sheets[0],
      cells = sheet.cells;
    assert.equal(cells.A1.fontFamily, sample.minorFont);
    assert.equal(cells.A1.fontColor, sample.colors.dk1);
    assert.equal(cells.B1.fontFamily, sample.majorFont);
    assert.equal(cells.B1.fontColor, sample.colors.accent1);
    assert.equal(cells.B1.fillColor, sample.colors.accent2);
    assert.equal(cells.C1.fontFamily, 'Courier New');
    assert.equal(cells.C1.fontColor, '11AA44');
    assert.equal(sheet.normalFont.fontFamily, sample.minorFont);
    assert.equal(sheet.normalFont.fontSize, 11);
    assert.deepEqual(sheet.hiddenRows, [3]);
    assert.equal(sheet.sourceView.gridlines, false);
    assert.equal(sheet.sheetFormat.defaultRowHeight, 19);
    assert(!view.limitations.some((message) => /theme|tint/i.test(message)));
    const metadata = readWorkbookMetadata(bytes);
    assert.equal(metadata.styles[1].fontFamily, sample.majorFont);
    assert.equal(metadata.styles[1].fillColor, sample.colors.accent2);
  }
  assert.throws(() => pageThemeSample('Office'), { code: 'internal-unknown-reference' });
  assert.throws(() => pageThemeXml('<bad>'), { code: 'internal-unknown-reference' });
  assert(Object.isFrozen(pageThemeSamples));
  assert(Object.isFrozen(pageThemeSample('forest').colors));
});

test('independent SheetJS theme interpretation agrees on all twelve indices and cell fill', () => {
  for (const sample of pageThemeSamples) {
    const { bytes } = fixture(sample.id),
      independent = XLSX.read(bytes, { type: 'buffer', cellStyles: true });
    const colors = independent.Themes.themeElements.clrScheme;
    assert.equal(colors.length, 12);
    for (let i = 0; i < 12; i++) assert.equal(colors[i].rgb, sample.colors[spreadsheetThemeColors[i]]);
    assert.equal(independent.Sheets.Sheet1.B1.s.fgColor.rgb, sample.colors.accent2);
  }
});

test('sample XML has the real color/font/format schemes and three style-matrix entries', () => {
  const D = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  for (const sample of pageThemeSamples) {
    const doc = new DOMParser().parseFromString(pageThemeXml(sample.id), 'application/xml');
    for (const name of ['themeElements', 'clrScheme', 'fontScheme', 'fmtScheme'])
      assert.equal(doc.getElementsByTagNameNS(D, name).length, 1);
    for (const name of ['fillStyleLst', 'lnStyleLst', 'effectStyleLst', 'bgFillStyleLst'])
      assert.equal(
        Array.from(doc.getElementsByTagNameNS(D, name)[0].childNodes).filter((n) => n.nodeType === 1).length,
        3
      );
  }
});

test('two PNG backgrounds independently decode to their actual pixel patterns and stable identities', () => {
  assert.equal(worksheetBackgroundSamples.length, 2);
  for (const sample of worksheetBackgroundSamples) {
    const bytes = Buffer.from(sample.base64, 'base64');
    assert.equal(sha(bytes), sample.sha256);
    assert.equal(bytes.subarray(0, 8).toString('hex'), '89504e470d0a1a0a');
    assert.equal(bytes.readUInt32BE(16), 32);
    assert.equal(bytes.readUInt32BE(20), 32);
    assert.deepEqual([...bytes.subarray(24, 29)], [8, 2, 0, 0, 0]);
    let cursor = 8,
      idat = [];
    while (cursor < bytes.length) {
      const length = bytes.readUInt32BE(cursor),
        tag = bytes.toString('ascii', cursor + 4, cursor + 8);
      if (tag === 'IDAT') idat.push(bytes.subarray(cursor + 8, cursor + 8 + length));
      cursor += length + 12;
    }
    const pixels = unzlibSync(Buffer.concat(idat));
    assert.equal(pixels.length, 32 * 97);
    for (let y = 0; y < 32; y++) {
      assert.equal(pixels[y * 97], 0);
      for (let x = 0; x < 32; x++) {
        const expected =
          sample.id === 'graph-paper'
            ? x % 8 === 0 || y % 8 === 0
              ? [224, 234, 241]
              : [250, 252, 254]
            : x % 8 === 3 && y % 8 === 3
              ? [226, 224, 235]
              : [252, 250, 254];
        assert.deepEqual([...pixels.subarray(y * 97 + 1 + x * 3, y * 97 + 4 + x * 3)], expected);
      }
    }
    assert(Object.isFrozen(sample));
  }
  assert.throws(() => worksheetBackgroundSample('custom'), { code: 'internal-unknown-reference' });
});

test('theme discovery is relationship-bound, namespace-aware and reads sysClr lastClr', () => {
  const { bytes } = fixture('forest', (parts) => {
    parts[themePart] = strToU8(
      pageThemeXml('forest')
        .replaceAll('a:', 'd:')
        .replace('xmlns:a=', 'xmlns:d=')
        .replace('<d:srgbClr val="182A20"/>', '<d:sysClr val="windowText" lastClr="182A20"/>')
    );
    parts['xl/theme/theme1.xml'] = strToU8(pageThemeXml('plum'));
  });
  const view = readWorkbook(bytes);
  assert.equal(view.sheets[0].cells.A1.fontColor, '182A20');
  assert.equal(view.sheets[0].cells.A1.fontFamily, 'Verdana');
});

test('unsupported tint and malformed/external/duplicate theme identities stay unresolved and preserve source', () => {
  const failures = [
    (parts) =>
      (parts[themePart] = strToU8(
        pageThemeXml('forest').replace('<a:accent1>', '<a:accent1><a:srgbClr val="123456"/>')
      )),
    (parts) => (parts[themePart] = strToU8(pageThemeXml('forest').replace('drawingml/2006/main', 'foreign/theme'))),
    (parts) => (parts[themePart] = strToU8('<!DOCTYPE theme [<!ENTITY evil "bad">]>' + pageThemeXml('forest'))),
    (parts) =>
      (parts['xl/_rels/workbook.xml.rels'] = strToU8(
        strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
          'Target="theme/workbook-sample.xml"',
          'Target="https://example.invalid/theme.xml" TargetMode="External"'
        )
      )),
    (parts) =>
      (parts['xl/_rels/workbook.xml.rels'] = strToU8(
        strFromU8(parts['xl/_rels/workbook.xml.rels']).replace(
          '</Relationships>',
          `<Relationship Id="secondTheme" Type="${REL}/theme" Target="theme/workbook-sample.xml"/></Relationships>`
        )
      )),
    (parts) => delete parts[themePart],
  ];
  for (const [index, mutate] of failures.entries()) {
    const { bytes } = fixture('forest', mutate),
      before = sha(bytes),
      view = readWorkbook(bytes);
    assert(view.limitations.some((message) => /theme/i.test(message)));
    assert.equal(view.sheets[0].cells.A1.fontFamily, index === 0 ? 'Verdana' : 'Old body');
    if (index === 0) assert.equal(view.sheets[0].cells.B1.fontColor, undefined);
    assert.equal(sha(bytes), before);
  }
  const { bytes } = fixture(
    'forest',
    (parts) =>
      (parts['xl/styles.xml'] = strToU8(strFromU8(parts['xl/styles.xml']).replace('theme="4"', 'theme="4" tint="0.3"')))
  );
  const view = readWorkbook(bytes);
  assert.equal(view.sheets[0].cells.B1.fontColor, undefined);
  assert(view.limitations.some((message) => /tint/i.test(message)));
});

test('oversized theme is refused before payload inflation; ZIP size respects replacement/removal', () => {
  const { bytes } = fixture('forest', (parts) => (parts[themePart] = new Uint8Array(64 * 1024 + 1)));
  const lazy = openZipParts(bytes);
  assert.equal(zipPartSize(lazy, themePart), 64 * 1024 + 1);
  // Corrupt that member's compressed stream. An attempted payload read must
  // fail inflation, while a pre-inflation size refusal still works.
  const broken = Buffer.from(bytes);
  let cursor = broken.readUInt32LE(broken.length - 6),
    changed = false;
  while (cursor < broken.length - 22 && broken.readUInt32LE(cursor) === 0x02014b50) {
    const length = broken.readUInt16LE(cursor + 28),
      name = broken.toString('utf8', cursor + 46, cursor + 46 + length);
    if (name === themePart) {
      const local = broken.readUInt32LE(cursor + 42),
        start = local + 30 + broken.readUInt16LE(local + 26) + broken.readUInt16LE(local + 28);
      broken[start] = 0xff;
      changed = true;
      break;
    }
    cursor += 46 + length + broken.readUInt16LE(cursor + 30) + broken.readUInt16LE(cursor + 32);
  }
  assert(changed);
  assert.throws(() => openZipParts(broken)[themePart]);
  const result = readWorkbookTheme(openZipParts(broken), [{ type: REL + '/theme', path: themePart }]);
  assert.match(result.issues[0], /exceeds 64 KiB/);
  const view = readWorkbook(bytes);
  assert(view.limitations.some((message) => /exceeds 64 KiB/.test(message)));
  lazy[themePart] = strToU8(pageThemeXml('neutral'));
  assert.equal(zipPartSize(lazy, themePart), lazy[themePart].length);
  delete lazy[themePart];
  assert.equal(zipPartSize(lazy, themePart), undefined);
  assert(!Object.keys(unzipSync(writeZipParts(lazy))).includes(themePart));
});

test('explicit font choice removes the inherited theme scheme without mutating neighboring styles', () => {
  const { bytes, parts } = fixture(),
    before = readWorkbook(bytes),
    edited = editWorkbook(bytes, [{ sheet: 'Sheet1', cell: 'A1', fontFamily: 'Courier New' }], before.revision),
    view = readWorkbook(edited);
  assert.equal(view.sheets[0].cells.A1.fontFamily, 'Courier New');
  assert.equal(view.sheets[0].cells.B1.fontFamily, 'Georgia');
  assert.equal(view.sheets[0].normalFont.fontFamily, 'Verdana');
  const saved = unzipSync(edited);
  assert.deepEqual(saved[themePart], parts[themePart]);
  assert.deepEqual(saved['custom/unrelated.bin'], parts['custom/unrelated.bin']);
});

test('actual native owners resolve all three themes; font edit Undo/Redo, Save and cold reopen preserve theme and unrelated parts', async (t) => {
  const root = await mkdtemp(path.join(os.tmpdir(), 'p3-themes-'));
  roots.push(root);
  for (const sample of pageThemeSamples) {
    const { bytes, parts } = fixture(sample.id),
      source = path.join(root, sample.id + '.xlsx'),
      saved = path.join(root, sample.id + '-saved.xlsx');
    await writeFile(source, bytes);
    const recoveryDirectory = path.join(root, 'recovery-' + sample.id),
      owner = new SpreadsheetDocument({ recoveryDirectory });
    let reopened;
    try {
      const status = await owner.open(source),
        initial = owner.status().state;
      const read = async (current) =>
        (await current.request({ operation: 'read', state: current.status().state, sheet: 'Sheet1', range: 'A1:C1' }))
          .cells;
      const cells = await read(owner);
      assert.equal(cells.A1.fontFamily, sample.minorFont);
      assert.equal(cells.B1.fontFamily, sample.majorFont);
      assert.equal(cells.B1.fillColor, sample.colors.accent2);
      assert.equal(owner.status().state, initial);
      assert.equal(owner.status().canUndo, false);
      await owner.request({
        operation: 'apply',
        state: initial,
        commands: [{ sheet: 'Sheet1', cell: 'A1', fontFamily: 'Courier New' }],
      });
      assert.equal((await read(owner)).A1.fontFamily, 'Courier New');
      await owner.request({ operation: 'undo' });
      assert.equal((await read(owner)).A1.fontFamily, sample.minorFont);
      await owner.request({ operation: 'redo' });
      assert.equal((await read(owner)).A1.fontFamily, 'Courier New');
      await owner.request({ operation: 'copy', destination: saved, expectedRevision: null });
      const savedParts = unzipSync(await readFile(saved));
      assert.deepEqual(savedParts[themePart], parts[themePart]);
      assert.deepEqual(savedParts['custom/unrelated.bin'], parts['custom/unrelated.bin']);
      assert.equal(sha(await readFile(source)), sha(bytes));
      await owner.close();
      reopened = new SpreadsheetDocument({ recoveryDirectory: path.join(root, 'cold-' + sample.id) });
      await reopened.open(saved);
      const cold = await read(reopened);
      assert.equal(cold.A1.fontFamily, 'Courier New');
      assert.equal(cold.B1.fontFamily, sample.majorFont);
      assert.equal(cold.B1.fillColor, sample.colors.accent2);
    } finally {
      await owner.close();
      await reopened?.close();
    }
  }
});
