/** Authored Values vectors. Expected destination types never come from the editor. */
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../../core/workbook-package.mjs';
export const valueSheet = 'Values';
export const valueNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const escape = (value) =>
  String(value).replace(
    /[&<>"'\t\r\n]/g,
    (c) =>
      ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&apos;',
        '\t': '&#9;',
        '\r': '&#13;',
        '\n': '&#10;',
      })[c]
  );
export const literalResults = [
  { kind: 'text', value: '00123' },
  { kind: 'text', value: '=SUM(A4)' },
  { kind: 'boolean', value: true, raw: '1' },
  { kind: 'number', value: 7.5, raw: '7.5' },
  { kind: 'text', value: '#DIV/0!' },
  { kind: 'error', value: '#N/A', raw: '#N/A' },
  { kind: 'text', value: '' },
  { kind: 'blank', value: null, raw: null },
];
export const formulaResults = [
  { kind: 'number', value: 8.5, raw: '8.5', formula: 'A4+1' },
  { kind: 'text', value: 'computed text', formula: '"computed text"' },
  { kind: 'boolean', value: true, raw: '1', formula: 'TRUE()' },
  { kind: 'error', value: '#DIV/0!', raw: '#DIV/0!', formula: '1/0' },
];
export const literalClipboard = '"00123"\n"=SUM(A4)"\n"true"\n"7.5"\n"#DIV/0!"\n"#N/A"\n""\n""';
export const formulaClipboard = '"=A4+1"\n"=""computed text"""\n"=TRUE()"\n"=1/0"';
export const valueCalculationTag = (mode) =>
  `<calcPr calcId="191029" calcMode="${mode}" calcOnSave="0" fullCalcOnLoad="0" forceFullCalc="0"/>`;
export function valueSourceCells({ mode = 'auto', boundary = null } = {}) {
  if (!['auto', 'manual'].includes(mode) || ![null, 'unknown-array', 'owned-spill'].includes(boundary))
    throw Error('Closed Values fixture options');
  const cells = {};
  for (let row = 1; row <= 20; row++)
    for (let col = 0; col < 8; col++)
      cells[String.fromCharCode(65 + col) + row] = {
        kind: 'blank',
        value: null,
        raw: null,
        style: col === 2 || col === 3 ? 1 : 0,
      };
  literalResults.forEach((entry, index) => {
    cells['A' + (index + 1)] = { ...entry, style: 0 };
  });
  formulaResults.forEach((entry, index) => {
    const old = [103, 'old computed text', false, '#N/A'][index];
    cells['B' + (index + 1)] = {
      ...entry,
      kind: 'formula',
      cacheKind: entry.kind,
      value: mode === 'manual' ? old : entry.value,
      raw:
        entry.kind === 'text'
          ? mode === 'manual'
            ? old
            : entry.value
          : entry.kind === 'boolean'
            ? mode === 'manual'
              ? '0'
              : '1'
            : mode === 'manual'
              ? index === 0
                ? '103.0'
                : '#N/A'
              : entry.raw,
      style: 0,
    };
  });
  cells.F1 = { kind: 'text', value: 'outside exact 雪 😀', style: 0 };
  cells.F2 = { kind: 'error', value: '#REF!', raw: '#REF!', style: 0 };
  cells.F3 = { kind: 'boolean', value: false, raw: '0', style: 0 };
  cells.F4 = { kind: 'number', value: 9007199254740992, raw: '9007199254740993', style: 0 };
  cells.H20 = { kind: 'text', value: 'outside owner/head setup', style: 0 };
  if (boundary === 'unknown-array') {
    cells.R1100 = {
      kind: 'formula',
      cacheKind: 'number',
      formula: 'UNKNOWN(1)',
      formulaAttrs: ' t="array" ref="R1100:R1101"',
      value: 913,
      raw: '913.0',
      style: 0,
    };
    cells.R1101 = { kind: 'number', value: 914, raw: '914.0', style: 0 };
  }
  if (boundary === 'owned-spill') {
    cells.A10 = { kind: 'formula', cacheKind: 'number', formula: 'SEQUENCE(2,2)', value: 1, raw: '1.0', style: 0 };
    for (const ref of ['B10', 'A11', 'B11']) cells[ref] = { kind: 'blank', value: null, raw: null, style: 0 };
  }
  return cells;
}
function cellXml(ref, c) {
  const type =
    c.kind === 'formula'
      ? c.cacheKind === 'text'
        ? 'str'
        : c.cacheKind === 'boolean'
          ? 'b'
          : c.cacheKind === 'error'
            ? 'e'
            : null
      : c.kind === 'text'
        ? 'inlineStr'
        : c.kind === 'boolean'
          ? 'b'
          : c.kind === 'error'
            ? 'e'
            : null;
  return `<c r="${ref}"${c.style ? ' s="' + c.style + '"' : ''}${type ? ' t="' + type + '"' : ''}>${c.formula ? `<f${c.formulaAttrs ?? ''}>${escape(c.formula)}</f>` : ''}${c.kind === 'text' ? `<is><t xml:space="preserve">${escape(c.value)}</t></is>` : c.kind === 'blank' ? '' : `<v>${escape(c.raw)}</v>`}</c>`;
}
export function valuePasteFixture({ mode = 'auto', readonly = false, boundary = null } = {}) {
  const cells = valueSourceCells({ mode, boundary }),
    parts = unzipSync(createWorkbook()),
    rows = new Map();
  for (const [ref, c] of Object.entries(cells)) {
    const row = +ref.match(/\d+$/)[0];
    if (!rows.has(row)) rows.set(row, []);
    rows.get(row).push([ref, c]);
  }
  const data = [...rows]
    .sort((a, b) => a[0] - b[0])
    .map(
      ([row, entries]) =>
        `<row r="${row}"${row === 8 ? ' ht="24" customHeight="1"' : ''}${row === 19 ? ' hidden="true"' : ''}>` +
        entries
          .sort(([a], [b]) => a.replace(/\d+$/, '').localeCompare(b.replace(/\d+$/, '')))
          .map(([ref, c]) => cellXml(ref, c))
          .join('') +
        '</row>'
    )
    .join('');
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="${valueNamespace}"><dimension ref="A1:${boundary === 'unknown-array' ? 'R1101' : 'H20'}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="18" width="14" customWidth="1"/></cols><sheetData>${data}</sheetData>${readonly ? '<sheetProtection sheet="1"/>' : ''}</worksheet>`
  );
  parts['xl/workbook.xml'] = strToU8(
    strFromU8(parts['xl/workbook.xml'])
      .replace('name="Sheet1"', `name="${valueSheet}"`)
      .replace(/<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/g, '')
      .replace('</workbook>', valueCalculationTag(mode) + '</workbook>')
  );
  parts['xl/styles.xml'] = strToU8(
    `<styleSheet xmlns="${valueNamespace}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><i/><sz val="12"/><color rgb="FF123456"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFABCDEF"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="2" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1"/></cellXfs></styleSheet>`
  );
  parts['custom/values-opaque.bin'] = new Uint8Array([0, 255, 19, 73, 4, 11]);
  return Buffer.from(zipSync(parts, { mtime: new Date(2000, 0, 1) }));
}
export function valueSavedExpectation({ literals = false, formulas = false, mode = 'auto' } = {}) {
  if (!['auto', 'manual'].includes(mode)) throw Error('Closed Values saved mode');
  const cells = {};
  if (literals)
    literalResults.forEach((c, index) => {
      cells['C' + (index + 1)] = { ...c, style: 1 };
    });
  if (formulas) {
    if (mode !== 'auto') throw Error('Pending formula values must refuse');
    formulaResults.forEach(({ formula, ...c }, index) => {
      cells['D' + (index + 1)] = { ...c, style: 1 };
    });
  }
  return {
    schema: 1,
    cells,
    ...(mode === 'manual'
      ? {
          calcPrAfter: `<calcPr calcId="191029" calcMode="manual" calcOnSave="0" fullCalcOnLoad="1" forceFullCalc="1" xmlns="${valueNamespace}"/>`,
        }
      : {}),
  };
}
