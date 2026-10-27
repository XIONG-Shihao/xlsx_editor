import test from 'node:test';
import assert from 'node:assert/strict';
import { qualifyPivotTarget, replacePivotOutput } from '../core/pivot-target.mjs';
const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const cell = (ref, value) => `<c r="${ref}"><v>${value}</v></c>`;
const worksheet = (rows, dimension = 'A1:B2') =>
  `<worksheet xmlns="${ns}"><dimension ref="${dimension}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>${rows}</sheetData><pivotTableParts count="1"/></worksheet>`;
const old = worksheet(
  `<row r="1">${cell('A1', 1)}${cell('B1', 2)}</row><row r="2">${cell('A2', 3)}${cell('B2', 4)}</row>`
);
const next = worksheet(
  `<row r="1">${cell('A1', 5)}${cell('B1', 6)}</row><row r="2">${cell('A2', 7)}${cell('B2', 8)}</row>`
);
const neighbor = '<c r="F6" t="inlineStr"><is><t xml:space="preserve">00123 =SUM(A1) &#38; &lt;kept&gt;</t></is></c>';
const decorated = old
  .replace('A1:B2', 'A1:F6')
  .replace(
    '</row></sheetData>',
    '<c r="D2"><f>B2*2</f><v>8</v></c></row><row r="6">' + neighbor + '</row></sheetData>'
  );
test('pivot replacement retains exact neighboring formula and escaped raw text', () => {
  const result = replacePivotOutput(decorated, old, 'A1:B2', next, 'A1:B2');
  assert.equal(
    result,
    decorated
      .replace('<v>1</v>', '<v>5</v>')
      .replace('<v>2</v>', '<v>6</v>')
      .replace('<v>3</v>', '<v>7</v>')
      .replace('<v>4</v>', '<v>8</v>')
  );
  assert(result.includes(neighbor));
  assert(result.includes('<c r="D2"><f>B2*2</f><v>8</v></c>'));
});
test('pivot shrink removes only released output; growth preserves empty row and refuses occupied neighbor', () => {
  const shrink = worksheet(`<row r="1">${cell('A1', 9)}${cell('B1', 10)}</row>`, 'A1:B1');
  const result = replacePivotOutput(decorated, old, 'A1:B2', shrink, 'A1:B1');
  assert(!result.includes('r="A2"'));
  assert(!result.includes('r="B2"'));
  assert(result.includes('r="D2"'));
  assert(result.includes(neighbor));
  const grow = worksheet(
    `<row r="1">${cell('A1', 5)}${cell('B1', 6)}</row><row r="2">${cell('A2', 7)}${cell('B2', 8)}</row><row r="3">${cell('A3', 9)}${cell('B3', 10)}</row>`,
    'A1:B3'
  );
  const empty = decorated.replace('<row r="6">', '<row r="3"/><row r="6">');
  const expanded = replacePivotOutput(empty, old, 'A1:B2', grow, 'A1:B3');
  assert(expanded.includes('<row r="3">' + cell('A3', 9) + cell('B3', 10) + '</row>'));
  assert(expanded.includes(neighbor));
  assert.throws(
    () =>
      replacePivotOutput(
        decorated.replace('<row r="6">', '<row r="3">' + cell('A3', 99) + '</row><row r="6">'),
        old,
        'A1:B2',
        grow,
        'A1:B3'
      ),
    /growth/
  );
});
test('unknown neighboring metadata, duplicate coordinates and special formula refuse', () => {
  for (const raw of [
    decorated.replace('r="F6"', 'r="F6" cm="1"'),
    decorated.replace('<f>B2*2</f>', '<f t="array" ref="D2">B2*2</f>'),
    decorated.replace('</row><row r="6">', cell('D2', 4) + '</row><row r="6">'),
    decorated.replace('A1:F6', 'A1:E6'),
  ])
    assert.throws(() => qualifyPivotTarget(raw, old, 'A1:B2'));
});
