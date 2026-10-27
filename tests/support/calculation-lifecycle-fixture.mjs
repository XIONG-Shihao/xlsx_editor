/** Independent authored caches and dependencies; no production evaluator supplies expectations. */
import { unzipSync, zipSync, strToU8, strFromU8 } from 'fflate';
import assert from 'node:assert/strict';
import { DOMParser } from '@xmldom/xmldom';
import { calculationTraceFixture } from './calculation-trace-fixture.mjs';

const esc = (value) => String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('"', '&quot;');
export const lifecycleSheets = Object.freeze({
  Inputs: 'xl/worksheets/sheet1.xml',
  Data: 'xl/worksheets/sheet2.xml',
  Report: 'xl/worksheets/sheet3.xml',
  Downstream: 'xl/worksheets/sheet4.xml',
});
export const lifecycleFormulas = Object.freeze({
  'Inputs!B1': 'A1+1',
  'Inputs!B2': 'SUM(A1:A2)',
  'Inputs!D1': 'Data!B2*2',
  'Report!A1': 'Inputs!B1*2',
  'Report!A2': 'Inputs!B2+Data!B2',
  'Report!B1': 'SUM(Rates[Amount])',
  'Report!B2': 'RateInput*3',
  'Downstream!A1': 'Report!A1+1',
});
const sourceNumbers = {
  'Inputs!B1': 3,
  'Inputs!B2': 5,
  'Inputs!D1': 6,
  'Report!A1': 6,
  'Report!A2': 8,
  'Report!B1': 15,
  'Report!B2': 6,
  'Downstream!A1': 7,
};
const staleNumbers = {
  'Inputs!B1': 103,
  'Inputs!B2': 105,
  'Inputs!D1': 106,
  'Report!A1': 106,
  'Report!A2': 108,
  'Report!B1': 115,
  'Report!B2': 106,
  'Downstream!A1': 107,
};
const afterNumbers = {
  'Inputs!B1': 6,
  'Inputs!B2': 8,
  'Inputs!D1': 18,
  'Report!A1': 12,
  'Report!A2': 17,
  'Report!B1': 21,
  'Report!B2': 15,
  'Downstream!A1': 13,
};
const cache = (formula, type, value) => ({ formula, cache: { type, value } });
function numbers(values) {
  return Object.fromEntries(
    Object.entries(values).map(([key, value]) => [key, cache(lifecycleFormulas[key], 'number', value)])
  );
}
export const lifecycleTypedFormulas = Object.freeze({
  'Inputs!E3': 'IF(A1=2,"source text","#N/A")',
  'Inputs!F3': 'A1=2',
  'Inputs!G3': 'A1/0',
  'Inputs!I3': '0.1+0.2',
});
export const lifecycleExpected = Object.freeze({
  source: numbers(sourceNumbers),
  staleSource: numbers(staleNumbers),
  afterAll: numbers(afterNumbers),
  edits: [
    { sheet: 'Inputs', cell: 'A1', value: 5 },
    { sheet: 'Data', cell: 'B2', value: 9 },
  ],
  reportClosure: ['Inputs!B1', 'Inputs!B2', 'Report!A1', 'Report!A2', 'Report!B1', 'Report!B2'],
  pendingAfterReport: ['Inputs!D1', 'Downstream!A1'],
  typedSource: {
    'Inputs!E3': cache(lifecycleTypedFormulas['Inputs!E3'], 'text', 'stale text'),
    'Inputs!F3': cache(lifecycleTypedFormulas['Inputs!F3'], 'boolean', false),
    'Inputs!G3': cache(lifecycleTypedFormulas['Inputs!G3'], 'error', '#N/A'),
    'Inputs!I3': cache(lifecycleTypedFormulas['Inputs!I3'], 'number', 0.25),
  },
  typedAfterAll: {
    'Inputs!E3': cache(lifecycleTypedFormulas['Inputs!E3'], 'text', '#N/A'),
    'Inputs!F3': cache(lifecycleTypedFormulas['Inputs!F3'], 'boolean', false),
    'Inputs!G3': cache(lifecycleTypedFormulas['Inputs!G3'], 'error', '#DIV/0!'),
    'Inputs!I3': cache(lifecycleTypedFormulas['Inputs!I3'], 'number', 0.30000000000000004),
  },
  changedFormula: {
    cell: 'Inputs!B1',
    formula: 'A1+10',
    pendingCache: { formula: 'A1+10', cache: { type: 'absent' } },
    afterAll: {
      ...numbers(afterNumbers),
      'Inputs!B1': cache('A1+10', 'number', 15),
      'Report!A1': cache(lifecycleFormulas['Report!A1'], 'number', 30),
      'Downstream!A1': cache(lifecycleFormulas['Downstream!A1'], 'number', 31),
    },
  },
  scope:
    'Authored local selected-sheet plus required formula ancestors; not captured Microsoft Excel worksheet-only behavior',
});

/** A Manual source deliberately containing stale, type-distinct saved caches. */
export function calculationLifecycleFixture({
  mode = 'manual',
  recalculateBeforeSave = true,
  stale = true,
  typed = false,
  markers = false,
  readonly = false,
} = {}) {
  const parts = unzipSync(
    calculationTraceFixture({ mode, recalcBeforeSave: recalculateBeforeSave, markers, readonly })
  );
  if (stale)
    for (const [key, value] of Object.entries(staleNumbers)) {
      const [sheet, ref] = key.split('!'),
        part = lifecycleSheets[sheet],
        xml = strFromU8(parts[part]);
      const pattern = new RegExp(`(<c r="${ref}"[^>]*><f>[^<]*</f><v>)[^<]*(</v></c>)`);
      if (!pattern.test(xml)) throw Error('Authored cache missing ' + key);
      parts[part] = strToU8(xml.replace(pattern, `$1${value}$2`));
    }
  if (typed) {
    const formula = (ref, text, type, value) =>
      `<c r="${ref}"${type === 'number' ? '' : ` t="${{ text: 'str', boolean: 'b', error: 'e' }[type]}"`}><f>${esc(text)}</f><v>${esc(type === 'boolean' ? (value ? '1' : '0') : value)}</v></c>`;
    const cells = Object.entries(lifecycleExpected.typedSource)
      .map(([key, value]) => formula(key.split('!')[1], value.formula, value.cache.type, value.cache.value))
      .join('');
    parts[lifecycleSheets.Inputs] = strToU8(
      strFromU8(parts[lifecycleSheets.Inputs]).replace(
        '</sheetData>',
        `<row r="3">${cells.replace('<c r="I3"', '<c r="H3" t="inlineStr"><is><t>#N/A</t></is></c><c r="I3"')}</row></sheetData>`
      )
    );
  }
  return Buffer.from(zipSync(parts, { mtime: new Date(2000, 0, 1) }));
}

export function lifecycleSaveExpected({
  mode = 'manual',
  recalculateBeforeSave = false,
  caches = lifecycleExpected.staleSource,
  literals = {},
  typed = false,
  typedCaches = lifecycleExpected.typedSource,
  formulaReplacement = false,
  fullCalcOnLoad = '0',
  forceFullCalc = '0',
} = {}) {
  const cells = { ...structuredClone(caches), ...(typed ? structuredClone(typedCaches) : {}) };
  if (formulaReplacement) cells['Inputs!B1'] = structuredClone(lifecycleExpected.changedFormula.pendingCache);
  return {
    schema: 'workbook.xlsx.calculation-lifecycle-expected.v1',
    sheetParts: lifecycleSheets,
    calcPr: { calcMode: mode, calcOnSave: recalculateBeforeSave ? '1' : '0', fullCalcOnLoad, forceFullCalc },
    cells,
    literals,
    authority: 'Saved cache bytes alone never establish current/pending owner status',
  };
}

/** Optional, independently sourced dimension expectations for this fixture
 * family. Read ORIGINAL authored XML only; never native/Saved results. All
 * physical cells in the qualified profile persist and retain their coordinates.
 * A source tableParts profile keeps its exact original dimension declaration.
 * lifecycleSaveExpected() stays unchanged, so callers must opt in explicitly.
 */
export function lifecycleDimensionExpectations(originalBytes) {
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    rels = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    parts = unzipSync(originalBytes),
    result = {};
  const children = (node, name) =>
    Array.from(node?.childNodes ?? []).filter(
      (child) => child.nodeType === 1 && child.namespaceURI === ns && child.localName === name
    );
  const attributes = (node) =>
    Array.from(node.attributes ?? [])
      .filter((attribute) => attribute.namespaceURI !== 'http://www.w3.org/2000/xmlns/')
      .map((attribute) => attribute.name);
  const parse = (bytes) => {
    const source = strFromU8(bytes);
    assert(!/<!DOCTYPE|<!ENTITY/i.test(source), 'Dimension profile cannot use XML entities');
    return new DOMParser({
      errorHandler: {
        warning() {},
        error(message) {
          throw Error(message);
        },
        fatalError(message) {
          throw Error(message);
        },
      },
    }).parseFromString(source, 'application/xml').documentElement;
  };
  const coordinate = (ref) => {
    const match = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(ref);
    assert(match, 'Noncanonical source coordinate');
    const column = [...match[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0),
      row = Number(match[2]);
    assert(column <= 16384 && row <= 1048576, 'Source coordinate outside worksheet');
    return { row, column };
  };
  const reference = (ref) => {
    assert(typeof ref === 'string');
    const pair = ref.split(':');
    assert(pair.length <= 2);
    const first = coordinate(pair[0]),
      last = coordinate(pair[1] ?? pair[0]);
    assert(first.row <= last.row && first.column <= last.column);
    return ref;
  };
  const columnName = (column) => {
    let value = '';
    for (let n = column; n; n = Math.floor((n - 1) / 26)) value = String.fromCharCode(65 + ((n - 1) % 26)) + value;
    return value;
  };
  const physical = Object.keys(parts)
    .filter((part) => /^xl\/worksheets\/[^/]+\.xml$/.test(part))
    .sort();
  assert.deepEqual(
    physical,
    Object.values(lifecycleSheets).slice().sort(),
    'No undeclared worksheet in the dimension profile'
  );
  for (const [sheet, part] of Object.entries(lifecycleSheets)) {
    assert(parts[part]);
    const root = parse(parts[part]);
    assert.equal(root.namespaceURI, ns);
    assert.equal(root.localName, 'worksheet');
    assert.equal(attributes(root).length, 0, 'Unknown worksheet root attributes are outside this profile');
    for (const node of Array.from(root.childNodes).filter((child) => child.nodeType === 1))
      assert(
        node.namespaceURI === ns &&
          ['dimension', 'sheetViews', 'sheetFormatPr', 'sheetData', 'tableParts'].includes(node.localName),
        'Axes, merges or unknown worksheet graph are outside this dimension profile'
      );
    const format = children(root, 'sheetFormatPr');
    assert.equal(format.length, 1);
    assert.deepEqual(attributes(format[0]), ['defaultRowHeight']);
    assert.equal(format[0].getAttribute('defaultRowHeight'), '15');
    assert.equal(format[0].childNodes.length, 0, 'Only the known fixture defaults qualify');
    const dimensions = children(root, 'dimension');
    assert.equal(dimensions.length, 1);
    assert.deepEqual(attributes(dimensions[0]), ['ref']);
    assert.equal(dimensions[0].childNodes.length, 0);
    const source = reference(dimensions[0].getAttribute('ref'));
    const data = children(root, 'sheetData');
    assert.equal(data.length, 1);
    const positions = [],
      seen = new Set(),
      seenRows = new Set();
    for (const row of Array.from(data[0].childNodes).filter((child) => child.nodeType === 1)) {
      assert.equal(row.namespaceURI, ns);
      assert.equal(row.localName, 'row');
      assert.deepEqual(attributes(row), ['r'], 'Authored row formatting is outside this dimension profile');
      const number = row.getAttribute('r');
      assert(/^[1-9]\d{0,6}$/.test(number) && Number(number) <= 1048576 && !seenRows.has(number));
      seenRows.add(number);
      for (const cell of Array.from(row.childNodes).filter((child) => child.nodeType === 1)) {
        assert.equal(cell.namespaceURI, ns);
        assert.equal(cell.localName, 'c');
        assert(attributes(cell).every((key) => ['r', 't', 's'].includes(key)));
        assert(
          !cell.hasAttribute('s') || cell.getAttribute('s') === '0',
          'Nondefault source styling is outside this dimension profile'
        );
        const ref = cell.getAttribute('r'),
          position = coordinate(ref);
        assert.equal(String(position.row), number);
        assert(!seen.has(ref));
        seen.add(ref);
        const formulas = children(cell, 'f'),
          values = children(cell, 'v'),
          inline = children(cell, 'is'),
          kind = cell.getAttribute('t') || 'n';
        assert(formulas.length <= 1 && values.length <= 1 && inline.length <= 1);
        assert(['n', 'str', 'inlineStr', 'b', 'e'].includes(kind), 'Unsupported scalar encoding in dimension profile');
        assert(
          Array.from(cell.childNodes)
            .filter((child) => child.nodeType === 1)
            .every((child) => child.namespaceURI === ns && ['f', 'v', 'is'].includes(child.localName))
        );
        if (formulas.length) {
          assert(
            formulas[0].textContent &&
              !attributes(formulas[0]).length &&
              !Array.from(formulas[0].childNodes).some((child) => child.nodeType === 1),
            'Only ordinary authored formulas qualify'
          );
        } else if (kind === 'inlineStr') {
          assert.equal(inline.length, 1);
          assert.equal(values.length, 0);
          assert.equal(children(inline[0], 't').length, 1);
        } else {
          assert.equal(values.length, 1, 'Removable blank tags cannot borrow persistent-coordinate expectations');
          if (kind === 'n') assert(values[0].textContent && Number.isFinite(Number(values[0].textContent)));
          if (kind === 'b') assert(['0', '1'].includes(values[0].textContent));
          if (kind === 'e') assert(values[0].textContent);
        }
        positions.push(position);
      }
    }
    assert(positions.length, 'Empty sheets require a separate qualified dimension profile');
    const tables = children(root, 'tableParts');
    assert(tables.length <= 1);
    let saved;
    if (tables.length) {
      assert.equal(sheet, 'Data', 'Only the exact authored Data table profile is preserved');
      assert.deepEqual(attributes(tables[0]), ['count']);
      assert.equal(tables[0].getAttribute('count'), '1');
      const table = children(tables[0], 'tablePart');
      assert.equal(table.length, 1);
      assert.equal(table[0].getAttributeNS(rels, 'id'), 'table1');
      assert.deepEqual(attributes(table[0]), ['r:id']);
      assert(parts['xl/tables/table1.xml'] && parts['xl/worksheets/_rels/sheet2.xml.rels']);
      saved = source;
    } else {
      const minRow = Math.min(...positions.map((p) => p.row)),
        maxRow = Math.max(...positions.map((p) => p.row)),
        minColumn = Math.min(...positions.map((p) => p.column)),
        maxColumn = Math.max(...positions.map((p) => p.column)),
        first = columnName(minColumn) + minRow,
        last = columnName(maxColumn) + maxRow;
      saved = first === last ? first : first + ':' + last;
    }
    result[sheet] = { source, saved };
  }
  return result;
}
