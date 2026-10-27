/** A new values-only workbook, independently re-read before publication. */
import { PASSIVE_FORMAT_LIMITS as LIMITS } from './passive-container.mjs';
import { normalizeDataOnlyXlsb } from './xlsb-values-normalize.mjs';
import { refusal } from './refusal-codes.mjs';
export const DATA_ONLY_LIMITS = Object.freeze({
  formats: ['xls', 'xlsb', 'xlsx'],
  sourcePositions: 1000000,
  estimatedBytes: 32 * 1024 * 1024,
  outputBytes: 64 * 1024 * 1024,
  text: 32767,
});
export const DATA_ONLY_OMISSIONS = Object.freeze([
  'Formulas become their stored values; no calculation is run.',
  'VBA, drawings, charts, images, comments, merges, formatting, names, links, tables, filters, rules and print settings are omitted.',
  'All worksheets are exported as visible; numeric date/time serials remain numbers without their original display formats.',
  'This is a new workbook. The original file and exact same-format Copy stay unchanged.',
]);
const fail = (message) => {
  throw refusal('package-bad-input', 'Data-only export refused: ' + message);
};
const errors = new Map([
  ['#NULL!', 0],
  ['#DIV/0!', 7],
  ['#VALUE!', 15],
  ['#REF!', 23],
  ['#NAME?', 29],
  ['#NUM!', 36],
  ['#N/A', 42],
  ['#GETTING_DATA', 43],
]);
function scalar(cell) {
  const formula = cell.formula !== undefined || cell.formulaType !== undefined;
  if (formula && cell.calculation?.status !== 'stored') fail('a formula has no qualified stored value');
  const value = formula ? cell.calculation.value : cell.value;
  if (cell.dataType === 'error') {
    const error = formula ? cell.calculation?.error : cell.error;
    if (typeof value !== 'string' || value !== error || !errors.has(value))
      fail('an error or malformed cell is not representable');
    return { t: 'e', v: errors.get(value) };
  }
  if (cell.error) fail('a malformed cell is protected');
  if (value === null) return null;
  if (typeof value === 'string') {
    if (value.length > DATA_ONLY_LIMITS.text || value.includes('\0')) fail('text exceeds the writer character budget');
    return { t: 's', v: value };
  }
  if (typeof value === 'boolean') return { t: 'b', v: value };
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) fail('a numeric value is nonfinite');
    if (
      typeof cell.rawValue === 'string' &&
      /^[+-]?\d+$/.test(cell.rawValue) &&
      Math.abs(value) > Number.MAX_SAFE_INTEGER
    )
      fail('an original decimal integer exceeds exact scalar precision');
    return { t: 'n', v: value };
  }
  fail('an unsupported stored scalar is protected');
}
/** XLSX is injected so this pure module cannot choose or upgrade a parser. */
export function dataOnlyWorkbook(XLSX, model, format) {
  if (XLSX?.version !== '0.20.3' || !DATA_ONLY_LIMITS.formats.includes(format))
    fail('writer identity or target format disagrees');
  const output = { SheetNames: [], Sheets: Object.create(null) },
    expected = new Map();
  let positions = 0,
    estimated = 0,
    formulaValues = 0,
    values = 0;
  for (const sheet of model.sheets) {
    const target = Object.create(null),
      entries = model.records.get(sheet.name);
    if (!Array.isArray(entries)) fail('source worksheet identity differs');
    let lastRow = 0,
      lastCol = 0;
    const comparison = new Map();
    for (const entry of entries) {
      if (++positions > DATA_ONLY_LIMITS.sourcePositions) fail('source position budget exceeded');
      const value = scalar(entry.cell);
      if (!value) continue;
      if (entry.row >= (format === 'xls' ? 65536 : 1048576) || entry.col >= (format === 'xls' ? 256 : 16384))
        fail('stored values exceed the target worksheet geometry');
      estimated += 128 + (typeof value.v === 'string' ? new TextEncoder().encode(value.v).length * 2 : 16);
      if (estimated > DATA_ONLY_LIMITS.estimatedBytes) fail('estimated output budget exceeded');
      target[entry.ref] = value;
      comparison.set(entry.ref, value);
      lastRow = Math.max(lastRow, entry.row);
      lastCol = Math.max(lastCol, entry.col);
      values++;
      if (entry.cell.formula !== undefined || entry.cell.formulaType !== undefined) formulaValues++;
    }
    target['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lastRow, c: lastCol } });
    output.SheetNames.push(sheet.name);
    output.Sheets[sheet.name] = target;
    expected.set(sheet.name, comparison);
  }
  if (!output.SheetNames.length || output.SheetNames.length > LIMITS.sheets) fail('worksheet count exceeds its budget');
  const written = XLSX.write(output, { type: 'buffer', bookType: format, compression: true, bookSST: true }),
    bytes = format === 'xlsb' ? normalizeDataOnlyXlsb(written) : written;
  if (bytes.byteLength > DATA_ONLY_LIMITS.outputBytes) fail('output exceeds its byte budget');
  const parsed = XLSX.read(bytes, {
    type: 'buffer',
    dense: false,
    cellFormula: true,
    cellNF: false,
    cellStyles: false,
    cellDates: false,
    sheetStubs: true,
    WTF: true,
  });
  if (
    parsed.vbaraw ||
    parsed.SheetNames.length !== output.SheetNames.length ||
    parsed.SheetNames.some((name, index) => name !== output.SheetNames[index])
  )
    fail('written workbook or worksheet identity differs');
  for (const name of output.SheetNames) {
    const source = parsed.Sheets[name],
      comparison = expected.get(name);
    let count = 0;
    for (const [ref, cell] of Object.entries(source)) {
      if (ref.startsWith('!')) continue;
      if (cell.t === 'z' && (cell.v === null || cell.v === undefined)) continue;
      const value = comparison.get(ref);
      if (!value || cell.f !== undefined || cell.F !== undefined || cell.t !== value.t || !Object.is(cell.v, value.v))
        fail('writer did not preserve the exact scalar type/value at ' + name + '!' + ref);
      count++;
    }
    if (count !== comparison.size) fail('writer omitted stored scalar values');
  }
  return {
    bytes,
    format,
    sheets: output.SheetNames.length,
    positions,
    values,
    formulaValues,
    omissions: [...DATA_ONLY_OMISSIONS],
    policy: 'Stored typed scalar values only; fresh workbook with no macros or original graphs',
  };
}
