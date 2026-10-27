/** Controlled authority for model/Grid checks. No native-import claim. */
export const providerReceipt = Object.freeze({
  documentId: 'doc',
  ownerId: 'indexed-provider-owner',
  state: 'accepted-head',
  revision: 'accepted-view-revision',
});
export const initialCells = () => ({
  A1: { value: 7, dataType: 'number' },
  A3: { value: 2, formula: '1+1', dataType: 'number' },
  C3: { value: '', formula: '""', dataType: 'text' },
  C4: { value: 'text', dataType: 'text' },
  A5: { value: 7, dataType: 'number' },
  C5: { value: true, dataType: 'boolean' },
  C6: { value: '#N/A', dataType: 'error' },
});
const point = (reference) => {
  const match = /^([A-Z]+)([1-9]\d*)$/.exec(reference);
  if (!match) throw Error('Fixture needs canonical cell.');
  return {
    row: Number(match[2]) - 1,
    column: [...match[1]].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0) - 1,
  };
};
export function bounds(range) {
  const [first, last = first] = range.split(':'),
    a = point(first),
    b = point(last);
  return { minRow: a.row, minColumn: a.column, maxRow: b.row, maxColumn: b.column };
}
export function contains(range, reference) {
  const outer = bounds(range),
    at = point(reference);
  return (
    at.row >= outer.minRow && at.row <= outer.maxRow && at.column >= outer.minColumn && at.column <= outer.maxColumn
  );
}
export function completeUsedRange(
  range = 'A1:C6',
  {
    receipt = providerReceipt,
    sheet = 'Sheet1',
    empty = false,
    merges = null,
    acceptedHistory = null,
    version = 2,
    sourceRevision = 'a'.repeat(64),
  } = {}
) {
  const box = bounds(range),
    last = range.split(':').at(-1);
  return {
    schema: 1,
    policy: 'xlsx-used-range-v1',
    status: 'complete',
    receipt: { ...receipt },
    sheet,
    native: { version, sourceRevision, sheetPath: 'xl/worksheets/original-path.xml' },
    scopeMinMax: box,
    range,
    lastCell: last,
    empty,
    provenance: {
      factsSchema: 'workbook.xlsx.used-range-facts.v1',
      declaration: { status: 'valid', ref: range, bounds: { ...box } },
      sourceCells: empty ? null : { ...box },
      sourceMerges: merges ? bounds(merges) : null,
      effectiveCells: empty ? null : { ...box },
      sourceAxisFormatting: null,
      acceptedHistory: acceptedHistory ? bounds(acceptedHistory) : null,
      sourceDefaultFormatting: false,
    },
    reasons: [],
  };
}
export function incompleteUsedRange({ receipt = providerReceipt, sheet = 'Sheet1' } = {}) {
  const reply = completeUsedRange('A1:C6', { receipt, sheet });
  return {
    ...reply,
    status: 'incomplete',
    scopeMinMax: null,
    range: null,
    lastCell: null,
    empty: null,
    reasons: [{ code: 'source-facts-unavailable', message: 'The complete source facts are unavailable.' }],
  };
}
export const sourceReader =
  (cells, calls = []) =>
  async (sheet, range) => {
    calls.push({ sheet, range });
    return {
      range,
      complete: true,
      cells: Object.fromEntries(
        Object.entries(cells)
          .filter(([reference]) => contains(range, reference))
          .map(([reference, value]) => [reference, structuredClone(value)])
      ),
    };
  };
