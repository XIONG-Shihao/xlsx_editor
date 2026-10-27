/** Independent fixed-width records and explicit expected field boundaries.
 * Literal test input only; no XLSX writer, calculation or native authority. */
export const fixedWidthFixture = Object.freeze({
  breaks: Object.freeze([5, 13, 18, 28]),
  records: Object.freeze(['00123   12.50TRUE 2011-07-05=SUM(A1)', 'abcde      -0FALSE2011-07-05=1+2']),
  fields: Object.freeze([
    Object.freeze(['00123', '   12.50', 'TRUE ', '2011-07-05', '=SUM(A1)']),
    Object.freeze(['abcde', '      -0', 'FALSE', '2011-07-05', '=1+2']),
  ]),
  policies: Object.freeze(['text', 'number', 'boolean', 'iso-date', 'text']),
  calendars: Object.freeze({ 1900: 40729, 1904: 39267 }),
  unicode: Object.freeze({
    record: 'A😀e\u0301Z',
    breaks: Object.freeze([1, 2, 3, 4]),
    fields: Object.freeze(['A', '😀', 'e', '\u0301', 'Z']),
  }),
  short: Object.freeze({
    record: 'ABCDE',
    breaks: Object.freeze([5, 8]),
    fields: Object.freeze(['ABCDE', '', '']),
    padded: Object.freeze([false, false, true]),
  }),
  empty: Object.freeze({
    record: '',
    breaks: Object.freeze([5, 8]),
    fields: Object.freeze(['', '', '']),
    padded: Object.freeze([false, true, true]),
  }),
});
