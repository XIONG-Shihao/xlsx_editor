import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  emptyPrintHeaderFooter,
  encodePrintSections,
  decodePrintSections,
  validatePrintHeaderFooter,
} from '../core/print-header-footer.mjs';
import { defaultPrintSettings, validatePrintSettings } from '../core/print-model.mjs';
import { composePrint, printDocumentHtml, printOverflowInspectionSource } from '../core/print-layout.mjs';
test('six sections roundtrip literal ampersands, XML text and Unicode without promoting text to dynamic fields', () => {
  const hf = emptyPrintHeaderFooter();
  hf.header = { left: 'R&D', center: 'Sales <Q1>', right: '😀' };
  hf.footer.center = 'Literal &P and &N';
  assert.equal(encodePrintSections(hf.header), '&LR&&D&CSales <Q1>&R😀');
  assert.deepEqual(decodePrintSections(encodePrintSections(hf.header)), hf.header);
  assert.deepEqual(decodePrintSections(encodePrintSections(hf.footer)), hf.footer);
  const copy = validatePrintHeaderFooter(hf);
  copy.header.center = 'changed';
  assert.equal(hf.header.center, 'Sales <Q1>');
});
test('formatted/first-even/dynamic codes, malformed Unicode and oversized encoded text refuse', () => {
  for (const raw of ['unsectioned', '&C&P', '&C&G', '&C&BStyled', '&C&', '&C&Q'])
    assert.throws(() => decodePrintSections(raw), /protected|Unqualified/);
  for (const value of ['bad\nline', 'bad\u0000text', '\ud800', '\ufffe'])
    assert.throws(() => encodePrintSections({ left: value, center: '', right: '' }), {
      code: /^(note-bad-input|print-bad-input)$/,
    });
  assert.throws(() => encodePrintSections({ left: '&'.repeat(127), center: '', right: '' }), {
    code: /^(print-limit|rules-limit)$/,
  });
  assert.equal(encodePrintSections({ left: '&'.repeat(126), center: '', right: '' }).length, 254);
  assert.throws(() => validatePrintHeaderFooter({ ...emptyPrintHeaderFooter(), scaleWithDoc: 1 }), /scaling/);
});
test('old settings remain unchanged and header/footer drafts normalize independently', () => {
  const old = defaultPrintSettings();
  assert.deepEqual(validatePrintSettings(old, { allowAbsentArea: true }), old);
  const value = { ...old, area: 'A1:B3', headerFooter: emptyPrintHeaderFooter() };
  assert.deepEqual(validatePrintSettings(value), value);
  assert.throws(
    () => validatePrintSettings({ ...value, headerFooter: { ...value.headerFooter, extra: 1 } }),
    /Invalid/
  );
});
function snapshot(hf, patch = {}) {
  return {
    schema: 1,
    fingerprint: 'test',
    revision: 'r1',
    sheet: 'Sales',
    settings: { ...defaultPrintSettings(), area: 'A1:B3', headerFooter: hf, ...patch },
    rows: [0, 1, 2].map((index) => ({ index, heightPt: 18, hidden: false })),
    columns: [0, 1].map((index) => ({ index, widthPt: 72, hidden: false })),
    cells: [0, 1, 2].flatMap((r) =>
      [0, 1].map((c) => ({
        ref: String.fromCharCode(65 + c) + (r + 1),
        text: r === 0 && c === 0 ? '42' : '',
        style: {},
      }))
    ),
    objects: [],
    merges: [],
  };
}
test('physical output renders three actual plain-text presets in separate header/footer bands', () => {
  const presets = [
    ['Sales', ''],
    ['Sales', 'Confidential'],
    ['Draft', 'Sales'],
  ];
  for (const [header, footer] of presets) {
    const hf = emptyPrintHeaderFooter();
    hf.header.center = header;
    hf.footer.right = footer;
    const layout = composePrint(snapshot(hf)),
      html = printDocumentHtml(layout);
    assert.ok(html.includes('data-print-furniture="header-center"'));
    assert.ok(html.includes('>' + header + '</div>'));
    if (footer) assert.ok(html.includes('data-print-furniture="footer-right"'));
    assert.deepEqual(layout.pages[0].cells[0].bbox, [36, 36, 72, 18]);
    assert.ok(html.includes('>42</span>'));
  }
});
test('margin overlap refuses while explicit unscaled furniture stays independent of cell scale', () => {
  const hf = emptyPrintHeaderFooter();
  hf.header.center = 'Header';
  hf.footer.center = '<script>& literal';
  hf.scaleWithDoc = false;
  const layout = composePrint(snapshot(hf, { scaling: { mode: 'scale', percent: 50 } })),
    html = printDocumentHtml(layout);
  assert.match(html, /font-size:10pt/);
  assert.ok(html.includes('&lt;script&gt;&amp; literal'));
  assert.equal(layout.scale, 0.5);
  assert.throws(
    () =>
      printDocumentHtml(
        composePrint(snapshot(hf, { marginsPt: { ...defaultPrintSettings().marginsPt, top: 20, header: 18 } }))
      ),
    (e) => e.code === 'PRINT_HEADER_FOOTER_MARGIN'
  );
  assert.match(printOverflowInspectionSource(), /PRINT_HEADER_FOOTER_OVERFLOW/);
});
