/** Actual native theme style admission and owner snapshots, not PDF/handoff proof. */
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { SpreadsheetDocument } from '../adapters/document-session.mjs';
import { IndexedWorkbook } from '../core/indexed-session.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
import { pageThemeSamples, worksheetBackgroundSamples } from '../core/page-style-samples.mjs';
import { defaultPrintSettings } from '../core/print-model.mjs';
import { pageThemePrintFixture } from './support/page-theme-print-fixture.mjs';
const roots = [],
  settings = { ...defaultPrintSettings(), area: 'A1:F1' };
after(async () => {
  await closeNativeWorkbookWorker();
  for (const root of roots) await fs.rm(root, { recursive: true, force: true });
});
async function fixture(t, options = {}) {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'theme-print-'));
  roots.push(root);
  const source = path.join(root, 'source.xlsx');
  await fs.writeFile(source, pageThemePrintFixture(options));
  const sourceBytes = await fs.readFile(source),
    owners = [];
  t.after(async () => {
    for (const o of owners) await o.close();
  });
  const open = async (file = source) => {
    const o = new SpreadsheetDocument({ recoveryDirectory: path.join(root, 'recovery-' + owners.length) });
    owners.push(o);
    await o.open(file, { fresh: true });
    return o;
  };
  return { root, source, sourceBytes, open };
}
const request = (o, operation, more = {}) =>
  o.request({ documentId: o.status().documentId, state: o.status().state, operation, ...more });
test('three themes print resolved body/heading fonts and colors, fill, literal styles and geometry; background omitted', async (t) => {
  for (const [i, sample] of pageThemeSamples.entries()) {
    const f = await fixture(t),
      o = await f.open();
    await request(o, 'apply', {
      commands: [{ type: 'applyPageTheme', sample: sample.id, expectedThemeRevision: null }],
    });
    const background = worksheetBackgroundSamples[i % 2];
    await request(o, 'apply', {
      commands: [
        { type: 'setSheetBackground', sheet: 'Sheet1', sample: background.id, expectedBackgroundRevision: null },
      ],
    });
    const before = structuredClone(o.status()),
      job = await request(o, 'printPrepare', { sheet: 'Sheet1', settings }),
      page = job.pages[0],
      cell = (ref) => page.cells.find((c) => c.ref === ref);
    assert.equal(cell('A1').style.fontFamily, sample.minorFont);
    assert.equal(cell('A1').style.fontColor, sample.colors.accent1);
    assert.equal(cell('E1').style.fontFamily, sample.majorFont);
    assert.equal(cell('E1').style.fontColor, sample.colors.accent2);
    assert.equal(cell('F1').style.fillColor, sample.colors.accent3);
    assert.equal(cell('C1').style.fontFamily, 'Courier New');
    assert.equal(cell('C1').style.fontColor, '11AA44');
    assert.equal(cell('B1').text, '6');
    assert.equal(page.objects.length, 0);
    assert.equal(job.geometryProfile.font.fontFamily, sample.minorFont);
    assert.equal(job.geometryProfile.font.fontSize, 11);
    const document = await request(o, 'printDocument', { sheet: 'Sheet1', jobId: job.id });
    assert(!document.html.includes(background.base64));
    assert(document.html.includes(sample.minorFont));
    assert(document.html.includes(sample.majorFont));
    assert.deepEqual(o.status(), before);
    await request(o, 'printRelease', { jobId: job.id });
    const saved = path.join(f.root, sample.id + '.xlsx');
    await request(o, 'copy', { destination: saved, expectedRevision: null });
    await o.close();
    const cold = await f.open(saved),
      fresh = await request(cold, 'printPrepare', { sheet: 'Sheet1', settings });
    assert.equal(fresh.pages[0].cells.find((c) => c.ref === 'A1').style.fontFamily, sample.minorFont);
    await request(cold, 'printRelease', { jobId: fresh.id });
    await cold.close();
    assert.deepEqual(await fs.readFile(f.source), f.sourceBytes);
  }
});
test('imported, modified and tinted themes refuse print without altering owner/source', async (t) => {
  for (const options of [
    { theme: 'forest', foreign: true },
    { theme: 'forest', modified: true },
    { theme: 'forest', tint: true },
  ]) {
    const f = await fixture(t, options),
      o = await f.open(),
      before = structuredClone(o.status());
    await assert.rejects(
      request(o, 'printPrepare', { sheet: 'Sheet1', settings }),
      /PRINT_UNSUPPORTED|Unsupported|unsupported|theme|THEME/
    );
    assert.deepEqual(o.status(), before);
    assert.deepEqual(await fs.readFile(f.source), f.sourceBytes);
    await o.close();
  }
});
test('mismatched native Normal-font or theme identity facts cannot produce a snapshot', async (t) => {
  for (const mode of ['font', 'theme']) {
    const f = await fixture(t, { theme: 'forest', formula: false });
    const bridge = async (input) => {
      const reply = await native(input);
      if (input.operation === 'print-check') {
        if (mode === 'font') reply.normalFont.fontFamily = 'Courier New';
        else reply.pageTheme.revision = '0'.repeat(64);
      }
      return reply;
    };
    const book = new IndexedWorkbook(bridge, () => assert.fail('No compatibility writer'));
    t.after(() => book.close());
    await book.open(f.source);
    const before = book.view('0');
    await assert.rejects(
      book.printSnapshot('0', 'Sheet1', settings),
      (e) => e.code === (mode === 'font' ? 'PRINT_UNRESOLVED_GEOMETRY' : 'PRINT_UNSUPPORTED_THEME')
    );
    assert.deepEqual(book.view('0'), before);
    assert.deepEqual(await fs.readFile(f.source), f.sourceBytes);
  }
});
