import './support/scratch-dir.mjs';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { JSDOM } from 'jsdom';
import { smartartSampleModel, normalizeSmartArt } from '../core/smartart-model.mjs';
const out = await fs.mkdtemp(path.resolve('.build/tests/smartart-preview-ui-')),
  file = path.join(out, 'ui.mjs');
await build({
  stdin: {
    contents:
      'export {WorkbookSmartArtPreview} from "./ui/WorkbookSmartArtPreview";export {WorkbookObjectPreviewView} from "./ui/WorkbookObjects";',
    resolveDir: process.cwd(),
    loader: 'tsx',
  },
  bundle: true,
  outfile: file,
  format: 'esm',
  platform: 'node',
  packages: 'external',
  jsx: 'automatic',
  logLevel: 'silent',
});
const { WorkbookSmartArtPreview, WorkbookObjectPreviewView } = await import(pathToFileURL(file));
test('three actual semantic sample models render separate nodes/connections without a bitmap', () => {
  for (const layout of ['process', 'hierarchy', 'cycle']) {
    const model = smartartSampleModel(layout),
      doc = new JSDOM(
        renderToStaticMarkup(
          React.createElement(WorkbookSmartArtPreview, { model, width: 600, height: 360, label: 'Sample ' + layout })
        )
      ).window.document;
    assert.equal(doc.querySelector('svg').getAttribute('data-smartart-layout'), layout);
    assert.equal(doc.querySelectorAll('[data-smartart-node]').length, 3);
    assert.equal(doc.querySelectorAll('[data-smartart-connection]').length, layout === 'cycle' ? 3 : 2);
    assert.equal(doc.querySelector('image,img'), null);
    assert.equal(doc.querySelector('[data-smartart-node="n1"] foreignObject').textContent, model.nodes[0].text);
    assert(!/(NaN|Infinity)/.test(doc.documentElement.outerHTML));
    doc.defaultView.close();
  }
});
test('renderer accepts only finite sample topologies up to six nodes and output text is escaped', () => {
  for (const layout of ['process', 'hierarchy', 'cycle'])
    for (const n of [2, 6]) {
      const model = normalizeSmartArt({
          layout,
          nodes: Array.from({ length: n }, (_, i) => ({ id: 'n' + (i + 1), text: 'Node <&é ' + i })),
        }),
        html = renderToStaticMarkup(React.createElement(WorkbookSmartArtPreview, { model, width: 300, height: 180 }));
      assert(!/(NaN|Infinity)/.test(html));
      assert.match(html, /&lt;&amp;é/);
    }
});
test('existing object preview renders a validated protected diagram with an explicit editing route', () => {
  const object = {
    name: 'Sample',
    kind: 'smartart',
    managed: true,
    canEdit: false,
    canDelete: false,
    protectedReason: 'SmartArt authoring is not available yet',
    smartart: smartartSampleModel('process'),
  };
  const html = renderToStaticMarkup(React.createElement(WorkbookObjectPreviewView, { object }));
  assert.match(html, /data-smartart-layout="process"/);
  assert.match(html, /view only/);
  assert.doesNotMatch(html, /Preview is unavailable/);
});
