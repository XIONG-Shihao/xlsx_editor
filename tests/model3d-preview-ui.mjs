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
import { MODEL3D_SAMPLES, model3DGlb, parseModel3DGlb } from '../core/model3d-model.mjs';
const phase = path.resolve('.build/tests/in-m3-model3d-foundation-phase-1'),
  bundle = path.join(phase, 'preview.mjs');
await fs.mkdir(phase, { recursive: true });
await build({
  stdin: {
    contents: 'export {WorkbookModel3DPreview} from "./ui/WorkbookModel3DPreview";',
    resolveDir: process.cwd(),
    loader: 'tsx',
  },
  bundle: true,
  outfile: bundle,
  format: 'esm',
  platform: 'node',
  packages: 'external',
  jsx: 'automatic',
  logLevel: 'silent',
});
const { WorkbookModel3DPreview } = await import(pathToFileURL(bundle));
test('actual decoded cube/pyramid/octahedron render lit mesh triangles with accessible labels and no bitmap', async () => {
  for (const sample of MODEL3D_SAMPLES) {
    const asset = parseModel3DGlb(await fs.readFile('core/assets/models/' + sample.id + '.glb')),
      html = renderToStaticMarkup(
        React.createElement(WorkbookModel3DPreview, { asset, label: 'Model <&é ' + sample.label })
      ),
      dom = new JSDOM(html),
      doc = dom.window.document;
    assert.equal(doc.querySelector('svg').getAttribute('data-model3d-sample'), sample.id);
    assert.equal(doc.querySelector('title').textContent, 'Model <&é ' + sample.label);
    assert(doc.querySelectorAll('polygon').length > 0);
    assert.equal(doc.querySelector('img,image,canvas'), null);
    assert(!/(NaN|Infinity)/.test(html));
    await fs.writeFile(path.join(phase, sample.id + '.svg'), html);
    dom.window.close();
  }
});
test('different actual GLB rotations/materials change rendered faces and colors rather than a fixed thumbnail', () => {
  for (const sample of MODEL3D_SAMPLES) {
    const make = (model) =>
        renderToStaticMarkup(
          React.createElement(WorkbookModel3DPreview, { asset: parseModel3DGlb(model3DGlb(model)) })
        ),
      a = make({ sample: sample.id, rotation: { x: 0, y: 0, z: 0 } }),
      b = make({ sample: sample.id, color: 'FF0000', rotation: { x: 25, y: -40, z: 15 } });
    assert.notEqual(a, b);
    const dom = new JSDOM(b),
      fills = [...dom.window.document.querySelectorAll('polygon')].map((n) => n.getAttribute('fill'));
    assert(fills.every((s) => /^#[0-9a-f]{2}0000$/.test(s)));
    dom.window.close();
  }
});
