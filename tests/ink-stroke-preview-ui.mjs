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
import { inkPen } from '../core/ink-stroke.mjs';
import { inkFoundationCases } from './support/ink-foundation-fixture.mjs';
const phase = path.resolve('.build/tests/dw-m1-stroke-foundation-phase-1');
await fs.mkdir(phase, { recursive: true });
const file = path.join(phase, 'preview.mjs');
await build({
  stdin: {
    contents: 'export {WorkbookInkPreview} from "./ui/WorkbookInkPreview";',
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
const { WorkbookInkPreview } = await import(pathToFileURL(file));
test('SVG renderer paints actual ordered vector points, width/caps/opacity and escaped labels for all three pen models', async () => {
  for (const c of inkFoundationCases()) {
    const label = 'Stroke <&é 😀 ' + c.label,
      html = renderToStaticMarkup(React.createElement(WorkbookInkPreview, { stroke: c.stroke, label })),
      dom = new JSDOM(html),
      doc = dom.window.document,
      svg = doc.querySelector('svg'),
      p = svg.querySelector('path'),
      pen = inkPen(c.pen);
    assert.equal(svg.getAttribute('viewBox'), `0 0 ${c.stroke.widthEmu} ${c.stroke.heightEmu}`);
    assert.equal(svg.getAttribute('aria-label'), label);
    assert.equal(svg.querySelector('title').textContent, label);
    assert.equal(p.getAttribute('d'), c.stroke.points.map(([x, y], i) => (i ? 'L' : 'M') + x + ' ' + y).join(' '));
    assert.equal(p.getAttribute('stroke'), '#' + pen.color);
    assert.equal(p.getAttribute('stroke-width'), String(pen.widthEmu));
    assert.equal(p.getAttribute('stroke-opacity'), String(pen.alpha / 100000));
    assert.equal(p.getAttribute('stroke-linecap'), pen.cap === 'round' ? 'round' : 'butt');
    assert.equal(p.getAttribute('fill'), 'none');
    assert.equal(doc.querySelector('img,image,canvas,foreignObject'), null);
    assert(!/(NaN|Infinity)/.test(html));
    await fs.writeFile(path.join(phase, c.pen + '.svg'), html);
    dom.window.close();
  }
});
test('renderer consumes changed path data and refuses an invalid model instead of painting a generic sample image', () => {
  const c = inkFoundationCases()[0],
    a = renderToStaticMarkup(React.createElement(WorkbookInkPreview, { stroke: c.stroke })),
    model = {
      ...c.stroke,
      points: [
        c.stroke.points[0],
        [c.stroke.points[1][0], c.stroke.points[2][1]],
        c.stroke.points[1],
        c.stroke.points[2],
      ],
    },
    b = renderToStaticMarkup(React.createElement(WorkbookInkPreview, { stroke: model }));
  assert.notEqual(a, b);
  assert.throws(() =>
    renderToStaticMarkup(React.createElement(WorkbookInkPreview, { stroke: { ...c.stroke, widthEmu: Infinity } }))
  );
});
