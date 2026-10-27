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
import { printPageHtml } from '../core/print-layout.mjs';
import { chartSeriesAppearance } from '../core/chart-styles.mjs';
const out = await fs.mkdtemp(path.resolve('.build/tests/chart-style-preview-')),
  file = path.join(out, 'ui.mjs');
await build({
  stdin: {
    contents: 'export {WorkbookObjectPreviewView} from "./ui/WorkbookObjects";',
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
const { WorkbookObjectPreviewView } = await import(pathToFileURL(file));
const object = (kind, style) => ({
  id: 'chart',
  name: 'Styled sample',
  kind: 'chart',
  managed: true,
  canEdit: true,
  canDelete: true,
  protectedReason: null,
  chart: { kind, style, range: 'A1:B4', title: 'Styled sample', legend: 'bottom' },
  data: { categories: ['A', 'B', 'C'], series: [{ name: 'One', values: [2, 5, 8] }], hidden: [false, false, false] },
});
function inspect(html, kind, style) {
  const doc = new JSDOM(html).window.document,
    svg = doc.querySelector('svg'),
    a = chartSeriesAppearance(style, 0);
  assert(svg);
  if (kind === 'line') {
    const line = svg.querySelector('polyline');
    assert(line);
    assert.equal(line.getAttribute('stroke'), a.color);
    assert.equal(line.getAttribute('stroke-dasharray'), a.dash ?? null);
  } else {
    const bars = svg.querySelectorAll(':scope > rect');
    assert.equal(bars.length, 3);
    assert.equal(bars[0].getAttribute('fill'), a.fill);
    assert.equal(bars[0].getAttribute('stroke'), a.stroke ?? null);
    assert.equal(bars[0].getAttribute('stroke-dasharray'), a.dash ?? null);
  }
  assert.equal(svg.querySelector('img,image'), null);
  return svg.outerHTML;
}
test('worksheet preview paints all three actual fill/line palettes for column/bar/line, rather than only relabelling styles', () => {
  for (const kind of ['column', 'bar', 'line']) {
    const variants = [];
    for (const style of ['colorful', 'monochrome', 'outline'])
      variants.push(
        inspect(
          renderToStaticMarkup(React.createElement(WorkbookObjectPreviewView, { object: object(kind, style) })),
          kind,
          style
        )
      );
    assert.equal(new Set(variants).size, 3);
  }
});
test('print SVG consumes the same explicit saved styles and preserves readable titles/categories/legends', () => {
  for (const kind of ['column', 'bar', 'line'])
    for (const style of ['colorful', 'monochrome', 'outline']) {
      const item = object(kind, style),
        html = printPageHtml(
          {
            sheet: 'Data',
            scale: 1,
            pageSizePt: [612, 792],
            settings: {},
            pages: [
              {
                number: 1,
                cells: [],
                objects: [
                  { id: 'chart', kind: 'chart', object: item, bbox: [0, 0, 400, 240], clipBox: [0, 0, 400, 240] },
                ],
              },
            ],
          },
          1
        );
      inspect(html, kind, style);
      assert.match(html, /Styled sample/);
      assert.match(html, />One</);
    }
});
