/** Compare encoder implementations on immutable source bytes in fresh processes.
 * Explicit cloning approximates one transport hop, not the Electron UI journey. */
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { readCsvWorkbook } from '@office-editors/xlsx-editor/core/csv-document.mjs';
import { decodeCsvTransfer } from '@office-editors/xlsx-editor/csv-transfer';
import { csvReadRange } from '@office-editors/xlsx-editor/core/csv-workbook.mjs';
const [mode, previous, output] = process.argv.slice(2);
assert(process.env.XLSX_EDITOR_CSV_SOURCE, 'Set XLSX_EDITOR_CSV_SOURCE to an explicit fixture path');
const source = process.env.XLSX_EDITOR_CSV_SOURCE,
  hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
if (mode === 'previous' || mode === 'current') {
  const module =
    mode === 'previous'
      ? path.resolve(previous)
      : fileURLToPath(import.meta.resolve('@office-editors/xlsx-editor/csv-transfer-node'));
  const { encodeCsvTransfer } = await import(pathToFileURL(module));
  const bytes = new Uint8Array(fs.readFileSync(source)),
    sourceSha256 = hash(bytes),
    timings = {};
  const time = (name, fn) => {
    const start = performance.now();
    const result = fn();
    timings[name] = performance.now() - start;
    return result;
  };
  const start = performance.now(),
    view = time('indexMs', () => readCsvWorkbook(bytes));
  const encoded = time('encodeMs', () => encodeCsvTransfer(view));
  const transferred = time('cloneMs', () => structuredClone(encoded));
  const decoded = time('decodeMs', () => decodeCsvTransfer(transferred));
  time('visibleRowsMs', () => {
    for (const row of [1, Math.floor(view.rows / 2), view.rows]) csvReadRange(decoded, `A${row}:O${row}`);
  });
  timings.totalMs = performance.now() - start;
  assert.deepEqual({ ...decoded, text: decoded.text.slice(0) }, view);
  assert.equal(hash(fs.readFileSync(source)), sourceSha256);
  console.log(
    JSON.stringify({
      mode,
      sourceSha256,
      sourceBytes: bytes.length,
      encoderSha256: hash(fs.readFileSync(module)),
      compressedBytes:
        encoded.transport === 'csv-gzip-v1'
          ? encoded.compressedText.length
          : encoded.blocks.reduce((n, b) => n + b.length, 0),
      timings,
      passed: true,
    })
  );
} else {
  assert.equal(mode, 'compare');
  assert(previous && output, 'Usage: compare previous-encoder.ts output.json');
  const runs = [];
  for (let i = 0; i < 3; i++)
    for (const version of i % 2 ? ['current', 'previous'] : ['previous', 'current']) {
      const child = spawnSync(
        process.execPath,
        [
          '--import',
          new URL('../ts-resolve-register.mjs', import.meta.url).href,
          import.meta.filename,
          version,
          previous,
        ],
        { encoding: 'utf8' }
      );
      assert.equal(child.status, 0, child.stderr);
      runs.push(JSON.parse(child.stdout));
    }
  const medians = Object.fromEntries(
    ['previous', 'current'].map((version) => [
      version,
      Object.fromEntries(
        Object.keys(runs[0].timings).map((key) => [
          key,
          runs
            .filter((r) => r.mode === version)
            .map((r) => r.timings[key])
            .sort((a, b) => a - b)[1],
        ])
      ),
    ])
  );
  const report = {
    schema: 'workbook.csv.transfer-benchmark.v2',
    recordedAt: new Date().toISOString(),
    method:
      'Three alternating fresh Node processes per encoder, identical 53 MB supplied CSV. Parse, encode, clone once, decode and read first/middle/last rows; full source equality/hashes checked outside timed interval. Block transport decodes only requested blocks. No Electron, forced GC or app RSS claim.',
    runs,
    medians,
    passed: true,
  };
  fs.mkdirSync(path.dirname(output), { recursive: true });
  fs.writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify(medians, null, 2));
}
