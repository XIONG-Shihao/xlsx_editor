/** Compare previous full-string writer against byte-span output on the same file. */
import fs from 'node:fs';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readCsvWorkbook, editCsvBytes } from '@office-editors/xlsx-editor/core/csv-document.mjs';
import { encodeCsvText } from '@office-editors/xlsx-editor/core/csv-codec.mjs';
import { applyCsvCommands, serializeCsv } from '@office-editors/xlsx-editor/core/csv-workbook.mjs';
assert(
  process.env.XLSX_EDITOR_CSV_SOURCE && process.env.XLSX_EDITOR_CSV_TARGETS,
  'Set XLSX_EDITOR_CSV_SOURCE and XLSX_EDITOR_CSV_TARGETS to explicit fixture paths'
);
const [mode, output] = process.argv.slice(2),
  hash = (b) => createHash('sha256').update(b).digest('hex');
if (['previous', 'current'].includes(mode)) {
  const bytes = new Uint8Array(fs.readFileSync(process.env.XLSX_EDITOR_CSV_SOURCE)),
    sourceSha256 = hash(bytes),
    commands = JSON.parse(fs.readFileSync(process.env.XLSX_EDITOR_CSV_TARGETS)).edits.map((e) => ({
      sheet: 'CSV',
      cell: e.cell,
      value: e.csvReplacementText,
    }));
  const start = performance.now();
  let result;
  if (mode === 'current') result = editCsvBytes(bytes, commands, sourceSha256);
  else {
    const original = readCsvWorkbook(bytes);
    assert.equal(original.revision, sourceSha256);
    const edited = applyCsvCommands(original, commands);
    const output = encodeCsvText(serializeCsv(edited.index), original.encoding, original.bom);
    result = { bytes: output, revision: hash(output) };
  }
  const elapsedMs = performance.now() - start;
  assert.equal(hash(bytes), sourceSha256);
  console.log(
    JSON.stringify({ mode, elapsedMs, sourceSha256, outputSha256: result.revision, outputBytes: result.bytes.length })
  );
} else {
  assert.equal(mode, 'compare');
  assert(output);
  const runs = [];
  for (let i = 0; i < 3; i++)
    for (const version of i % 2 ? ['current', 'previous'] : ['previous', 'current']) {
      const p = spawnSync(process.execPath, [import.meta.filename, version], { encoding: 'utf8' });
      assert.equal(p.status, 0, p.stderr);
      runs.push(JSON.parse(p.stdout));
    }
  assert(runs.every((r) => r.outputSha256 === runs[0].outputSha256 && r.sourceSha256 === runs[0].sourceSha256));
  const medians = Object.fromEntries(
    ['previous', 'current'].map((v) => [
      v,
      runs
        .filter((r) => r.mode === v)
        .map((r) => r.elapsedMs)
        .sort((a, b) => a - b)[1],
    ])
  );
  fs.writeFileSync(
    output,
    JSON.stringify(
      {
        passed: true,
        method:
          'Three alternating fresh Node processes each. Previous full-string serialization/encoding algorithm versus sparse byte-spans; both include source read/index/hash, command guards and output hash. Disk read before timing. No Electron/RSS claim.',
        runs,
        medians,
      },
      null,
      2
    ) + '\n'
  );
  console.log(medians);
}
