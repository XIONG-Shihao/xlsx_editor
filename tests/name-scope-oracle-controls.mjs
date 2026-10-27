/** Controls for the independent name oracle's narrowly allowed blank normalization. */
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { nameScopeOwnerFixture } from './support/name-scope-owner-fixture.mjs';
test('name saved-output oracle permits only inert style-zero markers and refuses real data/style/metadata changes', async (t) => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'name-oracle-controls-'));
  t.after(() => fs.rm(dir, { recursive: true, force: true }));
  const original = path.join(dir, 'source.xlsx'),
    saved = path.join(dir, 'saved.xlsx'),
    expected = path.join(dir, 'expected.json'),
    before = nameScopeOwnerFixture(),
    parts = unzipSync(before),
    part = 'xl/worksheets/sheet1.xml';
  await fs.writeFile(original, before);
  await fs.writeFile(expected, JSON.stringify({ [part]: { C1: { formula: 'Rate', value: 2 } } }));
  const insert = (xml) => xml.replace('<c r="C1" s="0"/>', '<c r="C1"><f>Rate</f><v>2</v></c>'),
    check = () =>
      execFileSync(
        process.env.XLSX_DESKTOP_PYTHON ?? 'python3',
        [fileURLToPath(new URL('./support/check-name-scope-owner.py', import.meta.url)), original, saved, expected],
        { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] }
      );
  parts[part] = strToU8(insert(strFromU8(parts[part])));
  await fs.writeFile(saved, zipSync(parts));
  assert.equal(JSON.parse(check()).passed, true);
  const normalized = strFromU8(parts[part])
    .replace(/<c r="C[2-6]" s="0"\/>/g, '')
    .replace('<row r="5"></row>', '');
  parts[part] = strToU8(normalized);
  await fs.writeFile(saved, zipSync(parts));
  assert.equal(JSON.parse(check()).passed, true);
  const untouched = 'xl/worksheets/sheet2.xml',
    untouchedOriginal = strFromU8(parts[untouched]);
  parts[untouched] = strToU8(
    untouchedOriginal.replace(
      '<c r="A1">',
      '<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="A1">'
    )
  );
  await fs.writeFile(saved, zipSync(parts));
  assert.deepEqual(JSON.parse(check()).unchangedWorksheetsSemanticallyVerified, [untouched]);
  parts[untouched] = strToU8(untouchedOriginal.replace('<v>23</v>', '<v>999</v>'));
  await fs.writeFile(saved, zipSync(parts));
  assert.throws(check, (error) => error.status === 1, 'untouched-sheet literal change');
  parts[untouched] = strToU8(untouchedOriginal);
  // A wider original hint can normalize only to the real saved-cell bounds.
  parts[untouched] = strToU8(untouchedOriginal.replace('ref="A1:F6"', 'ref="A1:Z99"'));
  const widerSource = unzipSync(before);
  widerSource[untouched] = parts[untouched];
  await fs.writeFile(original, zipSync(widerSource));
  parts[untouched] = strToU8(untouchedOriginal);
  await fs.writeFile(saved, zipSync(parts));
  assert.equal(JSON.parse(check()).normalizedWorksheetDimensions.length, 1);
  parts[untouched] = strToU8(untouchedOriginal.replace('ref="A1:F6"', 'ref="A1:B2"'));
  await fs.writeFile(saved, zipSync(parts));
  assert.throws(check, (error) => error.status === 1, 'wrong saved dimension');
  parts[untouched] = strToU8(untouchedOriginal);
  await fs.writeFile(original, before);
  const opaque = parts['custom/name-scope-preserved.bin'].slice();
  parts['custom/name-scope-preserved.bin'][0] = 1;
  await fs.writeFile(saved, zipSync(parts));
  assert.throws(check, (error) => error.status === 1, 'opaque bytes');
  parts['custom/name-scope-preserved.bin'] = opaque;
  for (const [label, xml] of [
    ['real style', normalized.replace('<row r="2">', '<row r="2"><c r="C2" s="1"/>')],
    ['opaque metadata', normalized.replace('<row r="2">', '<row r="2"><c r="C2" s="0" vm="1"/>')],
    ['literal changed', normalized.replace('<v>13</v>', '<v>999</v>')],
    ['wrong cache', normalized.replace('<f>Rate</f><v>2</v>', '<f>Rate</f><v>5</v>')],
  ]) {
    parts[part] = strToU8(xml);
    await fs.writeFile(saved, zipSync(parts));
    assert.throws(check, (error) => error.status === 1, label);
  }
});
