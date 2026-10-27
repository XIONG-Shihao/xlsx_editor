import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import { createWorkbook } from '../core/index.mjs';

test('independent saved-paste oracle rejects bool/text substitutions, precision drift, formulas and broken cell coverage', async (t) => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-staged-oracle-'));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const source = path.join(directory, 'source.xlsx'),
    saved = path.join(directory, 'saved.xlsx'),
    parts = unzipSync(createWorkbook());
  parts['custom/opaque.bin'] = new Uint8Array([0, 255, 17]);
  await fs.writeFile(source, zipSync(parts));
  const inspect = (cell) => {
    const contents = {
      ...parts,
      'xl/worksheets/sheet1.xml': strToU8(
        `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">${cell}</row></sheetData></worksheet>`
      ),
    };
    return fs.writeFile(saved, zipSync(contents)).then(() =>
      JSON.parse(
        execFileSync('python3', ['tests/support/staged-paste-oracle.py', source, saved, '1', '1'], {
          encoding: 'utf8',
          stdio: ['ignore', 'pipe', 'pipe'],
        })
      )
    );
  };
  assert.equal((await inspect('<c r="A1"><v>1.0</v></c>')).binary64Verified, true);
  for (const cell of [
    '<c r="A1" t="b"><v>1</v></c>',
    '<c r="A1" t="str"><v>1</v></c>',
    '<c r="A1"><v>1.0000000000000002</v></c>',
    '<c r="A1"><f>1+0</f><v>1</v></c>',
    '<c r="B1"><v>1</v></c>',
    '<c r="A1"><v>1</v></c><c r="A1"><v>1</v></c>',
    '',
  ])
    await assert.rejects(inspect(cell));
  await inspect('<c r="A1"><v>1</v></c>');
  const changed = { ...unzipSync(await fs.readFile(saved)), 'custom/opaque.bin': new Uint8Array([0, 255, 18]) };
  await fs.writeFile(saved, zipSync(changed));
  assert.throws(() =>
    execFileSync('python3', ['tests/support/staged-paste-oracle.py', source, saved, '1', '1'], {
      stdio: ['ignore', 'pipe', 'pipe'],
    })
  );
});
