// Every place that can refuse carries a code from the registry (core/refusal-codes.mjs).
// A new refusal written without one fails here; the fix is `node scripts/refusal-codes/apply.mjs`
// (and a rule in scripts/refusal-codes/mapping.mjs if no rule covers the new sentence), never an allowance.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { problems, readAllowList, uncoded } from '../scripts/refusal-codes/check.mjs';
import { collect } from '../scripts/refusal-codes/sites.mjs';
import { REFUSAL_CODES } from '../core/refusal-codes.mjs';
import { scanJs } from '../scripts/refusal-codes/scan-js.mjs';
import { scanRs } from '../scripts/refusal-codes/scan-rs.mjs';

test('no refusal site lacks a code beyond the allow-list, and the allow-list only shrinks', () => {
  assert.deepEqual(problems(), []);
});

test('the allow-list holds only files that still have uncoded sites', () => {
  const files = new Set(uncoded().map((site) => site.file));
  for (const file of Object.keys(readAllowList().files))
    assert.ok(files.has(file), `${file} is fully coded; remove it from the allow-list`);
});

test('every code the mapping can give exists in the registry, with a meaning', () => {
  for (const site of collect().filter((s) => s.status === 'code'))
    assert.ok(Object.hasOwn(REFUSAL_CODES, site.code), site.code);
  for (const [code, meaning] of Object.entries(REFUSAL_CODES)) {
    assert.match(code, /^[a-z]+(-[a-z]+)*$/);
    assert.ok(meaning.length > 20, `${code} needs a meaning`);
  }
});

test('the scanners find the refusal forms they are meant to', () => {
  const js = scanJs(
    'x.mjs',
    "throw Error('A'); throw new Error(`B ${x}`); const e = Object.assign(Error('C'), { code: 'x' }); fail(Error(m));"
  );
  assert.deepEqual(
    js.map((s) => [s.message, s.dynamic, s.legacyCode]),
    [
      ['A', false, false],
      ['B {}', false, false],
      ['C', false, true],
      ['{}', true, false],
    ]
  );
  const rs = scanRs(
    'x.rs',
    `fn f() -> Result<()> {
      Err("A".into())?;
      Err(format!("B {x}"))?;
      Err("C")?; // a &str marker, not a sentence
      x.ok_or("D")?;
      x.map_err(|e| format!("E {e}"))?;
      Err(crate::coded("c-d", "F"))?;
      if let Err(e) = y {}
      Err(reason)
    }
    #[cfg(test)]
    mod tests { fn t() { Err("G".into()) } }`
  );
  assert.deepEqual(
    rs.map((s) => [s.form, s.message, s.dynamic]),
    [
      ['Err', 'A', false],
      ['Err', 'B {}', false],
      ['ok_or', 'D', false],
      ['map_err', 'E {}', false],
      ['Err', '', true],
    ]
  );
  assert.equal(fs.existsSync(new URL('../scripts/refusal-codes/mapping.mjs', import.meta.url)), true);
});
