/** Independent native plan mutants: malformed quote, opaque ghost worksheet
 * and out-of-range local scope cannot establish a static-name target. */
import test, { after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8, strFromU8 } from 'fflate';
import { createWorkbook } from '../core/workbook-package.mjs';
import { nativeWorkbookOperationAsync as native, closeNativeWorkbookWorker } from '../core/native-runtime.mjs';
after(() => closeNativeWorkbookWorker());
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex'),
  prior = "'O''Brien'!$A$1:$B$4",
  next = "'O''Brien'!$C$3:$E$9";
async function fixture(t, scope = null, { ghost = false } = {}) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'name-admission-')),
    source = path.join(directory, 'source.xlsx'),
    parts = unzipSync(createWorkbook());
  let workbook = strFromU8(parts['xl/workbook.xml'])
    .replace('name="Sheet1"', 'name="O&apos;Brien"')
    .replace(
      '</sheets>',
      `</sheets><definedNames><definedName name="Inputs"${scope === null ? '' : ' localSheetId="' + scope + '"'}>${prior}</definedName></definedNames>`
    );
  if (ghost)
    workbook = workbook.replace(
      '</workbook>',
      '<extLst><ext uri="urn:opaque"><sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Ghost"/></ext></extLst></workbook>'
    );
  parts['xl/workbook.xml'] = strToU8(workbook);
  const bytes = zipSync(parts);
  await fs.writeFile(source, bytes);
  const root = path.join(directory, 'native');
  await fs.mkdir(root, { mode: 0o700 });
  await fs.writeFile(path.join(root, 'source.xlsx'), bytes);
  t.after(async () => {
    await fs.rm(directory, { recursive: true, force: true });
  });
  // Native indexes actual workbook/sheets children. The portable reader already
  // refuses the deliberately opaque ghost as a missing sheet, so bypass that
  // reader here to exercise the independent native plan admission itself.
  await native({ action: 'workbook', operation: 'index', root });
  const plan = {
    kind: 'updateName',
    name: 'Inputs',
    scope,
    priorFormula: prior,
    formula: next,
    replacements: { 'xl/workbook.xml': workbook.replace(prior, next) },
    removed: [],
    copies: {},
    fragments: {},
  };
  const attempt = async (value) => {
    await fs.writeFile(path.join(root, 'package-metadata.json'), JSON.stringify(value));
    return native({ action: 'workbook', operation: 'package-edit', root, version: 0, source: 'source.xlsx' });
  };
  return { directory, source, bytes, plan, attempt, root: () => root };
}
test('native static-name target requires canonical apostrophe quoting', async (t) => {
  const f = await fixture(t),
    valid = await f.attempt(f.plan),
    saved = path.join(f.root(), valid.file),
    accepted = await fs.readFile(saved),
    mutants = [];
  for (const formula of ["'O'Brien'!$C$3:$E$9"]) {
    const plan = structuredClone(f.plan);
    plan.formula = formula;
    plan.replacements['xl/workbook.xml'] = plan.replacements['xl/workbook.xml'].replace(next, formula);
    mutants.push(plan);
  }
  const outcomes = [];
  for (const plan of mutants) {
    let admitted = false,
      error;
    try {
      await f.attempt(plan);
      admitted = true;
    } catch (failure) {
      error = String(failure);
    }
    outcomes.push({
      formula: plan.formula,
      admitted,
      error,
      sourcePreserved: hash(await fs.readFile(f.source)) === hash(f.bytes),
      priorOutputPreserved: hash(await fs.readFile(saved)) === hash(accepted),
    });
  }
  assert(
    outcomes.every(
      (outcome) =>
        !outcome.admitted &&
        /name|quote|scope|worksheet|sheet|target/i.test(outcome.error) &&
        outcome.sourcePreserved &&
        outcome.priorOutputPreserved
    ),
    JSON.stringify(outcomes)
  );
});
test('opaque extension sheet cannot establish a native static-name target', async (t) => {
  const f = await fixture(t, null, { ghost: true }),
    plan = structuredClone(f.plan),
    formula = "'Ghost'!$C$3:$E$9";
  plan.formula = formula;
  plan.replacements['xl/workbook.xml'] = plan.replacements['xl/workbook.xml'].replace(next, formula);
  let admitted = false,
    error;
  try {
    await f.attempt(plan);
    admitted = true;
  } catch (failure) {
    error = String(failure);
  }
  const outcome = { formula, admitted, error, sourcePreserved: hash(await fs.readFile(f.source)) === hash(f.bytes) };
  assert(
    !admitted && /name|worksheet|sheet|reference|target/i.test(error) && outcome.sourcePreserved,
    JSON.stringify(outcome)
  );
});
test('native local-name scope must index a real workbook sheet', async (t) => {
  const f = await fixture(t, 999);
  let admitted = false,
    error;
  try {
    await f.attempt(f.plan);
    admitted = true;
  } catch (failure) {
    error = String(failure);
  }
  const outcome = { scope: 999, admitted, error, sourcePreserved: hash(await fs.readFile(f.source)) === hash(f.bytes) };
  assert(!admitted && /scope|worksheet|sheet|name/i.test(error) && outcome.sourcePreserved, JSON.stringify(outcome));
});
