import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import assert from 'node:assert/strict';
import { planCleanup, applyCleanup } from '../scripts/clean-artifacts.mjs';
const root = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), 'xlsx-artifact-cleanup-')));
const put = (file, value = 'keep') => {
  const p = path.join(root, file);
  fs.mkdirSync(path.dirname(p), { recursive: true });
  fs.writeFileSync(p, value);
};
try {
  put('artifacts/checkpoint-1/mac-arm64/XLSX Editor.app/app', 'old');
  put('artifacts/checkpoint-3/mac-arm64/XLSX Editor.app/app');
  put(
    'artifacts/current.json',
    JSON.stringify({
      app: path.join(root, 'artifacts/checkpoint-3/mac-arm64/XLSX Editor.app'),
      asarSha256: 'unchanged',
    })
  );
  put('artifacts/final/library.tgz');
  put('artifacts/bench/baseline/source.mjs');
  put('artifacts/npm-cache/item');
  put('artifacts/engine-package/staged');
  for (const i of [1, 2]) {
    put(`artifacts/large-desktop-${i}/book.xlsx`);
    put(`artifacts/large-desktop-${i}/profile/cache`);
    put(
      `artifacts/large-desktop-${i}/report.json`,
      JSON.stringify({
        passed: true,
        cases: [
          { source: path.join(root, 'original.xlsx'), file: path.join(root, `artifacts/large-desktop-${i}/book.xlsx`) },
        ],
      })
    );
  }
  put('original.xlsx', 'original');
  const plan = planCleanup(root);
  assert.throws(() => applyCleanup(plan, [`app ${root}/artifacts/checkpoint-3`]), /active process/);
  assert(fs.existsSync(path.join(root, 'artifacts/checkpoint-1')));
  const result = applyCleanup(plan, []);
  assert.equal(fs.readFileSync(path.join(root, 'original.xlsx'), 'utf8'), 'original');
  for (const file of [
    'artifacts/releases/checkpoint-3/mac-arm64/XLSX Editor.app/app',
    'artifacts/packages/library.tgz',
    '.build/runs/large-desktop-2/book.xlsx',
    '.build/runs/bench/baseline/source.mjs',
    '.build/cache/npm/item',
    '.build/staging/engine-package/staged',
  ])
    assert.equal(fs.readFileSync(path.join(root, file), 'utf8'), 'keep');
  assert(!fs.existsSync(path.join(root, '.build/runs/large-desktop-1/book.xlsx')));
  for (const i of [1, 2]) assert(!fs.existsSync(path.join(root, `.build/runs/large-desktop-${i}/profile`)));
  assert(fs.existsSync(path.join(root, '.build/runs/large-desktop-1/report.json')));
  assert.equal(JSON.parse(fs.readFileSync(path.join(root, 'artifacts/current.json'))).app, result.currentApp);
  const ledger = fs.readFileSync(path.join(root, '.build/artifact-cleanup.json'), 'utf8');
  const again = planCleanup(root);
  assert.equal(again.moves.length, 0);
  assert.equal(again.removals.length, 0);
  applyCleanup(again, []);
  assert.equal(fs.readFileSync(path.join(root, '.build/artifact-cleanup.json'), 'utf8'), ledger);
  fs.symlinkSync(os.tmpdir(), path.join(root, 'artifacts/checkpoint-9'));
  assert.throws(() => planCleanup(root), /symlink/);
  fs.unlinkSync(path.join(root, 'artifacts/checkpoint-9'));
  put('artifacts/checkpoint-9/mac-arm64/XLSX Editor.app/app', 'old');
  const changed = planCleanup(root);
  put('artifacts/current.json', JSON.stringify({ app: result.currentApp, changed: true }));
  assert.throws(() => applyCleanup(changed, []), /Changed after preview/);
  assert(fs.existsSync(path.join(root, 'artifacts/checkpoint-9')));
  console.log(
    'PASS qualified app/package/latest workbook/report preservation, source protection, process/symlink/changed-marker refusal and idempotence'
  );
} finally {
  fs.rmSync(root, { recursive: true, force: true });
}
