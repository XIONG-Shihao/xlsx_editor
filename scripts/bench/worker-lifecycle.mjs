/** Matched worker lifetime/retention check. Baseline must be an installed editor package. */
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import assert from 'node:assert/strict';
import { fork } from 'node:child_process';
import { Worker } from 'node:worker_threads';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import { createWorkbook } from '../../core/workbook-package.mjs';
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
async function ownedStorage() {
  const found = [];
  for (const [name, lease] of [
    ['workbook-indexes-v1', 'owner.json'],
    ['spreadsheet-workers-v1', 'worker-owner.json'],
  ]) {
    const directory = path.join(os.tmpdir(), name + (process.platform === 'linux' ? `-${process.getuid()}` : ''));
    for (const entry of await fs.readdir(directory, { withFileTypes: true }).catch((e) => {
      if (e.code === 'ENOENT') return [];
      throw e;
    }))
      if (entry.isDirectory()) {
        const root = path.join(directory, entry.name);
        try {
          if (JSON.parse(await fs.readFile(path.join(root, lease), 'utf8')).pid === process.pid) found.push(root);
        } catch (e) {
          if (e.code !== 'ENOENT') throw e;
        }
      }
  }
  return found;
}
async function bytesIn(directory) {
  let total = 0;
  for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
    const file = path.join(directory, entry.name);
    if (entry.isDirectory()) total += await bytesIn(file);
    else if (entry.isFile()) total += (await fs.stat(file)).size;
    else throw Error('Unexpected symlink in benchmark-owned storage');
  }
  return total;
}
if (process.argv[2] === 'worker') {
  const [mode, module, filename, recoveryDirectory] = process.argv.slice(3),
    { SpreadsheetWorker } = await import(pathToFileURL(module)),
    baseline = new Set(await ownedStorage());
  const bytes = await fs.readFile(filename),
    sourceSha256 = hash(bytes);
  let owner = new SpreadsheetWorker({ filename, recoveryDirectory });
  try {
    let start = performance.now();
    const status = await owner.open(),
      openMs = performance.now() - start;
    start = performance.now();
    await owner.request({
      operation: 'apply',
      state: status.state,
      commands: [{ sheet: 'Sheet1', cell: 'A1', value: 123 }],
    });
    const editMs = performance.now() - start;
    const original = Worker.prototype.postMessage;
    Worker.prototype.postMessage = function (message, ...args) {
      if (message.request?.simulateWorkerExit) {
        void this.terminate();
        return;
      }
      return original.call(this, message, ...args);
    };
    start = performance.now();
    try {
      await assert.rejects(owner.request({ operation: 'status', simulateWorkerExit: true }), /owner stopped/);
    } finally {
      Worker.prototype.postMessage = original;
    }
    await owner.close();
    const crashCleanupMs = performance.now() - start,
      retained = (await ownedStorage()).filter((root) => !baseline.has(root));
    let retainedBytes = 0;
    for (const root of retained) retainedBytes += await bytesIn(root);
    if (mode === 'current') assert.equal(retainedBytes, 0, 'current worker must clean its disposable native indexes');
    owner = new SpreadsheetWorker({ filename, recoveryDirectory });
    const recovered = await owner.open();
    assert.equal(recovered.recovered, true);
    const range = await owner.request({ operation: 'read', state: recovered.state, sheet: 'Sheet1', range: 'A1' });
    assert.equal(range.cells.A1.value, 123);
    await owner.close();
    assert.equal(hash(await fs.readFile(filename)), sourceSha256);
    // These roots were created by this disposable benchmark process; remove the
    // old implementation's proven residue after recording it, preserving no user files.
    for (const root of retained) await fs.rm(root, { recursive: true, force: true });
    process.send({
      mode,
      openMs,
      editMs,
      crashCleanupMs,
      retainedDirectories: retained.length,
      retainedBytes,
      sourceSha256,
      recoveredValue: 123,
      sourceUnchanged: true,
    });
  } finally {
    await owner.close();
  }
} else {
  const baseline = process.argv[2];
  assert(baseline, 'Pass an absolute installed baseline editor package directory');
  const output = path.resolve(
    process.argv[3] ?? fileURLToPath(new URL('../../.build/runs/bench/worker-lifecycle', import.meta.url))
  );
  await fs.mkdir(output, { recursive: true });
  const filename = path.join(output, 'input.xlsx'),
    parts = unzipSync(createWorkbook()),
    rows = 20000;
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${Array.from({ length: rows }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i + 1}</v></c>${i === 0 ? `<c r="B1"><f>SUM(A1:A${rows})</f></c>` : ''}</row>`).join('')}</sheetData></worksheet>`
  );
  await fs.writeFile(filename, zipSync(parts));
  const runs = [];
  for (const [index, mode] of ['previous', 'current', 'current', 'previous', 'previous', 'current'].entries()) {
    const module =
      mode === 'previous'
        ? path.join(path.resolve(baseline), 'adapters/worker-client.mjs')
        : fileURLToPath(new URL('../../adapters/worker-client.mjs', import.meta.url));
    const child = fork(
      import.meta.filename,
      ['worker', mode, module, filename, path.join(output, 'recovery-' + index)],
      { stdio: ['ignore', 'inherit', 'inherit', 'ipc'], execArgv: [] }
    );
    let result;
    child.on('message', (value) => {
      result = value;
    });
    const timeout = setTimeout(() => child.kill(), 120000);
    const code = await new Promise((resolve, reject) => {
      child.once('exit', resolve);
      child.once('error', reject);
    });
    clearTimeout(timeout);
    assert.equal(code, 0);
    assert(result);
    runs.push(result);
  }
  const medians = Object.fromEntries(
    ['previous', 'current'].map((mode) => [
      mode,
      Object.fromEntries(
        ['openMs', 'editMs', 'crashCleanupMs', 'retainedBytes'].map((key) => [
          key,
          runs
            .filter((r) => r.mode === mode)
            .map((r) => r[key])
            .sort((a, b) => a - b)[1],
        ])
      ),
    ])
  );
  const report = {
    passed: true,
    recordedAt: new Date().toISOString(),
    method:
      'Three alternating fresh Node host processes per version; 20,000 numeric inputs and one SUM; edit A1, abruptly terminate JS worker, await close, inspect own disposable index bytes, recover accepted edit and confirm unchanged source. No Electron or whole-app memory measurement. Filesystem caches uncontrolled.',
    baseline: path.resolve(baseline),
    runs,
    medians,
  };
  await fs.writeFile(path.join(output, 'report.json'), JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify(medians, null, 2));
}
