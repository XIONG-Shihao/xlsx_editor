import { spawnSync } from 'node:child_process';
import { Worker } from 'node:worker_threads';
import { performance } from 'node:perf_hooks';

const INPUT_BYTES = 100 * 1024 * 1024;
const modeIndex = process.argv.indexOf('--case');
const mode = modeIndex >= 0 ? process.argv[modeIndex + 1] : '';

if (mode) {
  if (!['copy', 'transfer'].includes(mode)) throw new Error(`Unknown transfer case: ${mode}`);
  const worker = new Worker(
    `
    const { parentPort } = require('node:worker_threads')
    parentPort.on('message', (bytes) => {
      let checksum = 0
      for (let offset = 0; offset < bytes.byteLength; offset += 4096) checksum ^= bytes[offset]
      setTimeout(() => parentPort.postMessage({ byteLength: bytes.byteLength, checksum }), 300)
    })
  `,
    { eval: true }
  );
  await new Promise((resolve, reject) => {
    worker.once('online', resolve);
    worker.once('error', reject);
  });
  global.gc?.();
  const bytes = new Uint8Array(INPUT_BYTES);
  for (let offset = 0; offset < bytes.byteLength; offset += 4096) bytes[offset] = offset & 0xff;
  const baselineRssBytes = process.memoryUsage().rss;
  let peakRssBytes = baselineRssBytes;
  const started = performance.now();
  const reply = new Promise((resolve, reject) => {
    worker.once('message', resolve);
    worker.once('error', reject);
  });
  worker.postMessage(bytes, mode === 'transfer' ? [bytes.buffer] : []);
  peakRssBytes = Math.max(peakRssBytes, process.memoryUsage().rss);
  const sampler = setInterval(() => {
    peakRssBytes = Math.max(peakRssBytes, process.memoryUsage().rss);
  }, 2);
  const result = await reply;
  clearInterval(sampler);
  peakRssBytes = Math.max(peakRssBytes, process.memoryUsage().rss);
  await worker.terminate();
  process.stdout.write(
    JSON.stringify({
      mode,
      inputBytes: INPUT_BYTES,
      detachedInputBytes: bytes.byteLength,
      resultBytes: result.byteLength,
      elapsedMs: performance.now() - started,
      baselineRssBytes,
      peakRssBytes,
      peakRssDeltaBytes: peakRssBytes - baselineRssBytes,
    })
  );
  process.exit(0);
}

function runCase(name) {
  const run = spawnSync(process.execPath, ['--expose-gc', import.meta.filename, '--case', name], {
    encoding: 'utf8',
    maxBuffer: 1024 * 1024,
  });
  if (run.status !== 0) throw new Error(`${name} benchmark failed: ${run.stderr}`);
  return JSON.parse(run.stdout);
}

const copied = runCase('copy');
const transferred = runCase('transfer');
process.stdout.write(
  JSON.stringify(
    {
      schemaVersion: 1,
      measuredAt: new Date().toISOString(),
      node: process.version,
      platform: `${process.platform}-${process.arch}`,
      description: 'Observational RSS benchmark; not a CI threshold.',
      copied,
      transferred,
      observedPeakRssSavingBytes: copied.peakRssDeltaBytes - transferred.peakRssDeltaBytes,
    },
    null,
    2
  ) + '\n'
);
