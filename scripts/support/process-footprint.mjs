/** Component-owned macOS diagnostic adapted from the existing host footprint
 * collector. Native aggregate accounting stays separate from process-tree RSS. */
import fs from 'node:fs/promises';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
const exec = promisify(execFile),
  bytes = (value) => (Number.isSafeInteger(value) && value >= 0 ? value : null);

export function footprintTargets(rootPid, sample) {
  if (!Number.isSafeInteger(rootPid) || rootPid < 1 || !sample?.processes?.some((item) => item.pid === rootPid))
    return [];
  const owned = new Set([rootPid]);
  let changed;
  do {
    changed = false;
    for (const item of sample.processes)
      if (owned.has(item.ppid) && Number.isSafeInteger(item.pid) && item.pid > 0 && !owned.has(item.pid)) {
        owned.add(item.pid);
        changed = true;
      }
  } while (changed);
  return [
    ...new Set(
      sample.processes.filter((item) => owned.has(item.pid) && item.executable !== '<defunct>').map((item) => item.pid)
    ),
  ];
}
export function parseFootprint(raw, requestedPids) {
  const requested = new Set(requestedPids),
    errors = [],
    processes = [];
  if (
    !requested.size ||
    requested.size !== requestedPids.length ||
    requestedPids.some((pid) => !Number.isSafeInteger(pid) || pid < 1)
  )
    return {
      status: 'unavailable',
      totalFootprintBytes: null,
      processes,
      errors: ['Invalid or empty requested process set'],
    };
  if (raw?.unit !== 'byte' || raw['bytes per unit'] !== 1 || !Array.isArray(raw.processes))
    return {
      status: 'unavailable',
      totalFootprintBytes: null,
      processes,
      errors: ['Unsupported footprint JSON units or process list'],
    };
  for (const entry of raw.processes) {
    if (!requested.has(entry.pid)) {
      errors.push(`Unexpected footprint process ${entry.pid}`);
      continue;
    }
    if (processes.some((item) => item.pid === entry.pid)) {
      errors.push(`Duplicate footprint process ${entry.pid}`);
      continue;
    }
    const footprintBytes = bytes(entry.auxiliary?.phys_footprint ?? entry.footprint);
    if (footprintBytes === null) {
      errors.push(`Missing physical footprint for ${entry.pid}`);
      continue;
    }
    processes.push({ pid: entry.pid, footprintBytes, peakFootprintBytes: bytes(entry.auxiliary?.phys_footprint_peak) });
  }
  for (const pid of requested)
    if (!processes.some((item) => item.pid === pid)) errors.push(`Footprint did not measure requested process ${pid}`);
  if (Array.isArray(raw.errors))
    errors.push(...raw.errors.map((value) => (typeof value === 'string' ? value : JSON.stringify(value))));
  const total = bytes(raw['total footprint']);
  if (total === null) errors.push('Missing native aggregate physical footprint');
  return {
    status: errors.length ? (processes.length ? 'partial' : 'unavailable') : 'measured',
    totalFootprintBytes: errors.length ? null : total,
    processes,
    errors: errors.slice(0, 64),
    warnings: Array.isArray(raw.warnings) ? raw.warnings.slice(0, 64) : [],
    nativeStartTime: raw.start_time ?? null,
    nativeEndTime: raw.end_time ?? null,
  };
}
export async function captureFootprint({ rootPid, sample, label, outputDirectory }) {
  if (!outputDirectory) return { status: 'disabled', totalFootprintBytes: null };
  const requestedPids = footprintTargets(rootPid, sample),
    result = {
      status: 'unavailable',
      totalFootprintBytes: null,
      requestedPids,
      processes: [],
      errors: [],
      startedAt: new Date().toISOString(),
      sourceSampleAt: sample?.at ?? null,
      method:
        'macOS footprint native aggregate, --noCategories -f bytes, for the sampled app and descendants; separate diagnostic, not RSS.',
    };
  if (process.platform !== 'darwin') return { ...result, errors: ['macOS footprint is unavailable on this platform'] };
  if (!requestedPids.length || requestedPids.length > 64)
    return { ...result, errors: ['Missing task process or diagnostic process-count budget exceeded'] };
  const directory = path.resolve(outputDirectory),
    name =
      String(label)
        .replace(/[^a-zA-Z0-9_-]/g, '_')
        .slice(0, 64) +
      '-' +
      randomUUID();
  try {
    await fs.mkdir(directory, { recursive: true });
  } catch (error) {
    return { ...result, errors: [String(error)] };
  }
  const rawPath = path.join(directory, name + '.json');
  result.rawPath = rawPath;
  let stdout = '',
    stderr = '';
  try {
    const reply = await exec(
      '/usr/bin/footprint',
      [...requestedPids.flatMap((pid) => ['-p', String(pid)]), '-j', rawPath, '--noCategories', '-f', 'bytes'],
      { timeout: 30000, maxBuffer: 256 * 1024 }
    );
    stdout = reply.stdout;
    stderr = reply.stderr;
    result.exitCode = 0;
  } catch (error) {
    stdout = error.stdout ?? '';
    stderr = error.stderr ?? '';
    result.exitCode = Number.isInteger(error.code) ? error.code : null;
    result.errors.push(String(error.message).slice(0, 4000));
  }
  try {
    await fs.writeFile(path.join(directory, name + '.stdout.txt'), stdout);
    await fs.writeFile(path.join(directory, name + '.stderr.txt'), stderr);
    if ((await fs.stat(rawPath)).size > 16 * 1024 ** 2) throw Error('Footprint JSON exceeds diagnostic read budget');
    const parsed = parseFootprint(JSON.parse(await fs.readFile(rawPath, 'utf8')), requestedPids),
      commandErrors = result.errors;
    Object.assign(result, parsed);
    result.errors = [...commandErrors, ...parsed.errors];
    if (result.errors.length && result.status === 'measured') {
      result.status = 'partial';
      result.totalFootprintBytes = null;
    }
  } catch (error) {
    result.errors.push(String(error.message).slice(0, 4000));
  }
  result.finishedAt = new Date().toISOString();
  try {
    await fs.writeFile(path.join(directory, name + '.result.json'), JSON.stringify(result, null, 2) + '\n');
  } catch (error) {
    result.errors.push(String(error));
    if (result.status === 'measured') result.status = 'partial';
    result.totalFootprintBytes = null;
  }
  return result;
}
