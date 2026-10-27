/** Independent test helper only. No production imports or workbook operations. */
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const oracle = fileURLToPath(new URL('./w7-object-oracle.py', import.meta.url));
function call(command, args, { python = process.env.XLSX_ORACLE_PYTHON ?? 'python3', timeout = 30000 } = {}) {
  const result = spawnSync(python, [oracle, command, ...args], {
    encoding: 'utf8',
    timeout,
    maxBuffer: 16 * 1024 * 1024,
  });
  if (result.error) throw result.error;
  let report;
  try {
    report = JSON.parse(result.stdout);
  } catch {
    throw Error(result.stderr || 'Independent object oracle produced no JSON.');
  }
  if (result.status !== 0 || report.passed === false) {
    const failure = Error('Independent W7 object check failed: ' + JSON.stringify(report.failures ?? report));
    failure.report = report;
    throw failure;
  }
  return report;
}
export const inventoryObjectPackage = (file, options) => call('inventory', [file], options);
/** expectationPath is authored by the test, rather than copied from a product response. */
export function checkObjectPackage(before, after, expectationPath, options = {}) {
  const report = call('verify', [before, after, expectationPath], options);
  if (options.decodeMedia) report.mediaDecode = call('decode-media', [after], options);
  return report;
}
export const checkDecodedObjectMedia = (file, options) => call('decode-media', [file], options);
