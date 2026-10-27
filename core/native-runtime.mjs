/** Public Node runtime: standalone framed helper; no app service or UI imports. */
import { existsSync, readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { createWorkbookNativeClient } from './native-client.mjs';
import { refusal } from './refusal-codes.mjs';
export function nativeWorkbookBinary() {
  const name = process.platform === 'win32' ? 'workbook-engine.exe' : 'workbook-engine';
  const root = fileURLToPath(new URL('../', import.meta.url));
  const candidates = [
    process.env.XLSX_EDITOR_NATIVE_BIN,
    path.join(root, 'native/bin', `${process.platform}-${process.arch}`, name),
    path.join(root, 'native/target/release', name),
  ].filter(Boolean);
  const binary = candidates.find(existsSync);
  if (!binary)
    throw refusal(
      'session-unsupported',
      `XLSX native engine unavailable for ${process.platform}-${process.arch}; build or install the matching component runtime.`
    );
  return binary;
}
export function nativeWorkbookIdentity() {
  const file = path.join(path.dirname(nativeWorkbookBinary()), 'identity.json');
  return existsSync(file) ? JSON.parse(readFileSync(file, 'utf8')) : null;
}
let worker;
export function nativeWorkbookOperationAsync(input) {
  worker ??= createWorkbookNativeClient(nativeWorkbookBinary());
  return worker.call(input);
}
export function closeNativeWorkbookWorker() {
  worker?.close();
  worker = undefined;
}
export function nativeWorkbookOperation(input) {
  const request = Buffer.from(JSON.stringify({ id: 1, input }));
  if (request.length > 8 * 1024 * 1024) throw refusal('session-limit', 'Workbook request exceeds transport bound');
  const header = Buffer.alloc(4);
  header.writeUInt32BE(request.length);
  const bytes = execFileSync(nativeWorkbookBinary(), [], {
    input: Buffer.concat([header, request]),
    maxBuffer: 16 * 1024 * 1024 + 4,
    timeout: 120000,
    windowsHide: true,
  });
  if (bytes.length < 4 || bytes.readUInt32BE(0) !== bytes.length - 4)
    throw refusal('internal-invalid-field', 'Invalid workbook reply');
  const reply = JSON.parse(bytes.subarray(4));
  if (reply.id !== 1 || typeof reply.ok !== 'boolean')
    throw refusal('internal-invalid-field', 'Invalid workbook reply');
  if (!reply.ok) throw Object.assign(Error(reply.error), typeof reply.code === 'string' ? { code: reply.code } : {});
  return reply.result;
}
