/** Real framed helper: an oversized read is refused without its content, then the same
 * child keeps serving ordinary reads. */
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { zipSync, unzipSync, strToU8 } from 'fflate';
import { createWorkbook } from '../core/workbook-package.mjs';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
const helper = nativeWorkbookBinary(),
  output = await fs.mkdtemp(path.join(os.tmpdir(), 'xlsx-oversized-reply-'));
const root = path.join(output, 'root');
await fs.mkdir(root);
const privateText = 'PRIVATE-WORKBOOK-CELL:' + 'x'.repeat(20 * 1024),
  parts = unzipSync(createWorkbook());
parts['xl/worksheets/sheet1.xml'] = strToU8(
  `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${Array.from({ length: 1024 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>${privateText}</t></is></c></row>`).join('')}<row r="2001"><c r="A2001"><v>42</v></c></row></sheetData></worksheet>`
);
await fs.writeFile(path.join(root, 'source.xlsx'), zipSync(parts));
const sourceBytes = await fs.readFile(path.join(root, 'source.xlsx'));
let buffer = Buffer.alloc(0),
  nextId = 0,
  stderr = '',
  closed = false;
const pending = new Map(),
  child = spawn(helper, [], { stdio: ['pipe', 'pipe', 'pipe'] });
child.stderr.on('data', (bytes) => {
  stderr += bytes.toString();
});
child.stdout.on('data', (bytes) => {
  buffer = Buffer.concat([buffer, bytes]);
  while (buffer.length >= 4) {
    const size = buffer.readUInt32BE();
    assert(size > 0 && size <= 16 * 1024 * 1024, 'Every framed reply stays bounded');
    if (buffer.length < size + 4) break;
    const raw = Buffer.from(buffer.subarray(4, size + 4)),
      reply = JSON.parse(raw);
    buffer = buffer.subarray(size + 4);
    const waiter = pending.get(reply.id);
    assert(waiter, 'Reply belongs to an issued request');
    pending.delete(reply.id);
    waiter.resolve({ reply, raw });
  }
});
child.once('exit', (code, signal) => {
  closed = true;
  for (const waiter of pending.values()) waiter.reject(Error(`Native child exited (${code}/${signal})`));
  pending.clear();
});
child.once('error', (error) => {
  for (const waiter of pending.values()) waiter.reject(error);
  pending.clear();
});
const call = (input) =>
  new Promise((resolve, reject) => {
    const id = ++nextId,
      bytes = Buffer.from(JSON.stringify({ id, input }));
    assert(bytes.length <= 8 * 1024 * 1024);
    const header = Buffer.alloc(4);
    header.writeUInt32BE(bytes.length);
    pending.set(id, { resolve, reject });
    child.stdin.write(Buffer.concat([header, bytes]));
  });
try {
  const indexed = await call({ action: 'workbook', operation: 'index', root });
  assert(indexed.reply.ok);
  const childPid = child.pid;
  for (let attempt = 0; attempt < 2; attempt++) {
    const refusal = await call({
      action: 'workbook',
      operation: 'range',
      root,
      version: 0,
      sheet: 'xl/worksheets/sheet1.xml',
      range: 'A1:A1024',
      limit: 2048,
    });
    assert.deepEqual(refusal.reply, {
      id: refusal.reply.id,
      ok: false,
      error: 'Native workbook reply exceeds transport bound',
    });
    assert(refusal.raw.length < 256);
    assert(!refusal.raw.includes(Buffer.from('PRIVATE-WORKBOOK-CELL')));
    assert.equal(child.pid, childPid);
    assert(!closed);
    const small = await call({
      action: 'workbook',
      operation: 'range',
      root,
      version: 0,
      sheet: 'xl/worksheets/sheet1.xml',
      range: 'A1',
      limit: 2048,
    });
    assert(small.reply.ok);
    assert.equal(small.reply.result.cells[0][1].value, privateText);
    const ordinary = await call({
      action: 'workbook',
      operation: 'range',
      root,
      version: 0,
      sheet: 'xl/worksheets/sheet1.xml',
      range: 'A2001',
      limit: 2048,
    });
    assert(ordinary.reply.ok);
    assert.equal(ordinary.reply.result.cells[0][1].value, 42);
    assert.equal(child.pid, childPid);
    assert(!closed);
  }
  assert.deepEqual(await fs.readFile(path.join(root, 'source.xlsx')), sourceBytes);
} finally {
  child.stdin.end();
  if (!closed) await new Promise((resolve) => child.once('exit', resolve));
  await fs.rm(output, { recursive: true, force: true });
}
assert.equal(stderr, '');
