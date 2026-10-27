import fs from 'node:fs';
import assert from 'node:assert/strict';
import { refusal } from '../core/refusal-codes.mjs';
// Numeric-only white-box checks of the private protocol's bounded buffers and
// timers. Real helper/port integration is also exercised by worker/session tests.
const text = fs.readFileSync(new URL('../core/native-client.mjs', import.meta.url), 'utf8');
const candidate = text.slice(
  text.indexOf('function protocol('),
  text.indexOf('export function createWorkbookNativeClient')
);
function harness(source, { failWrite = false } = {}) {
  let onData,
    onError,
    stops = 0,
    copied = 0,
    allocations = 0,
    timerId = 0;
  const timers = new Map(),
    writes = [];
  const counted = Object.create(Buffer);
  counted.concat = (items, ...args) => {
    copied += items.reduce((sum, b) => sum + b.length, 0);
    return Buffer.concat(items, ...args);
  };
  counted.allocUnsafe = (n) => {
    allocations += n;
    return Buffer.allocUnsafe(n);
  };
  const fields = 'retained:body?.buffer.byteLength??0,buffered:bodyUsed+headerUsed';
  const measured = source.replaceAll('chunk.copy(', 'copy(chunk,');
  const maker = new Function(
    'Buffer',
    'setTimeout',
    'clearTimeout',
    'copy',
    'refusal',
    (() => {
      const marker = 'return { call, close };';
      if (!measured.includes(marker)) throw Error('The protocol source no longer ends with its expected return');
      return measured.replace(marker, `return {call,close,diagnostic:()=>({${fields},pending:pending.size})};`);
    })() + ';return protocol'
  )(
    counted,
    (fn) => {
      timers.set(++timerId, fn);
      return timerId;
    },
    (id) => timers.delete(id),
    (chunk, target, to, from, end) => {
      copied += end - from;
      return chunk.copy(target, to, from, end);
    },
    refusal
  );
  const p = maker({
    write: (b) => {
      if (failWrite) throw Error('injected write failure');
      writes.push(b);
    },
    stop: () => stops++,
    onData: (f) => (onData = f),
    onError: (f) => (onError = f),
  });
  return {
    p,
    feed: (b) => onData(b),
    error: (e) => onError(e),
    timeout: () => [...timers.values()][0](),
    get stops() {
      return stops;
    },
    get timers() {
      return timers.size;
    },
    get stats() {
      return { ...p.diagnostic(), copied, allocations };
    },
    writes,
  };
}
function frame(value) {
  const body = Buffer.from(typeof value === 'string' ? value : JSON.stringify(value)),
    head = Buffer.alloc(4);
  head.writeUInt32BE(body.length);
  return Buffer.concat([head, body]);
}
async function check(source) {
  const cases = [];
  for (const split of [1, 2, 3, 4, 5, 9, 64]) {
    const h = harness(source),
      p = h.p.call({ test: true }),
      b = frame({ id: 1, ok: true, result: 'Unicode 🧪 中文' });
    for (let i = 0; i < b.length; i += split) h.feed(b.subarray(i, i + split));
    assert.equal(await p, 'Unicode 🧪 中文');
    assert.equal(h.timers, 0);
    assert.equal(h.stats.pending, 0);
    h.p.close();
    assert.equal(h.stops, 1);
  }
  cases.push('partial headers/payloads and split UTF-8');
  {
    const h = harness(source),
      a = h.p.call(1),
      b = h.p.call(2);
    h.feed(Buffer.concat([frame({ id: 2, ok: true, result: 22 }), frame({ id: 1, ok: true, result: 11 })]));
    assert.deepEqual(await Promise.all([a, b]), [11, 22]);
    h.p.close();
  }
  cases.push('coalesced out-of-order IDs');
  {
    const h = harness(source),
      a = h.p.call(1).catch((e) => String(e)),
      b = h.p.call(2);
    h.feed(
      Buffer.concat([frame({ id: 1, ok: false, error: 'expected failure' }), frame({ id: 2, ok: true, result: 22 })])
    );
    assert.match(await a, /expected failure/);
    assert.equal(await b, 22);
    assert.equal(h.stops, 0);
    h.p.close();
  }
  cases.push('application error preserves other requests');
  for (const invalid of [
    Buffer.alloc(4),
    Buffer.from([1, 0, 0, 1]),
    frame('{invalid'),
    frame({ id: 77, ok: true }),
    frame({ id: 1, ok: 'true' }),
  ]) {
    const h = harness(source),
      a = h.p.call(1).catch((e) => e),
      b = h.p.call(2).catch((e) => e);
    h.feed(invalid);
    assert((await a) instanceof Error);
    assert((await b) instanceof Error);
    assert.equal(h.stops, 1);
    assert.equal(h.timers, 0);
    await assert.rejects(h.p.call(3), /closed/);
  }
  cases.push('invalid length/JSON/reply IDs/ok field closes all');
  for (const mode of ['close', 'error', 'timeout']) {
    const h = harness(source),
      p = h.p.call(1).catch((e) => e),
      b = frame({ id: 1, ok: true, result: 'x'.repeat(1024 * 1024) });
    h.feed(b.subarray(0, 16384));
    if (mode === 'close') h.p.close();
    if (mode === 'error') h.error(Error('helper exited'));
    if (mode === 'timeout') h.timeout();
    assert((await p) instanceof Error);
    assert.equal(h.stops, 1);
    assert.equal(h.timers, 0);
    h.feed(b.subarray(16384));
    assert.equal(h.stops, 1);
    if (source === candidate) assert.equal(h.stats.retained, 0);
  }
  cases.push('close/helper exit/deadline during incomplete frame');
  {
    const h = harness(source),
      ps = Array.from({ length: 4 }, (_, i) => h.p.call(i).catch((e) => e));
    await assert.rejects(h.p.call(5), { code: 'session-refused' });
    h.p.close();
    await Promise.all(ps);
  }
  cases.push('queue admission');
  {
    const h = harness(source);
    await assert.rejects(h.p.call('x'.repeat(8 * 1024 * 1024)), { code: /^(session-limit|workbook-limit)$/ });
    const cycle = {};
    cycle.self = cycle;
    await assert.rejects(h.p.call(cycle));
    h.p.close();
  }
  cases.push('request size and serialization failures');
  {
    const h = harness(source, { failWrite: true });
    await assert.rejects(h.p.call(1), /write failure/);
    assert.equal(h.stops, 1);
    assert.equal(h.timers, 0);
  }
  cases.push('write failure');
  return cases;
}
const cases = await check(candidate);
console.log('PASS native client protocol: ' + cases.join('; '));
