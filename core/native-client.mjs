/** Bounded framed protocol. Direct hosts own a child; worker hosts may relay raw
 * bytes through a port while keeping reply parsing in the document worker. */
import { spawn } from 'node:child_process';
import { refusal } from './refusal-codes.mjs';
function protocol({ write, stop, onData, onError }) {
  let sequence = 0,
    closed = false,
    headerUsed = 0,
    length = 0,
    body,
    bodyUsed = 0;
  const header = Buffer.alloc(4),
    pending = new Map();
  const reset = () => {
    headerUsed = 0;
    length = 0;
    body = undefined;
    bodyUsed = 0;
  };
  const close = (error = refusal('session-refused', 'Native workbook client closed')) => {
    if (closed) return;
    closed = true;
    reset();
    for (const item of pending.values()) {
      clearTimeout(item.timer);
      item.reject(error);
    }
    pending.clear();
    stop();
  };
  function reply(bytes) {
    const message = JSON.parse(bytes),
      item = pending.get(message.id);
    if (!item || typeof message.ok !== 'boolean')
      throw refusal('internal-invalid-field', 'Invalid native workbook reply');
    pending.delete(message.id);
    clearTimeout(item.timer);
    if (message.ok) item.resolve(message.result);
    else
      item.reject(Object.assign(Error(message.error), typeof message.code === 'string' ? { code: message.code } : {}));
  }
  onError(close);
  onData((chunk) => {
    if (closed) return;
    let offset = 0;
    try {
      while (offset < chunk.length) {
        if (headerUsed < 4) {
          const count = Math.min(4 - headerUsed, chunk.length - offset);
          chunk.copy(header, headerUsed, offset, offset + count);
          headerUsed += count;
          offset += count;
          if (headerUsed < 4) break;
          length = header.readUInt32BE();
          if (!length || length > 16 * 1024 * 1024)
            throw refusal('internal-invalid-field', 'Invalid native workbook frame');
        }
        if (offset === chunk.length) break;
        // Parse complete payloads directly; copy fragmented payloads into one bounded buffer.
        if (!body && chunk.length - offset >= length) {
          const bytes = chunk.subarray(offset, offset + length);
          offset += length;
          reset();
          reply(bytes);
        } else {
          body ??= Buffer.allocUnsafe(length);
          const count = Math.min(length - bodyUsed, chunk.length - offset);
          chunk.copy(body, bodyUsed, offset, offset + count);
          bodyUsed += count;
          offset += count;
          if (bodyUsed === length) {
            const bytes = body;
            reset();
            reply(bytes);
          }
        }
      }
    } catch (error) {
      close(error);
    }
  });
  const call = (input) =>
    new Promise((resolve, reject) => {
      if (closed) {
        reject(refusal('session-refused', 'Native workbook client closed'));
        return;
      }
      if (pending.size >= 4) {
        reject(refusal('session-refused', 'Native workbook queue full'));
        return;
      }
      let bytes;
      try {
        bytes = Buffer.from(JSON.stringify({ id: sequence + 1, input }));
      } catch (error) {
        reject(error);
        return;
      }
      if (bytes.length > 8 * 1024 * 1024) {
        reject(refusal('session-limit', 'Native workbook request too large'));
        return;
      }
      const id = ++sequence,
        timer = setTimeout(() => close(refusal('operation-timed-out', 'Native workbook deadline exceeded')), 120000);
      pending.set(id, { resolve, reject, timer });
      const header = Buffer.alloc(4);
      header.writeUInt32BE(bytes.length);
      try {
        write(Buffer.concat([header, bytes]));
      } catch (error) {
        close(error);
      }
    });
  return { call, close };
}
export function createWorkbookNativeClient(binary) {
  const child = spawn(binary, [], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
  child.stderr.on('data', () => {});
  return protocol({
    write: (bytes) => child.stdin.write(bytes),
    stop: () => child.kill(),
    onData: (handler) => child.stdout.on('data', handler),
    onError: (handler) => {
      child.on('error', handler);
      child.stdin.on('error', handler);
      child.on('exit', () => handler(refusal('session-closed', 'Native workbook helper stopped')));
    },
  });
}
export function createWorkbookNativePortClient(port) {
  return protocol({
    write: (bytes) => {
      const frame = Uint8Array.from(bytes);
      port.postMessage({ bytes: frame }, [frame.buffer]);
    },
    stop: () => port.close(),
    onData: (handler) =>
      port.on('message', (message) => {
        if (message.bytes)
          handler(Buffer.from(message.bytes.buffer, message.bytes.byteOffset, message.bytes.byteLength));
      }),
    onError: (handler) => {
      port.on('close', () => handler(refusal('session-closed', 'Native workbook helper stopped')));
      port.on('messageerror', handler);
      port.on('message', (message) => {
        if (message.error) handler(Error(message.error));
      });
    },
  });
}
