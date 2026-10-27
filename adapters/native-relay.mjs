/** Own/reap the helper in the host; transfer raw frames, never parse cell graphs. */
import { spawn } from 'node:child_process';
import { nativeWorkbookBinary } from '../core/native-runtime.mjs';
import { refusal } from '../core/refusal-codes.mjs';
const reapedProofs = new WeakMap();
/** Host-private, unforgeable proof bound to one allocated worker storage root. */
export function knownReapedNativePid(proof, root) {
  const record = proof && typeof proof === 'object' ? reapedProofs.get(proof) : undefined;
  return record && record.root === root ? record.pid : undefined;
}
export function createNativeRelay(port, binary, { storageRoot } = {}) {
  let child,
    finished = Promise.resolve(),
    closing,
    closed = false;
  const report = (error) => {
    try {
      port.postMessage({ error: String(error.message ?? error) });
    } catch {}
  };
  function start() {
    child = spawn(binary ?? nativeWorkbookBinary(), [], { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
    finished = new Promise((resolve) => child.once('close', resolve));
    child.on('error', (error) => {
      report(error);
      void close();
    });
    child.stdin.on('error', (error) => {
      report(error);
      void close();
    });
    child.stderr.on('data', () => {});
    child.stdout.on('data', (chunk) => {
      if (closed) return;
      const bytes = Uint8Array.from(chunk);
      try {
        port.postMessage({ bytes }, [bytes.buffer]);
      } catch {
        void close();
      }
    });
    child.once('exit', () => {
      report(refusal('session-closed', 'Native workbook helper stopped'));
      port.close();
    });
  }
  function close() {
    if (closing) return closing;
    closed = true;
    closing = (async () => {
      port.close();
      if (!child) return { reapedNativePid: null };
      const pid = Number.isSafeInteger(child.pid) && child.pid > 0 ? child.pid : null;
      child.kill();
      const deadline = setTimeout(() => child.kill('SIGKILL'), 2000);
      try {
        await finished;
        const proof = Object.freeze({ reapedNativePid: pid });
        if (pid !== null && typeof storageRoot === 'string') reapedProofs.set(proof, { pid, root: storageRoot });
        return proof;
      } finally {
        clearTimeout(deadline);
      }
    })();
    return closing;
  }
  port.on('message', (message) => {
    if (closed) return;
    try {
      if (
        !(message.bytes instanceof Uint8Array) ||
        message.bytes.byteLength < 4 ||
        message.bytes.byteLength > 8 * 1024 * 1024 + 4
      )
        throw refusal('internal-invalid-field', 'Invalid relayed workbook frame');
      if (!child) start();
      child.stdin.write(message.bytes);
    } catch (error) {
      report(error);
      void close();
    }
  });
  port.on('close', () => {
    void close();
  });
  port.on('messageerror', (error) => {
    report(error);
    void close();
  });
  return { close };
}
