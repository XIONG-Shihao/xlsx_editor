import { parentPort, workerData } from 'node:worker_threads';
import { CsvFileSession } from '../core/indexed-csv-session.mjs';
import { refusal } from '../core/refusal-codes.mjs';

let session,
  queue = Promise.resolve();
const methods = new Set(['read', 'scan', 'query', 'selectionSummary', 'apply', 'compact', 'export', 'close']);
function diagnostics() {
  const stats = session?.diagnosticStats() ?? {};
  if (workerData?.diagnostics) {
    const memory = process.memoryUsage();
    // Worker heaps are thread-local; RSS belongs to the whole process and must
    // not be added once per CSV worker.
    Object.assign(stats, {
      workerHeapUsed: memory.heapUsed,
      workerHeapTotal: memory.heapTotal,
      workerExternal: memory.external,
      workerArrayBuffers: memory.arrayBuffers,
    });
  }
  return stats;
}
parentPort.on('message', ({ id, method, args }) => {
  queue = queue
    .then(async () => {
      try {
        let result;
        if (method === 'open') {
          if (session) throw refusal('csv-refused', 'CSV owner is already open');
          session = new CsvFileSession();
          result = await session.open(...args);
        } else if (method === 'close' && !session) {
          result = { closed: true };
        } else {
          if (!session || !methods.has(method)) throw refusal('csv-unsupported', 'Unsupported CSV owner request');
          result = await session[method](...args);
        }
        parentPort.postMessage({ id, result, stats: diagnostics() });
      } catch (error) {
        parentPort.postMessage({ id, error: String(error.message ?? error), stats: diagnostics() });
      }
    })
    .catch((error) => {
      parentPort.postMessage({ id, error: String(error.message ?? error) });
    });
});
