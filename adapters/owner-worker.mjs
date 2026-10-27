import { parentPort, workerData } from 'node:worker_threads';
import { SpreadsheetDocument } from './document-session.mjs';
import { createWorkbookNativePortClient } from '../core/native-client.mjs';
import { indexedImplementationIdentity } from '../core/indexed-cache.mjs';
import { operationGate } from './operation-gate.mjs';
import { refusal } from '../core/refusal-codes.mjs';
const native = createWorkbookNativePortClient(workerData.nativePort);

let document;
const operations = new Map();
parentPort.on('message', async ({ id, operation, request }) => {
  if (operation === 'cancel-request') {
    const gate = operations.get(request?.requestId),
      result = gate ? gate.authorize() : { granted: false, reason: 'completed' };
    parentPort.postMessage({ id, result });
    return;
  }
  const gate = operationGate(id, (event) => parentPort.postMessage({ event: 'progress', ...event }));
  operations.set(id, gate);
  const options = {
    onProgress: gate.progress,
    beforeCritical: gate.enterCritical,
    leaveCritical: gate.leaveCritical,
    checkCancelled: gate.check,
  };
  try {
    let result;
    if (operation === 'open') {
      if (document) throw refusal('session-refused', 'Document owner has already been opened');
      let cacheIdentity;
      if (workerData.nativeBinary) {
        try {
          cacheIdentity = await indexedImplementationIdentity(workerData.nativeBinary);
        } catch {}
      }
      document = new SpreadsheetDocument({
        recoveryDirectory: workerData.recoveryDirectory,
        maxRecoveryBytes: workerData.maxRecoveryBytes,
        maxIndexCacheBytes: workerData.maxIndexCacheBytes,
        indexDirectory: workerData.indexDirectory,
        native: native.call,
        cacheIdentity,
        initialCalculationPolicy: workerData.initialCalculationPolicy,
        ...options,
      });
      result =
        workerData.newWorkbook !== undefined
          ? await document.create(workerData.newWorkbook.name, { prepared: workerData.prepared })
          : workerData.recoveryId !== undefined
            ? await document.recover(workerData.recoveryId)
            : workerData.resume
              ? await document.resume(workerData.filename)
              : await document.open(workerData.filename, { fresh: workerData.fresh, prepared: workerData.prepared });
    } else if (operation === 'close') {
      await document?.close();
      native.close();
      result = { closed: true };
    } else if (operation === 'request') {
      if (!document) throw refusal('session-closed', 'Document owner is not open');
      result = await document.request(request, options);
    } else throw refusal('session-unsupported', 'Unsupported owner message.');
    parentPort.postMessage({ id, result });
  } catch (error) {
    parentPort.postMessage({
      id,
      error: String(error.message ?? error),
      ...(typeof error.code === 'string' ? { errorCode: error.code } : {}),
    });
  } finally {
    gate.finish();
    operations.delete(id);
  }
});
