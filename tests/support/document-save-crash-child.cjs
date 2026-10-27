/** Disposable test host: pause real filesystem operations without production hooks. */
const fs = require('node:fs/promises');
const {syncBuiltinESMExports} = require('node:module');
const {createHash} = require('node:crypto');
const path = require('node:path');
const config = JSON.parse(process.argv[2]);
const original = {open: fs.open, rename: fs.rename, copyFile: fs.copyFile, readFile: fs.readFile};
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
let armed = false, hit = false, exported, staged, manifest, receipt, published = false, accepted, resume;
process.on('message', message => {if (message.operation === 'continue') resume?.();});
const send = message => new Promise((resolve, reject) => process.send(message, error => error ? reject(error) : resolve()));

async function checkpoint(boundary) {
  if (!armed || hit || config.boundary !== boundary) return;
  hit = true;
  const continued = new Promise(resolve => {resume = resolve;});
  await send({type: 'checkpoint', boundary, exported, staged, manifest, accepted,
    candidateHash: exported ? hash(await original.readFile(exported)) : null,
    sourceHash: hash(await original.readFile(config.filename))});
  await continued;
  if (config.action === 'fail') {
    const error = Error(`Injected ${config.errorCode ?? 'ENOSPC'} at ${boundary}`);
    error.code = config.errorCode ?? 'ENOSPC';
    throw error;
  }
}

function isStage(filename) {
  return path.dirname(filename) === path.dirname(config.filename) &&
    path.basename(filename).startsWith('.' + path.basename(config.filename) + '.') && filename.endsWith('.tmp');
}

fs.copyFile = async function(source, destination, ...options) {
  if (armed && isStage(destination)) {
    exported = source;
    staged = destination;
    await checkpoint('export-complete');
    await checkpoint('before-stage-copy');
  }
  return original.copyFile(source, destination, ...options);
};

fs.open = async function(filename, ...options) {
  const handle = await original.open(filename, ...options);
  const sync = handle.sync.bind(handle);
  handle.sync = async function() {
    const result = await sync();
    if (!armed) return result;
    if (isStage(filename)) await checkpoint('stage-file-sync');
    if (path.basename(filename).startsWith('current.json.') && filename.endsWith('.tmp')) {
      const record = JSON.parse(await original.readFile(filename, 'utf8'));
      if (record.head === accepted.state) {
        if (record.pendingPublication) await checkpoint('pending-manifest-file-sync');
        else if (record.savedHead === accepted.state) await checkpoint('final-manifest-file-sync');
      }
    }
    if (manifest && filename === path.dirname(manifest)) {
      if (receipt === 'pending') await checkpoint('pending-receipt-sync');
      if (receipt === 'final') await checkpoint('final-manifest-sync');
    }
    if (published && filename === path.dirname(config.filename)) await checkpoint('publication-directory-sync');
    return result;
  };
  return handle;
};

fs.rename = async function(source, destination) {
  if (armed && destination === config.filename) await checkpoint('before-source-rename');
  const result = await original.rename(source, destination);
  if (!armed) return result;
  if (destination === config.filename) {
    published = true;
    await checkpoint('source-rename');
  } else if (path.basename(destination) === 'current.json') {
    const record = JSON.parse(await original.readFile(destination, 'utf8'));
    if (record.head === accepted.state) {
      manifest = destination;
      receipt = record.pendingPublication ? 'pending' : 'final';
    }
  }
  return result;
};

// Refresh ESM named imports before loading the unchanged document/storage code.
syncBuiltinESMExports();
(async () => {
  const {SpreadsheetDocument} = await import('../../adapters/document-session.mjs');
  const {nativeWorkbookOperationAsync, closeNativeWorkbookWorker} = await import('../../core/native-runtime.mjs');
  const owner = new SpreadsheetDocument({recoveryDirectory: config.recoveryDirectory, indexDirectory: config.indexDirectory,
    native: input => nativeWorkbookOperationAsync(input)});
  try {
    await owner.open(config.filename);
    for (const commands of [[{sheet: 'Sheet1', cell: 'A1', value: 7}], [{sheet: 'Sheet1', cell: 'B1', value: 4}]]) {
      await owner.request({operation: 'apply', state: owner.status().state, commands});
    }
    const {state, dirty, canUndo} = owner.status();
    accepted = {state, dirty, canUndo};
    armed = true;
    let error;
    try {await owner.request({operation: 'save'});} catch (failure) {error = {message: failure.message, code: failure.code};}
    const status = owner.status();
    await send({type: 'complete', hit, error, state: status.state, dirty: status.dirty, canUndo: status.canUndo});
  } finally {
    armed = false;
    await owner.close();
    closeNativeWorkbookWorker();
    process.disconnect();
  }
})().catch(async error => {
  if (process.connected) await send({type: 'fatal', error: String(error.stack ?? error)}).catch(() => {});
  process.exitCode = 1;
  if (process.connected) process.disconnect();
});
