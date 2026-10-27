import test from 'node:test';
import assert from 'node:assert/strict';
import { createStandaloneApi } from '../standalone/bridge.mjs';
import { desktopLaunchInventory } from './support/xlsx-desktop-launch-guard.mjs';

test('standalone bridge preserves document identity, command arguments and event retirement', async () => {
  const calls = [], events = new Map();
  const api = createStandaloneApi({
    invoke: async (channel, ...args) => { calls.push({ channel, args }); return { accepted: true }; },
    subscribe: (channel, callback) => {
      events.set(channel, callback);
      return () => events.delete(channel);
    },
  });
  const command = { operation: 'apply', documentId: 'document', state: 'revision', commands: [{ type: 'undo' }] };
  assert.deepEqual(await api.request('owner', command), { accepted: true });
  assert.deepEqual(calls.pop(), { channel: 'xlsx:request', args: ['owner', command] });
  assert.deepEqual(await api.confirmLeave('owner'), { accepted: true });
  assert.deepEqual(calls.pop(), { channel: 'xlsx:confirm-leave', args: ['owner', false] });
  const observed = [], unsubscribe = api.onProgress(value => observed.push(value));
  const progress = { documentId: 'document', requestId: 'operation', cancellable: true };
  events.get('xlsx:progress')(progress);
  assert.deepEqual(observed, [progress]);
  unsubscribe();
  assert.equal(events.size, 0);
});

test('native owner guard recognizes Tauri and workbook processes without blocking unrelated apps', () => {
  const native = { pid: 1, executable: '/native/XLSX Editor.app/Contents/MacOS/xlsx-desktop' };
  const workbook = { pid: 2, executable: '/native/workbook-engine' };
  const unrelated = { pid: 3, executable: '/other/Electron.app/Contents/MacOS/Electron' };
  assert.deepEqual(desktopLaunchInventory([native, workbook, unrelated]), {
    blockers: [native, workbook], foreignFixtures: [],
  });
});
