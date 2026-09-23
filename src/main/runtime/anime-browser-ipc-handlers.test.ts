import assert from 'node:assert/strict';
import test from 'node:test';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { registerAnimeBrowserIpcHandlers } from './anime-browser-ipc-handlers';

test('anime browser preference IPC coerces values at the renderer boundary', () => {
  const handlers = new Map<string, (event: unknown, ...args: unknown[]) => unknown>();
  const received: unknown[][] = [];
  registerAnimeBrowserIpcHandlers({
    ipcMain: {
      handle: (channel, listener) => handlers.set(channel, listener),
    },
    runtime: {
      setPreference: (...args: unknown[]) => received.push(args),
    } as never,
  });

  const setPreference = handlers.get(IPC_CHANNELS.request.animeBrowserSetPreference);
  assert.ok(setPreference);
  setPreference({}, 'source', 'text', 'value');
  setPreference({}, 'source', 'enabled', true);
  setPreference({}, 'source', 'choices', ['one', 2, false, 'two']);
  setPreference({}, 'source', 'invalid', { nested: 'value' });

  assert.deepEqual(received, [
    ['source', 'text', 'value'],
    ['source', 'enabled', true],
    ['source', 'choices', ['one', 'two']],
    ['source', 'invalid', ''],
  ]);
});

test('anime browser bulk update IPC returns the runtime result', async () => {
  const handlers = new Map<string, (event: unknown, ...args: unknown[]) => unknown>();
  registerAnimeBrowserIpcHandlers({
    ipcMain: {
      handle: (channel, listener) => handlers.set(channel, listener),
    },
    runtime: {
      updateAllExtensions: async () => 3,
    } as never,
  });

  const updateAll = handlers.get(IPC_CHANNELS.request.animeBrowserUpdateAllExtensions);
  assert.ok(updateAll);
  assert.equal(await updateAll({}), 3);
});
