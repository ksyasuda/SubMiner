import assert from 'node:assert/strict';
import test from 'node:test';
import { readMpvInputBindings } from './mpv-input-bindings';

test('discovery reads the connected player and preserves configured keys including disabled bindings', async () => {
  const client = {
    connected: true,
    requestProperty: async (name: string) => {
      assert.equal(name, 'input-bindings');
      return [{ key: 'r', cmd: 'script-binding replay/run', priority: 1 }];
    },
  };
  assert.deepEqual(
    await readMpvInputBindings({
      getMpvClient: () => client,
      getConfiguredKeybindings: () => [{ key: 'Ctrl+KeyR', command: null }],
      platform: 'linux',
    }),
    { keys: ['r'], blockedKeys: [{ code: 'KeyR', modifiers: ['ctrl'] }] },
  );
});

test('discovery safely handles unsupported properties and disconnects during a request', async () => {
  const client = {
    connected: true,
    requestProperty: async (): Promise<unknown> => {
      throw new Error('property unavailable');
    },
  };
  const deps = {
    getMpvClient: () => client,
    getConfiguredKeybindings: () => [],
    platform: 'linux',
  } satisfies Parameters<typeof readMpvInputBindings>[0];
  assert.deepEqual((await readMpvInputBindings(deps)).keys, []);
  client.requestProperty = async () => {
    client.connected = false;
    return [{ key: 'r', cmd: 'seek 5', priority: 1 }];
  };
  assert.deepEqual((await readMpvInputBindings(deps)).keys, []);
});
