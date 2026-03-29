import assert from 'node:assert/strict';
import test from 'node:test';

import { createDiscordRpcClient } from './discord-rpc-client';

test('createDiscordRpcClient forwards rich presence calls through client.user', async () => {
  const calls: Array<string> = [];
  const rpcClient = createDiscordRpcClient('123456789012345678', {
    createClient: () =>
      ({
        login: async () => {
          calls.push('login');
        },
        user: {
          setActivity: async () => {
            calls.push('setActivity');
          },
          clearActivity: async () => {
            calls.push('clearActivity');
          },
        },
        destroy: async () => {
          calls.push('destroy');
        },
      }) as never,
  });

  await rpcClient.login();
  await rpcClient.setActivity({
    details: 'Title',
    state: 'Playing 00:01 / 00:02',
    startTimestamp: 1_700_000_000,
  });
  await rpcClient.clearActivity();
  await rpcClient.destroy();

  assert.deepEqual(calls, ['login', 'setActivity', 'clearActivity', 'destroy']);
});
