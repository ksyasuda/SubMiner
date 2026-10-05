import assert from 'node:assert/strict';
import test from 'node:test';
import { createRestartJellyfinRemoteSessionAfterSetupLoginHandler } from './jellyfin-runtime-composer';

test('setup login restart uses auto-connect path without an active remote session', async () => {
  const startOptions: Array<{ explicit?: boolean } | undefined> = [];
  const restart = createRestartJellyfinRemoteSessionAfterSetupLoginHandler({
    getCurrentSession: () => null,
    startJellyfinRemoteSession: async (options) => {
      startOptions.push(options);
    },
  });

  await restart();

  assert.deepEqual(startOptions, [undefined]);
});

test('setup login restart explicitly refreshes an active remote session', async () => {
  const startOptions: Array<{ explicit?: boolean } | undefined> = [];
  const restart = createRestartJellyfinRemoteSessionAfterSetupLoginHandler({
    getCurrentSession: () => ({ stop: () => {} }),
    startJellyfinRemoteSession: async (options) => {
      startOptions.push(options);
    },
  });

  await restart();

  assert.deepEqual(startOptions, [{ explicit: true }]);
});
