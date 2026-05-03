import test from 'node:test';
import assert from 'node:assert/strict';
import { createHandleJellyfinRemoteAnnounceCommand } from './jellyfin-cli-remote-announce';

test('remote announce handler no-ops when flag is disabled', async () => {
  let started = false;
  const handleRemoteAnnounce = createHandleJellyfinRemoteAnnounceCommand({
    startJellyfinRemoteSession: async () => {
      started = true;
    },
    getRemoteSession: () => null,
    logInfo: () => {},
    logWarn: () => {},
  });

  const handled = await handleRemoteAnnounce({
    jellyfinRemoteAnnounce: false,
  } as never);

  assert.equal(handled, false);
  assert.equal(started, false);
});

test('remote announce handler warns when session is unavailable', async () => {
  const warnings: string[] = [];
  let startOptions: { explicit?: boolean } | undefined;
  const handleRemoteAnnounce = createHandleJellyfinRemoteAnnounceCommand({
    startJellyfinRemoteSession: async (options) => {
      startOptions = options;
    },
    getRemoteSession: () => null,
    logInfo: () => {},
    logWarn: (message) => warnings.push(message),
  });

  const handled = await handleRemoteAnnounce({
    jellyfinRemoteAnnounce: true,
  } as never);

  assert.equal(handled, true);
  assert.deepEqual(startOptions, { explicit: true });
  assert.deepEqual(warnings, ['Jellyfin remote session is not available.']);
});

test('remote announce handler reports visibility result', async () => {
  const infos: string[] = [];
  const warnings: string[] = [];
  const handleRemoteAnnounce = createHandleJellyfinRemoteAnnounceCommand({
    startJellyfinRemoteSession: async () => {},
    getRemoteSession: () => ({
      advertiseNow: async () => true,
    }),
    logInfo: (message) => infos.push(message),
    logWarn: (message) => warnings.push(message),
  });

  const handled = await handleRemoteAnnounce({
    jellyfinRemoteAnnounce: true,
  } as never);

  assert.equal(handled, true);
  assert.deepEqual(infos, ['Jellyfin cast target is visible in server sessions.']);
  assert.equal(warnings.length, 0);
});

test('remote announce handler warns when visibility is not confirmed', async () => {
  const warnings: string[] = [];
  const handleRemoteAnnounce = createHandleJellyfinRemoteAnnounceCommand({
    startJellyfinRemoteSession: async () => {},
    getRemoteSession: () => ({
      advertiseNow: async () => false,
    }),
    logInfo: () => {},
    logWarn: (message) => warnings.push(message),
  });

  const handled = await handleRemoteAnnounce({
    jellyfinRemoteAnnounce: true,
  } as never);

  assert.equal(handled, true);
  assert.deepEqual(warnings, [
    'Jellyfin remote announce sent, but cast target is not visible in server sessions yet.',
  ]);
});
