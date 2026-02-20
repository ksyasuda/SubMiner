import test from 'node:test';
import assert from 'node:assert/strict';
import { createHandleJellyfinAuthCommands } from './jellyfin-cli-auth';

test('jellyfin auth handler processes logout', async () => {
  const calls: string[] = [];
  const handleAuth = createHandleJellyfinAuthCommands({
    patchRawConfig: () => calls.push('patch'),
    saveStoredToken: () => calls.push('save'),
    clearStoredToken: () => calls.push('clear'),
    authenticateWithPassword: async () => {
      throw new Error('should not authenticate');
    },
    logInfo: (message) => calls.push(message),
  });

  const handled = await handleAuth({
    args: {
      jellyfinLogout: true,
      jellyfinLogin: false,
      jellyfinUsername: undefined,
      jellyfinPassword: undefined,
    } as never,
    jellyfinConfig: {
      serverUrl: '',
      username: '',
      accessToken: '',
      userId: '',
    },
    serverUrl: 'http://localhost',
    clientInfo: {
      deviceId: 'd1',
      clientName: 'SubMiner',
      clientVersion: '1.0',
    },
  });

  assert.equal(handled, true);
  assert.deepEqual(calls.slice(0, 2), ['clear', 'patch']);
});

test('jellyfin auth handler processes login', async () => {
  const calls: string[] = [];
  const handleAuth = createHandleJellyfinAuthCommands({
    patchRawConfig: () => calls.push('patch'),
    saveStoredToken: () => calls.push('save'),
    clearStoredToken: () => calls.push('clear'),
    authenticateWithPassword: async () => ({
      serverUrl: 'http://localhost',
      username: 'user',
      accessToken: 'token',
      userId: 'uid',
    }),
    logInfo: (message) => calls.push(message),
  });

  const handled = await handleAuth({
    args: {
      jellyfinLogout: false,
      jellyfinLogin: true,
      jellyfinUsername: 'user',
      jellyfinPassword: 'pw',
    } as never,
    jellyfinConfig: {
      serverUrl: '',
      username: '',
      accessToken: '',
      userId: '',
    },
    serverUrl: 'http://localhost',
    clientInfo: {
      deviceId: 'd1',
      clientName: 'SubMiner',
      clientVersion: '1.0',
    },
  });

  assert.equal(handled, true);
  assert.ok(calls.includes('save'));
  assert.ok(calls.includes('patch'));
  assert.ok(calls.some((entry) => entry.includes('Jellyfin login succeeded')));
});

test('jellyfin auth handler no-ops when no auth command', async () => {
  const handleAuth = createHandleJellyfinAuthCommands({
    patchRawConfig: () => {},
    saveStoredToken: () => {},
    clearStoredToken: () => {},
    authenticateWithPassword: async () => ({
      serverUrl: '',
      username: '',
      accessToken: '',
      userId: '',
    }),
    logInfo: () => {},
  });

  const handled = await handleAuth({
    args: {
      jellyfinLogout: false,
      jellyfinLogin: false,
      jellyfinUsername: undefined,
      jellyfinPassword: undefined,
    } as never,
    jellyfinConfig: {
      serverUrl: '',
      username: '',
      accessToken: '',
      userId: '',
    },
    serverUrl: 'http://localhost',
    clientInfo: {
      deviceId: 'd1',
      clientName: 'SubMiner',
      clientVersion: '1.0',
    },
  });

  assert.equal(handled, false);
});
