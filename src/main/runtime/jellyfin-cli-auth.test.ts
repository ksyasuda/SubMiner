import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createHandleJellyfinAuthCommands,
  persistJellyfinAuthSession,
} from './jellyfin-cli-auth';

test('jellyfin auth handler processes logout', async () => {
  const calls: string[] = [];
  const handleAuth = createHandleJellyfinAuthCommands({
    patchRawConfig: () => calls.push('patch'),
    saveStoredSession: () => calls.push('save'),
    clearStoredSession: () => calls.push('clear'),
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
  let patchPayload: unknown = null;
  let storedSession: unknown = null;
  const handleAuth = createHandleJellyfinAuthCommands({
    patchRawConfig: (patch) => {
      patchPayload = patch;
      calls.push('patch');
    },
    saveStoredSession: (session) => {
      storedSession = session;
      calls.push('save');
    },
    clearStoredSession: () => calls.push('clear'),
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
      recentServers: ['http://localhost'],
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
  assert.deepEqual(storedSession, { accessToken: 'token', userId: 'uid' });
  assert.deepEqual(patchPayload, {
    jellyfin: {
      enabled: true,
      serverUrl: 'http://localhost',
      username: 'user',
      deviceId: 'd1',
      clientName: 'SubMiner',
      clientVersion: '1.0',
      recentServers: ['http://localhost'],
    },
  });
  assert.ok(calls.some((entry) => entry.includes('Jellyfin login succeeded')));
});

test('persistJellyfinAuthSession stores client metadata and recent servers', () => {
  let patchPayload: unknown = null;
  let storedSession: unknown = null;

  persistJellyfinAuthSession({
    session: {
      serverUrl: 'http://localhost:8096',
      username: 'alice',
      accessToken: 'token',
      userId: 'uid',
    },
    clientInfo: {
      deviceId: 'device-1',
      clientName: 'SubMiner',
      clientVersion: '1.0',
    },
    existingRecentServers: [
      ' http://old.example:8096/ ',
      'http://localhost:8096',
      '',
      'http://another.example:8096',
    ],
    saveStoredSession: (session) => {
      storedSession = session;
    },
    patchRawConfig: (patch) => {
      patchPayload = patch;
    },
  });

  assert.deepEqual(storedSession, { accessToken: 'token', userId: 'uid' });
  assert.deepEqual(patchPayload, {
    jellyfin: {
      enabled: true,
      serverUrl: 'http://localhost:8096',
      username: 'alice',
      deviceId: 'device-1',
      clientName: 'SubMiner',
      clientVersion: '1.0',
      recentServers: [
        'http://localhost:8096',
        'http://old.example:8096',
        'http://another.example:8096',
      ],
    },
  });
});

test('jellyfin auth handler no-ops when no auth command', async () => {
  const handleAuth = createHandleJellyfinAuthCommands({
    patchRawConfig: () => {},
    saveStoredSession: () => {},
    clearStoredSession: () => {},
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
