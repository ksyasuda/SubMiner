import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createGetJellyfinClientInfoHandler,
  createGetResolvedJellyfinConfigHandler,
} from './jellyfin-client-info';

test('get resolved jellyfin config returns jellyfin section from resolved config', () => {
  const jellyfin = { url: 'https://jellyfin.local' } as never;
  const getConfig = createGetResolvedJellyfinConfigHandler({
    getResolvedConfig: () => ({ jellyfin }) as never,
    loadStoredSession: () => null,
    getEnv: () => undefined,
  });

  assert.equal(getConfig(), jellyfin);
});

test('get resolved jellyfin config falls back to stored session when env is unset', () => {
  const getConfig = createGetResolvedJellyfinConfigHandler({
    getResolvedConfig: () =>
      ({
        jellyfin: {
          serverUrl: 'http://localhost:8096',
        },
      }) as never,
    loadStoredSession: () => ({ accessToken: 'stored-token', userId: 'uid-1' }),
    getEnv: () => undefined,
  });

  assert.deepEqual(getConfig(), {
    serverUrl: 'http://localhost:8096',
    accessToken: 'stored-token',
    userId: 'uid-1',
  });
});

test('get resolved jellyfin config prefers env token and env user id over stored session', () => {
  const getConfig = createGetResolvedJellyfinConfigHandler({
    getResolvedConfig: () =>
      ({
        jellyfin: {
          serverUrl: 'http://localhost:8096',
        },
      }) as never,
    loadStoredSession: () => ({ accessToken: 'stored-token', userId: 'stored-user' }),
    getEnv: (key: string) =>
      key === 'SUBMINER_JELLYFIN_ACCESS_TOKEN'
        ? 'env-token'
        : key === 'SUBMINER_JELLYFIN_USER_ID'
          ? 'env-user'
          : undefined,
  });

  assert.deepEqual(getConfig(), {
    serverUrl: 'http://localhost:8096',
    accessToken: 'env-token',
    userId: 'env-user',
  });
});

test('get resolved jellyfin config uses stored user id when env token set without env user id', () => {
  const getConfig = createGetResolvedJellyfinConfigHandler({
    getResolvedConfig: () =>
      ({
        jellyfin: {
          serverUrl: 'http://localhost:8096',
        },
      }) as never,
    loadStoredSession: () => ({ accessToken: 'stored-token', userId: 'stored-user' }),
    getEnv: (key: string) => (key === 'SUBMINER_JELLYFIN_ACCESS_TOKEN' ? 'env-token' : undefined),
  });

  assert.deepEqual(getConfig(), {
    serverUrl: 'http://localhost:8096',
    accessToken: 'env-token',
    userId: 'stored-user',
  });
});

test('jellyfin client info resolves defaults when fields are missing', () => {
  const getClientInfo = createGetJellyfinClientInfoHandler({
    getResolvedJellyfinConfig: () => ({ clientName: '', clientVersion: '', deviceId: '' }) as never,
    getDefaultJellyfinConfig: () =>
      ({
        clientName: 'SubMiner',
        clientVersion: '1.0.0',
        deviceId: 'default-device',
      }) as never,
  });

  assert.deepEqual(getClientInfo(), {
    clientName: 'SubMiner',
    clientVersion: '1.0.0',
    deviceId: 'default-device',
  });
});

test('jellyfin client info keeps explicit config values', () => {
  const getClientInfo = createGetJellyfinClientInfoHandler({
    getResolvedJellyfinConfig: () =>
      ({
        clientName: 'Custom',
        clientVersion: '2.3.4',
        deviceId: 'custom-device',
      }) as never,
    getDefaultJellyfinConfig: () =>
      ({
        clientName: 'SubMiner',
        clientVersion: '1.0.0',
        deviceId: 'default-device',
      }) as never,
  });

  assert.deepEqual(getClientInfo(), {
    clientName: 'Custom',
    clientVersion: '2.3.4',
    deviceId: 'custom-device',
  });
});
