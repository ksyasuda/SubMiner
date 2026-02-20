import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createGetJellyfinClientInfoHandler,
  createGetResolvedJellyfinConfigHandler,
} from './jellyfin-client-info';

test('get resolved jellyfin config returns jellyfin section from resolved config', () => {
  const jellyfin = { url: 'https://jellyfin.local' } as never;
  const getConfig = createGetResolvedJellyfinConfigHandler({
    getResolvedConfig: () => ({ jellyfin } as never),
    loadStoredToken: () => null,
  });

  assert.equal(getConfig(), jellyfin);
});

test('get resolved jellyfin config falls back to stored token when config token is blank', () => {
  const getConfig = createGetResolvedJellyfinConfigHandler({
    getResolvedConfig: () =>
      ({
        jellyfin: {
          serverUrl: 'http://localhost:8096',
          accessToken: '   ',
          userId: 'uid-1',
        },
      }) as never,
    loadStoredToken: () => 'stored-token',
  });

  assert.deepEqual(getConfig(), {
    serverUrl: 'http://localhost:8096',
    accessToken: 'stored-token',
    userId: 'uid-1',
  });
});

test('jellyfin client info resolves defaults when fields are missing', () => {
  const getClientInfo = createGetJellyfinClientInfoHandler({
    getResolvedJellyfinConfig: () => ({ clientName: '', clientVersion: '', deviceId: '' } as never),
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
