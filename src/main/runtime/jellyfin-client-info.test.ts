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
  });

  assert.equal(getConfig(), jellyfin);
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
