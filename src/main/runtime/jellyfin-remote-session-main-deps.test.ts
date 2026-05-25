import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildStartJellyfinRemoteSessionMainDepsHandler,
  createBuildStopJellyfinRemoteSessionMainDepsHandler,
} from './jellyfin-remote-session-main-deps';

test('start jellyfin remote session main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const session = { start: () => {}, stop: () => {}, advertiseNow: async () => true };
  const deps = createBuildStartJellyfinRemoteSessionMainDepsHandler({
    getJellyfinConfig: () => ({ serverUrl: 'http://localhost' }) as never,
    getCurrentSession: () => null,
    setCurrentSession: () => calls.push('set-session'),
    createRemoteSessionService: () => session as never,
    getClientInfo: () =>
      ({
        deviceId: 'workstation',
        clientName: 'SubMiner',
        clientVersion: '1.0',
      }) as never,
    getHostName: () => 'workstation',
    defaultDeviceId: 'device',
    defaultClientName: 'SubMiner',
    defaultClientVersion: '1.0',
    handlePlay: async () => {
      calls.push('play');
    },
    handlePlaystate: async () => {
      calls.push('playstate');
    },
    handleGeneralCommand: async () => {
      calls.push('general');
    },
    logInfo: (message) => calls.push(`info:${message}`),
    logWarn: (message) => calls.push(`warn:${message}`),
    onSessionStateChanged: () => calls.push('state-changed'),
  })();

  assert.deepEqual(deps.getJellyfinConfig(), { serverUrl: 'http://localhost' });
  assert.equal(deps.defaultDeviceId, 'device');
  assert.equal(deps.defaultClientName, 'SubMiner');
  assert.equal(deps.defaultClientVersion, '1.0');
  assert.equal(deps.getHostName(), 'workstation');
  assert.deepEqual(deps.getClientInfo(), {
    deviceId: 'workstation',
    clientName: 'SubMiner',
    clientVersion: '1.0',
  });
  assert.equal(deps.createRemoteSessionService({} as never), session);
  await deps.handlePlay({});
  await deps.handlePlaystate({});
  await deps.handleGeneralCommand({});
  deps.logInfo('connected');
  deps.logWarn('missing');
  deps.onSessionStateChanged?.();
  assert.deepEqual(calls, [
    'play',
    'playstate',
    'general',
    'info:connected',
    'warn:missing',
    'state-changed',
  ]);
});

test('stop jellyfin remote session main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const session = { start: () => {}, stop: () => {}, advertiseNow: async () => true };
  const deps = createBuildStopJellyfinRemoteSessionMainDepsHandler({
    getCurrentSession: () => session as never,
    setCurrentSession: () => calls.push('set-null'),
    clearActivePlayback: () => calls.push('clear'),
    onSessionStateChanged: () => calls.push('state-changed'),
  })();

  assert.equal(deps.getCurrentSession(), session);
  deps.setCurrentSession(null);
  deps.clearActivePlayback();
  deps.onSessionStateChanged?.();
  assert.deepEqual(calls, ['set-null', 'clear', 'state-changed']);
});
