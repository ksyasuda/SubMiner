import test from 'node:test';
import assert from 'node:assert/strict';
import { composeJellyfinRemoteHandlers } from './jellyfin-remote-composer';

test('composeJellyfinRemoteHandlers returns callable jellyfin remote handlers', () => {
  let lastProgressAt = 0;
  const composed = composeJellyfinRemoteHandlers({
    getConfiguredSession: () => null,
    getClientInfo: () =>
      ({ clientName: 'SubMiner', clientVersion: 'test', deviceId: 'dev' }) as never,
    getJellyfinConfig: () => ({ enabled: false }) as never,
    playJellyfinItem: async () => {},
    logWarn: () => {},
    getMpvClient: () => null,
    sendMpvCommand: () => {},
    jellyfinTicksToSeconds: () => 0,
    getActivePlayback: () => null,
    clearActivePlayback: () => {},
    getSession: () => null,
    getNow: () => 0,
    getLastProgressAtMs: () => lastProgressAt,
    setLastProgressAtMs: (next) => {
      lastProgressAt = next;
    },
    progressIntervalMs: 3000,
    ticksPerSecond: 10_000_000,
    logDebug: () => {},
  });

  assert.equal(typeof composed.reportJellyfinRemoteProgress, 'function');
  assert.equal(typeof composed.reportJellyfinRemoteStopped, 'function');
  assert.equal(typeof composed.handleJellyfinRemotePlay, 'function');
  assert.equal(typeof composed.handleJellyfinRemotePlaystate, 'function');
  assert.equal(typeof composed.handleJellyfinRemoteGeneralCommand, 'function');
});
