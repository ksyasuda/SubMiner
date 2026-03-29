import test from 'node:test';
import assert from 'node:assert/strict';
import { composeJellyfinRemoteHandlers } from './jellyfin-remote-composer';

test('composeJellyfinRemoteHandlers returns callable jellyfin remote handlers', async () => {
  let lastProgressAt = 0;
  let activePlayback: unknown = { itemId: 'item-1', mediaSourceId: 'src-1', playMethod: 'DirectPlay', audioStreamIndex: null, subtitleStreamIndex: null };
  const calls: string[] = [];

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
    getActivePlayback: () => activePlayback as never,
    clearActivePlayback: () => {
      activePlayback = null;
      calls.push('clearActivePlayback');
    },
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

  // reportJellyfinRemoteStopped clears active playback when there is no connected session
  await composed.reportJellyfinRemoteStopped();
  assert.equal(activePlayback, null);
  assert.deepEqual(calls, ['clearActivePlayback']);
});
