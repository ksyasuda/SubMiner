import assert from 'node:assert/strict';
import test from 'node:test';
import type { ActiveJellyfinRemotePlaybackState } from '../jellyfin-remote-commands';
import { composeJellyfinRemoteHandlers } from './jellyfin-remote-composer';

function makeHandlers() {
  return composeJellyfinRemoteHandlers({
    getConfiguredSession: () => null,
    getClientInfo: () => ({ clientName: 'SubMiner', clientVersion: '1', deviceId: 'did' }),
    getJellyfinConfig: () => ({}),
    playJellyfinItem: async () => {},
    logWarn: () => {},
    getMpvClient: () => null,
    sendMpvCommand: () => {},
    jellyfinTicksToSeconds: (ticks) => ticks / 10_000_000,
    getActivePlayback: () => null,
    clearActivePlayback: () => {},
    getNow: () => 0,
    getLastProgressAtMs: () => 0,
    setLastProgressAtMs: () => {},
    progressIntervalMs: 3000,
    ticksPerSecond: 10_000_000,
    logDebug: () => {},
  });
}

test('a tracked episode start holds its own stop and the next transition', async () => {
  const handlers = makeHandlers();
  const events: string[] = [];
  const playback: ActiveJellyfinRemotePlaybackState = {
    itemId: 'ep1',
    playMethod: 'DirectPlay',
    reporter: {
      reportPlaying: async () => {},
      reportProgress: async () => {},
      reportStopped: async () => void events.push('stopped:ep1'),
    },
  };
  let finishStart = () => {};
  handlers.trackJellyfinPlaybackStart(
    new Promise<void>((resolve) => {
      finishStart = () => {
        events.push('started:ep1');
        resolve();
      };
    }),
    playback,
  );

  const stopped = handlers.reportJellyfinPlaybackSuperseded(playback);
  const transitionsDone = handlers.waitForJellyfinTransitionReports().then(() => {
    events.push('transitions settled');
  });
  await Promise.resolve();
  assert.deepEqual(events, [], 'the stop and the next start wait for the start in flight');

  finishStart();
  await Promise.all([stopped, transitionsDone]);
  assert.deepEqual(events, ['started:ep1', 'stopped:ep1', 'transitions settled']);
});
