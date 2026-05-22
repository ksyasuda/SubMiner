import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createReportJellyfinRemoteProgressHandler,
  createReportJellyfinRemoteStoppedHandler,
  secondsToJellyfinTicks,
} from './jellyfin-remote-playback';

test('secondsToJellyfinTicks converts seconds and clamps invalid values', () => {
  assert.equal(secondsToJellyfinTicks(1.25, 10_000_000), 12_500_000);
  assert.equal(secondsToJellyfinTicks(-3, 10_000_000), 0);
  assert.equal(secondsToJellyfinTicks(Number.NaN, 10_000_000), 0);
});

test('createReportJellyfinRemoteProgressHandler reports playback progress', async () => {
  let lastProgressAtMs = 0;
  const reportPayloads: Array<{ itemId: string; positionTicks: number; isPaused: boolean }> = [];

  const reportProgress = createReportJellyfinRemoteProgressHandler({
    getActivePlayback: () => ({
      itemId: 'item-1',
      mediaSourceId: undefined,
      playMethod: 'DirectPlay',
      audioStreamIndex: 1,
      subtitleStreamIndex: 2,
    }),
    clearActivePlayback: () => {},
    getSession: () => ({
      isConnected: () => true,
      reportProgress: async (payload) => {
        reportPayloads.push({
          itemId: payload.itemId,
          positionTicks: payload.positionTicks,
          isPaused: payload.isPaused,
        });
      },
      reportStopped: async () => {},
    }),
    getMpvClient: () => ({
      requestProperty: async (name: string) => (name === 'time-pos' ? 2.5 : true),
    }),
    getNow: () => 5000,
    getLastProgressAtMs: () => lastProgressAtMs,
    setLastProgressAtMs: (value) => {
      lastProgressAtMs = value;
    },
    progressIntervalMs: 3000,
    ticksPerSecond: 10_000_000,
    logDebug: () => {},
  });

  await reportProgress(true);

  assert.deepEqual(reportPayloads, [
    {
      itemId: 'item-1',
      positionTicks: 25_000_000,
      isPaused: true,
    },
  ]);
  assert.equal(lastProgressAtMs, 5000);
});

test('createReportJellyfinRemoteProgressHandler normalizes mpv pause strings', async () => {
  const reportPayloads: Array<{ isPaused: boolean }> = [];

  const reportProgress = createReportJellyfinRemoteProgressHandler({
    getActivePlayback: () => ({
      itemId: 'item-1',
      playMethod: 'DirectPlay',
    }),
    clearActivePlayback: () => {},
    getSession: () => ({
      isConnected: () => true,
      reportProgress: async (payload) => {
        reportPayloads.push({ isPaused: payload.isPaused });
      },
      reportStopped: async () => {},
    }),
    getMpvClient: () => ({
      requestProperty: async (name: string) => (name === 'pause' ? 'yes' : 3),
    }),
    getNow: () => 5000,
    getLastProgressAtMs: () => 0,
    setLastProgressAtMs: () => {},
    progressIntervalMs: 3000,
    ticksPerSecond: 10_000_000,
    logDebug: () => {},
  });

  await reportProgress(true);

  assert.deepEqual(reportPayloads, [{ isPaused: true }]);
});

test('createReportJellyfinRemoteProgressHandler respects debounce interval', async () => {
  let called = false;
  const reportProgress = createReportJellyfinRemoteProgressHandler({
    getActivePlayback: () => ({
      itemId: 'item-1',
      playMethod: 'DirectPlay',
    }),
    clearActivePlayback: () => {},
    getSession: () => ({
      isConnected: () => true,
      reportProgress: async () => {
        called = true;
      },
      reportStopped: async () => {},
    }),
    getMpvClient: () => ({
      requestProperty: async () => 1,
    }),
    getNow: () => 4000,
    getLastProgressAtMs: () => 3500,
    setLastProgressAtMs: () => {},
    progressIntervalMs: 3000,
    ticksPerSecond: 10_000_000,
    logDebug: () => {},
  });

  await reportProgress(false);
  assert.equal(called, false);
});

test('createReportJellyfinRemoteStoppedHandler reports stop and clears playback', async () => {
  let cleared = false;
  let stoppedItemId: string | null = null;
  const reportStopped = createReportJellyfinRemoteStoppedHandler({
    getActivePlayback: () => ({
      itemId: 'item-2',
      mediaSourceId: undefined,
      playMethod: 'Transcode',
      audioStreamIndex: null,
      subtitleStreamIndex: null,
    }),
    clearActivePlayback: () => {
      cleared = true;
    },
    getSession: () => ({
      isConnected: () => true,
      reportProgress: async () => {},
      reportStopped: async (payload) => {
        stoppedItemId = payload.itemId;
      },
    }),
    logDebug: () => {},
  });

  await reportStopped();
  assert.equal(stoppedItemId, 'item-2');
  assert.equal(cleared, true);
});
