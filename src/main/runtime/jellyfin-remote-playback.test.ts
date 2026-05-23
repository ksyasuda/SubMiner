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

test('createReportJellyfinRemoteProgressHandler reports while remote websocket is disconnected', async () => {
  const reportPayloads: Array<{ positionTicks: number; isPaused: boolean }> = [];

  const reportProgress = createReportJellyfinRemoteProgressHandler({
    getActivePlayback: () => ({
      itemId: 'item-1',
      playMethod: 'DirectPlay',
    }),
    clearActivePlayback: () => {},
    getSession: () => ({
      isConnected: () => false,
      reportProgress: async (payload) => {
        reportPayloads.push({
          positionTicks: payload.positionTicks,
          isPaused: payload.isPaused,
        });
      },
      reportStopped: async () => {},
    }),
    getMpvClient: () => ({
      currentTimePos: 42,
      requestProperty: async (name: string) => (name === 'pause' ? false : 42),
    }),
    getNow: () => 5000,
    getLastProgressAtMs: () => 0,
    setLastProgressAtMs: () => {},
    progressIntervalMs: 3000,
    ticksPerSecond: 10_000_000,
    logDebug: () => {},
  });

  await reportProgress(true);

  assert.deepEqual(reportPayloads, [{ positionTicks: 420_000_000, isPaused: false }]);
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

test('createReportJellyfinRemoteProgressHandler reports mpv seek jumps during debounce', async () => {
  let now = 5000;
  let lastProgressAtMs = 0;
  let position = 10;
  const reportPayloads: Array<{ positionTicks: number; eventName: string }> = [];

  const reportProgress = createReportJellyfinRemoteProgressHandler({
    getActivePlayback: () => ({
      itemId: 'item-1',
      playMethod: 'DirectPlay',
    }),
    clearActivePlayback: () => {},
    getSession: () => ({
      isConnected: () => true,
      reportProgress: async (payload) => {
        reportPayloads.push({
          positionTicks: payload.positionTicks,
          eventName: payload.eventName,
        });
      },
      reportStopped: async () => {},
    }),
    getMpvClient: () => ({
      currentTimePos: position,
      requestProperty: async (name: string) => (name === 'pause' ? false : position),
    }),
    getNow: () => now,
    getLastProgressAtMs: () => lastProgressAtMs,
    setLastProgressAtMs: (value) => {
      lastProgressAtMs = value;
    },
    progressIntervalMs: 3000,
    ticksPerSecond: 10_000_000,
    logDebug: () => {},
  });

  await reportProgress(true);
  now = 5500;
  position = 90;
  await reportProgress(false);

  assert.deepEqual(reportPayloads, [
    { positionTicks: 100_000_000, eventName: 'TimeUpdate' },
    { positionTicks: 900_000_000, eventName: 'TimeUpdate' },
  ]);
  assert.equal(lastProgressAtMs, 5500);
});

test('createReportJellyfinRemoteStoppedHandler reports stop and clears playback', async () => {
  let cleared = false;
  let stoppedPayload: {
    itemId: string;
    positionTicks?: number;
    failed?: boolean;
  } | null = null;
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
        stoppedPayload = {
          itemId: payload.itemId,
          positionTicks: payload.positionTicks,
          failed: payload.failed,
        };
      },
    }),
    getMpvClient: () => ({
      currentTimePos: 12.5,
      requestProperty: async () => {
        throw new Error('unloaded');
      },
    }),
    ticksPerSecond: 10_000_000,
    logDebug: () => {},
  });

  await reportStopped();
  assert.deepEqual(stoppedPayload, {
    itemId: 'item-2',
    positionTicks: 125_000_000,
    failed: false,
  });
  assert.equal(cleared, true);
});

test('createReportJellyfinRemoteStoppedHandler reports stop while remote websocket is disconnected', async () => {
  let cleared = false;
  let stoppedPayload: {
    itemId: string;
    positionTicks?: number;
    failed?: boolean;
  } | null = null;
  const reportStopped = createReportJellyfinRemoteStoppedHandler({
    getActivePlayback: () => ({
      itemId: 'item-2',
      mediaSourceId: undefined,
      playMethod: 'Transcode',
      audioStreamIndex: null,
      subtitleStreamIndex: null,
      loadedMediaPath: 'https://stream.example/video.m3u8',
    }),
    clearActivePlayback: () => {
      cleared = true;
    },
    getSession: () => ({
      isConnected: () => false,
      reportProgress: async () => {},
      reportStopped: async (payload) => {
        stoppedPayload = {
          itemId: payload.itemId,
          positionTicks: payload.positionTicks,
          failed: payload.failed,
        };
      },
    }),
    getMpvClient: () => ({
      currentTimePos: 12.5,
    }),
    ticksPerSecond: 10_000_000,
    logDebug: () => {},
  });

  await reportStopped();

  assert.deepEqual(stoppedPayload, {
    itemId: 'item-2',
    positionTicks: 125_000_000,
    failed: false,
  });
  assert.equal(cleared, true);
});

test('createReportJellyfinRemoteStoppedHandler ignores unloaded active playback', async () => {
  let cleared = false;
  let stopped = false;
  const reportStopped = createReportJellyfinRemoteStoppedHandler({
    getActivePlayback: () =>
      ({
        itemId: 'item-2',
        playMethod: 'Transcode',
        loadedMediaPath: null,
      }) as never,
    clearActivePlayback: () => {
      cleared = true;
    },
    getSession: () => ({
      isConnected: () => true,
      reportProgress: async () => {},
      reportStopped: async () => {
        stopped = true;
      },
    }),
    getMpvClient: () => ({
      currentTimePos: 0,
    }),
    ticksPerSecond: 10_000_000,
    logDebug: () => {},
  });

  await reportStopped();

  assert.equal(stopped, false);
  assert.equal(cleared, false);
});

test('createReportJellyfinRemoteStoppedHandler ignores startup stop churn before grace expires', async () => {
  let cleared = false;
  let stopped = false;
  const reportStopped = createReportJellyfinRemoteStoppedHandler({
    getActivePlayback: () =>
      ({
        itemId: 'item-2',
        playMethod: 'DirectPlay',
        loadedMediaPath: 'https://stream.example/video.m3u8',
        stopReportsAfterMs: 20_000,
      }) as never,
    clearActivePlayback: () => {
      cleared = true;
    },
    getSession: () => ({
      isConnected: () => true,
      reportProgress: async () => {},
      reportStopped: async () => {
        stopped = true;
      },
    }),
    getMpvClient: () => ({
      currentTimePos: 0,
    }),
    getNow: () => 12_000,
    ticksPerSecond: 10_000_000,
    logDebug: () => {},
  });

  await reportStopped();

  assert.equal(stopped, false);
  assert.equal(cleared, false);
});
