import test from 'node:test';
import assert from 'node:assert/strict';
import {
  markJellyfinRemotePlaybackLoaded,
  createJellyfinRemoteReportTracker,
  createReportJellyfinRemoteProgressHandler,
  createReportJellyfinRemoteStoppedHandler,
  secondsToJellyfinTicks,
  shouldAutoLoadSecondarySubTrackForJellyfinPlayback,
} from './jellyfin-remote-playback';
import type { ActiveJellyfinRemotePlaybackState } from './jellyfin-remote-commands';

type ProgressDeps = Parameters<typeof createReportJellyfinRemoteProgressHandler>[0];
type StoppedDeps = Parameters<typeof createReportJellyfinRemoteStoppedHandler>[0];
type Session = NonNullable<ReturnType<ProgressDeps['getSession']>>;

const TICKS_PER_SECOND = 10_000_000;

const makePlayback = (
  overrides: Partial<ActiveJellyfinRemotePlaybackState> = {},
): ActiveJellyfinRemotePlaybackState => ({
  itemId: 'item-1',
  playMethod: 'DirectPlay',
  ...overrides,
});

const makeSession = (overrides: Partial<Session> = {}): Session => ({
  isConnected: () => true,
  reportProgress: async () => {},
  reportStopped: async () => {},
  ...overrides,
});

function makeProgressDeps(overrides: Partial<ProgressDeps> = {}): ProgressDeps {
  return {
    getActivePlayback: () => makePlayback(),
    clearActivePlayback: () => {},
    getSession: () => makeSession(),
    getMpvClient: () => ({ requestProperty: async () => 1 }),
    getNow: () => 5000,
    getLastProgressAtMs: () => 0,
    setLastProgressAtMs: () => {},
    progressIntervalMs: 3000,
    ticksPerSecond: TICKS_PER_SECOND,
    logDebug: () => {},
    ...overrides,
  };
}

function makeStoppedDeps(overrides: Partial<StoppedDeps> = {}): StoppedDeps {
  return {
    getActivePlayback: () => makePlayback({ loadedMediaPath: 'https://stream.example/video.m3u8' }),
    clearActivePlayback: () => {},
    getSession: () => makeSession(),
    getMpvClient: () => ({ currentTimePos: 0 }),
    ticksPerSecond: TICKS_PER_SECOND,
    logDebug: () => {},
    ...overrides,
  };
}

test('secondsToJellyfinTicks converts seconds and clamps invalid values', () => {
  assert.equal(secondsToJellyfinTicks(1.25, 10_000_000), 12_500_000);
  assert.equal(secondsToJellyfinTicks(-3, 10_000_000), 0);
  assert.equal(secondsToJellyfinTicks(Number.NaN, 10_000_000), 0);
});

test('shouldAutoLoadSecondarySubTrackForJellyfinPlayback suppresses generic secondary autoload for active Jellyfin media', () => {
  assert.equal(shouldAutoLoadSecondarySubTrackForJellyfinPlayback(null, '/tmp/local.mkv'), true);
  assert.equal(
    shouldAutoLoadSecondarySubTrackForJellyfinPlayback(
      { itemId: 'item-1', playMethod: 'DirectPlay', loadedMediaPath: null },
      'http://pve-main:8096/Videos/item/stream',
    ),
    false,
  );
  assert.equal(
    shouldAutoLoadSecondarySubTrackForJellyfinPlayback(
      {
        itemId: 'item-1',
        playMethod: 'DirectPlay',
        loadedMediaPath: 'http://pve-main:8096/Videos/item/stream',
      },
      'http://pve-main:8096/Videos/item/stream',
    ),
    false,
  );
  assert.equal(
    shouldAutoLoadSecondarySubTrackForJellyfinPlayback(
      {
        itemId: 'item-1',
        playMethod: 'DirectPlay',
        loadedMediaPath: 'http://pve-main:8096/Videos/item/stream',
      },
      '/tmp/local.mkv',
    ),
    true,
  );
});

test('createReportJellyfinRemoteProgressHandler reports playback progress', async () => {
  let lastProgressAtMs = 0;
  const reportPayloads: Array<{ itemId: string; positionTicks: number; isPaused: boolean }> = [];

  const reportProgress = createReportJellyfinRemoteProgressHandler(
    makeProgressDeps({
      getSession: () =>
        makeSession({
          reportProgress: async (payload) => {
            reportPayloads.push({
              itemId: payload.itemId,
              positionTicks: payload.positionTicks,
              isPaused: payload.isPaused,
            });
          },
        }),
      getMpvClient: () => ({
        requestProperty: async (name: string) => (name === 'time-pos' ? 2.5 : true),
      }),
      setLastProgressAtMs: (value) => {
        lastProgressAtMs = value;
      },
    }),
  );

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

  const reportProgress = createReportJellyfinRemoteProgressHandler(
    makeProgressDeps({
      getSession: () =>
        makeSession({
          isConnected: () => false,
          reportProgress: async (payload) => {
            reportPayloads.push({
              positionTicks: payload.positionTicks,
              isPaused: payload.isPaused,
            });
          },
        }),
      getMpvClient: () => ({
        currentTimePos: 42,
        requestProperty: async (name: string) => (name === 'pause' ? false : 42),
      }),
    }),
  );

  await reportProgress(true);

  assert.deepEqual(reportPayloads, [{ positionTicks: 420_000_000, isPaused: false }]);
});

test('createReportJellyfinRemoteProgressHandler normalizes mpv pause strings', async () => {
  const reportPayloads: Array<{ isPaused: boolean }> = [];

  const reportProgress = createReportJellyfinRemoteProgressHandler(
    makeProgressDeps({
      getSession: () =>
        makeSession({
          reportProgress: async (payload) => {
            reportPayloads.push({ isPaused: payload.isPaused });
          },
        }),
      getMpvClient: () => ({
        requestProperty: async (name: string) => (name === 'pause' ? 'yes' : 3),
      }),
    }),
  );

  await reportProgress(true);

  assert.deepEqual(reportPayloads, [{ isPaused: true }]);
});

test('createReportJellyfinRemoteProgressHandler respects debounce interval', async () => {
  let called = false;
  const reportProgress = createReportJellyfinRemoteProgressHandler(
    makeProgressDeps({
      getSession: () =>
        makeSession({
          reportProgress: async () => {
            called = true;
          },
        }),
      getNow: () => 4000,
      getLastProgressAtMs: () => 3500,
    }),
  );

  await reportProgress(false);
  assert.equal(called, false);
});

test('createReportJellyfinRemoteProgressHandler reports mpv seek jumps during debounce', async () => {
  let now = 5000;
  let lastProgressAtMs = 0;
  let position = 10;
  const reportPayloads: Array<{ positionTicks: number; eventName: string }> = [];

  const reportProgress = createReportJellyfinRemoteProgressHandler(
    makeProgressDeps({
      getSession: () =>
        makeSession({
          reportProgress: async (payload) => {
            reportPayloads.push({
              positionTicks: payload.positionTicks,
              eventName: payload.eventName,
            });
          },
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
    }),
  );

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
  const reportStopped = createReportJellyfinRemoteStoppedHandler(
    makeStoppedDeps({
      getActivePlayback: () => makePlayback({ itemId: 'item-2', playMethod: 'Transcode' }),
      clearActivePlayback: () => {
        cleared = true;
      },
      getSession: () =>
        makeSession({
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
    }),
  );

  await reportStopped();
  assert.deepEqual(stoppedPayload, {
    itemId: 'item-2',
    positionTicks: 125_000_000,
    failed: false,
  });
  assert.equal(cleared, true);
});

test('createReportJellyfinRemoteStoppedHandler clears aborted playback that never loaded', async () => {
  let cleared = false;
  const reportStopped = createReportJellyfinRemoteStoppedHandler(
    makeStoppedDeps({
      getActivePlayback: () => makePlayback({ loadedMediaPath: null }),
      clearActivePlayback: () => {
        cleared = true;
      },
      getSession: () =>
        makeSession({
          reportStopped: async () => {
            throw new Error('should not report stopped for unloaded media');
          },
        }),
      getMpvClient: () => null,
    }),
  );

  await reportStopped();

  assert.equal(cleared, true);
});

test('createReportJellyfinRemoteStoppedHandler reports stop while remote websocket is disconnected', async () => {
  let cleared = false;
  let stoppedPayload: {
    itemId: string;
    positionTicks?: number;
    failed?: boolean;
  } | null = null;
  const reportStopped = createReportJellyfinRemoteStoppedHandler(
    makeStoppedDeps({
      clearActivePlayback: () => {
        cleared = true;
      },
      getSession: () =>
        makeSession({
          isConnected: () => false,
          reportStopped: async (payload) => {
            stoppedPayload = {
              itemId: payload.itemId,
              positionTicks: payload.positionTicks,
              failed: payload.failed,
            };
          },
        }),
      getMpvClient: () => ({ currentTimePos: 12.5 }),
    }),
  );

  await reportStopped();

  assert.deepEqual(stoppedPayload, {
    itemId: 'item-1',
    positionTicks: 125_000_000,
    failed: false,
  });
  assert.equal(cleared, true);
});

test('createReportJellyfinRemoteStoppedHandler uses cached position after mpv unload reset', async () => {
  let cleared = false;
  const calls: Array<{ event: string; positionTicks?: number }> = [];
  const reportStopped = createReportJellyfinRemoteStoppedHandler(
    makeStoppedDeps({
      getActivePlayback: () =>
        makePlayback({
          loadedMediaPath: 'https://stream.example/video.m3u8',
          lastKnownPositionSeconds: 72.25,
        }),
      clearActivePlayback: () => {
        cleared = true;
      },
      getSession: () =>
        makeSession({
          reportProgress: async (payload) => {
            calls.push({ event: 'progress', positionTicks: payload.positionTicks });
          },
          reportStopped: async (payload) => {
            calls.push({ event: 'stopped', positionTicks: payload.positionTicks });
          },
        }),
    }),
  );

  await reportStopped();

  assert.deepEqual(calls, [
    { event: 'progress', positionTicks: 722_500_000 },
    { event: 'stopped', positionTicks: 722_500_000 },
  ]);
  assert.equal(cleared, true);
});

test('createReportJellyfinRemoteProgressHandler caches last nonzero mpv position', async () => {
  let position = 42;
  const playback = makePlayback();
  const reportProgress = createReportJellyfinRemoteProgressHandler(
    makeProgressDeps({
      getActivePlayback: () => playback,
      getMpvClient: () => ({
        currentTimePos: position,
        requestProperty: async (name: string) => (name === 'pause' ? false : position),
      }),
    }),
  );

  await reportProgress(true);
  position = 0;
  await reportProgress(true);

  assert.equal(playback.lastKnownPositionSeconds, 42);
});

test('markJellyfinRemotePlaybackLoaded preserves the loaded marker on unload paths', () => {
  const playback = makePlayback({
    itemId: 'item-2',
    playMethod: 'Transcode',
    loadedMediaPath: 'https://stream.example/video.m3u8',
  });

  markJellyfinRemotePlaybackLoaded(playback, '');
  markJellyfinRemotePlaybackLoaded(playback, '   ');
  assert.equal(playback.loadedMediaPath, 'https://stream.example/video.m3u8');

  markJellyfinRemotePlaybackLoaded(playback, '  https://stream.example/next.m3u8  ');
  assert.equal(playback.loadedMediaPath, 'https://stream.example/next.m3u8');
});

test('createReportJellyfinRemoteStoppedHandler ignores startup stop churn before grace expires', async () => {
  let cleared = false;
  let stopped = false;
  const reportStopped = createReportJellyfinRemoteStoppedHandler(
    makeStoppedDeps({
      getActivePlayback: () =>
        makePlayback({
          loadedMediaPath: 'https://stream.example/video.m3u8',
          stopReportsAfterMs: 20_000,
        }),
      clearActivePlayback: () => {
        cleared = true;
      },
      getSession: () =>
        makeSession({
          reportStopped: async () => {
            stopped = true;
          },
        }),
      getNow: () => 12_000,
    }),
  );

  await reportStopped();

  assert.equal(stopped, false);
  assert.equal(cleared, false);
});

test('createReportJellyfinRemoteStoppedHandler clears playback before reporting and waits for in-flight progress', async () => {
  const tracker = createJellyfinRemoteReportTracker();
  let playback: { itemId: string; playMethod: 'DirectPlay'; loadedMediaPath: string } | null = {
    itemId: 'item-1',
    playMethod: 'DirectPlay',
    loadedMediaPath: 'http://pve-main:8096/Videos/item-1/stream',
  };
  const calls: string[] = [];
  let releaseProgress: () => void = () => undefined;
  const progressGate = new Promise<void>((resolve) => {
    releaseProgress = resolve;
  });
  const session = {
    isConnected: () => true,
    reportProgress: async ({ eventName }: { eventName: string }) => {
      calls.push(`progress:${eventName}:${playback ? 'active' : 'cleared'}`);
      if (calls.length === 1) await progressGate;
      return true;
    },
    reportStopped: async () => {
      calls.push(`stopped:${playback ? 'active' : 'cleared'}`);
      return true;
    },
  };
  const shared = {
    getActivePlayback: () => playback,
    clearActivePlayback: () => {
      playback = null;
    },
    getSession: () => session,
    getMpvClient: () => ({ currentTimePos: 42 }),
    ticksPerSecond: 10_000_000,
    logDebug: () => undefined,
    reportTracker: tracker,
  };
  const reportProgress = createReportJellyfinRemoteProgressHandler({
    ...shared,
    getNow: () => 10_000,
    getLastProgressAtMs: () => 0,
    setLastProgressAtMs: () => undefined,
    progressIntervalMs: 3000,
  });
  const reportStopped = createReportJellyfinRemoteStoppedHandler(shared);

  // A periodic tick is mid-request when the stop starts.
  const tick = reportProgress(true);
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.deepEqual(calls, ['progress:TimeUpdate:active']);
  const stop = reportStopped();
  await new Promise((resolve) => setTimeout(resolve, 0));
  assert.equal(playback, null);
  assert.deepEqual(calls, ['progress:TimeUpdate:active']);

  // A tick fired after the stop began must not report anything.
  await reportProgress(true);
  assert.deepEqual(calls, ['progress:TimeUpdate:active']);

  releaseProgress();
  await tick;
  await stop;
  assert.deepEqual(calls, [
    'progress:TimeUpdate:active',
    'progress:TimeUpdate:cleared',
    'stopped:cleared',
  ]);
});
