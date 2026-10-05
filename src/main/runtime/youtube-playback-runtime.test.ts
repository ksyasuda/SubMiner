import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createYoutubePlaybackRuntime,
  type YoutubePlaybackRuntimeDeps,
} from './youtube-playback-runtime';

function makeRuntime(overrides: Partial<YoutubePlaybackRuntimeDeps> = {}) {
  return createYoutubePlaybackRuntime({
    platform: 'linux',
    directPlaybackFormat: 'best',
    mpvYtdlFormat: 'bestvideo+bestaudio',
    autoLaunchTimeoutMs: 2_000,
    connectTimeoutMs: 1_000,
    getSocketPath: () => '/tmp/mpv.sock',
    getMpvConnected: () => true,
    invalidatePendingAutoplayReadyFallbacks: () => {},
    setAppOwnedFlowInFlight: () => {},
    ensureYoutubePlaybackRuntimeReady: async () => {},
    resolveYoutubePlaybackUrl: async () => {
      throw new Error('only windows resolves direct playback urls');
    },
    launchWindowsMpv: async () => ({ ok: false }),
    waitForYoutubeMpvConnected: async () => true,
    prepareYoutubePlaybackInMpv: async () => true,
    runYoutubePlaybackFlow: async () => {},
    logInfo: () => {},
    logWarn: () => {},
    schedule: () => 1 as never,
    clearScheduled: () => {},
    ...overrides,
  });
}

test('youtube playback runtime releases flow ownership and arms quit-on-disconnect after an initial run', async () => {
  const appOwned: boolean[] = [];
  const waitTimeouts: number[] = [];
  let invalidated = 0;
  let armQuitOnDisconnect: (() => void) | null = null;

  const runtime = makeRuntime({
    invalidatePendingAutoplayReadyFallbacks: () => {
      invalidated += 1;
    },
    setAppOwnedFlowInFlight: (next) => appOwned.push(next),
    waitForYoutubeMpvConnected: async (timeoutMs) => {
      waitTimeouts.push(timeoutMs);
      return true;
    },
    schedule: (callback) => {
      armQuitOnDisconnect = callback;
      return 1 as never;
    },
  });

  await runtime.runYoutubePlaybackFlow({ url: 'https://youtu.be/demo', source: 'initial' });

  assert.equal(invalidated, 1);
  assert.deepEqual(appOwned, [true, false]);
  assert.deepEqual(waitTimeouts, [1_000]);
  assert.equal(runtime.getQuitOnDisconnectArmed(), false);

  assert.ok(armQuitOnDisconnect);
  (armQuitOnDisconnect as () => void)();
  assert.equal(runtime.getQuitOnDisconnectArmed(), true);
});

test('youtube playback runtime resolves the socket path lazily for windows startup', async () => {
  let socketPath = '/tmp/initial.sock';
  const launchArgs: string[][] = [];
  const prepared: Array<{ url: string; sourceUrl: string }> = [];

  const runtime = makeRuntime({
    platform: 'win32',
    getSocketPath: () => socketPath,
    getMpvConnected: () => false,
    resolveYoutubePlaybackUrl: async () => 'https://example.com/direct',
    launchWindowsMpv: async (_playbackUrl, args) => {
      launchArgs.push(args);
      return { ok: true, mpvPath: '/usr/bin/mpv' };
    },
    prepareYoutubePlaybackInMpv: async (request) => {
      prepared.push(request);
      return true;
    },
  });

  socketPath = '/tmp/updated.sock';
  await runtime.runYoutubePlaybackFlow({ url: 'https://youtu.be/demo', source: 'initial' });

  assert.ok(launchArgs[0]?.includes('--input-ipc-server=/tmp/updated.sock'));
  // The page URL rides along so prepare can swap a queued entry for the direct stream in place.
  assert.deepEqual(prepared, [
    { url: 'https://example.com/direct', sourceUrl: 'https://youtu.be/demo' },
  ]);
});

test('youtube playback runtime maps resolved windows streams back to their page urls', async () => {
  const streamFor = (url: string) =>
    `https://rr1---sn.example.googlevideo.com/videoplayback?src=${encodeURIComponent(url)}`;
  const firstUrl = 'https://www.youtube.com/watch?v=abcdefghijk';
  const secondUrl = 'https://www.youtube.com/watch?v=bcdefghijkl';
  const runtime = makeRuntime({
    platform: 'win32',
    directPlaybackFormat: 'b',
    resolveYoutubePlaybackUrl: async (url) => streamFor(url),
    // The second video never loads, so the first stream keeps playing.
    prepareYoutubePlaybackInMpv: async ({ url }) => url === streamFor(firstUrl),
  });

  assert.equal(runtime.getYoutubeSourceUrlForStream(streamFor(firstUrl)), null);
  await runtime.runYoutubePlaybackFlow({ url: firstUrl, source: 'second-instance' });
  await runtime
    .runYoutubePlaybackFlow({ url: secondUrl, source: 'second-instance' })
    .catch(() => {});

  assert.equal(runtime.getYoutubeSourceUrlForStream(streamFor(firstUrl)), firstUrl);
  assert.equal(runtime.getYoutubeSourceUrlForStream(streamFor(secondUrl)), secondUrl);
  assert.equal(
    runtime.getYoutubeSourceUrlForStream('https://rr2---sn.example.googlevideo.com/other'),
    null,
  );
  assert.equal(runtime.getYoutubeSourceUrlForStream('/video/episode.mkv'), null);
});

test('youtube playback runtime starts media cache without blocking the subtitle flow', async () => {
  const calls: string[] = [];
  let resolveCache: (() => void) | undefined;
  const cachePromise = new Promise<void>((resolve) => {
    resolveCache = resolve;
  });

  const runtime = makeRuntime({
    prepareYoutubePlaybackInMpv: async () => {
      calls.push('prepare');
      return true;
    },
    startYoutubeMediaCache: async () => {
      calls.push('cache');
      await cachePromise;
      calls.push('cache-done');
    },
    runYoutubePlaybackFlow: async () => {
      calls.push('run-flow');
    },
  });

  await runtime.runYoutubePlaybackFlow({ url: 'https://youtu.be/demo', source: 'second-instance' });

  // The cache starts once media is ready, and the flow does not wait for it to finish.
  assert.deepEqual(calls, ['prepare', 'cache', 'run-flow']);
  resolveCache?.();
});

test('youtube playback runtime logs synchronous media cache startup failures', async () => {
  const warnings: string[] = [];
  let flowRan = false;

  const runtime = makeRuntime({
    startYoutubeMediaCache: () => {
      throw new Error('cache exploded');
    },
    runYoutubePlaybackFlow: async () => {
      flowRan = true;
    },
    logWarn: (message) => warnings.push(message),
  });

  await runtime.runYoutubePlaybackFlow({ url: 'https://youtu.be/demo', source: 'second-instance' });
  await Promise.resolve();

  assert.equal(flowRan, true);
  assert.ok(
    warnings.some((entry) =>
      entry.startsWith('Failed to start YouTube media cache: cache exploded'),
    ),
  );
});
