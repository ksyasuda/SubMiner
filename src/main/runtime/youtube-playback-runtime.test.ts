import assert from 'node:assert/strict';
import test from 'node:test';
import { createYoutubePlaybackRuntime } from './youtube-playback-runtime';

test('youtube playback runtime resets flow ownership after a successful run', async () => {
  const calls: string[] = [];
  let appOwnedFlowInFlight = false;
  let timeoutCallback: (() => void) | null = null;
  let socketPath = '/tmp/mpv.sock';

  const runtime = createYoutubePlaybackRuntime({
    platform: 'linux',
    directPlaybackFormat: 'best',
    mpvYtdlFormat: 'bestvideo+bestaudio',
    autoLaunchTimeoutMs: 2_000,
    connectTimeoutMs: 1_000,
    getSocketPath: () => socketPath,
    getMpvConnected: () => true,
    invalidatePendingAutoplayReadyFallbacks: () => {
      calls.push('invalidate-autoplay');
    },
    setAppOwnedFlowInFlight: (next) => {
      appOwnedFlowInFlight = next;
      calls.push(`app-owned:${next}`);
    },
    ensureYoutubePlaybackRuntimeReady: async () => {
      calls.push('ensure-runtime-ready');
    },
    resolveYoutubePlaybackUrl: async () => {
      throw new Error('linux path should not resolve direct playback url');
    },
    launchWindowsMpv: async () => ({ ok: false }),
    waitForYoutubeMpvConnected: async (timeoutMs) => {
      calls.push(`wait-connected:${timeoutMs}`);
      return true;
    },
    prepareYoutubePlaybackInMpv: async ({ url }) => {
      calls.push(`prepare:${url}`);
      return true;
    },
    runYoutubePlaybackFlow: async ({ url, mode }) => {
      calls.push(`run-flow:${url}:${mode}`);
    },
    logInfo: (message) => {
      calls.push(`info:${message}`);
    },
    logWarn: (message) => {
      calls.push(`warn:${message}`);
    },
    schedule: (callback) => {
      timeoutCallback = callback;
      calls.push('schedule-arm');
      return 1 as never;
    },
    clearScheduled: () => {
      calls.push('clear-scheduled');
    },
  });

  await runtime.runYoutubePlaybackFlow({
    url: 'https://youtu.be/demo',
    mode: 'download',
    source: 'initial',
  });

  assert.equal(appOwnedFlowInFlight, false);
  assert.equal(runtime.getQuitOnDisconnectArmed(), false);
  assert.deepEqual(calls.slice(0, 6), [
    'invalidate-autoplay',
    'app-owned:true',
    'ensure-runtime-ready',
    'wait-connected:1000',
    'schedule-arm',
    'prepare:https://youtu.be/demo',
  ]);

  assert.ok(timeoutCallback);
  const scheduledCallback = timeoutCallback as () => void;
  scheduledCallback();
  assert.equal(runtime.getQuitOnDisconnectArmed(), true);
});

test('youtube playback runtime resolves the socket path lazily for windows startup', async () => {
  const calls: string[] = [];
  let socketPath = '/tmp/initial.sock';

  const runtime = createYoutubePlaybackRuntime({
    platform: 'win32',
    directPlaybackFormat: 'best',
    mpvYtdlFormat: 'bestvideo+bestaudio',
    autoLaunchTimeoutMs: 2_000,
    connectTimeoutMs: 1_000,
    getSocketPath: () => socketPath,
    getMpvConnected: () => false,
    invalidatePendingAutoplayReadyFallbacks: () => {
      calls.push('invalidate-autoplay');
    },
    setAppOwnedFlowInFlight: (next) => {
      calls.push(`app-owned:${next}`);
    },
    ensureYoutubePlaybackRuntimeReady: async () => {
      calls.push('ensure-runtime-ready');
    },
    resolveYoutubePlaybackUrl: async (url, format) => {
      calls.push(`resolve:${url}:${format}`);
      return 'https://example.com/direct';
    },
    launchWindowsMpv: async (_playbackUrl, args) => {
      calls.push(`launch:${args.join(' ')}`);
      return { ok: true, mpvPath: '/usr/bin/mpv' };
    },
    waitForYoutubeMpvConnected: async (timeoutMs) => {
      calls.push(`wait-connected:${timeoutMs}`);
      return true;
    },
    prepareYoutubePlaybackInMpv: async ({ url }) => {
      calls.push(`prepare:${url}`);
      return true;
    },
    runYoutubePlaybackFlow: async ({ url, mode }) => {
      calls.push(`run-flow:${url}:${mode}`);
    },
    logInfo: (message) => {
      calls.push(`info:${message}`);
    },
    logWarn: (message) => {
      calls.push(`warn:${message}`);
    },
    schedule: (callback) => {
      calls.push('schedule-arm');
      callback();
      return 1 as never;
    },
    clearScheduled: () => {
      calls.push('clear-scheduled');
    },
  });

  socketPath = '/tmp/updated.sock';

  await runtime.runYoutubePlaybackFlow({
    url: 'https://youtu.be/demo',
    mode: 'download',
    source: 'initial',
  });

  assert.ok(calls.some((entry) => entry.includes('--input-ipc-server=/tmp/updated.sock')));
});

test('youtube playback runtime starts media cache without blocking the subtitle flow', async () => {
  const calls: string[] = [];
  let resolveCache: (() => void) | undefined;
  const cachePromise = new Promise<void>((resolve) => {
    resolveCache = resolve;
  });

  const runtime = createYoutubePlaybackRuntime({
    platform: 'linux',
    directPlaybackFormat: 'best',
    mpvYtdlFormat: 'bestvideo+bestaudio',
    autoLaunchTimeoutMs: 2_000,
    connectTimeoutMs: 1_000,
    getSocketPath: () => '/tmp/mpv.sock',
    getMpvConnected: () => true,
    invalidatePendingAutoplayReadyFallbacks: () => {
      calls.push('invalidate-autoplay');
    },
    setAppOwnedFlowInFlight: (next) => {
      calls.push(`app-owned:${next}`);
    },
    ensureYoutubePlaybackRuntimeReady: async () => {
      calls.push('ensure-runtime-ready');
    },
    resolveYoutubePlaybackUrl: async () => {
      throw new Error('linux path should not resolve direct playback url');
    },
    launchWindowsMpv: async () => ({ ok: false }),
    waitForYoutubeMpvConnected: async () => true,
    prepareYoutubePlaybackInMpv: async ({ url }) => {
      calls.push(`prepare:${url}`);
      return true;
    },
    startYoutubeMediaCache: async (url) => {
      calls.push(`cache:${url}`);
      await cachePromise;
      calls.push('cache-done');
    },
    runYoutubePlaybackFlow: async ({ url, mode }) => {
      calls.push(`run-flow:${url}:${mode}`);
    },
    logInfo: (message) => {
      calls.push(`info:${message}`);
    },
    logWarn: (message) => {
      calls.push(`warn:${message}`);
    },
    schedule: () => 1 as never,
    clearScheduled: () => {},
  });

  await runtime.runYoutubePlaybackFlow({
    url: 'https://youtu.be/demo',
    mode: 'download',
    source: 'second-instance',
  });

  assert.ok(
    calls.indexOf('prepare:https://youtu.be/demo') < calls.indexOf('cache:https://youtu.be/demo'),
  );
  assert.ok(
    calls.indexOf('cache:https://youtu.be/demo') <
      calls.indexOf('run-flow:https://youtu.be/demo:download'),
  );
  assert.equal(calls.includes('cache-done'), false);
  const resolveCacheNow = resolveCache;
  assert.ok(resolveCacheNow);
  resolveCacheNow();
});
