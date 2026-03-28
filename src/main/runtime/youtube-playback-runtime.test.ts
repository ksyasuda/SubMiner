import assert from 'node:assert/strict';
import test from 'node:test';
import { createYoutubePlaybackRuntime } from './youtube-playback-runtime';

test('youtube playback runtime resets flow ownership after a successful run', async () => {
  const calls: string[] = [];
  let appOwnedFlowInFlight = false;
  let timeoutCallback: (() => void) | null = null;

  const runtime = createYoutubePlaybackRuntime({
    platform: 'linux',
    directPlaybackFormat: 'best',
    mpvYtdlFormat: 'bestvideo+bestaudio',
    autoLaunchTimeoutMs: 2_000,
    connectTimeoutMs: 1_000,
    socketPath: '/tmp/mpv.sock',
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
    launchWindowsMpv: () => ({ ok: false }),
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
