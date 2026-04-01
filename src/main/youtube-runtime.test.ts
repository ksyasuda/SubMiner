import assert from 'node:assert/strict';
import test from 'node:test';

import { createYoutubeRuntime } from './youtube-runtime';

function createRuntime(overrides: Partial<Parameters<typeof createYoutubeRuntime>[0]> = {}) {
  const calls: string[] = [];

  const runtime = createYoutubeRuntime({
    flow: {
      probeYoutubeTracks: async () => ({ videoId: 'demo', title: 'Demo', tracks: [] }),
      acquireYoutubeSubtitleTrack: async () => ({ path: '/tmp/primary.vtt' }),
      acquireYoutubeSubtitleTracks: async () => new Map<string, string>(),
      openPicker: async () => true,
      pauseMpv: () => {},
      resumeMpv: () => {},
      sendMpvCommand: () => {},
      requestMpvProperty: async () => null,
      refreshCurrentSubtitle: () => {},
      startTokenizationWarmups: async () => {},
      waitForTokenizationReady: async () => {},
      waitForAnkiReady: async () => {},
      wait: async () => {},
      waitForPlaybackWindowReady: async () => {},
      waitForOverlayGeometryReady: async () => {},
      focusOverlayWindow: () => {},
      showMpvOsd: (message) => {
        calls.push(`flow-osd:${message}`);
      },
      warn: (message) => {
        calls.push(`warn:${message}`);
      },
      log: (message) => {
        calls.push(`log:${message}`);
      },
      getYoutubeOutputDir: () => '/tmp',
    },
    playback: {
      platform: 'linux',
      directPlaybackFormat: 'b',
      mpvYtdlFormat: 'best',
      autoLaunchTimeoutMs: 1000,
      connectTimeoutMs: 1000,
      getSocketPath: () => '/tmp/mpv.sock',
      getMpvConnected: () => true,
      ensureYoutubePlaybackRuntimeReady: async () => {},
      resolveYoutubePlaybackUrl: async (url) => url,
      launchWindowsMpv: () => ({ ok: false }),
      waitForYoutubeMpvConnected: async () => true,
      prepareYoutubePlaybackInMpv: async () => true,
      logInfo: () => {},
      logWarn: () => {},
      schedule: (callback) => setTimeout(callback, 0),
      clearScheduled: (timer) => clearTimeout(timer),
    },
    autoplay: {
      getCurrentMediaPath: () => null,
      getCurrentVideoPath: () => null,
      getPlaybackPaused: () => true,
      getMpvClient: () => null,
      signalPluginAutoplayReady: () => {
        calls.push('autoplay-ready');
      },
      schedule: (callback) => setTimeout(callback, 0),
      logDebug: () => {},
    },
    notification: {
      getPrimarySubtitleLanguages: () => ['ja'],
      schedule: (callback) => setTimeout(callback, 0),
      clearSchedule: (timer) => clearTimeout(timer as ReturnType<typeof setTimeout>),
    },
    getNotificationType: () => 'osd',
    getCurrentMediaPath: () => null,
    getCurrentVideoPath: () => null,
    showMpvOsd: (message) => {
      calls.push(`osd:${message}`);
    },
    showDesktopNotification: (_title, options) => {
      calls.push(`notify:${options.body}`);
    },
    broadcastYoutubePickerCancel: () => {
      calls.push('picker-cancel');
    },
    closeYoutubePickerModal: () => {
      calls.push('close-modal');
    },
    logWarn: (message) => {
      calls.push(`warn:${message}`);
    },
    ...overrides,
  });

  return {
    runtime,
    calls,
  };
}

test('youtube runtime gates manual picker availability by playback context', async () => {
  const inactive = createRuntime({
    getCurrentMediaPath: () => '/tmp/video.mkv',
    getCurrentVideoPath: () => null,
  });

  await inactive.runtime.openYoutubeTrackPickerFromPlayback();
  assert.ok(
    inactive.calls.includes(
      'osd:YouTube subtitle picker is only available during YouTube playback.',
    ),
  );

  const active = createRuntime({
    getCurrentMediaPath: () => 'https://www.youtube.com/watch?v=demo',
    getCurrentVideoPath: () => null,
    createFlowRuntime: () => ({
      runYoutubePlaybackFlow: async () => {},
      openManualPicker: async ({ url }) => {
        active.calls.push(`manual-picker:${url}`);
      },
      resolveActivePicker: async () => ({ ok: true, message: 'resolved' }),
      cancelActivePicker: () => true,
      hasActiveSession: () => false,
    }),
  });

  await active.runtime.openYoutubeTrackPickerFromPlayback();
  assert.ok(active.calls.includes('manual-picker:https://www.youtube.com/watch?v=demo'));
});

test('youtube runtime cancels active picker on mpv disconnect', () => {
  const harness = createRuntime({
    createFlowRuntime: () => ({
      runYoutubePlaybackFlow: async () => {},
      openManualPicker: async () => {},
      resolveActivePicker: async () => ({ ok: true, message: 'resolved' }),
      cancelActivePicker: () => {
        harness.calls.push('cancel-active');
        return true;
      },
      hasActiveSession: () => true,
    }),
  });

  harness.runtime.handleMpvConnectionChange(false);

  assert.deepEqual(harness.calls, ['cancel-active', 'picker-cancel', 'close-modal']);
});

test('youtube runtime delegates picker resolution to flow runtime', async () => {
  const harness = createRuntime({
    createFlowRuntime: () => ({
      runYoutubePlaybackFlow: async () => {},
      openManualPicker: async () => {},
      resolveActivePicker: async (request) => ({ request, ok: true, message: 'resolved' }),
      cancelActivePicker: () => true,
      hasActiveSession: () => false,
    }),
  });

  const request = {
    sessionId: 'session-1',
    action: 'use-selected' as const,
    primaryTrackId: 'ja',
    secondaryTrackId: null,
  };
  const result = await harness.runtime.resolveActivePicker(request);
  assert.deepEqual(result, { request, ok: true, message: 'resolved' });
});

test('youtube runtime routes subtitle failures through configured notification channels', () => {
  const harness = createRuntime({
    getNotificationType: () => 'both',
  });

  harness.runtime.reportYoutubeSubtitleFailure('Primary subtitles failed');

  assert.ok(harness.calls.includes('osd:Primary subtitles failed'));
  assert.ok(harness.calls.includes('notify:Primary subtitles failed'));
});
