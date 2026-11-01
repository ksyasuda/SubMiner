import assert from 'node:assert/strict';
import test from 'node:test';
import { DEFAULT_CONFIG } from '../../config/definitions';
import type { ResolvedConfig } from '../../types';
import { createYoutubeAssembly, type YoutubeAssemblyDeps } from './youtube';

type StatusCall = { message: string; title?: string };

// Every callback throws, so a test fails if the assembly calls a dep it should only call later.
function makeDeps(overrides: Partial<YoutubeAssemblyDeps> = {}): YoutubeAssemblyDeps {
  const lazy = (name: string) => (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
  return {
    appState: {
      mpvClient: null,
      currentMediaPath: null,
      ankiIntegration: null,
      windowTracker: null,
      mpvSocketPath: '/tmp/subminer-test.sock',
      initialArgs: null,
    },
    getConfig: lazy('getConfig'),
    logger: { info: () => {}, warn: () => {}, debug: () => {} },
    getWhisperModelDirectory: lazy('getWhisperModelDirectory'),
    getMainOverlayWindow: lazy('getMainOverlayWindow'),
    overlayModalRuntime: {
      sendToActiveOverlayWindow: lazy('sendToActiveOverlayWindow'),
      waitForModalOpen: lazy('waitForModalOpen'),
    },
    getOverlayNotifications: lazy('getOverlayNotifications'),
    getOverlayGeometry: lazy('getOverlayGeometry'),
    showMpvOsd: lazy('showMpvOsd'),
    showDesktopNotification: lazy('showDesktopNotification'),
    refreshCurrentSubtitle: lazy('refreshCurrentSubtitle'),
    refreshSubtitleSidebarSource: lazy('refreshSubtitleSidebarSource'),
    startTokenizationWarmups: lazy('startTokenizationWarmups'),
    waitForMediaTokenizationReady: lazy('waitForMediaTokenizationReady'),
    generateYoutubeSubtitles: lazy('generateYoutubeSubtitles'),
    openSubtitleGenerationModal: lazy('openSubtitleGenerationModal'),
    invalidatePendingAutoplayReadyFallbacks: lazy('invalidatePendingAutoplayReadyFallbacks'),
    ensurePlaybackRuntimeReady: lazy('ensurePlaybackRuntimeReady'),
    launchWindowsMpv: lazy('launchWindowsMpv'),
    ensureMpvConnectedForPlayback: lazy('ensureMpvConnectedForPlayback'),
    requestAppQuit: lazy('requestAppQuit'),
    browserWindow: {
      createBrowserWindow: lazy('createBrowserWindow'),
      preloadPath: '/tmp/preload-youtube-browser.js',
      ipcMain: { on: lazy('ipcMain.on') },
      showContextMenu: lazy('showContextMenu'),
      openExternal: lazy('openExternal'),
    },
    ...overrides,
  };
}

function makeNotifications(calls: StatusCall[]) {
  return {
    getConfiguredStatusNotificationType: () => 'overlay' as const,
    showConfiguredStatusNotification: (message: string, options?: { title?: string }) => {
      calls.push({ message, title: options?.title });
    },
    showOverlayNotification: () => {},
    showYoutubeFlowStatusNotification: () => {},
  };
}

const backgroundCacheConfig: ResolvedConfig = {
  ...DEFAULT_CONFIG,
  youtube: {
    ...DEFAULT_CONFIG.youtube,
    mediaCache: { ...DEFAULT_CONFIG.youtube.mediaCache, mode: 'background' },
  },
};

// main.ts builds the assembly before most siblings exist, so construction must stay lazy.
test('createYoutubeAssembly does not call its deps while constructing', () => {
  const youtube = createYoutubeAssembly(makeDeps());
  assert.equal(youtube.shouldQuitWhenMpvCloses(), false);
});

test('track picker refuses to open outside YouTube playback', async () => {
  const calls: StatusCall[] = [];
  const youtube = createYoutubeAssembly(
    makeDeps({
      appState: {
        mpvClient: null,
        currentMediaPath: '/media/episode-01.mkv',
        ankiIntegration: null,
        windowTracker: null,
        mpvSocketPath: '/tmp/subminer-test.sock',
        initialArgs: null,
      },
      getConfig: () => backgroundCacheConfig,
      getOverlayNotifications: () => makeNotifications(calls),
    }),
  );

  assert.equal(youtube.isPlaybackActiveNow(), false);
  assert.equal(youtube.shouldRequireMediaCacheForCurrentPlayback(), false);
  await youtube.openTrackPickerFromPlayback();
  assert.deepEqual(calls, [
    {
      message: 'YouTube subtitle picker is only available during YouTube playback.',
      title: 'YouTube subtitles',
    },
  ]);
});
