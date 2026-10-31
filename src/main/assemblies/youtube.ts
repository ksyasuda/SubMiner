import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { sendMpvCommandRuntime } from '../../core/services';
import {
  acquireYoutubeSubtitleTrack,
  acquireYoutubeSubtitleTracks,
} from '../../core/services/youtube/generate';
import { createYoutubeMediaCacheService } from '../../core/services/youtube/media-cache';
import { resolveYoutubePlaybackUrl } from '../../core/services/youtube/playback-resolve';
import { probeYoutubeTracks } from '../../core/services/youtube/track-probe';
import { createYoutubeWhisperSubtitleService } from '../../core/services/youtube/whisper-subtitles';
import type { BrowserWindow } from 'electron';
import type { Logger } from '../../logger';
import type { ResolvedConfig, WindowGeometry } from '../../types';
import type { OverlayModalRuntime } from '../overlay-runtime';
import type { createOverlayGeometryRuntime } from '../runtime/overlay-geometry-runtime';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import type { createSubtitleGenerationRuntime } from '../runtime/subtitle-generation-runtime';
import { createWaitForMpvConnectedHandler } from '../runtime/jellyfin-remote-connection';
import { createYoutubeBrowserPlaybackRuntime } from '../runtime/youtube-browser-playback';
import {
  createYoutubeBrowserWindowRuntime,
  type YoutubeBrowserWindowDeps,
} from '../runtime/youtube-browser-window';
import { createYoutubeFlowRuntime } from '../runtime/youtube-flow';
import { createYoutubeMediaCachePlaybackRuntime } from '../runtime/youtube-media-cache-playback';
import { openYoutubeTrackPicker } from '../runtime/youtube-picker-open';
import { isYoutubeMediaPath } from '../runtime/youtube-playback';
import { createPrepareYoutubePlaybackInMpvHandler } from '../runtime/youtube-playback-launch';
import {
  createYoutubePlaybackRuntime,
  type YoutubePlaybackRuntimeDeps,
} from '../runtime/youtube-playback-runtime';
import {
  clearYoutubePrimarySubtitleNotificationTimer,
  createYoutubePrimarySubtitleNotificationRuntime,
} from '../runtime/youtube-primary-subtitle-notification';
import type { AppState } from '../state';

const YOUTUBE_MPV_CONNECT_TIMEOUT_MS = 3000;
const YOUTUBE_MPV_AUTO_LAUNCH_TIMEOUT_MS = 10000;
// mpv's ytdl hook resolves the stream before tracks appear; a cold yt-dlp run takes well over 5s.
const YOUTUBE_MPV_LOAD_TIMEOUT_MS = 20000;
const YOUTUBE_MPV_YTDL_FORMAT = 'bestvideo*+bestaudio/best';
const YOUTUBE_DIRECT_PLAYBACK_FORMAT = 'b';

type OverlayNotifications = Pick<
  ReturnType<typeof createOverlayNotificationsRuntime>,
  | 'getConfiguredStatusNotificationType'
  | 'showConfiguredStatusNotification'
  | 'showOverlayNotification'
  | 'showYoutubeFlowStatusNotification'
>;
type OverlayGeometry = Pick<
  ReturnType<typeof createOverlayGeometryRuntime>,
  'geometryMatches' | 'getLastOverlayWindowGeometry'
>;
type YoutubeFlowDeps = Parameters<typeof createYoutubeFlowRuntime>[0];

/**
 * What the YouTube domain needs from the rest of the app. Siblings built after this assembly in
 * `main.ts` are reached through getters or callbacks, which are only called once the app runs.
 */
export interface YoutubeAssemblyDeps {
  appState: Pick<
    AppState,
    | 'mpvClient'
    | 'currentMediaPath'
    | 'ankiIntegration'
    | 'windowTracker'
    | 'mpvSocketPath'
    | 'initialArgs'
  >;
  getConfig: () => ResolvedConfig;
  logger: Pick<Logger, 'info' | 'warn' | 'debug'>;
  getWhisperModelDirectory: () => string;
  getMainOverlayWindow: () => BrowserWindow | null;
  overlayModalRuntime: Pick<OverlayModalRuntime, 'sendToActiveOverlayWindow' | 'waitForModalOpen'>;
  getOverlayNotifications: () => OverlayNotifications;
  getOverlayGeometry: () => OverlayGeometry;
  showMpvOsd: (text: string) => void;
  showDesktopNotification: (title: string, options: { body?: string }) => void;
  refreshCurrentSubtitle: YoutubeFlowDeps['refreshCurrentSubtitle'];
  refreshSubtitleSidebarSource: YoutubeFlowDeps['refreshSubtitleSidebarSource'];
  startTokenizationWarmups: () => Promise<void>;
  waitForMediaTokenizationReady: (mediaPath: string | null) => Promise<void>;
  generateYoutubeSubtitles: ReturnType<
    typeof createSubtitleGenerationRuntime
  >['generateYoutubeSubtitles'];
  openSubtitleGenerationModal: YoutubeFlowDeps['openSubtitleGenerationModal'];
  invalidatePendingAutoplayReadyFallbacks: () => void;
  ensurePlaybackRuntimeReady: () => Promise<void>;
  launchWindowsMpv: YoutubePlaybackRuntimeDeps['launchWindowsMpv'];
  ensureMpvConnectedForPlayback: () => Promise<boolean>;
  requestAppQuit: () => void;
  browserWindow: Pick<
    YoutubeBrowserWindowDeps,
    'createBrowserWindow' | 'preloadPath' | 'ipcMain' | 'showContextMenu' | 'openExternal'
  >;
}

export type YoutubeAssembly = ReturnType<typeof createYoutubeAssembly>;

/**
 * Owns YouTube playback, subtitle flow, media cache, primary-subtitle notifications and the
 * in-app YouTube browser. `main.ts` builds it once and reaches the domain through the result.
 */
export function createYoutubeAssembly(deps: YoutubeAssemblyDeps) {
  const { appState, logger } = deps;

  const whisperSubtitles = createYoutubeWhisperSubtitleService({
    getConfig: () => deps.getConfig().subtitleGeneration,
    getModelDirectory: deps.getWhisperModelDirectory,
  });
  const flow = createYoutubeFlowRuntime({
    probeYoutubeTracks: (url: string) => probeYoutubeTracks(url),
    acquireYoutubeSubtitleTrack: (input) => acquireYoutubeSubtitleTrack(input),
    acquireYoutubeSubtitleTracks: (input) => acquireYoutubeSubtitleTracks(input),
    openPicker: async (payload) => {
      return await openYoutubeTrackPicker(
        {
          sendToActiveOverlayWindow: (channel, nextPayload, runtimeOptions) =>
            deps.overlayModalRuntime.sendToActiveOverlayWindow(
              channel,
              nextPayload,
              runtimeOptions,
            ),
          waitForModalOpen: (modal, timeoutMs) =>
            deps.overlayModalRuntime.waitForModalOpen(modal, timeoutMs),
          logWarn: (message) => logger.warn(message),
        },
        payload,
      );
    },
    pauseMpv: () => {
      sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'pause', 'yes']);
    },
    resumeMpv: () => {
      sendMpvCommandRuntime(appState.mpvClient, ['set_property', 'pause', 'no']);
    },
    sendMpvCommand: (command) => {
      sendMpvCommandRuntime(appState.mpvClient, command);
    },
    requestMpvProperty: async (name: string) => {
      const client = appState.mpvClient;
      if (!client) return null;
      return await client.requestProperty(name);
    },
    refreshCurrentSubtitle: deps.refreshCurrentSubtitle,
    refreshSubtitleSidebarSource: deps.refreshSubtitleSidebarSource,
    startTokenizationWarmups: async () => {
      await deps.startTokenizationWarmups();
    },
    waitForTokenizationReady: async () => {
      await deps.waitForMediaTokenizationReady(
        appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || null,
      );
    },
    waitForAnkiReady: async () => {
      const integration = appState.ankiIntegration;
      if (!integration) {
        return;
      }
      try {
        await Promise.race([
          integration.waitUntilReady(),
          new Promise<never>((_, reject) => {
            setTimeout(
              () => reject(new Error('Timed out waiting for AnkiConnect integration')),
              2500,
            );
          }),
        ]);
      } catch (error) {
        logger.warn(
          'Continuing YouTube playback before AnkiConnect integration reported ready:',
          error instanceof Error ? error.message : String(error),
        );
      }
    },
    wait: (ms: number) => new Promise((resolve) => setTimeout(resolve, ms)),
    waitForPlaybackWindowReady: async () => {
      const deadline = Date.now() + 4000;
      let stableGeometry: WindowGeometry | null = null;
      let stableSinceMs = 0;
      while (Date.now() < deadline) {
        const tracker = appState.windowTracker;
        const trackerGeometry = tracker?.getGeometry() ?? null;
        const mediaPath =
          appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || '';
        const trackerFocused = tracker?.isTargetWindowFocused() ?? false;
        if (tracker && tracker.isTracking() && trackerGeometry && trackerFocused && mediaPath) {
          if (!deps.getOverlayGeometry().geometryMatches(stableGeometry, trackerGeometry)) {
            stableGeometry = trackerGeometry;
            stableSinceMs = Date.now();
          } else if (Date.now() - stableSinceMs >= 200) {
            return;
          }
        } else {
          stableGeometry = null;
          stableSinceMs = 0;
        }
        await new Promise((resolve) => setTimeout(resolve, 100));
      }
      logger.warn(
        'Timed out waiting for tracked playback window focus/media readiness before opening YouTube subtitle picker.',
      );
    },
    waitForOverlayGeometryReady: async () => {
      const deadline = Date.now() + 4000;
      while (Date.now() < deadline) {
        const tracker = appState.windowTracker;
        const trackerGeometry = tracker?.getGeometry() ?? null;
        const overlayGeometry = deps.getOverlayGeometry();
        if (
          trackerGeometry &&
          overlayGeometry.geometryMatches(
            overlayGeometry.getLastOverlayWindowGeometry(),
            trackerGeometry,
          )
        ) {
          return;
        }
        await new Promise((resolve) => setTimeout(resolve, 50));
      }
      logger.warn('Timed out waiting for overlay geometry to match tracked playback window.');
    },
    focusOverlayWindow: () => {
      const mainWindow = deps.getMainOverlayWindow();
      if (!mainWindow || mainWindow.isDestroyed()) {
        return;
      }
      mainWindow.setIgnoreMouseEvents(false);
      if (!mainWindow.isFocused()) {
        mainWindow.focus();
      }
      if (!mainWindow.webContents.isFocused()) {
        mainWindow.webContents.focus();
      }
    },
    showMpvOsd: (text: string) =>
      deps.getOverlayNotifications().showYoutubeFlowStatusNotification(text),
    reportSubtitleFailure: (message: string) => reportSubtitleFailure(message),
    notifyPrimarySubtitleLoaded: () =>
      primarySubtitleNotification.markCurrentMediaPrimarySubtitleLoaded(),
    warn: (message: string) => logger.warn(message),
    log: (message: string) => logger.info(message),
    getYoutubeOutputDir: () => path.join(os.homedir(), '.cache', 'subminer', 'youtube-subs'),
    createSubtitleTempDir: () =>
      fs.promises.mkdtemp(path.join(os.tmpdir(), 'subminer-youtube-subtitles-')),
    cleanupSubtitleTempDirs: (dirs) => {
      for (const dir of dirs) {
        fs.rmSync(dir, { recursive: true, force: true });
      }
    },
    getSubtitleSource: () => deps.getConfig().youtube.subtitleSource,
    // Runs as the subtitle generation modal's active job, so the modal shows it and can cancel it.
    generateWhisperSubtitles: ({ url, outputPath, signal, onProgress }) =>
      deps.generateYoutubeSubtitles({
        url,
        signal,
        generate: (jobSignal, reportProgress) =>
          whisperSubtitles.generate({
            url,
            outputPath,
            signal: jobSignal,
            onProgress: (progress) => {
              reportProgress(progress);
              onProgress(progress);
            },
          }),
      }),
    openSubtitleGenerationModal: deps.openSubtitleGenerationModal,
  });
  const prepareYoutubePlaybackInMpv = createPrepareYoutubePlaybackInMpvHandler({
    requestPath: async () => {
      const client = appState.mpvClient;
      if (!client) return null;
      const value = await client.requestProperty('path').catch(() => null);
      return typeof value === 'string' ? value : null;
    },
    requestProperty: async (name) => {
      const client = appState.mpvClient;
      if (!client) return null;
      return await client.requestProperty(name);
    },
    sendMpvCommand: (command) => {
      sendMpvCommandRuntime(appState.mpvClient, command);
    },
    wait: (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
  });
  const mediaCache = createYoutubeMediaCacheService({
    onDownloadStarted: (event) => {
      deps
        .getOverlayNotifications()
        .showConfiguredStatusNotification('YouTube media cache is downloading.', {
          id: 'youtube-media-cache-status',
          title: 'YouTube media cache',
          variant: 'progress',
          persistent: true,
        });
      logger.info(`YouTube media cache download notification shown for ${event.url}`);
    },
    onReady: (event) => {
      deps
        .getOverlayNotifications()
        .showConfiguredStatusNotification('YouTube media cache ready.', {
          id: 'youtube-media-cache-status',
          title: 'YouTube media cache',
          variant: 'success',
          persistent: false,
        });
      void appState.ankiIntegration
        ?.handleYoutubeMediaCacheReady(event.url, event.path, { notifyNoQueued: false })
        .catch((error) => {
          logger.warn(
            `Failed to apply queued YouTube media updates: ${
              error instanceof Error ? error.message : String(error)
            }`,
          );
        });
    },
    onFailed: (event) => {
      deps
        .getOverlayNotifications()
        .showConfiguredStatusNotification('YouTube media cache failed.', {
          id: 'youtube-media-cache-status',
          title: 'YouTube media cache',
          variant: 'error',
          persistent: false,
        });
      void appState.ankiIntegration
        ?.handleYoutubeMediaCacheFailed(event.url, { notifyStatus: false })
        .catch((error) => {
          logger.warn(
            `Failed to drain queued YouTube media updates after cache failure: ${
              error instanceof Error ? error.message : String(error)
            }`,
          );
        });
    },
    logInfo: (message) => logger.info(message),
    logWarn: (message) => logger.warn(message),
  });
  const mediaCachePlayback = createYoutubeMediaCachePlaybackRuntime({
    getMediaCacheConfig: () => deps.getConfig().youtube.mediaCache,
    requestMpvProperty: async (name) => {
      const client = appState.mpvClient;
      if (!client) return null;
      return await client.requestProperty(name);
    },
    startYoutubeMediaCache: (url, options) => {
      mediaCache.start(url, options);
    },
    logWarn: (message) => logger.warn(message),
  });
  const waitForYoutubeMpvConnected = createWaitForMpvConnectedHandler({
    getMpvClient: () => appState.mpvClient,
    now: () => Date.now(),
    sleep: (delayMs) => new Promise((resolve) => setTimeout(resolve, delayMs)),
  });

  const playback = createYoutubePlaybackRuntime({
    platform: process.platform,
    directPlaybackFormat: YOUTUBE_DIRECT_PLAYBACK_FORMAT,
    mpvYtdlFormat: YOUTUBE_MPV_YTDL_FORMAT,
    autoLaunchTimeoutMs: YOUTUBE_MPV_AUTO_LAUNCH_TIMEOUT_MS,
    connectTimeoutMs: YOUTUBE_MPV_CONNECT_TIMEOUT_MS,
    getSocketPath: () => appState.mpvSocketPath,
    getMpvConnected: () => Boolean(appState.mpvClient?.connected),
    invalidatePendingAutoplayReadyFallbacks: () => deps.invalidatePendingAutoplayReadyFallbacks(),
    setAppOwnedFlowInFlight: (next) => {
      primarySubtitleNotification.setAppOwnedFlowInFlight(next);
    },
    ensureYoutubePlaybackRuntimeReady: async () => {
      await deps.ensurePlaybackRuntimeReady();
    },
    resolveYoutubePlaybackUrl: (url, format) => resolveYoutubePlaybackUrl(url, format),
    launchWindowsMpv: deps.launchWindowsMpv,
    waitForYoutubeMpvConnected: (timeoutMs) => waitForYoutubeMpvConnected(timeoutMs),
    prepareYoutubePlaybackInMpv: (request) =>
      prepareYoutubePlaybackInMpv({ ...request, timeoutMs: YOUTUBE_MPV_LOAD_TIMEOUT_MS }),
    startYoutubeMediaCache: (url) => {
      const mediaCacheConfig = deps.getConfig().youtube.mediaCache;
      mediaCache.start(url, {
        mode: mediaCacheConfig.mode,
        maxHeight: mediaCacheConfig.maxHeight,
      });
    },
    runYoutubePlaybackFlow: (request) => flow.runYoutubePlaybackFlow(request),
    logInfo: (message) => logger.info(message),
    logWarn: (message) => logger.warn(message),
    schedule: (callback, delayMs) => setTimeout(callback, delayMs),
    clearScheduled: (timer) => clearTimeout(timer),
  });

  const primarySubtitleNotification = createYoutubePrimarySubtitleNotificationRuntime({
    getPrimarySubtitleLanguages: () => deps.getConfig().youtube.primarySubLanguages,
    notifyFailure: (message) => reportSubtitleFailure(message),
    schedule: (fn, delayMs) => setTimeout(fn, delayMs),
    clearSchedule: clearYoutubePrimarySubtitleNotificationTimer,
    getCurrentSubtitleState: async () => {
      const client = appState.mpvClient;
      if (!client?.connected) {
        return null;
      }
      const [sid, trackList] = await Promise.all([
        client.requestProperty('sid').catch(() => null),
        client.requestProperty('track-list').catch(() => null),
      ]);
      return {
        sid,
        trackList: Array.isArray(trackList) ? trackList : null,
      };
    },
  });

  function isPlaybackActiveNow(): boolean {
    return Boolean(getCurrentMediaCacheSourceUrlSnapshot());
  }

  function getCurrentMediaCacheSourceUrlSnapshot(): string | null {
    const currentMediaPath = appState.currentMediaPath?.trim() || null;
    if (isYoutubeMediaPath(currentMediaPath)) {
      return currentMediaPath;
    }

    const currentVideoPath = appState.mpvClient?.currentVideoPath?.trim() || null;
    if (isYoutubeMediaPath(currentVideoPath)) {
      return currentVideoPath;
    }

    return mediaCachePlayback.getActiveYoutubeSourceUrlSnapshot();
  }

  async function getCurrentMediaCacheSourceUrl(): Promise<string | null> {
    const currentMediaPath = appState.currentMediaPath?.trim() || null;
    if (isYoutubeMediaPath(currentMediaPath)) {
      return currentMediaPath;
    }

    const currentVideoPath = appState.mpvClient?.currentVideoPath?.trim() || null;
    if (isYoutubeMediaPath(currentVideoPath)) {
      return currentVideoPath;
    }

    return await mediaCachePlayback.getActiveYoutubeSourceUrl();
  }

  function shouldRequireMediaCacheForCurrentPlayback(): boolean {
    return deps.getConfig().youtube.mediaCache.mode === 'background' && isPlaybackActiveNow();
  }

  async function getCachedMediaPathForCurrentPlayback(
    currentVideoPath: string,
    _kind: 'audio' | 'video',
  ): Promise<string | null> {
    if (deps.getConfig().youtube.mediaCache.mode !== 'background') {
      return null;
    }
    const cacheSourceUrl = isYoutubeMediaPath(currentVideoPath)
      ? currentVideoPath
      : await getCurrentMediaCacheSourceUrl();
    if (!cacheSourceUrl) {
      return null;
    }
    // mpv can expose the resolved stream URL here while the cache key uses the original page URL.
    // Keep the active-cache fallback so current playback can still resolve the ready cached file.
    return (
      (await mediaCache.getCachedMediaPath(cacheSourceUrl)) ??
      (await mediaCache.getActiveCachedMediaPath())
    );
  }

  function reportSubtitleFailure(message: string): void {
    const notifications = deps.getOverlayNotifications();
    const type = notifications.getConfiguredStatusNotificationType();
    if (type === 'none') {
      return;
    }
    if (type === 'overlay' || type === 'both') {
      notifications.showOverlayNotification({
        title: 'SubMiner',
        body: message,
        variant: 'warning',
      });
    }
    if (type === 'osd' || type === 'osd-system') {
      deps.showMpvOsd(message);
    }
    if (type === 'system' || type === 'both' || type === 'osd-system') {
      try {
        deps.showDesktopNotification('SubMiner', { body: message });
      } catch {
        logger.warn(`Unable to show desktop notification: ${message}`);
      }
    }
  }

  async function openTrackPickerFromPlayback(): Promise<void> {
    if (flow.hasActiveSession()) {
      deps
        .getOverlayNotifications()
        .showConfiguredStatusNotification('YouTube subtitle flow already in progress.', {
          title: 'YouTube subtitles',
          variant: 'warning',
        });
      return;
    }
    const currentMediaPath =
      appState.currentMediaPath?.trim() || appState.mpvClient?.currentVideoPath?.trim() || '';
    if (!isPlaybackActiveNow() || !currentMediaPath) {
      deps
        .getOverlayNotifications()
        .showConfiguredStatusNotification(
          'YouTube subtitle picker is only available during YouTube playback.',
          {
            title: 'YouTube subtitles',
            variant: 'warning',
          },
        );
      return;
    }
    await flow.openManualPicker({
      url: currentMediaPath,
    });
  }

  let quitWhenMpvClosesAfterBrowser = false;
  const isMpvPlayingMedia = (): boolean =>
    Boolean(appState.mpvClient?.connected && appState.currentMediaPath?.trim());
  const browserPlayback = createYoutubeBrowserPlaybackRuntime({
    isMpvPlaying: isMpvPlayingMedia,
    // Windows playback bootstraps its own mpv inside the playback flow.
    ensureMpvReady: async () =>
      process.platform === 'win32' ? true : await deps.ensureMpvConnectedForPlayback(),
    runPlaybackFlow: (url) => playback.runYoutubePlaybackFlow({ url, source: 'youtube-browser' }),
    appendToMpvPlaylist: (url) => {
      sendMpvCommandRuntime(appState.mpvClient, ['loadfile', url, 'append']);
    },
    notifyFailure: (message) =>
      deps
        .getOverlayNotifications()
        .showConfiguredStatusNotification(message, { title: 'YouTube' }),
    logWarn: (message) => logger.warn(message),
    onStartupFailed: () => quitAfterBrowserIfIdle(),
  });
  const browserWindow = createYoutubeBrowserWindowRuntime({
    ...deps.browserWindow,
    openVideo: (request) => browserPlayback.openVideo(request),
    logWarn: (message, error) => logger.warn(message, error),
    logDebug: (message) => logger.debug(message),
    onClosed: () => {
      if (!appState.initialArgs?.youtubeBrowser) return;
      quitWhenMpvClosesAfterBrowser = true;
      quitAfterBrowserIfIdle();
    },
  });
  // A process started just for the browser (`--youtube-browser`) exits once the window is closed,
  // but keeps the overlay alive until mpv closes if a video is playing or still starting. Rechecked
  // when a start fails, so a failed start after the window closed does not leave the app running.
  const quitAfterBrowserIfIdle = (): void => {
    if (!quitWhenMpvClosesAfterBrowser || browserWindow.isOpen()) return;
    if (isMpvPlayingMedia() || browserPlayback.isStartingPlayback()) return;
    deps.requestAppQuit();
  };

  return {
    flow,
    playback,
    mediaCache,
    mediaCachePlayback,
    whisperSubtitles,
    primarySubtitleNotification,
    browserPlayback,
    browserWindow,
    isPlaybackActiveNow,
    getCurrentMediaCacheSourceUrl,
    shouldRequireMediaCacheForCurrentPlayback,
    getCachedMediaPathForCurrentPlayback,
    openTrackPickerFromPlayback,
    /** True once a `--youtube-browser` process's window has closed, so mpv closing ends the app. */
    shouldQuitWhenMpvCloses: () => quitWhenMpvClosesAfterBrowser && !browserWindow.isOpen(),
  };
}
