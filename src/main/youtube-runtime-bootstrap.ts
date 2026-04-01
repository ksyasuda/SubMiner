import os from 'node:os';
import path from 'node:path';

import type { OverlayHostedModal } from '../shared/ipc/contracts';
import type { WindowGeometry } from '../types';
import type { YoutubeRuntimeInput } from './youtube-runtime';
import { createWaitForMpvConnectedHandler } from './runtime/jellyfin-remote-connection';
import { createPrepareYoutubePlaybackInMpvHandler } from './runtime/youtube-playback-launch';
import { openYoutubeTrackPicker } from './runtime/youtube-picker-open';
import { createWindowsMpvLaunchDeps, launchWindowsMpv } from './runtime/windows-mpv-launch';

type MpvClientLike = {
  connected: boolean;
  currentVideoPath?: string | null;
  connect: () => void;
  requestProperty: (name: string) => Promise<unknown>;
  send: (payload: { command: Array<string | boolean> }) => void;
};

type AnkiIntegrationLike = {
  waitUntilReady: () => Promise<void>;
};

type WindowTrackerLike = {
  getGeometry: () => WindowGeometry | null;
  isTargetWindowFocused: () => boolean;
  isTracking: () => boolean;
};

type OverlayMainWindowLike = {
  isDestroyed: () => boolean;
  isFocused: () => boolean;
  focus: () => void;
  setIgnoreMouseEvents: (ignore: boolean) => void;
  webContents: {
    isFocused: () => boolean;
    focus: () => void;
  };
};

type OverlayUiLike = {
  sendToActiveOverlayWindow: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: {
      restoreOnModalClose?: OverlayHostedModal;
      preferModalWindow?: boolean;
    },
  ) => boolean;
  waitForModalOpen: (modal: OverlayHostedModal, timeoutMs: number) => Promise<boolean>;
  handleOverlayModalClosed: (modal: OverlayHostedModal) => void;
};

type OverlayGeometryLike = {
  geometryMatches: (left: WindowGeometry | null, right: WindowGeometry | null) => boolean;
  getLastOverlayWindowGeometry: () => WindowGeometry | null;
};

type SubtitleRuntimeLike = {
  refreshCurrentSubtitle: (text: string) => void;
  refreshSubtitleSidebarFromSource: (sourcePath: string) => Promise<void>;
};

type TokenizationGateLike = {
  waitUntilReady: (mediaPath: string | null) => Promise<void>;
};

export interface YoutubeRuntimeBootstrapInput {
  appState: {
    getMpvClient: () => MpvClientLike | null;
    getCurrentMediaPath: () => string | null;
    getPlaybackPaused: () => boolean | null;
    getWindowTracker: () => WindowTrackerLike | null;
    getAnkiIntegration: () => AnkiIntegrationLike | null;
  };
  overlay: {
    getOverlayUi: () => OverlayUiLike | null;
    getMainWindow: () => OverlayMainWindowLike | null;
    getOverlayGeometry: () => OverlayGeometryLike;
    broadcastYoutubePickerCancel: () => void;
  };
  getSubtitle: () => SubtitleRuntimeLike;
  tokenization: {
    startTokenizationWarmups: () => Promise<void>;
    getGate: () => TokenizationGateLike;
  };
  appReady: {
    ensureYoutubePlaybackRuntimeReady: () => Promise<void>;
  };
  services: {
    probeYoutubeTracks: YoutubeRuntimeInput['flow']['probeYoutubeTracks'];
    acquireYoutubeSubtitleTrack: YoutubeRuntimeInput['flow']['acquireYoutubeSubtitleTrack'];
    acquireYoutubeSubtitleTracks: YoutubeRuntimeInput['flow']['acquireYoutubeSubtitleTracks'];
    resolveYoutubePlaybackUrl: YoutubeRuntimeInput['playback']['resolveYoutubePlaybackUrl'];
    sendMpvCommand: YoutubeRuntimeInput['flow']['sendMpvCommand'];
    showMpvOsd: YoutubeRuntimeInput['showMpvOsd'];
    showDesktopNotification: YoutubeRuntimeInput['showDesktopNotification'];
    showErrorBox: (title: string, content: string) => void;
    logInfo: (message: string) => void;
    logWarn: (message: string, error?: unknown) => void;
    logDebug: (message: string) => void;
  };
  config: {
    platform: NodeJS.Platform;
    directPlaybackFormat: string;
    mpvYtdlFormat: string;
    autoLaunchTimeoutMs: number;
    connectTimeoutMs: number;
    logPath: string;
    getSocketPath: () => string;
    getNotificationType: () => YoutubeRuntimeInput['getNotificationType'] extends () => infer T
      ? T
      : string;
    getPrimarySubtitleLanguages: () => string[];
  };
}

export function createYoutubeRuntimeInput(
  input: YoutubeRuntimeBootstrapInput,
): YoutubeRuntimeInput {
  const prepareYoutubePlaybackInMpv = createPrepareYoutubePlaybackInMpvHandler({
    requestPath: async () => {
      const client = input.appState.getMpvClient();
      if (!client) return null;
      const value = await client.requestProperty('path').catch(() => null);
      return typeof value === 'string' ? value : null;
    },
    requestProperty: async (name) => {
      const client = input.appState.getMpvClient();
      if (!client) return null;
      return await client.requestProperty(name);
    },
    sendMpvCommand: (command) => {
      input.services.sendMpvCommand(command);
    },
    wait: (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
  });

  const waitForYoutubeMpvConnected = createWaitForMpvConnectedHandler({
    getMpvClient: () => input.appState.getMpvClient(),
    now: () => Date.now(),
    sleep: (delayMs) => new Promise((resolve) => setTimeout(resolve, delayMs)),
  });

  return {
    flow: {
      probeYoutubeTracks: (url) => input.services.probeYoutubeTracks(url),
      acquireYoutubeSubtitleTrack: (request) => input.services.acquireYoutubeSubtitleTrack(request),
      acquireYoutubeSubtitleTracks: (request) =>
        input.services.acquireYoutubeSubtitleTracks(request),
      openPicker: async (payload) =>
        await openYoutubeTrackPicker(
          {
            sendToActiveOverlayWindow: (channel, nextPayload, runtimeOptions) =>
              input.overlay
                .getOverlayUi()
                ?.sendToActiveOverlayWindow(channel, nextPayload, runtimeOptions) ?? false,
            waitForModalOpen: (modal, timeoutMs) =>
              input.overlay.getOverlayUi()?.waitForModalOpen(modal, timeoutMs) ??
              Promise.resolve(false),
            logWarn: (message) => input.services.logWarn(message),
          },
          payload,
        ),
      pauseMpv: () => {
        input.services.sendMpvCommand(['set_property', 'pause', 'yes']);
      },
      resumeMpv: () => {
        input.services.sendMpvCommand(['set_property', 'pause', 'no']);
      },
      sendMpvCommand: (command) => {
        input.services.sendMpvCommand(command);
      },
      requestMpvProperty: async (name) => {
        const client = input.appState.getMpvClient();
        if (!client) return null;
        return await client.requestProperty(name);
      },
      refreshCurrentSubtitle: (text) => {
        input.getSubtitle().refreshCurrentSubtitle(text);
      },
      refreshSubtitleSidebarSource: async (sourcePath) => {
        await input.getSubtitle().refreshSubtitleSidebarFromSource(sourcePath);
      },
      startTokenizationWarmups: async () => {
        await input.tokenization.startTokenizationWarmups();
      },
      waitForTokenizationReady: async () => {
        const currentMediaPath =
          input.appState.getCurrentMediaPath()?.trim() ||
          input.appState.getMpvClient()?.currentVideoPath?.trim() ||
          null;
        await input.tokenization.getGate().waitUntilReady(currentMediaPath);
      },
      waitForAnkiReady: async () => {
        const integration = input.appState.getAnkiIntegration();
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
          input.services.logWarn(
            'Continuing YouTube playback before AnkiConnect integration reported ready:',
            error instanceof Error ? error.message : String(error),
          );
        }
      },
      wait: (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
      waitForPlaybackWindowReady: async () => {
        const deadline = Date.now() + 4000;
        let stableGeometry: WindowGeometry | null = null;
        let stableSinceMs = 0;
        while (Date.now() < deadline) {
          const tracker = input.appState.getWindowTracker();
          const trackerGeometry = tracker?.getGeometry() ?? null;
          const mediaPath =
            input.appState.getCurrentMediaPath()?.trim() ||
            input.appState.getMpvClient()?.currentVideoPath?.trim() ||
            '';
          const trackerFocused = tracker?.isTargetWindowFocused() ?? false;
          if (tracker && tracker.isTracking() && trackerGeometry && trackerFocused && mediaPath) {
            if (
              !input.overlay.getOverlayGeometry().geometryMatches(stableGeometry, trackerGeometry)
            ) {
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
        input.services.logWarn(
          'Timed out waiting for tracked playback window focus/media readiness before opening YouTube subtitle picker.',
        );
      },
      waitForOverlayGeometryReady: async () => {
        const deadline = Date.now() + 4000;
        while (Date.now() < deadline) {
          const trackerGeometry = input.appState.getWindowTracker()?.getGeometry() ?? null;
          if (
            trackerGeometry &&
            input.overlay
              .getOverlayGeometry()
              .geometryMatches(
                input.overlay.getOverlayGeometry().getLastOverlayWindowGeometry(),
                trackerGeometry,
              )
          ) {
            return;
          }
          await new Promise((resolve) => setTimeout(resolve, 50));
        }
        input.services.logWarn(
          'Timed out waiting for overlay geometry to match tracked playback window.',
        );
      },
      focusOverlayWindow: () => {
        const mainWindow = input.overlay.getMainWindow();
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
      showMpvOsd: (text) => input.services.showMpvOsd(text),
      warn: (message) => input.services.logWarn(message),
      log: (message) => input.services.logInfo(message),
      getYoutubeOutputDir: () => path.join(os.homedir(), '.cache', 'subminer', 'youtube-subs'),
    },
    playback: {
      platform: input.config.platform,
      directPlaybackFormat: input.config.directPlaybackFormat,
      mpvYtdlFormat: input.config.mpvYtdlFormat,
      autoLaunchTimeoutMs: input.config.autoLaunchTimeoutMs,
      connectTimeoutMs: input.config.connectTimeoutMs,
      getSocketPath: () => input.config.getSocketPath(),
      getMpvConnected: () => Boolean(input.appState.getMpvClient()?.connected),
      ensureYoutubePlaybackRuntimeReady: async () => {
        await input.appReady.ensureYoutubePlaybackRuntimeReady();
      },
      resolveYoutubePlaybackUrl: (url, format) =>
        input.services.resolveYoutubePlaybackUrl(url, format),
      launchWindowsMpv: (playbackUrl, args) =>
        launchWindowsMpv(
          [playbackUrl],
          createWindowsMpvLaunchDeps({
            showError: (title, content) => input.services.showErrorBox(title, content),
          }),
          [...args, `--log-file=${input.config.logPath}`],
        ),
      waitForYoutubeMpvConnected: (timeoutMs) => waitForYoutubeMpvConnected(timeoutMs),
      prepareYoutubePlaybackInMpv: (request) => prepareYoutubePlaybackInMpv(request),
      logInfo: (message) => input.services.logInfo(message),
      logWarn: (message) => input.services.logWarn(message),
      schedule: (callback, delayMs) => setTimeout(callback, delayMs),
      clearScheduled: (timer) => clearTimeout(timer),
    },
    autoplay: {
      getCurrentMediaPath: () => input.appState.getCurrentMediaPath(),
      getCurrentVideoPath: () => input.appState.getMpvClient()?.currentVideoPath ?? null,
      getPlaybackPaused: () => input.appState.getPlaybackPaused(),
      getMpvClient: () => input.appState.getMpvClient(),
      signalPluginAutoplayReady: () => {
        input.services.sendMpvCommand(['script-message', 'subminer-autoplay-ready']);
      },
      schedule: (callback, delayMs) => setTimeout(callback, delayMs),
      logDebug: (message) => input.services.logDebug(message),
    },
    notification: {
      getPrimarySubtitleLanguages: () => input.config.getPrimarySubtitleLanguages(),
      schedule: (callback, delayMs) => setTimeout(callback, delayMs),
      clearSchedule: (timer) => clearTimeout(timer as ReturnType<typeof setTimeout>),
    },
    getNotificationType: () => input.config.getNotificationType(),
    getCurrentMediaPath: () => input.appState.getCurrentMediaPath(),
    getCurrentVideoPath: () => input.appState.getMpvClient()?.currentVideoPath ?? null,
    showMpvOsd: (message) => input.services.showMpvOsd(message),
    showDesktopNotification: (title, options) =>
      input.services.showDesktopNotification(title, options),
    broadcastYoutubePickerCancel: () => {
      input.overlay.broadcastYoutubePickerCancel();
    },
    closeYoutubePickerModal: () => {
      input.overlay.getOverlayUi()?.handleOverlayModalClosed('youtube-track-picker');
    },
    logWarn: (message) => input.services.logWarn(message),
  };
}
