import { IPC_CHANNELS } from '../shared/ipc/contracts';
import {
  acquireYoutubeSubtitleTrack,
  acquireYoutubeSubtitleTracks,
} from '../core/services/youtube/generate';
import { resolveYoutubePlaybackUrl } from '../core/services/youtube/playback-resolve';
import { probeYoutubeTracks } from '../core/services/youtube/track-probe';
import type { ResolvedConfig } from '../types';
import type { AppState } from './state';
import type { OverlayGeometryRuntime } from './overlay-geometry-runtime';
import type { OverlayUiRuntime } from './overlay-ui-runtime';
import type { SubtitleRuntime } from './subtitle-runtime';
import { createYoutubeRuntime } from './youtube-runtime';
import { createYoutubeRuntimeInput } from './youtube-runtime-bootstrap';

export interface YoutubeRuntimeCoordinatorInput {
  appState: {
    getMpvClient: () => Parameters<
      typeof createYoutubeRuntimeInput
    >[0]['appState']['getMpvClient'] extends () => infer T
      ? T
      : never;
    getCurrentMediaPath: () => string | null;
    getPlaybackPaused: () => boolean | null;
    getWindowTracker: () => Parameters<
      typeof createYoutubeRuntimeInput
    >[0]['appState']['getWindowTracker'] extends () => infer T
      ? T
      : never;
    getAnkiIntegration: () => Parameters<
      typeof createYoutubeRuntimeInput
    >[0]['appState']['getAnkiIntegration'] extends () => infer T
      ? T
      : never;
    getSocketPath: () => string;
  };
  overlay: {
    getOverlayUi: () => Parameters<
      typeof createYoutubeRuntimeInput
    >[0]['overlay']['getOverlayUi'] extends () => infer T
      ? T
      : never;
    getMainWindow: () => Parameters<
      typeof createYoutubeRuntimeInput
    >[0]['overlay']['getMainWindow'] extends () => infer T
      ? T
      : never;
    getOverlayGeometry: () => Parameters<
      typeof createYoutubeRuntimeInput
    >[0]['overlay']['getOverlayGeometry'] extends () => infer T
      ? T
      : never;
    broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
  };
  subtitle: {
    getSubtitle: Parameters<typeof createYoutubeRuntimeInput>[0]['getSubtitle'];
  };
  tokenization: {
    startTokenizationWarmups: () => Promise<void>;
    getGate: Parameters<typeof createYoutubeRuntimeInput>[0]['tokenization']['getGate'];
  };
  appReady: {
    ensureYoutubePlaybackRuntimeReady: () => Promise<void>;
  };
  services: {
    sendMpvCommand: (command: (string | number)[]) => void;
    showMpvOsd: (message: string) => void;
    showDesktopNotification: (title: string, options: { body?: string }) => void;
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
    getNotificationType: () => string;
    getPrimarySubtitleLanguages: () => string[];
  };
}

export function createYoutubeRuntimeCoordinator(input: YoutubeRuntimeCoordinatorInput) {
  return createYoutubeRuntime(
    createYoutubeRuntimeInput({
      appState: {
        getMpvClient: () => input.appState.getMpvClient(),
        getCurrentMediaPath: () => input.appState.getCurrentMediaPath(),
        getPlaybackPaused: () => input.appState.getPlaybackPaused(),
        getWindowTracker: () => input.appState.getWindowTracker(),
        getAnkiIntegration: () => input.appState.getAnkiIntegration(),
      },
      overlay: {
        getOverlayUi: () => input.overlay.getOverlayUi(),
        getMainWindow: () => input.overlay.getMainWindow(),
        getOverlayGeometry: () => input.overlay.getOverlayGeometry(),
        broadcastYoutubePickerCancel: () => {
          input.overlay.broadcastToOverlayWindows(IPC_CHANNELS.event.youtubePickerCancel, null);
        },
      },
      getSubtitle: () => input.subtitle.getSubtitle(),
      tokenization: {
        startTokenizationWarmups: () => input.tokenization.startTokenizationWarmups(),
        getGate: () => input.tokenization.getGate(),
      },
      appReady: {
        ensureYoutubePlaybackRuntimeReady: () => input.appReady.ensureYoutubePlaybackRuntimeReady(),
      },
      services: {
        probeYoutubeTracks: (url) => probeYoutubeTracks(url),
        acquireYoutubeSubtitleTrack: (request) => acquireYoutubeSubtitleTrack(request),
        acquireYoutubeSubtitleTracks: (request) => acquireYoutubeSubtitleTracks(request),
        resolveYoutubePlaybackUrl: (url, format) => resolveYoutubePlaybackUrl(url, format),
        sendMpvCommand: (command) => input.services.sendMpvCommand(command),
        showMpvOsd: (message) => input.services.showMpvOsd(message),
        showDesktopNotification: (title, options) =>
          input.services.showDesktopNotification(title, options),
        showErrorBox: (title, content) => input.services.showErrorBox(title, content),
        logInfo: (message) => input.services.logInfo(message),
        logWarn: (message, error) => input.services.logWarn(message, error),
        logDebug: (message) => input.services.logDebug(message),
      },
      config: {
        platform: input.config.platform,
        directPlaybackFormat: input.config.directPlaybackFormat,
        mpvYtdlFormat: input.config.mpvYtdlFormat,
        autoLaunchTimeoutMs: input.config.autoLaunchTimeoutMs,
        connectTimeoutMs: input.config.connectTimeoutMs,
        logPath: input.config.logPath,
        getSocketPath: () => input.appState.getSocketPath(),
        getNotificationType: () => input.config.getNotificationType(),
        getPrimarySubtitleLanguages: () => input.config.getPrimarySubtitleLanguages(),
      },
    }),
  );
}

export interface YoutubeRuntimeFromMainStateInput {
  platform: NodeJS.Platform;
  directPlaybackFormat: string;
  mpvYtdlFormat: string;
  autoLaunchTimeoutMs: number;
  connectTimeoutMs: number;
  logPath: string;
  appState: Pick<
    AppState,
    | 'mpvClient'
    | 'currentMediaPath'
    | 'playbackPaused'
    | 'windowTracker'
    | 'ankiIntegration'
    | 'mpvSocketPath'
  >;
  overlay: {
    getOverlayUi: () => OverlayUiRuntime<Electron.BrowserWindow> | null;
    getMainWindow: () => Electron.BrowserWindow | null;
    getOverlayGeometry: () => OverlayGeometryRuntime<Electron.BrowserWindow>;
    broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
  };
  subtitle: {
    getSubtitle: () => SubtitleRuntime;
  };
  tokenization: {
    startTokenizationWarmups: () => Promise<void>;
    getGate: Parameters<typeof createYoutubeRuntimeInput>[0]['tokenization']['getGate'];
  };
  appReady: {
    ensureYoutubePlaybackRuntimeReady: () => Promise<void>;
  };
  getResolvedConfig: () => ResolvedConfig;
  notifications: {
    showDesktopNotification: (title: string, options: { body?: string }) => void;
    showErrorBox: (title: string, content: string) => void;
  };
  mpv: {
    sendMpvCommand: (command: (string | number)[]) => void;
    showMpvOsd: (message: string) => void;
  };
  logger: {
    info: (message: string) => void;
    warn: (message: string, error?: unknown) => void;
    debug: (message: string) => void;
  };
}

export function createYoutubeRuntimeFromMainState(input: YoutubeRuntimeFromMainStateInput) {
  return createYoutubeRuntimeCoordinator({
    appState: {
      getMpvClient: () => input.appState.mpvClient,
      getCurrentMediaPath: () => input.appState.currentMediaPath,
      getPlaybackPaused: () => input.appState.playbackPaused,
      getWindowTracker: () => input.appState.windowTracker,
      getAnkiIntegration: () => input.appState.ankiIntegration,
      getSocketPath: () => input.appState.mpvSocketPath,
    },
    overlay: {
      getOverlayUi: () => input.overlay.getOverlayUi(),
      getMainWindow: () => input.overlay.getMainWindow(),
      getOverlayGeometry: () => input.overlay.getOverlayGeometry(),
      broadcastToOverlayWindows: (channel, payload) => {
        input.overlay.broadcastToOverlayWindows(channel, payload);
      },
    },
    subtitle: {
      getSubtitle: () => input.subtitle.getSubtitle(),
    },
    tokenization: {
      startTokenizationWarmups: () => input.tokenization.startTokenizationWarmups(),
      getGate: () => input.tokenization.getGate(),
    },
    appReady: {
      ensureYoutubePlaybackRuntimeReady: () => input.appReady.ensureYoutubePlaybackRuntimeReady(),
    },
    services: {
      sendMpvCommand: (command) => input.mpv.sendMpvCommand(command),
      showMpvOsd: (message) => input.mpv.showMpvOsd(message),
      showDesktopNotification: (title, options) =>
        input.notifications.showDesktopNotification(title, options),
      showErrorBox: (title, content) => input.notifications.showErrorBox(title, content),
      logInfo: (message) => input.logger.info(message),
      logWarn: (message, error) => input.logger.warn(message, error),
      logDebug: (message) => input.logger.debug(message),
    },
    config: {
      platform: input.platform,
      directPlaybackFormat: input.directPlaybackFormat,
      mpvYtdlFormat: input.mpvYtdlFormat,
      autoLaunchTimeoutMs: input.autoLaunchTimeoutMs,
      connectTimeoutMs: input.connectTimeoutMs,
      logPath: input.logPath,
      getNotificationType: () => input.getResolvedConfig().ankiConnect.behavior.notificationType,
      getPrimarySubtitleLanguages: () => input.getResolvedConfig().youtube.primarySubLanguages,
    },
  });
}
