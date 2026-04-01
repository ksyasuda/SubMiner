import type { BrowserWindow } from 'electron';

import type { ConfigService } from '../config';
import type { ResolvedConfig } from '../types';
import type { AppState } from './state';
import { createFirstRunRuntimeCoordinator } from './first-run-runtime-coordinator';
import { createStartupSupportFromMainState } from './startup-support-coordinator';
import { createYoutubeRuntimeFromMainState } from './youtube-runtime-coordinator';
import { createOverlayMpvSubtitleSuppressionRuntime } from './runtime/overlay-mpv-sub-visibility';
import { createDiscordPresenceRuntimeFromMainState } from './runtime/discord-presence-runtime';
import type { OverlayGeometryRuntime } from './overlay-geometry-runtime';
import type { OverlayHostedModal } from '../shared/ipc/contracts';
import type { SubtitleRuntime } from './subtitle-runtime';
import type { OverlayUiRuntime } from './overlay-ui-runtime';

export interface MainEarlyRuntimeInput {
  platform: NodeJS.Platform;
  configDir: string;
  homeDir: string;
  xdgConfigHome?: string;
  binaryPath: string;
  appPath: string;
  resourcesPath: string;
  appDataDir: string;
  desktopDir: string;
  defaultImmersionDbPath: string;
  defaultJimakuLanguagePreference: ResolvedConfig['jimaku']['languagePreference'];
  defaultJimakuMaxEntryResults: number;
  defaultJimakuApiBaseUrl: string;
  jellyfinLangPref: string;
  youtube: {
    directPlaybackFormat: string;
    mpvYtdlFormat: string;
    autoLaunchTimeoutMs: number;
    connectTimeoutMs: number;
    logPath: string;
  };
  discordPresenceAppId: string;
  appState: AppState;
  getResolvedConfig: () => ResolvedConfig;
  getFallbackDiscordMediaDurationSec: () => number | null;
  configService: Pick<ConfigService, 'reloadConfigStrict'>;
  overlay: {
    overlayManager: {
      getVisibleOverlayVisible: () => boolean;
      broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
      getMainWindow: () => BrowserWindow | null;
    };
    overlayModalRuntime: {
      sendToActiveOverlayWindow: (
        channel: string,
        payload?: unknown,
        runtimeOptions?: {
          restoreOnModalClose?: OverlayHostedModal;
          preferModalWindow?: boolean;
        },
      ) => void;
    };
    getOverlayUi: () => OverlayUiRuntime<BrowserWindow> | null;
    getOverlayGeometry: () => OverlayGeometryRuntime<BrowserWindow>;
    ensureTray: () => void;
    hasTray: () => boolean;
  };
  yomitan: {
    ensureYomitanExtensionLoaded: () => Promise<unknown>;
    getParserRuntimeDeps: () => Parameters<
      typeof import('../core/services').getYomitanDictionaryInfo
    >[0];
    openYomitanSettings: () => boolean;
  };
  subtitle: {
    getSubtitle: () => SubtitleRuntime;
  };
  tokenization: {
    startTokenizationWarmups: () => Promise<void>;
    getGate: Parameters<typeof createYoutubeRuntimeFromMainState>[0]['tokenization']['getGate'];
  };
  appReady: {
    ensureYoutubePlaybackRuntimeReady: () => Promise<void>;
  };
  shortcuts: {
    refreshGlobalAndOverlayShortcuts: () => void;
  };
  notifications: {
    showDesktopNotification: (title: string, options: { body?: string }) => void;
    showErrorBox: (title: string, content: string) => void;
  };
  mpv: {
    sendMpvCommandRuntime: (client: AppState['mpvClient'], command: (string | number)[]) => void;
    setSubVisibility: (visible: boolean) => void;
    showMpvOsd: (text: string) => void;
  };
  actions: {
    requestAppQuit: () => void;
    writeShortcutLink: (
      shortcutPath: string,
      operation: 'create' | 'update' | 'replace',
      details: {
        target: string;
        args?: string;
        cwd?: string;
        description?: string;
        icon?: string;
        iconIndex?: number;
      },
    ) => boolean;
  };
  logger: {
    error: (message: string, error?: unknown) => void;
    info: (message: string, ...args: unknown[]) => void;
    warn: (message: string, error?: unknown) => void;
    debug: (message: string, meta?: unknown) => void;
  };
}

export function createMainEarlyRuntime(input: MainEarlyRuntimeInput) {
  const firstRun = createFirstRunRuntimeCoordinator({
    platform: input.platform,
    configDir: input.configDir,
    homeDir: input.homeDir,
    xdgConfigHome: input.xdgConfigHome,
    binaryPath: input.binaryPath,
    appPath: input.appPath,
    resourcesPath: input.resourcesPath,
    appDataDir: input.appDataDir,
    desktopDir: input.desktopDir,
    appState: input.appState,
    getResolvedConfig: () => input.getResolvedConfig(),
    yomitan: input.yomitan,
    overlay: {
      ensureTray: () => input.overlay.ensureTray(),
      hasTray: () => input.overlay.hasTray(),
    },
    actions: {
      writeShortcutLink: (shortcutPath, operation, details) =>
        input.actions.writeShortcutLink(shortcutPath, operation, details),
      requestAppQuit: () => input.actions.requestAppQuit(),
    },
    logger: {
      error: (message, error) => input.logger.error(message, error),
      info: (message, ...args) => input.logger.info(message, ...args),
    },
  });

  const { discordPresenceRuntime, initializeDiscordPresenceService } =
    createDiscordPresenceRuntimeFromMainState({
      appId: input.discordPresenceAppId,
      appState: input.appState,
      getResolvedConfig: () => input.getResolvedConfig(),
      getFallbackMediaDurationSec: () => input.getFallbackDiscordMediaDurationSec(),
      logger: {
        debug: (message, meta) => input.logger.debug(message, meta),
      },
    });

  const overlaySubtitleSuppression = createOverlayMpvSubtitleSuppressionRuntime({
    appState: input.appState,
    getVisibleOverlayVisible: () => input.overlay.overlayManager.getVisibleOverlayVisible(),
    setMpvSubVisibility: (visible) => input.mpv.setSubVisibility(visible),
    logWarn: (message, error) => input.logger.warn(message, error),
  });

  const startupSupport = createStartupSupportFromMainState({
    platform: input.platform,
    defaultImmersionDbPath: input.defaultImmersionDbPath,
    defaultJimakuLanguagePreference: input.defaultJimakuLanguagePreference,
    defaultJimakuMaxEntryResults: input.defaultJimakuMaxEntryResults,
    defaultJimakuApiBaseUrl: input.defaultJimakuApiBaseUrl,
    jellyfinLangPref: input.jellyfinLangPref,
    getResolvedConfig: () => input.getResolvedConfig(),
    appState: input.appState,
    configService: input.configService,
    overlay: {
      broadcastToOverlayWindows: (channel, payload) =>
        input.overlay.overlayManager.broadcastToOverlayWindows(channel, payload),
      sendToActiveOverlayWindow: (channel, payload, runtimeOptions) =>
        input.overlay.overlayModalRuntime.sendToActiveOverlayWindow(
          channel,
          payload,
          runtimeOptions,
        ),
    },
    shortcuts: {
      refreshGlobalAndOverlayShortcuts: () => input.shortcuts.refreshGlobalAndOverlayShortcuts(),
    },
    notifications: {
      showDesktopNotification: (title, options) =>
        input.notifications.showDesktopNotification(title, options),
      showErrorBox: (title, details) => input.notifications.showErrorBox(title, details),
    },
    logger: {
      debug: (message) => input.logger.debug(message),
      info: (message) => input.logger.info(message),
      warn: (message, error) => input.logger.warn(message, error),
    },
    mpv: {
      sendMpvCommandRuntime: (client, command) => input.mpv.sendMpvCommandRuntime(client, command),
      showMpvOsd: (text) => input.mpv.showMpvOsd(text),
    },
  });

  const youtube = createYoutubeRuntimeFromMainState({
    platform: input.platform,
    directPlaybackFormat: input.youtube.directPlaybackFormat,
    mpvYtdlFormat: input.youtube.mpvYtdlFormat,
    autoLaunchTimeoutMs: input.youtube.autoLaunchTimeoutMs,
    connectTimeoutMs: input.youtube.connectTimeoutMs,
    logPath: input.youtube.logPath,
    appState: input.appState,
    overlay: {
      getOverlayUi: () => input.overlay.getOverlayUi(),
      getMainWindow: () => input.overlay.overlayManager.getMainWindow(),
      getOverlayGeometry: () => input.overlay.getOverlayGeometry(),
      broadcastToOverlayWindows: (channel, payload) =>
        input.overlay.overlayManager.broadcastToOverlayWindows(channel, payload),
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
    getResolvedConfig: () => input.getResolvedConfig(),
    notifications: {
      showDesktopNotification: (title, options) =>
        input.notifications.showDesktopNotification(title, options),
      showErrorBox: (title, content) => input.notifications.showErrorBox(title, content),
    },
    mpv: {
      sendMpvCommand: (command) =>
        input.mpv.sendMpvCommandRuntime(input.appState.mpvClient, command),
      showMpvOsd: (message) => input.mpv.showMpvOsd(message),
    },
    logger: {
      info: (message) => input.logger.info(message),
      warn: (message, error) => input.logger.warn(message, error),
      debug: (message) => input.logger.debug(message),
    },
  });

  return {
    firstRun,
    discordPresenceRuntime,
    initializeDiscordPresenceService,
    overlaySubtitleSuppression,
    startupSupport,
    youtube,
  };
}
