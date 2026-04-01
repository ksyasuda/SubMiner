import fs from 'node:fs';
import path from 'node:path';

import type { MpvIpcClient } from '../core/services/mpv';
import type {
  JimakuLanguagePreference,
  ResolvedConfig,
  SecondarySubMode,
  SubsyncManualPayload,
} from '../types';
import type { ConfigService } from '../config';
import type { RuntimeOptionsManager } from '../runtime-options';
import type { AppState } from './state';
import type { OverlayHostedModal } from '../shared/ipc/contracts';
import { createStartupSupportRuntime, type StartupSupportRuntime } from './startup-support-runtime';

export interface StartupSupportCoordinatorInput {
  platform: NodeJS.Platform;
  defaultImmersionDbPath: string;
  defaultJimakuLanguagePreference: JimakuLanguagePreference;
  defaultJimakuMaxEntryResults: number;
  defaultJimakuApiBaseUrl: string;
  jellyfinLangPref: string;
  getResolvedConfig: () => ResolvedConfig;
  appState: AppState;
  configService: Pick<ConfigService, 'reloadConfigStrict'>;
  actions: {
    sendMpvCommandRuntime: (client: MpvIpcClient, command: (string | number)[]) => void;
    showMpvOsd: (text: string) => void;
    openSubsyncManualPicker: (payload: SubsyncManualPayload) => void;
    refreshGlobalAndOverlayShortcuts: () => void;
    broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
    showDesktopNotification: (title: string, options: { body?: string }) => void;
    showErrorBox: (title: string, details: string) => void;
  };
  logger: StartupSupportRuntime['configHotReloadRuntime'] extends never
    ? never
    : Parameters<typeof createStartupSupportRuntime>[0]['logger'];
  watch: Parameters<typeof createStartupSupportRuntime>[0]['watch'];
  timers: Parameters<typeof createStartupSupportRuntime>[0]['timers'];
}

export interface StartupSupportFromMainStateInput {
  platform: NodeJS.Platform;
  defaultImmersionDbPath: string;
  defaultJimakuLanguagePreference: JimakuLanguagePreference;
  defaultJimakuMaxEntryResults: number;
  defaultJimakuApiBaseUrl: string;
  jellyfinLangPref: string;
  getResolvedConfig: () => ResolvedConfig;
  appState: AppState;
  configService: Pick<ConfigService, 'reloadConfigStrict'>;
  overlay: {
    broadcastToOverlayWindows: (channel: string, payload: unknown) => void;
    sendToActiveOverlayWindow: (
      channel: string,
      payload?: unknown,
      runtimeOptions?: {
        restoreOnModalClose?: OverlayHostedModal;
        preferModalWindow?: boolean;
      },
    ) => void;
  };
  shortcuts: {
    refreshGlobalAndOverlayShortcuts: () => void;
  };
  notifications: {
    showDesktopNotification: (title: string, options: { body?: string }) => void;
    showErrorBox: (title: string, details: string) => void;
  };
  logger: Parameters<typeof createStartupSupportRuntime>[0]['logger'];
  mpv: {
    sendMpvCommandRuntime: (client: MpvIpcClient, command: (string | number)[]) => void;
    showMpvOsd: (text: string) => void;
  };
}

export function createStartupSupportCoordinator(
  input: StartupSupportCoordinatorInput,
): StartupSupportRuntime {
  return createStartupSupportRuntime({
    platform: input.platform,
    defaultImmersionDbPath: input.defaultImmersionDbPath,
    defaultJimakuLanguagePreference: input.defaultJimakuLanguagePreference,
    defaultJimakuMaxEntryResults: input.defaultJimakuMaxEntryResults,
    defaultJimakuApiBaseUrl: input.defaultJimakuApiBaseUrl,
    jellyfinLangPref: input.jellyfinLangPref,
    getResolvedConfig: () => input.getResolvedConfig(),
    appState: {
      immersionTracker: input.appState.immersionTracker,
      mpvClient: input.appState.mpvClient,
      currentMediaPath: input.appState.currentMediaPath,
      currentMediaTitle: input.appState.currentMediaTitle,
      runtimeOptionsManager: input.appState.runtimeOptionsManager as RuntimeOptionsManager | null,
      subsyncInProgress: input.appState.subsyncInProgress,
      keybindings: input.appState.keybindings,
      ankiIntegration: input.appState.ankiIntegration,
    },
    mpv: {
      sendMpvCommandRuntime: (client, command) =>
        input.actions.sendMpvCommandRuntime(client as MpvIpcClient, command),
      showMpvOsd: (text) => input.actions.showMpvOsd(text),
    },
    config: {
      reloadConfigStrict: () => input.configService.reloadConfigStrict(),
    },
    subsync: {
      openManualPicker: (payload) => input.actions.openSubsyncManualPicker(payload),
    },
    hotReload: {
      setSecondarySubMode: (mode: SecondarySubMode) => {
        input.appState.secondarySubMode = mode;
      },
      refreshGlobalAndOverlayShortcuts: () => input.actions.refreshGlobalAndOverlayShortcuts(),
      broadcastToOverlayWindows: (channel, payload) =>
        input.actions.broadcastToOverlayWindows(channel, payload),
    },
    notifications: {
      showDesktopNotification: (title, options) =>
        input.actions.showDesktopNotification(title, options),
      showErrorBox: (title, details) => input.actions.showErrorBox(title, details),
    },
    logger: input.logger,
    watch: input.watch,
    timers: input.timers,
  });
}

export function createStartupSupportFromMainState(
  input: StartupSupportFromMainStateInput,
): StartupSupportRuntime {
  return createStartupSupportCoordinator({
    platform: input.platform,
    defaultImmersionDbPath: input.defaultImmersionDbPath,
    defaultJimakuLanguagePreference: input.defaultJimakuLanguagePreference,
    defaultJimakuMaxEntryResults: input.defaultJimakuMaxEntryResults,
    defaultJimakuApiBaseUrl: input.defaultJimakuApiBaseUrl,
    jellyfinLangPref: input.jellyfinLangPref,
    getResolvedConfig: () => input.getResolvedConfig(),
    appState: input.appState,
    configService: input.configService,
    actions: {
      sendMpvCommandRuntime: (client, command) => input.mpv.sendMpvCommandRuntime(client, command),
      showMpvOsd: (text) => input.mpv.showMpvOsd(text),
      openSubsyncManualPicker: (payload) => {
        input.overlay.sendToActiveOverlayWindow('subsync:open-manual', payload, {
          restoreOnModalClose: 'subsync',
        });
      },
      refreshGlobalAndOverlayShortcuts: () => {
        input.shortcuts.refreshGlobalAndOverlayShortcuts();
      },
      broadcastToOverlayWindows: (channel, payload) => {
        input.overlay.broadcastToOverlayWindows(channel, payload);
      },
      showDesktopNotification: (title, options) =>
        input.notifications.showDesktopNotification(title, options),
      showErrorBox: (title, details) => input.notifications.showErrorBox(title, details),
    },
    logger: input.logger,
    watch: {
      fileExists: (targetPath) => fs.existsSync(targetPath),
      dirname: (targetPath) => path.dirname(targetPath),
      watchPath: (targetPath, listener) => fs.watch(targetPath, listener),
    },
    timers: {
      setTimeout: (callback, delayMs) => setTimeout(callback, delayMs),
      clearTimeout: (timeout) => clearTimeout(timeout),
    },
  });
}
