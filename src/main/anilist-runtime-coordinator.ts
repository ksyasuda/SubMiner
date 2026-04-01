import path from 'node:path';

import { app, BrowserWindow, shell } from 'electron';

import { DEFAULT_MIN_WATCH_RATIO } from '../shared/watch-threshold';
import type { ResolvedConfig } from '../types';
import {
  guessAnilistMediaInfo,
  updateAnilistPostWatchProgress,
} from '../core/services/anilist/anilist-updater';
import type { AnilistSetupWindowLike } from './anilist-runtime';
import { createAnilistRuntime } from './anilist-runtime';
import {
  isAllowedAnilistExternalUrl,
  isAllowedAnilistSetupNavigationUrl,
} from './anilist-url-guard';

export interface AnilistRuntimeCoordinatorInput {
  getResolvedConfig: () => ResolvedConfig;
  isTrackingEnabled: (config: ResolvedConfig) => boolean;
  tokenStore: Parameters<typeof createAnilistRuntime>[0]['tokenStore'];
  updateQueue: Parameters<typeof createAnilistRuntime>[0]['updateQueue'];
  appState: {
    currentMediaPath: string | null;
    currentMediaTitle: string | null;
    mpvClient: {
      currentTimePos?: number | null;
      requestProperty: (name: string) => Promise<unknown>;
    } | null;
    anilistSetupWindow: BrowserWindow | null;
  };
  dictionarySupport: {
    resolveMediaPathForJimaku: (mediaPath: string | null) => string | null;
  };
  actions: {
    showMpvOsd: (message: string) => void;
    showDesktopNotification: (title: string, options: { body?: string }) => void;
  };
  logger: {
    info: (message: string) => void;
    warn: (message: string, details?: unknown) => void;
    error: (message: string, error?: unknown) => void;
    debug: (message: string, details?: unknown) => void;
  };
  constants: {
    authorizeUrl: string;
    clientId: string;
    responseType: string;
    redirectUri: string;
    developerSettingsUrl: string;
    durationRetryIntervalMs: number;
    minWatchSeconds: number;
    maxAttemptedUpdateKeys: number;
  };
}

export function createAnilistRuntimeCoordinator(input: AnilistRuntimeCoordinatorInput) {
  return createAnilistRuntime<ResolvedConfig, AnilistSetupWindowLike>({
    getResolvedConfig: () => input.getResolvedConfig(),
    isTrackingEnabled: (config) => input.isTrackingEnabled(config),
    tokenStore: input.tokenStore,
    updateQueue: input.updateQueue,
    getCurrentMediaPath: () => input.appState.currentMediaPath,
    getCurrentMediaTitle: () => input.appState.currentMediaTitle,
    getWatchedSeconds: () => input.appState.mpvClient?.currentTimePos ?? Number.NaN,
    hasMpvClient: () => Boolean(input.appState.mpvClient),
    requestMpvDuration: async () => input.appState.mpvClient?.requestProperty('duration'),
    resolveMediaPathForJimaku: (currentMediaPath) =>
      input.dictionarySupport.resolveMediaPathForJimaku(currentMediaPath),
    guessAnilistMediaInfo: (mediaPath, mediaTitle) => guessAnilistMediaInfo(mediaPath, mediaTitle),
    updateAnilistPostWatchProgress: (accessToken, title, episode) =>
      updateAnilistPostWatchProgress(accessToken, title, episode),
    createBrowserWindow: (options) => {
      const window = new BrowserWindow(options);
      input.appState.anilistSetupWindow = window;
      window.on('closed', () => {
        input.appState.anilistSetupWindow = null;
      });
      return window as unknown as AnilistSetupWindowLike;
    },
    authorizeUrl: input.constants.authorizeUrl,
    clientId: input.constants.clientId,
    responseType: input.constants.responseType,
    redirectUri: input.constants.redirectUri,
    developerSettingsUrl: input.constants.developerSettingsUrl,
    isAllowedExternalUrl: (url) => isAllowedAnilistExternalUrl(url),
    isAllowedNavigationUrl: (url) => isAllowedAnilistSetupNavigationUrl(url),
    openExternal: (url) => shell.openExternal(url),
    showMpvOsd: (message) => input.actions.showMpvOsd(message),
    showDesktopNotification: (title, options) =>
      input.actions.showDesktopNotification(title, options),
    logInfo: (message) => input.logger.info(message),
    logWarn: (message, details) => input.logger.warn(message, details),
    logError: (message, error) => input.logger.error(message, error),
    logDebug: (message, details) => input.logger.debug(message, details),
    isDefaultApp: () => Boolean(process.defaultApp),
    getArgv: () => process.argv,
    execPath: process.execPath,
    resolvePath: (value) => path.resolve(value),
    setAsDefaultProtocolClient: (scheme, appPath, args) =>
      appPath
        ? app.setAsDefaultProtocolClient(scheme, appPath, args)
        : app.setAsDefaultProtocolClient(scheme),
    now: () => Date.now(),
    durationRetryIntervalMs: input.constants.durationRetryIntervalMs,
    minWatchSeconds: input.constants.minWatchSeconds,
    minWatchRatio: DEFAULT_MIN_WATCH_RATIO,
    maxAttemptedUpdateKeys: input.constants.maxAttemptedUpdateKeys,
  });
}
