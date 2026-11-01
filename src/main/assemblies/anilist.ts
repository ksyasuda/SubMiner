import path from 'node:path';
import type { App, BrowserWindow, BrowserWindowConstructorOptions, Shell } from 'electron';
import type { ConfigService } from '../../config/service';
import type { createAnilistTokenStore } from '../../core/services/anilist/anilist-token-store';
import type { createAnilistUpdateQueue } from '../../core/services/anilist/anilist-update-queue';
import {
  guessAnilistMediaInfo,
  updateAnilistPostWatchProgress,
} from '../../core/services/anilist/anilist-updater';
import type { createAnilistRateLimiter } from '../../core/services/anilist/rate-limiter';
import type { Logger } from '../../logger';
import { DEFAULT_MIN_WATCH_RATIO } from '../../shared/watch-threshold';
import type { ResolvedConfig } from '../../types';
import {
  isAllowedAnilistExternalUrl,
  isAllowedAnilistSetupNavigationUrl,
} from '../anilist-url-guard';
import { composeAnilistSetupHandlers, composeAnilistTrackingHandlers } from '../runtime/composers';
import {
  buildAnilistSetupUrl,
  consumeAnilistSetupCallbackUrl,
  createAnilistStateRuntime,
  createMaybeFocusExistingAnilistSetupWindowHandler,
  createOpenAnilistSetupWindowHandler,
  isAnilistTrackingEnabled,
  loadAnilistManualTokenEntry,
  openAnilistSetupInBrowser,
  rememberAnilistAttemptedUpdateKey,
} from '../runtime/domains/anilist';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import { createCreateAnilistSetupWindowHandler } from '../runtime/setup-window-factory';
import type { CliSlice, IpcSlice } from './slices';
import {
  type AnilistMediaGuessRuntimeState,
  type AppState,
  createInitialAnilistMediaGuessRuntimeState,
  createInitialAnilistUpdateInFlightState,
  transitionAnilistClientSecretState,
  transitionAnilistMediaGuessRuntimeState,
  transitionAnilistRetryQueueLastAttemptAt,
  transitionAnilistRetryQueueLastError,
  transitionAnilistRetryQueueState,
  transitionAnilistUpdateInFlightState,
} from '../state';

const ANILIST_SETUP_CLIENT_ID_URL = 'https://anilist.co/api/v2/oauth/authorize';
const ANILIST_SETUP_RESPONSE_TYPE = 'token';
const ANILIST_DEFAULT_CLIENT_ID = '36084';
const ANILIST_REDIRECT_URI = 'https://anilist.subminer.moe/';
const ANILIST_DEVELOPER_SETTINGS_URL = 'https://anilist.co/settings/developer';
const ANILIST_UPDATE_MIN_WATCH_SECONDS = 10 * 60;
const ANILIST_DURATION_RETRY_INTERVAL_MS = 15_000;
const ANILIST_MAX_ATTEMPTED_UPDATE_KEYS = 1000;

type TrackingComposerOptions = Parameters<typeof composeAnilistTrackingHandlers>[0];

/**
 * What the AniList domain needs from the rest of the app. Siblings built after this assembly in
 * `main.ts` are reached through getters or callbacks, which are only called once the app runs.
 */
export interface AnilistAssemblyDeps {
  appState: Pick<
    AppState,
    | 'mpvClient'
    | 'currentMediaPath'
    | 'currentMediaTitle'
    | 'anilistClientSecretState'
    | 'anilistRetryQueueState'
    | 'anilistSetupPageOpened'
    | 'anilistSetupWindow'
  >;
  configService: Pick<ConfigService, 'getConfig'>;
  logger: Pick<Logger, 'info' | 'warn' | 'debug' | 'error'>;
  anilistTokenStore: ReturnType<typeof createAnilistTokenStore>;
  anilistUpdateQueue: ReturnType<typeof createAnilistUpdateQueue>;
  /** Shared with the stats cover-art fetcher, so `main.ts` owns it. */
  anilistRateLimiter: ReturnType<typeof createAnilistRateLimiter>;
  getOverlayNotifications: () => Pick<
    ReturnType<typeof createOverlayNotificationsRuntime>,
    | 'getConfiguredStatusNotificationType'
    | 'showConfiguredStatusNotification'
    | 'showOverlayNotification'
  >;
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
  resolveMediaPathForJimaku: TrackingComposerOptions['ensureMediaGuessMainDeps']['resolveMediaPathForJimaku'];
  resolvePinnedMediaId: NonNullable<
    TrackingComposerOptions['maybeRunPostWatchUpdateMainDeps']['resolvePinnedAnilistMediaId']
  >;
  createBrowserWindow: (options: BrowserWindowConstructorOptions) => BrowserWindow;
  shell: Pick<Shell, 'openExternal'>;
  app: Pick<App, 'setAsDefaultProtocolClient'>;
}

export type AnilistAssembly = ReturnType<typeof createAnilistAssembly>;

/**
 * Owns AniList login/setup, token and client-secret state, media tracking, and post-watch
 * progress updates with their retry queue. `main.ts` builds it once and calls through the result.
 */
export function createAnilistAssembly(deps: AnilistAssemblyDeps) {
  const {
    appState,
    configService,
    logger,
    anilistTokenStore,
    anilistUpdateQueue,
    anilistRateLimiter,
  } = deps;
  let anilistMediaGuessRuntimeState: AnilistMediaGuessRuntimeState =
    createInitialAnilistMediaGuessRuntimeState();
  let anilistUpdateInFlightState = createInitialAnilistUpdateInFlightState();
  const anilistAttemptedUpdateKeys = new Set<string>();
  let anilistCachedAccessToken: string | null = null;

  const anilistStateRuntime = createAnilistStateRuntime({
    getClientSecretState: () => appState.anilistClientSecretState,
    setClientSecretState: (next) => {
      appState.anilistClientSecretState = transitionAnilistClientSecretState(
        appState.anilistClientSecretState,
        next,
      );
    },
    getRetryQueueState: () => appState.anilistRetryQueueState,
    setRetryQueueState: (next) => {
      appState.anilistRetryQueueState = transitionAnilistRetryQueueState(
        appState.anilistRetryQueueState,
        next,
      );
    },
    getUpdateQueueSnapshot: () => anilistUpdateQueue.getSnapshot(),
    clearStoredToken: () => anilistTokenStore.clearToken(),
    clearCachedAccessToken: () => {
      anilistCachedAccessToken = null;
    },
  });

  const {
    notifyAnilistSetup,
    consumeAnilistSetupTokenFromUrl,
    handleAnilistSetupProtocolUrl,
    registerSubminerProtocolClient,
  } = composeAnilistSetupHandlers({
    notifyDeps: {
      getNotificationType: () =>
        deps.getOverlayNotifications().getConfiguredStatusNotificationType(),
      hasMpvClient: () => Boolean(appState.mpvClient),
      showMpvOsd: (message) =>
        deps
          .getOverlayNotifications()
          .showConfiguredStatusNotification(message, { title: 'AniList' }),
      showOverlayNotification: (payload) =>
        deps.getOverlayNotifications().showOverlayNotification(payload),
      showDesktopNotification: (title, options) => deps.showDesktopNotification(title, options),
      logInfo: (message) => logger.info(message),
    },
    consumeTokenDeps: {
      consumeAnilistSetupCallbackUrl,
      saveToken: (token) => anilistTokenStore.saveToken(token),
      setCachedToken: (token) => {
        anilistCachedAccessToken = token;
      },
      setResolvedState: (resolvedAt) => {
        anilistStateRuntime.setClientSecretState({
          status: 'resolved',
          source: 'stored',
          message: 'saved token from AniList login',
          resolvedAt,
          errorAt: null,
        });
      },
      setSetupPageOpened: (opened) => {
        appState.anilistSetupPageOpened = opened;
      },
      onSuccess: () => {
        notifyAnilistSetup('AniList login success');
      },
      closeWindow: () => {
        if (appState.anilistSetupWindow && !appState.anilistSetupWindow.isDestroyed()) {
          appState.anilistSetupWindow.close();
        }
      },
    },
    handleProtocolDeps: {
      consumeAnilistSetupTokenFromUrl: (rawUrl) => consumeAnilistSetupTokenFromUrl(rawUrl),
      logWarn: (message, details) => logger.warn(message, details),
    },
    registerProtocolClientDeps: {
      isDefaultApp: () => Boolean(process.defaultApp),
      getArgv: () => process.argv,
      execPath: process.execPath,
      resolvePath: (value) => path.resolve(value),
      setAsDefaultProtocolClient: (scheme, appPath, args) =>
        appPath
          ? deps.app.setAsDefaultProtocolClient(scheme, appPath, args)
          : deps.app.setAsDefaultProtocolClient(scheme),
      logDebug: (message, details) => logger.debug(message, details),
    },
  });

  const maybeFocusExistingAnilistSetupWindow = createMaybeFocusExistingAnilistSetupWindowHandler({
    getSetupWindow: () => appState.anilistSetupWindow,
  });
  const openAnilistSetupWindowHandler = createOpenAnilistSetupWindowHandler({
    maybeFocusExistingSetupWindow: maybeFocusExistingAnilistSetupWindow,
    createSetupWindow: createCreateAnilistSetupWindowHandler({
      createBrowserWindow: (options) => deps.createBrowserWindow(options),
    }),
    buildAuthorizeUrl: () =>
      buildAnilistSetupUrl({
        authorizeUrl: ANILIST_SETUP_CLIENT_ID_URL,
        clientId: ANILIST_DEFAULT_CLIENT_ID,
        responseType: ANILIST_SETUP_RESPONSE_TYPE,
      }),
    consumeCallbackUrl: (rawUrl) => consumeAnilistSetupTokenFromUrl(rawUrl),
    openSetupInBrowser: (authorizeUrl) =>
      openAnilistSetupInBrowser({
        authorizeUrl,
        openExternal: (url) => deps.shell.openExternal(url),
        logError: (message, error) => logger.error(message, error),
      }),
    loadManualTokenEntry: (setupWindow, authorizeUrl) =>
      loadAnilistManualTokenEntry({
        setupWindow: setupWindow as BrowserWindow,
        authorizeUrl,
        developerSettingsUrl: ANILIST_DEVELOPER_SETTINGS_URL,
        logWarn: (message, data) => logger.warn(message, data),
      }),
    redirectUri: ANILIST_REDIRECT_URI,
    developerSettingsUrl: ANILIST_DEVELOPER_SETTINGS_URL,
    isAllowedExternalUrl: (url) => isAllowedAnilistExternalUrl(url),
    isAllowedNavigationUrl: (url) => isAllowedAnilistSetupNavigationUrl(url),
    logWarn: (message, details) => logger.warn(message, details),
    logError: (message, details) => logger.error(message, details),
    clearSetupWindow: () => {
      appState.anilistSetupWindow = null;
    },
    setSetupPageOpened: (opened) => {
      appState.anilistSetupPageOpened = opened;
    },
    setSetupWindow: (setupWindow) => {
      appState.anilistSetupWindow = setupWindow as BrowserWindow;
    },
    openExternal: (url) => {
      void deps.shell.openExternal(url);
    },
  });

  const {
    refreshAnilistClientSecretState,
    getCurrentAnilistMediaKey,
    resetAnilistMediaTracking,
    getAnilistMediaGuessRuntimeState,
    setAnilistMediaGuessRuntimeState,
    recordAnilistMediaDuration,
    resetAnilistMediaGuessState,
    maybeProbeAnilistDuration,
    ensureAnilistMediaGuess,
    processNextAnilistRetryUpdate,
    maybeRunAnilistPostWatchUpdate,
  } = composeAnilistTrackingHandlers({
    refreshClientSecretMainDeps: {
      getResolvedConfig: () => configService.getConfig(),
      isAnilistTrackingEnabled: (config) => isAnilistTrackingEnabled(config as ResolvedConfig),
      getCachedAccessToken: () => anilistCachedAccessToken,
      setCachedAccessToken: (token) => {
        anilistCachedAccessToken = token;
      },
      saveStoredToken: (token) => {
        anilistTokenStore.saveToken(token);
      },
      loadStoredToken: () => anilistTokenStore.loadToken(),
      setClientSecretState: (state) => {
        anilistStateRuntime.setClientSecretState(state);
      },
      getAnilistSetupPageOpened: () => appState.anilistSetupPageOpened,
      setAnilistSetupPageOpened: (opened) => {
        appState.anilistSetupPageOpened = opened;
      },
      openAnilistSetupWindow: () => {
        openAnilistSetupWindowHandler();
      },
      now: () => Date.now(),
    },
    getCurrentMediaKeyMainDeps: {
      getCurrentMediaPath: () => appState.currentMediaPath,
    },
    resetMediaTrackingMainDeps: {
      setMediaKey: (value) => {
        anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
          anilistMediaGuessRuntimeState,
          { mediaKey: value },
        );
      },
      setMediaDurationSec: (value) => {
        anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
          anilistMediaGuessRuntimeState,
          { mediaDurationSec: value },
        );
      },
      setMediaGuess: (value) => {
        anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
          anilistMediaGuessRuntimeState,
          { mediaGuess: value },
        );
      },
      setMediaGuessPromise: (value) => {
        anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
          anilistMediaGuessRuntimeState,
          { mediaGuessPromise: value },
        );
      },
      setLastDurationProbeAtMs: (value) => {
        anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
          anilistMediaGuessRuntimeState,
          { lastDurationProbeAtMs: value },
        );
      },
    },
    getMediaGuessRuntimeStateMainDeps: {
      getMediaKey: () => anilistMediaGuessRuntimeState.mediaKey,
      getMediaDurationSec: () => anilistMediaGuessRuntimeState.mediaDurationSec,
      getMediaGuess: () => anilistMediaGuessRuntimeState.mediaGuess,
      getMediaGuessPromise: () => anilistMediaGuessRuntimeState.mediaGuessPromise,
      getLastDurationProbeAtMs: () => anilistMediaGuessRuntimeState.lastDurationProbeAtMs,
    },
    setMediaGuessRuntimeStateMainDeps: {
      setMediaKey: (value) => {
        anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
          anilistMediaGuessRuntimeState,
          { mediaKey: value },
        );
      },
      setMediaDurationSec: (value) => {
        anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
          anilistMediaGuessRuntimeState,
          { mediaDurationSec: value },
        );
      },
      setMediaGuess: (value) => {
        anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
          anilistMediaGuessRuntimeState,
          { mediaGuess: value },
        );
      },
      setMediaGuessPromise: (value) => {
        anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
          anilistMediaGuessRuntimeState,
          { mediaGuessPromise: value },
        );
      },
      setLastDurationProbeAtMs: (value) => {
        anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
          anilistMediaGuessRuntimeState,
          { lastDurationProbeAtMs: value },
        );
      },
    },
    recordMediaDurationMainDeps: {
      getCurrentMediaKey: () => getCurrentAnilistMediaKey(),
      getState: () => getAnilistMediaGuessRuntimeState(),
      setState: (state) => {
        setAnilistMediaGuessRuntimeState(state);
      },
    },
    resetMediaGuessStateMainDeps: {
      setMediaGuess: (value) => {
        anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
          anilistMediaGuessRuntimeState,
          { mediaGuess: value },
        );
      },
      setMediaGuessPromise: (value) => {
        anilistMediaGuessRuntimeState = transitionAnilistMediaGuessRuntimeState(
          anilistMediaGuessRuntimeState,
          { mediaGuessPromise: value },
        );
      },
    },
    maybeProbeDurationMainDeps: {
      getState: () => getAnilistMediaGuessRuntimeState(),
      setState: (state) => {
        setAnilistMediaGuessRuntimeState(state);
      },
      durationRetryIntervalMs: ANILIST_DURATION_RETRY_INTERVAL_MS,
      now: () => Date.now(),
      requestMpvDuration: async () => appState.mpvClient?.requestProperty('duration'),
      logWarn: (message, error) => logger.warn(message, error),
    },
    ensureMediaGuessMainDeps: {
      getState: () => getAnilistMediaGuessRuntimeState(),
      setState: (state) => {
        setAnilistMediaGuessRuntimeState(state);
      },
      resolveMediaPathForJimaku: (currentMediaPath) =>
        deps.resolveMediaPathForJimaku(currentMediaPath),
      getCurrentMediaPath: () => appState.currentMediaPath,
      getCurrentMediaTitle: () => appState.currentMediaTitle,
      guessAnilistMediaInfo: (mediaPath, mediaTitle) =>
        guessAnilistMediaInfo(mediaPath, mediaTitle),
    },
    processNextRetryUpdateMainDeps: {
      nextReady: () => anilistUpdateQueue.nextReady(),
      refreshRetryQueueState: () => anilistStateRuntime.refreshRetryQueueState(),
      setLastAttemptAt: (value) => {
        appState.anilistRetryQueueState = transitionAnilistRetryQueueLastAttemptAt(
          appState.anilistRetryQueueState,
          value,
        );
      },
      setLastError: (value) => {
        appState.anilistRetryQueueState = transitionAnilistRetryQueueLastError(
          appState.anilistRetryQueueState,
          value,
        );
      },
      refreshAnilistClientSecretState: () => refreshAnilistClientSecretState(),
      updateAnilistPostWatchProgress: (accessToken, title, episode, season, mediaId) =>
        updateAnilistPostWatchProgress(accessToken, title, episode, {
          rateLimiter: anilistRateLimiter,
          season,
          mediaId,
          logInfo: (message) => logger.info(message),
        }),
      markSuccess: (key) => {
        anilistUpdateQueue.markSuccess(key);
      },
      rememberAttemptedUpdateKey: (key) => {
        rememberAnilistAttemptedUpdate(key);
      },
      markFailure: (key, message) => {
        anilistUpdateQueue.markFailure(key, message);
      },
      logInfo: (message) => logger.info(message),
      now: () => Date.now(),
    },
    maybeRunPostWatchUpdateMainDeps: {
      getInFlight: () => anilistUpdateInFlightState.inFlight,
      setInFlight: (value) => {
        anilistUpdateInFlightState = transitionAnilistUpdateInFlightState(
          anilistUpdateInFlightState,
          value,
        );
      },
      getResolvedConfig: () => configService.getConfig(),
      isAnilistTrackingEnabled: (config) => isAnilistTrackingEnabled(config as ResolvedConfig),
      getCurrentMediaKey: () => getCurrentAnilistMediaKey(),
      hasMpvClient: () => Boolean(appState.mpvClient),
      getTrackedMediaKey: () => anilistMediaGuessRuntimeState.mediaKey,
      resetTrackedMedia: (mediaKey) => {
        resetAnilistMediaTracking(mediaKey);
      },
      getWatchedSeconds: () => appState.mpvClient?.currentTimePos ?? Number.NaN,
      maybeProbeAnilistDuration: (mediaKey, options) =>
        maybeProbeAnilistDuration(mediaKey, options),
      ensureAnilistMediaGuess: (mediaKey) => ensureAnilistMediaGuess(mediaKey),
      hasAttemptedUpdateKey: (key) => anilistAttemptedUpdateKeys.has(key),
      processNextAnilistRetryUpdate: () => processNextAnilistRetryUpdate(),
      refreshAnilistClientSecretState: () => refreshAnilistClientSecretState(),
      enqueueRetry: (key, title, episode, season, mediaId) => {
        anilistUpdateQueue.enqueue(key, title, episode, season, mediaId);
      },
      getCurrentMediaTitle: () => appState.currentMediaTitle,
      resolvePinnedAnilistMediaId: ({ mediaPath, mediaTitle, guess }) =>
        deps.resolvePinnedMediaId({ mediaPath, mediaTitle, guess }),
      markRetryFailure: (key, message) => {
        anilistUpdateQueue.markFailure(key, message);
      },
      markRetrySuccess: (key) => {
        anilistUpdateQueue.markSuccess(key);
      },
      refreshRetryQueueState: () => anilistStateRuntime.refreshRetryQueueState(),
      updateAnilistPostWatchProgress: (accessToken, title, episode, season, mediaId) =>
        updateAnilistPostWatchProgress(accessToken, title, episode, {
          rateLimiter: anilistRateLimiter,
          season,
          mediaId,
          logInfo: (message) => logger.info(message),
        }),
      rememberAttemptedUpdateKey: (key) => {
        rememberAnilistAttemptedUpdate(key);
      },
      showMpvOsd: (message) =>
        deps
          .getOverlayNotifications()
          .showConfiguredStatusNotification(message, { title: 'AniList' }),
      logInfo: (message) => logger.info(message),
      logWarn: (message) => logger.warn(message),
      minWatchSeconds: ANILIST_UPDATE_MIN_WATCH_SECONDS,
      minWatchRatio: DEFAULT_MIN_WATCH_RATIO,
    },
  });

  function refreshAnilistClientSecretStateIfEnabled(options?: {
    force?: boolean;
    allowSetupPrompt?: boolean;
  }): Promise<string | null> {
    if (!isAnilistTrackingEnabled(configService.getConfig())) {
      return Promise.resolve(null);
    }
    return refreshAnilistClientSecretState(options);
  }

  const rememberAnilistAttemptedUpdate = (key: string): void => {
    rememberAnilistAttemptedUpdateKey(
      anilistAttemptedUpdateKeys,
      key,
      ANILIST_MAX_ATTEMPTED_UPDATE_KEYS,
    );
  };

  const ipc: IpcSlice<
    | 'getAnilistStatus'
    | 'clearAnilistToken'
    | 'openAnilistSetup'
    | 'getAnilistQueueStatus'
    | 'retryAnilistQueueNow'
    | 'runAnilistPostWatchUpdateOnManualMark'
  > = {
    getAnilistStatus: () => anilistStateRuntime.getStatusSnapshot(),
    clearAnilistToken: () => anilistStateRuntime.clearTokenState(),
    openAnilistSetup: () => openAnilistSetupWindowHandler(),
    getAnilistQueueStatus: () => anilistStateRuntime.getQueueStatusSnapshot(),
    retryAnilistQueueNow: () => processNextAnilistRetryUpdate(),
    runAnilistPostWatchUpdateOnManualMark: () => maybeRunAnilistPostWatchUpdate({ force: true }),
  };

  const cli: CliSlice<
    | 'getAnilistStatus'
    | 'clearAnilistToken'
    | 'openAnilistSetupWindow'
    | 'getAnilistQueueStatus'
    | 'processNextAnilistRetryUpdate'
  > = {
    getAnilistStatus: () => anilistStateRuntime.getStatusSnapshot(),
    clearAnilistToken: () => anilistStateRuntime.clearTokenState(),
    openAnilistSetupWindow: () => openAnilistSetupWindowHandler(),
    getAnilistQueueStatus: () => anilistStateRuntime.getQueueStatusSnapshot(),
    processNextAnilistRetryUpdate: () => processNextAnilistRetryUpdate(),
  };

  return {
    ipc,
    cli,
    stateRuntime: anilistStateRuntime,
    openSetupWindow: openAnilistSetupWindowHandler,
    handleAnilistSetupProtocolUrl,
    registerSubminerProtocolClient,
    refreshClientSecretStateIfEnabled: refreshAnilistClientSecretStateIfEnabled,
    getCurrentAnilistMediaKey,
    resetAnilistMediaTracking,
    recordAnilistMediaDuration,
    resetAnilistMediaGuessState,
    maybeProbeAnilistDuration,
    ensureAnilistMediaGuess,
    processNextAnilistRetryUpdate,
    maybeRunAnilistPostWatchUpdate,
    /** Duration of the media AniList is tracking, used as Discord presence's fallback. */
    getMediaDurationSec: () => anilistMediaGuessRuntimeState.mediaDurationSec,
  };
}
