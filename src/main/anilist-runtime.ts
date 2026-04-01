import {
  createInitialAnilistMediaGuessRuntimeState,
  createInitialAnilistRetryQueueState,
  createInitialAnilistSecretResolutionState,
  createInitialAnilistUpdateInFlightState,
  type AnilistMediaGuessRuntimeState,
  type AnilistRetryQueueState,
  type AnilistSecretResolutionState,
} from './state';
import { createAnilistStateRuntime } from './runtime/anilist-state';
import { composeAnilistSetupHandlers } from './runtime/composers/anilist-setup-composer';
import { composeAnilistTrackingHandlers } from './runtime/composers/anilist-tracking-composer';
import {
  buildAnilistSetupUrl,
  consumeAnilistSetupCallbackUrl,
  loadAnilistManualTokenEntry,
  openAnilistSetupInBrowser,
} from './runtime/anilist-setup';
import {
  createMaybeFocusExistingAnilistSetupWindowHandler,
  createOpenAnilistSetupWindowHandler,
} from './runtime/anilist-setup-window';
import {
  buildAnilistAttemptKey,
  rememberAnilistAttemptedUpdateKey,
} from './runtime/anilist-post-watch';
import { createCreateAnilistSetupWindowHandler } from './runtime/setup-window-factory';
import type {
  AnilistMediaGuess,
  AnilistPostWatchUpdateResult,
} from '../core/services/anilist/anilist-updater';
import type { AnilistUpdateQueue } from '../core/services/anilist/anilist-update-queue';

export interface AnilistSetupWindowLike {
  focus: () => void;
  close: () => void;
  isDestroyed: () => boolean;
  on: (event: 'closed', handler: () => void) => void;
  loadURL: (url: string) => Promise<void> | void;
  webContents: {
    setWindowOpenHandler: (handler: (details: { url: string }) => { action: 'deny' }) => void;
    on: (event: string, handler: (...args: unknown[]) => void) => void;
    getURL: () => string;
  };
}

export interface AnilistTokenStoreLike {
  saveToken: (token: string) => void;
  loadToken: () => string | null | undefined;
  clearToken: () => void;
}

export interface AnilistRuntimeInput<
  TConfig extends { anilist: { accessToken: string; enabled?: boolean } } = {
    anilist: { accessToken: string; enabled?: boolean };
  },
  TWindow extends AnilistSetupWindowLike = AnilistSetupWindowLike,
> {
  getResolvedConfig: () => TConfig;
  isTrackingEnabled: (config: TConfig) => boolean;
  tokenStore: AnilistTokenStoreLike;
  updateQueue: AnilistUpdateQueue;
  getCurrentMediaPath: () => string | null;
  getCurrentMediaTitle: () => string | null;
  getWatchedSeconds: () => number;
  hasMpvClient: () => boolean;
  requestMpvDuration: () => Promise<unknown>;
  resolveMediaPathForJimaku: (currentMediaPath: string | null) => string | null;
  guessAnilistMediaInfo: (
    mediaPath: string | null,
    mediaTitle: string | null,
  ) => Promise<AnilistMediaGuess | null>;
  updateAnilistPostWatchProgress: (
    accessToken: string,
    title: string,
    episode: number,
  ) => Promise<AnilistPostWatchUpdateResult>;
  createBrowserWindow: (options: Electron.BrowserWindowConstructorOptions) => TWindow;
  authorizeUrl: string;
  clientId: string;
  responseType: string;
  redirectUri: string;
  developerSettingsUrl: string;
  isAllowedExternalUrl: (url: string) => boolean;
  isAllowedNavigationUrl: (url: string) => boolean;
  openExternal: (url: string) => Promise<unknown> | void;
  showMpvOsd: (message: string) => void;
  showDesktopNotification: (title: string, options: { body: string }) => void;
  logInfo: (message: string) => void;
  logWarn: (message: string, details?: unknown) => void;
  logError: (message: string, error: unknown) => void;
  logDebug: (message: string, details?: unknown) => void;
  isDefaultApp: () => boolean;
  getArgv: () => string[];
  execPath: string;
  resolvePath: (value: string) => string;
  setAsDefaultProtocolClient: (scheme: string, path?: string, args?: string[]) => boolean;
  now?: () => number;
  durationRetryIntervalMs?: number;
  minWatchSeconds?: number;
  minWatchRatio?: number;
  maxAttemptedUpdateKeys?: number;
}

export interface AnilistRuntime {
  notifyAnilistSetup: (message: string) => void;
  consumeAnilistSetupTokenFromUrl: (rawUrl: string) => boolean;
  handleAnilistSetupProtocolUrl: (rawUrl: string) => boolean;
  registerSubminerProtocolClient: () => void;
  openAnilistSetupWindow: () => void;
  refreshAnilistClientSecretState: (options?: {
    force?: boolean;
    allowSetupPrompt?: boolean;
  }) => Promise<string | null>;
  refreshAnilistClientSecretStateIfEnabled: (options?: {
    force?: boolean;
    allowSetupPrompt?: boolean;
  }) => Promise<string | null>;
  getCurrentAnilistMediaKey: () => string | null;
  resetAnilistMediaTracking: (mediaKey: string | null) => void;
  getAnilistMediaGuessRuntimeState: () => AnilistMediaGuessRuntimeState;
  setAnilistMediaGuessRuntimeState: (state: AnilistMediaGuessRuntimeState) => void;
  resetAnilistMediaGuessState: () => void;
  maybeProbeAnilistDuration: (mediaKey: string) => Promise<number | null>;
  ensureAnilistMediaGuess: (mediaKey: string) => Promise<AnilistMediaGuess | null>;
  processNextAnilistRetryUpdate: () => Promise<{ ok: boolean; message: string }>;
  maybeRunAnilistPostWatchUpdate: () => Promise<void>;
  setClientSecretState: (partial: Partial<AnilistSecretResolutionState>) => void;
  refreshRetryQueueState: () => void;
  getStatusSnapshot: () => {
    tokenStatus: AnilistSecretResolutionState['status'];
    tokenSource: AnilistSecretResolutionState['source'];
    tokenMessage: string | null;
    tokenResolvedAt: number | null;
    tokenErrorAt: number | null;
    queuePending: number;
    queueReady: number;
    queueDeadLetter: number;
    queueLastAttemptAt: number | null;
    queueLastError: string | null;
  };
  getQueueStatusSnapshot: () => AnilistRetryQueueState;
  clearTokenState: () => void;
  getSetupWindow: () => AnilistSetupWindowLike | null;
}

const DEFAULT_DURATION_RETRY_INTERVAL_MS = 15_000;
const DEFAULT_MIN_WATCH_SECONDS = 10 * 60;
const DEFAULT_MIN_WATCH_RATIO = 0.85;
const DEFAULT_MAX_ATTEMPTED_UPDATE_KEYS = 1000;

export function createAnilistRuntime<
  TConfig extends { anilist: { accessToken: string; enabled?: boolean } },
  TWindow extends AnilistSetupWindowLike,
>(input: AnilistRuntimeInput<TConfig, TWindow>): AnilistRuntime {
  const now = input.now ?? Date.now;

  let setupWindow: TWindow | null = null;
  let setupPageOpened = false;
  let cachedAccessToken: string | null = null;
  let clientSecretState = createInitialAnilistSecretResolutionState();
  let retryQueueState = createInitialAnilistRetryQueueState();
  let mediaGuessRuntimeState = createInitialAnilistMediaGuessRuntimeState();
  let updateInFlightState = createInitialAnilistUpdateInFlightState();
  const attemptedUpdateKeys = new Set<string>();

  const stateRuntime = createAnilistStateRuntime({
    getClientSecretState: () => clientSecretState,
    setClientSecretState: (next) => {
      clientSecretState = next;
    },
    getRetryQueueState: () => retryQueueState,
    setRetryQueueState: (next) => {
      retryQueueState = next;
    },
    getUpdateQueueSnapshot: () => input.updateQueue.getSnapshot(),
    clearStoredToken: () => input.tokenStore.clearToken(),
    clearCachedAccessToken: () => {
      cachedAccessToken = null;
    },
  });

  const rememberAttemptedUpdate = (key: string): void => {
    rememberAnilistAttemptedUpdateKey(
      attemptedUpdateKeys,
      key,
      input.maxAttemptedUpdateKeys ?? DEFAULT_MAX_ATTEMPTED_UPDATE_KEYS,
    );
  };

  const maybeFocusExistingSetupWindow = createMaybeFocusExistingAnilistSetupWindowHandler({
    getSetupWindow: () => setupWindow,
  });
  const createSetupWindow = createCreateAnilistSetupWindowHandler({
    createBrowserWindow: (options) => input.createBrowserWindow(options),
  });

  const {
    notifyAnilistSetup,
    consumeAnilistSetupTokenFromUrl,
    handleAnilistSetupProtocolUrl,
    registerSubminerProtocolClient,
  } = composeAnilistSetupHandlers({
    notifyDeps: {
      hasMpvClient: () => input.hasMpvClient(),
      showMpvOsd: (message) => input.showMpvOsd(message),
      showDesktopNotification: (title, options) => input.showDesktopNotification(title, options),
      logInfo: (message) => input.logInfo(message),
    },
    consumeTokenDeps: {
      consumeAnilistSetupCallbackUrl,
      saveToken: (token) => input.tokenStore.saveToken(token),
      setCachedToken: (token) => {
        cachedAccessToken = token;
      },
      setResolvedState: (resolvedAt) => {
        stateRuntime.setClientSecretState({
          status: 'resolved',
          source: 'stored',
          message: 'saved token from AniList login',
          resolvedAt,
          errorAt: null,
        });
      },
      setSetupPageOpened: (opened) => {
        setupPageOpened = opened;
      },
      onSuccess: () => {
        notifyAnilistSetup('AniList login success');
      },
      closeWindow: () => {
        if (setupWindow && !setupWindow.isDestroyed()) {
          setupWindow.close();
        }
      },
    },
    handleProtocolDeps: {
      consumeAnilistSetupTokenFromUrl: (rawUrl) => consumeAnilistSetupTokenFromUrl(rawUrl),
      logWarn: (message, details) => input.logWarn(message, details),
    },
    registerProtocolClientDeps: {
      isDefaultApp: () => input.isDefaultApp(),
      getArgv: () => input.getArgv(),
      execPath: input.execPath,
      resolvePath: (value) => input.resolvePath(value),
      setAsDefaultProtocolClient: (scheme, targetPath, args) =>
        input.setAsDefaultProtocolClient(scheme, targetPath, args),
      logDebug: (message, details) => input.logDebug(message, details),
    },
  });

  const openAnilistSetupWindow = createOpenAnilistSetupWindowHandler({
    maybeFocusExistingSetupWindow: () => maybeFocusExistingSetupWindow(),
    createSetupWindow: () => createSetupWindow(),
    buildAuthorizeUrl: () =>
      buildAnilistSetupUrl({
        authorizeUrl: input.authorizeUrl,
        clientId: input.clientId,
        responseType: input.responseType,
        redirectUri: input.redirectUri,
      }),
    consumeCallbackUrl: (rawUrl) => consumeAnilistSetupTokenFromUrl(rawUrl),
    openSetupInBrowser: (authorizeUrl) =>
      openAnilistSetupInBrowser({
        authorizeUrl,
        openExternal: async (url) => {
          await input.openExternal(url);
        },
        logError: (message, error) => input.logError(message, error),
      }),
    loadManualTokenEntry: (window, authorizeUrl) =>
      loadAnilistManualTokenEntry({
        setupWindow: window as never,
        authorizeUrl,
        developerSettingsUrl: input.developerSettingsUrl,
        logWarn: (message, details) => input.logWarn(message, details),
      }),
    redirectUri: input.redirectUri,
    developerSettingsUrl: input.developerSettingsUrl,
    isAllowedExternalUrl: (url) => input.isAllowedExternalUrl(url),
    isAllowedNavigationUrl: (url) => input.isAllowedNavigationUrl(url),
    logWarn: (message, details) => input.logWarn(message, details),
    logError: (message, details) => input.logError(message, details),
    clearSetupWindow: () => {
      setupWindow = null;
    },
    setSetupPageOpened: (opened) => {
      setupPageOpened = opened;
    },
    setSetupWindow: (window) => {
      setupWindow = window;
    },
    openExternal: (url) => {
      void input.openExternal(url);
    },
  });

  const trackingRuntime = composeAnilistTrackingHandlers({
    refreshClientSecretMainDeps: {
      getResolvedConfig: () => input.getResolvedConfig(),
      isAnilistTrackingEnabled: (config) => input.isTrackingEnabled(config as TConfig),
      getCachedAccessToken: () => cachedAccessToken,
      setCachedAccessToken: (token) => {
        cachedAccessToken = token;
      },
      saveStoredToken: (token) => {
        input.tokenStore.saveToken(token);
      },
      loadStoredToken: () => input.tokenStore.loadToken(),
      setClientSecretState: (state) => {
        clientSecretState = state;
      },
      getAnilistSetupPageOpened: () => setupPageOpened,
      setAnilistSetupPageOpened: (opened) => {
        setupPageOpened = opened;
      },
      openAnilistSetupWindow: () => {
        openAnilistSetupWindow();
      },
      now,
    },
    getCurrentMediaKeyMainDeps: {
      getCurrentMediaPath: () => input.getCurrentMediaPath(),
    },
    resetMediaTrackingMainDeps: {
      setMediaKey: (value) => {
        mediaGuessRuntimeState = { ...mediaGuessRuntimeState, mediaKey: value };
      },
      setMediaDurationSec: (value) => {
        mediaGuessRuntimeState = { ...mediaGuessRuntimeState, mediaDurationSec: value };
      },
      setMediaGuess: (value) => {
        mediaGuessRuntimeState = { ...mediaGuessRuntimeState, mediaGuess: value };
      },
      setMediaGuessPromise: (value) => {
        mediaGuessRuntimeState = { ...mediaGuessRuntimeState, mediaGuessPromise: value };
      },
      setLastDurationProbeAtMs: (value) => {
        mediaGuessRuntimeState = { ...mediaGuessRuntimeState, lastDurationProbeAtMs: value };
      },
    },
    getMediaGuessRuntimeStateMainDeps: {
      getMediaKey: () => mediaGuessRuntimeState.mediaKey,
      getMediaDurationSec: () => mediaGuessRuntimeState.mediaDurationSec,
      getMediaGuess: () => mediaGuessRuntimeState.mediaGuess,
      getMediaGuessPromise: () => mediaGuessRuntimeState.mediaGuessPromise,
      getLastDurationProbeAtMs: () => mediaGuessRuntimeState.lastDurationProbeAtMs,
    },
    setMediaGuessRuntimeStateMainDeps: {
      setMediaKey: (value) => {
        mediaGuessRuntimeState = { ...mediaGuessRuntimeState, mediaKey: value };
      },
      setMediaDurationSec: (value) => {
        mediaGuessRuntimeState = { ...mediaGuessRuntimeState, mediaDurationSec: value };
      },
      setMediaGuess: (value) => {
        mediaGuessRuntimeState = { ...mediaGuessRuntimeState, mediaGuess: value };
      },
      setMediaGuessPromise: (value) => {
        mediaGuessRuntimeState = { ...mediaGuessRuntimeState, mediaGuessPromise: value };
      },
      setLastDurationProbeAtMs: (value) => {
        mediaGuessRuntimeState = { ...mediaGuessRuntimeState, lastDurationProbeAtMs: value };
      },
    },
    resetMediaGuessStateMainDeps: {
      setMediaGuess: (value) => {
        mediaGuessRuntimeState = { ...mediaGuessRuntimeState, mediaGuess: value };
      },
      setMediaGuessPromise: (value) => {
        mediaGuessRuntimeState = { ...mediaGuessRuntimeState, mediaGuessPromise: value };
      },
    },
    maybeProbeDurationMainDeps: {
      getState: () => mediaGuessRuntimeState,
      setState: (state) => {
        mediaGuessRuntimeState = state;
      },
      durationRetryIntervalMs: input.durationRetryIntervalMs ?? DEFAULT_DURATION_RETRY_INTERVAL_MS,
      now,
      requestMpvDuration: () => input.requestMpvDuration(),
      logWarn: (message, error) => input.logWarn(message, error),
    },
    ensureMediaGuessMainDeps: {
      getState: () => mediaGuessRuntimeState,
      setState: (state) => {
        mediaGuessRuntimeState = state;
      },
      resolveMediaPathForJimaku: (currentMediaPath) =>
        input.resolveMediaPathForJimaku(currentMediaPath),
      getCurrentMediaPath: () => input.getCurrentMediaPath(),
      getCurrentMediaTitle: () => input.getCurrentMediaTitle(),
      guessAnilistMediaInfo: (mediaPath, mediaTitle) =>
        input.guessAnilistMediaInfo(mediaPath, mediaTitle),
    },
    processNextRetryUpdateMainDeps: {
      nextReady: () => input.updateQueue.nextReady(),
      refreshRetryQueueState: () => stateRuntime.refreshRetryQueueState(),
      setLastAttemptAt: (value) => {
        retryQueueState = { ...retryQueueState, lastAttemptAt: value };
      },
      setLastError: (value) => {
        retryQueueState = { ...retryQueueState, lastError: value };
      },
      refreshAnilistClientSecretState: () => trackingRuntime.refreshAnilistClientSecretState(),
      updateAnilistPostWatchProgress: (accessToken, title, episode) =>
        input.updateAnilistPostWatchProgress(accessToken, title, episode),
      markSuccess: (key) => {
        input.updateQueue.markSuccess(key);
      },
      rememberAttemptedUpdateKey: (key) => {
        rememberAttemptedUpdate(key);
      },
      markFailure: (key, message) => {
        input.updateQueue.markFailure(key, message);
      },
      logInfo: (message) => input.logInfo(message),
      now,
    },
    maybeRunPostWatchUpdateMainDeps: {
      getInFlight: () => updateInFlightState.inFlight,
      setInFlight: (value) => {
        updateInFlightState = { ...updateInFlightState, inFlight: value };
      },
      getResolvedConfig: () => input.getResolvedConfig(),
      isAnilistTrackingEnabled: (config) => input.isTrackingEnabled(config as TConfig),
      getCurrentMediaKey: () => trackingRuntime.getCurrentAnilistMediaKey(),
      hasMpvClient: () => input.hasMpvClient(),
      getTrackedMediaKey: () => mediaGuessRuntimeState.mediaKey,
      resetTrackedMedia: (mediaKey) => {
        trackingRuntime.resetAnilistMediaTracking(mediaKey);
      },
      getWatchedSeconds: () => input.getWatchedSeconds(),
      maybeProbeAnilistDuration: (mediaKey) => trackingRuntime.maybeProbeAnilistDuration(mediaKey),
      ensureAnilistMediaGuess: (mediaKey) => trackingRuntime.ensureAnilistMediaGuess(mediaKey),
      hasAttemptedUpdateKey: (key) => attemptedUpdateKeys.has(key),
      processNextAnilistRetryUpdate: () => trackingRuntime.processNextAnilistRetryUpdate(),
      refreshAnilistClientSecretState: () => trackingRuntime.refreshAnilistClientSecretState(),
      enqueueRetry: (key, title, episode) => {
        input.updateQueue.enqueue(key, title, episode);
      },
      markRetryFailure: (key, message) => {
        input.updateQueue.markFailure(key, message);
      },
      markRetrySuccess: (key) => {
        input.updateQueue.markSuccess(key);
      },
      refreshRetryQueueState: () => stateRuntime.refreshRetryQueueState(),
      updateAnilistPostWatchProgress: (accessToken, title, episode) =>
        input.updateAnilistPostWatchProgress(accessToken, title, episode),
      rememberAttemptedUpdateKey: (key) => {
        rememberAttemptedUpdate(key);
      },
      showMpvOsd: (message) => input.showMpvOsd(message),
      logInfo: (message) => input.logInfo(message),
      logWarn: (message) => input.logWarn(message),
      minWatchSeconds: input.minWatchSeconds ?? DEFAULT_MIN_WATCH_SECONDS,
      minWatchRatio: input.minWatchRatio ?? DEFAULT_MIN_WATCH_RATIO,
    },
  });

  return {
    notifyAnilistSetup,
    consumeAnilistSetupTokenFromUrl,
    handleAnilistSetupProtocolUrl,
    registerSubminerProtocolClient,
    openAnilistSetupWindow,
    refreshAnilistClientSecretState: (options) =>
      trackingRuntime.refreshAnilistClientSecretState(options),
    refreshAnilistClientSecretStateIfEnabled: (options) => {
      if (!input.isTrackingEnabled(input.getResolvedConfig())) {
        return Promise.resolve(null);
      }
      return trackingRuntime.refreshAnilistClientSecretState(options);
    },
    getCurrentAnilistMediaKey: () => trackingRuntime.getCurrentAnilistMediaKey(),
    resetAnilistMediaTracking: (mediaKey) => trackingRuntime.resetAnilistMediaTracking(mediaKey),
    getAnilistMediaGuessRuntimeState: () => trackingRuntime.getAnilistMediaGuessRuntimeState(),
    setAnilistMediaGuessRuntimeState: (state) =>
      trackingRuntime.setAnilistMediaGuessRuntimeState(state),
    resetAnilistMediaGuessState: () => trackingRuntime.resetAnilistMediaGuessState(),
    maybeProbeAnilistDuration: (mediaKey) => trackingRuntime.maybeProbeAnilistDuration(mediaKey),
    ensureAnilistMediaGuess: (mediaKey) => trackingRuntime.ensureAnilistMediaGuess(mediaKey),
    processNextAnilistRetryUpdate: () => trackingRuntime.processNextAnilistRetryUpdate(),
    maybeRunAnilistPostWatchUpdate: () => trackingRuntime.maybeRunAnilistPostWatchUpdate(),
    setClientSecretState: (partial) => stateRuntime.setClientSecretState(partial),
    refreshRetryQueueState: () => stateRuntime.refreshRetryQueueState(),
    getStatusSnapshot: () => stateRuntime.getStatusSnapshot(),
    getQueueStatusSnapshot: () => stateRuntime.getQueueStatusSnapshot(),
    clearTokenState: () => stateRuntime.clearTokenState(),
    getSetupWindow: () => setupWindow,
  };
}

export { buildAnilistAttemptKey };
