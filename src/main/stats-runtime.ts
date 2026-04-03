import path from 'node:path';
import {
  isBackgroundStatsServerProcessAlive,
  readBackgroundStatsServerState,
  removeBackgroundStatsServerState,
  resolveBackgroundStatsServerUrl,
  writeBackgroundStatsServerState,
  type BackgroundStatsServerState,
} from './runtime/stats-daemon';
import {
  createRunStatsCliCommandHandler,
  writeStatsCliCommandResponse,
} from './runtime/stats-cli-command';
import type { CliArgs, CliCommandSource } from '../cli/args';
import { ImmersionTrackerService } from '../core/services/immersion-tracker-service';
import { startStatsServer as startStatsServerCore } from '../core/services/stats-server';
import { createLogger } from '../logger';
import { createCoverArtFetcher } from '../core/services/anilist/cover-art-fetcher';
import { createAnilistRateLimiter } from '../core/services/anilist/rate-limiter';
import { resolveLegacyVocabularyPosFromTokens } from '../core/services/immersion-tracker/legacy-vocabulary-pos';
import type {
  LifetimeRebuildSummary,
  VocabularyCleanupSummary,
} from '../core/services/immersion-tracker/types';
import type { ResolvedConfig } from '../types';
import type { AppReadyImmersionInput } from './app-ready-runtime';
import { createImmersionTrackerStartupHandler } from './runtime/immersion-startup';

type StatsConfigLike = {
  immersionTracking?: {
    enabled?: boolean;
  };
  stats: {
    serverPort: number;
    autoOpenBrowser?: boolean;
  };
};

type StatsServerLike = {
  close: () => void;
};

type StatsTrackerLike = {
  cleanupVocabularyStats?: () => Promise<VocabularyCleanupSummary>;
  rebuildLifetimeSummaries?: () => Promise<LifetimeRebuildSummary>;
  recordCardsMined?: (count: number, noteIds?: number[]) => void;
};

type StatsBootstrapAppState = {
  mecabTokenizer: {
    tokenize: (text: string) => Promise<unknown[] | null>;
  } | null;
  immersionTracker: ImmersionTrackerService | null;
  mpvClient: unknown | null;
  mpvSocketPath: string;
  ankiIntegration: {
    resolveCurrentNoteId: (noteId: number) => number;
  } | null;
  statsOverlayVisible: boolean;
};

export interface StatsRuntimeInput<
  TConfig extends StatsConfigLike = StatsConfigLike,
  TTracker extends StatsTrackerLike = StatsTrackerLike,
  TServer extends StatsServerLike = StatsServerLike,
> {
  statsDaemonStatePath: string;
  getResolvedConfig: () => TConfig;
  getImmersionTracker: () => TTracker | null;
  ensureImmersionTrackerStartedCore: () => void;
  ensureVocabularyCleanupTokenizerReady?: () => Promise<void> | void;
  startStatsServer: (port: number) => TServer;
  openExternal: (url: string) => Promise<unknown>;
  exitAppWithCode: (code: number) => void;
  destroyStatsWindow?: () => void;
  logInfo: (message: string) => void;
  logWarn: (message: string, error?: unknown) => void;
  logError: (message: string, error: unknown) => void;
  now?: () => number;
  getCurrentPid?: () => number;
  isProcessAlive?: (pid: number) => boolean;
  killProcess?: (pid: number, signal: NodeJS.Signals) => void;
  wait?: (delayMs: number) => Promise<void>;
}

export interface StatsRuntime {
  readLiveBackgroundStatsDaemonState: () => BackgroundStatsServerState | null;
  ensureImmersionTrackerStarted: () => void;
  ensureStatsServerStarted: () => string;
  stopStatsServer: () => void;
  ensureBackgroundStatsServerStarted: () => {
    url: string;
    runningInCurrentProcess: boolean;
  };
  stopBackgroundStatsServer: () => Promise<{ ok: boolean; stale: boolean }>;
  runStatsCliCommand: (
    args: Pick<
      CliArgs,
      | 'statsResponsePath'
      | 'statsBackground'
      | 'statsStop'
      | 'statsCleanup'
      | 'statsCleanupVocab'
      | 'statsCleanupLifetime'
    >,
    source: CliCommandSource,
  ) => Promise<void>;
  cleanupBeforeQuit: () => void;
  getStatsServer: () => StatsServerLike | null;
  isStatsStartupInProgress: () => boolean;
}

export interface StatsRuntimeBootstrapInput {
  statsDaemonStatePath: string;
  statsDistPath: string;
  statsPreloadPath: string;
  userDataPath: string;
  appState: StatsBootstrapAppState;
  getResolvedConfig: () => ResolvedConfig;
  dictionarySupport: {
    getConfiguredDbPath: () => string;
    seedImmersionMediaFromCurrentMedia: () => Promise<void> | void;
  };
  overlay: {
    getOverlayGeometry: () => { getCurrentOverlayGeometry: () => Electron.Rectangle };
    updateVisibleOverlayVisibility: () => void;
    registerStatsOverlayToggle: (options: {
      staticDir: string;
      preloadPath: string;
      getApiBaseUrl: () => string;
      getToggleKey: () => string;
      resolveBounds: () => Electron.Rectangle;
      onVisibilityChanged: (visible: boolean) => void;
    }) => void;
  };
  createMecabTokenizerAndCheck: () => Promise<void>;
  addYomitanNote: (word: string) => Promise<number | null>;
  openExternal: (url: string) => Promise<unknown>;
  requestAppQuit: () => void;
  destroyStatsWindow: () => void;
  logger: {
    info: (message: string) => void;
    warn: (message: string, error?: unknown) => void;
    error: (message: string, error?: unknown) => void;
    debug: (message: string, details?: unknown) => void;
  };
}

export interface StatsRuntimeBootstrap {
  stats: StatsRuntime;
  immersion: AppReadyImmersionInput;
  recordTrackedCardsMined: (count: number, noteIds?: number[]) => void;
}

export function createStatsRuntime<
  TConfig extends StatsConfigLike,
  TTracker extends StatsTrackerLike,
  TServer extends StatsServerLike,
>(input: StatsRuntimeInput<TConfig, TTracker, TServer>): StatsRuntime {
  const now = input.now ?? Date.now;
  const getCurrentPid = input.getCurrentPid ?? (() => process.pid);
  const isProcessAlive = input.isProcessAlive ?? isBackgroundStatsServerProcessAlive;
  const killProcess =
    input.killProcess ??
    ((pid: number, signal: NodeJS.Signals) => {
      process.kill(pid, signal);
    });
  const wait =
    input.wait ??
    (async (delayMs: number) => {
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    });

  let statsServer: TServer | null = null;
  let statsStartupInProgress = false;
  let hasAttemptedImmersionTrackerStartup = false;

  const readLiveBackgroundStatsDaemonState = (): BackgroundStatsServerState | null => {
    const state = readBackgroundStatsServerState(input.statsDaemonStatePath);
    if (!state) {
      removeBackgroundStatsServerState(input.statsDaemonStatePath);
      return null;
    }
    if (state.pid === getCurrentPid() && !statsServer) {
      removeBackgroundStatsServerState(input.statsDaemonStatePath);
      return null;
    }
    if (!isProcessAlive(state.pid)) {
      removeBackgroundStatsServerState(input.statsDaemonStatePath);
      return null;
    }
    return state;
  };

  const clearOwnedBackgroundStatsDaemonState = (): void => {
    const state = readBackgroundStatsServerState(input.statsDaemonStatePath);
    if (state?.pid === getCurrentPid()) {
      removeBackgroundStatsServerState(input.statsDaemonStatePath);
    }
  };

  const stopStatsServer = (): void => {
    if (!statsServer) {
      return;
    }
    statsServer.close();
    statsServer = null;
    clearOwnedBackgroundStatsDaemonState();
  };

  const ensureImmersionTrackerStarted = (): void => {
    if (hasAttemptedImmersionTrackerStartup || input.getImmersionTracker()) {
      return;
    }
    hasAttemptedImmersionTrackerStartup = true;
    statsStartupInProgress = true;
    try {
      input.ensureImmersionTrackerStartedCore();
    } finally {
      statsStartupInProgress = false;
    }
  };

  const ensureStatsServerStarted = (): string => {
    const liveDaemon = readLiveBackgroundStatsDaemonState();
    if (liveDaemon && liveDaemon.pid !== getCurrentPid()) {
      return resolveBackgroundStatsServerUrl(liveDaemon);
    }

    const tracker = input.getImmersionTracker();
    if (!tracker) {
      throw new Error('Immersion tracker failed to initialize.');
    }

    if (!statsServer) {
      statsServer = input.startStatsServer(input.getResolvedConfig().stats.serverPort);
    }

    return `http://127.0.0.1:${input.getResolvedConfig().stats.serverPort}`;
  };

  const ensureBackgroundStatsServerStarted = (): {
    url: string;
    runningInCurrentProcess: boolean;
  } => {
    const liveDaemon = readLiveBackgroundStatsDaemonState();
    if (liveDaemon && liveDaemon.pid !== getCurrentPid()) {
      return {
        url: resolveBackgroundStatsServerUrl(liveDaemon),
        runningInCurrentProcess: false,
      };
    }

    ensureImmersionTrackerStarted();
    const url = ensureStatsServerStarted();
    writeBackgroundStatsServerState(input.statsDaemonStatePath, {
      pid: getCurrentPid(),
      port: input.getResolvedConfig().stats.serverPort,
      startedAtMs: now(),
    });
    return {
      url,
      runningInCurrentProcess: true,
    };
  };

  const stopBackgroundStatsServer = async (): Promise<{ ok: boolean; stale: boolean }> => {
    const state = readBackgroundStatsServerState(input.statsDaemonStatePath);
    if (!state) {
      removeBackgroundStatsServerState(input.statsDaemonStatePath);
      return { ok: true, stale: true };
    }
    if (state.pid === getCurrentPid()) {
      if (!statsServer) {
        removeBackgroundStatsServerState(input.statsDaemonStatePath);
        return { ok: true, stale: true };
      }

      stopStatsServer();
      return { ok: true, stale: false };
    }
    if (!isProcessAlive(state.pid)) {
      removeBackgroundStatsServerState(input.statsDaemonStatePath);
      return { ok: true, stale: true };
    }

    try {
      killProcess(state.pid, 'SIGTERM');
    } catch (error) {
      if ((error as NodeJS.ErrnoException)?.code === 'ESRCH') {
        removeBackgroundStatsServerState(input.statsDaemonStatePath);
        return { ok: true, stale: true };
      }
      if ((error as NodeJS.ErrnoException)?.code === 'EPERM') {
        throw new Error(
          `Insufficient permissions to stop background stats server (pid ${state.pid}).`,
        );
      }
      throw error;
    }

    const deadline = now() + 2_000;
    while (now() < deadline) {
      if (!isProcessAlive(state.pid)) {
        removeBackgroundStatsServerState(input.statsDaemonStatePath);
        return { ok: true, stale: false };
      }
      await wait(50);
    }

    throw new Error('Timed out stopping background stats server.');
  };

  const runStatsCliCommand = createRunStatsCliCommandHandler({
    getResolvedConfig: () => input.getResolvedConfig(),
    ensureImmersionTrackerStarted: () => ensureImmersionTrackerStarted(),
    ensureVocabularyCleanupTokenizerReady: input.ensureVocabularyCleanupTokenizerReady,
    getImmersionTracker: () => input.getImmersionTracker(),
    ensureStatsServerStarted: () => ensureStatsServerStarted(),
    ensureBackgroundStatsServerStarted: () => ensureBackgroundStatsServerStarted(),
    stopBackgroundStatsServer: () => stopBackgroundStatsServer(),
    openExternal: async (url) => await input.openExternal(url),
    writeResponse: (responsePath, payload) => {
      writeStatsCliCommandResponse(responsePath, payload);
    },
    exitAppWithCode: (code) => input.exitAppWithCode(code),
    logInfo: (message) => input.logInfo(message),
    logWarn: (message, error) => input.logWarn(message, error),
    logError: (message, error) => input.logError(message, error),
  });

  const cleanupBeforeQuit = (): void => {
    input.destroyStatsWindow?.();
    stopStatsServer();
  };

  return {
    readLiveBackgroundStatsDaemonState,
    ensureImmersionTrackerStarted,
    ensureStatsServerStarted,
    stopStatsServer,
    ensureBackgroundStatsServerStarted,
    stopBackgroundStatsServer,
    runStatsCliCommand,
    cleanupBeforeQuit,
    getStatsServer: () => statsServer,
    isStatsStartupInProgress: () => statsStartupInProgress,
  };
}

export function createStatsRuntimeBootstrap(
  input: StatsRuntimeBootstrapInput,
): StatsRuntimeBootstrap {
  const statsCoverArtFetcher = createCoverArtFetcher(
    createAnilistRateLimiter(),
    createLogger('main:stats-cover-art'),
  );
  const resolveLegacyVocabularyPos = async (row: {
    headword: string;
    word: string;
    reading: string | null;
  }) => {
    const tokenizer = input.appState.mecabTokenizer;
    if (!tokenizer) {
      return null;
    }

    const lookupTexts = [...new Set([row.headword, row.word, row.reading ?? ''])]
      .map((value) => value.trim())
      .filter((value) => value.length > 0);

    for (const lookupText of lookupTexts) {
      const tokens = await tokenizer.tokenize(lookupText);
      if (!tokens) {
        continue;
      }
      const resolved = resolveLegacyVocabularyPosFromTokens(lookupText, tokens as never);
      if (resolved) {
        return resolved;
      }
    }

    return null;
  };

  let stats: StatsRuntime | null = null;
  const immersionMainDeps: Parameters<typeof createImmersionTrackerStartupHandler>[0] = {
    getResolvedConfig: () => input.getResolvedConfig(),
    getConfiguredDbPath: () => input.dictionarySupport.getConfiguredDbPath(),
    createTrackerService: (params) =>
      new ImmersionTrackerService({
        ...params,
        resolveLegacyVocabularyPos,
      }),
    setTracker: (tracker) => {
      const trackerHasChanged =
        input.appState.immersionTracker !== null && input.appState.immersionTracker !== tracker;
      if (trackerHasChanged && stats?.getStatsServer()) {
        stats.stopStatsServer();
      }

      input.appState.immersionTracker = tracker as ImmersionTrackerService | null;
      input.appState.immersionTracker?.setCoverArtFetcher?.(statsCoverArtFetcher);
      if (!tracker) {
        return;
      }

      if (!stats?.getStatsServer() && input.getResolvedConfig().stats.autoStartServer) {
        stats?.ensureStatsServerStarted();
      }

      input.overlay.registerStatsOverlayToggle({
        staticDir: input.statsDistPath,
        preloadPath: input.statsPreloadPath,
        getApiBaseUrl: () => stats!.ensureStatsServerStarted(),
        getToggleKey: () => input.getResolvedConfig().stats.toggleKey,
        resolveBounds: () => input.overlay.getOverlayGeometry().getCurrentOverlayGeometry(),
        onVisibilityChanged: (visible) => {
          input.appState.statsOverlayVisible = visible;
          input.overlay.updateVisibleOverlayVisibility();
        },
      });
    },
    getMpvClient: () => input.appState.mpvClient as never,
    shouldAutoConnectMpv: () => !stats?.isStatsStartupInProgress(),
    seedTrackerFromCurrentMedia: () => {
      void input.dictionarySupport.seedImmersionMediaFromCurrentMedia();
    },
    logInfo: (message) => input.logger.info(message),
    logDebug: (message) => input.logger.debug(message),
    logWarn: (message, details) => input.logger.warn(message, details),
  };
  const createImmersionTrackerStartup = createImmersionTrackerStartupHandler(immersionMainDeps);

  stats = createStatsRuntime({
    statsDaemonStatePath: input.statsDaemonStatePath,
    getResolvedConfig: () => input.getResolvedConfig(),
    getImmersionTracker: () => input.appState.immersionTracker,
    ensureImmersionTrackerStartedCore: () => {
      createImmersionTrackerStartup();
    },
    ensureVocabularyCleanupTokenizerReady: async () => {
      await input.createMecabTokenizerAndCheck();
    },
    startStatsServer: (port) =>
      startStatsServerCore({
        port,
        staticDir: input.statsDistPath,
        tracker: input.appState.immersionTracker as ImmersionTrackerService,
        knownWordCachePath: path.join(input.userDataPath, 'known-words-cache.json'),
        mpvSocketPath: input.appState.mpvSocketPath,
        ankiConnectConfig: input.getResolvedConfig().ankiConnect,
        resolveAnkiNoteId: (noteId: number) =>
          input.appState.ankiIntegration?.resolveCurrentNoteId(noteId) ?? noteId,
        addYomitanNote: (word: string) => input.addYomitanNote(word),
      }),
    openExternal: (url) => input.openExternal(url),
    exitAppWithCode: (code) => {
      process.exitCode = code;
      input.requestAppQuit();
    },
    destroyStatsWindow: () => {
      input.destroyStatsWindow();
    },
    logInfo: (message) => input.logger.info(message),
    logWarn: (message, error) => input.logger.warn(message, error),
    logError: (message, error) => input.logger.error(message, error),
  });

  return {
    stats,
    immersion: immersionMainDeps as AppReadyImmersionInput,
    recordTrackedCardsMined: (count, noteIds) => {
      stats.ensureImmersionTrackerStarted();
      input.appState.immersionTracker?.recordCardsMined?.(count, noteIds);
    },
  };
}
