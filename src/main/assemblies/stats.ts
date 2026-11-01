import path from 'node:path';
import { ImmersionTrackerService } from '../../core/services';
import { createCoverArtFetcher } from '../../core/services/anilist/cover-art-fetcher';
import type { createAnilistRateLimiter } from '../../core/services/anilist/rate-limiter';
import { resolveLegacyVocabularyPosFromTokens } from '../../core/services/immersion-tracker/legacy-vocabulary-pos';
import type { StatsWindowOptions } from '../../core/services/stats-window';
import { readBundledTmdbApiKey } from '../../core/services/tmdb/bundled-api-key';
import { createLiveActionMetadataResolver } from '../../core/services/tmdb/live-action-resolver';
import { createTmdbApiKeyResolver, createTmdbClient } from '../../core/services/tmdb/tmdb-client';
import { createLogger, type Logger } from '../../logger';
import type { ResolvedConfig } from '../../types';
import { createEnsureBackgroundStatsServerHandler } from '../runtime/background-stats-startup';
import { createImmersionMediaRuntime } from '../runtime/domains/startup';
import {
  createImmersionTrackerStartupHandler,
  type ImmersionTrackerStartupDeps,
} from '../runtime/immersion-startup';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import {
  createRunStatsCliCommandHandler,
  writeStatsCliCommandResponse,
} from '../runtime/stats-cli-command';
import {
  createStatsServerRuntime,
  type StatsServerRuntimeDeps,
} from '../runtime/stats-server-runtime';
import type { AppState } from '../state';
import type { PlaybackAssembly } from './playback';
import type { YomitanAssembly } from './yomitan';

/**
 * What the stats domain needs from the rest of the app. Siblings built after this assembly in
 * `main.ts` are reached through callbacks.
 */
export interface StatsAssemblyDeps {
  appState: Pick<
    AppState,
    | 'immersionTracker'
    | 'statsServer'
    | 'statsStartupInProgress'
    | 'mpvClient'
    | 'mpvSocketPath'
    | 'mecabTokenizer'
    | 'ankiIntegration'
    | 'currentMediaPath'
    | 'currentMediaTitle'
  >;
  getConfig: () => ResolvedConfig;
  logger: Pick<Logger, 'info' | 'warn' | 'error' | 'debug'>;
  userDataPath: string;
  /** Compiled main-process directory; the stats bundle, preload and TMDB key sit beside it. */
  mainDirname: string;
  defaultImmersionDbPath: string;
  anilistRateLimiter: ReturnType<typeof createAnilistRateLimiter>;
  /** The stats window module reaches Electron's `dialog`, so it is injected. */
  statsWindow: {
    registerToggle: (options: StatsWindowOptions) => void;
    toggle: (options: StatsWindowOptions) => Promise<void>;
  };
  yomitanParserRuntimeDeps: ReturnType<YomitanAssembly['getParserRuntimeDeps']>;
  getYomitanAnkiDeckName: () => Promise<string>;
  getTokenizeSubtitle: PlaybackAssembly['getTokenizeSubtitleDeferred'];
  createMecabTokenizerAndCheck: () => Promise<void>;
  getOverlayNotifications: () => Pick<
    ReturnType<typeof createOverlayNotificationsRuntime>,
    'showConfiguredStatusNotification'
  >;
  getOverlayBounds: () => ReturnType<StatsWindowOptions['resolveBounds']>;
  onStatsOverlayVisibilityChanged: (visible: boolean) => void;
  /** Forced AniList post-watch update after a manual watched mark. */
  runAnilistPostWatchUpdate: () => Promise<void>;
  openExternal: (url: string) => Promise<void>;
  requestAppQuit: () => void;
}

export type StatsAssembly = ReturnType<typeof createStatsAssembly>;

/**
 * Owns the immersion tracker's startup, the stats HTTP server (in-process or background), the
 * stats overlay window options, and the `--stats` CLI command.
 */
export function createStatsAssembly(deps: StatsAssemblyDeps) {
  const { appState, logger } = deps;
  // Release builds stage a project TMDB key next to the compiled main process.
  const bundledTmdbApiKey = readBundledTmdbApiKey(deps.mainDirname);
  let hasAttemptedImmersionTrackerStartup = false;

  async function resolveSentenceSearchHeadwords(term: string): Promise<string[]> {
    const fallback = term.trim() ? [term.trim()] : [];
    try {
      const tokenizeSubtitleDeferred = deps.getTokenizeSubtitle();
      const tokenized = tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(term) : null;
      const tokens = tokenized?.tokens ?? [];
      if (tokens.length === 0) return fallback;

      const seen = new Set<string>();
      const headwords: string[] = [];
      for (const token of tokens) {
        const headword = (token.headword || token.surface).trim();
        if (!headword || seen.has(headword)) continue;
        seen.add(headword);
        headwords.push(headword);
      }
      return headwords.length > 0 ? headwords : fallback;
    } catch (error) {
      logger.debug(
        'Failed to resolve sentence-search headwords:',
        error instanceof Error ? error.message : String(error),
      );
      return fallback;
    }
  }

  const statsServerRuntime = createStatsServerRuntime({
    userDataPath: deps.userDataPath,
    statsDistPath: path.join(deps.mainDirname, '..', 'stats', 'dist'),
    getResolvedConfig: () => deps.getConfig(),
    getImmersionTracker: () => appState.immersionTracker,
    setAppStateStatsServer: (server) => {
      appState.statsServer = server;
    },
    getMpvSocketPath: () => appState.mpvSocketPath,
    ...(deps.yomitanParserRuntimeDeps satisfies Pick<
      StatsServerRuntimeDeps,
      keyof ReturnType<YomitanAssembly['getParserRuntimeDeps']>
    >),
    getYomitanAnkiDeckName: () => deps.getYomitanAnkiDeckName(),
    getAnilistRateLimiter: () => deps.anilistRateLimiter,
    getBundledTmdbApiKey: () => bundledTmdbApiKey,
    resolveAnkiNoteId: (noteId) => appState.ankiIntegration?.resolveCurrentNoteId(noteId) ?? noteId,
    trackDuplicateNoteIdsForNote: (noteId, duplicateNoteIds) => {
      appState.ankiIntegration?.trackDuplicateNoteIdsForNote(noteId, duplicateNoteIds);
    },
    resolveSentenceSearchHeadwords: (term) => resolveSentenceSearchHeadwords(term),
    ensureImmersionTrackerStarted: () => ensureImmersionTrackerStarted(),
    setStatsStartupInProgress: (inProgress) => {
      appState.statsStartupInProgress = inProgress;
    },
  });
  const {
    stopStatsServer,
    ensureStatsServerStarted,
    ensureBackgroundStatsServerStarted,
    stopBackgroundStatsServer,
  } = statsServerRuntime;

  const immersionMediaRuntime = createImmersionMediaRuntime({
    getResolvedConfig: () => deps.getConfig(),
    defaultImmersionDbPath: deps.defaultImmersionDbPath,
    getTracker: () => appState.immersionTracker,
    getMpvClient: () => appState.mpvClient,
    getCurrentMediaPath: () => appState.currentMediaPath,
    getCurrentMediaTitle: () => appState.currentMediaTitle,
    logDebug: (message) => logger.debug(message),
    logInfo: (message) => logger.info(message),
  });

  const coverArtFetcher = createCoverArtFetcher(
    deps.anilistRateLimiter,
    createLogger('main:stats-cover-art'),
    {
      liveAction: createLiveActionMetadataResolver(
        createTmdbClient({
          resolveApiKey: createTmdbApiKeyResolver(
            () => deps.getConfig().tmdb,
            () => bundledTmdbApiKey,
          ),
        }),
        createLogger('main:tmdb'),
      ),
    },
  );

  // One options object serves the overlay's IPC toggle and the session-action toggle.
  const statsOverlayOptions: StatsWindowOptions = {
    preloadPath: path.join(deps.mainDirname, 'preload-stats.js'),
    getApiBaseUrl: async () => (await ensureStatsServerStarted()).url,
    onStartupError: (error) =>
      deps
        .getOverlayNotifications()
        .showConfiguredStatusNotification(
          `Stats server startup failed: ${error instanceof Error ? error.message : String(error)}`,
          { title: 'Stats' },
        ),
    getToggleKey: () => deps.getConfig().stats.toggleKey,
    resolveBounds: () => deps.getOverlayBounds(),
    onVisibilityChanged: (visible) => {
      deps.onStatsOverlayVisibilityChanged(visible);
    },
  };

  const resolveLegacyVocabularyPos = async (row: {
    headword: string;
    word: string;
    reading: string | null;
  }) => {
    const tokenizer = appState.mecabTokenizer;
    if (!tokenizer) {
      return null;
    }

    const lookupTexts = [...new Set([row.headword, row.word, row.reading ?? ''])]
      .map((value) => value.trim())
      .filter((value) => value.length > 0);

    for (const lookupText of lookupTexts) {
      const tokens = await tokenizer.tokenize(lookupText);
      const resolved = resolveLegacyVocabularyPosFromTokens(lookupText, tokens);
      if (resolved) {
        return resolved;
      }
    }

    return null;
  };

  const immersionTrackerStartupDeps: ImmersionTrackerStartupDeps = {
    getResolvedConfig: () => deps.getConfig(),
    getConfiguredDbPath: () => immersionMediaRuntime.getConfiguredDbPath(),
    createTrackerService: (params) =>
      new ImmersionTrackerService({
        ...params,
        resolveLegacyVocabularyPos,
      }),
    setTracker: (tracker) => {
      const trackerHasChanged =
        appState.immersionTracker !== null && appState.immersionTracker !== tracker;
      if (trackerHasChanged && appState.statsServer) {
        void stopStatsServer().catch((error: unknown) => {
          logger.warn('Failed to stop stats server while replacing immersion tracker.', error);
        });
        appState.statsServer = null;
      }

      appState.immersionTracker = tracker as ImmersionTrackerService | null;
      appState.immersionTracker?.setCoverArtFetcher(coverArtFetcher);
      if (tracker) {
        // Start HTTP stats server
        if (!appState.statsServer) {
          const config = deps.getConfig();
          if (config.stats.autoStartServer) {
            void ensureStatsServerStarted().catch((error: unknown) => {
              logger.warn('Failed to auto-start stats server.', error);
            });
          }
        }

        // Register stats overlay toggle IPC handler (idempotent)
        deps.statsWindow.registerToggle(statsOverlayOptions);
      }
    },
    getMpvClient: () => appState.mpvClient,
    shouldAutoConnectMpv: () => !appState.statsStartupInProgress,
    seedTrackerFromCurrentMedia: () => {
      void immersionMediaRuntime.seedFromCurrentMedia();
    },
    logInfo: (message) => logger.info(message),
    logDebug: (message) => logger.debug(message),
    logWarn: (message, details) => logger.warn(message, details),
  };
  const createImmersionTrackerStartup = createImmersionTrackerStartupHandler(
    immersionTrackerStartupDeps,
  );

  /** Starts the tracker at most once per process; later calls are no-ops. */
  function ensureImmersionTrackerStarted(): void {
    if (hasAttemptedImmersionTrackerStartup || appState.immersionTracker) {
      return;
    }
    hasAttemptedImmersionTrackerStartup = true;
    createImmersionTrackerStartup();
  }

  // Stats startup must not auto-connect mpv, so the tracker starts under this flag.
  function ensureImmersionTrackerStartedForStats(): void {
    appState.statsStartupInProgress = true;
    try {
      ensureImmersionTrackerStarted();
    } finally {
      appState.statsStartupInProgress = false;
    }
  }

  const ensureBackgroundStatsServer = createEnsureBackgroundStatsServerHandler({
    isStatsAutoStartEnabled: () => deps.getConfig().stats.autoStartServer,
    isImmersionTrackingEnabled: () => deps.getConfig().immersionTracking?.enabled !== false,
    ensureBackgroundStatsServerStarted: () => ensureBackgroundStatsServerStarted(),
    logInfo: (message) => logger.info(message),
    logWarn: (message, error) => logger.warn(message, error),
  });

  const runStatsCliCommand = createRunStatsCliCommandHandler({
    getResolvedConfig: () => deps.getConfig(),
    ensureImmersionTrackerStarted: () => ensureImmersionTrackerStartedForStats(),
    ensureVocabularyCleanupTokenizerReady: async () => {
      await deps.createMecabTokenizerAndCheck();
    },
    getImmersionTracker: () => appState.immersionTracker,
    ensureStatsServerStarted: async () => (await ensureStatsServerStarted()).url,
    ensureBackgroundStatsServerStarted: () => ensureBackgroundStatsServerStarted(),
    stopBackgroundStatsServer: () => stopBackgroundStatsServer(),
    openExternal: (url: string) => deps.openExternal(url),
    writeResponse: (responsePath, payload) => {
      writeStatsCliCommandResponse(responsePath, payload);
    },
    exitAppWithCode: (code) => {
      process.exitCode = code;
      deps.requestAppQuit();
    },
    logInfo: (message) => logger.info(message),
    logWarn: (message, error) => logger.warn(message, error),
    logError: (message, error) => logger.error(message, error),
  });

  return {
    immersionMediaRuntime,
    immersionTrackerStartupDeps,
    ensureImmersionTrackerStarted,
    stopStatsServer,
    ensureBackgroundStatsServer,
    runStatsCliCommand,
    toggleStatsOverlay: () => deps.statsWindow.toggle(statsOverlayOptions),
    recordCardsMined: (count: number, noteIds?: number[]): void => {
      ensureImmersionTrackerStarted();
      appState.immersionTracker?.recordCardsMined(count, noteIds);
    },
    /** Marks the active video watched, then forces the AniList post-watch update. */
    markActiveVideoWatched: async (): Promise<boolean> => {
      ensureImmersionTrackerStarted();
      const marked = (await appState.immersionTracker?.markActiveVideoWatched()) ?? false;
      if (marked) {
        try {
          await deps.runAnilistPostWatchUpdate();
        } catch (error) {
          logger.warn('Failed to run AniList post-watch update after manual watched mark:', error);
        }
      }
      return marked;
    },
  };
}
