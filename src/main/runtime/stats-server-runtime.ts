import { generateSentenceFurigana } from '../../core/services/tokenizer/sentence-furigana';
import path from 'node:path';
import type { BrowserWindow } from 'electron';
import {
  addYomitanNoteViaSearch,
  syncYomitanDefaultAnkiServer as syncYomitanDefaultAnkiServerCore,
} from '../../core/services';
import { startStatsServer, type StatsServer } from '../../core/services/stats-server';
import { createTmdbClient, createTmdbApiKeyResolver } from '../../core/services/tmdb/tmdb-client';
import { createLogger } from '../../logger';
import type { ResolvedConfig } from '../../types/config';
import type { AppState } from '../state';
import {
  isBackgroundStatsServerProcessAlive as defaultIsBackgroundStatsServerProcessAlive,
  readBackgroundStatsServerState as defaultReadBackgroundStatsServerState,
  removeBackgroundStatsServerState as defaultRemoveBackgroundStatsServerState,
  resolveBackgroundStatsServerUrl,
  verifyBackgroundStatsServerIdentity as defaultVerifyBackgroundStatsServerIdentity,
  writeBackgroundStatsServerState,
} from './stats-daemon';
import { createEnsureStatsServerUrlHandler } from './stats-server-routing';
import {
  getPreferredYomitanAnkiServerUrl,
  shouldForceOverrideYomitanAnkiServer,
} from './yomitan-anki-server';

export function isSelfOwnedBackgroundStatsDaemonState(state: {
  pid: number;
  port?: number;
  startedAtMs?: number;
}): boolean {
  return state.pid === process.pid;
}

export interface StatsServerRuntimeDeps {
  userDataPath: string;
  statsDistPath: string;
  getResolvedConfig: () => ResolvedConfig;
  getImmersionTracker: () => AppState['immersionTracker'];
  setAppStateStatsServer: (server: AppState['statsServer']) => void;
  getMpvSocketPath: () => AppState['mpvSocketPath'];
  getYomitanExt: () => AppState['yomitanExt'];
  getYomitanSession: () => AppState['yomitanSession'];
  getYomitanParserWindow: () => AppState['yomitanParserWindow'];
  setYomitanParserWindow: (w: BrowserWindow | null) => void;
  getYomitanParserReadyPromise: () => AppState['yomitanParserReadyPromise'];
  setYomitanParserReadyPromise: (p: Promise<void> | null) => void;
  getYomitanParserInitPromise: () => AppState['yomitanParserInitPromise'];
  setYomitanParserInitPromise: (p: Promise<boolean> | null) => void;
  getYomitanAnkiDeckName: () => Promise<string>;
  getAnilistRateLimiter: () => NonNullable<
    Parameters<typeof startStatsServer>[0]['anilistRateLimiter']
  >;
  /** Project TMDB key staged into release builds; null for source builds. */
  getBundledTmdbApiKey?: () => string | null;
  resolveAnkiNoteId: (noteId: number) => number;
  trackDuplicateNoteIdsForNote: (noteId: number, duplicateNoteIds: number[]) => void;
  resolveSentenceSearchHeadwords: (term: string) => Promise<string[]>;
  ensureImmersionTrackerStarted: () => void;
  setStatsStartupInProgress: (inProgress: boolean) => void;
  readBackgroundStatsServerState?: typeof defaultReadBackgroundStatsServerState;
  removeBackgroundStatsServerState?: typeof defaultRemoveBackgroundStatsServerState;
  isBackgroundStatsServerProcessAlive?: typeof defaultIsBackgroundStatsServerProcessAlive;
  verifyBackgroundStatsServerIdentity?: typeof defaultVerifyBackgroundStatsServerIdentity;
  killProcess?: (pid: number, signal: NodeJS.Signals) => void;
  startServer?: typeof startStatsServer;
}

export function createStatsServerRuntime(deps: StatsServerRuntimeDeps): {
  stopStatsServer: () => Promise<void>;
  ensureStatsServerStarted: ReturnType<typeof createEnsureStatsServerUrlHandler>;
  ensureBackgroundStatsServerStarted: () => Promise<{
    url: string;
    runningInCurrentProcess: boolean;
  }>;
  stopBackgroundStatsServer: () => Promise<{ ok: boolean; stale: boolean }>;
} {
  type LocalStatsServerState =
    | { kind: 'stopped' }
    | { kind: 'starting'; token: symbol; promise: Promise<void> }
    | { kind: 'running'; server: StatsServer }
    | { kind: 'stopping'; token: symbol; promise: Promise<void> };

  let localStatsServerState: LocalStatsServerState = { kind: 'stopped' };
  const pendingBackgroundStarts = new Set<symbol>();
  const statsDaemonStatePath = path.join(deps.userDataPath, 'stats-daemon.json');
  const startServer = deps.startServer ?? startStatsServer;
  const readDaemonState =
    deps.readBackgroundStatsServerState ??
    ((statePath: string) => defaultReadBackgroundStatsServerState(statePath));
  const removeDaemonState =
    deps.removeBackgroundStatsServerState ??
    ((statePath: string) => defaultRemoveBackgroundStatsServerState(statePath));
  const isDaemonAlive =
    deps.isBackgroundStatsServerProcessAlive ??
    ((pid: number) => defaultIsBackgroundStatsServerProcessAlive(pid));
  const verifyDaemonIdentity =
    deps.verifyBackgroundStatsServerIdentity ??
    ((pid: number, startedAtMs: number) =>
      defaultVerifyBackgroundStatsServerIdentity(pid, startedAtMs));
  const killProcess = deps.killProcess ?? ((pid, signal) => process.kill(pid, signal));

  function readLiveBackgroundStatsDaemonState(): {
    pid: number;
    port: number;
    startedAtMs: number;
  } | null {
    const state = readDaemonState(statsDaemonStatePath);
    if (!state) {
      removeDaemonState(statsDaemonStatePath);
      return null;
    }
    if (state.pid === process.pid && localStatsServerState.kind !== 'running') {
      removeDaemonState(statsDaemonStatePath);
      return null;
    }
    if (!isDaemonAlive(state.pid)) {
      removeDaemonState(statsDaemonStatePath);
      return null;
    }
    return state;
  }

  function clearOwnedBackgroundStatsDaemonState(): void {
    const state = readDaemonState(statsDaemonStatePath);
    if (state?.pid === process.pid) {
      removeDaemonState(statsDaemonStatePath);
    }
  }

  const buildStatsServerConfig = (): Parameters<typeof startStatsServer>[0] => {
    const tracker = deps.getImmersionTracker();
    if (!tracker) {
      throw new Error('Immersion tracker failed to initialize.');
    }
    const yomitanDeps = {
      getYomitanExt: () => deps.getYomitanExt(),
      getYomitanSession: () => deps.getYomitanSession(),
      getYomitanParserWindow: () => deps.getYomitanParserWindow(),
      setYomitanParserWindow: (w: BrowserWindow | null) => {
        deps.setYomitanParserWindow(w);
      },
      getYomitanParserReadyPromise: () => deps.getYomitanParserReadyPromise(),
      setYomitanParserReadyPromise: (p: Promise<void> | null) => {
        deps.setYomitanParserReadyPromise(p);
      },
      getYomitanParserInitPromise: () => deps.getYomitanParserInitPromise(),
      setYomitanParserInitPromise: (p: Promise<boolean> | null) => {
        deps.setYomitanParserInitPromise(p);
      },
    };
    const yomitanLogger = createLogger('main:yomitan-stats');
    return {
      port: deps.getResolvedConfig().stats.serverPort,
      staticDir: deps.statsDistPath,
      tracker,
      knownWordCachePath: path.join(deps.userDataPath, 'known-words-cache.json'),
      mpvSocketPath: deps.getMpvSocketPath(),
      getAnkiConnectConfig: () => deps.getResolvedConfig().ankiConnect,
      getYomitanAnkiDeckName: deps.getYomitanAnkiDeckName,
      getSecondarySubtitleLanguages: () =>
        deps.getResolvedConfig().secondarySub.secondarySubLanguages,
      getStatsMiningAlassPath: () => deps.getResolvedConfig().subsync.alass_path,
      anilistRateLimiter: deps.getAnilistRateLimiter(),
      tmdbClient: createTmdbClient({
        resolveApiKey: createTmdbApiKeyResolver(
          () => deps.getResolvedConfig().tmdb,
          () => deps.getBundledTmdbApiKey?.() ?? null,
        ),
      }),
      resolveAnkiNoteId: (noteId: number) => deps.resolveAnkiNoteId(noteId),
      resolveSentenceSearchHeadwords: (term: string) => deps.resolveSentenceSearchHeadwords(term),
      generateSentenceFurigana: (text, highlightedText) =>
        generateSentenceFurigana(text, highlightedText, yomitanDeps, yomitanLogger),
      addYomitanNote: async (word: string) => {
        const ankiConnectConfig = deps.getResolvedConfig().ankiConnect;
        const ankiUrl = getPreferredYomitanAnkiServerUrl(ankiConnectConfig);
        await syncYomitanDefaultAnkiServerCore(ankiUrl, yomitanDeps, yomitanLogger, {
          forceOverride: shouldForceOverrideYomitanAnkiServer(ankiConnectConfig),
          deck: ankiConnectConfig.deck,
          ankiConfig: ankiConnectConfig,
        });
        const result = await addYomitanNoteViaSearch(word, yomitanDeps, yomitanLogger);
        if (result.noteId && result.duplicateNoteIds.length > 0) {
          deps.trackDuplicateNoteIdsForNote(result.noteId, result.duplicateNoteIds);
        }
        return result.noteId;
      },
    };
  };

  const beginLocalStatsServerStartup = (): Promise<void> => {
    const token = Symbol('stats-server-startup');
    const promise = startServer(buildStatsServerConfig())
      .then(async (server) => {
        const state = localStatsServerState;
        if (state.kind !== 'starting' || state.token !== token) {
          await server.close();
          throw new Error('Stats server startup was cancelled.');
        }
        localStatsServerState = { kind: 'running', server };
        deps.setAppStateStatsServer(server);
      })
      .catch((error: unknown) => {
        const state = localStatsServerState;
        if (state.kind === 'starting' && state.token === token) {
          localStatsServerState = { kind: 'stopped' };
          deps.setAppStateStatsServer(null);
        }
        throw error;
      });
    localStatsServerState = { kind: 'starting', token, promise };
    return promise;
  };

  const startLocalStatsServer = async (): Promise<void> => {
    while (localStatsServerState.kind === 'stopping') {
      await localStatsServerState.promise;
    }
    if (localStatsServerState.kind === 'running') {
      deps.setAppStateStatsServer(localStatsServerState.server);
      return;
    }
    if (localStatsServerState.kind === 'starting') {
      await localStatsServerState.promise;
      return;
    }
    await beginLocalStatsServerStartup();
  };

  function stopStatsServer(): Promise<void> {
    const state = localStatsServerState;
    if (state.kind === 'stopped') {
      deps.setAppStateStatsServer(null);
      clearOwnedBackgroundStatsDaemonState();
      return Promise.resolve();
    }
    if (state.kind === 'stopping') {
      return state.promise;
    }

    const token = Symbol('stats-server-shutdown');
    const promise = Promise.resolve()
      .then(async () => {
        if (state.kind === 'starting') {
          try {
            await state.promise;
          } catch {
            // Startup owns cleanup of a server that finishes binding after cancellation.
          }
          return;
        }
        await state.server.close();
      })
      .finally(() => {
        const current = localStatsServerState;
        if (current.kind === 'stopping' && current.token === token) {
          localStatsServerState = { kind: 'stopped' };
        }
        deps.setAppStateStatsServer(null);
        clearOwnedBackgroundStatsDaemonState();
      });
    localStatsServerState = { kind: 'stopping', token, promise };
    deps.setAppStateStatsServer(null);
    return promise;
  }

  const ensureStatsServerStarted = createEnsureStatsServerUrlHandler({
    currentPid: process.pid,
    readBackgroundState: () => readDaemonState(statsDaemonStatePath),
    removeBackgroundState: () => {
      removeDaemonState(statsDaemonStatePath);
    },
    isProcessAlive: (pid) => isDaemonAlive(pid),
    hasLocalStatsServer: () => localStatsServerState.kind === 'running',
    startLocalStatsServer,
    getConfiguredPort: () => deps.getResolvedConfig().stats.serverPort,
  });

  const ensureBackgroundStatsServerStarted = async (): Promise<{
    url: string;
    runningInCurrentProcess: boolean;
  }> => {
    const liveDaemon = readLiveBackgroundStatsDaemonState();
    if (liveDaemon && liveDaemon.pid !== process.pid) {
      return {
        url: resolveBackgroundStatsServerUrl(liveDaemon),
        runningInCurrentProcess: false,
      };
    }

    deps.setStatsStartupInProgress(true);
    try {
      deps.ensureImmersionTrackerStarted();
    } finally {
      deps.setStatsStartupInProgress(false);
    }

    const request = Symbol('background-stats-startup');
    pendingBackgroundStarts.add(request);
    try {
      const port = deps.getResolvedConfig().stats.serverPort;
      const result = await ensureStatsServerStarted();
      if (result.source === 'local') {
        if (localStatsServerState.kind !== 'running') {
          throw new Error('Stats server startup was cancelled.');
        }
        writeBackgroundStatsServerState(statsDaemonStatePath, {
          pid: process.pid,
          port,
          startedAtMs: Date.now(),
        });
      }
      return { url: result.url, runningInCurrentProcess: result.source === 'local' };
    } finally {
      pendingBackgroundStarts.delete(request);
    }
  };

  const stopBackgroundStatsServer = async (): Promise<{ ok: boolean; stale: boolean }> => {
    const state = readDaemonState(statsDaemonStatePath);
    if (!state) {
      if (pendingBackgroundStarts.size > 0) {
        await stopStatsServer();
        return { ok: true, stale: false };
      }
      removeDaemonState(statsDaemonStatePath);
      return { ok: true, stale: true };
    }
    if (isSelfOwnedBackgroundStatsDaemonState(state)) {
      await stopStatsServer();
      return { ok: true, stale: false };
    }
    if (!isDaemonAlive(state.pid)) {
      removeDaemonState(statsDaemonStatePath);
      return { ok: true, stale: true };
    }
    if (!verifyDaemonIdentity(state.pid, state.startedAtMs)) {
      removeDaemonState(statsDaemonStatePath);
      return { ok: true, stale: true };
    }

    try {
      killProcess(state.pid, 'SIGTERM');
    } catch (error) {
      if ((error as NodeJS.ErrnoException)?.code === 'ESRCH') {
        removeDaemonState(statsDaemonStatePath);
        return { ok: true, stale: true };
      }
      if ((error as NodeJS.ErrnoException)?.code === 'EPERM') {
        throw new Error(
          `Insufficient permissions to stop background stats server (pid ${state.pid}).`,
        );
      }
      throw error;
    }

    const deadline = Date.now() + 2_000;
    while (Date.now() < deadline) {
      if (!isDaemonAlive(state.pid)) {
        removeDaemonState(statsDaemonStatePath);
        return { ok: true, stale: false };
      }
      await new Promise((resolve) => setTimeout(resolve, 50));
    }

    throw new Error('Timed out stopping background stats server.');
  };

  return {
    stopStatsServer,
    ensureStatsServerStarted,
    ensureBackgroundStatsServerStarted,
    stopBackgroundStatsServer,
  };
}
