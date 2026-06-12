import path from 'node:path';
import type { BrowserWindow } from 'electron';
import {
  addYomitanNoteViaSearch,
  syncYomitanDefaultAnkiServer as syncYomitanDefaultAnkiServerCore,
} from '../../core/services';
import { startStatsServer } from '../../core/services/stats-server';
import { createLogger } from '../../logger';
import type { ResolvedConfig } from '../../types/config';
import type { AppState } from '../state';
import {
  isBackgroundStatsServerProcessAlive,
  readBackgroundStatsServerState,
  removeBackgroundStatsServerState,
  resolveBackgroundStatsServerUrl,
  writeBackgroundStatsServerState,
} from './stats-daemon';
import { createEnsureStatsServerUrlHandler } from './stats-server-routing';
import { shouldForceOverrideYomitanAnkiServer } from './yomitan-anki-server';

export function isSelfOwnedBackgroundStatsDaemonState(state: {
  pid: number;
  port?: number;
  startedAtMs?: number;
}): boolean {
  return state.pid === process.pid;
}

export function shouldClearAppStateStatsServerOnStop(options: {
  hadStatsServer: boolean;
}): boolean {
  return options.hadStatsServer;
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
  resolveAnkiNoteId: (noteId: number) => number;
  trackDuplicateNoteIdsForNote: (noteId: number, duplicateNoteIds: number[]) => void;
  resolveSentenceSearchHeadwords: (term: string) => Promise<string[]>;
  ensureImmersionTrackerStarted: () => void;
  setStatsStartupInProgress: (inProgress: boolean) => void;
}

export function createStatsServerRuntime(deps: StatsServerRuntimeDeps): {
  stopStatsServer: () => void;
  ensureStatsServerStarted: ReturnType<typeof createEnsureStatsServerUrlHandler>;
  ensureBackgroundStatsServerStarted: () => {
    url: string;
    runningInCurrentProcess: boolean;
  };
  stopBackgroundStatsServer: () => Promise<{ ok: boolean; stale: boolean }>;
} {
  let statsServer: ReturnType<typeof startStatsServer> | null = null;
  const statsDaemonStatePath = path.join(deps.userDataPath, 'stats-daemon.json');

  function readLiveBackgroundStatsDaemonState(): {
    pid: number;
    port: number;
    startedAtMs: number;
  } | null {
    const state = readBackgroundStatsServerState(statsDaemonStatePath);
    if (!state) {
      removeBackgroundStatsServerState(statsDaemonStatePath);
      return null;
    }
    if (state.pid === process.pid && !statsServer) {
      removeBackgroundStatsServerState(statsDaemonStatePath);
      return null;
    }
    if (!isBackgroundStatsServerProcessAlive(state.pid)) {
      removeBackgroundStatsServerState(statsDaemonStatePath);
      return null;
    }
    return state;
  }

  function clearOwnedBackgroundStatsDaemonState(): void {
    const state = readBackgroundStatsServerState(statsDaemonStatePath);
    if (state?.pid === process.pid) {
      removeBackgroundStatsServerState(statsDaemonStatePath);
    }
  }

  function stopStatsServer(): void {
    if (!statsServer) {
      return;
    }
    statsServer.close();
    statsServer = null;
    if (shouldClearAppStateStatsServerOnStop({ hadStatsServer: true })) {
      deps.setAppStateStatsServer(null);
    }
    clearOwnedBackgroundStatsDaemonState();
  }

  const startLocalStatsServer = (): void => {
    const tracker = deps.getImmersionTracker();
    if (!tracker) {
      throw new Error('Immersion tracker failed to initialize.');
    }
    if (!statsServer) {
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
      statsServer = startStatsServer({
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
        resolveAnkiNoteId: (noteId: number) => deps.resolveAnkiNoteId(noteId),
        resolveSentenceSearchHeadwords: (term: string) => deps.resolveSentenceSearchHeadwords(term),
        addYomitanNote: async (word: string) => {
          const ankiConnectConfig = deps.getResolvedConfig().ankiConnect;
          const ankiUrl = ankiConnectConfig.url || 'http://127.0.0.1:8765';
          await syncYomitanDefaultAnkiServerCore(ankiUrl, yomitanDeps, yomitanLogger, {
            forceOverride: shouldForceOverrideYomitanAnkiServer(ankiConnectConfig),
            deck: ankiConnectConfig.deck,
          });
          const result = await addYomitanNoteViaSearch(word, yomitanDeps, yomitanLogger);
          if (result.noteId && result.duplicateNoteIds.length > 0) {
            deps.trackDuplicateNoteIdsForNote(result.noteId, result.duplicateNoteIds);
          }
          return result.noteId;
        },
      });
      deps.setAppStateStatsServer(statsServer);
    }
    deps.setAppStateStatsServer(statsServer);
  };

  const ensureStatsServerStarted = createEnsureStatsServerUrlHandler({
    currentPid: process.pid,
    readBackgroundState: () => readBackgroundStatsServerState(statsDaemonStatePath),
    removeBackgroundState: () => {
      removeBackgroundStatsServerState(statsDaemonStatePath);
    },
    isProcessAlive: (pid) => isBackgroundStatsServerProcessAlive(pid),
    hasLocalStatsServer: () => statsServer !== null,
    startLocalStatsServer,
    getConfiguredPort: () => deps.getResolvedConfig().stats.serverPort,
  });

  const ensureBackgroundStatsServerStarted = (): {
    url: string;
    runningInCurrentProcess: boolean;
  } => {
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

    const port = deps.getResolvedConfig().stats.serverPort;
    const result = ensureStatsServerStarted();
    if (result.source === 'local') {
      writeBackgroundStatsServerState(statsDaemonStatePath, {
        pid: process.pid,
        port,
        startedAtMs: Date.now(),
      });
    }
    return { url: result.url, runningInCurrentProcess: result.source === 'local' };
  };

  const stopBackgroundStatsServer = async (): Promise<{ ok: boolean; stale: boolean }> => {
    const state = readBackgroundStatsServerState(statsDaemonStatePath);
    if (!state) {
      removeBackgroundStatsServerState(statsDaemonStatePath);
      return { ok: true, stale: true };
    }
    if (isSelfOwnedBackgroundStatsDaemonState(state)) {
      removeBackgroundStatsServerState(statsDaemonStatePath);
      return { ok: true, stale: true };
    }
    if (!isBackgroundStatsServerProcessAlive(state.pid)) {
      removeBackgroundStatsServerState(statsDaemonStatePath);
      return { ok: true, stale: true };
    }

    try {
      process.kill(state.pid, 'SIGTERM');
    } catch (error) {
      if ((error as NodeJS.ErrnoException)?.code === 'ESRCH') {
        removeBackgroundStatsServerState(statsDaemonStatePath);
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
      if (!isBackgroundStatsServerProcessAlive(state.pid)) {
        removeBackgroundStatsServerState(statsDaemonStatePath);
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
