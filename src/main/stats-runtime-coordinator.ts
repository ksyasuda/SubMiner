import type { BrowserWindow } from 'electron';
import { shell } from 'electron';
import path from 'node:path';

import type { CliArgs, CliCommandSource } from '../cli/args';
import type { ResolvedConfig } from '../types';
import {
  addYomitanNoteViaSearch,
  syncYomitanDefaultAnkiServer as syncYomitanDefaultAnkiServerCore,
} from '../core/services';
import { createLogger } from '../logger';
import type { AppState } from './state';
import {
  createStatsRuntimeBootstrap,
  type StatsRuntime,
  type StatsRuntimeBootstrap,
} from './stats-runtime';
import { registerStatsOverlayToggle, destroyStatsWindow } from '../core/services/stats-window.js';

export interface StatsRuntimeCoordinatorInput {
  statsDaemonStatePath: string;
  statsDistPath: string;
  statsPreloadPath: string;
  userDataPath: string;
  appState: AppState;
  getResolvedConfig: () => ResolvedConfig;
  dictionarySupport: {
    getConfiguredDbPath: () => string;
    seedImmersionMediaFromCurrentMedia: () => Promise<void> | void;
  };
  overlay: {
    getOverlayGeometry: () => { getCurrentOverlayGeometry: () => Electron.Rectangle };
    updateVisibleOverlayVisibility: () => void;
    registerStatsOverlayToggle: StatsRuntimeBootstrap['stats'] extends never
      ? never
      : Parameters<typeof createStatsRuntimeBootstrap>[0]['overlay']['registerStatsOverlayToggle'];
  };
  mpvRuntime: {
    createMecabTokenizerAndCheck: () => Promise<void>;
  };
  actions: {
    openExternal: (url: string) => Promise<unknown>;
    requestAppQuit: () => void;
    destroyStatsWindow: () => void;
  };
  logger: {
    info: (message: string) => void;
    warn: (message: string, error?: unknown) => void;
    error: (message: string, error?: unknown) => void;
    debug: (message: string, details?: unknown) => void;
  };
}

export interface StatsRuntimeCoordinator {
  statsBootstrap: StatsRuntimeBootstrap;
  stats: StatsRuntime;
  ensureStatsServerStarted: () => string;
  ensureBackgroundStatsServerStarted: () => {
    url: string;
    runningInCurrentProcess: boolean;
  };
  stopBackgroundStatsServer: () => Promise<{ ok: boolean; stale: boolean }>;
  ensureImmersionTrackerStarted: () => void;
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
}

export function createStatsRuntimeCoordinator(
  input: StatsRuntimeCoordinatorInput,
): StatsRuntimeCoordinator {
  const statsBootstrap = createStatsRuntimeBootstrap({
    statsDaemonStatePath: input.statsDaemonStatePath,
    statsDistPath: input.statsDistPath,
    statsPreloadPath: input.statsPreloadPath,
    userDataPath: input.userDataPath,
    appState: input.appState,
    getResolvedConfig: () => input.getResolvedConfig(),
    dictionarySupport: input.dictionarySupport,
    overlay: input.overlay,
    createMecabTokenizerAndCheck: async () => {
      await input.mpvRuntime.createMecabTokenizerAndCheck();
    },
    addYomitanNote: async (word: string) => {
      const yomitanDeps = {
        getYomitanExt: () => input.appState.yomitanExt,
        getYomitanSession: () => input.appState.yomitanSession,
        getYomitanParserWindow: () => input.appState.yomitanParserWindow,
        setYomitanParserWindow: (window: BrowserWindow | null) => {
          input.appState.yomitanParserWindow = window;
        },
        getYomitanParserReadyPromise: () => input.appState.yomitanParserReadyPromise,
        setYomitanParserReadyPromise: (promise: Promise<void> | null) => {
          input.appState.yomitanParserReadyPromise = promise;
        },
        getYomitanParserInitPromise: () => input.appState.yomitanParserInitPromise,
        setYomitanParserInitPromise: (promise: Promise<boolean> | null) => {
          input.appState.yomitanParserInitPromise = promise;
        },
      };
      const yomitanLogger = createLogger('main:yomitan-stats');
      const ankiUrl = input.getResolvedConfig().ankiConnect.url || 'http://127.0.0.1:8765';
      await syncYomitanDefaultAnkiServerCore(ankiUrl, yomitanDeps, yomitanLogger, {
        forceOverride: true,
      });
      return addYomitanNoteViaSearch(word, yomitanDeps, yomitanLogger);
    },
    openExternal: input.actions.openExternal,
    requestAppQuit: input.actions.requestAppQuit,
    destroyStatsWindow: input.actions.destroyStatsWindow,
    logger: input.logger,
  });

  const stats = statsBootstrap.stats;

  return {
    statsBootstrap,
    stats,
    ensureStatsServerStarted: () => stats.ensureStatsServerStarted(),
    ensureBackgroundStatsServerStarted: () => stats.ensureBackgroundStatsServerStarted(),
    stopBackgroundStatsServer: async () => await stats.stopBackgroundStatsServer(),
    ensureImmersionTrackerStarted: () => {
      stats.ensureImmersionTrackerStarted();
    },
    runStatsCliCommand: async (args, source) => {
      await stats.runStatsCliCommand(args, source);
    },
  };
}

export interface StatsRuntimeFromMainStateInput {
  dirname: string;
  userDataPath: string;
  appState: AppState;
  getResolvedConfig: () => ResolvedConfig;
  dictionarySupport: StatsRuntimeCoordinatorInput['dictionarySupport'];
  overlay: {
    getOverlayGeometry: () => { getCurrentOverlayGeometry: () => Electron.Rectangle };
    updateVisibleOverlayVisibility: () => void;
  };
  mpvRuntime: {
    createMecabTokenizerAndCheck: () => Promise<void>;
  };
  actions: {
    requestAppQuit: () => void;
  };
  logger: StatsRuntimeCoordinatorInput['logger'];
}

export function createStatsRuntimeFromMainState(
  input: StatsRuntimeFromMainStateInput,
): StatsRuntimeCoordinator {
  return createStatsRuntimeCoordinator({
    statsDaemonStatePath: path.join(input.userDataPath, 'stats-daemon.json'),
    statsDistPath: path.join(input.dirname, '..', 'stats', 'dist'),
    statsPreloadPath: path.join(input.dirname, 'preload-stats.js'),
    userDataPath: input.userDataPath,
    appState: input.appState,
    getResolvedConfig: () => input.getResolvedConfig(),
    dictionarySupport: input.dictionarySupport,
    overlay: {
      getOverlayGeometry: () => input.overlay.getOverlayGeometry(),
      updateVisibleOverlayVisibility: () => input.overlay.updateVisibleOverlayVisibility(),
      registerStatsOverlayToggle,
    },
    mpvRuntime: {
      createMecabTokenizerAndCheck: () => input.mpvRuntime.createMecabTokenizerAndCheck(),
    },
    actions: {
      openExternal: (url) => shell.openExternal(url),
      requestAppQuit: () => input.actions.requestAppQuit(),
      destroyStatsWindow,
    },
    logger: input.logger,
  });
}
