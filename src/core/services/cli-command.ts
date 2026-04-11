import { CliArgs, CliCommandSource, commandNeedsOverlayRuntime } from '../../cli/args';
import type { SessionActionDispatchRequest } from '../../types/runtime';

export interface CliCommandServiceDeps {
  setLogLevel?: (level: NonNullable<CliArgs['logLevel']>) => void;
  getMpvSocketPath: () => string;
  setMpvSocketPath: (socketPath: string) => void;
  setMpvClientSocketPath: (socketPath: string) => void;
  hasMpvClient: () => boolean;
  connectMpvClient: () => void;
  isTexthookerRunning: () => boolean;
  setTexthookerPort: (port: number) => void;
  getTexthookerPort: () => number;
  getTexthookerWebsocketUrl: () => string | undefined;
  shouldOpenTexthookerBrowser: () => boolean;
  ensureTexthookerRunning: (port: number, websocketUrl?: string) => void;
  openTexthookerInBrowser: (url: string) => void;
  stopApp: () => void;
  isOverlayRuntimeInitialized: () => boolean;
  initializeOverlayRuntime: () => void;
  toggleVisibleOverlay: () => void;
  openFirstRunSetup: () => void;
  openYomitanSettingsDelayed: (delayMs: number) => void;
  setVisibleOverlayVisible: (visible: boolean) => void;
  copyCurrentSubtitle: () => void;
  startPendingMultiCopy: (timeoutMs: number) => void;
  mineSentenceCard: () => Promise<void>;
  startPendingMineSentenceMultiple: (timeoutMs: number) => void;
  updateLastCardFromClipboard: () => Promise<void>;
  refreshKnownWords: () => Promise<void>;
  cycleSecondarySubMode: () => void;
  triggerFieldGrouping: () => Promise<void>;
  triggerSubsyncFromConfig: () => Promise<void>;
  markLastCardAsAudioCard: () => Promise<void>;
  openRuntimeOptionsPalette: () => void;
  dispatchSessionAction: (request: SessionActionDispatchRequest) => Promise<void>;
  getAnilistStatus: () => {
    tokenStatus: 'not_checked' | 'resolved' | 'error';
    tokenSource: 'none' | 'literal' | 'stored';
    tokenMessage: string | null;
    tokenResolvedAt: number | null;
    tokenErrorAt: number | null;
    queuePending: number;
    queueReady: number;
    queueDeadLetter: number;
    queueLastAttemptAt: number | null;
    queueLastError: string | null;
  };
  clearAnilistToken: () => void;
  openAnilistSetup: () => void;
  openJellyfinSetup: () => void;
  getAnilistQueueStatus: () => {
    pending: number;
    ready: number;
    deadLetter: number;
    lastAttemptAt: number | null;
    lastError: string | null;
  };
  retryAnilistQueue: () => Promise<{ ok: boolean; message: string }>;
  generateCharacterDictionary: (targetPath?: string) => Promise<{
    zipPath: string;
    fromCache: boolean;
    mediaId: number;
    mediaTitle: string;
    entryCount: number;
  }>;
  runStatsCommand: (args: CliArgs, source: CliCommandSource) => Promise<void>;
  runJellyfinCommand: (args: CliArgs) => Promise<void>;
  runYoutubePlaybackFlow: (request: {
    url: string;
    mode: NonNullable<CliArgs['youtubeMode']>;
    source: CliCommandSource;
  }) => Promise<void>;
  printHelp: () => void;
  hasMainWindow: () => boolean;
  getMultiCopyTimeoutMs: () => number;
  showMpvOsd: (text: string) => void;
  log: (message: string) => void;
  warn: (message: string) => void;
  error: (message: string, err: unknown) => void;
}

interface MpvClientLike {
  setSocketPath: (socketPath: string) => void;
  connect: () => void;
}

interface TexthookerServiceLike {
  isRunning: () => boolean;
  start: (port: number, websocketUrl?: string) => void;
}

interface MpvCliRuntime {
  getSocketPath: () => string;
  setSocketPath: (socketPath: string) => void;
  getClient: () => MpvClientLike | null;
  showOsd: (text: string) => void;
}

interface TexthookerCliRuntime {
  service: TexthookerServiceLike;
  getPort: () => number;
  setPort: (port: number) => void;
  getWebsocketUrl: () => string | undefined;
  shouldOpenBrowser: () => boolean;
  openInBrowser: (url: string) => void;
}

interface OverlayCliRuntime {
  isInitialized: () => boolean;
  initialize: () => void;
  toggleVisible: () => void;
  setVisible: (visible: boolean) => void;
}

interface MiningCliRuntime {
  copyCurrentSubtitle: () => void;
  startPendingMultiCopy: (timeoutMs: number) => void;
  mineSentenceCard: () => Promise<void>;
  startPendingMineSentenceMultiple: (timeoutMs: number) => void;
  updateLastCardFromClipboard: () => Promise<void>;
  refreshKnownWords: () => Promise<void>;
  triggerFieldGrouping: () => Promise<void>;
  triggerSubsyncFromConfig: () => Promise<void>;
  markLastCardAsAudioCard: () => Promise<void>;
}

interface UiCliRuntime {
  openFirstRunSetup: () => void;
  openYomitanSettings: () => void;
  cycleSecondarySubMode: () => void;
  openRuntimeOptionsPalette: () => void;
  printHelp: () => void;
}

interface AnilistCliRuntime {
  getStatus: CliCommandServiceDeps['getAnilistStatus'];
  clearToken: CliCommandServiceDeps['clearAnilistToken'];
  openSetup: CliCommandServiceDeps['openAnilistSetup'];
  getQueueStatus: CliCommandServiceDeps['getAnilistQueueStatus'];
  retryQueueNow: CliCommandServiceDeps['retryAnilistQueue'];
}

interface AppCliRuntime {
  stop: () => void;
  hasMainWindow: () => boolean;
  runYoutubePlaybackFlow: CliCommandServiceDeps['runYoutubePlaybackFlow'];
}

export interface CliCommandDepsRuntimeOptions {
  setLogLevel?: (level: NonNullable<CliArgs['logLevel']>) => void;
  mpv: MpvCliRuntime;
  texthooker: TexthookerCliRuntime;
  overlay: OverlayCliRuntime;
  mining: MiningCliRuntime;
  anilist: AnilistCliRuntime;
  dictionary: {
    generate: (targetPath?: string) => Promise<{
      zipPath: string;
      fromCache: boolean;
      mediaId: number;
      mediaTitle: string;
      entryCount: number;
    }>;
  };
  jellyfin: {
    openSetup: () => void;
    runStatsCommand: (args: CliArgs, source: CliCommandSource) => Promise<void>;
    runCommand: (args: CliArgs) => Promise<void>;
  };
  ui: UiCliRuntime;
  app: AppCliRuntime;
  dispatchSessionAction: (request: SessionActionDispatchRequest) => Promise<void>;
  getMultiCopyTimeoutMs: () => number;
  schedule: (fn: () => void, delayMs: number) => unknown;
  log: (message: string) => void;
  warn: (message: string) => void;
  error: (message: string, err: unknown) => void;
}

export function createCliCommandDepsRuntime(
  options: CliCommandDepsRuntimeOptions,
): CliCommandServiceDeps {
  return {
    setLogLevel: options.setLogLevel,
    getMpvSocketPath: options.mpv.getSocketPath,
    setMpvSocketPath: options.mpv.setSocketPath,
    setMpvClientSocketPath: (socketPath) => {
      const client = options.mpv.getClient();
      if (!client) return;
      client.setSocketPath(socketPath);
    },
    hasMpvClient: () => Boolean(options.mpv.getClient()),
    connectMpvClient: () => {
      const client = options.mpv.getClient();
      if (!client) return;
      client.connect();
    },
    isTexthookerRunning: () => options.texthooker.service.isRunning(),
    setTexthookerPort: options.texthooker.setPort,
    getTexthookerPort: options.texthooker.getPort,
    getTexthookerWebsocketUrl: options.texthooker.getWebsocketUrl,
    shouldOpenTexthookerBrowser: options.texthooker.shouldOpenBrowser,
    ensureTexthookerRunning: (port, websocketUrl) => {
      if (!options.texthooker.service.isRunning()) {
        options.texthooker.service.start(port, websocketUrl);
      }
    },
    openTexthookerInBrowser: options.texthooker.openInBrowser,
    stopApp: options.app.stop,
    isOverlayRuntimeInitialized: options.overlay.isInitialized,
    initializeOverlayRuntime: options.overlay.initialize,
    toggleVisibleOverlay: options.overlay.toggleVisible,
    openFirstRunSetup: options.ui.openFirstRunSetup,
    openYomitanSettingsDelayed: (delayMs) => {
      options.schedule(() => {
        options.ui.openYomitanSettings();
      }, delayMs);
    },
    setVisibleOverlayVisible: options.overlay.setVisible,
    copyCurrentSubtitle: options.mining.copyCurrentSubtitle,
    startPendingMultiCopy: options.mining.startPendingMultiCopy,
    mineSentenceCard: options.mining.mineSentenceCard,
    startPendingMineSentenceMultiple: options.mining.startPendingMineSentenceMultiple,
    updateLastCardFromClipboard: options.mining.updateLastCardFromClipboard,
    refreshKnownWords: options.mining.refreshKnownWords,
    cycleSecondarySubMode: options.ui.cycleSecondarySubMode,
    triggerFieldGrouping: options.mining.triggerFieldGrouping,
    triggerSubsyncFromConfig: options.mining.triggerSubsyncFromConfig,
    markLastCardAsAudioCard: options.mining.markLastCardAsAudioCard,
    openRuntimeOptionsPalette: options.ui.openRuntimeOptionsPalette,
    dispatchSessionAction: options.dispatchSessionAction,
    getAnilistStatus: options.anilist.getStatus,
    clearAnilistToken: options.anilist.clearToken,
    openAnilistSetup: options.anilist.openSetup,
    openJellyfinSetup: options.jellyfin.openSetup,
    getAnilistQueueStatus: options.anilist.getQueueStatus,
    retryAnilistQueue: options.anilist.retryQueueNow,
    generateCharacterDictionary: options.dictionary.generate,
    runStatsCommand: options.jellyfin.runStatsCommand,
    runJellyfinCommand: options.jellyfin.runCommand,
    runYoutubePlaybackFlow: options.app.runYoutubePlaybackFlow,
    printHelp: options.ui.printHelp,
    hasMainWindow: options.app.hasMainWindow,
    getMultiCopyTimeoutMs: options.getMultiCopyTimeoutMs,
    showMpvOsd: options.mpv.showOsd,
    log: options.log,
    warn: options.warn,
    error: options.error,
  };
}

function formatTimestamp(value: number | null): string {
  if (!value) return 'never';
  return new Date(value).toISOString();
}

function runAsyncWithOsd(
  task: () => Promise<void>,
  deps: CliCommandServiceDeps,
  logLabel: string,
  osdLabel: string,
): void {
  task().catch((err) => {
    deps.error(`${logLabel} failed:`, err);
    deps.showMpvOsd(`${osdLabel}: ${(err as Error).message}`);
  });
}

export function handleCliCommand(
  args: CliArgs,
  source: CliCommandSource = 'initial',
  deps: CliCommandServiceDeps,
): void {
  const dispatchCliSessionAction = (
    request: SessionActionDispatchRequest,
    logLabel: string,
    osdLabel: string,
  ): void => {
    runAsyncWithOsd(
      () => deps.dispatchSessionAction?.(request) ?? Promise.resolve(),
      deps,
      logLabel,
      osdLabel,
    );
  };

  if (args.logLevel) {
    deps.setLogLevel?.(args.logLevel);
  }

  const reuseSecondInstanceStart =
    source === 'second-instance' && args.start && deps.isOverlayRuntimeInitialized();
  const shouldConnectMpv = args.start;
  const needsOverlayRuntime = commandNeedsOverlayRuntime(args);
  const shouldInitializeOverlayRuntime = needsOverlayRuntime || args.start;

  if (args.socketPath !== undefined) {
    deps.setMpvSocketPath(args.socketPath);
    deps.setMpvClientSocketPath(args.socketPath);
  }

  if (args.texthookerPort !== undefined) {
    if (deps.isTexthookerRunning()) {
      deps.warn('Ignoring --port override because the texthooker server is already running.');
    } else {
      deps.setTexthookerPort(args.texthookerPort);
    }
  }

  if (args.stop) {
    deps.log('Stopping SubMiner...');
    deps.stopApp();
    return;
  }

  if (reuseSecondInstanceStart) {
    deps.log('Reusing running SubMiner instance for --start.');
  }

  if (shouldInitializeOverlayRuntime && !deps.isOverlayRuntimeInitialized()) {
    deps.initializeOverlayRuntime();
  }

  if (shouldConnectMpv && deps.hasMpvClient()) {
    const socketPath = deps.getMpvSocketPath();
    deps.setMpvClientSocketPath(socketPath);
    deps.connectMpvClient();
    deps.log(`Starting MPV IPC connection on socket: ${socketPath}`);
  }

  if (args.toggle || args.toggleVisibleOverlay) {
    deps.toggleVisibleOverlay();
  } else if (args.setup) {
    deps.openFirstRunSetup();
    deps.log('Opened first-run setup flow.');
  } else if (args.settings) {
    deps.openYomitanSettingsDelayed(1000);
  } else if (args.show || args.showVisibleOverlay) {
    deps.setVisibleOverlayVisible(true);
  } else if (args.hide || args.hideVisibleOverlay) {
    deps.setVisibleOverlayVisible(false);
  } else if (args.copySubtitle) {
    deps.copyCurrentSubtitle();
  } else if (args.copySubtitleMultiple) {
    deps.startPendingMultiCopy(deps.getMultiCopyTimeoutMs());
  } else if (args.mineSentence) {
    runAsyncWithOsd(
      () => deps.mineSentenceCard(),
      deps,
      'mineSentenceCard',
      'Mine sentence failed',
    );
  } else if (args.mineSentenceMultiple) {
    deps.startPendingMineSentenceMultiple(deps.getMultiCopyTimeoutMs());
  } else if (args.updateLastCardFromClipboard) {
    runAsyncWithOsd(
      () => deps.updateLastCardFromClipboard(),
      deps,
      'updateLastCardFromClipboard',
      'Update failed',
    );
  } else if (args.refreshKnownWords) {
    const shouldStopAfterRun = source === 'initial' && !deps.hasMainWindow();
    deps
      .refreshKnownWords()
      .catch((err) => {
        deps.error('refreshKnownWords failed:', err);
        deps.showMpvOsd(`Refresh known words failed: ${(err as Error).message}`);
      })
      .finally(() => {
        if (shouldStopAfterRun) {
          deps.stopApp();
        }
      });
  } else if (args.toggleSecondarySub) {
    deps.cycleSecondarySubMode();
  } else if (args.triggerFieldGrouping) {
    runAsyncWithOsd(
      () => deps.triggerFieldGrouping(),
      deps,
      'triggerFieldGrouping',
      'Field grouping failed',
    );
  } else if (args.triggerSubsync) {
    runAsyncWithOsd(
      () => deps.triggerSubsyncFromConfig(),
      deps,
      'triggerSubsyncFromConfig',
      'Subsync failed',
    );
  } else if (args.markAudioCard) {
    runAsyncWithOsd(
      () => deps.markLastCardAsAudioCard(),
      deps,
      'markLastCardAsAudioCard',
      'Audio card failed',
    );
  } else if (args.toggleStatsOverlay) {
    dispatchCliSessionAction(
      { actionId: 'toggleStatsOverlay' },
      'toggleStatsOverlay',
      'Stats toggle failed',
    );
  } else if (args.toggleSubtitleSidebar) {
    dispatchCliSessionAction(
      { actionId: 'toggleSubtitleSidebar' },
      'toggleSubtitleSidebar',
      'Subtitle sidebar toggle failed',
    );
  } else if (args.openRuntimeOptions) {
    deps.openRuntimeOptionsPalette();
  } else if (args.openSessionHelp) {
    dispatchCliSessionAction(
      { actionId: 'openSessionHelp' },
      'openSessionHelp',
      'Open session help failed',
    );
  } else if (args.openControllerSelect) {
    dispatchCliSessionAction(
      { actionId: 'openControllerSelect' },
      'openControllerSelect',
      'Open controller select failed',
    );
  } else if (args.openControllerDebug) {
    dispatchCliSessionAction(
      { actionId: 'openControllerDebug' },
      'openControllerDebug',
      'Open controller debug failed',
    );
  } else if (args.openJimaku) {
    dispatchCliSessionAction({ actionId: 'openJimaku' }, 'openJimaku', 'Open jimaku failed');
  } else if (args.openYoutubePicker) {
    dispatchCliSessionAction(
      { actionId: 'openYoutubePicker' },
      'openYoutubePicker',
      'Open YouTube picker failed',
    );
  } else if (args.openPlaylistBrowser) {
    dispatchCliSessionAction(
      { actionId: 'openPlaylistBrowser' },
      'openPlaylistBrowser',
      'Open playlist browser failed',
    );
  } else if (args.replayCurrentSubtitle) {
    dispatchCliSessionAction(
      { actionId: 'replayCurrentSubtitle' },
      'replayCurrentSubtitle',
      'Replay subtitle failed',
    );
  } else if (args.playNextSubtitle) {
    dispatchCliSessionAction(
      { actionId: 'playNextSubtitle' },
      'playNextSubtitle',
      'Play next subtitle failed',
    );
  } else if (args.shiftSubDelayPrevLine) {
    dispatchCliSessionAction(
      { actionId: 'shiftSubDelayPrevLine' },
      'shiftSubDelayPrevLine',
      'Shift subtitle delay failed',
    );
  } else if (args.shiftSubDelayNextLine) {
    dispatchCliSessionAction(
      { actionId: 'shiftSubDelayNextLine' },
      'shiftSubDelayNextLine',
      'Shift subtitle delay failed',
    );
  } else if (args.cycleRuntimeOptionId !== undefined) {
    dispatchCliSessionAction(
      {
        actionId: 'cycleRuntimeOption',
        payload: {
          runtimeOptionId: args.cycleRuntimeOptionId,
          direction: args.cycleRuntimeOptionDirection ?? 1,
        },
      },
      'cycleRuntimeOption',
      'Runtime option change failed',
    );
  } else if (args.copySubtitleCount !== undefined) {
    dispatchCliSessionAction(
      { actionId: 'copySubtitleMultiple', payload: { count: args.copySubtitleCount } },
      'copySubtitleMultiple',
      'Copy failed',
    );
  } else if (args.mineSentenceCount !== undefined) {
    dispatchCliSessionAction(
      { actionId: 'mineSentenceMultiple', payload: { count: args.mineSentenceCount } },
      'mineSentenceMultiple',
      'Mine sentence failed',
    );
  } else if (args.anilistStatus) {
    const status = deps.getAnilistStatus();
    deps.log(`AniList token status: ${status.tokenStatus} (source=${status.tokenSource})`);
    if (status.tokenMessage) {
      deps.log(`AniList token message: ${status.tokenMessage}`);
    }
    deps.log(
      `AniList token timestamps: resolved=${formatTimestamp(status.tokenResolvedAt)}, error=${formatTimestamp(status.tokenErrorAt)}`,
    );
    deps.log(
      `AniList queue: pending=${status.queuePending}, ready=${status.queueReady}, deadLetter=${status.queueDeadLetter}`,
    );
    deps.log(`AniList queue timestamps: lastAttempt=${formatTimestamp(status.queueLastAttemptAt)}`);
    if (status.queueLastError) {
      deps.warn(`AniList queue last error: ${status.queueLastError}`);
    }
  } else if (args.anilistLogout) {
    deps.clearAnilistToken();
    deps.log('Cleared stored AniList token.');
  } else if (args.anilistSetup) {
    deps.openAnilistSetup();
    deps.log('Opened AniList setup flow.');
  } else if (args.jellyfin) {
    deps.openJellyfinSetup();
    deps.log('Opened Jellyfin setup flow.');
  } else if (args.youtubePlay) {
    const youtubeUrl = args.youtubePlay;
    runAsyncWithOsd(
      () =>
        deps.runYoutubePlaybackFlow({
          url: youtubeUrl,
          mode: args.youtubeMode ?? 'download',
          source,
        }),
      deps,
      'runYoutubePlaybackFlow',
      'YouTube playback failed',
    );
  } else if (args.dictionary) {
    const shouldStopAfterRun = source === 'initial' && !deps.hasMainWindow();
    deps.log('Generating character dictionary for current anime...');
    deps
      .generateCharacterDictionary(args.dictionaryTarget)
      .then((result) => {
        const cacheLabel = result.fromCache ? 'cache hit' : 'generated';
        deps.log(
          `Character dictionary ${cacheLabel}: AniList ${result.mediaId} (${result.mediaTitle}), entries=${result.entryCount}`,
        );
        deps.log(`Dictionary ZIP: ${result.zipPath}`);
      })
      .catch((error) => {
        deps.error('generateCharacterDictionary failed:', error);
        deps.warn(
          `Dictionary generation failed: ${error instanceof Error ? error.message : String(error)}`,
        );
      })
      .finally(() => {
        if (shouldStopAfterRun) {
          deps.stopApp();
        }
      });
  } else if (args.stats) {
    void deps.runStatsCommand(args, source);
  } else if (args.anilistRetryQueue) {
    const queueStatus = deps.getAnilistQueueStatus();
    deps.log(
      `AniList queue before retry: pending=${queueStatus.pending}, ready=${queueStatus.ready}, deadLetter=${queueStatus.deadLetter}`,
    );
    runAsyncWithOsd(
      async () => {
        const result = await deps.retryAnilistQueue();
        if (result.ok) deps.log(result.message);
        else deps.warn(result.message);
      },
      deps,
      'retryAnilistQueue',
      'AniList retry failed',
    );
  } else if (
    args.jellyfinLogin ||
    args.jellyfinLogout ||
    args.jellyfinLibraries ||
    args.jellyfinItems ||
    args.jellyfinSubtitles ||
    args.jellyfinPlay ||
    args.jellyfinRemoteAnnounce
  ) {
    runAsyncWithOsd(
      () => deps.runJellyfinCommand(args),
      deps,
      'runJellyfinCommand',
      'Jellyfin command failed',
    );
  } else if (args.texthooker) {
    const texthookerPort = deps.getTexthookerPort();
    deps.ensureTexthookerRunning(texthookerPort, deps.getTexthookerWebsocketUrl());
    if (deps.shouldOpenTexthookerBrowser()) {
      deps.openTexthookerInBrowser(`http://127.0.0.1:${texthookerPort}`);
    }
    deps.log(`Texthooker available at http://127.0.0.1:${texthookerPort}`);
  } else if (args.help) {
    deps.printHelp();
    if (!deps.hasMainWindow()) deps.stopApp();
  }
}
