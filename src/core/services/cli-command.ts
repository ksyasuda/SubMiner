import {
  CliArgs,
  CliCommandSource,
  commandNeedsOverlayRuntime,
} from "../../cli/args";

export interface CliCommandServiceDeps {
  getMpvSocketPath: () => string;
  setMpvSocketPath: (socketPath: string) => void;
  setMpvClientSocketPath: (socketPath: string) => void;
  hasMpvClient: () => boolean;
  connectMpvClient: () => void;
  isTexthookerRunning: () => boolean;
  setTexthookerPort: (port: number) => void;
  getTexthookerPort: () => number;
  shouldOpenTexthookerBrowser: () => boolean;
  ensureTexthookerRunning: (port: number) => void;
  openTexthookerInBrowser: (url: string) => void;
  stopApp: () => void;
  isOverlayRuntimeInitialized: () => boolean;
  initializeOverlayRuntime: () => void;
  toggleVisibleOverlay: () => void;
  toggleInvisibleOverlay: () => void;
  openYomitanSettingsDelayed: (delayMs: number) => void;
  setVisibleOverlayVisible: (visible: boolean) => void;
  setInvisibleOverlayVisible: (visible: boolean) => void;
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
  getAnilistStatus: () => {
    tokenStatus: "not_checked" | "resolved" | "error";
    tokenSource: "none" | "literal" | "stored";
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
  getAnilistQueueStatus: () => {
    pending: number;
    ready: number;
    deadLetter: number;
    lastAttemptAt: number | null;
    lastError: string | null;
  };
  retryAnilistQueue: () => Promise<{ ok: boolean; message: string }>;
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
  start: (port: number) => void;
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
  shouldOpenBrowser: () => boolean;
  openInBrowser: (url: string) => void;
}

interface OverlayCliRuntime {
  isInitialized: () => boolean;
  initialize: () => void;
  toggleVisible: () => void;
  toggleInvisible: () => void;
  setVisible: (visible: boolean) => void;
  setInvisible: (visible: boolean) => void;
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
  openYomitanSettings: () => void;
  cycleSecondarySubMode: () => void;
  openRuntimeOptionsPalette: () => void;
  printHelp: () => void;
}

interface AnilistCliRuntime {
  getStatus: CliCommandServiceDeps["getAnilistStatus"];
  clearToken: CliCommandServiceDeps["clearAnilistToken"];
  openSetup: CliCommandServiceDeps["openAnilistSetup"];
  getQueueStatus: CliCommandServiceDeps["getAnilistQueueStatus"];
  retryQueueNow: CliCommandServiceDeps["retryAnilistQueue"];
}

interface AppCliRuntime {
  stop: () => void;
  hasMainWindow: () => boolean;
}

export interface CliCommandDepsRuntimeOptions {
  mpv: MpvCliRuntime;
  texthooker: TexthookerCliRuntime;
  overlay: OverlayCliRuntime;
  mining: MiningCliRuntime;
  anilist: AnilistCliRuntime;
  ui: UiCliRuntime;
  app: AppCliRuntime;
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
    shouldOpenTexthookerBrowser: options.texthooker.shouldOpenBrowser,
    ensureTexthookerRunning: (port) => {
      if (!options.texthooker.service.isRunning()) {
        options.texthooker.service.start(port);
      }
    },
    openTexthookerInBrowser: options.texthooker.openInBrowser,
    stopApp: options.app.stop,
    isOverlayRuntimeInitialized: options.overlay.isInitialized,
    initializeOverlayRuntime: options.overlay.initialize,
    toggleVisibleOverlay: options.overlay.toggleVisible,
    toggleInvisibleOverlay: options.overlay.toggleInvisible,
    openYomitanSettingsDelayed: (delayMs) => {
      options.schedule(() => {
        options.ui.openYomitanSettings();
      }, delayMs);
    },
    setVisibleOverlayVisible: options.overlay.setVisible,
    setInvisibleOverlayVisible: options.overlay.setInvisible,
    copyCurrentSubtitle: options.mining.copyCurrentSubtitle,
    startPendingMultiCopy: options.mining.startPendingMultiCopy,
    mineSentenceCard: options.mining.mineSentenceCard,
    startPendingMineSentenceMultiple:
      options.mining.startPendingMineSentenceMultiple,
    updateLastCardFromClipboard: options.mining.updateLastCardFromClipboard,
    refreshKnownWords: options.mining.refreshKnownWords,
    cycleSecondarySubMode: options.ui.cycleSecondarySubMode,
    triggerFieldGrouping: options.mining.triggerFieldGrouping,
    triggerSubsyncFromConfig: options.mining.triggerSubsyncFromConfig,
    markLastCardAsAudioCard: options.mining.markLastCardAsAudioCard,
    openRuntimeOptionsPalette: options.ui.openRuntimeOptionsPalette,
    getAnilistStatus: options.anilist.getStatus,
    clearAnilistToken: options.anilist.clearToken,
    openAnilistSetup: options.anilist.openSetup,
    getAnilistQueueStatus: options.anilist.getQueueStatus,
    retryAnilistQueue: options.anilist.retryQueueNow,
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
  if (!value) return "never";
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
  source: CliCommandSource = "initial",
  deps: CliCommandServiceDeps,
): void {
  const hasNonStartAction =
    args.stop ||
    args.toggle ||
    args.toggleVisibleOverlay ||
    args.toggleInvisibleOverlay ||
    args.settings ||
    args.show ||
    args.hide ||
    args.showVisibleOverlay ||
    args.hideVisibleOverlay ||
    args.showInvisibleOverlay ||
    args.hideInvisibleOverlay ||
    args.copySubtitle ||
    args.copySubtitleMultiple ||
    args.mineSentence ||
    args.mineSentenceMultiple ||
    args.updateLastCardFromClipboard ||
    args.refreshKnownWords ||
    args.toggleSecondarySub ||
    args.triggerFieldGrouping ||
    args.triggerSubsync ||
    args.markAudioCard ||
    args.openRuntimeOptions ||
    args.anilistStatus ||
    args.anilistLogout ||
    args.anilistSetup ||
    args.anilistRetryQueue ||
    args.texthooker ||
    args.help;
  const ignoreStartOnly = source === "second-instance" && args.start && !hasNonStartAction;
  if (ignoreStartOnly) {
    deps.log("Ignoring --start because SubMiner is already running.");
    return;
  }

  const shouldStart =
    args.start ||
    (source === "initial" &&
      (args.toggle ||
        args.toggleVisibleOverlay ||
        args.toggleInvisibleOverlay));
  const needsOverlayRuntime = commandNeedsOverlayRuntime(args);

  if (args.socketPath !== undefined) {
    deps.setMpvSocketPath(args.socketPath);
    deps.setMpvClientSocketPath(args.socketPath);
  }

  if (args.texthookerPort !== undefined) {
    if (deps.isTexthookerRunning()) {
      deps.warn(
        "Ignoring --port override because the texthooker server is already running.",
      );
    } else {
      deps.setTexthookerPort(args.texthookerPort);
    }
  }

  if (args.stop) {
    deps.log("Stopping SubMiner...");
    deps.stopApp();
    return;
  }

  if (needsOverlayRuntime && !deps.isOverlayRuntimeInitialized()) {
    deps.initializeOverlayRuntime();
  }

  if (shouldStart && deps.hasMpvClient()) {
    const socketPath = deps.getMpvSocketPath();
    deps.setMpvClientSocketPath(socketPath);
    deps.connectMpvClient();
    deps.log(`Starting MPV IPC connection on socket: ${socketPath}`);
  }

  if (args.toggle || args.toggleVisibleOverlay) {
    deps.toggleVisibleOverlay();
  } else if (args.toggleInvisibleOverlay) {
    deps.toggleInvisibleOverlay();
  } else if (args.settings) {
    deps.openYomitanSettingsDelayed(1000);
  } else if (args.show || args.showVisibleOverlay) {
    deps.setVisibleOverlayVisible(true);
  } else if (args.hide || args.hideVisibleOverlay) {
    deps.setVisibleOverlayVisible(false);
  } else if (args.showInvisibleOverlay) {
    deps.setInvisibleOverlayVisible(true);
  } else if (args.hideInvisibleOverlay) {
    deps.setInvisibleOverlayVisible(false);
  } else if (args.copySubtitle) {
    deps.copyCurrentSubtitle();
  } else if (args.copySubtitleMultiple) {
    deps.startPendingMultiCopy(deps.getMultiCopyTimeoutMs());
  } else if (args.mineSentence) {
    runAsyncWithOsd(
      () => deps.mineSentenceCard(),
      deps,
      "mineSentenceCard",
      "Mine sentence failed",
    );
  } else if (args.mineSentenceMultiple) {
    deps.startPendingMineSentenceMultiple(deps.getMultiCopyTimeoutMs());
  } else if (args.updateLastCardFromClipboard) {
    runAsyncWithOsd(
      () => deps.updateLastCardFromClipboard(),
      deps,
      "updateLastCardFromClipboard",
      "Update failed",
    );
  } else if (args.refreshKnownWords) {
    runAsyncWithOsd(
      () => deps.refreshKnownWords(),
      deps,
      "refreshKnownWords",
      "Refresh known words failed",
    );
  } else if (args.toggleSecondarySub) {
    deps.cycleSecondarySubMode();
  } else if (args.triggerFieldGrouping) {
    runAsyncWithOsd(
      () => deps.triggerFieldGrouping(),
      deps,
      "triggerFieldGrouping",
      "Field grouping failed",
    );
  } else if (args.triggerSubsync) {
    runAsyncWithOsd(
      () => deps.triggerSubsyncFromConfig(),
      deps,
      "triggerSubsyncFromConfig",
      "Subsync failed",
    );
  } else if (args.markAudioCard) {
    runAsyncWithOsd(
      () => deps.markLastCardAsAudioCard(),
      deps,
      "markLastCardAsAudioCard",
      "Audio card failed",
    );
  } else if (args.openRuntimeOptions) {
    deps.openRuntimeOptionsPalette();
  } else if (args.anilistStatus) {
    const status = deps.getAnilistStatus();
    deps.log(
      `AniList token status: ${status.tokenStatus} (source=${status.tokenSource})`,
    );
    if (status.tokenMessage) {
      deps.log(`AniList token message: ${status.tokenMessage}`);
    }
    deps.log(
      `AniList token timestamps: resolved=${formatTimestamp(status.tokenResolvedAt)}, error=${formatTimestamp(status.tokenErrorAt)}`,
    );
    deps.log(
      `AniList queue: pending=${status.queuePending}, ready=${status.queueReady}, deadLetter=${status.queueDeadLetter}`,
    );
    deps.log(
      `AniList queue timestamps: lastAttempt=${formatTimestamp(status.queueLastAttemptAt)}`,
    );
    if (status.queueLastError) {
      deps.warn(`AniList queue last error: ${status.queueLastError}`);
    }
  } else if (args.anilistLogout) {
    deps.clearAnilistToken();
    deps.log("Cleared stored AniList token.");
  } else if (args.anilistSetup) {
    deps.openAnilistSetup();
    deps.log("Opened AniList setup flow.");
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
      "retryAnilistQueue",
      "AniList retry failed",
    );
  } else if (args.texthooker) {
    const texthookerPort = deps.getTexthookerPort();
    deps.ensureTexthookerRunning(texthookerPort);
    if (deps.shouldOpenTexthookerBrowser()) {
      deps.openTexthookerInBrowser(`http://127.0.0.1:${texthookerPort}`);
    }
    deps.log(`Texthooker available at http://127.0.0.1:${texthookerPort}`);
  } else if (args.help) {
    deps.printHelp();
    if (!deps.hasMainWindow()) deps.stopApp();
  }
}
