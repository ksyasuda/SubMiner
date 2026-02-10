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
  cycleSecondarySubMode: () => void;
  triggerFieldGrouping: () => Promise<void>;
  triggerSubsyncFromConfig: () => Promise<void>;
  markLastCardAsAudioCard: () => Promise<void>;
  openRuntimeOptionsPalette: () => void;
  printHelp: () => void;
  hasMainWindow: () => boolean;
  getMultiCopyTimeoutMs: () => number;
  showMpvOsd: (text: string) => void;
  log: (message: string) => void;
  warn: (message: string) => void;
  error: (message: string, err: unknown) => void;
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

export function handleCliCommandService(
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
    args.toggleSecondarySub ||
    args.triggerFieldGrouping ||
    args.triggerSubsync ||
    args.markAudioCard ||
    args.openRuntimeOptions ||
    args.texthooker ||
    args.help;
  const ignoreStart = source === "second-instance" && args.start;
  if (ignoreStart && !hasNonStartAction) {
    deps.log("Ignoring --start because SubMiner is already running.");
    return;
  }

  const shouldStart =
    !ignoreStart &&
    (args.start ||
      (source === "initial" &&
        (args.toggle ||
          args.toggleVisibleOverlay ||
          args.toggleInvisibleOverlay)));
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
