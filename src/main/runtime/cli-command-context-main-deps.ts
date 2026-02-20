import type { CliArgs } from '../../cli/args';

export function createBuildCliCommandContextMainDepsHandler(deps: {
  appState: {
    mpvSocketPath: string;
    mpvClient: unknown | null;
    texthookerPort: number;
    overlayRuntimeInitialized: boolean;
  };
  texthookerService: unknown;
  getResolvedConfig: () => { texthooker?: { openBrowser?: boolean } };
  openExternal: (url: string) => Promise<unknown>;
  logBrowserOpenError: (url: string, error: unknown) => void;
  showMpvOsd: (text: string) => void;

  initializeOverlayRuntime: () => void;
  toggleVisibleOverlay: () => void;
  toggleInvisibleOverlay: () => void;
  setVisibleOverlayVisible: (visible: boolean) => void;
  setInvisibleOverlayVisible: (visible: boolean) => void;

  copyCurrentSubtitle: () => void;
  startPendingMultiCopy: (timeoutMs: number) => void;
  mineSentenceCard: () => Promise<void>;
  startPendingMineSentenceMultiple: (timeoutMs: number) => void;
  updateLastCardFromClipboard: () => Promise<void>;
  refreshKnownWordCache: () => Promise<void>;
  triggerFieldGrouping: () => Promise<void>;
  triggerSubsyncFromConfig: () => Promise<void>;
  markLastCardAsAudioCard: () => Promise<void>;

  getAnilistStatus: () => unknown;
  clearAnilistToken: () => void;
  openAnilistSetupWindow: () => void;
  openJellyfinSetupWindow: () => void;
  getAnilistQueueStatus: () => unknown;
  processNextAnilistRetryUpdate: () => Promise<{ ok: boolean; message: string }>;
  runJellyfinCommand: (args: CliArgs) => Promise<void>;

  openYomitanSettings: () => void;
  cycleSecondarySubMode: () => void;
  openRuntimeOptionsPalette: () => void;
  printHelp: () => void;
  stopApp: () => void;
  hasMainWindow: () => boolean;
  getMultiCopyTimeoutMs: () => number;
  schedule: (fn: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  logInfo: (message: string) => void;
  logWarn: (message: string) => void;
  logError: (message: string, err: unknown) => void;
}) {
  return () => ({
    getSocketPath: () => deps.appState.mpvSocketPath,
    setSocketPath: (socketPath: string) => {
      deps.appState.mpvSocketPath = socketPath;
    },
    getMpvClient: () => deps.appState.mpvClient as never,
    showOsd: (text: string) => deps.showMpvOsd(text),
    texthookerService: deps.texthookerService as never,
    getTexthookerPort: () => deps.appState.texthookerPort,
    setTexthookerPort: (port: number) => {
      deps.appState.texthookerPort = port;
    },
    shouldOpenBrowser: () => deps.getResolvedConfig().texthooker?.openBrowser !== false,
    openExternal: (url: string) => deps.openExternal(url),
    logBrowserOpenError: (url: string, error: unknown) => deps.logBrowserOpenError(url, error),
    isOverlayInitialized: () => deps.appState.overlayRuntimeInitialized,
    initializeOverlay: () => deps.initializeOverlayRuntime(),
    toggleVisibleOverlay: () => deps.toggleVisibleOverlay(),
    toggleInvisibleOverlay: () => deps.toggleInvisibleOverlay(),
    setVisibleOverlay: (visible: boolean) => deps.setVisibleOverlayVisible(visible),
    setInvisibleOverlay: (visible: boolean) => deps.setInvisibleOverlayVisible(visible),
    copyCurrentSubtitle: () => deps.copyCurrentSubtitle(),
    startPendingMultiCopy: (timeoutMs: number) => deps.startPendingMultiCopy(timeoutMs),
    mineSentenceCard: () => deps.mineSentenceCard(),
    startPendingMineSentenceMultiple: (timeoutMs: number) =>
      deps.startPendingMineSentenceMultiple(timeoutMs),
    updateLastCardFromClipboard: () => deps.updateLastCardFromClipboard(),
    refreshKnownWordCache: () => deps.refreshKnownWordCache(),
    triggerFieldGrouping: () => deps.triggerFieldGrouping(),
    triggerSubsyncFromConfig: () => deps.triggerSubsyncFromConfig(),
    markLastCardAsAudioCard: () => deps.markLastCardAsAudioCard(),
    getAnilistStatus: () => deps.getAnilistStatus() as never,
    clearAnilistToken: () => deps.clearAnilistToken(),
    openAnilistSetup: () => deps.openAnilistSetupWindow(),
    openJellyfinSetup: () => deps.openJellyfinSetupWindow(),
    getAnilistQueueStatus: () => deps.getAnilistQueueStatus() as never,
    retryAnilistQueueNow: () => deps.processNextAnilistRetryUpdate(),
    runJellyfinCommand: (args: CliArgs) => deps.runJellyfinCommand(args),
    openYomitanSettings: () => deps.openYomitanSettings(),
    cycleSecondarySubMode: () => deps.cycleSecondarySubMode(),
    openRuntimeOptionsPalette: () => deps.openRuntimeOptionsPalette(),
    printHelp: () => deps.printHelp(),
    stopApp: () => deps.stopApp(),
    hasMainWindow: () => deps.hasMainWindow(),
    getMultiCopyTimeoutMs: () => deps.getMultiCopyTimeoutMs(),
    schedule: (fn: () => void, delayMs: number) => deps.schedule(fn, delayMs),
    logInfo: (message: string) => deps.logInfo(message),
    logWarn: (message: string) => deps.logWarn(message),
    logError: (message: string, err: unknown) => deps.logError(message, err),
  });
}
