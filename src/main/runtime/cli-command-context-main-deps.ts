import type { CliArgs, CliCommandSource } from '../../cli/args';
import { resolveTexthookerWebsocketUrl } from '../../core/services/startup';
import type { CliCommandContextFactoryDeps } from './cli-command-context';

type CliCommandContextMainState = {
  mpvSocketPath: string;
  mpvClient: ReturnType<CliCommandContextFactoryDeps['getMpvClient']>;
  texthookerPort: number;
  overlayRuntimeInitialized: boolean;
};

export function createBuildCliCommandContextMainDepsHandler(deps: {
  appState: CliCommandContextMainState;
  setLogLevel?: (level: NonNullable<CliArgs['logLevel']>) => void;
  texthookerService: CliCommandContextFactoryDeps['texthookerService'];
  getResolvedConfig: () => {
    texthooker?: { openBrowser?: boolean };
    websocket?: { enabled?: boolean | 'auto'; port?: number };
    annotationWebsocket?: { enabled?: boolean; port?: number };
  };
  defaultWebsocketPort: number;
  defaultAnnotationWebsocketPort: number;
  hasMpvWebsocketPlugin: () => boolean;
  openExternal: (url: string) => Promise<unknown>;
  logBrowserOpenError: (url: string, error: unknown) => void;
  showMpvOsd: (text: string) => void;

  initializeOverlayRuntime: () => void;
  toggleVisibleOverlay: () => void;
  togglePrimarySubtitleBar: () => void;
  openFirstRunSetupWindow: (force?: boolean) => void;
  setVisibleOverlayVisible: (visible: boolean) => void;

  copyCurrentSubtitle: () => void;
  startPendingMultiCopy: (timeoutMs: number) => void;
  mineSentenceCard: () => Promise<void>;
  startPendingMineSentenceMultiple: (timeoutMs: number) => void;
  updateLastCardFromClipboard: () => Promise<void>;
  refreshKnownWordCache: () => Promise<void>;
  triggerFieldGrouping: () => Promise<void>;
  triggerSubsyncFromConfig: () => Promise<void>;
  markLastCardAsAudioCard: () => Promise<void>;
  dispatchSessionAction: CliCommandContextFactoryDeps['dispatchSessionAction'];

  getAnilistStatus: CliCommandContextFactoryDeps['getAnilistStatus'];
  clearAnilistToken: () => void;
  openAnilistSetupWindow: () => void;
  openJellyfinSetupWindow: () => void;
  getAnilistQueueStatus: CliCommandContextFactoryDeps['getAnilistQueueStatus'];
  processNextAnilistRetryUpdate: CliCommandContextFactoryDeps['retryAnilistQueueNow'];
  generateCharacterDictionary: CliCommandContextFactoryDeps['generateCharacterDictionary'];
  getCharacterDictionarySelection?: CliCommandContextFactoryDeps['getCharacterDictionarySelection'];
  setCharacterDictionarySelection?: CliCommandContextFactoryDeps['setCharacterDictionarySelection'];
  runStatsCommand: CliCommandContextFactoryDeps['runStatsCommand'];
  runJellyfinCommand: (args: CliArgs) => Promise<void>;
  runUpdateCommand: CliCommandContextFactoryDeps['runUpdateCommand'];
  runYoutubePlaybackFlow: CliCommandContextFactoryDeps['runYoutubePlaybackFlow'];

  openYomitanSettings: () => void;
  openConfigSettingsWindow: () => void;
  cycleSecondarySubMode: () => void;
  openRuntimeOptionsPalette: () => void;
  printHelp: () => void;
  stopApp: () => void;
  hasMainWindow: () => boolean;
  getMultiCopyTimeoutMs: () => number;
  schedule: (fn: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  logInfo: (message: string) => void;
  logDebug: (message: string) => void;
  logWarn: (message: string) => void;
  logError: (message: string, err: unknown) => void;
}) {
  return (): CliCommandContextFactoryDeps => ({
    setLogLevel: deps.setLogLevel,
    getSocketPath: () => deps.appState.mpvSocketPath,
    setSocketPath: (socketPath: string) => {
      deps.appState.mpvSocketPath = socketPath;
    },
    getMpvClient: () => deps.appState.mpvClient,
    showOsd: (text: string) => deps.showMpvOsd(text),
    texthookerService: deps.texthookerService,
    getTexthookerPort: () => deps.appState.texthookerPort,
    setTexthookerPort: (port: number) => {
      deps.appState.texthookerPort = port;
    },
    getTexthookerWebsocketUrl: () =>
      resolveTexthookerWebsocketUrl(
        deps.getResolvedConfig(),
        {
          defaultWebsocketPort: deps.defaultWebsocketPort,
          defaultAnnotationWebsocketPort: deps.defaultAnnotationWebsocketPort,
        },
        deps.hasMpvWebsocketPlugin(),
      ),
    shouldOpenBrowser: () => deps.getResolvedConfig().texthooker?.openBrowser !== false,
    openExternal: (url: string) => deps.openExternal(url),
    logBrowserOpenError: (url: string, error: unknown) => deps.logBrowserOpenError(url, error),
    isOverlayInitialized: () => deps.appState.overlayRuntimeInitialized,
    initializeOverlay: () => deps.initializeOverlayRuntime(),
    toggleVisibleOverlay: () => deps.toggleVisibleOverlay(),
    togglePrimarySubtitleBar: () => deps.togglePrimarySubtitleBar(),
    openFirstRunSetup: (force?: boolean) => deps.openFirstRunSetupWindow(force),
    setVisibleOverlay: (visible: boolean) => deps.setVisibleOverlayVisible(visible),
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
    dispatchSessionAction: (request) => deps.dispatchSessionAction(request),
    getAnilistStatus: () => deps.getAnilistStatus(),
    clearAnilistToken: () => deps.clearAnilistToken(),
    openAnilistSetup: () => deps.openAnilistSetupWindow(),
    openJellyfinSetup: () => deps.openJellyfinSetupWindow(),
    getAnilistQueueStatus: () => deps.getAnilistQueueStatus(),
    retryAnilistQueueNow: () => deps.processNextAnilistRetryUpdate(),
    generateCharacterDictionary: (targetPath?: string) =>
      deps.generateCharacterDictionary(targetPath),
    getCharacterDictionarySelection: deps.getCharacterDictionarySelection,
    setCharacterDictionarySelection: deps.setCharacterDictionarySelection,
    runStatsCommand: (args: CliArgs, source) => deps.runStatsCommand(args, source),
    runJellyfinCommand: (args: CliArgs) => deps.runJellyfinCommand(args),
    runUpdateCommand: (args: CliArgs, source: CliCommandSource) =>
      deps.runUpdateCommand(args, source),
    runYoutubePlaybackFlow: (request) => deps.runYoutubePlaybackFlow(request),
    openYomitanSettings: () => deps.openYomitanSettings(),
    openConfigSettingsWindow: () => deps.openConfigSettingsWindow(),
    cycleSecondarySubMode: () => deps.cycleSecondarySubMode(),
    openRuntimeOptionsPalette: () => deps.openRuntimeOptionsPalette(),
    printHelp: () => deps.printHelp(),
    stopApp: () => deps.stopApp(),
    hasMainWindow: () => deps.hasMainWindow(),
    getMultiCopyTimeoutMs: () => deps.getMultiCopyTimeoutMs(),
    schedule: (fn: () => void, delayMs: number) => deps.schedule(fn, delayMs),
    logInfo: (message: string) => deps.logInfo(message),
    logDebug: (message: string) => deps.logDebug(message),
    logWarn: (message: string) => deps.logWarn(message),
    logError: (message: string, err: unknown) => deps.logError(message, err),
  });
}
