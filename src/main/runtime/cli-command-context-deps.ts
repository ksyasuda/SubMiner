import type { CliArgs } from '../../cli/args';
import type { CliCommandContextFactoryDeps } from './cli-command-context';

export function createBuildCliCommandContextDepsHandler(deps: {
  setLogLevel?: (level: NonNullable<CliArgs['logLevel']>) => void;
  getSocketPath: () => string;
  setSocketPath: (socketPath: string) => void;
  getMpvClient: CliCommandContextFactoryDeps['getMpvClient'];
  showOsd: (text: string) => void;
  showPlaybackFeedback?: (text: string) => void;
  texthookerService: CliCommandContextFactoryDeps['texthookerService'];
  getTexthookerPort: () => number;
  setTexthookerPort: (port: number) => void;
  getTexthookerWebsocketUrl: () => string | undefined;
  shouldOpenBrowser: () => boolean;
  openExternal: (url: string) => Promise<unknown>;
  logBrowserOpenError: (url: string, error: unknown) => void;
  isOverlayInitialized: () => boolean;
  initializeOverlay: () => void;
  toggleVisibleOverlay: () => void;
  togglePrimarySubtitleBar: () => void;
  openFirstRunSetup: (force?: boolean) => void;
  setVisibleOverlay: (visible: boolean) => void;
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
  clearAnilistToken: CliCommandContextFactoryDeps['clearAnilistToken'];
  openAnilistSetup: CliCommandContextFactoryDeps['openAnilistSetup'];
  openJellyfinSetup: CliCommandContextFactoryDeps['openJellyfinSetup'];
  getAnilistQueueStatus: CliCommandContextFactoryDeps['getAnilistQueueStatus'];
  retryAnilistQueueNow: CliCommandContextFactoryDeps['retryAnilistQueueNow'];
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
    getSocketPath: deps.getSocketPath,
    setSocketPath: deps.setSocketPath,
    getMpvClient: deps.getMpvClient,
    showOsd: deps.showOsd,
    showPlaybackFeedback: deps.showPlaybackFeedback,
    texthookerService: deps.texthookerService,
    getTexthookerPort: deps.getTexthookerPort,
    setTexthookerPort: deps.setTexthookerPort,
    getTexthookerWebsocketUrl: deps.getTexthookerWebsocketUrl,
    shouldOpenBrowser: deps.shouldOpenBrowser,
    openExternal: deps.openExternal,
    logBrowserOpenError: deps.logBrowserOpenError,
    isOverlayInitialized: deps.isOverlayInitialized,
    initializeOverlay: deps.initializeOverlay,
    toggleVisibleOverlay: deps.toggleVisibleOverlay,
    togglePrimarySubtitleBar: deps.togglePrimarySubtitleBar,
    openFirstRunSetup: deps.openFirstRunSetup,
    setVisibleOverlay: deps.setVisibleOverlay,
    copyCurrentSubtitle: deps.copyCurrentSubtitle,
    startPendingMultiCopy: deps.startPendingMultiCopy,
    mineSentenceCard: deps.mineSentenceCard,
    startPendingMineSentenceMultiple: deps.startPendingMineSentenceMultiple,
    updateLastCardFromClipboard: deps.updateLastCardFromClipboard,
    refreshKnownWordCache: deps.refreshKnownWordCache,
    triggerFieldGrouping: deps.triggerFieldGrouping,
    triggerSubsyncFromConfig: deps.triggerSubsyncFromConfig,
    markLastCardAsAudioCard: deps.markLastCardAsAudioCard,
    dispatchSessionAction: deps.dispatchSessionAction,
    getAnilistStatus: deps.getAnilistStatus,
    clearAnilistToken: deps.clearAnilistToken,
    openAnilistSetup: deps.openAnilistSetup,
    openJellyfinSetup: deps.openJellyfinSetup,
    getAnilistQueueStatus: deps.getAnilistQueueStatus,
    retryAnilistQueueNow: deps.retryAnilistQueueNow,
    generateCharacterDictionary: deps.generateCharacterDictionary,
    getCharacterDictionarySelection: deps.getCharacterDictionarySelection,
    setCharacterDictionarySelection: deps.setCharacterDictionarySelection,
    runStatsCommand: deps.runStatsCommand,
    runJellyfinCommand: deps.runJellyfinCommand,
    runUpdateCommand: deps.runUpdateCommand,
    runYoutubePlaybackFlow: deps.runYoutubePlaybackFlow,
    openYomitanSettings: deps.openYomitanSettings,
    openConfigSettingsWindow: deps.openConfigSettingsWindow,
    cycleSecondarySubMode: deps.cycleSecondarySubMode,
    openRuntimeOptionsPalette: deps.openRuntimeOptionsPalette,
    printHelp: deps.printHelp,
    stopApp: deps.stopApp,
    hasMainWindow: deps.hasMainWindow,
    getMultiCopyTimeoutMs: deps.getMultiCopyTimeoutMs,
    schedule: deps.schedule,
    logInfo: deps.logInfo,
    logDebug: deps.logDebug,
    logWarn: deps.logWarn,
    logError: deps.logError,
  });
}
