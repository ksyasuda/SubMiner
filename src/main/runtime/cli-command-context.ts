import type { CliArgs } from '../../cli/args';
import type {
  CliCommandRuntimeServiceContext,
  CliCommandRuntimeServiceContextHandlers,
} from '../cli-runtime';

type MpvClientLike = CliCommandRuntimeServiceContext['getClient'] extends () => infer T ? T : never;

export type CliCommandContextFactoryDeps = {
  setLogLevel?: (level: NonNullable<CliArgs['logLevel']>) => void;
  getSocketPath: () => string;
  setSocketPath: (socketPath: string) => void;
  getMpvClient: () => MpvClientLike;
  showOsd: (text: string) => void;
  showPlaybackFeedback?: (text: string) => void;
  texthookerService: CliCommandRuntimeServiceContextHandlers['texthookerService'];
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
  dispatchSessionAction: CliCommandRuntimeServiceContext['dispatchSessionAction'];
  getAnilistStatus: CliCommandRuntimeServiceContext['getAnilistStatus'];
  clearAnilistToken: CliCommandRuntimeServiceContext['clearAnilistToken'];
  openAnilistSetup: CliCommandRuntimeServiceContext['openAnilistSetup'];
  openJellyfinSetup: CliCommandRuntimeServiceContext['openJellyfinSetup'];
  getAnilistQueueStatus: CliCommandRuntimeServiceContext['getAnilistQueueStatus'];
  retryAnilistQueueNow: CliCommandRuntimeServiceContext['retryAnilistQueueNow'];
  generateCharacterDictionary: CliCommandRuntimeServiceContext['generateCharacterDictionary'];
  getCharacterDictionarySelection?: CliCommandRuntimeServiceContext['getCharacterDictionarySelection'];
  setCharacterDictionarySelection?: CliCommandRuntimeServiceContext['setCharacterDictionarySelection'];
  runStatsCommand: CliCommandRuntimeServiceContext['runStatsCommand'];
  runJellyfinCommand: (args: CliArgs) => Promise<void>;
  runUpdateCommand: CliCommandRuntimeServiceContext['runUpdateCommand'];
  runEnsureLinuxRuntimePluginAssetsCommand: CliCommandRuntimeServiceContext['runEnsureLinuxRuntimePluginAssetsCommand'];
  runYoutubePlaybackFlow: CliCommandRuntimeServiceContext['runYoutubePlaybackFlow'];
  ensureBackgroundStatsServer?: CliCommandRuntimeServiceContext['ensureBackgroundStatsServer'];
  openYomitanSettings: () => void;
  openConfigSettingsWindow: () => void;
  openSyncUiWindow: () => void;
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
};

export function createCliCommandContext(
  deps: CliCommandContextFactoryDeps,
): CliCommandRuntimeServiceContext & CliCommandRuntimeServiceContextHandlers {
  return {
    setLogLevel: deps.setLogLevel,
    getSocketPath: deps.getSocketPath,
    setSocketPath: deps.setSocketPath,
    getClient: deps.getMpvClient,
    showOsd: deps.showOsd,
    showPlaybackFeedback: deps.showPlaybackFeedback,
    texthookerService: deps.texthookerService,
    getTexthookerPort: deps.getTexthookerPort,
    setTexthookerPort: deps.setTexthookerPort,
    getTexthookerWebsocketUrl: deps.getTexthookerWebsocketUrl,
    shouldOpenBrowser: deps.shouldOpenBrowser,
    openInBrowser: (url: string) => {
      void deps.openExternal(url).catch((error) => {
        deps.logBrowserOpenError(url, error);
      });
    },
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
    getCharacterDictionarySelection:
      deps.getCharacterDictionarySelection ??
      (async () => ({
        seriesKey: '',
        guessTitle: null,
        current: null,
        override: null,
        candidates: [],
      })),
    setCharacterDictionarySelection:
      deps.setCharacterDictionarySelection ??
      (async () => ({
        ok: false,
        seriesKey: '',
        selected: { id: 0, title: '', episodes: null },
        staleMediaIds: [],
      })),
    runStatsCommand: deps.runStatsCommand,
    runJellyfinCommand: deps.runJellyfinCommand,
    runUpdateCommand: deps.runUpdateCommand,
    runEnsureLinuxRuntimePluginAssetsCommand: deps.runEnsureLinuxRuntimePluginAssetsCommand,
    runYoutubePlaybackFlow: deps.runYoutubePlaybackFlow,
    ensureBackgroundStatsServer: deps.ensureBackgroundStatsServer,
    openYomitanSettings: deps.openYomitanSettings,
    openConfigSettingsWindow: deps.openConfigSettingsWindow,
    openSyncUiWindow: deps.openSyncUiWindow,
    cycleSecondarySubMode: deps.cycleSecondarySubMode,
    openRuntimeOptionsPalette: deps.openRuntimeOptionsPalette,
    printHelp: deps.printHelp,
    stopApp: deps.stopApp,
    hasMainWindow: deps.hasMainWindow,
    getMultiCopyTimeoutMs: deps.getMultiCopyTimeoutMs,
    schedule: deps.schedule,
    log: deps.logInfo,
    logDebug: deps.logDebug,
    warn: deps.logWarn,
    error: deps.logError,
  };
}
