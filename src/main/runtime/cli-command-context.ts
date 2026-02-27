import type { CliArgs } from '../../cli/args';
import type {
  CliCommandRuntimeServiceContext,
  CliCommandRuntimeServiceContextHandlers,
} from '../cli-runtime';

type MpvClientLike = CliCommandRuntimeServiceContext['getClient'] extends () => infer T ? T : never;

export type CliCommandContextFactoryDeps = {
  getSocketPath: () => string;
  setSocketPath: (socketPath: string) => void;
  getMpvClient: () => MpvClientLike;
  showOsd: (text: string) => void;
  texthookerService: CliCommandRuntimeServiceContextHandlers['texthookerService'];
  getTexthookerPort: () => number;
  setTexthookerPort: (port: number) => void;
  shouldOpenBrowser: () => boolean;
  openExternal: (url: string) => Promise<unknown>;
  logBrowserOpenError: (url: string, error: unknown) => void;
  isOverlayInitialized: () => boolean;
  initializeOverlay: () => void;
  toggleVisibleOverlay: () => void;
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
  getAnilistStatus: CliCommandRuntimeServiceContext['getAnilistStatus'];
  clearAnilistToken: CliCommandRuntimeServiceContext['clearAnilistToken'];
  openAnilistSetup: CliCommandRuntimeServiceContext['openAnilistSetup'];
  openJellyfinSetup: CliCommandRuntimeServiceContext['openJellyfinSetup'];
  getAnilistQueueStatus: CliCommandRuntimeServiceContext['getAnilistQueueStatus'];
  retryAnilistQueueNow: CliCommandRuntimeServiceContext['retryAnilistQueueNow'];
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
};

export function createCliCommandContext(
  deps: CliCommandContextFactoryDeps,
): CliCommandRuntimeServiceContext & CliCommandRuntimeServiceContextHandlers {
  return {
    getSocketPath: deps.getSocketPath,
    setSocketPath: deps.setSocketPath,
    getClient: deps.getMpvClient,
    showOsd: deps.showOsd,
    texthookerService: deps.texthookerService,
    getTexthookerPort: deps.getTexthookerPort,
    setTexthookerPort: deps.setTexthookerPort,
    shouldOpenBrowser: deps.shouldOpenBrowser,
    openInBrowser: (url: string) => {
      void deps.openExternal(url).catch((error) => {
        deps.logBrowserOpenError(url, error);
      });
    },
    isOverlayInitialized: deps.isOverlayInitialized,
    initializeOverlay: deps.initializeOverlay,
    toggleVisibleOverlay: deps.toggleVisibleOverlay,
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
    getAnilistStatus: deps.getAnilistStatus,
    clearAnilistToken: deps.clearAnilistToken,
    openAnilistSetup: deps.openAnilistSetup,
    openJellyfinSetup: deps.openJellyfinSetup,
    getAnilistQueueStatus: deps.getAnilistQueueStatus,
    retryAnilistQueueNow: deps.retryAnilistQueueNow,
    runJellyfinCommand: deps.runJellyfinCommand,
    openYomitanSettings: deps.openYomitanSettings,
    cycleSecondarySubMode: deps.cycleSecondarySubMode,
    openRuntimeOptionsPalette: deps.openRuntimeOptionsPalette,
    printHelp: deps.printHelp,
    stopApp: deps.stopApp,
    hasMainWindow: deps.hasMainWindow,
    getMultiCopyTimeoutMs: deps.getMultiCopyTimeoutMs,
    schedule: deps.schedule,
    log: deps.logInfo,
    warn: deps.logWarn,
    error: deps.logError,
  };
}
