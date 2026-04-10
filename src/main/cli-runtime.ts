import { handleCliCommand, createCliCommandDepsRuntime } from '../core/services';
import type { CliArgs, CliCommandSource } from '../cli/args';
import {
  createCliCommandRuntimeServiceDeps,
  CliCommandRuntimeServiceDepsParams,
} from './dependencies';

export interface CliCommandRuntimeServiceContext {
  setLogLevel?: (level: NonNullable<CliArgs['logLevel']>) => void;
  getSocketPath: () => string;
  setSocketPath: (socketPath: string) => void;
  getClient: CliCommandRuntimeServiceDepsParams['mpv']['getClient'];
  showOsd: CliCommandRuntimeServiceDepsParams['mpv']['showOsd'];
  getTexthookerPort: () => number;
  setTexthookerPort: (port: number) => void;
  getTexthookerWebsocketUrl: () => string | undefined;
  shouldOpenBrowser: () => boolean;
  openInBrowser: (url: string) => void;
  isOverlayInitialized: () => boolean;
  initializeOverlay: () => void;
  toggleVisibleOverlay: () => void;
  openFirstRunSetup: () => void;
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
  dispatchSessionAction: CliCommandRuntimeServiceDepsParams['dispatchSessionAction'];
  getAnilistStatus: CliCommandRuntimeServiceDepsParams['anilist']['getStatus'];
  clearAnilistToken: CliCommandRuntimeServiceDepsParams['anilist']['clearToken'];
  openAnilistSetup: CliCommandRuntimeServiceDepsParams['anilist']['openSetup'];
  getAnilistQueueStatus: CliCommandRuntimeServiceDepsParams['anilist']['getQueueStatus'];
  retryAnilistQueueNow: CliCommandRuntimeServiceDepsParams['anilist']['retryQueueNow'];
  generateCharacterDictionary: CliCommandRuntimeServiceDepsParams['dictionary']['generate'];
  openJellyfinSetup: CliCommandRuntimeServiceDepsParams['jellyfin']['openSetup'];
  runStatsCommand: CliCommandRuntimeServiceDepsParams['jellyfin']['runStatsCommand'];
  runJellyfinCommand: CliCommandRuntimeServiceDepsParams['jellyfin']['runCommand'];
  runYoutubePlaybackFlow: CliCommandRuntimeServiceDepsParams['app']['runYoutubePlaybackFlow'];
  openYomitanSettings: () => void;
  cycleSecondarySubMode: () => void;
  openRuntimeOptionsPalette: () => void;
  printHelp: () => void;
  stopApp: () => void;
  hasMainWindow: () => boolean;
  getMultiCopyTimeoutMs: () => number;
  schedule: (fn: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  log: (message: string) => void;
  warn: (message: string) => void;
  error: (message: string, err: unknown) => void;
}

export interface CliCommandRuntimeServiceContextHandlers {
  texthookerService: CliCommandRuntimeServiceDepsParams['texthooker']['service'];
}

function createCliCommandDepsFromContext(
  context: CliCommandRuntimeServiceContext & CliCommandRuntimeServiceContextHandlers,
): CliCommandRuntimeServiceDepsParams {
  return {
    setLogLevel: context.setLogLevel,
    mpv: {
      getSocketPath: context.getSocketPath,
      setSocketPath: context.setSocketPath,
      getClient: context.getClient,
      showOsd: context.showOsd,
    },
    texthooker: {
      service: context.texthookerService,
      getPort: context.getTexthookerPort,
      setPort: context.setTexthookerPort,
      getWebsocketUrl: context.getTexthookerWebsocketUrl,
      shouldOpenBrowser: context.shouldOpenBrowser,
      openInBrowser: context.openInBrowser,
    },
    overlay: {
      isInitialized: context.isOverlayInitialized,
      initialize: context.initializeOverlay,
      toggleVisible: context.toggleVisibleOverlay,
      setVisible: context.setVisibleOverlay,
    },
    mining: {
      copyCurrentSubtitle: context.copyCurrentSubtitle,
      startPendingMultiCopy: context.startPendingMultiCopy,
      mineSentenceCard: context.mineSentenceCard,
      startPendingMineSentenceMultiple: context.startPendingMineSentenceMultiple,
      updateLastCardFromClipboard: context.updateLastCardFromClipboard,
      refreshKnownWords: context.refreshKnownWordCache,
      triggerFieldGrouping: context.triggerFieldGrouping,
      triggerSubsyncFromConfig: context.triggerSubsyncFromConfig,
      markLastCardAsAudioCard: context.markLastCardAsAudioCard,
    },
    anilist: {
      getStatus: context.getAnilistStatus,
      clearToken: context.clearAnilistToken,
      openSetup: context.openAnilistSetup,
      getQueueStatus: context.getAnilistQueueStatus,
      retryQueueNow: context.retryAnilistQueueNow,
    },
    dictionary: {
      generate: context.generateCharacterDictionary,
    },
    jellyfin: {
      openSetup: context.openJellyfinSetup,
      runStatsCommand: context.runStatsCommand,
      runCommand: context.runJellyfinCommand,
    },
    app: {
      stop: context.stopApp,
      hasMainWindow: context.hasMainWindow,
      runYoutubePlaybackFlow: context.runYoutubePlaybackFlow,
    },
    dispatchSessionAction: context.dispatchSessionAction,
    ui: {
      openFirstRunSetup: context.openFirstRunSetup,
      openYomitanSettings: context.openYomitanSettings,
      cycleSecondarySubMode: context.cycleSecondarySubMode,
      openRuntimeOptionsPalette: context.openRuntimeOptionsPalette,
      printHelp: context.printHelp,
    },
    getMultiCopyTimeoutMs: context.getMultiCopyTimeoutMs,
    schedule: context.schedule,
    log: context.log,
    warn: context.warn,
    error: context.error,
  };
}

export function handleCliCommandRuntimeService(
  args: CliArgs,
  source: CliCommandSource,
  params: CliCommandRuntimeServiceDepsParams,
): void {
  const deps = createCliCommandDepsRuntime(createCliCommandRuntimeServiceDeps(params));
  handleCliCommand(args, source, deps);
}

export function handleCliCommandRuntimeServiceWithContext(
  args: CliArgs,
  source: CliCommandSource,
  context: CliCommandRuntimeServiceContext & CliCommandRuntimeServiceContextHandlers,
): void {
  handleCliCommandRuntimeService(args, source, createCliCommandDepsFromContext(context));
}
