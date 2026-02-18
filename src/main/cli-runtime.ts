import {
  handleCliCommand,
  createCliCommandDepsRuntime,
} from "../core/services";
import type { CliArgs, CliCommandSource } from "../cli/args";
import {
  createCliCommandRuntimeServiceDeps,
  CliCommandRuntimeServiceDepsParams,
} from "./dependencies";

export interface CliCommandRuntimeServiceContext {
  getSocketPath: () => string;
  setSocketPath: (socketPath: string) => void;
  getClient: CliCommandRuntimeServiceDepsParams["mpv"]["getClient"];
  showOsd: CliCommandRuntimeServiceDepsParams["mpv"]["showOsd"];
  getTexthookerPort: () => number;
  setTexthookerPort: (port: number) => void;
  shouldOpenBrowser: () => boolean;
  openInBrowser: (url: string) => void;
  isOverlayInitialized: () => boolean;
  initializeOverlay: () => void;
  toggleVisibleOverlay: () => void;
  toggleInvisibleOverlay: () => void;
  setVisibleOverlay: (visible: boolean) => void;
  setInvisibleOverlay: (visible: boolean) => void;
  copyCurrentSubtitle: () => void;
  startPendingMultiCopy: (timeoutMs: number) => void;
  mineSentenceCard: () => Promise<void>;
  startPendingMineSentenceMultiple: (timeoutMs: number) => void;
  updateLastCardFromClipboard: () => Promise<void>;
  refreshKnownWordCache: () => Promise<void>;
  triggerFieldGrouping: () => Promise<void>;
  triggerSubsyncFromConfig: () => Promise<void>;
  markLastCardAsAudioCard: () => Promise<void>;
  getAnilistStatus: CliCommandRuntimeServiceDepsParams["anilist"]["getStatus"];
  clearAnilistToken: CliCommandRuntimeServiceDepsParams["anilist"]["clearToken"];
  openAnilistSetup: CliCommandRuntimeServiceDepsParams["anilist"]["openSetup"];
  getAnilistQueueStatus: CliCommandRuntimeServiceDepsParams["anilist"]["getQueueStatus"];
  retryAnilistQueueNow: CliCommandRuntimeServiceDepsParams["anilist"]["retryQueueNow"];
  openJellyfinSetup: CliCommandRuntimeServiceDepsParams["jellyfin"]["openSetup"];
  runJellyfinCommand: CliCommandRuntimeServiceDepsParams["jellyfin"]["runCommand"];
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
  texthookerService: CliCommandRuntimeServiceDepsParams["texthooker"]["service"];
}

function createCliCommandDepsFromContext(
  context: CliCommandRuntimeServiceContext &
    CliCommandRuntimeServiceContextHandlers,
): CliCommandRuntimeServiceDepsParams {
  return {
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
      shouldOpenBrowser: context.shouldOpenBrowser,
      openInBrowser: context.openInBrowser,
    },
    overlay: {
      isInitialized: context.isOverlayInitialized,
      initialize: context.initializeOverlay,
      toggleVisible: context.toggleVisibleOverlay,
      toggleInvisible: context.toggleInvisibleOverlay,
      setVisible: context.setVisibleOverlay,
      setInvisible: context.setInvisibleOverlay,
    },
    mining: {
      copyCurrentSubtitle: context.copyCurrentSubtitle,
      startPendingMultiCopy: context.startPendingMultiCopy,
      mineSentenceCard: context.mineSentenceCard,
      startPendingMineSentenceMultiple:
        context.startPendingMineSentenceMultiple,
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
    jellyfin: {
      openSetup: context.openJellyfinSetup,
      runCommand: context.runJellyfinCommand,
    },
    ui: {
      openYomitanSettings: context.openYomitanSettings,
      cycleSecondarySubMode: context.cycleSecondarySubMode,
      openRuntimeOptionsPalette: context.openRuntimeOptionsPalette,
      printHelp: context.printHelp,
    },
    app: {
      stop: context.stopApp,
      hasMainWindow: context.hasMainWindow,
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
  const deps = createCliCommandDepsRuntime(
    createCliCommandRuntimeServiceDeps(params),
  );
  handleCliCommand(args, source, deps);
}

export function handleCliCommandRuntimeServiceWithContext(
  args: CliArgs,
  source: CliCommandSource,
  context: CliCommandRuntimeServiceContext &
    CliCommandRuntimeServiceContextHandlers,
): void {
  handleCliCommandRuntimeService(
    args,
    source,
    createCliCommandDepsFromContext(context),
  );
}
