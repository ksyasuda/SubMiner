import type { CliArgs, CliCommandSource } from '../cli/args';
import type {
  CliCommandRuntimeServiceContext,
  CliCommandRuntimeServiceContextHandlers,
} from './cli-runtime';
import { composeCliStartupHandlers } from './runtime/composers/cli-startup-composer';

/** Mpv client shape required by the CLI command context (appState.mpvClient). */
type AppStateMpvClientLike = {
  setSocketPath: (socketPath: string) => void;
  connect: () => void;
} | null;

/** Mpv client shape required by the initial-args handler (getMpvClient). */
type InitialArgsMpvClientLike = { connected: boolean; connect: () => void } | null;

/** Resolved config shape consumed by the CLI command context builder. */
type ResolvedConfigLike = {
  texthooker?: { openBrowser?: boolean };
  websocket?: { enabled?: boolean | 'auto'; port?: number };
  annotationWebsocket?: { enabled?: boolean; port?: number };
};

/** Mutable app state consumed by the CLI command context builder. */
type CliCommandContextMainStateLike = {
  mpvSocketPath: string;
  mpvClient: AppStateMpvClientLike;
  texthookerPort: number;
  overlayRuntimeInitialized: boolean;
};

export interface CliStartupAppStateInput {
  appState: CliCommandContextMainStateLike;
  getInitialArgs: () => CliArgs | null | undefined;
  isBackgroundMode: () => boolean;
  isTexthookerOnlyMode: () => boolean;
  setTexthookerOnlyMode: (enabled: boolean) => void;
  hasImmersionTracker: () => boolean;
  getMpvClient: () => InitialArgsMpvClientLike;
  isOverlayRuntimeInitialized: () => boolean;
}

export interface CliStartupConfigInput {
  defaultConfig: {
    websocket: { port: number };
    annotationWebsocket: { port: number };
  };
  getResolvedConfig: () => ResolvedConfigLike;
  setCliLogLevel: (level: NonNullable<CliArgs['logLevel']>) => void;
  hasMpvWebsocketPlugin: () => boolean;
}

export interface CliStartupIoInput {
  texthookerService: CliCommandRuntimeServiceContextHandlers['texthookerService'];
  openExternal: (url: string) => Promise<void>;
  logBrowserOpenError: (url: string, error: unknown) => void;
  showMpvOsd: (text: string) => void;
  schedule: (fn: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  logInfo: (message: string) => void;
  logWarn: (message: string) => void;
  logError: (message: string, err: unknown) => void;
}

export interface CliStartupCommandsInput {
  initializeOverlayRuntime: () => void;
  toggleVisibleOverlay: () => void;
  openFirstRunSetupWindow: () => void;
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
  getAnilistStatus: CliCommandRuntimeServiceContext['getAnilistStatus'];
  clearAnilistToken: () => void;
  openAnilistSetupWindow: () => void;
  openJellyfinSetupWindow: () => void;
  getAnilistQueueStatus: CliCommandRuntimeServiceContext['getAnilistQueueStatus'];
  processNextAnilistRetryUpdate: CliCommandRuntimeServiceContext['retryAnilistQueueNow'];
  generateCharacterDictionary: CliCommandRuntimeServiceContext['generateCharacterDictionary'];
  runJellyfinCommand: (argsFromCommand: CliArgs) => Promise<void>;
  runStatsCommand: CliCommandRuntimeServiceContext['runStatsCommand'];
  runYoutubePlaybackFlow: (request: {
    url: string;
    mode: NonNullable<CliArgs['youtubeMode']>;
    source: CliCommandSource;
  }) => Promise<void>;
  openYomitanSettings: () => void;
  cycleSecondarySubMode: () => void;
  openRuntimeOptionsPalette: () => void;
  printHelp: () => void;
  stopApp: () => void;
  hasMainWindow: () => boolean;
  getMultiCopyTimeoutMs: () => number;
}

export interface CliStartupStartupInput {
  shouldEnsureTrayOnStartup: () => boolean;
  shouldRunHeadlessInitialCommand: (args: CliArgs) => boolean;
  ensureTray: () => void;
  commandNeedsOverlayStartupPrereqs: (args: CliArgs) => boolean;
  commandNeedsOverlayRuntime: (args: CliArgs) => boolean;
  ensureOverlayStartupPrereqs: () => void;
  startBackgroundWarmups: () => void;
}

export interface CliStartupRuntimeInput {
  appState: CliStartupAppStateInput;
  config: CliStartupConfigInput;
  io: CliStartupIoInput;
  commands: CliStartupCommandsInput;
  startup: CliStartupStartupInput;
  handleCliCommandRuntimeServiceWithContext: (
    args: CliArgs,
    source: CliCommandSource,
    context: CliCommandRuntimeServiceContext & CliCommandRuntimeServiceContextHandlers,
  ) => void;
}

export interface CliStartupRuntime {
  handleCliCommand: (args: CliArgs, source?: CliCommandSource) => void;
  handleInitialArgs: () => void;
}

export function createCliStartupRuntime(input: CliStartupRuntimeInput): CliStartupRuntime {
  const { handleCliCommand, handleInitialArgs } = composeCliStartupHandlers({
    cliCommandContextMainDeps: {
      appState: input.appState.appState,
      setLogLevel: (level) => input.config.setCliLogLevel(level),
      texthookerService: input.io.texthookerService,
      getResolvedConfig: () => input.config.getResolvedConfig(),
      defaultWebsocketPort: input.config.defaultConfig.websocket.port,
      defaultAnnotationWebsocketPort: input.config.defaultConfig.annotationWebsocket.port,
      hasMpvWebsocketPlugin: () => input.config.hasMpvWebsocketPlugin(),
      openExternal: (url: string) => input.io.openExternal(url),
      logBrowserOpenError: (url: string, error: unknown) =>
        input.io.logBrowserOpenError(url, error),
      showMpvOsd: (text: string) => input.io.showMpvOsd(text),
      initializeOverlayRuntime: () => input.commands.initializeOverlayRuntime(),
      toggleVisibleOverlay: () => input.commands.toggleVisibleOverlay(),
      openFirstRunSetupWindow: () => input.commands.openFirstRunSetupWindow(),
      setVisibleOverlayVisible: (visible: boolean) =>
        input.commands.setVisibleOverlayVisible(visible),
      copyCurrentSubtitle: () => input.commands.copyCurrentSubtitle(),
      startPendingMultiCopy: (timeoutMs: number) => input.commands.startPendingMultiCopy(timeoutMs),
      mineSentenceCard: () => input.commands.mineSentenceCard(),
      startPendingMineSentenceMultiple: (timeoutMs: number) =>
        input.commands.startPendingMineSentenceMultiple(timeoutMs),
      updateLastCardFromClipboard: () => input.commands.updateLastCardFromClipboard(),
      refreshKnownWordCache: () => input.commands.refreshKnownWordCache(),
      triggerFieldGrouping: () => input.commands.triggerFieldGrouping(),
      triggerSubsyncFromConfig: () => input.commands.triggerSubsyncFromConfig(),
      markLastCardAsAudioCard: () => input.commands.markLastCardAsAudioCard(),
      getAnilistStatus: () => input.commands.getAnilistStatus(),
      clearAnilistToken: () => input.commands.clearAnilistToken(),
      openAnilistSetupWindow: () => input.commands.openAnilistSetupWindow(),
      openJellyfinSetupWindow: () => input.commands.openJellyfinSetupWindow(),
      getAnilistQueueStatus: () => input.commands.getAnilistQueueStatus(),
      processNextAnilistRetryUpdate: () => input.commands.processNextAnilistRetryUpdate(),
      generateCharacterDictionary: (targetPath?: string) =>
        input.commands.generateCharacterDictionary(targetPath),
      runJellyfinCommand: (argsFromCommand: CliArgs) =>
        input.commands.runJellyfinCommand(argsFromCommand),
      runStatsCommand: (argsFromCommand: CliArgs, source: CliCommandSource) =>
        input.commands.runStatsCommand(argsFromCommand, source),
      runYoutubePlaybackFlow: (request) => input.commands.runYoutubePlaybackFlow(request),
      openYomitanSettings: () => input.commands.openYomitanSettings(),
      cycleSecondarySubMode: () => input.commands.cycleSecondarySubMode(),
      openRuntimeOptionsPalette: () => input.commands.openRuntimeOptionsPalette(),
      printHelp: () => input.commands.printHelp(),
      stopApp: () => input.commands.stopApp(),
      hasMainWindow: () => input.commands.hasMainWindow(),
      getMultiCopyTimeoutMs: () => input.commands.getMultiCopyTimeoutMs(),
      schedule: (fn: () => void, delayMs: number) => input.io.schedule(fn, delayMs),
      logInfo: (message: string) => input.io.logInfo(message),
      logWarn: (message: string) => input.io.logWarn(message),
      logError: (message: string, err: unknown) => input.io.logError(message, err),
    },
    cliCommandRuntimeHandlerMainDeps: {
      handleTexthookerOnlyModeTransitionMainDeps: {
        isTexthookerOnlyMode: () => input.appState.isTexthookerOnlyMode(),
        ensureOverlayStartupPrereqs: () => input.startup.ensureOverlayStartupPrereqs(),
        setTexthookerOnlyMode: (enabled) => input.appState.setTexthookerOnlyMode(enabled),
        commandNeedsOverlayStartupPrereqs: (args) =>
          input.startup.commandNeedsOverlayStartupPrereqs(args),
        startBackgroundWarmups: () => input.startup.startBackgroundWarmups(),
        logInfo: (message: string) => input.io.logInfo(message),
      },
      handleCliCommandRuntimeServiceWithContext: (args, source, context) =>
        input.handleCliCommandRuntimeServiceWithContext(args, source, context),
    },
    initialArgsRuntimeHandlerMainDeps: {
      getInitialArgs: () => input.appState.getInitialArgs() ?? null,
      isBackgroundMode: () => input.appState.isBackgroundMode(),
      shouldEnsureTrayOnStartup: () => input.startup.shouldEnsureTrayOnStartup(),
      shouldRunHeadlessInitialCommand: (args) =>
        input.startup.shouldRunHeadlessInitialCommand(args),
      ensureTray: () => input.startup.ensureTray(),
      isTexthookerOnlyMode: () => input.appState.isTexthookerOnlyMode(),
      hasImmersionTracker: () => input.appState.hasImmersionTracker(),
      getMpvClient: () => input.appState.getMpvClient(),
      commandNeedsOverlayStartupPrereqs: (args) =>
        input.startup.commandNeedsOverlayStartupPrereqs(args),
      commandNeedsOverlayRuntime: (args) => input.startup.commandNeedsOverlayRuntime(args),
      ensureOverlayStartupPrereqs: () => input.startup.ensureOverlayStartupPrereqs(),
      isOverlayRuntimeInitialized: () => input.appState.isOverlayRuntimeInitialized(),
      initializeOverlayRuntime: () => input.commands.initializeOverlayRuntime(),
      logInfo: (message) => input.io.logInfo(message),
    },
  });

  return {
    handleCliCommand,
    handleInitialArgs,
  };
}
