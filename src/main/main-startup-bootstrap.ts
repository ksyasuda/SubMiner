import type { CliArgs, CliCommandSource } from '../cli/args';
import type { LogLevelSource } from '../logger';
import type { ConfigValidationWarning, ResolvedConfig, SubtitleData } from '../types';
import type { StartupBootstrapRuntimeDeps } from '../core/services/startup';
import { resolveKeybindings } from '../core/utils';
import { RuntimeOptionsManager } from '../runtime-options';
import { SubtitleTimingTracker } from '../subtitle-timing-tracker';
import type { AppReadyRuntimeInput } from './app-ready-runtime';
import type {
  CliCommandRuntimeServiceContext,
  CliCommandRuntimeServiceContextHandlers,
} from './cli-runtime';
import type {
  StartupBootstrapAppStateLike,
  StartupBootstrapMpvClientLike,
  StartupBootstrapOverlayUiLike,
  StartupBootstrapSubtitleWebsocketLike,
} from './main-startup-bootstrap-types';
import type { MainStartupRuntime } from './main-startup-runtime';
import { createMainStartupRuntime } from './main-startup-runtime';

export interface MainStartupBootstrapInput<TStartupState> {
  appState: StartupBootstrapAppStateLike;
  appLifecycle: {
    app: unknown;
    argv: string[];
    platform: NodeJS.Platform;
  };
  config: {
    configService: {
      reloadConfigStrict: AppReadyRuntimeInput['reload']['reloadConfigStrict'];
      getConfigPath: () => string;
      getWarnings: () => ConfigValidationWarning[];
      getConfig: () => ResolvedConfig;
    };
    configHotReloadRuntime: {
      start: () => void;
    };
    configDerivedRuntime: {
      shouldAutoInitializeOverlayRuntimeFromConfig: () => boolean;
    };
    ensureDefaultConfigBootstrap: (options: {
      configDir: string;
      configFilePaths: unknown;
      generateTemplate: () => string;
    }) => void;
    getDefaultConfigFilePaths: (configDir: string) => unknown;
    generateConfigTemplate: (config: ResolvedConfig) => string;
    defaultConfig: ResolvedConfig;
    defaultKeybindings: unknown;
    configDir: string;
  };
  logging: {
    appLogger: {
      logInfo: (message: string) => void;
      logWarning: (message: string) => void;
      logConfigWarning: (warning: ConfigValidationWarning) => void;
      logNoRunningInstance: () => void;
    };
    logger: {
      info: (message: string) => void;
      warn: (message: string, error?: unknown) => void;
      error: (message: string, error?: unknown) => void;
      debug: (message: string) => void;
    };
    setLogLevel: (level: string, source: LogLevelSource) => void;
  };
  shell: {
    dialog: {
      showErrorBox: (title: string, message: string) => void;
    };
    shell: {
      openExternal: (url: string) => Promise<void>;
    };
    showDesktopNotification: (title: string, options: { body: string }) => void;
  };
  runtime: {
    subtitle: {
      loadSubtitlePosition: () => void;
      invalidateTokenizationCache: () => void;
      refreshSubtitlePrefetchFromActiveTrack: () => Promise<void>;
    };
    overlayUi: {
      get: () => StartupBootstrapOverlayUiLike | undefined;
    };
    overlayManager: {
      getMainWindow: () => unknown | null;
    };
    firstRun: {
      ensureSetupStateInitialized: () => Promise<{ state: { status: string } }>;
      openFirstRunSetupWindow: () => void;
    };
    anilist: {
      refreshAnilistClientSecretStateIfEnabled: (options: {
        force: boolean;
        allowSetupPrompt?: boolean;
      }) => Promise<unknown>;
      openAnilistSetupWindow: () => void;
      getStatusSnapshot: CliCommandRuntimeServiceContext['getAnilistStatus'];
      clearTokenState: () => void;
      getQueueStatusSnapshot: CliCommandRuntimeServiceContext['getAnilistQueueStatus'];
      processNextAnilistRetryUpdate: CliCommandRuntimeServiceContext['retryAnilistQueueNow'];
    };
    jellyfin: {
      startJellyfinRemoteSession: () => Promise<void>;
      openJellyfinSetupWindow: () => void;
      runJellyfinCommand: (argsFromCommand: CliArgs) => Promise<void>;
    };
    stats: {
      ensureImmersionTrackerStarted: () => void;
      runStatsCliCommand: CliCommandRuntimeServiceContext['runStatsCommand'];
    };
    mining: {
      copyCurrentSubtitle: () => void;
      mineSentenceCard: () => Promise<void>;
      updateLastCardFromClipboard: () => Promise<void>;
      refreshKnownWordCache: () => Promise<void>;
      triggerFieldGrouping: () => Promise<void>;
      markLastCardAsAudioCard: () => Promise<void>;
    };
    yomitan: {
      loadYomitanExtension: () => Promise<unknown>;
      ensureYomitanExtensionLoaded: () => Promise<unknown>;
      openYomitanSettings: () => void;
    };
    subsyncRuntime: {
      triggerFromConfig: () => Promise<void>;
    };
    dictionarySupport: {
      generateCharacterDictionaryForCurrentMedia: CliCommandRuntimeServiceContext['generateCharacterDictionary'];
    };
    texthookerService: CliCommandRuntimeServiceContextHandlers['texthookerService'];
    subtitleWsService: StartupBootstrapSubtitleWebsocketLike;
    annotationSubtitleWsService: StartupBootstrapSubtitleWebsocketLike;
    immersion: AppReadyRuntimeInput['immersion'];
  };
  commands: {
    createMpvClientRuntimeService: () => StartupBootstrapMpvClientLike;
    createMecabTokenizerAndCheck: () => Promise<void>;
    prewarmSubtitleDictionaries: () => Promise<void>;
    startBackgroundWarmupsIfAllowed: () => void;
    startBackgroundWarmups: () => void;
    runHeadlessInitialCommand: () => Promise<void>;
    startPendingMultiCopy: (timeoutMs: number) => void;
    startPendingMineSentenceMultiple: (timeoutMs: number) => void;
    cycleSecondarySubMode: () => void;
    refreshOverlayShortcuts: () => void;
    hasMpvWebsocketPlugin: () => boolean;
    startTexthooker: (port: number, websocketUrl?: string) => void;
    showMpvOsd: (text: string) => void;
    shouldAutoOpenFirstRunSetup: (args: CliArgs) => boolean;
    generateCharacterDictionary: CliCommandRuntimeServiceContext['generateCharacterDictionary'];
    runYoutubePlaybackFlow: (request: {
      url: string;
      mode: NonNullable<CliArgs['youtubeMode']>;
      source: CliCommandSource;
    }) => Promise<void>;
    getMultiCopyTimeoutMs: () => number;
    shouldEnsureTrayOnStartupForInitialArgs: (
      platform: NodeJS.Platform,
      initialArgs: CliArgs | null | undefined,
    ) => boolean;
    isHeadlessInitialCommand: (args: CliArgs) => boolean;
    commandNeedsOverlayStartupPrereqs: (args: CliArgs) => boolean;
    commandNeedsOverlayRuntime: (args: CliArgs) => boolean;
    handleCliCommandRuntimeServiceWithContext: (
      args: CliArgs,
      source: CliCommandSource,
      context: CliCommandRuntimeServiceContext & CliCommandRuntimeServiceContextHandlers,
    ) => void;
    shouldStartApp: (args: CliArgs) => boolean;
    parseArgs: (argv: string[]) => CliArgs;
    printHelp: (defaultTexthookerPort: number) => void;
    onWillQuitCleanupHandler: () => void;
    shouldRestoreWindowsOnActivateHandler: () => boolean;
    restoreWindowsOnActivateHandler: () => void;
    forceX11Backend: (args: CliArgs) => void;
    enforceUnsupportedWaylandMode: (args: CliArgs) => void;
    getDefaultSocketPathHandler: () => string;
    generateDefaultConfigFile: (
      args: CliArgs,
      options: {
        configDir: string;
        defaultConfig: unknown;
        generateTemplate: (config: unknown) => string;
      },
    ) => Promise<number>;
    runStartupBootstrapRuntime: (deps: StartupBootstrapRuntimeDeps) => TStartupState;
    applyStartupState: (startupState: TStartupState) => void;
    getStartupModeFlags: (initialArgs: CliArgs | null | undefined) => {
      shouldUseMinimalStartup: boolean;
      shouldSkipHeavyStartup: boolean;
    };
    requestAppQuit: () => void;
  };
  constants: {
    defaultTexthookerPort: number;
  };
}

export function createMainStartupBootstrap<TStartupState>(
  input: MainStartupBootstrapInput<TStartupState>,
): MainStartupRuntime<TStartupState> {
  let startup: MainStartupRuntime<TStartupState> | null = null;
  const getStartup = (): MainStartupRuntime<TStartupState> => {
    if (!startup) {
      throw new Error('Main startup runtime not initialized');
    }
    return startup;
  };
  const getOverlayUi = (): StartupBootstrapOverlayUiLike | undefined =>
    input.runtime.overlayUi.get();
  const getSubtitlePayload = (): SubtitleData | null =>
    input.appState.currentSubtitleData ??
    (input.appState.currentSubText
      ? {
          text: input.appState.currentSubText,
          tokens: null,
          startTime: input.appState.mpvClient?.currentSubStart ?? null,
          endTime: input.appState.mpvClient?.currentSubEnd ?? null,
        }
      : null);
  const getSubtitleFrequencyOptions = () => ({
    enabled: input.config.configService.getConfig().subtitleStyle.frequencyDictionary.enabled,
    topX: input.config.configService.getConfig().subtitleStyle.frequencyDictionary.topX,
    mode: input.config.configService.getConfig().subtitleStyle.frequencyDictionary.mode,
  });

  startup = createMainStartupRuntime<TStartupState>({
    appReady: {
      reload: {
        reloadConfigStrict: () => input.config.configService.reloadConfigStrict(),
        logInfo: (message) => input.logging.appLogger.logInfo(message),
        logWarning: (message) => input.logging.appLogger.logWarning(message),
        showDesktopNotification: (title, options) =>
          input.shell.showDesktopNotification(title, options),
        startConfigHotReload: () => input.config.configHotReloadRuntime.start(),
        refreshAnilistClientSecretState: (options) =>
          input.runtime.anilist.refreshAnilistClientSecretStateIfEnabled(options),
        failHandlers: {
          logError: (details) => input.logging.logger.error(details),
          showErrorBox: (title, details) => input.shell.dialog.showErrorBox(title, details),
          quit: () => input.commands.requestAppQuit(),
        },
      },
      criticalConfig: {
        getConfigPath: () => input.config.configService.getConfigPath(),
        failHandlers: {
          logError: (message) => input.logging.logger.error(message),
          showErrorBox: (title, message) => input.shell.dialog.showErrorBox(title, message),
          quit: () => input.commands.requestAppQuit(),
        },
      },
      runner: {
        ensureDefaultConfigBootstrap: () => {
          input.config.ensureDefaultConfigBootstrap({
            configDir: input.config.configDir,
            configFilePaths: input.config.getDefaultConfigFilePaths(input.config.configDir),
            generateTemplate: () => input.config.generateConfigTemplate(input.config.defaultConfig),
          });
        },
        getSubtitlePosition: () => input.appState.subtitlePosition,
        loadSubtitlePosition: () => input.runtime.subtitle.loadSubtitlePosition(),
        getKeybindingsCount: () => input.appState.keybindings.length,
        resolveKeybindings: () => {
          input.appState.keybindings = resolveKeybindings(
            input.config.configService.getConfig(),
            input.config.defaultKeybindings as never,
          );
        },
        hasMpvClient: () => Boolean(input.appState.mpvClient),
        createMpvClient: () => {
          input.appState.mpvClient = input.commands.createMpvClientRuntimeService();
        },
        getRuntimeOptionsManager: () => input.appState.runtimeOptionsManager,
        getResolvedConfig: () => input.config.configService.getConfig(),
        getConfigWarnings: () => input.config.configService.getWarnings(),
        logConfigWarning: (warning) => input.logging.appLogger.logConfigWarning(warning),
        setLogLevel: (level, source) => input.logging.setLogLevel(level, source),
        initRuntimeOptionsManager: () => {
          input.appState.runtimeOptionsManager = new RuntimeOptionsManager(
            () => input.config.configService.getConfig().ankiConnect,
            {
              applyAnkiPatch: (patch: unknown) => {
                (
                  input.appState.ankiIntegration as {
                    applyRuntimeConfigPatch?: (patch: unknown) => void;
                  } | null
                )?.applyRuntimeConfigPatch?.(patch);
              },
              getSubtitleStyleConfig: () => input.config.configService.getConfig().subtitleStyle,
              onOptionsChanged: () => {
                input.runtime.subtitle.invalidateTokenizationCache();
                void input.runtime.subtitle.refreshSubtitlePrefetchFromActiveTrack();
                getOverlayUi()?.broadcastRuntimeOptionsChanged();
                input.commands.refreshOverlayShortcuts();
              },
            },
          );
        },
        getSubtitleTimingTracker: () => input.appState.subtitleTimingTracker,
        createSubtitleTimingTracker: () => {
          input.appState.subtitleTimingTracker = new SubtitleTimingTracker();
        },
        setSecondarySubMode: (mode) => {
          input.appState.secondarySubMode = mode;
        },
        defaultSecondarySubMode: 'hover',
        defaultWebsocketPort: input.config.defaultConfig.websocket.port,
        defaultAnnotationWebsocketPort: input.config.defaultConfig.annotationWebsocket.port,
        defaultTexthookerPort: input.constants.defaultTexthookerPort,
        hasMpvWebsocketPlugin: () => input.commands.hasMpvWebsocketPlugin(),
        startSubtitleWebsocket: (port) => {
          input.runtime.subtitleWsService.start(
            port,
            getSubtitlePayload,
            getSubtitleFrequencyOptions,
          );
        },
        startAnnotationWebsocket: (port) => {
          input.runtime.annotationSubtitleWsService.start(
            port,
            getSubtitlePayload,
            getSubtitleFrequencyOptions,
          );
        },
        startTexthooker: (port, websocketUrl) => input.commands.startTexthooker(port, websocketUrl),
        log: (message) => input.logging.appLogger.logInfo(message),
        createMecabTokenizerAndCheck: () => input.commands.createMecabTokenizerAndCheck(),
        createImmersionTracker: () => {
          input.runtime.stats.ensureImmersionTrackerStarted();
        },
        startJellyfinRemoteSession: () => input.runtime.jellyfin.startJellyfinRemoteSession(),
        loadYomitanExtension: async () => {
          await input.runtime.yomitan.loadYomitanExtension();
        },
        ensureYomitanExtensionLoaded: async () => {
          await input.runtime.yomitan.ensureYomitanExtensionLoaded();
        },
        handleFirstRunSetup: async () => {
          const snapshot = await input.runtime.firstRun.ensureSetupStateInitialized();
          input.appState.firstRunSetupCompleted = snapshot.state.status === 'completed';
          if (
            input.appState.initialArgs &&
            input.commands.shouldAutoOpenFirstRunSetup(input.appState.initialArgs) &&
            snapshot.state.status !== 'completed'
          ) {
            input.runtime.firstRun.openFirstRunSetupWindow();
          }
        },
        prewarmSubtitleDictionaries: () => input.commands.prewarmSubtitleDictionaries(),
        startBackgroundWarmups: () => input.commands.startBackgroundWarmupsIfAllowed(),
        texthookerOnlyMode: input.appState.texthookerOnlyMode,
        shouldAutoInitializeOverlayRuntimeFromConfig: () =>
          input.appState.backgroundMode
            ? false
            : input.config.configDerivedRuntime.shouldAutoInitializeOverlayRuntimeFromConfig(),
        setVisibleOverlayVisible: (visible) => getOverlayUi()?.setVisibleOverlayVisible(visible),
        initializeOverlayRuntime: () => getOverlayUi()?.initializeOverlayRuntime(),
        ensureOverlayWindowsReadyForVisibilityActions: () =>
          getOverlayUi()?.ensureOverlayWindowsReadyForVisibilityActions(),
        runHeadlessInitialCommand: () => input.commands.runHeadlessInitialCommand(),
        handleInitialArgs: () => getStartup().handleInitialArgs(),
        shouldRunHeadlessInitialCommand: () =>
          Boolean(
            input.appState.initialArgs &&
            input.commands.isHeadlessInitialCommand(input.appState.initialArgs),
          ),
        shouldUseMinimalStartup: () =>
          input.commands.getStartupModeFlags(input.appState.initialArgs).shouldUseMinimalStartup,
        shouldSkipHeavyStartup: () =>
          input.commands.getStartupModeFlags(input.appState.initialArgs).shouldSkipHeavyStartup,
        logDebug: (message) => input.logging.logger.debug(message),
        now: () => Date.now(),
      },
      immersion: input.runtime.immersion,
      isOverlayRuntimeInitialized: () => input.appState.overlayRuntimeInitialized,
    },
    cli: {
      appState: {
        appState: input.appState,
        getInitialArgs: () => input.appState.initialArgs,
        isBackgroundMode: () => input.appState.backgroundMode,
        isTexthookerOnlyMode: () => input.appState.texthookerOnlyMode,
        setTexthookerOnlyMode: (enabled) => {
          input.appState.texthookerOnlyMode = enabled;
        },
        hasImmersionTracker: () => Boolean(input.appState.immersionTracker),
        getMpvClient: () => input.appState.mpvClient,
        isOverlayRuntimeInitialized: () => input.appState.overlayRuntimeInitialized,
      },
      config: {
        defaultConfig: input.config.defaultConfig,
        getResolvedConfig: () => input.config.configService.getConfig(),
        setCliLogLevel: (level) => input.logging.setLogLevel(level, 'cli'),
        hasMpvWebsocketPlugin: () => true,
      },
      io: {
        texthookerService: input.runtime.texthookerService,
        openExternal: (url) => input.shell.shell.openExternal(url),
        logBrowserOpenError: (url, error) =>
          input.logging.logger.error(`Failed to open browser for texthooker URL: ${url}`, error),
        showMpvOsd: (text) => input.commands.showMpvOsd(text),
        schedule: (fn, delayMs) => setTimeout(fn, delayMs),
        logInfo: (message) => input.logging.logger.info(message),
        logWarn: (message) => input.logging.logger.warn(message),
        logError: (message, err) => input.logging.logger.error(message, err),
      },
      commands: {
        initializeOverlayRuntime: () => getOverlayUi()?.initializeOverlayRuntime(),
        toggleVisibleOverlay: () => getOverlayUi()?.toggleVisibleOverlay(),
        openFirstRunSetupWindow: () => input.runtime.firstRun.openFirstRunSetupWindow(),
        setVisibleOverlayVisible: (visible) => getOverlayUi()?.setVisibleOverlayVisible(visible),
        copyCurrentSubtitle: () => input.runtime.mining.copyCurrentSubtitle(),
        startPendingMultiCopy: (timeoutMs) => input.commands.startPendingMultiCopy(timeoutMs),
        mineSentenceCard: () => input.runtime.mining.mineSentenceCard(),
        startPendingMineSentenceMultiple: (timeoutMs) =>
          input.commands.startPendingMineSentenceMultiple(timeoutMs),
        updateLastCardFromClipboard: () => input.runtime.mining.updateLastCardFromClipboard(),
        refreshKnownWordCache: () => input.runtime.mining.refreshKnownWordCache(),
        triggerFieldGrouping: () => input.runtime.mining.triggerFieldGrouping(),
        triggerSubsyncFromConfig: () => input.runtime.subsyncRuntime.triggerFromConfig(),
        markLastCardAsAudioCard: () => input.runtime.mining.markLastCardAsAudioCard(),
        getAnilistStatus: () => input.runtime.anilist.getStatusSnapshot(),
        clearAnilistToken: () => input.runtime.anilist.clearTokenState(),
        openAnilistSetupWindow: () => input.runtime.anilist.openAnilistSetupWindow(),
        openJellyfinSetupWindow: () => input.runtime.jellyfin.openJellyfinSetupWindow(),
        getAnilistQueueStatus: () => input.runtime.anilist.getQueueStatusSnapshot(),
        processNextAnilistRetryUpdate: () => input.runtime.anilist.processNextAnilistRetryUpdate(),
        generateCharacterDictionary: (targetPath?: string) =>
          input.commands.generateCharacterDictionary(targetPath),
        runJellyfinCommand: (argsFromCommand) =>
          input.runtime.jellyfin.runJellyfinCommand(argsFromCommand),
        runStatsCommand: (argsFromCommand, source) =>
          input.runtime.stats.runStatsCliCommand(argsFromCommand, source),
        runYoutubePlaybackFlow: (request) => input.commands.runYoutubePlaybackFlow(request),
        openYomitanSettings: () => input.runtime.yomitan.openYomitanSettings(),
        cycleSecondarySubMode: () => input.commands.cycleSecondarySubMode(),
        openRuntimeOptionsPalette: () => getOverlayUi()?.openRuntimeOptionsPalette(),
        printHelp: () => input.commands.printHelp(input.constants.defaultTexthookerPort),
        stopApp: () => input.commands.requestAppQuit(),
        hasMainWindow: () => Boolean(input.runtime.overlayManager.getMainWindow()),
        getMultiCopyTimeoutMs: () => input.commands.getMultiCopyTimeoutMs(),
      },
      startup: {
        shouldEnsureTrayOnStartup: () =>
          input.commands.shouldEnsureTrayOnStartupForInitialArgs(
            input.appLifecycle.platform,
            input.appState.initialArgs,
          ),
        shouldRunHeadlessInitialCommand: (args) => input.commands.isHeadlessInitialCommand(args),
        ensureTray: () => getOverlayUi()?.ensureTray(),
        commandNeedsOverlayStartupPrereqs: (args) =>
          input.commands.commandNeedsOverlayStartupPrereqs(args),
        commandNeedsOverlayRuntime: (args) => input.commands.commandNeedsOverlayRuntime(args),
        ensureOverlayStartupPrereqs: () => getStartup().appReady.ensureOverlayStartupPrereqs(),
        startBackgroundWarmups: () => input.commands.startBackgroundWarmups(),
      },
      handleCliCommandRuntimeServiceWithContext: (args, source, context) =>
        input.commands.handleCliCommandRuntimeServiceWithContext(args, source, context),
    },
    headless: {
      appLifecycleRuntimeRunnerMainDeps: {
        app: input.appLifecycle.app as never,
        platform: input.appLifecycle.platform,
        shouldStartApp: (nextArgs) => input.commands.shouldStartApp(nextArgs),
        parseArgs: (argv) => input.commands.parseArgs(argv),
        handleCliCommand: (nextArgs, source) => getStartup().handleCliCommand(nextArgs, source),
        printHelp: () => input.commands.printHelp(input.constants.defaultTexthookerPort),
        logNoRunningInstance: () => input.logging.appLogger.logNoRunningInstance(),
        onReady: (): Promise<void> => getStartup().appReady.runAppReady(),
        onWillQuitCleanup: () => input.commands.onWillQuitCleanupHandler(),
        shouldRestoreWindowsOnActivate: () =>
          input.commands.shouldRestoreWindowsOnActivateHandler(),
        restoreWindowsOnActivate: () => input.commands.restoreWindowsOnActivateHandler(),
        shouldQuitOnWindowAllClosed: () => !input.appState.backgroundMode,
      },
      bootstrap: {
        argv: input.appLifecycle.argv,
        parseArgs: (argv) => input.commands.parseArgs(argv),
        setLogLevel: (level, source) => input.logging.setLogLevel(level, source),
        forceX11Backend: (args) => input.commands.forceX11Backend(args),
        enforceUnsupportedWaylandMode: (args) => input.commands.enforceUnsupportedWaylandMode(args),
        shouldStartApp: (args) => input.commands.shouldStartApp(args),
        getDefaultSocketPath: () => input.commands.getDefaultSocketPathHandler(),
        defaultTexthookerPort: input.constants.defaultTexthookerPort,
        configDir: input.config.configDir,
        defaultConfig: input.config.defaultConfig,
        generateConfigTemplate: (config) => input.config.generateConfigTemplate(config),
        generateDefaultConfigFile: (args, options) =>
          input.commands.generateDefaultConfigFile(args, options),
        setExitCode: (exitCode) => {
          process.exitCode = exitCode;
        },
        quitApp: () => input.commands.requestAppQuit(),
        logGenerateConfigError: (message) => input.logging.logger.error(message),
        startAppLifecycle: () => {},
      },
      runStartupBootstrapRuntime: (deps) => input.commands.runStartupBootstrapRuntime(deps),
      applyStartupState: (startupState) => input.commands.applyStartupState(startupState),
    },
  });

  return startup;
}
