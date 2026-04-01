import type { LogLevelSource } from '../logger';
import type { ConfigValidationWarning, SecondarySubMode } from '../types';
import { composeAppReadyRuntime } from './runtime/composers/app-ready-composer';

type AppReadyConfigLike = {
  logging?: {
    level?: string;
  };
};

type SubtitlePositionLike = unknown;
type RuntimeOptionsManagerLike = unknown;
type SubtitleTimingTrackerLike = unknown;
type ImmersionTrackingConfigLike = {
  immersionTracking?: {
    enabled?: boolean;
  };
};
type MpvClientLike = {
  connected: boolean;
  connect: () => void;
};

export interface AppReadyReloadConfigInput {
  reloadConfigStrict: () =>
    | { ok: true; path: string; warnings: ConfigValidationWarning[] }
    | { ok: false; path: string; error: string };
  logInfo: (message: string) => void;
  logWarning: (message: string) => void;
  showDesktopNotification: (title: string, options: { body: string }) => void;
  startConfigHotReload: () => void;
  refreshAnilistClientSecretState: (options: { force: boolean }) => Promise<unknown>;
  failHandlers: {
    logError: (details: string) => void;
    showErrorBox: (title: string, details: string) => void;
    setExitCode?: (code: number) => void;
    quit: () => void;
  };
}

export interface AppReadyCriticalConfigInput {
  getConfigPath: () => string;
  failHandlers: {
    logError: (details: string) => void;
    showErrorBox: (title: string, details: string) => void;
    setExitCode?: (code: number) => void;
    quit: () => void;
  };
}

export interface AppReadyImmersionInput {
  getResolvedConfig: () => ImmersionTrackingConfigLike;
  getConfiguredDbPath: () => string;
  createTrackerService: (params: {
    dbPath: string;
    policy: {
      batchSize: number;
      flushIntervalMs: number;
      queueCap: number;
      payloadCapBytes: number;
      maintenanceIntervalMs: number;
      retention: {
        eventsDays: number;
        telemetryDays: number;
        sessionsDays: number;
        dailyRollupsDays: number;
        monthlyRollupsDays: number;
        vacuumIntervalDays: number;
      };
    };
  }) => unknown;
  setTracker: (tracker: unknown | null) => void;
  getMpvClient: () => MpvClientLike | null;
  shouldAutoConnectMpv?: () => boolean;
  seedTrackerFromCurrentMedia: () => void;
  logInfo: (message: string) => void;
  logDebug: (message: string) => void;
  logWarn: (message: string, details: unknown) => void;
}

export interface AppReadyRunnerInput<TConfig extends AppReadyConfigLike = AppReadyConfigLike> {
  ensureDefaultConfigBootstrap: () => void;
  getSubtitlePosition: () => SubtitlePositionLike | null;
  loadSubtitlePosition: () => void;
  getKeybindingsCount: () => number;
  resolveKeybindings: () => void;
  hasMpvClient: () => boolean;
  createMpvClient: () => void;
  getRuntimeOptionsManager: () => RuntimeOptionsManagerLike | null;
  initRuntimeOptionsManager: () => void;
  getSubtitleTimingTracker: () => SubtitleTimingTrackerLike | null;
  createSubtitleTimingTracker: () => void;
  getResolvedConfig: () => TConfig;
  getConfigWarnings: () => ConfigValidationWarning[];
  logConfigWarning: (warning: ConfigValidationWarning) => void;
  setLogLevel: (level: string, source: LogLevelSource) => void;
  setSecondarySubMode: (mode: SecondarySubMode) => void;
  defaultSecondarySubMode: SecondarySubMode;
  defaultWebsocketPort: number;
  defaultAnnotationWebsocketPort: number;
  defaultTexthookerPort: number;
  hasMpvWebsocketPlugin: () => boolean;
  startSubtitleWebsocket: (port: number) => void;
  startAnnotationWebsocket: (port: number) => void;
  startTexthooker: (port: number, websocketUrl?: string) => void;
  log: (message: string) => void;
  createMecabTokenizerAndCheck: () => Promise<void>;
  createImmersionTracker?: () => void;
  startJellyfinRemoteSession?: () => Promise<void>;
  loadYomitanExtension: () => Promise<void>;
  ensureYomitanExtensionLoaded: () => Promise<unknown>;
  handleFirstRunSetup: () => Promise<void>;
  prewarmSubtitleDictionaries?: () => Promise<void>;
  startBackgroundWarmups: () => void;
  texthookerOnlyMode: boolean;
  shouldAutoInitializeOverlayRuntimeFromConfig: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  initializeOverlayRuntime: () => void;
  ensureOverlayWindowsReadyForVisibilityActions: () => void;
  runHeadlessInitialCommand?: () => Promise<void>;
  handleInitialArgs: () => void;
  onCriticalConfigErrors?: (errors: string[]) => void;
  logDebug?: (message: string) => void;
  now?: () => number;
  shouldRunHeadlessInitialCommand?: () => boolean;
  shouldUseMinimalStartup?: () => boolean;
  shouldSkipHeavyStartup?: () => boolean;
}

export interface AppReadyRuntimeInput<TConfig extends AppReadyConfigLike = AppReadyConfigLike> {
  reload: AppReadyReloadConfigInput;
  criticalConfig: AppReadyCriticalConfigInput;
  immersion: AppReadyImmersionInput;
  runner: AppReadyRunnerInput<TConfig>;
  isOverlayRuntimeInitialized: () => boolean;
}

export interface AppReadyRuntime {
  reloadConfig: () => void;
  criticalConfigError: (errors: string[]) => never;
  ensureOverlayStartupPrereqs: () => void;
  ensureYoutubePlaybackRuntimeReady: () => Promise<void>;
  runAppReady: () => Promise<void>;
}

export function createAppReadyRuntime<TConfig extends AppReadyConfigLike>(
  input: AppReadyRuntimeInput<TConfig>,
): AppReadyRuntime {
  const ensureSubtitlePositionLoaded = (): void => {
    if (input.runner.getSubtitlePosition() === null) {
      input.runner.loadSubtitlePosition();
    }
  };

  const ensureKeybindingsResolved = (): void => {
    if (input.runner.getKeybindingsCount() === 0) {
      input.runner.resolveKeybindings();
    }
  };

  const ensureMpvClientCreated = (): void => {
    if (!input.runner.hasMpvClient()) {
      input.runner.createMpvClient();
    }
  };

  const ensureRuntimeOptionsManagerInitialized = (): void => {
    if (!input.runner.getRuntimeOptionsManager()) {
      input.runner.initRuntimeOptionsManager();
    }
  };

  const ensureSubtitleTimingTrackerCreated = (): void => {
    if (!input.runner.getSubtitleTimingTracker()) {
      input.runner.createSubtitleTimingTracker();
    }
  };

  const ensureOverlayStartupPrereqs = (): void => {
    ensureSubtitlePositionLoaded();
    ensureKeybindingsResolved();
    ensureMpvClientCreated();
    ensureRuntimeOptionsManagerInitialized();
    ensureSubtitleTimingTrackerCreated();
  };

  const ensureYoutubePlaybackRuntimeReady = async (): Promise<void> => {
    ensureOverlayStartupPrereqs();
    await input.runner.ensureYomitanExtensionLoaded();
    if (!input.isOverlayRuntimeInitialized()) {
      input.runner.initializeOverlayRuntime();
      return;
    }
    input.runner.ensureOverlayWindowsReadyForVisibilityActions();
  };

  const createImmersionTracker = input.runner.createImmersionTracker;
  const startJellyfinRemoteSession = input.runner.startJellyfinRemoteSession;
  const prewarmSubtitleDictionaries = input.runner.prewarmSubtitleDictionaries;
  const runHeadlessInitialCommand = input.runner.runHeadlessInitialCommand;

  const { reloadConfig, criticalConfigError, appReadyRuntimeRunner } = composeAppReadyRuntime({
    reloadConfigMainDeps: input.reload,
    criticalConfigErrorMainDeps: input.criticalConfig,
    immersionTrackerStartupMainDeps: input.immersion as never,
    appReadyRuntimeMainDeps: {
      ensureDefaultConfigBootstrap: () => input.runner.ensureDefaultConfigBootstrap(),
      loadSubtitlePosition: () => ensureSubtitlePositionLoaded(),
      resolveKeybindings: () => ensureKeybindingsResolved(),
      createMpvClient: () => ensureMpvClientCreated(),
      initRuntimeOptionsManager: () => ensureRuntimeOptionsManagerInitialized(),
      createSubtitleTimingTracker: () => ensureSubtitleTimingTrackerCreated(),
      getResolvedConfig: () => input.runner.getResolvedConfig() as never,
      getConfigWarnings: () => input.runner.getConfigWarnings(),
      logConfigWarning: (warning) => input.runner.logConfigWarning(warning),
      setLogLevel: (level, source) => input.runner.setLogLevel(level, source),
      setSecondarySubMode: (mode) => input.runner.setSecondarySubMode(mode),
      defaultSecondarySubMode: input.runner.defaultSecondarySubMode,
      defaultWebsocketPort: input.runner.defaultWebsocketPort,
      defaultAnnotationWebsocketPort: input.runner.defaultAnnotationWebsocketPort,
      defaultTexthookerPort: input.runner.defaultTexthookerPort,
      hasMpvWebsocketPlugin: () => input.runner.hasMpvWebsocketPlugin(),
      startSubtitleWebsocket: (port) => input.runner.startSubtitleWebsocket(port),
      startAnnotationWebsocket: (port) => input.runner.startAnnotationWebsocket(port),
      startTexthooker: (port, websocketUrl) => input.runner.startTexthooker(port, websocketUrl),
      log: (message) => input.runner.log(message),
      createMecabTokenizerAndCheck: () => input.runner.createMecabTokenizerAndCheck(),
      createImmersionTracker: createImmersionTracker ? () => createImmersionTracker() : undefined,
      startJellyfinRemoteSession: startJellyfinRemoteSession
        ? () => startJellyfinRemoteSession()
        : undefined,
      loadYomitanExtension: () => input.runner.loadYomitanExtension(),
      handleFirstRunSetup: () => input.runner.handleFirstRunSetup(),
      prewarmSubtitleDictionaries: prewarmSubtitleDictionaries
        ? () => prewarmSubtitleDictionaries()
        : undefined,
      startBackgroundWarmups: () => input.runner.startBackgroundWarmups(),
      texthookerOnlyMode: input.runner.texthookerOnlyMode,
      shouldAutoInitializeOverlayRuntimeFromConfig: () =>
        input.runner.shouldAutoInitializeOverlayRuntimeFromConfig(),
      setVisibleOverlayVisible: (visible) => input.runner.setVisibleOverlayVisible(visible),
      initializeOverlayRuntime: () => input.runner.initializeOverlayRuntime(),
      runHeadlessInitialCommand: runHeadlessInitialCommand
        ? () => runHeadlessInitialCommand()
        : undefined,
      handleInitialArgs: () => input.runner.handleInitialArgs(),
      logDebug: input.runner.logDebug,
      now: input.runner.now,
      shouldRunHeadlessInitialCommand: input.runner.shouldRunHeadlessInitialCommand,
      shouldUseMinimalStartup: input.runner.shouldUseMinimalStartup,
      shouldSkipHeavyStartup: input.runner.shouldSkipHeavyStartup,
    },
  });

  return {
    reloadConfig,
    criticalConfigError,
    ensureOverlayStartupPrereqs,
    ensureYoutubePlaybackRuntimeReady,
    runAppReady: async () => {
      await appReadyRuntimeRunner();
    },
  };
}
