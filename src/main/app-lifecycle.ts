import type { CliArgs, CliCommandSource } from '../cli/args';
import { runAppReadyRuntime } from '../core/services/startup';
import type { AppReadyRuntimeDeps } from '../core/services/startup';
import type { AppLifecycleDepsRuntimeOptions } from '../core/services/app-lifecycle';

export interface AppLifecycleRuntimeDepsFactoryInput {
  app: AppLifecycleDepsRuntimeOptions['app'];
  platform: NodeJS.Platform;
  shouldStartApp: (args: CliArgs) => boolean;
  parseArgs: (argv: string[]) => CliArgs;
  handleCliCommand: (nextArgs: CliArgs, source: CliCommandSource) => void;
  printHelp: () => void;
  logNoRunningInstance: () => void;
  onReady: () => Promise<void>;
  onWillQuitCleanup: () => void;
  shouldRestoreWindowsOnActivate: () => boolean;
  restoreWindowsOnActivate: () => void;
  shouldQuitOnWindowAllClosed: () => boolean;
}

export interface AppReadyRuntimeDepsFactoryInput {
  loadSubtitlePosition: AppReadyRuntimeDeps['loadSubtitlePosition'];
  resolveKeybindings: AppReadyRuntimeDeps['resolveKeybindings'];
  createMpvClient: AppReadyRuntimeDeps['createMpvClient'];
  reloadConfig: AppReadyRuntimeDeps['reloadConfig'];
  getResolvedConfig: AppReadyRuntimeDeps['getResolvedConfig'];
  getConfigWarnings: AppReadyRuntimeDeps['getConfigWarnings'];
  logConfigWarning: AppReadyRuntimeDeps['logConfigWarning'];
  initRuntimeOptionsManager: AppReadyRuntimeDeps['initRuntimeOptionsManager'];
  setSecondarySubMode: AppReadyRuntimeDeps['setSecondarySubMode'];
  defaultSecondarySubMode: AppReadyRuntimeDeps['defaultSecondarySubMode'];
  defaultWebsocketPort: AppReadyRuntimeDeps['defaultWebsocketPort'];
  hasMpvWebsocketPlugin: AppReadyRuntimeDeps['hasMpvWebsocketPlugin'];
  startSubtitleWebsocket: AppReadyRuntimeDeps['startSubtitleWebsocket'];
  log: AppReadyRuntimeDeps['log'];
  setLogLevel: AppReadyRuntimeDeps['setLogLevel'];
  createMecabTokenizerAndCheck: AppReadyRuntimeDeps['createMecabTokenizerAndCheck'];
  createSubtitleTimingTracker: AppReadyRuntimeDeps['createSubtitleTimingTracker'];
  createImmersionTracker?: AppReadyRuntimeDeps['createImmersionTracker'];
  startJellyfinRemoteSession?: AppReadyRuntimeDeps['startJellyfinRemoteSession'];
  loadYomitanExtension: AppReadyRuntimeDeps['loadYomitanExtension'];
  prewarmSubtitleDictionaries?: AppReadyRuntimeDeps['prewarmSubtitleDictionaries'];
  startBackgroundWarmups: AppReadyRuntimeDeps['startBackgroundWarmups'];
  texthookerOnlyMode: AppReadyRuntimeDeps['texthookerOnlyMode'];
  shouldAutoInitializeOverlayRuntimeFromConfig: AppReadyRuntimeDeps['shouldAutoInitializeOverlayRuntimeFromConfig'];
  setVisibleOverlayVisible: AppReadyRuntimeDeps['setVisibleOverlayVisible'];
  initializeOverlayRuntime: AppReadyRuntimeDeps['initializeOverlayRuntime'];
  handleInitialArgs: AppReadyRuntimeDeps['handleInitialArgs'];
  onCriticalConfigErrors?: AppReadyRuntimeDeps['onCriticalConfigErrors'];
  logDebug?: AppReadyRuntimeDeps['logDebug'];
  now?: AppReadyRuntimeDeps['now'];
  shouldSkipHeavyStartup?: AppReadyRuntimeDeps['shouldSkipHeavyStartup'];
}

export function createAppLifecycleRuntimeDeps(
  params: AppLifecycleRuntimeDepsFactoryInput,
): AppLifecycleDepsRuntimeOptions {
  return {
    app: params.app,
    platform: params.platform,
    shouldStartApp: params.shouldStartApp,
    parseArgs: params.parseArgs,
    handleCliCommand: params.handleCliCommand,
    printHelp: params.printHelp,
    logNoRunningInstance: params.logNoRunningInstance,
    onReady: params.onReady,
    onWillQuitCleanup: params.onWillQuitCleanup,
    shouldRestoreWindowsOnActivate: params.shouldRestoreWindowsOnActivate,
    restoreWindowsOnActivate: params.restoreWindowsOnActivate,
    shouldQuitOnWindowAllClosed: params.shouldQuitOnWindowAllClosed,
  };
}

export function createAppReadyRuntimeDeps(
  params: AppReadyRuntimeDepsFactoryInput,
): AppReadyRuntimeDeps {
  return {
    loadSubtitlePosition: params.loadSubtitlePosition,
    resolveKeybindings: params.resolveKeybindings,
    createMpvClient: params.createMpvClient,
    reloadConfig: params.reloadConfig,
    getResolvedConfig: params.getResolvedConfig,
    getConfigWarnings: params.getConfigWarnings,
    logConfigWarning: params.logConfigWarning,
    initRuntimeOptionsManager: params.initRuntimeOptionsManager,
    setSecondarySubMode: params.setSecondarySubMode,
    defaultSecondarySubMode: params.defaultSecondarySubMode,
    defaultWebsocketPort: params.defaultWebsocketPort,
    hasMpvWebsocketPlugin: params.hasMpvWebsocketPlugin,
    startSubtitleWebsocket: params.startSubtitleWebsocket,
    log: params.log,
    setLogLevel: params.setLogLevel,
    createMecabTokenizerAndCheck: params.createMecabTokenizerAndCheck,
    createSubtitleTimingTracker: params.createSubtitleTimingTracker,
    createImmersionTracker: params.createImmersionTracker,
    startJellyfinRemoteSession: params.startJellyfinRemoteSession,
    loadYomitanExtension: params.loadYomitanExtension,
    prewarmSubtitleDictionaries: params.prewarmSubtitleDictionaries,
    startBackgroundWarmups: params.startBackgroundWarmups,
    texthookerOnlyMode: params.texthookerOnlyMode,
    shouldAutoInitializeOverlayRuntimeFromConfig:
      params.shouldAutoInitializeOverlayRuntimeFromConfig,
    setVisibleOverlayVisible: params.setVisibleOverlayVisible,
    initializeOverlayRuntime: params.initializeOverlayRuntime,
    handleInitialArgs: params.handleInitialArgs,
    onCriticalConfigErrors: params.onCriticalConfigErrors,
    logDebug: params.logDebug,
    now: params.now,
    shouldSkipHeavyStartup: params.shouldSkipHeavyStartup,
  };
}

export function createAppReadyRuntimeRunner(
  params: AppReadyRuntimeDepsFactoryInput,
): () => Promise<void> {
  return async () => {
    await runAppReadyRuntime(createAppReadyRuntimeDeps(params));
  };
}
