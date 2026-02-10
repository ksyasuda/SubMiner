import {
  AppReadyRuntimeDeps,
} from "./app-ready-runtime-service";
import {
  AppShutdownRuntimeDeps,
} from "./app-shutdown-runtime-service";

export type StartupAppReadyDepsRuntimeOptions = AppReadyRuntimeDeps;
export type StartupAppShutdownDepsRuntimeOptions = AppShutdownRuntimeDeps;

export function createStartupAppReadyDepsRuntimeService(
  options: StartupAppReadyDepsRuntimeOptions,
): AppReadyRuntimeDeps {
  return {
    loadSubtitlePosition: options.loadSubtitlePosition,
    resolveKeybindings: options.resolveKeybindings,
    createMpvClient: options.createMpvClient,
    reloadConfig: options.reloadConfig,
    getResolvedConfig: options.getResolvedConfig,
    getConfigWarnings: options.getConfigWarnings,
    logConfigWarning: options.logConfigWarning,
    initRuntimeOptionsManager: options.initRuntimeOptionsManager,
    setSecondarySubMode: options.setSecondarySubMode,
    defaultSecondarySubMode: options.defaultSecondarySubMode,
    defaultWebsocketPort: options.defaultWebsocketPort,
    hasMpvWebsocketPlugin: options.hasMpvWebsocketPlugin,
    startSubtitleWebsocket: options.startSubtitleWebsocket,
    log: options.log,
    createMecabTokenizerAndCheck: options.createMecabTokenizerAndCheck,
    createSubtitleTimingTracker: options.createSubtitleTimingTracker,
    loadYomitanExtension: options.loadYomitanExtension,
    texthookerOnlyMode: options.texthookerOnlyMode,
    shouldAutoInitializeOverlayRuntimeFromConfig:
      options.shouldAutoInitializeOverlayRuntimeFromConfig,
    initializeOverlayRuntime: options.initializeOverlayRuntime,
    handleInitialArgs: options.handleInitialArgs,
  };
}

export function createStartupAppShutdownDepsRuntimeService(
  options: StartupAppShutdownDepsRuntimeOptions,
): AppShutdownRuntimeDeps {
  return {
    unregisterAllGlobalShortcuts: options.unregisterAllGlobalShortcuts,
    stopSubtitleWebsocket: options.stopSubtitleWebsocket,
    stopTexthookerService: options.stopTexthookerService,
    destroyYomitanParserWindow: options.destroyYomitanParserWindow,
    clearYomitanParserPromises: options.clearYomitanParserPromises,
    stopWindowTracker: options.stopWindowTracker,
    destroyMpvSocket: options.destroyMpvSocket,
    clearReconnectTimer: options.clearReconnectTimer,
    destroySubtitleTimingTracker: options.destroySubtitleTimingTracker,
    destroyAnkiIntegration: options.destroyAnkiIntegration,
  };
}
