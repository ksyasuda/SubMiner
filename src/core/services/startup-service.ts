import { CliArgs } from "../../cli/args";
import { ConfigValidationWarning, SecondarySubMode } from "../../types";

export interface StartupBootstrapRuntimeState {
  initialArgs: CliArgs;
  mpvSocketPath: string;
  texthookerPort: number;
  backendOverride: string | null;
  autoStartOverlay: boolean;
  texthookerOnlyMode: boolean;
}

export interface StartupBootstrapRuntimeDeps {
  argv: string[];
  parseArgs: (argv: string[]) => CliArgs;
  setLogLevelEnv: (level: string) => void;
  enableVerboseLogging: () => void;
  forceX11Backend: (args: CliArgs) => void;
  enforceUnsupportedWaylandMode: (args: CliArgs) => void;
  getDefaultSocketPath: () => string;
  defaultTexthookerPort: number;
  runGenerateConfigFlow: (args: CliArgs) => boolean;
  startAppLifecycle: (args: CliArgs) => void;
}

export function runStartupBootstrapRuntimeService(
  deps: StartupBootstrapRuntimeDeps,
): StartupBootstrapRuntimeState {
  const initialArgs = deps.parseArgs(deps.argv);

  if (initialArgs.logLevel) {
    deps.setLogLevelEnv(initialArgs.logLevel);
  } else if (initialArgs.verbose) {
    deps.enableVerboseLogging();
  }

  deps.forceX11Backend(initialArgs);
  deps.enforceUnsupportedWaylandMode(initialArgs);

  const state: StartupBootstrapRuntimeState = {
    initialArgs,
    mpvSocketPath: initialArgs.socketPath ?? deps.getDefaultSocketPath(),
    texthookerPort: initialArgs.texthookerPort ?? deps.defaultTexthookerPort,
    backendOverride: initialArgs.backend ?? null,
    autoStartOverlay: initialArgs.autoStartOverlay,
    texthookerOnlyMode: initialArgs.texthooker,
  };

  if (!deps.runGenerateConfigFlow(initialArgs)) {
    deps.startAppLifecycle(initialArgs);
  }

  return state;
}

interface AppReadyConfigLike {
  secondarySub?: {
    defaultMode?: SecondarySubMode;
  };
  websocket?: {
    enabled?: boolean | "auto";
    port?: number;
  };
}

export interface AppReadyRuntimeDeps {
  loadSubtitlePosition: () => void;
  resolveKeybindings: () => void;
  createMpvClient: () => void;
  reloadConfig: () => void;
  getResolvedConfig: () => AppReadyConfigLike;
  getConfigWarnings: () => ConfigValidationWarning[];
  logConfigWarning: (warning: ConfigValidationWarning) => void;
  initRuntimeOptionsManager: () => void;
  setSecondarySubMode: (mode: SecondarySubMode) => void;
  defaultSecondarySubMode: SecondarySubMode;
  defaultWebsocketPort: number;
  hasMpvWebsocketPlugin: () => boolean;
  startSubtitleWebsocket: (port: number) => void;
  log: (message: string) => void;
  createMecabTokenizerAndCheck: () => Promise<void>;
  createSubtitleTimingTracker: () => void;
  loadYomitanExtension: () => Promise<void>;
  texthookerOnlyMode: boolean;
  shouldAutoInitializeOverlayRuntimeFromConfig: () => boolean;
  initializeOverlayRuntime: () => void;
  handleInitialArgs: () => void;
}

export async function runAppReadyRuntimeService(
  deps: AppReadyRuntimeDeps,
): Promise<void> {
  deps.loadSubtitlePosition();
  deps.resolveKeybindings();
  deps.createMpvClient();

  deps.reloadConfig();
  const config = deps.getResolvedConfig();
  for (const warning of deps.getConfigWarnings()) {
    deps.logConfigWarning(warning);
  }
  deps.initRuntimeOptionsManager();
  deps.setSecondarySubMode(
    config.secondarySub?.defaultMode ?? deps.defaultSecondarySubMode,
  );

  const wsConfig = config.websocket || {};
  const wsEnabled = wsConfig.enabled ?? "auto";
  const wsPort = wsConfig.port || deps.defaultWebsocketPort;

  if (
    wsEnabled === true ||
    (wsEnabled === "auto" && !deps.hasMpvWebsocketPlugin())
  ) {
    deps.startSubtitleWebsocket(wsPort);
  } else if (wsEnabled === "auto") {
    deps.log("mpv_websocket detected, skipping built-in WebSocket server");
  }

  await deps.createMecabTokenizerAndCheck();
  deps.createSubtitleTimingTracker();
  await deps.loadYomitanExtension();

  if (deps.texthookerOnlyMode) {
    deps.log("Texthooker-only mode enabled; skipping overlay window.");
  } else if (deps.shouldAutoInitializeOverlayRuntimeFromConfig()) {
    deps.initializeOverlayRuntime();
  } else {
    deps.log("Overlay runtime deferred: waiting for explicit overlay command.");
  }

  deps.handleInitialArgs();
}
