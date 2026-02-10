import { ConfigValidationWarning, SecondarySubMode } from "../../types";

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

  if (wsEnabled === true || (wsEnabled === "auto" && !deps.hasMpvWebsocketPlugin())) {
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
