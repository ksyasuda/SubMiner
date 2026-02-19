import { CliArgs } from '../../cli/args';
import type { LogLevelSource } from '../../logger';
import { ConfigValidationWarning, ResolvedConfig, SecondarySubMode } from '../../types';

export interface StartupBootstrapRuntimeState {
  initialArgs: CliArgs;
  mpvSocketPath: string;
  texthookerPort: number;
  backendOverride: string | null;
  autoStartOverlay: boolean;
  texthookerOnlyMode: boolean;
  backgroundMode: boolean;
}

interface RuntimeAutoUpdateOptionManagerLike {
  getOptionValue: (id: 'anki.autoUpdateNewCards') => unknown;
}

export interface RuntimeConfigLike {
  auto_start_overlay?: boolean;
  bind_visible_overlay_to_mpv_sub_visibility: boolean;
  invisibleOverlay: {
    startupVisibility: 'visible' | 'hidden' | 'platform-default';
  };
  ankiConnect?: {
    behavior?: {
      autoUpdateNewCards?: boolean;
    };
  };
}

export interface StartupBootstrapRuntimeDeps {
  argv: string[];
  parseArgs: (argv: string[]) => CliArgs;
  setLogLevel: (level: string, source: LogLevelSource) => void;
  forceX11Backend: (args: CliArgs) => void;
  enforceUnsupportedWaylandMode: (args: CliArgs) => void;
  getDefaultSocketPath: () => string;
  defaultTexthookerPort: number;
  runGenerateConfigFlow: (args: CliArgs) => boolean;
  startAppLifecycle: (args: CliArgs) => void;
}

export function runStartupBootstrapRuntime(
  deps: StartupBootstrapRuntimeDeps,
): StartupBootstrapRuntimeState {
  const initialArgs = deps.parseArgs(deps.argv);

  if (initialArgs.logLevel) {
    deps.setLogLevel(initialArgs.logLevel, 'cli');
  } else if (initialArgs.background) {
    deps.setLogLevel('warn', 'cli');
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
    backgroundMode: initialArgs.background,
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
    enabled?: boolean | 'auto';
    port?: number;
  };
  logging?: {
    level?: 'debug' | 'info' | 'warn' | 'error';
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
  setLogLevel: (level: string, source: LogLevelSource) => void;
  initRuntimeOptionsManager: () => void;
  setSecondarySubMode: (mode: SecondarySubMode) => void;
  defaultSecondarySubMode: SecondarySubMode;
  defaultWebsocketPort: number;
  hasMpvWebsocketPlugin: () => boolean;
  startSubtitleWebsocket: (port: number) => void;
  log: (message: string) => void;
  createMecabTokenizerAndCheck: () => Promise<void>;
  createSubtitleTimingTracker: () => void;
  createImmersionTracker?: () => void;
  startJellyfinRemoteSession?: () => Promise<void>;
  loadYomitanExtension: () => Promise<void>;
  prewarmSubtitleDictionaries?: () => Promise<void>;
  startBackgroundWarmups: () => void;
  texthookerOnlyMode: boolean;
  shouldAutoInitializeOverlayRuntimeFromConfig: () => boolean;
  initializeOverlayRuntime: () => void;
  handleInitialArgs: () => void;
  logDebug?: (message: string) => void;
  now?: () => number;
}

export function getInitialInvisibleOverlayVisibility(
  config: RuntimeConfigLike,
  platform: NodeJS.Platform,
): boolean {
  const visibility = config.invisibleOverlay.startupVisibility;
  if (visibility === 'visible') return true;
  if (visibility === 'hidden') return false;
  if (platform === 'linux') return false;
  return true;
}

export function shouldAutoInitializeOverlayRuntimeFromConfig(config: RuntimeConfigLike): boolean {
  if (config.auto_start_overlay === true) return true;
  if (config.invisibleOverlay.startupVisibility === 'visible') return true;
  return false;
}

export function shouldBindVisibleOverlayToMpvSubVisibility(config: RuntimeConfigLike): boolean {
  return config.bind_visible_overlay_to_mpv_sub_visibility;
}

export function isAutoUpdateEnabledRuntime(
  config: ResolvedConfig | RuntimeConfigLike,
  runtimeOptionsManager: RuntimeAutoUpdateOptionManagerLike | null,
): boolean {
  const value = runtimeOptionsManager?.getOptionValue('anki.autoUpdateNewCards');
  if (typeof value === 'boolean') return value;
  return (config as ResolvedConfig).ankiConnect?.behavior?.autoUpdateNewCards !== false;
}

export async function runAppReadyRuntime(deps: AppReadyRuntimeDeps): Promise<void> {
  const now = deps.now ?? (() => Date.now());
  const startupStartedAtMs = now();
  deps.logDebug?.('App-ready critical path started.');

  deps.loadSubtitlePosition();
  deps.resolveKeybindings();
  deps.createMpvClient();

  deps.reloadConfig();
  const config = deps.getResolvedConfig();
  deps.setLogLevel(config.logging?.level ?? 'info', 'config');
  for (const warning of deps.getConfigWarnings()) {
    deps.logConfigWarning(warning);
  }
  deps.initRuntimeOptionsManager();
  deps.setSecondarySubMode(config.secondarySub?.defaultMode ?? deps.defaultSecondarySubMode);

  const wsConfig = config.websocket || {};
  const wsEnabled = wsConfig.enabled ?? 'auto';
  const wsPort = wsConfig.port || deps.defaultWebsocketPort;

  if (wsEnabled === true || (wsEnabled === 'auto' && !deps.hasMpvWebsocketPlugin())) {
    deps.startSubtitleWebsocket(wsPort);
  } else if (wsEnabled === 'auto') {
    deps.log('mpv_websocket detected, skipping built-in WebSocket server');
  }

  deps.createSubtitleTimingTracker();
  if (deps.createImmersionTracker) {
    deps.log('Runtime ready: invoking createImmersionTracker.');
    try {
      deps.createImmersionTracker();
    } catch (error) {
      deps.log(`Runtime ready: createImmersionTracker failed: ${(error as Error).message}`);
    }
  } else {
    deps.log('Runtime ready: createImmersionTracker dependency is missing.');
  }

  if (deps.texthookerOnlyMode) {
    deps.log('Texthooker-only mode enabled; skipping overlay window.');
  } else if (deps.shouldAutoInitializeOverlayRuntimeFromConfig()) {
    deps.initializeOverlayRuntime();
  } else {
    deps.log('Overlay runtime deferred: waiting for explicit overlay command.');
  }

  deps.handleInitialArgs();
  deps.startBackgroundWarmups();
  deps.logDebug?.(`App-ready critical path finished in ${now() - startupStartedAtMs}ms.`);
}
