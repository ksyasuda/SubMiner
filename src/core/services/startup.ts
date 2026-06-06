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
  setLogRotation?: (rotation: number) => void;
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
  } else if (initialArgs.background || initialArgs.update) {
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
  annotationWebsocket?: {
    enabled?: boolean;
    port?: number;
  };
  texthooker?: {
    launchAtStartup?: boolean;
  };
  secondarySub?: {
    defaultMode?: SecondarySubMode;
  };
  ankiConnect?: {
    enabled?: boolean;
    fields?: {
      audio?: string;
      image?: string;
      sentence?: string;
      miscInfo?: string;
      translation?: string;
    };
  };
  websocket?: {
    enabled?: boolean | 'auto';
    port?: number;
  };
  logging?: {
    level?: 'debug' | 'info' | 'warn' | 'error';
    rotation?: number;
    files?: {
      app?: boolean;
      launcher?: boolean;
      mpv?: boolean;
    };
  };
}

type TexthookerWebsocketConfigLike = Pick<AppReadyConfigLike, 'annotationWebsocket' | 'websocket'>;

type TexthookerWebsocketDefaults = {
  defaultWebsocketPort: number;
  defaultAnnotationWebsocketPort: number;
};

export interface AppReadyRuntimeDeps {
  ensureDefaultConfigBootstrap: () => void;
  loadSubtitlePosition: () => void;
  resolveKeybindings: () => void;
  createMpvClient: () => void;
  reloadConfig: () => void;
  getResolvedConfig: () => AppReadyConfigLike;
  getConfigWarnings: () => ConfigValidationWarning[];
  logConfigWarning: (warning: ConfigValidationWarning) => void;
  setLogLevel: (level: string, source: LogLevelSource) => void;
  setLogRotation?: (rotation: number) => void;
  setLogFileToggles?: (
    files: { app?: boolean; launcher?: boolean; mpv?: boolean } | undefined,
  ) => void;
  initRuntimeOptionsManager: () => void;
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
  createSubtitleTimingTracker: () => void;
  createImmersionTracker?: () => void;
  startJellyfinRemoteSession?: () => Promise<void>;
  loadYomitanExtension: () => Promise<void>;
  ensureYomitanExtensionLoaded?: () => Promise<void>;
  handleFirstRunSetup: () => Promise<void>;
  prewarmSubtitleDictionaries?: () => Promise<void>;
  startBackgroundWarmups: () => void;
  texthookerOnlyMode: boolean;
  shouldAutoInitializeOverlayRuntimeFromConfig: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  initializeOverlayRuntime: () => void;
  runHeadlessInitialCommand?: () => Promise<void>;
  handleInitialArgs: () => void;
  logDebug?: (message: string) => void;
  onCriticalConfigErrors?: (errors: string[]) => void;
  now?: () => number;
  shouldRunHeadlessInitialCommand?: () => boolean;
  shouldUseMinimalStartup?: () => boolean;
  shouldSkipHeavyStartup?: () => boolean;
  shouldHandleInitialArgsBeforeDeferredOverlayWarmup?: () => boolean;
}

const REQUIRED_ANKI_FIELD_MAPPING_KEYS = [
  'audio',
  'image',
  'sentence',
  'miscInfo',
  'translation',
] as const;

function getStartupCriticalConfigErrors(config: AppReadyConfigLike): string[] {
  if (!config.ankiConnect?.enabled) {
    return [];
  }

  const errors: string[] = [];
  const fields = config.ankiConnect.fields ?? {};

  for (const key of REQUIRED_ANKI_FIELD_MAPPING_KEYS) {
    const value = fields[key];
    if (typeof value !== 'string' || value.trim().length === 0) {
      errors.push(
        `ankiConnect.fields.${key} must be a non-empty string when ankiConnect is enabled.`,
      );
    }
  }

  return errors;
}

export function resolveTexthookerWebsocketUrl(
  config: TexthookerWebsocketConfigLike,
  defaults: TexthookerWebsocketDefaults,
  hasMpvWebsocketPlugin: boolean,
): string | undefined {
  const wsConfig = config.websocket || {};
  const wsEnabled = wsConfig.enabled ?? 'auto';
  const wsPort = wsConfig.port || defaults.defaultWebsocketPort;
  const annotationWsConfig = config.annotationWebsocket || {};
  const annotationWsEnabled = annotationWsConfig.enabled !== false;
  const annotationWsPort = annotationWsConfig.port || defaults.defaultAnnotationWebsocketPort;

  if (annotationWsEnabled) {
    return `ws://127.0.0.1:${annotationWsPort}`;
  }

  if (wsEnabled === true || (wsEnabled === 'auto' && !hasMpvWebsocketPlugin)) {
    return `ws://127.0.0.1:${wsPort}`;
  }

  return undefined;
}

export function shouldAutoInitializeOverlayRuntimeFromConfig(config: RuntimeConfigLike): boolean {
  return config.auto_start_overlay === true;
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
  const ensureYomitanExtensionReady =
    deps.ensureYomitanExtensionLoaded ?? deps.loadYomitanExtension;
  let firstRunSetupHandled = false;
  let initialArgsHandled = false;
  const handleFirstRunSetupOnce = async (): Promise<void> => {
    if (firstRunSetupHandled) {
      return;
    }
    firstRunSetupHandled = true;
    await deps.handleFirstRunSetup();
  };
  const handleInitialArgsOnce = (): void => {
    if (initialArgsHandled) {
      return;
    }
    initialArgsHandled = true;
    deps.handleInitialArgs();
  };

  deps.ensureDefaultConfigBootstrap();
  if (deps.shouldRunHeadlessInitialCommand?.()) {
    deps.reloadConfig();
    deps.initRuntimeOptionsManager();
    if (deps.runHeadlessInitialCommand) {
      await deps.runHeadlessInitialCommand();
    } else {
      deps.createMpvClient();
      deps.createSubtitleTimingTracker();
      await ensureYomitanExtensionReady();
      deps.initializeOverlayRuntime();
      deps.handleInitialArgs();
    }
    return;
  }

  if (deps.shouldUseMinimalStartup?.()) {
    deps.reloadConfig();
    handleInitialArgsOnce();
    return;
  }

  deps.logDebug?.('App-ready critical path started.');

  if (deps.shouldSkipHeavyStartup?.()) {
    await ensureYomitanExtensionReady();
    deps.reloadConfig();
    await handleFirstRunSetupOnce();
    handleInitialArgsOnce();
    deps.logDebug?.(`App-ready critical path finished in ${now() - startupStartedAtMs}ms.`);
    return;
  }

  deps.reloadConfig();
  const config = deps.getResolvedConfig();
  const criticalConfigErrors = getStartupCriticalConfigErrors(config);
  if (criticalConfigErrors.length > 0) {
    deps.onCriticalConfigErrors?.(criticalConfigErrors);
    deps.logDebug?.(
      `App-ready critical path aborted after config validation in ${now() - startupStartedAtMs}ms.`,
    );
    return;
  }

  deps.setLogLevel(config.logging?.level ?? 'info', 'config');
  deps.setLogRotation?.(config.logging?.rotation ?? 7);
  deps.setLogFileToggles?.(config.logging?.files);
  for (const warning of deps.getConfigWarnings()) {
    deps.logConfigWarning(warning);
  }
  deps.startBackgroundWarmups();

  deps.loadSubtitlePosition();
  deps.resolveKeybindings();
  deps.createMpvClient();
  deps.initRuntimeOptionsManager();
  deps.setSecondarySubMode(config.secondarySub?.defaultMode ?? deps.defaultSecondarySubMode);

  const wsConfig = config.websocket || {};
  const wsEnabled = wsConfig.enabled ?? 'auto';
  const wsPort = wsConfig.port || deps.defaultWebsocketPort;
  const annotationWsConfig = config.annotationWebsocket || {};
  const annotationWsEnabled = annotationWsConfig.enabled !== false;
  const annotationWsPort = annotationWsConfig.port || deps.defaultAnnotationWebsocketPort;
  const texthookerPort = deps.defaultTexthookerPort;
  const texthookerWebsocketUrl = resolveTexthookerWebsocketUrl(
    config,
    {
      defaultWebsocketPort: deps.defaultWebsocketPort,
      defaultAnnotationWebsocketPort: deps.defaultAnnotationWebsocketPort,
    },
    deps.hasMpvWebsocketPlugin(),
  );

  if (wsEnabled === true || (wsEnabled === 'auto' && !deps.hasMpvWebsocketPlugin())) {
    deps.startSubtitleWebsocket(wsPort);
  } else if (wsEnabled === 'auto') {
    deps.log('mpv_websocket detected, skipping built-in WebSocket server');
  }

  if (annotationWsEnabled) {
    deps.startAnnotationWebsocket(annotationWsPort);
  }

  if (config.texthooker?.launchAtStartup !== false) {
    deps.startTexthooker(texthookerPort, texthookerWebsocketUrl);
  }

  deps.createSubtitleTimingTracker();
  if (deps.createImmersionTracker) {
    deps.log('Runtime ready: immersion tracker startup requested.');
  } else {
    deps.log('Runtime ready: immersion tracker dependency is missing.');
  }

  if (deps.texthookerOnlyMode) {
    deps.log('Texthooker-only mode enabled; skipping overlay window.');
  } else if (deps.shouldAutoInitializeOverlayRuntimeFromConfig()) {
    await ensureYomitanExtensionReady();
    deps.setVisibleOverlayVisible(true);
    deps.initializeOverlayRuntime();
  } else {
    deps.log('Overlay runtime deferred: waiting for explicit overlay command.');
    if (deps.shouldHandleInitialArgsBeforeDeferredOverlayWarmup?.()) {
      await handleFirstRunSetupOnce();
      handleInitialArgsOnce();
    } else {
      await ensureYomitanExtensionReady();
    }
  }

  await handleFirstRunSetupOnce();
  handleInitialArgsOnce();
  deps.logDebug?.(`App-ready critical path finished in ${now() - startupStartedAtMs}ms.`);
}
