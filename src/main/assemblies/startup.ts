import path from 'node:path';
import { AnkiIntegration } from '../../anki-integration';
import { isHeadlessInitialCommand } from '../../cli/args';
import { DEFAULT_CONFIG, DEFAULT_KEYBINDINGS, generateConfigTemplate } from '../../config';
import type { ConfigService } from '../../config/service';
import {
  initializeOverlayRuntime as initializeOverlayRuntimeCore,
  type createOverlayManager,
  type SubtitleWebSocket,
  type Texthooker,
} from '../../core/services';
import { resolveKeybindings, showDesktopNotification } from '../../core/utils';
import type { Logger } from '../../logger';
import { RuntimeOptionsManager } from '../../runtime-options';
import { ensureDefaultConfigBootstrap, getDefaultConfigFilePaths } from '../../shared/setup-state';
import { SubtitleTimingTracker } from '../../subtitle-timing-tracker';
import type { SecondarySubMode, WindowGeometry } from '../../types';
import { mergeAiConfig } from '../../ai/config';
import type { OverlayModalRuntime } from '../overlay-runtime';
import {
  composeAppReadyRuntime,
  type AppReadyComposerOptions,
  type composeOverlayVisibilityRuntime,
} from '../runtime/composers';
import type { createConfigDerivedRuntime } from '../runtime/domains/startup';
import { shouldAutoOpenFirstRunSetup } from '../runtime/first-run-setup-service';
import type { createOverlayGeometryRuntime } from '../runtime/overlay-geometry-runtime';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import { createOverlayRuntimeBootstrapHandlers } from '../runtime/overlay-runtime-bootstrap-handlers';
import {
  getStartupModeFlags,
  shouldHandleInitialArgsBeforeDeferredOverlayWarmup,
  shouldRefreshAnilistOnConfigReload,
} from '../runtime/startup-mode-flags';
import type { VisibleOverlayInteractionRuntime } from '../runtime/visible-overlay-interaction-runtime';
import type { AppState } from '../state';
import type { AnilistAssembly } from './anilist';
import type { FirstRunSetupAssembly } from './first-run-setup';
import type { JellyfinAssembly } from './jellyfin';
import type { MiningAssembly } from './mining';
import type { OverlayAssembly } from './overlay';
import type { PlaybackAssembly } from './playback';
import type { SettingsAssembly } from './settings';
import type { ShortcutsAssembly } from './shortcuts';
import type { StatsAssembly } from './stats';
import type { YomitanAssembly } from './yomitan';
import type { YoutubeAssembly } from './youtube';

type AppReadyMainDeps = AppReadyComposerOptions['appReadyRuntimeMainDeps'];
type OverlayManager = ReturnType<typeof createOverlayManager>;
type FatalErrorReporter = (error: unknown, options: { title: string; context: string }) => void;

/**
 * What startup needs from the rest of the app. It touches nearly every domain, so most deps are
 * sibling assemblies; those built after this one are reached through getters.
 */
export interface StartupAssemblyDeps {
  appState: AppState;
  configService: Pick<
    ConfigService,
    'getConfig' | 'getConfigPath' | 'getWarnings' | 'reloadConfigStrict'
  >;
  logger: Pick<Logger, 'debug' | 'info' | 'warn' | 'error'>;
  logConfigWarning: AppReadyMainDeps['logConfigWarning'];
  configDir: string;
  userDataPath: string;
  defaultTexthookerPort: number;
  /** Electron's `dialog.showErrorBox`, injected so the assembly loads under `bun test`. */
  showErrorBox: (title: string, content: string) => void;
  reportFatalError: FatalErrorReporter;
  requestAppQuit: () => void;
  logging: Pick<AppReadyMainDeps, 'setLogLevel' | 'setLogRotation' | 'setLogFileToggles'>;
  overlayManager: Pick<OverlayManager, 'getVisibleOverlayVisible' | 'getOverlayWindows'>;
  overlayModalRuntime: Pick<OverlayModalRuntime, 'primeModalWindow'>;
  subtitleWsService: Pick<SubtitleWebSocket, 'start'>;
  annotationSubtitleWsService: Pick<SubtitleWebSocket, 'start'>;
  texthookerService: Pick<Texthooker, 'isRunning' | 'start'>;
  configDerivedRuntime: Pick<
    ReturnType<typeof createConfigDerivedRuntime>,
    'shouldAutoInitializeOverlayRuntimeFromConfig'
  >;
  hasMpvWebsocketPlugin: () => boolean;
  loadSubtitlePosition: () => void;
  setSecondarySubMode: (mode: SecondarySubMode) => void;
  syncOverlayMpvSubtitleSuppression: () => void;
  yomitan: Pick<YomitanAssembly, 'loadExtension' | 'ensureExtensionLoaded'>;
  stats: Pick<StatsAssembly, 'ensureImmersionTrackerStarted' | 'immersionTrackerStartupDeps'>;
  firstRunSetup: Pick<FirstRunSetupAssembly, 'service' | 'openWindow'>;
  youtube: Pick<
    YoutubeAssembly,
    | 'getCachedMediaPathForCurrentPlayback'
    | 'shouldRequireMediaCacheForCurrentPlayback'
    | 'getCurrentMediaCacheSourceUrl'
  >;
  overlay: Pick<
    OverlayAssembly,
    | 'overlayVisibilityRuntime'
    | 'createMainWindow'
    | 'setVisibleOverlayVisible'
    | 'ensureOverlayWindowsReadyForVisibilityActions'
  >;
  shortcuts: Pick<
    ShortcutsAssembly,
    | 'overlayShortcutsRuntime'
    | 'refreshCurrentSessionBindings'
    | 'registerGlobalShortcuts'
    | 'refreshOverlayShortcuts'
  >;
  settings: { hotReload: Pick<SettingsAssembly['hotReload'], 'start'> };
  mining: Pick<MiningAssembly, 'attachToAnkiIntegration' | 'createFieldGroupingCallback'>;
  getAnilist: () => Pick<AnilistAssembly, 'refreshClientSecretStateIfEnabled'>;
  getJellyfin: () => Pick<JellyfinAssembly, 'startJellyfinRemoteSession'>;
  getPlayback: () => Pick<
    PlaybackAssembly,
    | 'createMpvClient'
    | 'createMecabTokenizerAndCheck'
    | 'prewarmSubtitleDictionaries'
    | 'startBackgroundWarmups'
    | 'subtitleProcessingController'
    | 'getSubtitlePrefetchService'
    | 'getLastObservedTimePos'
  >;
  getPluginHost: () => { loadAll: () => Promise<unknown> };
  getOverlayNotifications: () => Pick<
    ReturnType<typeof createOverlayNotificationsRuntime>,
    'showOverlayNotification' | 'dismissOverlayNotification'
  >;
  getOverlayVisibilityComposer: () => Pick<
    ReturnType<typeof composeOverlayVisibilityRuntime>,
    'broadcastRuntimeOptionsChanged'
  >;
  /** Pointer interaction and geometry stay in `main.ts`: both import Electron at module load. */
  getVisibleOverlayInteraction: () => Pick<
    VisibleOverlayInteractionRuntime,
    'createOverlayWindowTracker' | 'bindVisibleOverlayOwner' | 'releaseVisibleOverlayOwner'
  >;
  getOverlayGeometry: () => Pick<
    ReturnType<typeof createOverlayGeometryRuntime>,
    'updateVisibleOverlayBounds'
  >;
  /** Runs the CLI command the app was launched with. */
  handleInitialArgs: () => void;
}

export type StartupAssembly = ReturnType<typeof createStartupAssembly>;

/**
 * Owns how the app starts: the app-ready sequence, the headless known-word refresh, and bringing
 * up the overlay runtime and its prerequisites on demand (CLI commands, YouTube playback, the
 * tray). `main.ts` hands `runAppReady` to the app lifecycle.
 */
export function createStartupAssembly(deps: StartupAssemblyDeps) {
  const { appState, configService, logger } = deps;
  const isHeadlessLaunch = (): boolean =>
    Boolean(appState.initialArgs && isHeadlessInitialCommand(appState.initialArgs));
  const knownWordCacheStatePath = path.join(deps.userDataPath, 'known-words-cache.json');

  function createRuntimeOptionsManager(): RuntimeOptionsManager {
    return new RuntimeOptionsManager(() => configService.getConfig().ankiConnect, {
      applyAnkiPatch: (patch) => {
        if (appState.ankiIntegration) {
          appState.ankiIntegration.applyRuntimeConfigPatch(patch);
        }
      },
      getSubtitleStyleConfig: () => configService.getConfig().subtitleStyle,
      onOptionsChanged: () => {
        const playback = deps.getPlayback();
        playback.subtitleProcessingController.invalidateTokenizationCache();
        playback.getSubtitlePrefetchService()?.onSeek(playback.getLastObservedTimePos());
        deps.getOverlayVisibilityComposer().broadcastRuntimeOptionsChanged();
        deps.shortcuts.refreshOverlayShortcuts();
      },
    });
  }

  function resolveSessionKeybindings(): void {
    appState.keybindings = resolveKeybindings(configService.getConfig(), DEFAULT_KEYBINDINGS);
    deps.shortcuts.refreshCurrentSessionBindings();
  }

  /** What the subtitle websockets send a client that connects mid-line. */
  function getCurrentSubtitleForWebsocket() {
    return (
      appState.currentSubtitleData ??
      (appState.currentSubText
        ? {
            text: appState.currentSubText,
            tokens: null,
            startTime: appState.mpvClient?.currentSubStart ?? null,
            endTime: appState.mpvClient?.currentSubEnd ?? null,
          }
        : null)
    );
  }
  function getFrequencyOptionsForWebsocket() {
    const frequencyDictionary = configService.getConfig().subtitleStyle.frequencyDictionary;
    return {
      enabled: frequencyDictionary.enabled,
      topX: frequencyDictionary.topX,
      mode: frequencyDictionary.mode,
    };
  }

  const { initializeOverlayRuntime: initializeOverlayRuntimeHandler } =
    createOverlayRuntimeBootstrapHandlers({
      initializeOverlayRuntimeMainDeps: {
        appState,
        overlayManager: {
          getVisibleOverlayVisible: () => deps.overlayManager.getVisibleOverlayVisible(),
        },
        overlayVisibilityRuntime: {
          updateVisibleOverlayVisibility: () =>
            deps.overlay.overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
        },
        refreshCurrentSubtitle: () => {
          deps
            .getPlayback()
            .subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);
        },
        overlayShortcutsRuntime: {
          syncOverlayShortcuts: () => deps.shortcuts.overlayShortcutsRuntime.syncOverlayShortcuts(),
        },
        createMainWindow: () => {
          if (isHeadlessLaunch()) {
            return;
          }
          deps.overlay.createMainWindow();
        },
        registerGlobalShortcuts: () => {
          if (isHeadlessLaunch()) {
            return;
          }
          deps.shortcuts.registerGlobalShortcuts();
        },
        createWindowTracker: (override, targetMpvSocketPath) =>
          deps
            .getVisibleOverlayInteraction()
            .createOverlayWindowTracker(override, targetMpvSocketPath),
        updateVisibleOverlayBounds: (geometry: WindowGeometry) =>
          deps.getOverlayGeometry().updateVisibleOverlayBounds(geometry),
        bindOverlayOwner: () => deps.getVisibleOverlayInteraction().bindVisibleOverlayOwner(),
        releaseOverlayOwner: () => deps.getVisibleOverlayInteraction().releaseVisibleOverlayOwner(),
        getOverlayWindows: () => deps.overlayManager.getOverlayWindows(),
        getResolvedConfig: () => configService.getConfig(),
        showDesktopNotification,
        showOverlayNotification: (payload) =>
          deps.getOverlayNotifications().showOverlayNotification(payload),
        dismissOverlayNotification: (id) =>
          deps.getOverlayNotifications().dismissOverlayNotification(id),
        createFieldGroupingCallback: () => deps.mining.createFieldGroupingCallback(),
        getKnownWordCacheStatePath: () => knownWordCacheStatePath,
        getCachedMediaPath: (currentVideoPath, kind) =>
          deps.youtube.getCachedMediaPathForCurrentPlayback(currentVideoPath, kind),
        shouldRequireRemoteMediaCache: () =>
          deps.youtube.shouldRequireMediaCacheForCurrentPlayback(),
        getYoutubeMediaSourceUrl: () => deps.youtube.getCurrentMediaCacheSourceUrl(),
        shouldStartAnkiIntegration: () => !isHeadlessLaunch(),
      },
      initializeOverlayRuntimeBootstrapDeps: {
        isOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
        initializeOverlayRuntimeCore,
        setOverlayRuntimeInitialized: (initialized) => {
          appState.overlayRuntimeInitialized = initialized;
        },
        startBackgroundWarmups: () => {
          if (isHeadlessLaunch()) {
            return;
          }
          deps.getPlayback().startBackgroundWarmups();
        },
      },
    });

  /** Creates the overlay windows and runtime once, then wires the Anki integration it built. */
  function initializeOverlayRuntime(): void {
    initializeOverlayRuntimeHandler();
    if (!isHeadlessLaunch()) {
      deps.overlayModalRuntime.primeModalWindow();
    }
    if (appState.ankiIntegration) {
      deps.mining.attachToAnkiIntegration(appState.ankiIntegration);
    }
    deps.syncOverlayMpvSubtitleSuppression();
  }

  /** Creates whatever overlay commands need that app-ready skipped on a deferred start. */
  function ensureOverlayStartupPrereqs(): void {
    if (appState.subtitlePosition === null) {
      deps.loadSubtitlePosition();
    }
    if (appState.keybindings.length === 0) {
      resolveSessionKeybindings();
    } else if (!appState.sessionBindingsInitialized) {
      deps.shortcuts.refreshCurrentSessionBindings();
    }
    if (!appState.mpvClient) {
      appState.mpvClient = deps.getPlayback().createMpvClient();
    }
    if (!appState.runtimeOptionsManager) {
      appState.runtimeOptionsManager = createRuntimeOptionsManager();
    }
    if (!appState.subtitleTimingTracker) {
      appState.subtitleTimingTracker = new SubtitleTimingTracker();
    }
  }

  async function ensureYoutubePlaybackRuntimeReady(): Promise<void> {
    ensureOverlayStartupPrereqs();
    await deps.yomitan.ensureExtensionLoaded();
    if (!appState.overlayRuntimeInitialized) {
      initializeOverlayRuntime();
      return;
    }
    deps.overlay.ensureOverlayWindowsReadyForVisibilityActions();
  }

  /** `--refresh-known-words` without the overlay: refresh the cache, then quit. */
  async function runHeadlessInitialCommand(): Promise<void> {
    if (!appState.initialArgs?.refreshKnownWords) {
      deps.handleInitialArgs();
      return;
    }

    const resolvedConfig = configService.getConfig();
    if (resolvedConfig.ankiConnect.enabled !== true) {
      logger.error('Headless known-word refresh failed: AnkiConnect integration not enabled');
      process.exitCode = 1;
      deps.requestAppQuit();
      return;
    }

    const effectiveAnkiConfig =
      appState.runtimeOptionsManager?.getEffectiveAnkiConnectConfig(resolvedConfig.ankiConnect) ??
      resolvedConfig.ankiConnect;
    const integration = new AnkiIntegration(
      effectiveAnkiConfig,
      new SubtitleTimingTracker(),
      { send: () => undefined } as never,
      undefined,
      undefined,
      async () => ({
        keepNoteId: 0,
        deleteNoteId: 0,
        deleteDuplicate: false,
        cancelled: true,
      }),
      knownWordCacheStatePath,
      mergeAiConfig(resolvedConfig.ai, resolvedConfig.ankiConnect?.ai),
    );

    try {
      await integration.refreshKnownWordCache();
    } catch (error) {
      logger.error('Headless known-word refresh failed:', error);
      process.exitCode = 1;
    } finally {
      integration.stop();
      deps.requestAppQuit();
    }
  }

  const failHandlers = {
    logError: (details: string) => logger.error(details),
    showErrorBox: (title: string, details: string) => deps.showErrorBox(title, details),
    quit: () => deps.requestAppQuit(),
  };

  const { appReadyRuntimeRunner } = composeAppReadyRuntime({
    reloadConfigMainDeps: {
      reloadConfigStrict: () => configService.reloadConfigStrict(),
      logInfo: (message) => logger.info(message),
      logDebug: (message) => logger.debug(message),
      logWarning: (message) => logger.warn(message),
      showDesktopNotification: (title, options) => showDesktopNotification(title, options),
      startConfigHotReload: () => deps.settings.hotReload.start(),
      shouldRefreshAnilistClientSecretState: () =>
        shouldRefreshAnilistOnConfigReload(appState.initialArgs),
      refreshAnilistClientSecretState: (options) =>
        deps.getAnilist().refreshClientSecretStateIfEnabled(options),
      failHandlers,
    },
    criticalConfigErrorMainDeps: {
      getConfigPath: () => configService.getConfigPath(),
      failHandlers,
    },
    appReadyRuntimeMainDeps: {
      ensureDefaultConfigBootstrap: () => {
        ensureDefaultConfigBootstrap({
          configDir: deps.configDir,
          configFilePaths: getDefaultConfigFilePaths(deps.configDir),
          generateTemplate: () => generateConfigTemplate(DEFAULT_CONFIG),
        });
      },
      loadSubtitlePosition: () => deps.loadSubtitlePosition(),
      resolveKeybindings: () => resolveSessionKeybindings(),
      createMpvClient: () => {
        appState.mpvClient = deps.getPlayback().createMpvClient();
      },
      startPlugins: () => {
        void deps.getPluginHost().loadAll();
      },
      getResolvedConfig: () => configService.getConfig(),
      getConfigWarnings: () => configService.getWarnings(),
      logConfigWarning: (warning) => deps.logConfigWarning(warning),
      ...deps.logging,
      initRuntimeOptionsManager: () => {
        appState.runtimeOptionsManager = createRuntimeOptionsManager();
      },
      setSecondarySubMode: (mode: SecondarySubMode) => {
        deps.setSecondarySubMode(mode);
      },
      defaultSecondarySubMode: 'hover',
      defaultWebsocketPort: DEFAULT_CONFIG.websocket.port,
      defaultAnnotationWebsocketPort: DEFAULT_CONFIG.annotationWebsocket.port,
      defaultTexthookerPort: deps.defaultTexthookerPort,
      hasMpvWebsocketPlugin: () => deps.hasMpvWebsocketPlugin(),
      startSubtitleWebsocket: (port: number) => {
        deps.subtitleWsService.start(
          port,
          getCurrentSubtitleForWebsocket,
          getFrequencyOptionsForWebsocket,
        );
      },
      startAnnotationWebsocket: (port: number) => {
        deps.annotationSubtitleWsService.start(
          port,
          getCurrentSubtitleForWebsocket,
          getFrequencyOptionsForWebsocket,
        );
      },
      startTexthooker: (port: number, websocketUrl?: string) => {
        if (!deps.texthookerService.isRunning()) {
          deps.texthookerService.start(port, websocketUrl);
        }
      },
      log: (message) => logger.info(message),
      createMecabTokenizerAndCheck: async () => {
        await deps.getPlayback().createMecabTokenizerAndCheck();
      },
      createSubtitleTimingTracker: () => {
        appState.subtitleTimingTracker = new SubtitleTimingTracker();
      },
      loadYomitanExtension: async () => {
        await deps.yomitan.loadExtension();
      },
      ensureYomitanExtensionLoaded: async () => {
        await deps.yomitan.ensureExtensionLoaded();
      },
      handleFirstRunSetup: async () => {
        const snapshot = await deps.firstRunSetup.service.ensureSetupStateInitialized();
        appState.firstRunSetupCompleted = snapshot.state.status === 'completed';
        const args = appState.initialArgs;
        if (args && shouldAutoOpenFirstRunSetup(args)) {
          const force = Boolean(args.setup);
          if (force || snapshot.state.status !== 'completed') {
            deps.firstRunSetup.openWindow(force);
          }
        }
      },
      startJellyfinRemoteSession: async () => {
        await deps.getJellyfin().startJellyfinRemoteSession();
      },
      prewarmSubtitleDictionaries: async () => {
        await deps.getPlayback().prewarmSubtitleDictionaries();
      },
      startBackgroundWarmups: () => {
        deps.getPlayback().startBackgroundWarmups();
      },
      // Captured when the runner is composed, before main.ts applies the startup state.
      texthookerOnlyMode: appState.texthookerOnlyMode,
      shouldAutoInitializeOverlayRuntimeFromConfig: () =>
        appState.backgroundMode
          ? false
          : deps.configDerivedRuntime.shouldAutoInitializeOverlayRuntimeFromConfig(),
      setVisibleOverlayVisible: (visible: boolean) =>
        deps.overlay.setVisibleOverlayVisible(visible),
      initializeOverlayRuntime: () => initializeOverlayRuntime(),
      runHeadlessInitialCommand: () => runHeadlessInitialCommand(),
      handleInitialArgs: () => deps.handleInitialArgs(),
      shouldRunHeadlessInitialCommand: () => isHeadlessLaunch(),
      shouldUseMinimalStartup: () =>
        getStartupModeFlags(appState.initialArgs).shouldUseMinimalStartup,
      shouldSkipHeavyStartup: () =>
        getStartupModeFlags(appState.initialArgs).shouldSkipHeavyStartup,
      shouldHandleInitialArgsBeforeDeferredOverlayWarmup: () =>
        shouldHandleInitialArgsBeforeDeferredOverlayWarmup(appState.initialArgs),
      createImmersionTracker: () => {
        deps.stats.ensureImmersionTrackerStarted();
      },
      logDebug: (message: string) => {
        logger.debug(message);
      },
      now: () => Date.now(),
    },
    immersionTrackerStartupMainDeps: deps.stats.immersionTrackerStartupDeps,
  });

  /** The app-ready sequence; a failure is reported as fatal and quits the app. */
  async function runAppReady(): Promise<void> {
    try {
      await appReadyRuntimeRunner();
    } catch (error) {
      deps.reportFatalError(error, {
        title: 'SubMiner startup failed',
        context: 'SubMiner failed during app-ready startup.',
      });
      process.exitCode = 1;
      deps.requestAppQuit();
    }
  }

  return {
    runAppReady,
    runHeadlessInitialCommand,
    initializeOverlayRuntime,
    ensureOverlayStartupPrereqs,
    ensureYoutubePlaybackRuntimeReady,
  };
}
