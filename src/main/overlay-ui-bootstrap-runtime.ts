import type { BrowserWindow, Session } from 'electron';
import type {
  AnkiConnectConfig,
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
  RuntimeOptionState,
  WindowGeometry,
} from '../types';
import type { BaseWindowTracker } from '../window-trackers';
import {
  createOverlayGeometryRuntime,
  type OverlayGeometryRuntime,
} from './overlay-geometry-runtime';
import { createOverlayUiBootstrapRuntimeInput } from './overlay-ui-bootstrap-runtime-input';
import type { OverlayModalRuntime } from './overlay-runtime';
import { createOverlayUiRuntime, type OverlayUiRuntime } from './overlay-ui-runtime';

type WindowLike = {
  isDestroyed: () => boolean;
};

type OverlayWindowKind = 'visible' | 'modal';

type ScreenLike = {
  getCursorScreenPoint: () => { x: number; y: number };
  getDisplayNearestPoint: (point: { x: number; y: number }) => {
    workArea: { x: number; y: number; width: number; height: number };
  };
};

type OverlayWindowTrackerLike = BaseWindowTracker | null;

type OverlayRuntimeOptionsManagerLike = {
  listOptions: () => RuntimeOptionState[];
  getEffectiveAnkiConnectConfig: (config?: AnkiConnectConfig) => AnkiConnectConfig;
} | null;

type OverlayMpvClientLike = {
  connected: boolean;
  restorePreviousSecondarySubVisibility: () => void;
  send?: (payload: { command: string[] }) => void;
} | null;

type OverlayManagerLike<TWindow extends WindowLike> = {
  getMainWindow: () => TWindow | null;
  setMainWindow: (window: TWindow | null) => void;
  getModalWindow: () => TWindow | null;
  setModalWindow: (window: TWindow | null) => void;
  getVisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  setOverlayWindowBounds: (geometry: WindowGeometry) => void;
  setModalWindowBounds: (geometry: WindowGeometry) => void;
  broadcastToOverlayWindows: (channel: string, ...args: unknown[]) => void;
  getOverlayWindows: () => BrowserWindow[];
};

type OverlayModalInputStateLike = {
  getModalInputExclusive: () => boolean;
  handleModalInputStateChange: (active: boolean) => void;
};

type OverlayShortcutsRuntimeLike = {
  syncOverlayShortcuts: () => void;
  tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
};

type DictionarySupportLike = {
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
};

type FirstRunLike = {
  isSetupCompleted: () => boolean;
  openFirstRunSetupWindow: () => void;
};

type YomitanLike = {
  openYomitanSettings: () => void;
};

type JellyfinLike = {
  openJellyfinSetupWindow: () => void;
};

type AnilistLike = {
  openAnilistSetupWindow: () => void;
};

type BootstrapTrayIconLike = {
  isEmpty: () => boolean;
  resize: (options: {
    width: number;
    height: number;
    quality?: 'best' | 'better' | 'good';
  }) => BootstrapTrayIconLike;
  setTemplateImage: (enabled: boolean) => void;
};

type BootstrapTrayLike = {
  setContextMenu: (menu: any) => void;
  setToolTip: (tooltip: string) => void;
  on: (event: 'click', handler: () => void) => void;
  destroy: () => void;
};

export interface OverlayUiBootstrapAppStateInput {
  backendOverride: string | null;
  windowTracker: OverlayWindowTrackerLike;
  subtitleTimingTracker: unknown;
  mpvClient: OverlayMpvClientLike;
  mpvSocketPath: string;
  runtimeOptionsManager: OverlayRuntimeOptionsManagerLike;
  ankiIntegration: unknown;
  overlayRuntimeInitialized: boolean;
  overlayDebugVisualizationEnabled: boolean;
  statsOverlayVisible: boolean;
  trackerNotReadyWarningShown: boolean;
  yomitanSession: Session | null;
}

export interface OverlayUiBootstrapElectronInput<
  TWindow extends WindowLike,
  TMenuItem = unknown,
  TMenu = unknown,
> {
  screen: ScreenLike;
  appPath: string;
  resourcesPath: string;
  dirname: string;
  platform: NodeJS.Platform;
  joinPath: (...parts: string[]) => string;
  fileExists: (candidate: string) => boolean;
  createImageFromPath: (iconPath: string) => BootstrapTrayIconLike;
  createEmptyImage: () => BootstrapTrayIconLike;
  createTray: (icon: BootstrapTrayIconLike) => BootstrapTrayLike;
  buildMenuFromTemplate: (template: TMenuItem[]) => TMenu;
}

export interface OverlayUiBootstrapInput<TWindow extends WindowLike> {
  appState: OverlayUiBootstrapAppStateInput;
  overlayManager: OverlayManagerLike<TWindow>;
  overlayModalInputState: OverlayModalInputStateLike;
  overlayModalRuntime: OverlayModalRuntime;
  overlayShortcutsRuntime: OverlayShortcutsRuntimeLike;
  dictionarySupport: DictionarySupportLike;
  firstRun: FirstRunLike;
  yomitan: YomitanLike;
  jellyfin: JellyfinLike;
  anilist: AnilistLike;
  electron: OverlayUiBootstrapElectronInput<TWindow>;
  windowing: {
    isDev: boolean;
    createOverlayWindowCore: (
      kind: OverlayWindowKind,
      options: {
        isDev: boolean;
        ensureOverlayWindowLevel: (window: TWindow) => void;
        onRuntimeOptionsChanged: () => void;
        setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
        isOverlayVisible: (windowKind: OverlayWindowKind) => boolean;
        tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
        forwardTabToMpv: () => void;
        onWindowClosed: (windowKind: OverlayWindowKind) => void;
        yomitanSession?: Electron.Session | null;
      },
    ) => TWindow;
    ensureOverlayWindowLevelCore: (window: TWindow) => void;
    syncOverlayWindowLayer: (window: TWindow, layer: 'visible') => void;
    enforceOverlayLayerOrderCore: (params: {
      visibleOverlayVisible: boolean;
      mainWindow: TWindow | null;
      ensureOverlayWindowLevel: (window: TWindow) => void;
    }) => void;
    createWindowTrackerCore: (
      override?: string | null,
      targetMpvSocketPath?: string | null,
    ) => BaseWindowTracker | null;
  };
  actions: {
    showMpvOsd: (message: string) => void;
    showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
    sendMpvCommand: (command: (string | number)[]) => void;
    broadcastRuntimeOptionsChangedRuntime: (
      getRuntimeOptionsState: () => RuntimeOptionState[],
      broadcastToOverlayWindows: (channel: string, ...args: unknown[]) => void,
    ) => void;
    setOverlayDebugVisualizationEnabledRuntime: (
      currentEnabled: boolean,
      nextEnabled: boolean,
      setCurrentEnabled: (enabled: boolean) => void,
    ) => void;
    resolveTrayIconPathRuntime: (options: {
      platform: string;
      resourcesPath: string;
      appPath: string;
      dirname: string;
      joinPath: (...parts: string[]) => string;
      fileExists: (path: string) => boolean;
    }) => string | null;
    buildTrayMenuTemplateRuntime: (handlers: {
      openOverlay: () => void;
      openFirstRunSetup: () => void;
      showFirstRunSetup: boolean;
      openWindowsMpvLauncherSetup: () => void;
      showWindowsMpvLauncherSetup: boolean;
      openYomitanSettings: () => void;
      openRuntimeOptions: () => void;
      openJellyfinSetup: () => void;
      openAnilistSetup: () => void;
      quitApp: () => void;
    }) => unknown[];
    initializeOverlayRuntimeCore: (options: unknown) => void;
    ensureOverlayMpvSubtitlesHidden: () => Promise<void> | void;
    syncOverlayMpvSubtitleSuppression: () => void;
    registerGlobalShortcuts: () => void;
    startBackgroundWarmups: () => void;
    getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig };
    requestAppQuit: () => void;
  };
  trayState: {
    getTray: () => BootstrapTrayLike | null;
    setTray: (tray: BootstrapTrayLike | null) => void;
    trayTooltip: string;
    logWarn: (message: string) => void;
  };
  startup: {
    shouldSkipHeadlessOverlayBootstrap: () => boolean;
    getKnownWordCacheStatePath: () => string;
    onInitialized?: () => void;
  };
}

export interface OverlayUiBootstrapRuntime<TWindow extends WindowLike> {
  overlayGeometry: OverlayGeometryRuntime<TWindow>;
  overlayUi: OverlayUiRuntime<TWindow>;
  syncOverlayVisibilityForModal: () => void;
}

export function createOverlayUiBootstrapRuntime<TWindow extends WindowLike>(
  input: OverlayUiBootstrapInput<TWindow>,
): OverlayUiBootstrapRuntime<TWindow> {
  const overlayGeometry = createOverlayGeometryRuntime<TWindow>({
    screen: input.electron.screen,
    windowState: {
      getMainWindow: () => input.overlayManager.getMainWindow(),
      setOverlayWindowBounds: (geometry) => input.overlayManager.setOverlayWindowBounds(geometry),
      setModalWindowBounds: (geometry) => input.overlayManager.setModalWindowBounds(geometry),
      getVisibleOverlayVisible: () => input.overlayManager.getVisibleOverlayVisible(),
    },
    getWindowTracker: () => input.appState.windowTracker,
    ensureOverlayWindowLevelCore: (window) => input.windowing.ensureOverlayWindowLevelCore(window),
    syncOverlayWindowLayer: (window, layer) =>
      input.windowing.syncOverlayWindowLayer(window, layer),
    enforceOverlayLayerOrderCore: (params) => input.windowing.enforceOverlayLayerOrderCore(params),
  });

  let overlayUi: OverlayUiRuntime<TWindow> | undefined;

  overlayUi = createOverlayUiRuntime(
    createOverlayUiBootstrapRuntimeInput<TWindow>({
      windows: {
        state: {
          getMainWindow: () => input.overlayManager.getMainWindow(),
          setMainWindow: (window) => input.overlayManager.setMainWindow(window),
          getModalWindow: () => input.overlayManager.getModalWindow(),
          setModalWindow: (window) => input.overlayManager.setModalWindow(window),
          getVisibleOverlayVisible: () => input.overlayManager.getVisibleOverlayVisible(),
          setVisibleOverlayVisible: (visible) =>
            input.overlayManager.setVisibleOverlayVisible(visible),
          getOverlayDebugVisualizationEnabled: () =>
            input.appState.overlayDebugVisualizationEnabled,
          setOverlayDebugVisualizationEnabled: (enabled) => {
            input.appState.overlayDebugVisualizationEnabled = enabled;
          },
        },
        geometry: {
          getCurrentOverlayGeometry: () => overlayGeometry.getCurrentOverlayGeometry(),
        },
        modal: {
          setModalWindowBounds: (geometry) => input.overlayManager.setModalWindowBounds(geometry),
          onModalStateChange: (active) => {
            input.overlayModalInputState.handleModalInputStateChange(active);
          },
        },
        modalRuntime: input.overlayModalRuntime as never,
        visibility: {
          service: {
            getModalActive: () => input.overlayModalInputState.getModalInputExclusive(),
            getForceMousePassthrough: () => input.appState.statsOverlayVisible,
            getWindowTracker: () => input.appState.windowTracker,
            getTrackerNotReadyWarningShown: () => input.appState.trackerNotReadyWarningShown,
            setTrackerNotReadyWarningShown: (shown) => {
              input.appState.trackerNotReadyWarningShown = shown;
            },
            updateVisibleOverlayBounds: (geometry) =>
              overlayGeometry.updateVisibleOverlayBounds(geometry),
            ensureOverlayWindowLevel: (window) => overlayGeometry.ensureOverlayWindowLevel(window),
            syncPrimaryOverlayWindowLayer: (layer) =>
              overlayGeometry.syncPrimaryOverlayWindowLayer(layer),
            enforceOverlayLayerOrder: () => overlayGeometry.enforceOverlayLayerOrder(),
            syncOverlayShortcuts: () => input.overlayShortcutsRuntime.syncOverlayShortcuts(),
            isMacOSPlatform: () => input.electron.platform === 'darwin',
            isWindowsPlatform: () => input.electron.platform === 'win32',
            showOverlayLoadingOsd: (message) => input.actions.showMpvOsd(message),
            resolveFallbackBounds: () => overlayGeometry.getOverlayGeometryFallback(),
          },
          overlayWindows: {
            createOverlayWindowCore: (kind, options) =>
              input.windowing.createOverlayWindowCore(kind, options),
            isDev: input.windowing.isDev,
            ensureOverlayWindowLevel: (window) => overlayGeometry.ensureOverlayWindowLevel(window),
            onRuntimeOptionsChanged: () => {
              overlayUi?.broadcastRuntimeOptionsChanged();
            },
            setOverlayDebugVisualizationEnabled: (enabled) => {
              overlayUi?.setOverlayDebugVisualizationEnabled(enabled);
            },
            isOverlayVisible: (windowKind) =>
              windowKind === 'visible' ? input.overlayManager.getVisibleOverlayVisible() : false,
            getYomitanSession: () => input.appState.yomitanSession,
            tryHandleOverlayShortcutLocalFallback: (overlayInput) =>
              input.overlayShortcutsRuntime.tryHandleOverlayShortcutLocalFallback(overlayInput),
            forwardTabToMpv: () => input.actions.sendMpvCommand(['keypress', 'TAB']),
            onWindowClosed: (windowKind) => {
              if (windowKind === 'visible') {
                input.overlayManager.setMainWindow(null);
                return;
              }
              input.overlayManager.setModalWindow(null);
            },
          },
          actions: {
            setVisibleOverlayVisibleCore: ({
              visible,
              setVisibleOverlayVisibleState,
              updateVisibleOverlayVisibility,
            }) => {
              setVisibleOverlayVisibleState(visible);
              updateVisibleOverlayVisibility();
            },
          },
        },
      },
      overlayActions: {
        getRuntimeOptionsManager: () => input.appState.runtimeOptionsManager,
        getMpvClient: () => input.appState.mpvClient,
        broadcastRuntimeOptionsChangedRuntime: (
          getRuntimeOptionsState,
          broadcastToOverlayWindows,
        ) =>
          input.actions.broadcastRuntimeOptionsChangedRuntime(
            getRuntimeOptionsState,
            broadcastToOverlayWindows,
          ),
        broadcastToOverlayWindows: (channel, ...args) =>
          input.overlayManager.broadcastToOverlayWindows(channel, ...args),
        setOverlayDebugVisualizationEnabledRuntime: (
          currentEnabled,
          nextEnabled,
          setCurrentEnabled,
        ) =>
          input.actions.setOverlayDebugVisualizationEnabledRuntime(
            currentEnabled,
            nextEnabled,
            setCurrentEnabled,
          ),
      },
      tray: {
        resolveTrayIconPathDeps: {
          resolveTrayIconPathRuntime: input.actions.resolveTrayIconPathRuntime,
          platform: input.electron.platform,
          resourcesPath: input.electron.resourcesPath,
          appPath: input.electron.appPath,
          dirname: input.electron.dirname,
          joinPath: (...parts) => input.electron.joinPath(...parts),
          fileExists: (candidate) => input.electron.fileExists(candidate),
        },
        buildTrayMenuTemplateDeps: {
          buildTrayMenuTemplateRuntime: input.actions.buildTrayMenuTemplateRuntime,
          initializeOverlayRuntime: () => {
            overlayUi?.initializeOverlayRuntime();
          },
          isOverlayRuntimeInitialized: () => input.appState.overlayRuntimeInitialized,
          setVisibleOverlayVisible: (visible) => {
            overlayUi?.setVisibleOverlayVisible(visible);
          },
          showFirstRunSetup: () => !input.firstRun.isSetupCompleted(),
          openFirstRunSetupWindow: () => input.firstRun.openFirstRunSetupWindow(),
          showWindowsMpvLauncherSetup: () => input.electron.platform === 'win32',
          openYomitanSettings: () => input.yomitan.openYomitanSettings(),
          openRuntimeOptionsPalette: () => {
            overlayUi?.openRuntimeOptionsPalette();
          },
          openJellyfinSetupWindow: () => input.jellyfin.openJellyfinSetupWindow(),
          openAnilistSetupWindow: () => input.anilist.openAnilistSetupWindow(),
          quitApp: () => input.actions.requestAppQuit(),
        },
        ensureTrayDeps: {
          getTray: () => input.trayState.getTray(),
          setTray: (tray) => input.trayState.setTray(tray),
          createImageFromPath: (iconPath) => input.electron.createImageFromPath(iconPath),
          createEmptyImage: () => input.electron.createEmptyImage(),
          createTray: (icon) => input.electron.createTray(icon),
          trayTooltip: input.trayState.trayTooltip,
          platform: input.electron.platform,
          logWarn: (message) => input.trayState.logWarn(message),
          initializeOverlayRuntime: () => {
            overlayUi?.initializeOverlayRuntime();
          },
          isOverlayRuntimeInitialized: () => input.appState.overlayRuntimeInitialized,
          setVisibleOverlayVisible: (visible) => {
            overlayUi?.setVisibleOverlayVisible(visible);
          },
        },
        destroyTrayDeps: {
          getTray: () => input.trayState.getTray(),
          setTray: (tray) => input.trayState.setTray(tray),
        },
        buildMenuFromTemplate: (template) => input.electron.buildMenuFromTemplate(template),
      },
      bootstrap: {
        initializeOverlayRuntimeMainDeps: {
          appState: input.appState,
          overlayManager: {
            getVisibleOverlayVisible: () => input.overlayManager.getVisibleOverlayVisible(),
          },
          overlayVisibilityRuntime: {
            updateVisibleOverlayVisibility: () => {
              overlayUi?.updateVisibleOverlayVisibility();
            },
          },
          overlayShortcutsRuntime: {
            syncOverlayShortcuts: () => input.overlayShortcutsRuntime.syncOverlayShortcuts(),
          },
          createMainWindow: () => {
            if (input.startup.shouldSkipHeadlessOverlayBootstrap()) {
              return;
            }
            overlayUi?.createMainWindow();
          },
          registerGlobalShortcuts: () => {
            if (input.startup.shouldSkipHeadlessOverlayBootstrap()) {
              return;
            }
            input.actions.registerGlobalShortcuts();
          },
          createWindowTracker: (override, targetMpvSocketPath) => {
            if (input.startup.shouldSkipHeadlessOverlayBootstrap()) {
              return null;
            }
            return input.windowing.createWindowTrackerCore(
              override as string | null | undefined,
              targetMpvSocketPath as string | null | undefined,
            );
          },
          updateVisibleOverlayBounds: (geometry) =>
            overlayGeometry.updateVisibleOverlayBounds(geometry),
          getOverlayWindows: () => input.overlayManager.getOverlayWindows(),
          getResolvedConfig: () => input.actions.getResolvedConfig(),
          showDesktopNotification: (title, options) =>
            input.actions.showDesktopNotification(title, options),
          createFieldGroupingCallback: () => input.dictionarySupport.createFieldGroupingCallback(),
          getKnownWordCacheStatePath: () => input.startup.getKnownWordCacheStatePath(),
          shouldStartAnkiIntegration: () => !input.startup.shouldSkipHeadlessOverlayBootstrap(),
        },
        initializeOverlayRuntimeBootstrapDeps: {
          isOverlayRuntimeInitialized: () => input.appState.overlayRuntimeInitialized,
          initializeOverlayRuntimeCore: (options) =>
            input.actions.initializeOverlayRuntimeCore(options),
          setOverlayRuntimeInitialized: (initialized) => {
            input.appState.overlayRuntimeInitialized = initialized;
          },
          startBackgroundWarmups: () => {
            if (input.startup.shouldSkipHeadlessOverlayBootstrap()) {
              return;
            }
            input.actions.startBackgroundWarmups();
          },
        },
        onInitialized: input.startup.onInitialized,
      },
      runtimeState: {
        isOverlayRuntimeInitialized: () => input.appState.overlayRuntimeInitialized,
        setOverlayRuntimeInitialized: (initialized) => {
          input.appState.overlayRuntimeInitialized = initialized;
        },
      },
      mpvSubtitle: {
        ensureOverlayMpvSubtitlesHidden: () => input.actions.ensureOverlayMpvSubtitlesHidden(),
        syncOverlayMpvSubtitleSuppression: () => input.actions.syncOverlayMpvSubtitleSuppression(),
      },
    }),
  );

  return {
    overlayGeometry,
    overlayUi,
    syncOverlayVisibilityForModal: () => {
      overlayUi.updateVisibleOverlayVisibility();
    },
  };
}
