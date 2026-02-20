import type { AnkiConnectConfig } from '../../types';

export function createBuildInitializeOverlayRuntimeMainDepsHandler(deps: {
  appState: {
    backendOverride: string | null;
    windowTracker: unknown | null;
    subtitleTimingTracker: unknown | null;
    mpvClient: unknown | null;
    mpvSocketPath: string;
    runtimeOptionsManager: unknown | null;
    ankiIntegration: unknown | null;
  };
  overlayManager: {
    getVisibleOverlayVisible: () => boolean;
    getInvisibleOverlayVisible: () => boolean;
  };
  overlayVisibilityRuntime: {
    updateVisibleOverlayVisibility: () => void;
    updateInvisibleOverlayVisibility: () => void;
  };
  overlayShortcutsRuntime: {
    syncOverlayShortcuts: () => void;
  };
  getInitialInvisibleOverlayVisibility: () => boolean;
  createMainWindow: () => void;
  createInvisibleWindow: () => void;
  registerGlobalShortcuts: () => void;
  updateVisibleOverlayBounds: (geometry: { x: number; y: number; width: number; height: number }) => void;
  updateInvisibleOverlayBounds: (geometry: {
    x: number;
    y: number;
    width: number;
    height: number;
  }) => void;
  getOverlayWindows: () => unknown[];
  getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig };
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
  createFieldGroupingCallback: () => unknown;
  getKnownWordCacheStatePath: () => string;
}) {
  return () => ({
    getBackendOverride: () => deps.appState.backendOverride,
    getInitialInvisibleOverlayVisibility: () => deps.getInitialInvisibleOverlayVisibility(),
    createMainWindow: () => deps.createMainWindow(),
    createInvisibleWindow: () => deps.createInvisibleWindow(),
    registerGlobalShortcuts: () => deps.registerGlobalShortcuts(),
    updateVisibleOverlayBounds: (geometry: { x: number; y: number; width: number; height: number }) =>
      deps.updateVisibleOverlayBounds(geometry),
    updateInvisibleOverlayBounds: (geometry: {
      x: number;
      y: number;
      width: number;
      height: number;
    }) => deps.updateInvisibleOverlayBounds(geometry),
    isVisibleOverlayVisible: () => deps.overlayManager.getVisibleOverlayVisible(),
    isInvisibleOverlayVisible: () => deps.overlayManager.getInvisibleOverlayVisible(),
    updateVisibleOverlayVisibility: () => deps.overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    updateInvisibleOverlayVisibility: () =>
      deps.overlayVisibilityRuntime.updateInvisibleOverlayVisibility(),
    getOverlayWindows: () => deps.getOverlayWindows() as never,
    syncOverlayShortcuts: () => deps.overlayShortcutsRuntime.syncOverlayShortcuts(),
    setWindowTracker: (tracker: unknown | null) => {
      deps.appState.windowTracker = tracker;
    },
    getResolvedConfig: () => deps.getResolvedConfig(),
    getSubtitleTimingTracker: () => deps.appState.subtitleTimingTracker,
    getMpvClient: () =>
      (deps.appState.mpvClient as { send?: (payload: { command: string[] }) => void } | null),
    getMpvSocketPath: () => deps.appState.mpvSocketPath,
    getRuntimeOptionsManager: () =>
      deps.appState.runtimeOptionsManager as
        | { getEffectiveAnkiConnectConfig: (config?: AnkiConnectConfig) => AnkiConnectConfig }
        | null,
    setAnkiIntegration: (integration: unknown | null) => {
      deps.appState.ankiIntegration = integration;
    },
    showDesktopNotification: deps.showDesktopNotification,
    createFieldGroupingCallback: () => deps.createFieldGroupingCallback() as never,
    getKnownWordCacheStatePath: () => deps.getKnownWordCacheStatePath(),
  });
}
