import type { AnkiConnectConfig } from '../../types';
import type { createBuildInitializeOverlayRuntimeOptionsHandler } from './overlay-runtime-options';

type OverlayRuntimeOptionsMainDeps = Parameters<
  typeof createBuildInitializeOverlayRuntimeOptionsHandler
>[0];

export function createBuildInitializeOverlayRuntimeMainDepsHandler(deps: {
  appState: {
    backendOverride: string | null;
    windowTracker: Parameters<OverlayRuntimeOptionsMainDeps['setWindowTracker']>[0];
    subtitleTimingTracker: ReturnType<OverlayRuntimeOptionsMainDeps['getSubtitleTimingTracker']>;
    mpvClient: ReturnType<OverlayRuntimeOptionsMainDeps['getMpvClient']>;
    mpvSocketPath: string;
    runtimeOptionsManager: ReturnType<OverlayRuntimeOptionsMainDeps['getRuntimeOptionsManager']>;
    ankiIntegration: Parameters<OverlayRuntimeOptionsMainDeps['setAnkiIntegration']>[0];
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
  updateVisibleOverlayBounds: (geometry: {
    x: number;
    y: number;
    width: number;
    height: number;
  }) => void;
  updateInvisibleOverlayBounds: (geometry: {
    x: number;
    y: number;
    width: number;
    height: number;
  }) => void;
  getOverlayWindows: OverlayRuntimeOptionsMainDeps['getOverlayWindows'];
  getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig };
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
  createFieldGroupingCallback: OverlayRuntimeOptionsMainDeps['createFieldGroupingCallback'];
  getKnownWordCacheStatePath: () => string;
}) {
  return (): OverlayRuntimeOptionsMainDeps => ({
    getBackendOverride: () => deps.appState.backendOverride,
    getInitialInvisibleOverlayVisibility: () => deps.getInitialInvisibleOverlayVisibility(),
    createMainWindow: () => deps.createMainWindow(),
    createInvisibleWindow: () => deps.createInvisibleWindow(),
    registerGlobalShortcuts: () => deps.registerGlobalShortcuts(),
    updateVisibleOverlayBounds: (geometry: {
      x: number;
      y: number;
      width: number;
      height: number;
    }) => deps.updateVisibleOverlayBounds(geometry),
    updateInvisibleOverlayBounds: (geometry: {
      x: number;
      y: number;
      width: number;
      height: number;
    }) => deps.updateInvisibleOverlayBounds(geometry),
    isVisibleOverlayVisible: () => deps.overlayManager.getVisibleOverlayVisible(),
    isInvisibleOverlayVisible: () => deps.overlayManager.getInvisibleOverlayVisible(),
    updateVisibleOverlayVisibility: () =>
      deps.overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    updateInvisibleOverlayVisibility: () =>
      deps.overlayVisibilityRuntime.updateInvisibleOverlayVisibility(),
    getOverlayWindows: () => deps.getOverlayWindows(),
    syncOverlayShortcuts: () => deps.overlayShortcutsRuntime.syncOverlayShortcuts(),
    setWindowTracker: (tracker) => {
      deps.appState.windowTracker = tracker;
    },
    getResolvedConfig: () => deps.getResolvedConfig(),
    getSubtitleTimingTracker: () => deps.appState.subtitleTimingTracker,
    getMpvClient: () => deps.appState.mpvClient,
    getMpvSocketPath: () => deps.appState.mpvSocketPath,
    getRuntimeOptionsManager: () => deps.appState.runtimeOptionsManager,
    setAnkiIntegration: (integration) => {
      deps.appState.ankiIntegration = integration;
    },
    showDesktopNotification: deps.showDesktopNotification,
    createFieldGroupingCallback: () => deps.createFieldGroupingCallback(),
    getKnownWordCacheStatePath: () => deps.getKnownWordCacheStatePath(),
  });
}
