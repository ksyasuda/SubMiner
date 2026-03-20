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
  };
  overlayVisibilityRuntime: {
    updateVisibleOverlayVisibility: () => void;
  };
  overlayShortcutsRuntime: {
    syncOverlayShortcuts: () => void;
  };
  createMainWindow: () => void;
  registerGlobalShortcuts: () => void;
  updateVisibleOverlayBounds: (geometry: {
    x: number;
    y: number;
    width: number;
    height: number;
  }) => void;
  getOverlayWindows: OverlayRuntimeOptionsMainDeps['getOverlayWindows'];
  createWindowTracker?: OverlayRuntimeOptionsMainDeps['createWindowTracker'];
  getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig };
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
  createFieldGroupingCallback: OverlayRuntimeOptionsMainDeps['createFieldGroupingCallback'];
  getKnownWordCacheStatePath: () => string;
  shouldStartAnkiIntegration: () => boolean;
}) {
  return (): OverlayRuntimeOptionsMainDeps => ({
    getBackendOverride: () => deps.appState.backendOverride,
    createMainWindow: () => deps.createMainWindow(),
    registerGlobalShortcuts: () => deps.registerGlobalShortcuts(),
    updateVisibleOverlayBounds: (geometry: {
      x: number;
      y: number;
      width: number;
      height: number;
    }) => deps.updateVisibleOverlayBounds(geometry),
    isVisibleOverlayVisible: () => deps.overlayManager.getVisibleOverlayVisible(),
    updateVisibleOverlayVisibility: () =>
      deps.overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    getOverlayWindows: () => deps.getOverlayWindows(),
    syncOverlayShortcuts: () => deps.overlayShortcutsRuntime.syncOverlayShortcuts(),
    setWindowTracker: (tracker) => {
      deps.appState.windowTracker = tracker;
    },
    createWindowTracker: deps.createWindowTracker,
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
    shouldStartAnkiIntegration: () => deps.shouldStartAnkiIntegration(),
  });
}
