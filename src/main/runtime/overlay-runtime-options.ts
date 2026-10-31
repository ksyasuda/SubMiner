import type {
  AnkiConnectConfig,
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
} from '../../types/anki';
import type { BrowserWindow } from 'electron';
import type { WindowGeometry } from '../../types/runtime';
import type { OverlayNotificationPayload } from '../../types/notification';
import type { BaseWindowTracker } from '../../window-trackers';

type OverlayRuntimeOptions = {
  backendOverride: string | null;
  createMainWindow: () => void;
  registerGlobalShortcuts: () => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  isVisibleOverlayVisible: () => boolean;
  updateVisibleOverlayVisibility: () => void;
  refreshCurrentSubtitle?: () => void;
  getOverlayWindows: () => BrowserWindow[];
  syncOverlayShortcuts: () => void;
  setWindowTracker: (tracker: BaseWindowTracker | null) => void;
  createWindowTracker?: (
    override?: string | null,
    targetMpvSocketPath?: string | null,
  ) => BaseWindowTracker | null;
  getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig };
  getSubtitleTimingTracker: () => unknown | null;
  getMpvClient: () => { send?: (payload: { command: string[] }) => void } | null;
  getMpvSocketPath: () => string;
  getRuntimeOptionsManager: () => {
    getEffectiveAnkiConnectConfig: (config?: AnkiConnectConfig) => AnkiConnectConfig;
  } | null;
  setAnkiIntegration: (integration: unknown | null) => void;
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
  showOverlayNotification?: (payload: OverlayNotificationPayload) => void;
  dismissOverlayNotification?: (id: string) => void;
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
  getKnownWordCacheStatePath: () => string;
  getCachedMediaPath?: (
    currentVideoPath: string,
    kind: 'audio' | 'video',
  ) => Promise<string | null>;
  shouldRequireRemoteMediaCache?: () => boolean;
  getYoutubeMediaSourceUrl?: () => Promise<string | null | undefined> | string | null | undefined;
  shouldStartAnkiIntegration: () => boolean;
  bindOverlayOwner?: () => void;
  releaseOverlayOwner?: () => void;
};

// Builds the options for core initializeOverlayRuntime from main-process state. appState
// fields are read when the options are built (or lazily via getters), and the window tracker
// and Anki integration produced by the core are written back to appState.
export function createBuildInitializeOverlayRuntimeOptionsHandler(deps: {
  appState: {
    backendOverride: string | null;
    windowTracker: Parameters<OverlayRuntimeOptions['setWindowTracker']>[0];
    subtitleTimingTracker: ReturnType<OverlayRuntimeOptions['getSubtitleTimingTracker']>;
    mpvClient: ReturnType<OverlayRuntimeOptions['getMpvClient']>;
    mpvSocketPath: string;
    runtimeOptionsManager: ReturnType<OverlayRuntimeOptions['getRuntimeOptionsManager']>;
    ankiIntegration: Parameters<OverlayRuntimeOptions['setAnkiIntegration']>[0];
  };
  overlayManager: {
    getVisibleOverlayVisible: () => boolean;
  };
  overlayVisibilityRuntime: {
    updateVisibleOverlayVisibility: () => void;
  };
  refreshCurrentSubtitle?: () => void;
  overlayShortcutsRuntime: {
    syncOverlayShortcuts: () => void;
  };
  createMainWindow: () => void;
  registerGlobalShortcuts: () => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  getOverlayWindows: () => BrowserWindow[];
  createWindowTracker?: OverlayRuntimeOptions['createWindowTracker'];
  getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig };
  showDesktopNotification: OverlayRuntimeOptions['showDesktopNotification'];
  showOverlayNotification?: (payload: OverlayNotificationPayload) => void;
  dismissOverlayNotification?: (id: string) => void;
  createFieldGroupingCallback: OverlayRuntimeOptions['createFieldGroupingCallback'];
  getKnownWordCacheStatePath: () => string;
  getCachedMediaPath?: OverlayRuntimeOptions['getCachedMediaPath'];
  shouldRequireRemoteMediaCache?: () => boolean;
  getYoutubeMediaSourceUrl?: OverlayRuntimeOptions['getYoutubeMediaSourceUrl'];
  shouldStartAnkiIntegration: () => boolean;
  bindOverlayOwner?: () => void;
  releaseOverlayOwner?: () => void;
}) {
  return (): OverlayRuntimeOptions => ({
    backendOverride: deps.appState.backendOverride,
    createMainWindow: () => deps.createMainWindow(),
    registerGlobalShortcuts: () => deps.registerGlobalShortcuts(),
    updateVisibleOverlayBounds: (geometry) => deps.updateVisibleOverlayBounds(geometry),
    isVisibleOverlayVisible: () => deps.overlayManager.getVisibleOverlayVisible(),
    updateVisibleOverlayVisibility: () =>
      deps.overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    refreshCurrentSubtitle: () => deps.refreshCurrentSubtitle?.(),
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
    showOverlayNotification: deps.showOverlayNotification,
    dismissOverlayNotification: deps.dismissOverlayNotification,
    createFieldGroupingCallback: () => deps.createFieldGroupingCallback(),
    getKnownWordCacheStatePath: () => deps.getKnownWordCacheStatePath(),
    ...(deps.getCachedMediaPath ? { getCachedMediaPath: deps.getCachedMediaPath } : {}),
    ...(deps.shouldRequireRemoteMediaCache
      ? { shouldRequireRemoteMediaCache: deps.shouldRequireRemoteMediaCache }
      : {}),
    ...(deps.getYoutubeMediaSourceUrl
      ? { getYoutubeMediaSourceUrl: deps.getYoutubeMediaSourceUrl }
      : {}),
    shouldStartAnkiIntegration: () => deps.shouldStartAnkiIntegration(),
    bindOverlayOwner: deps.bindOverlayOwner,
    releaseOverlayOwner: deps.releaseOverlayOwner,
  });
}
