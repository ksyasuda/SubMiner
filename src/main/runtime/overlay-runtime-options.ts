import type {
  AnkiConnectConfig,
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
  WindowGeometry,
} from '../../types';
import type { BrowserWindow } from 'electron';
import type { BaseWindowTracker } from '../../window-trackers';

type OverlayRuntimeOptions = {
  backendOverride: string | null;
  createMainWindow: () => void;
  registerGlobalShortcuts: () => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  isVisibleOverlayVisible: () => boolean;
  updateVisibleOverlayVisibility: () => void;
  getOverlayWindows: () => BrowserWindow[];
  syncOverlayShortcuts: () => void;
  setWindowTracker: (tracker: BaseWindowTracker | null) => void;
  getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig };
  getSubtitleTimingTracker: () => unknown | null;
  getMpvClient: () => { send?: (payload: { command: string[] }) => void } | null;
  getMpvSocketPath: () => string;
  getRuntimeOptionsManager: () => {
    getEffectiveAnkiConnectConfig: (config?: AnkiConnectConfig) => AnkiConnectConfig;
  } | null;
  setAnkiIntegration: (integration: unknown | null) => void;
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
  getKnownWordCacheStatePath: () => string;
};

export function createBuildInitializeOverlayRuntimeOptionsHandler(deps: {
  getBackendOverride: () => string | null;
  createMainWindow: () => void;
  registerGlobalShortcuts: () => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  isVisibleOverlayVisible: () => boolean;
  updateVisibleOverlayVisibility: () => void;
  getOverlayWindows: () => BrowserWindow[];
  syncOverlayShortcuts: () => void;
  setWindowTracker: (tracker: BaseWindowTracker | null) => void;
  getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig };
  getSubtitleTimingTracker: () => unknown | null;
  getMpvClient: () => { send?: (payload: { command: string[] }) => void } | null;
  getMpvSocketPath: () => string;
  getRuntimeOptionsManager: () => {
    getEffectiveAnkiConnectConfig: (config?: AnkiConnectConfig) => AnkiConnectConfig;
  } | null;
  setAnkiIntegration: (integration: unknown | null) => void;
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
  getKnownWordCacheStatePath: () => string;
}) {
  return (): OverlayRuntimeOptions => ({
    backendOverride: deps.getBackendOverride(),
    createMainWindow: deps.createMainWindow,
    registerGlobalShortcuts: deps.registerGlobalShortcuts,
    updateVisibleOverlayBounds: deps.updateVisibleOverlayBounds,
    isVisibleOverlayVisible: deps.isVisibleOverlayVisible,
    updateVisibleOverlayVisibility: deps.updateVisibleOverlayVisibility,
    getOverlayWindows: deps.getOverlayWindows,
    syncOverlayShortcuts: deps.syncOverlayShortcuts,
    setWindowTracker: deps.setWindowTracker,
    getResolvedConfig: deps.getResolvedConfig,
    getSubtitleTimingTracker: deps.getSubtitleTimingTracker,
    getMpvClient: deps.getMpvClient,
    getMpvSocketPath: deps.getMpvSocketPath,
    getRuntimeOptionsManager: deps.getRuntimeOptionsManager,
    setAnkiIntegration: deps.setAnkiIntegration,
    showDesktopNotification: deps.showDesktopNotification,
    createFieldGroupingCallback: deps.createFieldGroupingCallback,
    getKnownWordCacheStatePath: deps.getKnownWordCacheStatePath,
  });
}
