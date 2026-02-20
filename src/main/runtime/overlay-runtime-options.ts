import type {
  AnkiConnectConfig,
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
  WindowGeometry,
} from '../../types';
import type { BrowserWindow } from 'electron';

type OverlayRuntimeOptions = {
  backendOverride: string | null;
  getInitialInvisibleOverlayVisibility: () => boolean;
  createMainWindow: () => void;
  createInvisibleWindow: () => void;
  registerGlobalShortcuts: () => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  updateInvisibleOverlayBounds: (geometry: WindowGeometry) => void;
  isVisibleOverlayVisible: () => boolean;
  isInvisibleOverlayVisible: () => boolean;
  updateVisibleOverlayVisibility: () => void;
  updateInvisibleOverlayVisibility: () => void;
  getOverlayWindows: () => BrowserWindow[];
  syncOverlayShortcuts: () => void;
  setWindowTracker: (tracker: unknown | null) => void;
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
  getInitialInvisibleOverlayVisibility: () => boolean;
  createMainWindow: () => void;
  createInvisibleWindow: () => void;
  registerGlobalShortcuts: () => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  updateInvisibleOverlayBounds: (geometry: WindowGeometry) => void;
  isVisibleOverlayVisible: () => boolean;
  isInvisibleOverlayVisible: () => boolean;
  updateVisibleOverlayVisibility: () => void;
  updateInvisibleOverlayVisibility: () => void;
  getOverlayWindows: () => BrowserWindow[];
  syncOverlayShortcuts: () => void;
  setWindowTracker: (tracker: unknown | null) => void;
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
    getInitialInvisibleOverlayVisibility: deps.getInitialInvisibleOverlayVisibility,
    createMainWindow: deps.createMainWindow,
    createInvisibleWindow: deps.createInvisibleWindow,
    registerGlobalShortcuts: deps.registerGlobalShortcuts,
    updateVisibleOverlayBounds: deps.updateVisibleOverlayBounds,
    updateInvisibleOverlayBounds: deps.updateInvisibleOverlayBounds,
    isVisibleOverlayVisible: deps.isVisibleOverlayVisible,
    isInvisibleOverlayVisible: deps.isInvisibleOverlayVisible,
    updateVisibleOverlayVisibility: deps.updateVisibleOverlayVisibility,
    updateInvisibleOverlayVisibility: deps.updateInvisibleOverlayVisibility,
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
