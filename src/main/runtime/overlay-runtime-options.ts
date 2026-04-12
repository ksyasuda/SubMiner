import type {
  AnkiConnectConfig,
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
} from '../../types/anki';
import type { BrowserWindow } from 'electron';
import type { WindowGeometry } from '../../types/runtime';
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
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
  getKnownWordCacheStatePath: () => string;
  shouldStartAnkiIntegration: () => boolean;
  bindOverlayOwner?: () => void;
  releaseOverlayOwner?: () => void;
};

export function createBuildInitializeOverlayRuntimeOptionsHandler(deps: {
  getBackendOverride: () => string | null;
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
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
  getKnownWordCacheStatePath: () => string;
  shouldStartAnkiIntegration: () => boolean;
  bindOverlayOwner?: () => void;
  releaseOverlayOwner?: () => void;
}) {
  return (): OverlayRuntimeOptions => ({
    backendOverride: deps.getBackendOverride(),
    createMainWindow: deps.createMainWindow,
    registerGlobalShortcuts: deps.registerGlobalShortcuts,
    updateVisibleOverlayBounds: deps.updateVisibleOverlayBounds,
    isVisibleOverlayVisible: deps.isVisibleOverlayVisible,
    updateVisibleOverlayVisibility: deps.updateVisibleOverlayVisibility,
    refreshCurrentSubtitle: deps.refreshCurrentSubtitle,
    getOverlayWindows: deps.getOverlayWindows,
    syncOverlayShortcuts: deps.syncOverlayShortcuts,
    setWindowTracker: deps.setWindowTracker,
    createWindowTracker: deps.createWindowTracker,
    getResolvedConfig: deps.getResolvedConfig,
    getSubtitleTimingTracker: deps.getSubtitleTimingTracker,
    getMpvClient: deps.getMpvClient,
    getMpvSocketPath: deps.getMpvSocketPath,
    getRuntimeOptionsManager: deps.getRuntimeOptionsManager,
    setAnkiIntegration: deps.setAnkiIntegration,
    showDesktopNotification: deps.showDesktopNotification,
    createFieldGroupingCallback: deps.createFieldGroupingCallback,
    getKnownWordCacheStatePath: deps.getKnownWordCacheStatePath,
    shouldStartAnkiIntegration: deps.shouldStartAnkiIntegration,
    bindOverlayOwner: deps.bindOverlayOwner,
    releaseOverlayOwner: deps.releaseOverlayOwner,
  });
}
