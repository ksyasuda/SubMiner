import { BrowserWindow } from "electron";
import { BaseWindowTracker } from "../../window-trackers";
import {
  AnkiConnectConfig,
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
  WindowGeometry,
} from "../../types";
import { createOverlayWindowService } from "./overlay-window-service";
import { initializeOverlayRuntimeService } from "./overlay-runtime-init-service";
import {
  updateInvisibleOverlayVisibilityService,
  updateVisibleOverlayVisibilityService,
} from "./overlay-visibility-service";

interface MpvCommandSender {
  command: Array<string | number>;
  request_id?: number;
}

export interface OverlayWindowRuntimeDepsOptions {
  isDev: boolean;
  getOverlayDebugVisualizationEnabled: () => boolean;
  ensureOverlayWindowLevel: (window: BrowserWindow) => void;
  onRuntimeOptionsChanged: () => void;
  setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
  getVisibleOverlayVisible: () => boolean;
  getInvisibleOverlayVisible: () => boolean;
  tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
  onWindowClosed: (kind: "visible" | "invisible") => void;
}

export interface InitializeOverlayRuntimeDepsOptions {
  backendOverride: string | null;
  getInitialInvisibleOverlayVisibility: () => boolean;
  createMainWindow: () => void;
  createInvisibleWindow: () => void;
  registerGlobalShortcuts: () => void;
  updateOverlayBounds: (geometry: WindowGeometry) => void;
  isVisibleOverlayVisible: () => boolean;
  isInvisibleOverlayVisible: () => boolean;
  updateVisibleOverlayVisibility: () => void;
  updateInvisibleOverlayVisibility: () => void;
  getOverlayWindows: () => BrowserWindow[];
  syncOverlayShortcuts: () => void;
  setWindowTracker: (tracker: BaseWindowTracker | null) => void;
  getResolvedConfig: () => { ankiConnect?: AnkiConnectConfig };
  getSubtitleTimingTracker: () => unknown | null;
  getMpvClient: () => { send?: (payload: { command: string[] }) => void } | null;
  getRuntimeOptionsManager: () => {
    getEffectiveAnkiConnectConfig: (config?: AnkiConnectConfig) => AnkiConnectConfig;
  } | null;
  setAnkiIntegration: (integration: unknown | null) => void;
  showDesktopNotification: (title: string, options: { body?: string; icon?: string }) => void;
  createFieldGroupingCallback: () => (
    data: KikuFieldGroupingRequestData,
  ) => Promise<KikuFieldGroupingChoice>;
}

export interface VisibleOverlayVisibilityDepsRuntimeOptions {
  getVisibleOverlayVisible: () => boolean;
  getMainWindow: () => BrowserWindow | null;
  getWindowTracker: () => BaseWindowTracker | null;
  getTrackerNotReadyWarningShown: () => boolean;
  setTrackerNotReadyWarningShown: (shown: boolean) => void;
  shouldBindVisibleOverlayToMpvSubVisibility: () => boolean;
  getPreviousSecondarySubVisibility: () => boolean | null;
  setPreviousSecondarySubVisibility: (value: boolean | null) => void;
  isMpvConnected: () => boolean;
  mpvSend: (payload: MpvCommandSender) => void;
  secondarySubVisibilityRequestId: number;
  updateOverlayBounds: (geometry: WindowGeometry) => void;
  ensureOverlayWindowLevel: (window: BrowserWindow) => void;
  enforceOverlayLayerOrder: () => void;
  syncOverlayShortcuts: () => void;
}

export interface InvisibleOverlayVisibilityDepsRuntimeOptions {
  getInvisibleWindow: () => BrowserWindow | null;
  getVisibleOverlayVisible: () => boolean;
  getInvisibleOverlayVisible: () => boolean;
  getWindowTracker: () => BaseWindowTracker | null;
  updateOverlayBounds: (geometry: WindowGeometry) => void;
  ensureOverlayWindowLevel: (window: BrowserWindow) => void;
  enforceOverlayLayerOrder: () => void;
  syncOverlayShortcuts: () => void;
}

export function createOverlayWindowRuntimeDepsService(
  options: OverlayWindowRuntimeDepsOptions,
): Parameters<typeof createOverlayWindowService>[1] {
  return {
    isDev: options.isDev,
    overlayDebugVisualizationEnabled:
      options.getOverlayDebugVisualizationEnabled(),
    ensureOverlayWindowLevel: options.ensureOverlayWindowLevel,
    onRuntimeOptionsChanged: options.onRuntimeOptionsChanged,
    setOverlayDebugVisualizationEnabled:
      options.setOverlayDebugVisualizationEnabled,
    isOverlayVisible: (kind) =>
      kind === "visible"
        ? options.getVisibleOverlayVisible()
        : options.getInvisibleOverlayVisible(),
    tryHandleOverlayShortcutLocalFallback:
      options.tryHandleOverlayShortcutLocalFallback,
    onWindowClosed: options.onWindowClosed,
  };
}

export function createInitializeOverlayRuntimeDepsService(
  options: InitializeOverlayRuntimeDepsOptions,
): Parameters<typeof initializeOverlayRuntimeService>[0] {
  return {
    backendOverride: options.backendOverride,
    getInitialInvisibleOverlayVisibility:
      options.getInitialInvisibleOverlayVisibility,
    createMainWindow: options.createMainWindow,
    createInvisibleWindow: options.createInvisibleWindow,
    registerGlobalShortcuts: options.registerGlobalShortcuts,
    updateOverlayBounds: options.updateOverlayBounds,
    isVisibleOverlayVisible: options.isVisibleOverlayVisible,
    isInvisibleOverlayVisible: options.isInvisibleOverlayVisible,
    updateVisibleOverlayVisibility: options.updateVisibleOverlayVisibility,
    updateInvisibleOverlayVisibility: options.updateInvisibleOverlayVisibility,
    getOverlayWindows: options.getOverlayWindows,
    syncOverlayShortcuts: options.syncOverlayShortcuts,
    setWindowTracker: options.setWindowTracker,
    getResolvedConfig: options.getResolvedConfig,
    getSubtitleTimingTracker: options.getSubtitleTimingTracker,
    getMpvClient: options.getMpvClient,
    getRuntimeOptionsManager: options.getRuntimeOptionsManager,
    setAnkiIntegration: options.setAnkiIntegration,
    showDesktopNotification: options.showDesktopNotification,
    createFieldGroupingCallback: options.createFieldGroupingCallback,
  };
}

export function createVisibleOverlayVisibilityDepsRuntimeService(
  options: VisibleOverlayVisibilityDepsRuntimeOptions,
): Parameters<typeof updateVisibleOverlayVisibilityService>[0] {
  return {
    visibleOverlayVisible: options.getVisibleOverlayVisible(),
    mainWindow: options.getMainWindow(),
    windowTracker: options.getWindowTracker(),
    trackerNotReadyWarningShown: options.getTrackerNotReadyWarningShown(),
    setTrackerNotReadyWarningShown: options.setTrackerNotReadyWarningShown,
    shouldBindVisibleOverlayToMpvSubVisibility:
      options.shouldBindVisibleOverlayToMpvSubVisibility(),
    previousSecondarySubVisibility: options.getPreviousSecondarySubVisibility(),
    setPreviousSecondarySubVisibility: options.setPreviousSecondarySubVisibility,
    mpvConnected: options.isMpvConnected(),
    mpvSend: options.mpvSend,
    secondarySubVisibilityRequestId: options.secondarySubVisibilityRequestId,
    updateOverlayBounds: options.updateOverlayBounds,
    ensureOverlayWindowLevel: options.ensureOverlayWindowLevel,
    enforceOverlayLayerOrder: options.enforceOverlayLayerOrder,
    syncOverlayShortcuts: options.syncOverlayShortcuts,
  };
}

export function createInvisibleOverlayVisibilityDepsRuntimeService(
  options: InvisibleOverlayVisibilityDepsRuntimeOptions,
): Parameters<typeof updateInvisibleOverlayVisibilityService>[0] {
  return {
    invisibleWindow: options.getInvisibleWindow(),
    visibleOverlayVisible: options.getVisibleOverlayVisible(),
    invisibleOverlayVisible: options.getInvisibleOverlayVisible(),
    windowTracker: options.getWindowTracker(),
    updateOverlayBounds: options.updateOverlayBounds,
    ensureOverlayWindowLevel: options.ensureOverlayWindowLevel,
    enforceOverlayLayerOrder: options.enforceOverlayLayerOrder,
    syncOverlayShortcuts: options.syncOverlayShortcuts,
  };
}
