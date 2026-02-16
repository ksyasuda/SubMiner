import type { BrowserWindow } from "electron";

import type { BaseWindowTracker } from "../window-trackers";
import type { WindowGeometry } from "../types";
import {
  syncInvisibleOverlayMousePassthroughService,
  updateInvisibleOverlayVisibilityService,
  updateVisibleOverlayVisibilityService,
} from "../core/services";

export interface OverlayVisibilityRuntimeDeps {
  getMainWindow: () => BrowserWindow | null;
  getInvisibleWindow: () => BrowserWindow | null;
  getVisibleOverlayVisible: () => boolean;
  getInvisibleOverlayVisible: () => boolean;
  getWindowTracker: () => BaseWindowTracker | null;
  getTrackerNotReadyWarningShown: () => boolean;
  setTrackerNotReadyWarningShown: (shown: boolean) => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  updateInvisibleOverlayBounds: (geometry: WindowGeometry) => void;
  ensureOverlayWindowLevel: (window: BrowserWindow) => void;
  enforceOverlayLayerOrder: () => void;
  syncOverlayShortcuts: () => void;
}

export interface OverlayVisibilityRuntimeService {
  updateVisibleOverlayVisibility: () => void;
  updateInvisibleOverlayVisibility: () => void;
  syncInvisibleOverlayMousePassthrough: () => void;
}

export function createOverlayVisibilityRuntimeService(
  deps: OverlayVisibilityRuntimeDeps,
): OverlayVisibilityRuntimeService {
  const hasInvisibleWindow = (): boolean => {
    const invisibleWindow = deps.getInvisibleWindow();
    return Boolean(invisibleWindow && !invisibleWindow.isDestroyed());
  };

  const setIgnoreMouseEvents = (
    ignore: boolean,
    options?: Parameters<BrowserWindow["setIgnoreMouseEvents"]>[1],
  ): void => {
    const invisibleWindow = deps.getInvisibleWindow();
    if (!invisibleWindow || invisibleWindow.isDestroyed()) return;
    invisibleWindow.setIgnoreMouseEvents(ignore, options);
  };

  return {
    updateVisibleOverlayVisibility(): void {
      updateVisibleOverlayVisibilityService({
        visibleOverlayVisible: deps.getVisibleOverlayVisible(),
        mainWindow: deps.getMainWindow(),
        windowTracker: deps.getWindowTracker(),
        trackerNotReadyWarningShown: deps.getTrackerNotReadyWarningShown(),
        setTrackerNotReadyWarningShown: (shown: boolean) => {
          deps.setTrackerNotReadyWarningShown(shown);
        },
        updateVisibleOverlayBounds: (geometry: WindowGeometry) =>
          deps.updateVisibleOverlayBounds(geometry),
        ensureOverlayWindowLevel: (window: BrowserWindow) =>
          deps.ensureOverlayWindowLevel(window),
        enforceOverlayLayerOrder: () => deps.enforceOverlayLayerOrder(),
        syncOverlayShortcuts: () => deps.syncOverlayShortcuts(),
      });
    },

    updateInvisibleOverlayVisibility(): void {
      updateInvisibleOverlayVisibilityService({
        invisibleWindow: deps.getInvisibleWindow(),
        visibleOverlayVisible: deps.getVisibleOverlayVisible(),
        invisibleOverlayVisible: deps.getInvisibleOverlayVisible(),
        windowTracker: deps.getWindowTracker(),
        updateInvisibleOverlayBounds: (geometry: WindowGeometry) =>
          deps.updateInvisibleOverlayBounds(geometry),
        ensureOverlayWindowLevel: (window: BrowserWindow) =>
          deps.ensureOverlayWindowLevel(window),
        enforceOverlayLayerOrder: () => deps.enforceOverlayLayerOrder(),
        syncOverlayShortcuts: () => deps.syncOverlayShortcuts(),
      });
    },

    syncInvisibleOverlayMousePassthrough(): void {
      syncInvisibleOverlayMousePassthroughService({
        hasInvisibleWindow,
        setIgnoreMouseEvents,
        visibleOverlayVisible: deps.getVisibleOverlayVisible(),
        invisibleOverlayVisible: deps.getInvisibleOverlayVisible(),
      });
    },
  };
}
