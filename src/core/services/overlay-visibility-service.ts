import { BrowserWindow, screen } from "electron";
import { BaseWindowTracker } from "../../window-trackers";
import { WindowGeometry } from "../../types";

interface MpvCommandSender {
  command: Array<string | number>;
  request_id?: number;
}

export function updateVisibleOverlayVisibilityService(args: {
  visibleOverlayVisible: boolean;
  mainWindow: BrowserWindow | null;
  windowTracker: BaseWindowTracker | null;
  trackerNotReadyWarningShown: boolean;
  setTrackerNotReadyWarningShown: (shown: boolean) => void;
  shouldBindVisibleOverlayToMpvSubVisibility: boolean;
  previousSecondarySubVisibility: boolean | null;
  setPreviousSecondarySubVisibility: (value: boolean | null) => void;
  mpvConnected: boolean;
  mpvSend: (payload: MpvCommandSender) => void;
  secondarySubVisibilityRequestId: number;
  updateOverlayBounds: (geometry: WindowGeometry) => void;
  ensureOverlayWindowLevel: (window: BrowserWindow) => void;
  enforceOverlayLayerOrder: () => void;
  syncOverlayShortcuts: () => void;
}): void {
  console.log(
    "updateVisibleOverlayVisibility called, visibleOverlayVisible:",
    args.visibleOverlayVisible,
  );
  if (!args.mainWindow || args.mainWindow.isDestroyed()) {
    console.log("mainWindow not available");
    return;
  }

  if (!args.visibleOverlayVisible) {
    console.log("Hiding visible overlay");
    args.mainWindow.hide();

    if (
      args.shouldBindVisibleOverlayToMpvSubVisibility &&
      args.previousSecondarySubVisibility !== null &&
      args.mpvConnected
    ) {
      args.mpvSend({
        command: [
          "set_property",
          "secondary-sub-visibility",
          args.previousSecondarySubVisibility ? "yes" : "no",
        ],
      });
      args.setPreviousSecondarySubVisibility(null);
    } else if (!args.shouldBindVisibleOverlayToMpvSubVisibility) {
      args.setPreviousSecondarySubVisibility(null);
    }
    args.syncOverlayShortcuts();
    return;
  }

  console.log(
    "Should show visible overlay, isTracking:",
    args.windowTracker?.isTracking(),
  );

  if (args.shouldBindVisibleOverlayToMpvSubVisibility && args.mpvConnected) {
    args.mpvSend({
      command: ["get_property", "secondary-sub-visibility"],
      request_id: args.secondarySubVisibilityRequestId,
    });
  }

  if (args.windowTracker && args.windowTracker.isTracking()) {
    args.setTrackerNotReadyWarningShown(false);
    const geometry = args.windowTracker.getGeometry();
    console.log("Geometry:", geometry);
    if (geometry) {
      args.updateOverlayBounds(geometry);
    }
    console.log("Showing visible overlay mainWindow");
    args.ensureOverlayWindowLevel(args.mainWindow);
    args.mainWindow.show();
    args.mainWindow.focus();
    args.enforceOverlayLayerOrder();
    args.syncOverlayShortcuts();
    return;
  }

  if (!args.windowTracker) {
    args.setTrackerNotReadyWarningShown(false);
    args.ensureOverlayWindowLevel(args.mainWindow);
    args.mainWindow.show();
    args.mainWindow.focus();
    args.enforceOverlayLayerOrder();
    args.syncOverlayShortcuts();
    return;
  }

  if (!args.trackerNotReadyWarningShown) {
    console.warn(
      "Window tracker exists but is not tracking yet; using fallback bounds until tracking starts",
    );
    args.setTrackerNotReadyWarningShown(true);
  }
  const cursorPoint = screen.getCursorScreenPoint();
  const display = screen.getDisplayNearestPoint(cursorPoint);
  const fallbackBounds = display.workArea;
  args.updateOverlayBounds({
    x: fallbackBounds.x,
    y: fallbackBounds.y,
    width: fallbackBounds.width,
    height: fallbackBounds.height,
  });
  args.ensureOverlayWindowLevel(args.mainWindow);
  args.mainWindow.show();
  args.mainWindow.focus();
  args.enforceOverlayLayerOrder();
  args.syncOverlayShortcuts();
}

export function updateInvisibleOverlayVisibilityService(args: {
  invisibleWindow: BrowserWindow | null;
  visibleOverlayVisible: boolean;
  invisibleOverlayVisible: boolean;
  windowTracker: BaseWindowTracker | null;
  updateOverlayBounds: (geometry: WindowGeometry) => void;
  ensureOverlayWindowLevel: (window: BrowserWindow) => void;
  enforceOverlayLayerOrder: () => void;
  syncOverlayShortcuts: () => void;
}): void {
  if (!args.invisibleWindow || args.invisibleWindow.isDestroyed()) {
    return;
  }

  if (args.visibleOverlayVisible) {
    args.invisibleWindow.hide();
    args.syncOverlayShortcuts();
    return;
  }

  const showInvisibleWithoutFocus = (): void => {
    args.ensureOverlayWindowLevel(args.invisibleWindow!);
    if (typeof args.invisibleWindow!.showInactive === "function") {
      args.invisibleWindow!.showInactive();
    } else {
      args.invisibleWindow!.show();
    }
    args.enforceOverlayLayerOrder();
  };

  if (!args.invisibleOverlayVisible) {
    args.invisibleWindow.hide();
    args.syncOverlayShortcuts();
    return;
  }

  if (args.windowTracker && args.windowTracker.isTracking()) {
    const geometry = args.windowTracker.getGeometry();
    if (geometry) {
      args.updateOverlayBounds(geometry);
    }
    showInvisibleWithoutFocus();
    args.syncOverlayShortcuts();
    return;
  }

  if (!args.windowTracker) {
    showInvisibleWithoutFocus();
    args.syncOverlayShortcuts();
    return;
  }

  const cursorPoint = screen.getCursorScreenPoint();
  const display = screen.getDisplayNearestPoint(cursorPoint);
  const fallbackBounds = display.workArea;
  args.updateOverlayBounds({
    x: fallbackBounds.x,
    y: fallbackBounds.y,
    width: fallbackBounds.width,
    height: fallbackBounds.height,
  });
  showInvisibleWithoutFocus();
  args.syncOverlayShortcuts();
}
