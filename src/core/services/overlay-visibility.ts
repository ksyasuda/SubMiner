import { BrowserWindow, screen } from 'electron';
import { BaseWindowTracker } from '../../window-trackers';
import { WindowGeometry } from '../../types';

export function updateVisibleOverlayVisibility(args: {
  visibleOverlayVisible: boolean;
  mainWindow: BrowserWindow | null;
  windowTracker: BaseWindowTracker | null;
  trackerNotReadyWarningShown: boolean;
  setTrackerNotReadyWarningShown: (shown: boolean) => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  ensureOverlayWindowLevel: (window: BrowserWindow) => void;
  enforceOverlayLayerOrder: () => void;
  syncOverlayShortcuts: () => void;
}): void {
  if (!args.mainWindow || args.mainWindow.isDestroyed()) {
    return;
  }

  if (!args.visibleOverlayVisible) {
    args.mainWindow.hide();
    args.syncOverlayShortcuts();
    return;
  }

  if (args.windowTracker && args.windowTracker.isTracking()) {
    args.setTrackerNotReadyWarningShown(false);
    const geometry = args.windowTracker.getGeometry();
    if (geometry) {
      args.updateVisibleOverlayBounds(geometry);
    }
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
    args.setTrackerNotReadyWarningShown(true);
  }
  const cursorPoint = screen.getCursorScreenPoint();
  const display = screen.getDisplayNearestPoint(cursorPoint);
  const fallbackBounds = display.workArea;
  args.updateVisibleOverlayBounds({
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

export function updateInvisibleOverlayVisibility(args: {
  invisibleWindow: BrowserWindow | null;
  visibleOverlayVisible: boolean;
  invisibleOverlayVisible: boolean;
  windowTracker: BaseWindowTracker | null;
  updateInvisibleOverlayBounds: (geometry: WindowGeometry) => void;
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
    if (typeof args.invisibleWindow!.showInactive === 'function') {
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
      args.updateInvisibleOverlayBounds(geometry);
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
  args.updateInvisibleOverlayBounds({
    x: fallbackBounds.x,
    y: fallbackBounds.y,
    width: fallbackBounds.width,
    height: fallbackBounds.height,
  });
  showInvisibleWithoutFocus();
  args.syncOverlayShortcuts();
}

export function syncInvisibleOverlayMousePassthrough(options: {
  hasInvisibleWindow: () => boolean;
  setIgnoreMouseEvents: (ignore: boolean, extra?: { forward: boolean }) => void;
  visibleOverlayVisible: boolean;
  invisibleOverlayVisible: boolean;
}): void {
  if (!options.hasInvisibleWindow()) return;
  if (options.visibleOverlayVisible) {
    options.setIgnoreMouseEvents(true, { forward: true });
  } else if (options.invisibleOverlayVisible) {
    options.setIgnoreMouseEvents(false);
  }
}

export function setVisibleOverlayVisible(options: {
  visible: boolean;
  setVisibleOverlayVisibleState: (visible: boolean) => void;
  updateVisibleOverlayVisibility: () => void;
  updateInvisibleOverlayVisibility: () => void;
  syncInvisibleOverlayMousePassthrough: () => void;
  shouldBindVisibleOverlayToMpvSubVisibility: () => boolean;
  isMpvConnected: () => boolean;
  setMpvSubVisibility: (visible: boolean) => void;
}): void {
  options.setVisibleOverlayVisibleState(options.visible);
  options.updateVisibleOverlayVisibility();
  options.updateInvisibleOverlayVisibility();
  options.syncInvisibleOverlayMousePassthrough();
  if (options.shouldBindVisibleOverlayToMpvSubVisibility() && options.isMpvConnected()) {
    options.setMpvSubVisibility(!options.visible);
  }
}

export function setInvisibleOverlayVisible(options: {
  visible: boolean;
  setInvisibleOverlayVisibleState: (visible: boolean) => void;
  updateInvisibleOverlayVisibility: () => void;
  syncInvisibleOverlayMousePassthrough: () => void;
}): void {
  options.setInvisibleOverlayVisibleState(options.visible);
  options.updateInvisibleOverlayVisibility();
  options.syncInvisibleOverlayMousePassthrough();
}
