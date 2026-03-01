import type { BrowserWindow } from 'electron';
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
  syncPrimaryOverlayWindowLayer: (layer: 'visible') => void;
  enforceOverlayLayerOrder: () => void;
  syncOverlayShortcuts: () => void;
  isMacOSPlatform?: boolean;
  showOverlayLoadingOsd?: (message: string) => void;
  resolveFallbackBounds?: () => WindowGeometry;
}): void {
  if (!args.mainWindow || args.mainWindow.isDestroyed()) {
    return;
  }

  if (!args.visibleOverlayVisible) {
    args.setTrackerNotReadyWarningShown(false);
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
    args.syncPrimaryOverlayWindowLayer('visible');
    args.mainWindow.setIgnoreMouseEvents(false);
    args.ensureOverlayWindowLevel(args.mainWindow);
    args.mainWindow.show();
    args.mainWindow.focus();
    args.enforceOverlayLayerOrder();
    args.syncOverlayShortcuts();
    return;
  }

  if (!args.windowTracker) {
    if (args.isMacOSPlatform) {
      if (!args.trackerNotReadyWarningShown) {
        args.setTrackerNotReadyWarningShown(true);
        args.showOverlayLoadingOsd?.('Overlay loading...');
      }
      args.mainWindow.hide();
      args.syncOverlayShortcuts();
      return;
    }
    args.setTrackerNotReadyWarningShown(false);
    args.syncPrimaryOverlayWindowLayer('visible');
    args.mainWindow.setIgnoreMouseEvents(false);
    args.ensureOverlayWindowLevel(args.mainWindow);
    args.mainWindow.show();
    args.mainWindow.focus();
    args.enforceOverlayLayerOrder();
    args.syncOverlayShortcuts();
    return;
  }

  if (!args.trackerNotReadyWarningShown) {
    args.setTrackerNotReadyWarningShown(true);
    if (args.isMacOSPlatform) {
      args.showOverlayLoadingOsd?.('Overlay loading...');
    }
  }

  if (args.isMacOSPlatform) {
    args.mainWindow.hide();
    args.syncOverlayShortcuts();
    return;
  }

  const fallbackBounds = args.resolveFallbackBounds?.();
  if (!fallbackBounds) return;

  args.updateVisibleOverlayBounds(fallbackBounds);
  args.syncPrimaryOverlayWindowLayer('visible');
  args.mainWindow.setIgnoreMouseEvents(false);
  args.ensureOverlayWindowLevel(args.mainWindow);
  args.mainWindow.show();
  args.mainWindow.focus();
  args.enforceOverlayLayerOrder();
  args.syncOverlayShortcuts();
}

export function setVisibleOverlayVisible(options: {
  visible: boolean;
  setVisibleOverlayVisibleState: (visible: boolean) => void;
  updateVisibleOverlayVisibility: () => void;
}): void {
  options.setVisibleOverlayVisibleState(options.visible);
  options.updateVisibleOverlayVisibility();
}
