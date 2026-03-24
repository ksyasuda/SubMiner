import type { BrowserWindow } from 'electron';
import { BaseWindowTracker } from '../../window-trackers';
import { WindowGeometry } from '../../types';

export function updateVisibleOverlayVisibility(args: {
  visibleOverlayVisible: boolean;
  modalActive?: boolean;
  forceMousePassthrough?: boolean;
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
  isWindowsPlatform?: boolean;
  showOverlayLoadingOsd?: (message: string) => void;
  shouldShowOverlayLoadingOsd?: () => boolean;
  markOverlayLoadingOsdShown?: () => void;
  resetOverlayLoadingOsdSuppression?: () => void;
  resolveFallbackBounds?: () => WindowGeometry;
}): void {
  if (!args.mainWindow || args.mainWindow.isDestroyed()) {
    return;
  }

  const mainWindow = args.mainWindow;

  if (args.modalActive) {
    mainWindow.hide();
    args.syncOverlayShortcuts();
    return;
  }

  const showPassiveVisibleOverlay = (): void => {
    const forceMousePassthrough = args.forceMousePassthrough === true;
    if (args.isWindowsPlatform || forceMousePassthrough) {
      mainWindow.setIgnoreMouseEvents(true, { forward: true });
    } else {
      mainWindow.setIgnoreMouseEvents(false);
    }
    args.ensureOverlayWindowLevel(mainWindow);
    mainWindow.show();
    if (!args.isWindowsPlatform && !args.isMacOSPlatform && !forceMousePassthrough) {
      mainWindow.focus();
    }
  };

  const maybeShowOverlayLoadingOsd = (): void => {
    if (!args.isMacOSPlatform || !args.showOverlayLoadingOsd) {
      return;
    }
    if (args.shouldShowOverlayLoadingOsd && !args.shouldShowOverlayLoadingOsd()) {
      return;
    }
    args.showOverlayLoadingOsd('Overlay loading...');
    args.markOverlayLoadingOsdShown?.();
  };

  if (!args.visibleOverlayVisible) {
    args.setTrackerNotReadyWarningShown(false);
    args.resetOverlayLoadingOsdSuppression?.();
    mainWindow.hide();
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
    showPassiveVisibleOverlay();
    args.enforceOverlayLayerOrder();
    args.syncOverlayShortcuts();
    return;
  }

  if (!args.windowTracker) {
    if (args.isMacOSPlatform || args.isWindowsPlatform) {
      if (!args.trackerNotReadyWarningShown) {
        args.setTrackerNotReadyWarningShown(true);
        maybeShowOverlayLoadingOsd();
      }
      mainWindow.hide();
      args.syncOverlayShortcuts();
      return;
    }
    args.setTrackerNotReadyWarningShown(false);
    args.syncPrimaryOverlayWindowLayer('visible');
    showPassiveVisibleOverlay();
    args.enforceOverlayLayerOrder();
    args.syncOverlayShortcuts();
    return;
  }

  if (!args.trackerNotReadyWarningShown) {
    args.setTrackerNotReadyWarningShown(true);
    maybeShowOverlayLoadingOsd();
  }

  mainWindow.hide();
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
