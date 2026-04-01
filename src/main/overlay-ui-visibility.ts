import type { WindowGeometry } from '../types';

export type OverlayUiVisibilityBridgeWindowLike = {
  isDestroyed: () => boolean;
  hide?: () => void;
  show?: () => void;
  focus?: () => void;
  setIgnoreMouseEvents?: (ignore: boolean, options?: { forward?: boolean }) => void;
};

export interface OverlayUiVisibilityBridgeInput<
  TWindow extends OverlayUiVisibilityBridgeWindowLike = OverlayUiVisibilityBridgeWindowLike,
> {
  getMainWindow: () => TWindow | null;
  getVisibleOverlayVisible: () => boolean;
  getModalActive: () => boolean;
  getForceMousePassthrough: () => boolean;
  getWindowTracker: () => unknown;
  getTrackerNotReadyWarningShown: () => boolean;
  setTrackerNotReadyWarningShown: (shown: boolean) => void;
  updateVisibleOverlayBounds: (geometry: WindowGeometry) => void;
  ensureOverlayWindowLevel: (window: TWindow) => void;
  syncPrimaryOverlayWindowLayer: (layer: 'visible') => void;
  enforceOverlayLayerOrder: () => void;
  syncOverlayShortcuts: () => void;
  isMacOSPlatform: () => boolean;
  isWindowsPlatform: () => boolean;
  showOverlayLoadingOsd: (message: string) => void;
}

export function createOverlayVisibilityRuntimeBridge<
  TWindow extends OverlayUiVisibilityBridgeWindowLike,
>(input: OverlayUiVisibilityBridgeInput<TWindow>) {
  let lastOverlayLoadingOsdAtMs: number | null = null;

  return {
    updateVisibleOverlayVisibility(): void {
      const mainWindow = input.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed()) {
        return;
      }

      if (input.getModalActive()) {
        mainWindow.hide?.();
        input.syncOverlayShortcuts();
        return;
      }

      const showPassiveVisibleOverlay = (): void => {
        const forceMousePassthrough = input.getForceMousePassthrough() === true;
        if (input.isWindowsPlatform() || forceMousePassthrough) {
          mainWindow.setIgnoreMouseEvents?.(true, { forward: true });
        } else {
          mainWindow.setIgnoreMouseEvents?.(false);
        }
        input.ensureOverlayWindowLevel(mainWindow);
        mainWindow.show?.();
        if (!input.isWindowsPlatform() && !input.isMacOSPlatform() && !forceMousePassthrough) {
          mainWindow.focus?.();
        }
      };

      const maybeShowOverlayLoadingOsd = (): void => {
        if (!input.isMacOSPlatform()) {
          return;
        }
        if (lastOverlayLoadingOsdAtMs !== null && Date.now() - lastOverlayLoadingOsdAtMs < 30_000) {
          return;
        }
        input.showOverlayLoadingOsd('Overlay loading...');
        lastOverlayLoadingOsdAtMs = Date.now();
      };

      if (!input.getVisibleOverlayVisible()) {
        input.setTrackerNotReadyWarningShown(false);
        lastOverlayLoadingOsdAtMs = null;
        mainWindow.hide?.();
        input.syncOverlayShortcuts();
        return;
      }

      const windowTracker = input.getWindowTracker() as {
        isTracking: () => boolean;
        getGeometry: () => WindowGeometry | null;
      } | null;

      if (windowTracker && windowTracker.isTracking()) {
        input.setTrackerNotReadyWarningShown(false);
        const geometry = windowTracker.getGeometry();
        if (geometry) {
          input.updateVisibleOverlayBounds(geometry);
        }
        input.syncPrimaryOverlayWindowLayer('visible');
        showPassiveVisibleOverlay();
        input.enforceOverlayLayerOrder();
        input.syncOverlayShortcuts();
        return;
      }

      if (!windowTracker) {
        if (input.isMacOSPlatform() || input.isWindowsPlatform()) {
          if (!input.getTrackerNotReadyWarningShown()) {
            input.setTrackerNotReadyWarningShown(true);
            maybeShowOverlayLoadingOsd();
          }
          mainWindow.hide?.();
          input.syncOverlayShortcuts();
          return;
        }

        input.setTrackerNotReadyWarningShown(false);
        input.syncPrimaryOverlayWindowLayer('visible');
        showPassiveVisibleOverlay();
        input.enforceOverlayLayerOrder();
        input.syncOverlayShortcuts();
        return;
      }

      if (!input.getTrackerNotReadyWarningShown()) {
        input.setTrackerNotReadyWarningShown(true);
        maybeShowOverlayLoadingOsd();
      }

      mainWindow.hide?.();
      input.syncOverlayShortcuts();
    },
  };
}
