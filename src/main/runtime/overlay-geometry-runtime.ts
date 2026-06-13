import { type BrowserWindow, screen } from 'electron';
import type { WindowGeometry } from '../../types';
import { hasHyprlandWindowPlacementBoundsMismatch } from '../../core/services/hyprland-window-placement';
import { normalizeOverlayWindowBoundsForPlatform } from '../../core/services/overlay-window-bounds';
import {
  enforceOverlayLayerOrder as enforceOverlayLayerOrderCore,
  ensureOverlayWindowLevel as ensureOverlayWindowLevelCore,
  syncOverlayWindowLayer,
} from '../../core/services/overlay-window';
import { promoteStatsOverlayAbovePlayback } from '../../core/services/stats-window.js';
import { restoreLinuxOverlayWindowShape } from './linux-overlay-window-shape';
import { shouldRunLinuxOverlayZOrderKeepAlive } from './linux-overlay-zorder-keepalive';
import {
  shouldExitFullscreenOverrideForTrackedGeometry,
  type LinuxVisibleOverlayWindowMode,
} from './linux-visible-overlay-window-mode';
import {
  createEnforceOverlayLayerOrderHandler,
  createEnsureOverlayWindowLevelHandler,
  createUpdateVisibleOverlayBoundsHandler,
  hasLiveOverlayWindowBoundsMismatch,
} from './overlay-window-layout';
import {
  createBuildEnforceOverlayLayerOrderMainDepsHandler,
  createBuildEnsureOverlayWindowLevelMainDepsHandler,
  createBuildUpdateVisibleOverlayBoundsMainDepsHandler,
} from './overlay-window-layout-main-deps';
import { shouldSuppressVisibleOverlayRaiseForSeparateWindow } from './settings-window-z-order';

const LINUX_VISIBLE_OVERLAY_FULLSCREEN_GEOMETRY_GRACE_MS = 1_200;

export interface OverlayGeometryRuntimeDeps {
  overlayManager: {
    getMainWindow: () => BrowserWindow | null;
    getModalWindow: () => BrowserWindow | null;
    getVisibleOverlayVisible: () => boolean;
    setOverlayWindowBounds: (geometry: WindowGeometry) => void;
    setModalWindowBounds: (geometry: WindowGeometry) => void;
  };
  getTrackedWindowGeometry: () => WindowGeometry | null;
  getTrackedWindowMediaSourceId: () => string | null | undefined;
  getTrackedWindowNativeId: () => string | null | undefined;
  getStatsOverlayVisible: () => boolean;
  getOverlayForegroundSeparateWindows: () => BrowserWindow[];
  getLinuxVisibleOverlayWindowMode: () => LinuxVisibleOverlayWindowMode;
  getLinuxTrackedMpvFullscreen: () => boolean;
  getLinuxTrackedMpvFullscreenChangedAtMs: () => number;
  syncLinuxVisibleOverlayMpvFullscreenMode: (fullscreen: boolean) => void;
  getLinuxVisibleOverlayOwnerBindingKey: () => string | null;
  setLinuxVisibleOverlayOwnerBindingKey: (key: string | null) => void;
  clearVisibleOverlayX11OwnerBinding: (window: BrowserWindow) => void;
  getNativeWindowHandleDecimal: (window: BrowserWindow) => string;
  enqueueVisibleOverlayX11OwnerBindingOperation: (
    window: BrowserWindow,
    args: string[],
    onError?: (error: Error) => void,
  ) => void;
  scheduleWindowsVisibleOverlayZOrderSyncBurst: () => void;
  logDebug: (message: string, ...args: unknown[]) => void;
}

export function createOverlayGeometryRuntime(deps: OverlayGeometryRuntimeDeps) {
  const { overlayManager } = deps;

  let lastOverlayWindowGeometry: WindowGeometry | null = null;

  function getOverlayGeometryFallback(): WindowGeometry {
    const cursorPoint = screen.getCursorScreenPoint();
    const display = screen.getDisplayNearestPoint(cursorPoint);
    const bounds = display.workArea;
    return {
      x: bounds.x,
      y: bounds.y,
      width: bounds.width,
      height: bounds.height,
    };
  }

  function getCurrentOverlayGeometry(): WindowGeometry {
    if (lastOverlayWindowGeometry) return lastOverlayWindowGeometry;
    const trackerGeometry = deps.getTrackedWindowGeometry();
    if (trackerGeometry) return trackerGeometry;
    return getOverlayGeometryFallback();
  }

  function getCurrentTrackedOverlayGeometry(): WindowGeometry | null {
    return deps.getTrackedWindowGeometry();
  }

  function geometryMatches(a: WindowGeometry | null, b: WindowGeometry | null): boolean {
    if (!a || !b) return false;
    return a.x === b.x && a.y === b.y && a.width === b.width && a.height === b.height;
  }

  function applyOverlayRegions(geometry: WindowGeometry): void {
    lastOverlayWindowGeometry = geometry;
    maybeExitLinuxFullscreenOverrideForTrackedGeometry(geometry);
    overlayManager.setOverlayWindowBounds(geometry);
    overlayManager.setModalWindowBounds(geometry);
  }

  function shouldExitLinuxFullscreenOverrideForGeometry(geometry: WindowGeometry): boolean {
    if (!shouldRunLinuxOverlayZOrderKeepAlive()) {
      return false;
    }
    if (
      deps.getLinuxTrackedMpvFullscreenChangedAtMs() > 0 &&
      Date.now() - deps.getLinuxTrackedMpvFullscreenChangedAtMs() <
        LINUX_VISIBLE_OVERLAY_FULLSCREEN_GEOMETRY_GRACE_MS
    ) {
      return false;
    }

    const displayBounds = screen.getDisplayMatching(geometry).bounds;
    return shouldExitFullscreenOverrideForTrackedGeometry({
      currentMode: deps.getLinuxVisibleOverlayWindowMode(),
      trackedFullscreen: deps.getLinuxTrackedMpvFullscreen(),
      geometry,
      displayBounds,
    });
  }

  function maybeExitLinuxFullscreenOverrideForTrackedGeometry(geometry: WindowGeometry): void {
    if (!shouldExitLinuxFullscreenOverrideForGeometry(geometry)) {
      return;
    }

    deps.logDebug(
      'Tracked mpv geometry no longer covers its display; exiting Linux fullscreen overlay override',
    );
    deps.syncLinuxVisibleOverlayMpvFullscreenMode(false);
  }

  function hasHyprlandOverlayWindowPlacementMismatch(geometry: WindowGeometry): boolean {
    if (process.platform !== 'linux') {
      return false;
    }

    return [overlayManager.getMainWindow(), overlayManager.getModalWindow()].some((window) => {
      if (!window || window.isDestroyed()) {
        return false;
      }
      return hasHyprlandWindowPlacementBoundsMismatch({
        title: window.getTitle(),
        bounds: normalizeOverlayWindowBoundsForPlatform(geometry, process.platform, screen, window),
      });
    });
  }

  const buildUpdateVisibleOverlayBoundsMainDepsHandler =
    createBuildUpdateVisibleOverlayBoundsMainDepsHandler({
      getCurrentOverlayWindowBounds: () => lastOverlayWindowGeometry,
      shouldRefreshUnchangedGeometry: (geometry) =>
        shouldExitLinuxFullscreenOverrideForGeometry(geometry) ||
        (process.platform === 'linux' &&
          (hasLiveOverlayWindowBoundsMismatch(
            [overlayManager.getMainWindow(), overlayManager.getModalWindow()],
            geometry,
          ) ||
            hasHyprlandOverlayWindowPlacementMismatch(geometry))),
      setOverlayWindowBounds: (geometry) => applyOverlayRegions(geometry),
      afterSetOverlayWindowBounds: () => {
        if (!overlayManager.getVisibleOverlayVisible()) {
          return;
        }
        if (process.platform === 'win32') {
          deps.scheduleWindowsVisibleOverlayZOrderSyncBurst();
          return;
        }
        const mainWindow = overlayManager.getMainWindow();
        if (!mainWindow || mainWindow.isDestroyed()) {
          return;
        }
        if (process.platform === 'linux') {
          restoreLinuxOverlayWindowShape(mainWindow);
        }
        ensureOverlayWindowLevel(mainWindow);
      },
    });
  const updateVisibleOverlayBoundsMainDeps = buildUpdateVisibleOverlayBoundsMainDepsHandler();
  const updateVisibleOverlayBounds = createUpdateVisibleOverlayBoundsHandler(
    updateVisibleOverlayBoundsMainDeps,
  );

  const buildEnsureOverlayWindowLevelMainDepsHandler =
    createBuildEnsureOverlayWindowLevelMainDepsHandler({
      shouldSuppressOverlayWindowLevel: (window) => {
        const mainWindow = overlayManager.getMainWindow();
        return (
          (deps.getStatsOverlayVisible() && window === mainWindow) ||
          shouldSuppressVisibleOverlayRaiseForSeparateWindow({
            window,
            mainWindow,
            separateWindows: deps.getOverlayForegroundSeparateWindows(),
          })
        );
      },
      ensureOverlayWindowLevelCore: (window) =>
        ensureOverlayWindowLevelCore(window as BrowserWindow),
      afterEnsureOverlayWindowLevel: () => {
        const mainWindow = overlayManager.getMainWindow();
        if (mainWindow && !mainWindow.isDestroyed()) {
          moveVisibleOverlayAboveTrackedPlaybackWindow(mainWindow);
        }
        promoteStatsOverlayAbovePlayback();
      },
    });
  const ensureOverlayWindowLevelMainDeps = buildEnsureOverlayWindowLevelMainDepsHandler();
  const ensureOverlayWindowLevel = createEnsureOverlayWindowLevelHandler(
    ensureOverlayWindowLevelMainDeps,
  );

  function syncPrimaryOverlayWindowLayer(layer: 'visible'): void {
    const mainWindow = overlayManager.getMainWindow();
    if (!mainWindow || mainWindow.isDestroyed()) return;
    syncOverlayWindowLayer(mainWindow, layer);
  }

  function moveVisibleOverlayAboveTrackedPlaybackWindow(window: BrowserWindow): void {
    if (process.platform !== 'linux') return;
    if (window !== overlayManager.getMainWindow()) return;

    bindVisibleOverlayToTrackedX11Window(window);

    const mediaSourceId = deps.getTrackedWindowMediaSourceId();
    if (!mediaSourceId) return;

    try {
      window.moveAbove(mediaSourceId);
    } catch (error) {
      deps.logDebug(
        'Failed to move visible overlay above tracked playback window:',
        error instanceof Error ? error.message : String(error),
      );
    }
  }

  function bindVisibleOverlayToTrackedX11Window(window: BrowserWindow): void {
    const targetWindowId = deps.getTrackedWindowNativeId();
    if (!targetWindowId) {
      if (deps.getLinuxVisibleOverlayOwnerBindingKey() !== null) {
        deps.clearVisibleOverlayX11OwnerBinding(window);
      }
      deps.setLinuxVisibleOverlayOwnerBindingKey(null);
      return;
    }

    const overlayWindowId = deps.getNativeWindowHandleDecimal(window);
    const bindingKey = `${overlayWindowId}:${targetWindowId}`;
    if (deps.getLinuxVisibleOverlayOwnerBindingKey() === bindingKey) {
      return;
    }
    deps.setLinuxVisibleOverlayOwnerBindingKey(bindingKey);

    deps.enqueueVisibleOverlayX11OwnerBindingOperation(
      window,
      [
        '-id',
        overlayWindowId,
        '-f',
        'WM_TRANSIENT_FOR',
        '32x',
        '-set',
        'WM_TRANSIENT_FOR',
        targetWindowId,
      ],
      (error) => {
        if (deps.getLinuxVisibleOverlayOwnerBindingKey() === bindingKey) {
          deps.setLinuxVisibleOverlayOwnerBindingKey(null);
        }
        deps.logDebug(
          'Failed to bind visible overlay as transient for tracked X11 playback window:',
          error instanceof Error ? error.message : String(error),
        );
      },
    );
  }

  const buildEnforceOverlayLayerOrderMainDepsHandler =
    createBuildEnforceOverlayLayerOrderMainDepsHandler({
      enforceOverlayLayerOrderCore: (params) =>
        enforceOverlayLayerOrderCore({
          visibleOverlayVisible: params.visibleOverlayVisible,
          mainWindow: params.mainWindow as BrowserWindow | null,
          ensureOverlayWindowLevel: (window) =>
            params.ensureOverlayWindowLevel(window as BrowserWindow),
        }),
      getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
      getMainWindow: () => overlayManager.getMainWindow(),
      ensureOverlayWindowLevel: (window) => ensureOverlayWindowLevel(window as BrowserWindow),
    });
  const enforceOverlayLayerOrderMainDeps = buildEnforceOverlayLayerOrderMainDepsHandler();
  const enforceOverlayLayerOrder = createEnforceOverlayLayerOrderHandler(
    enforceOverlayLayerOrderMainDeps,
  );

  return {
    getLastOverlayWindowGeometry: () => lastOverlayWindowGeometry,
    resetLastOverlayWindowGeometry: () => {
      lastOverlayWindowGeometry = null;
    },
    getOverlayGeometryFallback,
    getCurrentOverlayGeometry,
    getCurrentTrackedOverlayGeometry,
    geometryMatches,
    applyOverlayRegions,
    shouldExitLinuxFullscreenOverrideForGeometry,
    maybeExitLinuxFullscreenOverrideForTrackedGeometry,
    hasHyprlandOverlayWindowPlacementMismatch,
    moveVisibleOverlayAboveTrackedPlaybackWindow,
    bindVisibleOverlayToTrackedX11Window,
    syncPrimaryOverlayWindowLayer,
    updateVisibleOverlayBounds,
    ensureOverlayWindowLevel,
    enforceOverlayLayerOrder,
  };
}

export type OverlayGeometryRuntime = ReturnType<typeof createOverlayGeometryRuntime>;
