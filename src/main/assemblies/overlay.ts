import type { BrowserWindow, Screen, Session } from 'electron';
import {
  type createOverlayWindow as createOverlayWindowCoreFn,
  sendMpvCommandRuntime,
  setVisibleOverlayVisible as setVisibleOverlayVisibleCore,
  type createOverlayManager,
} from '../../core/services';
import type { Logger } from '../../logger';
import type { WindowGeometry } from '../../types';
import { ensureWindowsOverlayTransparency } from '../../window-trackers/windows-helper';
import type { createOverlayShortcutsRuntimeService } from '../overlay-shortcuts-runtime';
import { createOverlayVisibilityRuntimeService } from '../overlay-visibility-runtime';
import type { createAutoplayReadyGate } from '../runtime/autoplay-ready-gate';
import type { composeOverlayVisibilityRuntime } from '../runtime/composers';
import { createOverlayVisibilityRuntime } from '../runtime/domains/overlay';
import {
  type CancelLinuxMpvFullscreenOverlayRefreshBurst,
  clearLinuxMpvFullscreenOverlayRefreshTimeouts,
  updateLinuxMpvFullscreenOverlayRefreshBurst,
} from '../runtime/linux-mpv-fullscreen-overlay-refresh';
import { createLinuxOverlayModeRuntime } from '../runtime/linux-overlay-mode-runtime';
import { shouldRunLinuxOverlayZOrderKeepAlive } from '../runtime/linux-overlay-zorder-keepalive';
import type { createMediaTimingReviewRuntime } from '../runtime/media-timing-review';
import type { createOverlayModalInputState } from '../runtime/overlay-modal-input-state';
import type { createOverlayGeometryRuntime } from '../runtime/overlay-geometry-runtime';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import { createOverlayWindowRuntimeHandlers } from '../runtime/overlay-window-runtime-handlers';
import type { VisibleOverlayInteractionRuntime } from '../runtime/visible-overlay-interaction-runtime';
import type { AppState } from '../state';
import type { PlaybackAssembly } from './playback';

type OverlayManager = ReturnType<typeof createOverlayManager>;

/**
 * What overlay core needs from the rest of `main.ts`. Everything here except the getters and
 * callbacks is built before the assembly; siblings built later are reached through getters, and
 * construction calls none of them.
 */
export interface OverlayAssemblyDeps {
  appState: AppState;
  logger: Pick<Logger, 'debug' | 'info' | 'warn'>;
  overlayManager: Pick<
    OverlayManager,
    | 'getMainWindow'
    | 'setMainWindow'
    | 'getModalWindow'
    | 'setModalWindow'
    | 'getVisibleOverlayVisible'
    | 'setVisibleOverlayVisible'
    | 'setOverlayWindowBounds'
    | 'setModalWindowBounds'
  >;
  overlayModalInputState: Pick<
    ReturnType<typeof createOverlayModalInputState>,
    'getModalInputExclusive'
  >;
  autoplayReadyGate: Pick<
    ReturnType<typeof createAutoplayReadyGate>,
    'markCurrentMediaAutoplayReady' | 'flushPendingAutoplayReadySignal'
  >;
  screen: Pick<Screen, 'getCursorScreenPoint' | 'getDisplayNearestPoint'>;
  isDev: boolean;
  /** Builds an Electron overlay window; injected so the assembly stays free of Electron calls. */
  createOverlayWindowCore: typeof createOverlayWindowCoreFn;
  getYomitanSession: () => Session | null;
  /** Pointer interaction and geometry stay in `main.ts`: both import Electron at module load. */
  getVisibleOverlayInteraction: () => Pick<
    VisibleOverlayInteractionRuntime,
    | 'getLinuxOverlayInputShapeActive'
    | 'getMacOSVisibleOverlayForegroundProbeActive'
    | 'getVisibleOverlayInteractionActive'
    | 'getWindowsNativeWindowHandleNumber'
    | 'primeLinuxOverlayPointerInteractionAfterFirstMeasurement'
    | 'requestLinuxOverlayZOrderFollow'
    | 'requestWindowsVisibleOverlayZOrderSync'
    | 'resetLinuxVisibleOverlayStartupInputPrimer'
    | 'resetVisibleOverlayInputState'
    | 'restoreVisibleOverlayWindowShapeForShow'
    | 'scheduleVisibleOverlayBlurRefresh'
    | 'startLinuxVisibleOverlayStartupInputGrace'
  >;
  getOverlayGeometry: () => Pick<
    ReturnType<typeof createOverlayGeometryRuntime>,
    | 'enforceOverlayLayerOrder'
    | 'ensureOverlayWindowLevel'
    | 'getCurrentOverlayGeometry'
    | 'getCurrentTrackedOverlayGeometry'
    | 'syncPrimaryOverlayWindowLayer'
    | 'updateVisibleOverlayBounds'
  >;
  getOverlayShortcuts: () => Pick<
    ReturnType<typeof createOverlayShortcutsRuntimeService>,
    'syncOverlayShortcuts' | 'tryHandleOverlayShortcutLocalFallback'
  >;
  getOverlayNotifications: () => Pick<
    ReturnType<typeof createOverlayNotificationsRuntime>,
    | 'showOverlayLoadingStatusNotification'
    | 'dismissOverlayLoadingStatusNotification'
    | 'flushQueuedOverlayNotifications'
    | 'maybeStartOverlayLoadingOsd'
  >;
  getOverlayVisibilityComposer: () => Pick<
    ReturnType<typeof composeOverlayVisibilityRuntime>,
    'broadcastRuntimeOptionsChanged' | 'setOverlayDebugVisualizationEnabled'
  >;
  getMediaTimingReviewRuntime: () => Pick<
    ReturnType<typeof createMediaTimingReviewRuntime>,
    'dispose'
  >;
  getSubtitleProcessingController: () => Pick<
    PlaybackAssembly['subtitleProcessingController'],
    'refreshCurrentSubtitle'
  >;
  getAutoplaySubtitlePriming: () => Pick<
    PlaybackAssembly['autoplaySubtitlePrimingRuntime'],
    'cancelVisibleOverlaySubtitleRefreshAfterFirstPaint' | 'primeCurrentSubtitleForVisibleOverlay'
  >;
  ensureOverlayMpvSubtitlesHidden: () => Promise<void>;
  syncOverlayMpvSubtitleSuppression: () => void;
  initializeOverlayRuntime: () => void;
}

export type OverlayAssembly = ReturnType<typeof createOverlayAssembly>;

/**
 * Owns overlay core: the visible/modal windows, their visibility, the Linux fullscreen window
 * mode and refresh burst, and the show/hide paths. `main.ts` builds it once and reaches the
 * runtimes through the result; overlay bootstrap, pointer interaction and geometry stay there.
 */
export function createOverlayAssembly(deps: OverlayAssemblyDeps) {
  const { appState, logger, overlayManager, overlayModalInputState, autoplayReadyGate } = deps;

  let cancelLinuxMpvFullscreenOverlayRefreshBurst: CancelLinuxMpvFullscreenOverlayRefreshBurst | null =
    null;
  const linuxOverlayModeRuntime = createLinuxOverlayModeRuntime({
    isEnabled: shouldRunLinuxOverlayZOrderKeepAlive,
    isVisible: () => overlayManager.getVisibleOverlayVisible(),
    getWindow: () => overlayManager.getMainWindow(),
    clearWindow: () => overlayManager.setMainWindow(null),
    createWindow: () => {
      deps.getVisibleOverlayInteraction().resetVisibleOverlayInputState();
      createMainWindow();
    },
    refreshWindow: () => {
      const trackedGeometry = deps.getOverlayGeometry().getCurrentTrackedOverlayGeometry();
      if (trackedGeometry) overlayManager.setOverlayWindowBounds(trackedGeometry);
      overlayVisibilityRuntime.updateVisibleOverlayVisibility();
      void deps.ensureOverlayMpvSubtitlesHidden();
      if (appState.currentSubText.trim()) {
        deps.getSubtitleProcessingController().refreshCurrentSubtitle(appState.currentSubText);
      }
    },
    now: Date.now,
    logDebug: (message) => logger.debug(message),
  });
  function cancelPendingLinuxMpvFullscreenOverlayRefreshBurst(): void {
    cancelLinuxMpvFullscreenOverlayRefreshBurst?.();
    cancelLinuxMpvFullscreenOverlayRefreshBurst = null;
  }

  const overlayVisibilityRuntime = createOverlayVisibilityRuntimeService({
    getMainWindow: () => overlayManager.getMainWindow(),
    getModalActive: () => overlayModalInputState.getModalInputExclusive(),
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    getForceMousePassthrough: () => appState.statsOverlayVisible,
    getNonNativeInputRegionActive: () =>
      process.platform === 'linux' &&
      deps.getVisibleOverlayInteraction().getLinuxOverlayInputShapeActive(),
    getSuspendVisibleOverlay: () => appState.statsOverlayVisible,
    getOverlayInteractionActive: () =>
      deps.getVisibleOverlayInteraction().getVisibleOverlayInteractionActive(),
    getWindowTracker: () => appState.windowTracker,
    getMacOSForegroundProbeActive: () =>
      deps.getVisibleOverlayInteraction().getMacOSVisibleOverlayForegroundProbeActive(),
    getTrackerNotReadyWarningShown: () => appState.trackerNotReadyWarningShown,
    setTrackerNotReadyWarningShown: (shown: boolean) => {
      appState.trackerNotReadyWarningShown = shown;
    },
    updateVisibleOverlayBounds: (geometry: WindowGeometry) =>
      deps.getOverlayGeometry().updateVisibleOverlayBounds(geometry),
    ensureOverlayWindowLevel: (window) => {
      deps.getOverlayGeometry().ensureOverlayWindowLevel(window);
    },
    syncWindowsOverlayToMpvZOrder: (_window) => {
      deps.getVisibleOverlayInteraction().requestWindowsVisibleOverlayZOrderSync();
    },
    syncPrimaryOverlayWindowLayer: (layer) => {
      deps.getOverlayGeometry().syncPrimaryOverlayWindowLayer(layer);
    },
    enforceOverlayLayerOrder: () => {
      deps.getOverlayGeometry().enforceOverlayLayerOrder();
    },
    syncOverlayShortcuts: () => {
      deps.getOverlayShortcuts().syncOverlayShortcuts();
    },
    isMacOSPlatform: () => process.platform === 'darwin',
    isWindowsPlatform: () => process.platform === 'win32',
    showOverlayLoadingOsd: (_message: string) => {
      deps.getOverlayNotifications().showOverlayLoadingStatusNotification();
    },
    dismissOverlayLoadingOsd: () => {
      deps.getOverlayNotifications().dismissOverlayLoadingStatusNotification();
    },
    hideNonNativeOverlayWhenTargetUnfocused: () =>
      shouldRunLinuxOverlayZOrderKeepAlive() &&
      linuxOverlayModeRuntime.mode === 'fullscreen-override',
    resolveFallbackBounds: () => {
      const cursorPoint = deps.screen.getCursorScreenPoint();
      const display = deps.screen.getDisplayNearestPoint(cursorPoint);
      const fallbackBounds = display.workArea;
      return {
        x: fallbackBounds.x,
        y: fallbackBounds.y,
        width: fallbackBounds.width,
        height: fallbackBounds.height,
      };
    },
  });

  function createModalWindow(): BrowserWindow {
    const existingWindow = overlayManager.getModalWindow();
    if (existingWindow && !existingWindow.isDestroyed()) {
      return existingWindow;
    }
    const window = createModalWindowHandler();
    overlayManager.setModalWindowBounds(deps.getOverlayGeometry().getCurrentOverlayGeometry());
    return window;
  }

  function createMainWindow(): BrowserWindow {
    const window = createMainWindowHandler();
    if (process.platform === 'win32') {
      const overlayHwnd = deps
        .getVisibleOverlayInteraction()
        .getWindowsNativeWindowHandleNumber(window);
      if (!ensureWindowsOverlayTransparency(overlayHwnd)) {
        logger.warn('Failed to eagerly extend Windows overlay transparency via koffi');
      }
    }
    return window;
  }

  const {
    setVisibleOverlayVisible: setVisibleOverlayVisibleHandler,
    toggleVisibleOverlay: toggleVisibleOverlayHandler,
    setOverlayVisible: setOverlayVisibleHandler,
  } = createOverlayVisibilityRuntime({
    setVisibleOverlayVisibleDeps: {
      setVisibleOverlayVisibleCore,
      setVisibleOverlayVisibleState: (nextVisible) => {
        overlayManager.setVisibleOverlayVisible(nextVisible);
      },
      updateVisibleOverlayVisibility: () =>
        overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    },
    getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
  });

  const { createMainWindow: createMainWindowHandler, createModalWindow: createModalWindowHandler } =
    createOverlayWindowRuntimeHandlers<BrowserWindow>({
      createOverlayWindowDeps: {
        createOverlayWindowCore: (kind, options) => deps.createOverlayWindowCore(kind, options),
        isDev: deps.isDev,
        ensureOverlayWindowLevel: (window) =>
          deps.getOverlayGeometry().ensureOverlayWindowLevel(window),
        onRuntimeOptionsChanged: () =>
          deps.getOverlayVisibilityComposer().broadcastRuntimeOptionsChanged(),
        setOverlayDebugVisualizationEnabled: (enabled) =>
          deps.getOverlayVisibilityComposer().setOverlayDebugVisualizationEnabled(enabled),
        isOverlayVisible: (windowKind) =>
          windowKind === 'visible' ? overlayManager.getVisibleOverlayVisible() : false,
        getYomitanSession: () => deps.getYomitanSession(),
        tryHandleOverlayShortcutLocalFallback: (input) =>
          deps.getOverlayShortcuts().tryHandleOverlayShortcutLocalFallback(input),
        forwardTabToMpv: () => sendMpvCommandRuntime(appState.mpvClient, ['keypress', 'TAB']),
        getLinuxX11FullscreenOverlay: () =>
          shouldRunLinuxOverlayZOrderKeepAlive() &&
          linuxOverlayModeRuntime.fullscreen &&
          linuxOverlayModeRuntime.mode === 'fullscreen-override',
        onVisibleWindowBlurred: () =>
          deps.getVisibleOverlayInteraction().scheduleVisibleOverlayBlurRefresh(),
        onVisibleWindowFocused: () =>
          deps.getVisibleOverlayInteraction().requestLinuxOverlayZOrderFollow(),
        onWindowDidFinishLoad: () => {
          deps.getOverlayNotifications().flushQueuedOverlayNotifications();
        },
        onWindowContentReady: () => {
          deps.getOverlayNotifications().dismissOverlayLoadingStatusNotification();
          deps.getOverlayNotifications().flushQueuedOverlayNotifications();
          overlayVisibilityRuntime.updateVisibleOverlayVisibility();
          deps
            .getVisibleOverlayInteraction()
            .primeLinuxOverlayPointerInteractionAfterFirstMeasurement();
          autoplayReadyGate.flushPendingAutoplayReadySignal();
        },
        onWindowClosed: (windowKind, window) => {
          if (windowKind === 'visible') {
            if (overlayManager.getMainWindow() !== window) {
              return;
            }
            cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
            overlayManager.setMainWindow(null);
          } else {
            if (overlayManager.getModalWindow() !== window) {
              return;
            }
            void deps.getMediaTimingReviewRuntime().dispose();
            overlayManager.setModalWindow(null);
          }
        },
      },
      getMainWindow: () => overlayManager.getMainWindow(),
      isWindowDestroyed: (window) => window.isDestroyed(),
      setMainWindow: (window) => overlayManager.setMainWindow(window),
      setModalWindow: (window) => overlayManager.setModalWindow(window),
    });

  function ensureOverlayWindowsReadyForVisibilityActions(): void {
    if (!appState.overlayRuntimeInitialized) {
      deps.initializeOverlayRuntime();
      return;
    }

    const mainWindow = overlayManager.getMainWindow();
    if (!mainWindow || mainWindow.isDestroyed()) {
      createMainWindow();
    }
  }

  function notifyMpvPluginVisibleOverlayVisibility(visible: boolean): void {
    sendMpvCommandRuntime(appState.mpvClient, [
      'script-message',
      visible ? 'subminer-visible-overlay-shown' : 'subminer-visible-overlay-hidden',
    ]);
  }

  function setVisibleOverlayVisible(visible: boolean): void {
    ensureOverlayWindowsReadyForVisibilityActions();
    if (!visible) {
      deps.getOverlayNotifications().dismissOverlayLoadingStatusNotification();
      autoplayReadyGate.markCurrentMediaAutoplayReady();
      deps.getAutoplaySubtitlePriming().cancelVisibleOverlaySubtitleRefreshAfterFirstPaint();
      cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
      deps.getVisibleOverlayInteraction().resetVisibleOverlayInputState();
    }
    if (visible) {
      deps.getOverlayNotifications().maybeStartOverlayLoadingOsd();
      deps.getVisibleOverlayInteraction().resetLinuxVisibleOverlayStartupInputPrimer();
      deps.getVisibleOverlayInteraction().startLinuxVisibleOverlayStartupInputGrace();
      deps.getVisibleOverlayInteraction().restoreVisibleOverlayWindowShapeForShow();
      void deps.ensureOverlayMpvSubtitlesHidden();
      void deps.getAutoplaySubtitlePriming().primeCurrentSubtitleForVisibleOverlay();
    }
    setVisibleOverlayVisibleHandler(visible);
    notifyMpvPluginVisibleOverlayVisibility(visible);
    deps.syncOverlayMpvSubtitleSuppression();
  }

  function toggleVisibleOverlay(): void {
    ensureOverlayWindowsReadyForVisibilityActions();
    const nextVisible = !overlayManager.getVisibleOverlayVisible();
    if (!nextVisible) {
      deps.getOverlayNotifications().dismissOverlayLoadingStatusNotification();
      autoplayReadyGate.markCurrentMediaAutoplayReady();
      deps.getAutoplaySubtitlePriming().cancelVisibleOverlaySubtitleRefreshAfterFirstPaint();
      cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
      deps.getVisibleOverlayInteraction().resetVisibleOverlayInputState();
    } else {
      deps.getOverlayNotifications().maybeStartOverlayLoadingOsd();
      deps.getVisibleOverlayInteraction().resetLinuxVisibleOverlayStartupInputPrimer();
      deps.getVisibleOverlayInteraction().startLinuxVisibleOverlayStartupInputGrace();
      deps.getVisibleOverlayInteraction().restoreVisibleOverlayWindowShapeForShow();
      void deps.ensureOverlayMpvSubtitlesHidden();
      void deps.getAutoplaySubtitlePriming().primeCurrentSubtitleForVisibleOverlay();
    }
    toggleVisibleOverlayHandler();
    notifyMpvPluginVisibleOverlayVisibility(nextVisible);
    deps.syncOverlayMpvSubtitleSuppression();
  }
  function setOverlayVisible(visible: boolean): void {
    if (!visible) {
      deps.getOverlayNotifications().dismissOverlayLoadingStatusNotification();
      deps.getAutoplaySubtitlePriming().cancelVisibleOverlaySubtitleRefreshAfterFirstPaint();
      deps.getVisibleOverlayInteraction().resetVisibleOverlayInputState();
      autoplayReadyGate.markCurrentMediaAutoplayReady();
      cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
    }
    if (visible) {
      deps.getOverlayNotifications().maybeStartOverlayLoadingOsd();
      deps.getVisibleOverlayInteraction().resetLinuxVisibleOverlayStartupInputPrimer();
      deps.getVisibleOverlayInteraction().startLinuxVisibleOverlayStartupInputGrace();
      deps.getVisibleOverlayInteraction().restoreVisibleOverlayWindowShapeForShow();
      void deps.ensureOverlayMpvSubtitlesHidden();
      void deps.getAutoplaySubtitlePriming().primeCurrentSubtitleForVisibleOverlay();
    }
    setOverlayVisibleHandler(visible);
    notifyMpvPluginVisibleOverlayVisibility(visible);
    deps.syncOverlayMpvSubtitleSuppression();
  }

  function handleMpvFullscreenChange(fullscreen: boolean): void {
    cancelLinuxMpvFullscreenOverlayRefreshBurst = updateLinuxMpvFullscreenOverlayRefreshBurst(
      fullscreen,
      {
        overlayManager: {
          getMainWindow: () => overlayManager.getMainWindow(),
          getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
        },
        overlayVisibilityRuntime,
        syncVisibleOverlayMpvFullscreenMode: (nextFullscreen) =>
          linuxOverlayModeRuntime.sync(nextFullscreen),
        getOverlayInteractionActive: () =>
          deps.getVisibleOverlayInteraction().getVisibleOverlayInteractionActive() ||
          deps.getVisibleOverlayInteraction().getLinuxOverlayInputShapeActive(),
        ensureOverlayWindowLevel: (window) =>
          deps.getOverlayGeometry().ensureOverlayWindowLevel(window),
      },
      cancelLinuxMpvFullscreenOverlayRefreshBurst,
    );
  }

  /** Quit cleanup: drops any pending Linux fullscreen transition and refresh burst. */
  function clearLinuxFullscreenOverlayRefresh(): void {
    linuxOverlayModeRuntime.cancelPendingTransition();
    cancelLinuxMpvFullscreenOverlayRefreshBurst = null;
    clearLinuxMpvFullscreenOverlayRefreshTimeouts();
  }

  return {
    linuxOverlayModeRuntime,
    overlayVisibilityRuntime,
    createMainWindow,
    createModalWindow,
    ensureOverlayWindowsReadyForVisibilityActions,
    setVisibleOverlayVisible,
    toggleVisibleOverlay,
    setOverlayVisible,
    handleMpvFullscreenChange,
    clearLinuxFullscreenOverlayRefresh,
  };
}
