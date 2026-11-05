import type { BrowserWindow, Screen, Session } from 'electron';
import {
  broadcastRuntimeOptionsChangedRuntime,
  type createOverlayContentMeasurementStore,
  type createOverlayWindow as createOverlayWindowCoreFn,
  sendMpvCommandRuntime,
  setMpvSubVisibilityRuntime,
  setOverlayDebugVisualizationEnabledRuntime,
  setVisibleOverlayVisible as setVisibleOverlayVisibleCore,
  type createOverlayManager,
} from '../../core/services';
import { applyOverlayClickThrough } from '../../core/services/overlay-click-through';
import type { Logger } from '../../logger';
import type { WindowGeometry } from '../../types';
import { ensureWindowsOverlayTransparency } from '../../window-trackers/windows-helper';
import type { createOverlayShortcutsRuntimeService } from '../overlay-shortcuts-runtime';
import type { OverlayModalRuntime } from '../overlay-runtime';
import { createOverlayVisibilityRuntimeService } from '../overlay-visibility-runtime';
import type { createAutoplayReadyGate } from '../runtime/autoplay-ready-gate';
import { composeOverlayVisibilityRuntime } from '../runtime/composers';
import { createOverlayVisibilityRuntime } from '../runtime/domains/overlay';
import {
  type CancelLinuxMpvFullscreenOverlayRefreshBurst,
  clearLinuxMpvFullscreenOverlayRefreshTimeouts,
  updateLinuxMpvFullscreenOverlayRefreshBurst,
} from '../runtime/linux-mpv-fullscreen-overlay-refresh';
import { createLinuxOverlayModeRuntime } from '../runtime/linux-overlay-mode-runtime';
import { shouldRunLinuxOverlayZOrderKeepAlive } from '../runtime/linux-overlay-zorder-keepalive';
import { focusMacOSOverlayWindow } from '../runtime/macos-overlay-window-focus';
import type { createMediaTimingReviewRuntime } from '../runtime/media-timing-review';
import type { createJimakuAutoOpenRuntime } from '../runtime/jimaku-auto-open';
import type { createOverlayModalInputState } from '../runtime/overlay-modal-input-state';
import {
  createEnsureOverlayMpvSubtitlesHiddenHandler,
  createRestoreOverlayMpvSubtitlesHandler,
} from '../runtime/overlay-mpv-sub-visibility';
import { createOverlayGeometryRuntime } from '../runtime/overlay-geometry-runtime';
import type { createOverlayNotificationsRuntime } from '../runtime/overlay-notifications-runtime';
import { createOverlayWindowRuntimeHandlers } from '../runtime/overlay-window-runtime-handlers';
import { createVisibleOverlayInteractionRuntime } from '../runtime/visible-overlay-interaction-runtime';
import type { AppState } from '../state';
import type { PlaybackAssembly } from './playback';
import type { IpcSlice } from './slices';

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
    | 'getOverlayWindows'
    | 'broadcastToOverlayWindows'
  >;
  overlayModalInputState: Pick<
    ReturnType<typeof createOverlayModalInputState>,
    'getModalInputExclusive'
  >;
  overlayModalRuntime: Pick<
    OverlayModalRuntime,
    | 'notifyOverlayModalOpened'
    | 'handleOverlayModalClosed'
    | 'sendToActiveOverlayWindow'
    | 'openRuntimeOptionsPalette'
  >;
  overlayContentMeasurementStore: Pick<
    ReturnType<typeof createOverlayContentMeasurementStore>,
    'report' | 'clear' | 'getLatestByLayer'
  >;
  getAutoplayReadyGate: () => Pick<
    ReturnType<typeof createAutoplayReadyGate>,
    'markCurrentMediaAutoplayReady' | 'flushPendingAutoplayReadySignal'
  >;
  screen: Pick<
    Screen,
    'getCursorScreenPoint' | 'getDisplayNearestPoint' | 'getDisplayMatching' | 'screenToDipPoint'
  >;
  isDev: boolean;
  /** Builds an Electron overlay window; injected so the assembly stays free of Electron calls. */
  createOverlayWindowCore: typeof createOverlayWindowCoreFn;
  getYomitanSession: () => Session | null;
  getOverlayShortcuts: () => Pick<
    ReturnType<typeof createOverlayShortcutsRuntimeService>,
    'syncOverlayShortcuts' | 'tryHandleOverlayShortcutLocalFallback'
  >;
  /** Settings/setup windows the overlay must not raise above. */
  getOverlayForegroundSeparateWindows: () => BrowserWindow[];
  /** Keeps the stats window above playback after the overlay re-asserts its level. */
  promoteStatsOverlayAbovePlayback: () => void;
  getOverlayNotifications: () => Pick<
    ReturnType<typeof createOverlayNotificationsRuntime>,
    | 'showOverlayLoadingStatusNotification'
    | 'dismissOverlayLoadingStatusNotification'
    | 'flushQueuedOverlayNotifications'
    | 'maybeStartOverlayLoadingOsd'
  >;
  /** Runtime option snapshot broadcast to overlay windows when an option changes. */
  getRuntimeOptionsState: () => ReturnType<
    Parameters<
      typeof composeOverlayVisibilityRuntime
    >[0]['broadcastRuntimeOptionsChangedMainDeps']['getRuntimeOptionsState']
  >;
  getMediaTimingReviewRuntime: () => Pick<
    ReturnType<typeof createMediaTimingReviewRuntime>,
    'dispose'
  >;
  getJimakuAutoOpenRuntime: () => Pick<
    ReturnType<typeof createJimakuAutoOpenRuntime>,
    'handleJimakuModalOpened' | 'handleJimakuModalClosed'
  >;
  getSubtitleProcessingController: () => Pick<
    PlaybackAssembly['subtitleProcessingController'],
    'refreshCurrentSubtitle'
  >;
  getAutoplaySubtitlePriming: () => Pick<
    PlaybackAssembly['autoplaySubtitlePrimingRuntime'],
    | 'cancelVisibleOverlaySubtitleRefreshAfterFirstPaint'
    | 'primeCurrentSubtitleForVisibleOverlay'
    | 'scheduleVisibleOverlaySubtitleRefreshAfterFirstPaint'
  >;
  /** A late close for a replaced plugin panel reports false and must be ignored. */
  getPluginPanel: () => { handleModalClosed: (panelId: number) => boolean };
  /** The renderer reports the subtitle sidebar as a modal; playback remembers its state. */
  setSubtitleSidebarRequestedOpen: (open: boolean) => void;
  /** Brings SubMiner to the foreground (`app.focus({ steal: true })`). */
  stealAppFocus: () => void;
  initializeOverlayRuntime: () => void;
}

export type OverlayAssembly = ReturnType<typeof createOverlayAssembly>;

/**
 * Owns overlay core: the visible/modal windows, their visibility, geometry and pointer
 * interaction, the Linux fullscreen window mode and refresh burst, mpv subtitle suppression
 * while the overlay shows, and the show/hide paths. `main.ts` builds it once and reaches the
 * runtimes through the result; overlay bootstrap stays in the startup assembly.
 */
export function createOverlayAssembly(deps: OverlayAssemblyDeps) {
  const { appState, logger, overlayManager, overlayModalInputState } = deps;

  let cancelLinuxMpvFullscreenOverlayRefreshBurst: CancelLinuxMpvFullscreenOverlayRefreshBurst | null =
    null;
  const linuxOverlayModeRuntime = createLinuxOverlayModeRuntime({
    isEnabled: shouldRunLinuxOverlayZOrderKeepAlive,
    isVisible: () => overlayManager.getVisibleOverlayVisible(),
    getWindow: () => overlayManager.getMainWindow(),
    clearWindow: () => overlayManager.setMainWindow(null),
    createWindow: () => {
      interaction.resetVisibleOverlayInputState();
      createMainWindow();
    },
    refreshWindow: () => {
      const trackedGeometry = geometry.getCurrentTrackedOverlayGeometry();
      if (trackedGeometry) overlayManager.setOverlayWindowBounds(trackedGeometry);
      overlayVisibilityRuntime.updateVisibleOverlayVisibility();
      void ensureOverlayMpvSubtitlesHidden();
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
      process.platform === 'linux' && interaction.getLinuxOverlayInputShapeActive(),
    getSuspendVisibleOverlay: () => appState.statsOverlayVisible,
    getOverlayInteractionActive: () => interaction.getVisibleOverlayInteractionActive(),
    getWindowTracker: () => appState.windowTracker,
    getMacOSForegroundProbeActive: () => interaction.getMacOSVisibleOverlayForegroundProbeActive(),
    getTrackerNotReadyWarningShown: () => appState.trackerNotReadyWarningShown,
    setTrackerNotReadyWarningShown: (shown: boolean) => {
      appState.trackerNotReadyWarningShown = shown;
    },
    updateVisibleOverlayBounds: (bounds: WindowGeometry) =>
      geometry.updateVisibleOverlayBounds(bounds),
    ensureOverlayWindowLevel: (window) => {
      geometry.ensureOverlayWindowLevel(window);
    },
    syncWindowsOverlayToMpvZOrder: (_window) => {
      interaction.requestWindowsVisibleOverlayZOrderSync();
    },
    syncPrimaryOverlayWindowLayer: (layer) => {
      geometry.syncPrimaryOverlayWindowLayer(layer);
    },
    enforceOverlayLayerOrder: () => {
      geometry.enforceOverlayLayerOrder();
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
    overlayManager.setModalWindowBounds(geometry.getCurrentOverlayGeometry());
    return window;
  }

  function createMainWindow(): BrowserWindow {
    const window = createMainWindowHandler();
    if (process.platform === 'win32') {
      const overlayHwnd = interaction.getWindowsNativeWindowHandleNumber(window);
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
        ensureOverlayWindowLevel: (window) => geometry.ensureOverlayWindowLevel(window),
        onRuntimeOptionsChanged: () => visibilityComposer.broadcastRuntimeOptionsChanged(),
        setOverlayDebugVisualizationEnabled: (enabled) =>
          visibilityComposer.setOverlayDebugVisualizationEnabled(enabled),
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
        onVisibleWindowBlurred: () => interaction.scheduleVisibleOverlayBlurRefresh(),
        onVisibleWindowFocused: () => interaction.requestLinuxOverlayZOrderFollow(),
        onWindowDidFinishLoad: () => {
          deps.getOverlayNotifications().flushQueuedOverlayNotifications();
        },
        onWindowContentReady: () => {
          deps.getOverlayNotifications().dismissOverlayLoadingStatusNotification();
          deps.getOverlayNotifications().flushQueuedOverlayNotifications();
          overlayVisibilityRuntime.updateVisibleOverlayVisibility();
          interaction.primeLinuxOverlayPointerInteractionAfterFirstMeasurement();
          deps.getAutoplayReadyGate().flushPendingAutoplayReadySignal();
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
            // A destroyed modal window never reports its modals closed.
            deps.getJimakuAutoOpenRuntime().handleJimakuModalClosed();
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
      deps.getAutoplayReadyGate().markCurrentMediaAutoplayReady();
      deps.getAutoplaySubtitlePriming().cancelVisibleOverlaySubtitleRefreshAfterFirstPaint();
      cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
      interaction.resetVisibleOverlayInputState();
    }
    if (visible) {
      deps.getOverlayNotifications().maybeStartOverlayLoadingOsd();
      interaction.resetLinuxVisibleOverlayStartupInputPrimer();
      interaction.startLinuxVisibleOverlayStartupInputGrace();
      interaction.restoreVisibleOverlayWindowShapeForShow();
      void ensureOverlayMpvSubtitlesHidden();
      void deps.getAutoplaySubtitlePriming().primeCurrentSubtitleForVisibleOverlay();
    }
    setVisibleOverlayVisibleHandler(visible);
    notifyMpvPluginVisibleOverlayVisibility(visible);
    syncOverlayMpvSubtitleSuppression();
  }

  function toggleVisibleOverlay(): void {
    ensureOverlayWindowsReadyForVisibilityActions();
    const nextVisible = !overlayManager.getVisibleOverlayVisible();
    if (!nextVisible) {
      deps.getOverlayNotifications().dismissOverlayLoadingStatusNotification();
      deps.getAutoplayReadyGate().markCurrentMediaAutoplayReady();
      deps.getAutoplaySubtitlePriming().cancelVisibleOverlaySubtitleRefreshAfterFirstPaint();
      cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
      interaction.resetVisibleOverlayInputState();
    } else {
      deps.getOverlayNotifications().maybeStartOverlayLoadingOsd();
      interaction.resetLinuxVisibleOverlayStartupInputPrimer();
      interaction.startLinuxVisibleOverlayStartupInputGrace();
      interaction.restoreVisibleOverlayWindowShapeForShow();
      void ensureOverlayMpvSubtitlesHidden();
      void deps.getAutoplaySubtitlePriming().primeCurrentSubtitleForVisibleOverlay();
    }
    toggleVisibleOverlayHandler();
    notifyMpvPluginVisibleOverlayVisibility(nextVisible);
    syncOverlayMpvSubtitleSuppression();
  }
  function setOverlayVisible(visible: boolean): void {
    if (!visible) {
      deps.getOverlayNotifications().dismissOverlayLoadingStatusNotification();
      deps.getAutoplaySubtitlePriming().cancelVisibleOverlaySubtitleRefreshAfterFirstPaint();
      interaction.resetVisibleOverlayInputState();
      deps.getAutoplayReadyGate().markCurrentMediaAutoplayReady();
      cancelPendingLinuxMpvFullscreenOverlayRefreshBurst();
    }
    if (visible) {
      deps.getOverlayNotifications().maybeStartOverlayLoadingOsd();
      interaction.resetLinuxVisibleOverlayStartupInputPrimer();
      interaction.startLinuxVisibleOverlayStartupInputGrace();
      interaction.restoreVisibleOverlayWindowShapeForShow();
      void ensureOverlayMpvSubtitlesHidden();
      void deps.getAutoplaySubtitlePriming().primeCurrentSubtitleForVisibleOverlay();
    }
    setOverlayVisibleHandler(visible);
    notifyMpvPluginVisibleOverlayVisibility(visible);
    syncOverlayMpvSubtitleSuppression();
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
          interaction.getVisibleOverlayInteractionActive() ||
          interaction.getLinuxOverlayInputShapeActive(),
        ensureOverlayWindowLevel: (window) => geometry.ensureOverlayWindowLevel(window),
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

  const geometry = createOverlayGeometryRuntime({
    screen: deps.screen,
    promoteStatsOverlayAbovePlayback: () => deps.promoteStatsOverlayAbovePlayback(),
    overlayManager: {
      getMainWindow: () => overlayManager.getMainWindow(),
      getModalWindow: () => overlayManager.getModalWindow(),
      getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
      setOverlayWindowBounds: (bounds) => overlayManager.setOverlayWindowBounds(bounds),
      setModalWindowBounds: (bounds) => overlayManager.setModalWindowBounds(bounds),
    },
    getTrackedWindowGeometry: () => appState.windowTracker?.getGeometry() ?? null,
    getTrackedWindowMediaSourceId: () => appState.windowTracker?.getTargetWindowMediaSourceId?.(),
    getTrackedWindowNativeId: () => appState.windowTracker?.getTargetWindowNativeId?.(),
    getStatsOverlayVisible: () => appState.statsOverlayVisible,
    getOverlayForegroundSeparateWindows: () => deps.getOverlayForegroundSeparateWindows(),
    getLinuxVisibleOverlayWindowMode: () => linuxOverlayModeRuntime.mode,
    getLinuxTrackedMpvFullscreen: () => linuxOverlayModeRuntime.fullscreen,
    getLinuxTrackedMpvFullscreenChangedAtMs: () => linuxOverlayModeRuntime.fullscreenChangedAtMs,
    syncLinuxVisibleOverlayMpvFullscreenMode: (fullscreen) =>
      linuxOverlayModeRuntime.sync(fullscreen),
    getLinuxVisibleOverlayOwnerBindingKey: () => linuxOverlayModeRuntime.ownerBindingKey,
    setLinuxVisibleOverlayOwnerBindingKey: (key) => {
      linuxOverlayModeRuntime.ownerBindingKey = key;
    },
    clearVisibleOverlayX11OwnerBinding: (window) =>
      interaction.clearVisibleOverlayX11OwnerBinding(window),
    getNativeWindowHandleDecimal: (window) => interaction.getNativeWindowHandleDecimal(window),
    enqueueVisibleOverlayX11OwnerBindingOperation: (window, args, onError) =>
      interaction.enqueueVisibleOverlayX11OwnerBindingOperation(window, args, onError),
    scheduleWindowsVisibleOverlayZOrderSyncBurst: () =>
      interaction.scheduleWindowsVisibleOverlayZOrderSyncBurst(),
    logDebug: (message, ...args) => logger.debug(message, ...args),
  });

  // Starts the Windows foreground poll and Linux z-order keep-alive loops on those platforms.
  const interaction = createVisibleOverlayInteractionRuntime({
    screen: deps.screen,
    overlayManager: {
      getMainWindow: () => overlayManager.getMainWindow(),
      getVisibleOverlayVisible: () => overlayManager.getVisibleOverlayVisible(),
    },
    overlayContentMeasurementStore: {
      clear: (layer) => deps.overlayContentMeasurementStore.clear(layer),
      getLatestByLayer: (layer) => deps.overlayContentMeasurementStore.getLatestByLayer(layer),
    },
    logger: {
      info: (message, ...args) => logger.info(message, ...args),
      warn: (message, ...args) => logger.warn(message, ...args),
      debug: (message, ...args) => logger.debug(message, ...args),
    },
    updateVisibleOverlayVisibility: () => overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    getModalInputExclusive: () => overlayModalInputState.getModalInputExclusive(),
    getStatsOverlayVisible: () => appState.statsOverlayVisible,
    setStatsOverlayVisible: (visible) => {
      appState.statsOverlayVisible = visible;
    },
    getWindowTracker: () => appState.windowTracker,
    setWindowTracker: (tracker) => {
      appState.windowTracker = tracker;
    },
    setTrackerNotReadyWarningShown: (shown) => {
      appState.trackerNotReadyWarningShown = shown;
    },
    getMpvSocketPath: () => appState.mpvSocketPath,
    getBackendOverride: () => appState.backendOverride,
    getInitialArgs: () => appState.initialArgs,
    getOverlayRuntimeInitialized: () => appState.overlayRuntimeInitialized,
    getLinuxVisibleOverlayWindowMode: () => linuxOverlayModeRuntime.mode,
    setLinuxVisibleOverlayOwnerBindingKey: (key) => {
      linuxOverlayModeRuntime.ownerBindingKey = key;
    },
    bindVisibleOverlayToTrackedX11Window: (window) =>
      geometry.bindVisibleOverlayToTrackedX11Window(window),
    updateVisibleOverlayBounds: (bounds) => geometry.updateVisibleOverlayBounds(bounds),
    refreshCurrentSubtitle: () => {
      deps.getSubtitleProcessingController().refreshCurrentSubtitle(appState.currentSubText);
    },
    getOverlayWindows: () => overlayManager.getOverlayWindows(),
    syncOverlayShortcuts: () => deps.getOverlayShortcuts().syncOverlayShortcuts(),
    resetLastOverlayWindowGeometry: () => geometry.resetLastOverlayWindowGeometry(),
    enforceOverlayLayerOrder: () => {
      geometry.enforceOverlayLayerOrder();
    },
    getOverlayForegroundSeparateWindows: () => deps.getOverlayForegroundSeparateWindows(),
  });

  const visibilityComposer = composeOverlayVisibilityRuntime({
    overlayVisibilityRuntime,
    restorePreviousSecondarySubVisibilityMainDeps: {
      getMpvClient: () => appState.mpvClient,
    },
    broadcastRuntimeOptionsChangedMainDeps: {
      broadcastRuntimeOptionsChangedRuntime,
      getRuntimeOptionsState: () => deps.getRuntimeOptionsState(),
      broadcastToOverlayWindows: (channel, ...args) =>
        overlayManager.broadcastToOverlayWindows(channel, ...args),
    },
    sendToActiveOverlayWindowMainDeps: {
      sendToActiveOverlayWindowRuntime: (channel, payload, runtimeOptions) =>
        deps.overlayModalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
    },
    setOverlayDebugVisualizationEnabledMainDeps: {
      setOverlayDebugVisualizationEnabledRuntime,
      getCurrentEnabled: () => appState.overlayDebugVisualizationEnabled,
      setCurrentEnabled: (next) => {
        appState.overlayDebugVisualizationEnabled = next;
      },
    },
    openRuntimeOptionsPaletteMainDeps: {
      openRuntimeOptionsPaletteRuntime: () => deps.overlayModalRuntime.openRuntimeOptionsPalette(),
    },
  });

  // mpv's own subtitles stay hidden while the visible overlay shows and come back when it hides.
  const ensureOverlayMpvSubtitlesHidden = createEnsureOverlayMpvSubtitlesHiddenHandler({
    getMpvClient: () => appState.mpvClient,
    getSavedSubVisibility: () => appState.overlaySavedMpvSubVisibility,
    setSavedSubVisibility: (visible) => {
      appState.overlaySavedMpvSubVisibility = visible;
    },
    getRevision: () => appState.overlayMpvSubVisibilityRevision,
    setRevision: (revision) => {
      appState.overlayMpvSubVisibilityRevision = revision;
    },
    setMpvSubVisibility: (visible) => {
      setMpvSubVisibilityRuntime(appState.mpvClient, visible);
    },
    logWarn: (message, error) => {
      logger.warn(message, error);
    },
  });
  const restoreOverlayMpvSubtitles = createRestoreOverlayMpvSubtitlesHandler({
    getSavedSubVisibility: () => appState.overlaySavedMpvSubVisibility,
    setSavedSubVisibility: (visible) => {
      appState.overlaySavedMpvSubVisibility = visible;
    },
    getRevision: () => appState.overlayMpvSubVisibilityRevision,
    setRevision: (revision) => {
      appState.overlayMpvSubVisibilityRevision = revision;
    },
    isMpvConnected: () => Boolean(appState.mpvClient?.connected),
    shouldKeepSuppressedFromVisibleOverlayBinding: () => overlayManager.getVisibleOverlayVisible(),
    setMpvSubVisibility: (visible) => {
      setMpvSubVisibilityRuntime(appState.mpvClient, visible);
    },
  });

  function syncOverlayMpvSubtitleSuppression(): void {
    if (overlayManager.getVisibleOverlayVisible()) {
      void ensureOverlayMpvSubtitlesHidden();
      return;
    }

    restoreOverlayMpvSubtitles();
  }

  /** Renderer reports about overlay windows: focus, modal lifecycle, pointer and measurement. */
  const ipc: IpcSlice<
    | 'getMainWindow'
    | 'focusMainWindow'
    | 'activatePlaybackWindowForOverlayInteraction'
    | 'onOverlayModalClosed'
    | 'onOverlayModalOpened'
    | 'onOverlayMouseInteractionChanged'
    | 'onOverlayInteractiveHint'
    | 'reportOverlayContentBounds'
  > = {
    getMainWindow: () => overlayManager.getMainWindow(),
    focusMainWindow: () => {
      const mainWindow = overlayManager.getMainWindow();
      if (!mainWindow || mainWindow.isDestroyed()) return;
      if (process.platform === 'darwin') {
        focusMacOSOverlayWindow({
          platform: process.platform,
          getOverlayWindow: () => mainWindow,
          stealAppFocus: () => deps.stealAppFocus(),
          warn: (message, details) => logger.warn(message, details),
        });
        return;
      }
      if (!mainWindow.isFocused()) {
        mainWindow.focus();
      }
    },
    activatePlaybackWindowForOverlayInteraction: async () => {
      try {
        const raised = (await appState.windowTracker?.raiseTargetWindow?.()) ?? false;
        geometry.enforceOverlayLayerOrder();
        return raised;
      } catch (error) {
        logger.debug(
          'Failed to raise tracked mpv window for overlay interaction:',
          error instanceof Error ? error.message : String(error),
        );
        geometry.enforceOverlayLayerOrder();
        return false;
      }
    },
    onOverlayModalClosed: (modal, senderWindow, instanceId) => {
      // A late close for a replaced plugin panel must not hide the panel now showing.
      if (
        modal === 'plugin-panel' &&
        instanceId !== undefined &&
        !deps.getPluginPanel().handleModalClosed(instanceId)
      ) {
        return;
      }
      if (modal === 'media-timing-review') {
        void deps.getMediaTimingReviewRuntime().dispose();
      }
      if (modal === 'jimaku') {
        deps.getJimakuAutoOpenRuntime().handleJimakuModalClosed();
      }
      if (modal === 'subtitle-sidebar' && senderWindow === overlayManager.getMainWindow()) {
        deps.setSubtitleSidebarRequestedOpen(false);
      }
      const modalWindow = overlayManager.getModalWindow();
      if (
        senderWindow &&
        modalWindow &&
        senderWindow === modalWindow &&
        !senderWindow.isDestroyed()
      ) {
        applyOverlayClickThrough(senderWindow);
        senderWindow.hide();
      }
      deps.overlayModalRuntime.handleOverlayModalClosed(modal);
    },
    onOverlayModalOpened: (modal, senderWindow) => {
      if (modal === 'subtitle-sidebar' && senderWindow === overlayManager.getMainWindow()) {
        deps.setSubtitleSidebarRequestedOpen(true);
      }
      if (modal === 'jimaku') {
        deps.getJimakuAutoOpenRuntime().handleJimakuModalOpened();
      }
      deps.overlayModalRuntime.notifyOverlayModalOpened(modal);
    },
    onOverlayMouseInteractionChanged: (active, senderWindow) => {
      const mainWindow = overlayManager.getMainWindow();
      if (!mainWindow || senderWindow !== mainWindow) {
        return;
      }
      const previousActive = interaction.getVisibleOverlayInteractionActive();
      interaction.setVisibleOverlayInteractionActive(active);
      if (previousActive === active) {
        if (active && process.platform === 'darwin' && !mainWindow.isFocused()) {
          overlayVisibilityRuntime.updateVisibleOverlayVisibility();
        }
        return;
      }
      overlayVisibilityRuntime.updateVisibleOverlayVisibility();
    },
    onOverlayInteractiveHint: (interactive, senderWindow) => {
      const mainWindow = overlayManager.getMainWindow();
      if (!mainWindow || senderWindow !== mainWindow) {
        return;
      }
      interaction.setLinuxOverlayInteractiveHint(interactive);
      interaction.applyLinuxOverlayInputShapeFromLatestMeasurement();
    },
    // An accepted measurement refreshes pointer passthrough at once, then releases autoplay,
    // then schedules the first-paint subtitle refresh.
    reportOverlayContentBounds: (payload: unknown) => {
      if (deps.overlayContentMeasurementStore.report(payload)) {
        interaction.tickLinuxOverlayPointerInteractionNow();
        interaction.tickWindowsOverlayPointerInteractionNow();
        interaction.primeLinuxOverlayPointerInteractionAfterFirstMeasurement();
        deps.getAutoplayReadyGate().flushPendingAutoplayReadySignal();
        deps.getAutoplaySubtitlePriming().scheduleVisibleOverlaySubtitleRefreshAfterFirstPaint();
      }
    },
  };

  return {
    ipc,
    geometry,
    interaction,
    visibilityComposer,
    restoreOverlayMpvSubtitles,
    syncOverlayMpvSubtitleSuppression,
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
