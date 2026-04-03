import type { Session } from 'electron';
import type { OverlayHostedModal } from '../shared/ipc/contracts';
import type { RuntimeOptionState, WindowGeometry } from '../types';
import type { OverlayModalRuntime } from './overlay-runtime';
import {
  normalizeOverlayUiRuntimeInput,
  type OverlayUiRuntimeInputLike,
} from './overlay-ui-runtime-input';
import {
  createOverlayVisibilityRuntimeBridge,
  type OverlayUiVisibilityBridgeWindowLike,
} from './overlay-ui-visibility';
import { createOverlayVisibilityRuntime } from './runtime/overlay-visibility-runtime';
import { createOverlayWindowRuntimeHandlers } from './runtime/overlay-window-runtime-handlers';
import { createTrayRuntimeHandlers } from './runtime/tray-runtime-handlers';
import { createOverlayRuntimeBootstrapHandlers } from './runtime/overlay-runtime-bootstrap-handlers';
import { composeOverlayVisibilityRuntime } from './runtime/composers/overlay-visibility-runtime-composer';
import {
  createBuildBroadcastRuntimeOptionsChangedMainDepsHandler,
  createBuildGetRuntimeOptionsStateMainDepsHandler,
  createBuildOpenRuntimeOptionsPaletteMainDepsHandler,
  createBuildRestorePreviousSecondarySubVisibilityMainDepsHandler,
  createBuildSendToActiveOverlayWindowMainDepsHandler,
  createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler,
} from './runtime/overlay-runtime-main-actions-main-deps';
import { createGetRuntimeOptionsStateHandler } from './runtime/overlay-runtime-main-actions';

type OverlayWindowKind = 'visible' | 'modal';

type WindowLike = OverlayUiVisibilityBridgeWindowLike;

type RuntimeOptionsManagerLike = {
  listOptions: () => RuntimeOptionState[];
};

type MpvClientLike = {
  connected: boolean;
  restorePreviousSecondarySubVisibility: () => void;
};

type TrayHandlersDeps = Parameters<typeof createTrayRuntimeHandlers>[0];
type BootstrapHandlersDeps = Parameters<typeof createOverlayRuntimeBootstrapHandlers>[0];

type OverlayWindowCreateOptions<TWindow extends WindowLike> = {
  isDev: boolean;
  ensureOverlayWindowLevel: (window: TWindow) => void;
  onRuntimeOptionsChanged: () => void;
  setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
  isOverlayVisible: (windowKind: OverlayWindowKind) => boolean;
  tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
  forwardTabToMpv: () => void;
  onWindowClosed: (windowKind: OverlayWindowKind) => void;
  yomitanSession?: Electron.Session | null;
};

export interface OverlayUiWindowState<TWindow extends WindowLike = WindowLike> {
  getMainWindow: () => TWindow | null;
  setMainWindow: (window: TWindow | null) => void;
  getModalWindow: () => TWindow | null;
  setModalWindow: (window: TWindow | null) => void;
  getVisibleOverlayVisible: () => boolean;
  setVisibleOverlayVisible: (visible: boolean) => void;
  getOverlayDebugVisualizationEnabled: () => boolean;
  setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
}

export interface OverlayUiGeometryInput {
  getCurrentOverlayGeometry: () => WindowGeometry;
}

export interface OverlayUiModalInput {
  setModalWindowBounds?: (geometry: WindowGeometry) => void;
  onModalStateChange?: (active: boolean) => void;
}

export interface OverlayUiVisibilityServiceInput<TWindow extends WindowLike = WindowLike> {
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
  resolveFallbackBounds: () => WindowGeometry;
}

export interface OverlayUiWindowsInput<TWindow extends WindowLike = WindowLike> {
  createOverlayWindowCore: (
    kind: OverlayWindowKind,
    options: OverlayWindowCreateOptions<TWindow>,
  ) => TWindow;
  isDev: boolean;
  ensureOverlayWindowLevel: (window: TWindow) => void;
  onRuntimeOptionsChanged: () => void;
  setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
  isOverlayVisible: (windowKind: OverlayWindowKind) => boolean;
  getYomitanSession: () => Session | null;
  tryHandleOverlayShortcutLocalFallback: (input: Electron.Input) => boolean;
  forwardTabToMpv: () => void;
  onWindowClosed: (windowKind: OverlayWindowKind) => void;
}

export interface OverlayUiVisibilityActionsInput {
  setVisibleOverlayVisibleCore: (options: {
    visible: boolean;
    setVisibleOverlayVisibleState: (visible: boolean) => void;
    updateVisibleOverlayVisibility: () => void;
  }) => void;
}

export interface OverlayUiActionsInput {
  getRuntimeOptionsManager: () => RuntimeOptionsManagerLike | null;
  getMpvClient: () => MpvClientLike | null;

  broadcastRuntimeOptionsChangedRuntime: (
    getRuntimeOptionsState: () => RuntimeOptionState[],
    broadcastToOverlayWindows: (channel: string, ...args: unknown[]) => void,
  ) => void;
  broadcastToOverlayWindows: (channel: string, ...args: unknown[]) => void;
  setOverlayDebugVisualizationEnabledRuntime: (
    currentEnabled: boolean,
    nextEnabled: boolean,
    setCurrentEnabled: (enabled: boolean) => void,
  ) => void;
}

export interface OverlayUiTrayInput {
  resolveTrayIconPathDeps: TrayHandlersDeps['resolveTrayIconPathDeps'];
  buildTrayMenuTemplateDeps: TrayHandlersDeps['buildTrayMenuTemplateDeps'];
  ensureTrayDeps: TrayHandlersDeps['ensureTrayDeps'];
  destroyTrayDeps: TrayHandlersDeps['destroyTrayDeps'];
  buildMenuFromTemplate: TrayHandlersDeps['buildMenuFromTemplate'];
}

export interface OverlayUiBootstrapInput {
  initializeOverlayRuntimeMainDeps: BootstrapHandlersDeps['initializeOverlayRuntimeMainDeps'];
  initializeOverlayRuntimeBootstrapDeps: BootstrapHandlersDeps['initializeOverlayRuntimeBootstrapDeps'];
  onInitialized?: () => void;
}

export interface OverlayUiRuntimeStateInput {
  isOverlayRuntimeInitialized: () => boolean;
  setOverlayRuntimeInitialized: (initialized: boolean) => void;
}

export interface OverlayUiMpvSubtitleInput {
  ensureOverlayMpvSubtitlesHidden: () => Promise<void> | void;
  syncOverlayMpvSubtitleSuppression: () => void;
}

export interface OverlayUiRuntimeInput<TWindow extends WindowLike = WindowLike> {
  windowState: OverlayUiWindowState<TWindow>;
  geometry: OverlayUiGeometryInput;
  modal: OverlayUiModalInput;
  modalRuntime: OverlayModalRuntime;
  visibilityService: OverlayUiVisibilityServiceInput<TWindow>;
  overlayWindows: OverlayUiWindowsInput<TWindow>;
  visibilityActions: OverlayUiVisibilityActionsInput;
  overlayActions: OverlayUiActionsInput;
  tray: OverlayUiTrayInput | null;
  bootstrap: OverlayUiBootstrapInput;
  runtimeState: OverlayUiRuntimeStateInput;
  mpvSubtitle: OverlayUiMpvSubtitleInput;
}

export interface OverlayUiRuntime<TWindow extends WindowLike = WindowLike> {
  createMainWindow: () => TWindow;
  createModalWindow: () => TWindow;
  ensureTray: () => void;
  destroyTray: () => void;
  initializeOverlayRuntime: () => void;
  ensureOverlayWindowsReadyForVisibilityActions: () => void;
  setVisibleOverlayVisible: (visible: boolean) => void;
  toggleVisibleOverlay: () => void;
  setOverlayVisible: (visible: boolean) => void;
  handleOverlayModalClosed: (modal: OverlayHostedModal) => void;
  notifyOverlayModalOpened: (modal: OverlayHostedModal) => void;
  waitForModalOpen: (modal: OverlayHostedModal, timeoutMs: number) => Promise<boolean>;
  getRestoreVisibleOverlayOnModalClose: () => Set<OverlayHostedModal>;
  updateVisibleOverlayVisibility: () => void;
  sendToActiveOverlayWindow: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: {
      restoreOnModalClose?: OverlayHostedModal;
      preferModalWindow?: boolean;
    },
  ) => boolean;
  openRuntimeOptionsPalette: () => void;
  broadcastRuntimeOptionsChanged: () => void;
  setOverlayDebugVisualizationEnabled: (enabled: boolean) => void;
  restorePreviousSecondarySubVisibility: () => void;
}

export function createOverlayUiRuntime<TWindow extends WindowLike>(
  input: OverlayUiRuntimeInputLike<TWindow>,
): OverlayUiRuntime<TWindow> {
  const runtimeInput = normalizeOverlayUiRuntimeInput(input);
  const overlayVisibilityRuntime = createOverlayVisibilityRuntimeBridge({
    getMainWindow: () => runtimeInput.windowState.getMainWindow(),
    getVisibleOverlayVisible: () => runtimeInput.windowState.getVisibleOverlayVisible(),
    getModalActive: () => runtimeInput.visibilityService.getModalActive(),
    getForceMousePassthrough: () => runtimeInput.visibilityService.getForceMousePassthrough(),
    getWindowTracker: () => runtimeInput.visibilityService.getWindowTracker(),
    getTrackerNotReadyWarningShown: () =>
      runtimeInput.visibilityService.getTrackerNotReadyWarningShown(),
    setTrackerNotReadyWarningShown: (shown) =>
      runtimeInput.visibilityService.setTrackerNotReadyWarningShown(shown),
    updateVisibleOverlayBounds: (geometry) =>
      runtimeInput.visibilityService.updateVisibleOverlayBounds(geometry),
    ensureOverlayWindowLevel: (window) =>
      runtimeInput.visibilityService.ensureOverlayWindowLevel(window),
    syncPrimaryOverlayWindowLayer: (layer) =>
      runtimeInput.visibilityService.syncPrimaryOverlayWindowLayer(layer),
    enforceOverlayLayerOrder: () => runtimeInput.visibilityService.enforceOverlayLayerOrder(),
    syncOverlayShortcuts: () => runtimeInput.visibilityService.syncOverlayShortcuts(),
    isMacOSPlatform: () => runtimeInput.visibilityService.isMacOSPlatform(),
    isWindowsPlatform: () => runtimeInput.visibilityService.isWindowsPlatform(),
    showOverlayLoadingOsd: (message) =>
      runtimeInput.visibilityService.showOverlayLoadingOsd(message),
  });

  const overlayWindowHandlers = createOverlayWindowRuntimeHandlers<TWindow>({
    createOverlayWindowDeps: {
      createOverlayWindowCore: (kind, options) =>
        runtimeInput.overlayWindows.createOverlayWindowCore(kind, options),
      isDev: runtimeInput.overlayWindows.isDev,
      ensureOverlayWindowLevel: (window) =>
        runtimeInput.overlayWindows.ensureOverlayWindowLevel(window),
      onRuntimeOptionsChanged: () => runtimeInput.overlayWindows.onRuntimeOptionsChanged(),
      setOverlayDebugVisualizationEnabled: (enabled) =>
        runtimeInput.overlayWindows.setOverlayDebugVisualizationEnabled(enabled),
      isOverlayVisible: (windowKind) => runtimeInput.overlayWindows.isOverlayVisible(windowKind),
      getYomitanSession: () => runtimeInput.overlayWindows.getYomitanSession(),
      tryHandleOverlayShortcutLocalFallback: (overlayInput) =>
        runtimeInput.overlayWindows.tryHandleOverlayShortcutLocalFallback(overlayInput),
      forwardTabToMpv: () => runtimeInput.overlayWindows.forwardTabToMpv(),
      onWindowClosed: (windowKind) => runtimeInput.overlayWindows.onWindowClosed(windowKind),
    },
    setMainWindow: (window) => runtimeInput.windowState.setMainWindow(window),
    setModalWindow: (window) => runtimeInput.windowState.setModalWindow(window),
  });

  const visibilityActions = createOverlayVisibilityRuntime({
    setVisibleOverlayVisibleDeps: {
      setVisibleOverlayVisibleCore: (options) =>
        runtimeInput.visibilityActions.setVisibleOverlayVisibleCore(options),
      setVisibleOverlayVisibleState: (visible) =>
        runtimeInput.windowState.setVisibleOverlayVisible(visible),
      updateVisibleOverlayVisibility: () =>
        overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    },
    getVisibleOverlayVisible: () => runtimeInput.windowState.getVisibleOverlayVisible(),
  });

  const getRuntimeOptionsState = createGetRuntimeOptionsStateHandler(
    createBuildGetRuntimeOptionsStateMainDepsHandler({
      getRuntimeOptionsManager: () => runtimeInput.overlayActions.getRuntimeOptionsManager(),
    })(),
  );

  const overlayActions = composeOverlayVisibilityRuntime({
    overlayVisibilityRuntime,
    restorePreviousSecondarySubVisibilityMainDeps:
      createBuildRestorePreviousSecondarySubVisibilityMainDepsHandler({
        getMpvClient: () => runtimeInput.overlayActions.getMpvClient(),
      })(),
    broadcastRuntimeOptionsChangedMainDeps:
      createBuildBroadcastRuntimeOptionsChangedMainDepsHandler({
        broadcastRuntimeOptionsChangedRuntime: (getState, broadcast) =>
          runtimeInput.overlayActions.broadcastRuntimeOptionsChangedRuntime(getState, broadcast),
        getRuntimeOptionsState: () => getRuntimeOptionsState(),
        broadcastToOverlayWindows: (channel, ...args) =>
          runtimeInput.overlayActions.broadcastToOverlayWindows(channel, ...args),
      })(),
    sendToActiveOverlayWindowMainDeps: createBuildSendToActiveOverlayWindowMainDepsHandler({
      sendToActiveOverlayWindowRuntime: (channel, payload, runtimeOptions) =>
        runtimeInput.modalRuntime.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
    })(),
    setOverlayDebugVisualizationEnabledMainDeps:
      createBuildSetOverlayDebugVisualizationEnabledMainDepsHandler({
        setOverlayDebugVisualizationEnabledRuntime: (currentEnabled, nextEnabled, setCurrent) =>
          runtimeInput.overlayActions.setOverlayDebugVisualizationEnabledRuntime(
            currentEnabled,
            nextEnabled,
            setCurrent,
          ),
        getCurrentEnabled: () => runtimeInput.windowState.getOverlayDebugVisualizationEnabled(),
        setCurrentEnabled: (enabled) =>
          runtimeInput.windowState.setOverlayDebugVisualizationEnabled(enabled),
      })(),
    openRuntimeOptionsPaletteMainDeps: createBuildOpenRuntimeOptionsPaletteMainDepsHandler({
      openRuntimeOptionsPaletteRuntime: () => runtimeInput.modalRuntime.openRuntimeOptionsPalette(),
    })(),
  });

  const trayHandlers = runtimeInput.tray
    ? createTrayRuntimeHandlers({
        resolveTrayIconPathDeps: runtimeInput.tray.resolveTrayIconPathDeps,
        buildTrayMenuTemplateDeps: runtimeInput.tray.buildTrayMenuTemplateDeps,
        ensureTrayDeps: runtimeInput.tray.ensureTrayDeps,
        destroyTrayDeps: runtimeInput.tray.destroyTrayDeps,
        buildMenuFromTemplate: (template) => runtimeInput.tray!.buildMenuFromTemplate(template),
      })
    : null;

  const { initializeOverlayRuntime: initializeOverlayRuntimeHandler } =
    createOverlayRuntimeBootstrapHandlers({
      initializeOverlayRuntimeMainDeps: runtimeInput.bootstrap.initializeOverlayRuntimeMainDeps,
      initializeOverlayRuntimeBootstrapDeps:
        runtimeInput.bootstrap.initializeOverlayRuntimeBootstrapDeps,
    });

  function createMainWindow(): TWindow {
    return overlayWindowHandlers.createMainWindow();
  }

  function createModalWindow(): TWindow {
    const existingWindow = runtimeInput.windowState.getModalWindow();
    if (existingWindow && !existingWindow.isDestroyed()) {
      return existingWindow;
    }
    const window = overlayWindowHandlers.createModalWindow();
    runtimeInput.modal.setModalWindowBounds?.(runtimeInput.geometry.getCurrentOverlayGeometry());
    return window;
  }

  function initializeOverlayRuntime(): void {
    initializeOverlayRuntimeHandler();
    runtimeInput.bootstrap.onInitialized?.();
    runtimeInput.mpvSubtitle.syncOverlayMpvSubtitleSuppression();
  }

  function ensureOverlayWindowsReadyForVisibilityActions(): void {
    if (!runtimeInput.runtimeState.isOverlayRuntimeInitialized()) {
      initializeOverlayRuntime();
      return;
    }

    const mainWindow = runtimeInput.windowState.getMainWindow();
    if (!mainWindow || mainWindow.isDestroyed()) {
      createMainWindow();
    }
  }

  function setVisibleOverlayVisible(visible: boolean): void {
    ensureOverlayWindowsReadyForVisibilityActions();
    if (visible) {
      void runtimeInput.mpvSubtitle.ensureOverlayMpvSubtitlesHidden();
    }
    visibilityActions.setVisibleOverlayVisible(visible);
    runtimeInput.mpvSubtitle.syncOverlayMpvSubtitleSuppression();
  }

  function toggleVisibleOverlay(): void {
    ensureOverlayWindowsReadyForVisibilityActions();
    if (!runtimeInput.windowState.getVisibleOverlayVisible()) {
      void runtimeInput.mpvSubtitle.ensureOverlayMpvSubtitlesHidden();
    }
    visibilityActions.toggleVisibleOverlay();
    runtimeInput.mpvSubtitle.syncOverlayMpvSubtitleSuppression();
  }

  function setOverlayVisible(visible: boolean): void {
    ensureOverlayWindowsReadyForVisibilityActions();
    if (visible) {
      void runtimeInput.mpvSubtitle.ensureOverlayMpvSubtitlesHidden();
    }
    visibilityActions.setOverlayVisible(visible);
    runtimeInput.mpvSubtitle.syncOverlayMpvSubtitleSuppression();
  }

  return {
    createMainWindow,
    createModalWindow,
    ensureTray: () => {
      trayHandlers?.ensureTray();
    },
    destroyTray: () => {
      trayHandlers?.destroyTray();
    },
    initializeOverlayRuntime,
    ensureOverlayWindowsReadyForVisibilityActions,
    setVisibleOverlayVisible,
    toggleVisibleOverlay,
    setOverlayVisible,
    handleOverlayModalClosed: (modal) => runtimeInput.modalRuntime.handleOverlayModalClosed(modal),
    notifyOverlayModalOpened: (modal) => runtimeInput.modalRuntime.notifyOverlayModalOpened(modal),
    waitForModalOpen: (modal, timeoutMs) =>
      runtimeInput.modalRuntime.waitForModalOpen(modal, timeoutMs),
    getRestoreVisibleOverlayOnModalClose: () =>
      runtimeInput.modalRuntime.getRestoreVisibleOverlayOnModalClose(),
    updateVisibleOverlayVisibility: () => overlayVisibilityRuntime.updateVisibleOverlayVisibility(),
    sendToActiveOverlayWindow: (channel, payload, runtimeOptions) =>
      overlayActions.sendToActiveOverlayWindow(channel, payload, runtimeOptions),
    openRuntimeOptionsPalette: () => overlayActions.openRuntimeOptionsPalette(),
    broadcastRuntimeOptionsChanged: () => overlayActions.broadcastRuntimeOptionsChanged(),
    setOverlayDebugVisualizationEnabled: (enabled) =>
      overlayActions.setOverlayDebugVisualizationEnabled(enabled),
    restorePreviousSecondarySubVisibility: () =>
      overlayActions.restorePreviousSecondarySubVisibility(),
  };
}
