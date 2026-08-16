import type { BrowserWindow } from 'electron';
import type { OverlayHostedModal } from '../shared/ipc/contracts';
import type { WindowGeometry } from '../types';
import type { HyprlandPlacementStatus } from '../core/services/hyprland-window-placement';
import { applyOverlayClickThrough } from '../core/services/overlay-click-through';
import {
  OVERLAY_WINDOW_CONTENT_READY_FLAG,
  OVERLAY_WINDOW_DOCUMENT_LOADED_FLAG,
} from '../core/services/overlay-window-flags';

const MODAL_REVEAL_FALLBACK_DELAY_MS = 250;
// The dedicated modal window maps asynchronously on Wayland; a single reconcile can fire
// before the compositor has a client to place, leaving the modal buried under fullscreen mpv.
// Re-assert placement across this ladder until the Hyprland client is found (or attempts run out).
const MODAL_POST_SHOW_BOUNDS_RECONCILE_DELAYS_MS = [50, 120, 250, 500, 900, 1400];

function requestOverlayApplicationFocus(): void {
  try {
    const electron = require('electron') as {
      app?: {
        focus?: (options?: { steal?: boolean }) => void;
      };
    };
    electron.app?.focus?.({ steal: true });
  } catch {
    // Ignore focus-steal failures in non-Electron test environments.
  }
}

function setWindowFocusable(window: BrowserWindow): void {
  const maybeFocusableWindow = window as BrowserWindow & {
    setFocusable?: (focusable: boolean) => void;
  };
  maybeFocusableWindow.setFocusable?.(true);
}

export interface OverlayWindowResolver {
  getMainWindow: () => BrowserWindow | null;
  getModalWindow: () => BrowserWindow | null;
  createModalWindow: () => BrowserWindow | null;
  getModalGeometry: () => WindowGeometry;
  setModalWindowBounds: (geometry: WindowGeometry) => HyprlandPlacementStatus | void;
}

export interface OverlayModalRuntime {
  primeModalWindow: () => boolean;
  sendToActiveOverlayWindow: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: {
      restoreOnModalClose?: OverlayHostedModal;
      preferModalWindow?: boolean;
    },
  ) => boolean;
  openRuntimeOptionsPalette: () => void;
  openJimaku: () => void;
  openTsukihime: () => void;
  handleOverlayModalClosed: (modal: OverlayHostedModal) => void;
  notifyOverlayModalOpened: (modal: OverlayHostedModal) => void;
  waitForModalOpen: (modal: OverlayHostedModal, timeoutMs: number) => Promise<boolean>;
  getRestoreVisibleOverlayOnModalClose: () => Set<OverlayHostedModal>;
}

type RevealFallbackHandle = NonNullable<Parameters<typeof globalThis.clearTimeout>[0]>;

export interface OverlayModalRuntimeOptions {
  platform?: NodeJS.Platform;
  focusApplication?: () => void;
  onModalStateChange?: (isActive: boolean) => void;
  onFinalModalClosed?: () => void;
  scheduleRevealFallback?: (callback: () => void, delayMs: number) => RevealFallbackHandle;
  clearRevealFallback?: (timeout: RevealFallbackHandle) => void;
}

export function createOverlayModalRuntimeService(
  deps: OverlayWindowResolver,
  options: OverlayModalRuntimeOptions = {},
): OverlayModalRuntime {
  const restoreVisibleOverlayOnModalClose = new Set<OverlayHostedModal>();
  const modalOpenWaiters = new Map<OverlayHostedModal, Array<(opened: boolean) => void>>();
  const openedModals = new Set<OverlayHostedModal>();
  let modalActive = false;
  let mainWindowMousePassthroughForcedByModal = false;
  let mainWindowHiddenByModal = false;
  let modalWindowPrimedForImmediateShow = false;
  let pendingModalWindowReveal: BrowserWindow | null = null;
  let pendingModalWindowRevealTimeout: RevealFallbackHandle | null = null;
  const modalWindowBoundsReconcileGenerations = new WeakMap<BrowserWindow, number>();
  const modalWindowPrimeListenersRegistered = new WeakSet<BrowserWindow>();
  const platform = options.platform ?? process.platform;
  const shouldPrimeModalWindow = platform === 'darwin' || platform === 'win32';
  const reuseModalWindowAfterClose = platform === 'darwin';
  const focusApplication = options.focusApplication ?? requestOverlayApplicationFocus;
  const scheduleRevealFallback = (callback: () => void, delayMs: number): RevealFallbackHandle =>
    (options.scheduleRevealFallback ?? globalThis.setTimeout)(callback, delayMs);
  const clearRevealFallback = (timeout: RevealFallbackHandle): void =>
    (options.clearRevealFallback ?? globalThis.clearTimeout)(timeout);

  const notifyModalStateChange = (nextState: boolean): void => {
    if (modalActive === nextState) return;
    modalActive = nextState;
    options.onModalStateChange?.(nextState);
  };

  const resolveModalWindow = (): BrowserWindow | null => {
    const existingWindow = deps.getModalWindow();
    if (existingWindow && !existingWindow.isDestroyed()) {
      return existingWindow;
    }
    const createdWindow = deps.createModalWindow();
    if (!createdWindow || createdWindow.isDestroyed()) {
      return null;
    }
    return createdWindow;
  };

  const getTargetOverlayWindow = (): BrowserWindow | null => {
    const visibleMainWindow = deps.getMainWindow();

    if (visibleMainWindow && !visibleMainWindow.isDestroyed() && visibleMainWindow.isVisible()) {
      return visibleMainWindow;
    }
    return null;
  };

  const getActiveOverlayWindowForModalInput = (): BrowserWindow | null => {
    const modalWindow = deps.getModalWindow();
    if (modalWindow && !modalWindow.isDestroyed()) {
      return modalWindow;
    }

    const visibleMainWindow = deps.getMainWindow();
    if (visibleMainWindow && !visibleMainWindow.isDestroyed()) {
      return visibleMainWindow;
    }

    return null;
  };

  const isWindowReadyForIpc = (window: BrowserWindow): boolean => {
    if (window.isDestroyed()) {
      return false;
    }
    if (window.webContents.isLoading()) {
      return false;
    }
    const overlayWindow = window as BrowserWindow & {
      [OVERLAY_WINDOW_CONTENT_READY_FLAG]?: boolean;
      [OVERLAY_WINDOW_DOCUMENT_LOADED_FLAG]?: boolean;
    };
    if (overlayWindow[OVERLAY_WINDOW_DOCUMENT_LOADED_FLAG] === false) {
      return false;
    }
    if (
      typeof overlayWindow[OVERLAY_WINDOW_CONTENT_READY_FLAG] === 'boolean' &&
      overlayWindow[OVERLAY_WINDOW_CONTENT_READY_FLAG] !== true
    ) {
      return false;
    }
    const currentURL = window.webContents.getURL();
    return currentURL !== '' && currentURL !== 'about:blank';
  };

  const isWindowLoadedForIpc = (window: BrowserWindow): boolean => {
    if (window.isDestroyed() || window.webContents.isLoading()) {
      return false;
    }
    const overlayWindow = window as BrowserWindow & {
      [OVERLAY_WINDOW_DOCUMENT_LOADED_FLAG]?: boolean;
    };
    if (overlayWindow[OVERLAY_WINDOW_DOCUMENT_LOADED_FLAG] !== true) {
      return false;
    }
    const currentURL = window.webContents.getURL();
    return currentURL !== '' && currentURL !== 'about:blank';
  };

  const markModalWindowPrimed = (window: BrowserWindow): void => {
    if (deps.getModalWindow() !== window || !isWindowLoadedForIpc(window)) {
      return;
    }
    modalWindowPrimedForImmediateShow = true;
  };

  const primeModalWindow = (): boolean => {
    if (!shouldPrimeModalWindow) {
      return false;
    }
    const modalWindow = resolveModalWindow();
    if (!modalWindow) {
      return false;
    }

    deps.setModalWindowBounds(deps.getModalGeometry());
    if (isWindowReadyForIpc(modalWindow)) {
      modalWindowPrimedForImmediateShow = true;
      return true;
    }

    if (!modalWindowPrimeListenersRegistered.has(modalWindow)) {
      modalWindowPrimeListenersRegistered.add(modalWindow);
      modalWindow.webContents.once('did-finish-load', () => markModalWindowPrimed(modalWindow));
      modalWindow.once('ready-to-show', () => markModalWindowPrimed(modalWindow));
    }
    return true;
  };

  const elevateModalWindow = (window: BrowserWindow): void => {
    if (window.isDestroyed()) return;
    window.setAlwaysOnTop(true, 'screen-saver', 3);
    window.moveTop();
  };

  const reconcileModalWindowBounds = (window: BrowserWindow): HyprlandPlacementStatus | void => {
    const modalWindow = deps.getModalWindow();
    if (!modalWindow || modalWindow !== window || window.isDestroyed()) {
      return;
    }
    return deps.setModalWindowBounds(deps.getModalGeometry());
  };

  const nextModalWindowBoundsReconcileGeneration = (window: BrowserWindow): number => {
    const generation = (modalWindowBoundsReconcileGenerations.get(window) ?? 0) + 1;
    modalWindowBoundsReconcileGenerations.set(window, generation);
    return generation;
  };

  const isCurrentModalWindowBoundsReconcileGeneration = (
    window: BrowserWindow,
    generation: number,
  ): boolean => modalWindowBoundsReconcileGenerations.get(window) === generation;

  const scheduleModalWindowBoundsReconcile = (
    window: BrowserWindow,
    generation: number,
    attempt = 0,
  ): void => {
    if (attempt >= MODAL_POST_SHOW_BOUNDS_RECONCILE_DELAYS_MS.length) {
      return;
    }
    const timeout = setTimeout(() => {
      if (!isCurrentModalWindowBoundsReconcileGeneration(window, generation)) {
        return;
      }
      if (window.isDestroyed() || !window.isVisible()) {
        return;
      }
      const status = reconcileModalWindowBounds(window);
      // Keep retrying only while a Hyprland placement is applicable but the compositor has
      // not mapped the window yet. Once the client is found (or we're not on Hyprland), stop.
      if (status && status.applicable && !status.clientFound) {
        scheduleModalWindowBoundsReconcile(window, generation, attempt + 1);
      }
    }, MODAL_POST_SHOW_BOUNDS_RECONCILE_DELAYS_MS[attempt]);
    timeout.unref?.();
  };

  const sendOrQueueForWindow = (
    window: BrowserWindow,
    sendNow: (window: BrowserWindow) => void,
  ): void => {
    if (isWindowReadyForIpc(window)) {
      sendNow(window);
      return;
    }

    let delivered = false;
    const deliver = (isReady: () => boolean): void => {
      if (delivered || window.isDestroyed() || !isReady()) {
        return;
      }
      delivered = true;
      sendNow(window);
    };

    // A hidden macOS panel may not emit ready-to-show until it is presented. The
    // renderer can safely receive IPC as soon as its document has finished loading.
    window.webContents.once('did-finish-load', () => deliver(() => isWindowLoadedForIpc(window)));
    window.once('ready-to-show', () => deliver(() => isWindowReadyForIpc(window)));
    deliver(() => isWindowLoadedForIpc(window));
  };

  const showModalWindow = (
    window: BrowserWindow,
    options: {
      passThroughMouseEvents: boolean;
    } = { passThroughMouseEvents: false },
  ): void => {
    setWindowFocusable(window);
    const wasVisible = window.isVisible();
    if (!wasVisible && platform === 'darwin') {
      // Mapping the panel first keeps it attached to mpv's active fullscreen Space.
      window.showInactive();
      focusApplication();
    } else {
      focusApplication();
    }
    if (!wasVisible && platform !== 'darwin') {
      window.show();
    }
    elevateModalWindow(window);
    if (options.passThroughMouseEvents) {
      applyOverlayClickThrough(window, platform === 'win32');
    } else {
      window.setIgnoreMouseEvents(false);
    }
    window.focus();
    if (!window.webContents.isFocused()) {
      window.webContents.focus();
    }
    const reconcileGeneration = nextModalWindowBoundsReconcileGeneration(window);
    reconcileModalWindowBounds(window);
    scheduleModalWindowBoundsReconcile(window, reconcileGeneration);
  };

  const ensureModalWindowInteractive = (window: BrowserWindow): void => {
    setWindowFocusable(window);
    window.setIgnoreMouseEvents(false);
    elevateModalWindow(window);

    if (window.isVisible()) {
      focusApplication();
      window.focus();
      window.webContents.focus();
      const reconcileGeneration = nextModalWindowBoundsReconcileGeneration(window);
      reconcileModalWindowBounds(window);
      scheduleModalWindowBoundsReconcile(window, reconcileGeneration);
      return;
    }

    showModalWindow(window);
  };

  const showOverlayWindowForModal = (window: BrowserWindow): void => {
    window.show();
    if (!window.isFocused()) {
      window.focus();
    }
  };

  const clearPendingModalWindowReveal = (): void => {
    if (pendingModalWindowRevealTimeout === null) {
      pendingModalWindowReveal = null;
      return;
    }

    clearRevealFallback(pendingModalWindowRevealTimeout);
    pendingModalWindowRevealTimeout = null;
    pendingModalWindowReveal = null;
  };

  const setMainWindowMousePassthroughForModal = (enabled: boolean): void => {
    const mainWindow = deps.getMainWindow();
    if (!mainWindow || mainWindow.isDestroyed()) {
      mainWindowMousePassthroughForcedByModal = false;
      return;
    }

    if (enabled) {
      if (!mainWindow.isVisible()) {
        mainWindowMousePassthroughForcedByModal = false;
        return;
      }
      applyOverlayClickThrough(mainWindow, platform === 'win32');
      mainWindowMousePassthroughForcedByModal = true;
      return;
    }

    if (!mainWindowMousePassthroughForcedByModal) {
      return;
    }
    mainWindow.setIgnoreMouseEvents(false);
    mainWindowMousePassthroughForcedByModal = false;
  };

  const setMainWindowVisibilityForModal = (hidden: boolean): void => {
    const mainWindow = deps.getMainWindow();
    if (!mainWindow || mainWindow.isDestroyed()) {
      mainWindowHiddenByModal = false;
      return;
    }

    if (hidden) {
      if (!mainWindow.isVisible()) {
        mainWindowHiddenByModal = false;
        return;
      }
      mainWindow.hide();
      mainWindowHiddenByModal = true;
      return;
    }

    if (!mainWindowHiddenByModal) {
      return;
    }
    mainWindow.show();
    mainWindowHiddenByModal = false;
  };

  const scheduleModalWindowReveal = (window: BrowserWindow): void => {
    pendingModalWindowReveal = window;
    if (pendingModalWindowRevealTimeout !== null) {
      return;
    }

    pendingModalWindowRevealTimeout = scheduleRevealFallback(() => {
      const targetWindow = pendingModalWindowReveal;
      clearPendingModalWindowReveal();
      if (!targetWindow || targetWindow.isDestroyed() || targetWindow.isVisible()) {
        return;
      }
      if (!isWindowReadyForIpc(targetWindow)) {
        return;
      }
      showModalWindow(targetWindow, { passThroughMouseEvents: false });
    }, MODAL_REVEAL_FALLBACK_DELAY_MS);
  };

  const sendToActiveOverlayWindow = (
    channel: string,
    payload?: unknown,
    runtimeOptions?: {
      restoreOnModalClose?: OverlayHostedModal;
      preferModalWindow?: boolean;
    },
  ): boolean => {
    const restoreOnModalClose = runtimeOptions?.restoreOnModalClose;
    const preferModalWindow = runtimeOptions?.preferModalWindow === true;

    const sendNow = (window: BrowserWindow): void => {
      ensureModalWindowInteractive(window);
      if (payload === undefined) {
        window.webContents.send(channel);
      } else {
        window.webContents.send(channel, payload);
      }
    };

    if (restoreOnModalClose) {
      const mainWindow = getTargetOverlayWindow();
      if (!preferModalWindow && mainWindow && !mainWindow.isDestroyed() && mainWindow.isVisible()) {
        restoreVisibleOverlayOnModalClose.add(restoreOnModalClose);
        sendOrQueueForWindow(mainWindow, (window) => {
          if (payload === undefined) {
            window.webContents.send(channel);
          } else {
            window.webContents.send(channel, payload);
          }
        });
        return true;
      }

      const modalWindow = resolveModalWindow();
      if (!modalWindow) return false;

      restoreVisibleOverlayOnModalClose.add(restoreOnModalClose);
      deps.setModalWindowBounds(deps.getModalGeometry());
      const wasVisible = modalWindow.isVisible();
      if (!wasVisible) {
        if (modalWindowPrimedForImmediateShow && isWindowReadyForIpc(modalWindow)) {
          showModalWindow(modalWindow);
        } else {
          scheduleModalWindowReveal(modalWindow);
        }
      } else if (!modalWindow.isFocused()) {
        showModalWindow(modalWindow);
      }

      sendOrQueueForWindow(modalWindow, (window) => {
        if (window.isVisible()) {
          ensureModalWindowInteractive(window);
        }
        if (payload === undefined) {
          window.webContents.send(channel);
        } else {
          window.webContents.send(channel, payload);
        }
      });
      return true;
    }

    const target = getTargetOverlayWindow();
    if (!target) return false;

    const wasVisible = target.isVisible();
    if (!wasVisible) {
      showOverlayWindowForModal(target);
    }

    sendOrQueueForWindow(target, sendNow);
    return true;
  };

  const openRuntimeOptionsPalette = (): void => {
    sendToActiveOverlayWindow('runtime-options:open', undefined, {
      restoreOnModalClose: 'runtime-options',
    });
  };

  const openJimaku = (): void => {
    sendToActiveOverlayWindow('jimaku:open', undefined, {
      restoreOnModalClose: 'jimaku',
    });
  };

  const openTsukihime = (): void => {
    sendToActiveOverlayWindow('tsukihime:open', undefined, {
      restoreOnModalClose: 'tsukihime',
    });
  };

  const handleOverlayModalClosed = (modal: OverlayHostedModal): void => {
    openedModals.delete(modal);
    if (!restoreVisibleOverlayOnModalClose.has(modal)) return;
    restoreVisibleOverlayOnModalClose.delete(modal);
    const modalWindow = deps.getModalWindow();
    if (restoreVisibleOverlayOnModalClose.size === 0) {
      clearPendingModalWindowReveal();
      if (modalWindow && !modalWindow.isDestroyed()) {
        if (reuseModalWindowAfterClose) {
          applyOverlayClickThrough(modalWindow, false);
          modalWindow.hide();
          markModalWindowPrimed(modalWindow);
        } else {
          modalWindow.destroy();
          modalWindowPrimedForImmediateShow = false;
          // Reusing a transparent click-through BrowserWindow can leave later modal sessions
          // non-interactive on Windows. Recycle the renderer after every close, then warm its
          // replacement so the next shortcut still opens promptly.
          if (platform === 'win32') {
            primeModalWindow();
          }
        }
      }
      mainWindowMousePassthroughForcedByModal = false;
      setMainWindowVisibilityForModal(false);
      try {
        options.onFinalModalClosed?.();
      } catch {
        // Modal state still needs to deactivate if focus handoff fails.
      } finally {
        notifyModalStateChange(false);
      }
    }
  };

  const notifyOverlayModalOpened = (modal: OverlayHostedModal): void => {
    if (!restoreVisibleOverlayOnModalClose.has(modal)) return;
    openedModals.add(modal);
    const waiters = modalOpenWaiters.get(modal) ?? [];
    modalOpenWaiters.delete(modal);
    for (const resolve of waiters) {
      resolve(true);
    }
    notifyModalStateChange(true);
    const targetWindow = getActiveOverlayWindowForModalInput();
    clearPendingModalWindowReveal();
    if (!targetWindow || targetWindow.isDestroyed()) {
      return;
    }

    const modalWindow = deps.getModalWindow();
    if (targetWindow.isVisible()) {
      ensureModalWindowInteractive(targetWindow);
    } else {
      showModalWindow(targetWindow);
    }

    if (modalWindow && !modalWindow.isDestroyed() && targetWindow === modalWindow) {
      setMainWindowMousePassthroughForModal(true);
      setMainWindowVisibilityForModal(true);
    }
  };

  const waitForModalOpen = async (modal: OverlayHostedModal, timeoutMs: number): Promise<boolean> =>
    await new Promise<boolean>((resolve) => {
      if (openedModals.has(modal)) {
        resolve(true);
        return;
      }
      const waiters = modalOpenWaiters.get(modal) ?? [];
      const finish = (opened: boolean): void => {
        clearTimeout(timeout);
        resolve(opened);
      };
      waiters.push(finish);
      modalOpenWaiters.set(modal, waiters);
      const timeout = setTimeout(() => {
        const current = modalOpenWaiters.get(modal) ?? [];
        modalOpenWaiters.set(
          modal,
          current.filter((candidate) => candidate !== finish),
        );
        resolve(false);
      }, timeoutMs);
    });

  return {
    primeModalWindow,
    sendToActiveOverlayWindow,
    openRuntimeOptionsPalette,
    openJimaku,
    openTsukihime,
    handleOverlayModalClosed,
    notifyOverlayModalOpened,
    waitForModalOpen,
    getRestoreVisibleOverlayOnModalClose: () => restoreVisibleOverlayOnModalClose,
  };
}
