import type { BrowserWindow } from 'electron';
import type { WindowGeometry } from '../types';

const MODAL_REVEAL_FALLBACK_DELAY_MS = 250;

type OverlayHostedModal =
  | 'runtime-options'
  | 'subsync'
  | 'jimaku'
  | 'kiku'
  | 'controller-select'
  | 'controller-debug';

export interface OverlayWindowResolver {
  getMainWindow: () => BrowserWindow | null;
  getModalWindow: () => BrowserWindow | null;
  createModalWindow: () => BrowserWindow | null;
  getModalGeometry: () => WindowGeometry;
  setModalWindowBounds: (geometry: WindowGeometry) => void;
}

export interface OverlayModalRuntime {
  sendToActiveOverlayWindow: (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: OverlayHostedModal },
  ) => boolean;
  openRuntimeOptionsPalette: () => void;
  handleOverlayModalClosed: (modal: OverlayHostedModal) => void;
  notifyOverlayModalOpened: (modal: OverlayHostedModal) => void;
  getRestoreVisibleOverlayOnModalClose: () => Set<OverlayHostedModal>;
}

export interface OverlayModalRuntimeOptions {
  onModalStateChange?: (isActive: boolean) => void;
}

export function createOverlayModalRuntimeService(
  deps: OverlayWindowResolver,
  options: OverlayModalRuntimeOptions = {},
): OverlayModalRuntime {
  const restoreVisibleOverlayOnModalClose = new Set<OverlayHostedModal>();
  let modalActive = false;
  let pendingModalWindowReveal: BrowserWindow | null = null;
  let pendingModalWindowRevealTimeout: ReturnType<typeof setTimeout> | null = null;

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
    if (window.webContents.isLoading()) {
      return false;
    }
    const currentURL = window.webContents.getURL();
    return currentURL !== '' && currentURL !== 'about:blank';
  };

  const elevateModalWindow = (window: BrowserWindow): void => {
    if (window.isDestroyed()) return;
    window.setAlwaysOnTop(true, 'screen-saver', 1);
    window.moveTop();
  };

  const sendOrQueueForWindow = (
    window: BrowserWindow,
    sendNow: (window: BrowserWindow) => void,
  ): void => {
    if (isWindowReadyForIpc(window)) {
      sendNow(window);
      return;
    }

    window.webContents.once('did-finish-load', () => {
      if (!window.isDestroyed() && !window.webContents.isLoading()) {
        sendNow(window);
      }
    });
  };

  const showModalWindow = (
    window: BrowserWindow,
    options: {
      passThroughMouseEvents: boolean;
    } = { passThroughMouseEvents: false },
  ): void => {
    if (!window.isVisible()) {
      window.show();
    }
    elevateModalWindow(window);
    if (options.passThroughMouseEvents) {
      window.setIgnoreMouseEvents(true, { forward: true });
    } else {
      window.setIgnoreMouseEvents(false);
    }
    window.focus();
    if (!window.webContents.isFocused()) {
      window.webContents.focus();
    }
  };

  const ensureModalWindowInteractive = (window: BrowserWindow): void => {
    if (window.isVisible()) {
      window.setIgnoreMouseEvents(false);
      if (!window.isFocused()) {
        window.focus();
      }
      if (!window.webContents.isFocused()) {
        window.webContents.focus();
      }
      elevateModalWindow(window);
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

    clearTimeout(pendingModalWindowRevealTimeout);
    pendingModalWindowRevealTimeout = null;
    pendingModalWindowReveal = null;
  };

  const scheduleModalWindowReveal = (window: BrowserWindow): void => {
    pendingModalWindowReveal = window;
    if (pendingModalWindowRevealTimeout !== null) {
      return;
    }

    pendingModalWindowRevealTimeout = setTimeout(() => {
      const targetWindow = pendingModalWindowReveal;
      clearPendingModalWindowReveal();
      if (!targetWindow || targetWindow.isDestroyed() || targetWindow.isVisible()) {
        return;
      }
      showModalWindow(targetWindow, { passThroughMouseEvents: false });
    }, MODAL_REVEAL_FALLBACK_DELAY_MS);
  };

  const sendToActiveOverlayWindow = (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: OverlayHostedModal },
  ): boolean => {
    const restoreOnModalClose = runtimeOptions?.restoreOnModalClose;

    const sendNow = (window: BrowserWindow): void => {
      ensureModalWindowInteractive(window);
      if (payload === undefined) {
        window.webContents.send(channel);
      } else {
        window.webContents.send(channel, payload);
      }
    };

    if (restoreOnModalClose) {
      restoreVisibleOverlayOnModalClose.add(restoreOnModalClose);
      const mainWindow = getTargetOverlayWindow();
      if (mainWindow && !mainWindow.isDestroyed() && mainWindow.isVisible()) {
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

      deps.setModalWindowBounds(deps.getModalGeometry());
      const wasVisible = modalWindow.isVisible();
      if (!wasVisible) {
        scheduleModalWindowReveal(modalWindow);
      } else if (!modalWindow.isFocused()) {
        showModalWindow(modalWindow);
      }

      sendOrQueueForWindow(modalWindow, (window) => {
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

  const handleOverlayModalClosed = (modal: OverlayHostedModal): void => {
    if (!restoreVisibleOverlayOnModalClose.has(modal)) return;
    restoreVisibleOverlayOnModalClose.delete(modal);
    const modalWindow = deps.getModalWindow();
    if (restoreVisibleOverlayOnModalClose.size === 0) {
      clearPendingModalWindowReveal();
      notifyModalStateChange(false);
      if (modalWindow && !modalWindow.isDestroyed()) {
        modalWindow.hide();
      }
    }
  };

  const notifyOverlayModalOpened = (modal: OverlayHostedModal): void => {
    if (!restoreVisibleOverlayOnModalClose.has(modal)) return;
    notifyModalStateChange(true);
    const targetWindow = getActiveOverlayWindowForModalInput();
    clearPendingModalWindowReveal();
    if (!targetWindow || targetWindow.isDestroyed()) {
      return;
    }

    if (targetWindow.isVisible()) {
      targetWindow.setIgnoreMouseEvents(false);
      elevateModalWindow(targetWindow);
      if (!targetWindow.isFocused()) {
        targetWindow.focus();
      }
      if (!targetWindow.webContents.isFocused()) {
        targetWindow.webContents.focus();
      }
      return;
    }

    showModalWindow(targetWindow);
  };

  return {
    sendToActiveOverlayWindow,
    openRuntimeOptionsPalette,
    handleOverlayModalClosed,
    notifyOverlayModalOpened,
    getRestoreVisibleOverlayOnModalClose: () => restoreVisibleOverlayOnModalClose,
  };
}

export type { OverlayHostedModal };
