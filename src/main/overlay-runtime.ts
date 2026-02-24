import type { BrowserWindow } from 'electron';
import type { WindowGeometry } from '../types';

type OverlayHostedModal = 'runtime-options' | 'subsync' | 'jimaku' | 'kiku';
type OverlayHostLayer = 'visible' | 'invisible';

export interface OverlayWindowResolver {
  getMainWindow: () => BrowserWindow | null;
  getInvisibleWindow: () => BrowserWindow | null;
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

  const getTargetOverlayWindow = (): {
    window: BrowserWindow;
    layer: OverlayHostLayer;
  } | null => {
    const visibleMainWindow = deps.getMainWindow();
    const invisibleWindow = deps.getInvisibleWindow();

    if (visibleMainWindow && !visibleMainWindow.isDestroyed()) {
      return { window: visibleMainWindow, layer: 'visible' };
    }

    if (invisibleWindow && !invisibleWindow.isDestroyed()) {
      return { window: invisibleWindow, layer: 'invisible' };
    }

    return null;
  };

  const showModalWindow = (window: BrowserWindow): void => {
    window.show();
    window.setIgnoreMouseEvents(false);
    window.focus();
    if (!window.webContents.isFocused()) {
      window.webContents.focus();
    }
  };

  const showOverlayWindowForModal = (window: BrowserWindow, layer: OverlayHostLayer): void => {
    if (layer === 'invisible' && typeof window.showInactive === 'function') {
      window.showInactive();
    } else {
      window.show();
    }
    if (!window.isFocused()) {
      window.focus();
    }
  };

  const sendToActiveOverlayWindow = (
    channel: string,
    payload?: unknown,
    runtimeOptions?: { restoreOnModalClose?: OverlayHostedModal },
  ): boolean => {
    const restoreOnModalClose = runtimeOptions?.restoreOnModalClose;

    const sendNow = (window: BrowserWindow): void => {
      if (payload === undefined) {
        window.webContents.send(channel);
      } else {
        window.webContents.send(channel, payload);
      }
    };

    if (restoreOnModalClose) {
      const modalWindow = resolveModalWindow();
      if (!modalWindow) return false;

      deps.setModalWindowBounds(deps.getModalGeometry());
      const wasVisible = modalWindow.isVisible();
      const wasModalActive = restoreVisibleOverlayOnModalClose.size > 0;
      restoreVisibleOverlayOnModalClose.add(restoreOnModalClose);
      if (!wasModalActive) {
        notifyModalStateChange(true);
      }

      if (!wasVisible) {
        showModalWindow(modalWindow);
      } else if (!modalWindow.isFocused()) {
        showModalWindow(modalWindow);
      }

      if (modalWindow.webContents.isLoading()) {
        modalWindow.webContents.once('did-finish-load', () => {
          if (modalWindow && !modalWindow.isDestroyed() && !modalWindow.webContents.isLoading()) {
            sendNow(modalWindow);
          }
        });
        return true;
      }

      sendNow(modalWindow);
      return true;
    }

    const target = getTargetOverlayWindow();
    if (!target) return false;

    const { window: targetWindow, layer } = target;
    const wasVisible = targetWindow.isVisible();
    if (!wasVisible) {
      showOverlayWindowForModal(targetWindow, layer);
    }

    if (targetWindow.webContents.isLoading()) {
      targetWindow.webContents.once('did-finish-load', () => {
        if (targetWindow && !targetWindow.isDestroyed() && !targetWindow.webContents.isLoading()) {
          sendNow(targetWindow);
        }
      });
      return true;
    }

    sendNow(targetWindow);
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
    if (!modalWindow || modalWindow.isDestroyed()) return;
    if (restoreVisibleOverlayOnModalClose.size === 0) {
      notifyModalStateChange(false);
    }
    if (restoreVisibleOverlayOnModalClose.size === 0) {
      modalWindow.hide();
    }
  };

  return {
    sendToActiveOverlayWindow,
    openRuntimeOptionsPalette,
    handleOverlayModalClosed,
    getRestoreVisibleOverlayOnModalClose: () => restoreVisibleOverlayOnModalClose,
  };
}

export type { OverlayHostedModal };
