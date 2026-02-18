import type { BrowserWindow } from "electron";

type OverlayHostedModal = "runtime-options" | "subsync" | "jimaku";
type OverlayHostLayer = "visible" | "invisible";

export interface OverlayWindowResolver {
  getMainWindow: () => BrowserWindow | null;
  getInvisibleWindow: () => BrowserWindow | null;
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

export function createOverlayModalRuntimeService(
  deps: OverlayWindowResolver,
): OverlayModalRuntime {
  const restoreVisibleOverlayOnModalClose = new Set<OverlayHostedModal>();
  const overlayModalAutoShownLayer = new Map<
    OverlayHostedModal,
    OverlayHostLayer
  >();

  const getTargetOverlayWindow = (): {
    window: BrowserWindow;
    layer: OverlayHostLayer;
  } | null => {
    const visibleMainWindow = deps.getMainWindow();
    const invisibleWindow = deps.getInvisibleWindow();

    if (visibleMainWindow && !visibleMainWindow.isDestroyed()) {
      return { window: visibleMainWindow, layer: "visible" };
    }

    if (invisibleWindow && !invisibleWindow.isDestroyed()) {
      return { window: invisibleWindow, layer: "invisible" };
    }

    return null;
  };

  const showOverlayWindowForModal = (
    window: BrowserWindow,
    layer: OverlayHostLayer,
  ): void => {
    if (layer === "invisible" && typeof window.showInactive === "function") {
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
    const target = getTargetOverlayWindow();
    if (!target) return false;

    const { window: targetWindow, layer } = target;
    const wasVisible = targetWindow.isVisible();
    const restoreOnModalClose = runtimeOptions?.restoreOnModalClose;

    const sendNow = (): void => {
      if (payload === undefined) {
        targetWindow.webContents.send(channel);
      } else {
        targetWindow.webContents.send(channel, payload);
      }
    };

    if (!wasVisible) {
      showOverlayWindowForModal(targetWindow, layer);
    }
    if (!wasVisible && restoreOnModalClose) {
      restoreVisibleOverlayOnModalClose.add(restoreOnModalClose);
      overlayModalAutoShownLayer.set(restoreOnModalClose, layer);
    }

    if (targetWindow.webContents.isLoading()) {
      targetWindow.webContents.once("did-finish-load", () => {
        if (
          targetWindow &&
          !targetWindow.isDestroyed() &&
          !targetWindow.webContents.isLoading()
        ) {
          sendNow();
        }
      });
      return true;
    }

    sendNow();
    return true;
  };

  const openRuntimeOptionsPalette = (): void => {
    sendToActiveOverlayWindow("runtime-options:open", undefined, {
      restoreOnModalClose: "runtime-options",
    });
  };

  const handleOverlayModalClosed = (modal: OverlayHostedModal): void => {
    if (!restoreVisibleOverlayOnModalClose.has(modal)) return;
    restoreVisibleOverlayOnModalClose.delete(modal);
    const layer = overlayModalAutoShownLayer.get(modal);
    overlayModalAutoShownLayer.delete(modal);
    if (!layer) return;
    const shouldKeepLayerVisible = [...restoreVisibleOverlayOnModalClose].some(
      (pendingModal) => overlayModalAutoShownLayer.get(pendingModal) === layer,
    );
    if (shouldKeepLayerVisible) return;

    if (layer === "visible") {
      const mainWindow = deps.getMainWindow();
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.hide();
      }
      return;
    }
    const invisibleWindow = deps.getInvisibleWindow();
    if (invisibleWindow && !invisibleWindow.isDestroyed()) {
      invisibleWindow.hide();
    }
  };

  return {
    sendToActiveOverlayWindow,
    openRuntimeOptionsPalette,
    handleOverlayModalClosed,
    getRestoreVisibleOverlayOnModalClose: () =>
      restoreVisibleOverlayOnModalClose,
  };
}

export type { OverlayHostedModal };
